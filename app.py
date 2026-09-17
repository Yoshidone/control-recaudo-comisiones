import io
import os
import shutil
import sqlite3
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from core import (
    TITLE, create_db, process_to_db, style, Filtro, build_conn,
    query_duplicates, query_details,
)

st.set_page_config(page_title=TITLE, page_icon="💰", layout="wide")

# --------------------------------------------------------------------------
# Estilo moderno (solo visual — no toca ningún cálculo)
# --------------------------------------------------------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.6rem; max-width: 1400px;}
.kashio-hero {
    background: linear-gradient(120deg, #103a5e 0%, #1f6fa8 100%);
    padding: 1.4rem 1.8rem; border-radius: 16px; color: white; margin-bottom: 1.1rem;
    box-shadow: 0 8px 24px rgba(16,58,94,.25);
}
.kashio-hero h1 {margin:0; font-size: 1.6rem;}
.kashio-hero p {margin:.35rem 0 0; opacity:.85; font-size:.92rem;}
[data-testid="stMetric"] {
    background: white; border: 1px solid #e6ebf1; border-radius: 14px;
    padding: .9rem .9rem .6rem; box-shadow: 0 2px 8px rgba(16,58,94,.06);
}
[data-testid="stMetricLabel"] {font-weight:600; color:#4b5c6b;}
.stTabs [data-baseweb="tab"] {font-weight:600; padding: .5rem 1rem;}
.igv-card {
    background:#fff7e6; border:1px solid #ffdf99; border-radius:12px;
    padding:.8rem 1.1rem; font-weight:600; color:#8a5a00; margin:.6rem 0 1rem;
}
.note {color:#8a5a00; font-size:.88rem;}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="kashio-hero">
  <h1>💰 {TITLE}</h1>
  <p>Sube uno o varios Excel, concilia recaudo y comisiones, y comparte resultados con tu equipo desde cualquier navegador.</p>
</div>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# Estado de sesión
# --------------------------------------------------------------------------
ss = st.session_state
ss.setdefault("db_path", None)
ss.setdefault("db_con", None)
ss.setdefault("total_rows", 0)
ss.setdefault("run_dir", None)
ss.setdefault("processed_names", [])
ss.setdefault("period_text", "—")
ss.setdefault("errors", [])


def reset_state():
    if ss.db_con is not None:
        try:
            ss.db_con.close()
        except Exception:
            pass
    if ss.run_dir and os.path.isdir(ss.run_dir):
        shutil.rmtree(ss.run_dir, ignore_errors=True)
    ss.db_path = None
    ss.db_con = None
    ss.total_rows = 0
    ss.run_dir = None
    ss.processed_names = []
    ss.period_text = "—"
    ss.errors = []


def process_uploaded_files(uploaded_files):
    reset_state()
    run_dir = tempfile.mkdtemp(prefix="control_recaudo_comisiones_")
    saved_paths = []
    for f in uploaded_files:
        p = os.path.join(run_dir, f.name)
        with open(p, "wb") as out:
            out.write(f.getbuffer())
        saved_paths.append(p)

    progress_bar = st.progress(0.0)
    status = st.empty()
    source_paths = [None] * len(saved_paths)
    errors = []
    total_rows = 0
    done = 0

    def work(i, path):
        src = os.path.join(run_dir, f"source_{i}.sqlite")
        db = create_db(src)
        n = process_to_db(path, db)
        db.execute("CREATE INDEX IF NOT EXISTS idx_psp ON operations(psp_tin)")
        db.execute("CREATE INDEX IF NOT EXISTS idx_com_fecha ON operations(comercio,fecha)")
        db.commit()
        db.close()
        return i, src, n

    # Lectura en paralelo de los distintos Excel: cada archivo escribe en su
    # propio SQLite, así que no hay contención entre hilos. Con varios Excel
    # grandes esto reduce notablemente el tiempo total de carga.
    with ThreadPoolExecutor(max_workers=min(4, len(saved_paths)) or 1) as ex:
        futures = {ex.submit(work, i, p): p for i, p in enumerate(saved_paths, 1)}
        for fut in as_completed(futures):
            path = futures[fut]
            done += 1
            try:
                i, src, n = fut.result()
                source_paths[i - 1] = src
                total_rows += n
                status.write(f"✅ {os.path.basename(path)} — {n:,} filas")
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")
                status.write(f"⚠️ {os.path.basename(path)}: {e}")
            progress_bar.progress(done / len(saved_paths))

    progress_bar.empty()
    status.empty()

    master_path = os.path.join(run_dir, "master.sqlite")
    master = sqlite3.connect(master_path)
    master.execute("CREATE TABLE IF NOT EXISTS sources (idx INTEGER PRIMARY KEY, path TEXT)")
    master.executemany(
        "INSERT INTO sources(idx,path) VALUES (?,?)",
        [(i, p) for i, p in enumerate(source_paths) if p]
    )
    master.commit()
    master.close()

    ss.run_dir = run_dir
    ss.db_path = master_path
    ss.total_rows = total_rows
    ss.processed_names = [os.path.basename(p) for p in saved_paths]
    ss.errors = errors


def get_conn():
    if ss.db_con is not None:
        return ss.db_con
    con = build_conn(ss.db_path)
    ss.db_con = con
    # Periodo del Excel (idéntico al cálculo original)
    f1, f2 = con.execute(
        "SELECT MIN(fecha),MAX(fecha) FROM operations WHERE fecha IS NOT NULL AND fecha<>''"
    ).fetchone()
    if f1 and f2:
        try:
            a, b = pd.to_datetime(f1), pd.to_datetime(f2)
            ss.period_text = f"{a:%d/%m/%Y %H:%M:%S} → {b:%d/%m/%Y %H:%M:%S}"
        except Exception:
            ss.period_text = f"{f1} → {f2}"
    return con


def to_excel_bytes(df, sheet_name):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31] or "Detalle")
    buf.seek(0)
    wb = load_workbook(buf)
    style(wb)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------
# Barra lateral: carga de archivos
# --------------------------------------------------------------------------
with st.sidebar:
    st.subheader("📂 Archivos")
    uploaded = st.file_uploader(
        "Excel de recaudo/comisiones", type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="Puedes subir varios Excel a la vez; se procesan y consolidan juntos.",
    )
    c1, c2 = st.columns(2)
    procesar = c1.button("🚀 Procesar", use_container_width=True, disabled=not uploaded)
    limpiar = c2.button("🗑 Limpiar", use_container_width=True)

    if procesar and uploaded:
        with st.spinner("Leyendo Excel..."):
            process_uploaded_files(uploaded)
        st.rerun()

    if limpiar:
        reset_state()
        st.rerun()

    if ss.db_path:
        st.success(f"Leídas: {ss.total_rows:,} filas")
        st.caption("Archivos: " + ", ".join(ss.processed_names))
        st.caption(f"Periodo (TX_GMT_Peru): {ss.period_text}")
        if ss.errors:
            st.warning("\n".join(ss.errors))
    else:
        st.info("Sube uno o varios Excel y pulsa **Procesar**.")

    if ss.db_path:
        st.divider()
        st.subheader("💱 Moneda")
        currency = st.radio("Filtrar por moneda", ["Todas", "PEN", "USD"], horizontal=True, label_visibility="collapsed")
    else:
        currency = "Todas"

if not ss.db_path:
    st.markdown(
        "### 👋 Empieza aquí\n"
        "1. Sube uno o varios archivos Excel desde la barra lateral (`.xlsx` / `.xlsm`).\n"
        "2. Pulsa **Procesar**.\n"
        "3. Revisa el dashboard, exporta reportes y comparte el enlace de esta app con tu equipo.\n"
    )
    st.stop()

con = get_conn()
filtro = Filtro(currency)

# --------------------------------------------------------------------------
# Reglas de prueba (exclusiones manuales) — misma tabla/lógica que el original
# --------------------------------------------------------------------------
comercios_all = [r[0] for r in con.execute(
    f"SELECT DISTINCT comercio FROM operations WHERE {filtro.main_where()} AND TRIM(COALESCE(comercio,''))<>'' ORDER BY comercio"
).fetchall()]

with st.sidebar:
    st.divider()
    with st.expander("🧪 Operaciones de prueba (opcional)"):
        st.caption("Solo se excluyen del cálculo cuando agregas una regla aquí. Por defecto no se excluye ningún comercio.")
        tr_com = st.selectbox("Comercio", comercios_all, index=None, placeholder="Elige un comercio", key="tr_com")
        tr_d1 = st.date_input("Desde", value=None, key="tr_d1", format="DD/MM/YYYY")
        tr_d2 = st.date_input("Hasta", value=None, key="tr_d2", format="DD/MM/YYYY")
        if st.button("➕ Agregar regla", use_container_width=True):
            if not tr_com:
                st.warning("Selecciona un comercio.")
            elif not tr_d1 or not tr_d2:
                st.warning("Completa ambas fechas.")
            elif tr_d1 > tr_d2:
                st.warning("La fecha Desde no puede ser mayor que Hasta.")
            else:
                con.execute("INSERT INTO test_rules(comercio,fecha_desde,fecha_hasta) VALUES (?,?,?)",
                            (tr_com, tr_d1.strftime("%Y-%m-%d"), tr_d2.strftime("%Y-%m-%d")))
                con.commit()
                st.rerun()

        rules = con.execute("SELECT id,comercio,fecha_desde,fecha_hasta FROM test_rules ORDER BY fecha_desde,comercio").fetchall()
        if rules:
            rules_df = pd.DataFrame(rules, columns=["id", "Comercio", "Desde", "Hasta"])
            st.dataframe(rules_df.drop(columns="id"), hide_index=True, use_container_width=True)
            to_remove = st.multiselect("Quitar regla(s)", rules_df["id"].tolist(),
                                        format_func=lambda i: rules_df.set_index("id").loc[i, "Comercio"])
            if to_remove and st.button("🗑 Quitar seleccionadas", use_container_width=True):
                con.executemany("DELETE FROM test_rules WHERE id=?", [(i,) for i in to_remove])
                con.commit()
                st.rerun()
        else:
            st.caption("No hay reglas de prueba agregadas.")

# --------------------------------------------------------------------------
# Consultas agregadas (idénticas al método refresh() original)
# --------------------------------------------------------------------------
w = filtro.calc_where()
rec, sf, net, ops, coms = con.execute(
    f"SELECT COALESCE(SUM(recaudo),0),COALESCE(SUM(comision),0),COALESCE(SUM(neto),0),COUNT(*),"
    f"COUNT(DISTINCT comercio) FROM operations WHERE {w}"
).fetchone()
neg = int(con.execute(f"SELECT COUNT(*) FROM operations WHERE {filtro.currency_sql()} AND es_negativo=1").fetchone()[0])
dup_ids = int(con.execute(
    f"SELECT COUNT(*) FROM dup_keys WHERE UPPER(TRIM(COALESCE(moneda,'')))={currency.strip().upper()!r}"
    if currency in ("PEN", "USD") else "SELECT COUNT(*) FROM dup_keys"
).fetchone()[0])

tab_dash, tab_com, tab_mes, tab_neg = st.tabs(["📊 Dashboard", "🏪 Comercio + Mes", "📅 Mes", "🔴 Negativos"])

# --------------------------------------------------------------------------
# DASHBOARD
# --------------------------------------------------------------------------
with tab_dash:
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Comercios", f"{coms:,}")
    if currency == "Todas":
        cur_rows = con.execute(f"SELECT moneda,SUM(recaudo),SUM(comision),SUM(neto) FROM operations WHERE {w} GROUP BY moneda ORDER BY moneda").fetchall()
        vals = {str(r[0] or ""): (float(r[1] or 0), float(r[2] or 0), float(r[3] or 0)) for r in cur_rows}
        pen = vals.get("PEN", (0, 0, 0))
        usd = vals.get("USD", (0, 0, 0))
        m2.metric("Recaudo total", f"S/ {pen[0]:,.2f}", f"US$ {usd[0]:,.2f}")
        m3.metric("Comisión total", f"S/ {pen[1]:,.2f}", f"US$ {usd[1]:,.2f}")
        m4.metric("Neto", f"S/ {pen[2]:,.2f}", f"US$ {usd[2]:,.2f}")
        sinpen = pen[1] - (pen[1] * 18 / 118)
        sinusd = usd[1] - (usd[1] * 18 / 118)
        igv_text = f"Comisión sin IGV — PEN: S/ {sinpen:,.2f}&nbsp;&nbsp;|&nbsp;&nbsp;USD: US$ {sinusd:,.2f}"
    else:
        sym = "S/" if currency == "PEN" else "US$"
        m2.metric("Recaudo total", f"{sym} {rec:,.2f}")
        m3.metric("Comisión total", f"{sym} {sf:,.2f}")
        m4.metric("Neto", f"{sym} {net:,.2f}")
        igv_text = f"Comisión sin IGV: {sym} {sf - (sf * 18 / 118):,.2f}"
    m5.metric("Negativos", f"{neg:,}")
    m6.metric("PSP_TIN duplicados", f"{dup_ids:,}")

    st.markdown(f'<div class="igv-card">💰 {igv_text}</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="note">La comisión final usa el SF_amount neto de las operaciones válidas. '
        'En PSP_TIN repetidos, solo una positiva cuenta; el reverso negativo reduce el resultado '
        'y todas las filas se muestran en Duplicados.</p>', unsafe_allow_html=True)

    rows = con.execute(
        f"SELECT comercio,moneda,SUM(recaudo) total_recaudo,SUM(comision) total_comision,SUM(neto) total_neto,"
        f"COUNT(*) operaciones FROM operations WHERE {w} GROUP BY comercio,moneda ORDER BY comercio,moneda"
    ).fetchall()
    resumen = pd.DataFrame(rows, columns=["comercio", "moneda", "total_recaudo", "total_comision", "total_neto", "operaciones"])
    if not resumen.empty:
        resumen["comision_sin_igv"] = resumen.total_comision - (resumen.total_comision * 18 / 118)
        n = pd.read_sql_query(
            f"SELECT comercio,moneda,COUNT(*) negativos FROM operations WHERE {filtro.currency_sql()} AND es_negativo=1 GROUP BY comercio,moneda", con)
        resumen = resumen.merge(n, on=["comercio", "moneda"], how="left")
        resumen["negativos"] = resumen.negativos.fillna(0).astype(int)
        resumen = resumen[["comercio", "moneda", "total_recaudo", "total_comision", "comision_sin_igv", "total_neto", "operaciones", "negativos"]]

    st.dataframe(resumen, use_container_width=True, hide_index=True)
    st.download_button(
        "📥 Exportar resumen", data=to_excel_bytes(resumen, "Resumen_Comercio"),
        file_name="Resumen_por_Comercio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=resumen.empty,
    )

    st.markdown("#### 🔁 PSP_TIN duplicados — solo una operación cuenta en el cálculo")
    st.markdown(
        '<p class="note">Todos los PSP_TIN repetidos se muestran aquí. En el cálculo se toma una sola '
        'positiva por PSP_TIN + Moneda y el reverso negativo reduce el resultado neto.</p>', unsafe_allow_html=True)
    dup_df = query_duplicates(con, filtro, limit=5000)
    st.dataframe(dup_df, use_container_width=True, hide_index=True)
    st.download_button(
        "📥 Exportar PSP_TIN duplicados", data=to_excel_bytes(query_duplicates(con, filtro, None), "PSP_TIN_Duplicados"),
        file_name="PSP_TIN_Duplicados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=dup_df.empty,
    )
    st.caption(f"Leídas: {ss.total_rows:,} | Base principal: {ops:,} | Negativos: {neg:,} | PSP_TIN duplicados: {dup_ids:,}")

# --------------------------------------------------------------------------
# COMERCIO + MES
# --------------------------------------------------------------------------
with tab_com:
    cc1, cc2, cc3 = st.columns([2, 1, 1.4])
    comercio_sel = cc1.selectbox("Comercio", comercios_all, index=0 if comercios_all else None, key="comercio_sel")
    meses_com = [r[0] for r in con.execute(
        f"SELECT DISTINCT substr(fecha,1,7) FROM operations WHERE {filtro.main_where()} AND comercio=? AND fecha<>'' ORDER BY 1",
        (comercio_sel,)
    ).fetchall()] if comercio_sel else []
    mes_sel = cc2.selectbox("Mes", meses_com, index=0 if meses_com else None, key="mes_com_sel")
    incluir_neg = cc3.checkbox("Incluir negativos/reversos en detalle (no afecta cálculos)", key="incluir_neg")

    if comercio_sel:
        where = f"{filtro.main_where()} AND comercio=?"
        params = [comercio_sel]
        if mes_sel:
            where += " AND substr(fecha,1,7)=?"
            params.append(mes_sel)
        if incluir_neg:
            base_sql = f"({where}) OR ({filtro.currency_sql()} AND {filtro.test_exclusion_sql()} AND es_negativo=1 AND comercio=?"
            neg_params = [comercio_sel]
            if mes_sel:
                base_sql += " AND substr(fecha,1,7)=?"
                neg_params.append(mes_sel)
            base_sql += ")"
            where = base_sql
            params = params + neg_params
        det_com = pd.read_sql_query(
            f"SELECT fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',"
            f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
            f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {where} ORDER BY id LIMIT 5000",
            con, params=params)
        st.dataframe(det_com, use_container_width=True, hide_index=True)
        st.caption("Máximo 5,000 filas visibles en pantalla; la exportación incluye todo.")

        d1, d2 = st.columns(2)
        if mes_sel:
            full_where = f"{filtro.main_where()} AND comercio=? AND substr(fecha,1,7)=?"
            full_params = [comercio_sel, mes_sel]
            full_df = pd.read_sql_query(
                f"SELECT fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',"
                f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
                f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {full_where} ORDER BY id",
                con, params=full_params)
            d1.download_button(
                "📥 Exportar comercio + mes", data=to_excel_bytes(full_df, "Detalle"),
                file_name=f"Detalle_{comercio_sel}_{mes_sel}.xlsx".replace("|", "_"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=full_df.empty,
            )
        full_com_df = pd.read_sql_query(
            f"SELECT fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',"
            f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
            f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {filtro.main_where()} AND comercio=? ORDER BY id",
            con, params=[comercio_sel])
        d2.download_button(
            "📥 Exportar comercio completo", data=to_excel_bytes(full_com_df, "Detalle"),
            file_name=f"Detalle_{comercio_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=full_com_df.empty,
        )
    else:
        st.info("No hay comercios disponibles con los filtros actuales.")

# --------------------------------------------------------------------------
# MES
# --------------------------------------------------------------------------
with tab_mes:
    meses_all = [r[0] for r in con.execute(
        f"SELECT DISTINCT substr(fecha,1,7) FROM operations WHERE {filtro.main_where()} AND fecha<>'' ORDER BY 1"
    ).fetchall()]
    mes_general = st.selectbox("Mes", meses_all, index=0 if meses_all else None, key="mes_general")
    if mes_general:
        det_mes = pd.read_sql_query(
            f"SELECT fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',"
            f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
            f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {filtro.main_where()} "
            f"AND substr(fecha,1,7)=? ORDER BY id LIMIT 5000", con, params=[mes_general])
        st.dataframe(det_mes, use_container_width=True, hide_index=True)
        full_mes_df = pd.read_sql_query(
            f"SELECT fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',"
            f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
            f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {filtro.main_where()} "
            f"AND substr(fecha,1,7)=? ORDER BY id", con, params=[mes_general])
        st.download_button(
            "📥 Exportar mes", data=to_excel_bytes(full_mes_df, str(mes_general)),
            file_name=f"Detalle_{mes_general}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=full_mes_df.empty,
        )
    else:
        st.info("No hay meses disponibles con los filtros actuales.")

# --------------------------------------------------------------------------
# NEGATIVOS
# --------------------------------------------------------------------------
with tab_neg:
    st.markdown('<p class="note">Negativos separados. No aparecen en el detalle principal.</p>', unsafe_allow_html=True)
    neg_where = f"{filtro.currency_sql()} AND es_negativo=1"
    neg_df = query_details(con, neg_where, 5000)
    st.dataframe(neg_df, use_container_width=True, hide_index=True)
    st.download_button(
        "📥 Descargar negativos", data=to_excel_bytes(query_details(con, neg_where, None), "Negativos"),
        file_name="Negativos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=neg_df.empty,
    )
