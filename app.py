import io
import os
import shutil
import sqlite3
import tempfile
import gc

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
[data-testid="stMetricValue"] {font-size: clamp(1.15rem, 2.2vw, 1.9rem); white-space: nowrap;}
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
    """Procesa los Excel uno por uno para reducir el uso máximo de RAM."""
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

    for i, path in enumerate(saved_paths, 1):
        name = os.path.basename(path)
        status.write(f"⏳ Procesando {i}/{len(saved_paths)}: {name}")

        src = os.path.join(run_dir, f"source_{i}.sqlite")
        db = None
        try:
            db = create_db(src)
            n = process_to_db(path, db)

            db.execute("CREATE INDEX IF NOT EXISTS idx_psp ON operations(psp_tin)")
            db.execute("CREATE INDEX IF NOT EXISTS idx_com_fecha ON operations(comercio,fecha)")
            # Identificador real del comercio: se toma directamente de Com_Public_ID del Excel.
            db.execute("CREATE INDEX IF NOT EXISTS idx_com_public_id ON operations(com_public_id)")
            db.commit()
            db.close()
            db = None

            source_paths[i - 1] = src
            total_rows += n
            status.write(f"✅ {i}/{len(saved_paths)} — {name} — {n:,} filas")
        except Exception as e:
            if db is not None:
                try:
                    db.close()
                except Exception:
                    pass
            errors.append(f"{name}: {e}")
            status.write(f"⚠️ {i}/{len(saved_paths)} — {name}: {e}")

        gc.collect()
        progress_bar.progress(i / len(saved_paths))

    progress_bar.empty()
    status.empty()

    valid_sources = [(i, p) for i, p in enumerate(source_paths, 1) if p]
    if not valid_sources:
        shutil.rmtree(run_dir, ignore_errors=True)
        raise RuntimeError("No se pudo procesar ningún archivo. Revisa los mensajes de error.")

    master_path = os.path.join(run_dir, "master.sqlite")
    master = sqlite3.connect(master_path)
    master.execute("CREATE TABLE IF NOT EXISTS sources (idx INTEGER PRIMARY KEY, path TEXT)")
    master.executemany("INSERT INTO sources(idx,path) VALUES (?,?)", valid_sources)
    master.commit()
    master.close()

    ss.run_dir = run_dir
    ss.db_path = master_path
    ss.total_rows = total_rows
    ss.processed_names = [os.path.basename(p) for p in saved_paths]
    ss.errors = errors
    gc.collect()
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
    """Genera el XLSX solo cuando el usuario solicita una exportación."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31] or "Detalle")
    buf.seek(0)
    wb = load_workbook(buf)
    style(wb)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def lazy_download_button(label, df_factory, file_name, sheet_name, key, disabled=False):
    """Evita crear Excel grandes durante cada rerun de Streamlit."""
    if disabled:
        return
    if st.button(f"🛠️ Preparar {label}", key=f"prepare_{key}"):
        with st.spinner("Generando Excel..."):
            df = df_factory()
            if df is None or df.empty:
                st.info("No hay datos para exportar.")
            else:
                data = to_excel_bytes(df, sheet_name)
                st.download_button(
                    label,
                    data=data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{key}",
                    on_click="ignore",
                )
                del df
                gc.collect()


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
    procesar = c1.button("🚀 Procesar", width="stretch", disabled=not uploaded)
    limpiar = c2.button("🗑 Limpiar", width="stretch")

    if procesar and uploaded:
        try:
            with st.spinner("Procesando Excel uno por uno para proteger la memoria..."):
                process_uploaded_files(uploaded)
        except Exception as e:
            ss.errors = [f"Error general de procesamiento: {e}"]
            st.error(
                "No se pudo completar el procesamiento. "
                "La aplicación sigue abierta; revisa el detalle del error."
            )
        else:
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
        if st.button("➕ Agregar regla", width="stretch"):
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
            st.dataframe(rules_df.drop(columns="id"), hide_index=True, width="stretch")
            to_remove = st.multiselect("Quitar regla(s)", rules_df["id"].tolist(),
                                        format_func=lambda i: rules_df.set_index("id").loc[i, "Comercio"])
            if to_remove and st.button("🗑 Quitar seleccionadas", width="stretch"):
                con.executemany("DELETE FROM test_rules WHERE id=?", [(i,) for i in to_remove])
                con.commit()
                st.rerun()
        else:
            st.caption("No hay reglas de prueba agregadas.")

# --------------------------------------------------------------------------
# Com_Public_ID: identificador real del comercio tomado del Excel.
# IMPORTANTE: core.process_to_db debe guardar la columna Excel Com_Public_ID
# en operations.com_public_id; no se usa ni se crea el campo Com.
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
    m1, m2, m3, m4 = st.columns(4)
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
    m5, m6 = st.columns(2)
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
    resumen = pd.DataFrame(rows, columns=["Com_Public_ID", "comercio", "moneda", "total_recaudo", "total_comision", "total_neto", "operaciones"])
    if not resumen.empty:
        resumen["comision_sin_igv"] = resumen.total_comision - (resumen.total_comision * 18 / 118)
        n = pd.read_sql_query(
            f"SELECT com_public_id Com_Public_ID,comercio,moneda,COUNT(*) negativos FROM operations "
f"WHERE {filtro.currency_sql()} AND es_negativo=1 GROUP BY com_public_id,comercio,moneda", con)
        resumen = resumen.merge(n, on=["Com_Public_ID", "comercio", "moneda"], how="left")
        resumen["negativos"] = resumen.negativos.fillna(0).astype(int)
        resumen = resumen[["Com_Public_ID", "comercio", "moneda", "total_recaudo", "total_comision", "comision_sin_igv", "total_neto", "operaciones", "negativos"]]

    st.dataframe(resumen, width="stretch", hide_index=True)
    lazy_download_button(
        "📥 Descargar resumen",
        lambda: resumen.copy(),
        "Resumen_por_Comercio.xlsx",
        "Resumen_Comercio",
        key="resumen",
        disabled=resumen.empty,
    )

    st.markdown("#### 🔁 PSP_TIN duplicados — solo una operación cuenta en el cálculo")
    st.markdown(
        '<p class="note">Todos los PSP_TIN repetidos se muestran aquí. En el cálculo se toma una sola '
        'positiva por PSP_TIN + Moneda y el reverso negativo reduce el resultado neto.</p>', unsafe_allow_html=True)
    dup_df = query_duplicates(con, filtro, limit=5000)
    st.dataframe(dup_df, width="stretch", hide_index=True)
    lazy_download_button(
        "📥 Descargar PSP_TIN duplicados",
        lambda: query_duplicates(con, filtro, None),
        "PSP_TIN_Duplicados.xlsx",
        "PSP_TIN_Duplicados",
        key="psp_duplicates",
        disabled=dup_df.empty,
    )
    st.caption(f"Leídas: {ss.total_rows:,} | Base principal: {ops:,} | Negativos: {neg:,} | PSP_TIN duplicados: {dup_ids:,}")

# --------------------------------------------------------------------------
# COMERCIO + MES
# --------------------------------------------------------------------------
with tab_com:
    cc1, cc2, cc3 = st.columns([2, 1, 1.4])
    comercio_sel = cc1.selectbox("Comercio", comercios_all, index=0 if comercios_all else None, key="comercio_sel")
    # El comercio se identifica por Com_Public_ID; el nombre solo se muestra como referencia.
    ids_comercio = [r[0] for r in con.execute(
        f"SELECT DISTINCT com_public_id FROM operations WHERE {filtro.main_where()} AND comercio=? "
        "AND TRIM(COALESCE(com_public_id,''))<>'' ORDER BY com_public_id", (comercio_sel,)
    ).fetchall()] if comercio_sel else []
    com_public_sel = cc1.selectbox("Com_Public_ID", ids_comercio, index=0 if ids_comercio else None, key="com_public_sel")
    meses_com = [r[0] for r in con.execute(
        f"SELECT DISTINCT substr(fecha,1,7) FROM operations WHERE {filtro.main_where()} AND comercio=? AND com_public_id=? AND fecha<>'' ORDER BY 1",
        (comercio_sel, com_public_sel),
    ).fetchall()] if comercio_sel else []
    mes_sel = cc2.selectbox("Mes", meses_com, index=0 if meses_com else None, key="mes_com_sel")
    incluir_neg = cc3.checkbox("Incluir negativos/reversos en detalle (no afecta cálculos)", key="incluir_neg")

    if comercio_sel:
        where = f"{filtro.main_where()} AND comercio=? AND com_public_id=?"
        params = [comercio_sel, com_public_sel]
        if mes_sel:
            where += " AND substr(fecha,1,7)=?"
            params.append(mes_sel)
        if incluir_neg:
            base_sql = f"({where}) OR ({filtro.currency_sql()} AND {filtro.test_exclusion_sql()} AND es_negativo=1 AND comercio=?"
            neg_params = [comercio_sel, com_public_sel]
            if mes_sel:
                base_sql += " AND substr(fecha,1,7)=?"
                neg_params.append(mes_sel)
            base_sql += ")"
            where = base_sql
            params = params + neg_params
        det_com = pd.read_sql_query(
            f"SELECT com_public_id Com_Public_ID,fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',
            f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
            f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {where} ORDER BY id LIMIT 5000",
            con, params=params)
        st.dataframe(det_com, width="stretch", hide_index=True)
        st.caption("Máximo 5,000 filas visibles en pantalla; la exportación incluye todo.")

        d1, d2 = st.columns(2)
        if mes_sel:
            def make_com_month_df():
                full_where = f"{filtro.main_where()} AND comercio=? AND com_public_id=? AND substr(fecha,1,7)=?"
                return pd.read_sql_query(
                    f"SELECT com_public_id Com_Public_ID,fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',
                    f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
                    f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {full_where} ORDER BY id",
                    con, params=[comercio_sel, com_public_sel, mes_sel])

            with d1:
                lazy_download_button(
                    "📥 Descargar comercio + mes",
                    make_com_month_df,
                    f"Detalle_{comercio_sel}_{mes_sel}.xlsx".replace("|", "_"),
                    "Detalle",
                    key="comercio_mes",
                )

        def make_com_full_df():
            return pd.read_sql_query(
                f"SELECT com_public_id Com_Public_ID,fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',
                f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
                f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {filtro.main_where()} AND comercio=? ORDER BY id",
                con, params=[comercio_sel])

        with d2:
            lazy_download_button(
                "📥 Descargar comercio completo",
                make_com_full_df,
                f"Detalle_{comercio_sel}.xlsx",
                "Detalle",
                key="comercio_completo",
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
    ids_mes = [r[0] for r in con.execute(
        f"SELECT DISTINCT com_public_id FROM operations WHERE {filtro.main_where()} AND substr(fecha,1,7)=? "
        "AND TRIM(COALESCE(com_public_id,''))<>'' ORDER BY com_public_id", (mes_general,)
    ).fetchall()] if mes_general else []
    com_public_mes = st.selectbox("Com_Public_ID", ids_mes, index=0 if ids_mes else None, key="com_public_mes")
    if mes_general:
        det_mes = pd.read_sql_query(
            f"SELECT com_public_id Com_Public_ID,fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',
            f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
            f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {filtro.main_where()} "
            f"AND substr(fecha,1,7)=? ORDER BY id LIMIT 5000", con, params=[mes_general])
        st.dataframe(det_mes, width="stretch", hide_index=True)
        def make_month_df():
            return pd.read_sql_query(
                f"SELECT com_public_id Com_Public_ID,fecha FECHA,comercio Com_Nombre,Deb_Doc,Deb_Nombre,psp_tin,metodo_pago 'Método de Pago',
                f"SET_referencia,Fecha_Transferencia 'Fecha Transferencia',Banco_Transferencia 'Banco Transferencia',"
                f"recaudo RECAUDO,comision COMISION,neto NETO FROM operations WHERE {filtro.main_where()} "
                f"AND substr(fecha,1,7)=? ORDER BY id", con, params=[mes_general])

        lazy_download_button(
            "📥 Descargar mes",
            make_month_df,
            f"Detalle_{mes_general}.xlsx",
            str(mes_general),
            key="mes",
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
    if not neg_df.empty and "Com_Public_ID" not in neg_df.columns:
        # Se conserva query_details y solo se añade el identificador desde operations.
        neg_df = pd.read_sql_query(
            f"SELECT com_public_id Com_Public_ID, fecha FECHA, comercio Com_Nombre, Deb_Doc, Deb_Nombre, psp_tin, "
            f"metodo_pago 'Método de Pago', SET_referencia, Fecha_Transferencia 'Fecha Transferencia', "
            f"Banco_Transferencia 'Banco Transferencia', recaudo RECAUDO, comision COMISION, neto NETO "
            f"FROM operations WHERE {neg_where} ORDER BY id LIMIT 5000", con)
    st.dataframe(neg_df, width="stretch", hide_index=True)
    lazy_download_button(
        "📥 Descargar negativos",
        lambda: pd.read_sql_query(
            f"SELECT com_public_id Com_Public_ID, fecha FECHA, comercio Com_Nombre, Deb_Doc, Deb_Nombre, psp_tin, "
            f"metodo_pago 'Método de Pago', SET_referencia, Fecha_Transferencia 'Fecha Transferencia', "
            f"Banco_Transferencia 'Banco Transferencia', recaudo RECAUDO, comision COMISION, neto NETO "
            f"FROM operations WHERE {neg_where} ORDER BY id", con),
        "Negativos.xlsx",
        "Negativos",
        key="negativos",
        disabled=neg_df.empty,
    )
