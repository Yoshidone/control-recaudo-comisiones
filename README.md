# Control de Recaudo y Comisiones

Aplicación web en Streamlit para cargar uno o varios Excel de recaudo/comisiones,
procesarlos y generar reportes.

## Archivos principales
- `app.py`: interfaz Streamlit.
- `core.py`: motor de procesamiento y cálculos.
- `requirements.txt`: dependencias.
- `.streamlit/config.toml`: configuración de la aplicación.

## Desplegar en Streamlit Community Cloud

1. Crea un repositorio nuevo en GitHub.
2. Sube **el contenido de esta carpeta** al repositorio (de modo que `app.py`
   quede en la raíz del repositorio).
3. En Streamlit Community Cloud selecciona el repositorio y el archivo:
   `app.py`.
4. Pulsa **Deploy**.
5. Streamlit te dará una dirección pública `https://....streamlit.app`.

No necesitas ejecutar Python, CMD, Cloudflare ni dejar tu PC encendida para
que la versión desplegada esté disponible.

## Importante sobre los datos

Los Excel que cada usuario sube se procesan en su propia sesión de Streamlit.
La aplicación no usa una ruta local de Windows ni una base SQLite fija de tu PC.

Si quieres que varios compañeros compartan una **misma base de datos persistente**
de clientes/tarifas/reglas entre sesiones, esa parte debe conectarse a una base
de datos en la nube; no se debe resolver subiendo un `clientes.db` local al
repositorio.

## Cálculos

El motor mantiene la lógica implementada en `core.py`, incluyendo el tratamiento
de negativos/reversos y la deduplicación por PSP_TIN + moneda.
