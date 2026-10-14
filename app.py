import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Dashboard de Control de Viajes y Fletes",
    page_icon="🚚",
    layout="wide"
)

st.title("🚚 Dashboard de Control de Viajes y Fletes")
st.markdown("Carga tu archivo de Excel para gestionar viajes pendientes de facturar, analizar rutas y descargar reportes organizados.")

# Carga de archivo
st.sidebar.header("📁 Cargar Archivo de Excel")
uploaded_file = st.sidebar.file_uploader("Selecciona tu archivo (.xlsx o .xls)", type=["xlsx", "xls", "csv"])

# Datos de muestra adaptados con columna de estatus y campos clave
def cargar_datos_muestra():
    data = {
        "Viaje": [82373, 83637, 83884, 83889, 83915, 83920, 83927, 83932, 83941, 83944, 83946],
        "Fecha": ["2026-08-24", "2026-09-25", "2026-10-02", "2026-10-02", "2026-10-13", "2026-10-14", "2026-10-02", "2026-10-01", "2026-10-05", "2026-10-05", "2026-10-02"],
        "Cliente": [
            "TRANSPORTES CALPI, S.A. DE C.V.",
            "DERIVADOS DE MAIZ DE HONDURAS, S.A. DE C.V.",
            "DISTRIBUIDORA ZABLAH S.A. DE C.V.",
            "DISTRIBUIDORA ZABLAH S.A. DE C.V.",
            "DISTRIBUIDORA CHAVARRIA, S.A. DE C.V",
            "DISTRIBUIDORA CHAVARRIA, S.A. DE C.V",
            "HARISA, S.A. DE C.V.",
            "CERVECERIA SALVADOREÑA, S.A. DE C.V.",
            "MERCANTIL DE COMERCIO DE EL SALVADOR, S.A.",
            "MOLINOS MODERNOS, S.A.",
            "HARISA, S.A. DE C.V."
        ],
        "Completado": ["ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO", "ASIGNADO"],
        "Valor de Flete": [0, 950.00, 1650.00, 1650.00, 700.00, 700.00, 1203.75, 205.00, 1458.00, 827.50, 1203.75],
        "Observaciones": ["", "", "", "", "", "", "CON PROFORMA", "", "", "", "CON PROFORMA"],
        "Tipo Viaje": ["VIAJE LOCAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE LOCAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL"],
        "Motorista": ["FERMIN DE JESUS", "ROBERTO CARLOS", "JUAN PEREZ", "JUAN PEREZ", "CARLOS GOMEZ", "CARLOS GOMEZ", "MARIO CRUZ", "ROBERTO RIVAS", "ANA LOPEZ", "MARIO CRUZ", "MARIO CRUZ"]
    }
    return pd.DataFrame(data)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.sidebar.success("¡Archivo cargado con éxito!")
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()
else:
    st.sidebar.info("📌 Mostrando datos de demostración. Sube tu Excel para analizar tu información real.")
    df = cargar_datos_muestra()

# Detección inteligente de columnas
def buscar_columna(patrones, df_cols, defecto=None):
    for col in df_cols:
        col_clean = str(col).lower().strip()
        for pat in patrones:
            if pat in col_clean:
                return col
    return defecto

col_num_viaje = buscar_columna(["# de viaje", "num_viaje", "num", "núm", "numero", "número", "id_viaje", "viaje"], df.columns, df.columns[0])
col_fecha = buscar_columna(["fecha", "date"], df.columns, None)
col_cliente = buscar_columna(["cliente", "empresa", "razon_social"], df.columns, "Cliente")
col_flete = buscar_columna(["valor de flete", "flete", "valor", "monto", "precio", "tarifa"], df.columns, None)
col_obs = buscar_columna(["observaciones", "observacion", "obs", "nota", "comentario"], df.columns, None)
col_tipo_viaje = buscar_columna(["tipo viaje", "tipo_viaje", "tipo", "mercancia", "servicio"], df.columns, "Tipo Viaje")
col_completado = buscar_columna(["completado", "completo", "estado", "estatus", "status"], df.columns, None)
col_motorista = buscar_columna(["motorista", "conductor", "chofer"], df.columns, "Motorista")

# Limpieza y conversión obligatoria de la columna Flete
if col_flete and col_flete in df.columns:
    df[col_flete] = (
        df[col_flete]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df[col_flete] = pd.to_numeric(df[col_flete], errors="coerce").fillna(0.0)
else:
    df["Valor de Flete"] = 0.0
    col_flete = "Valor de Flete"

if col_fecha and col_fecha in df.columns:
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
    df["Año-Mes"] = df[col_fecha].dt.to_period("M").astype(str)
else:
    df["Año-Mes"] = "Sin Fecha"

# Asegurar columna Observaciones
if not col_obs or col_obs not in df.columns:
    df["Observaciones"] = ""
    col_obs = "Observaciones"
else:
    df[col_obs] = df[col_obs].fillna("")

# Pestañas principales de navegación
tab_reporte, tab_dashboard, tab_general = st.tabs([
    "📑 Reporte de Viajes Pendientes",
    "📊 Dashboard y Métricas",
    "🔍 Consulta General"
])

# --- PESTAÑA 1: REPORTE DE VIAJES PENDIENTES DE FACTURAR ---
with tab_reporte:
    # Formatear fecha actual en español para el título del reporte
    meses_es = {
        "JANUARY": "ENERO", "FEBRUARY": "FEBRERO", "MARCH
