import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Gestión y Control de Viajes | Dashboard",
    page_icon="🚚",
    layout="wide"
)

st.title("🚚 Sistema de Control de Viajes, Fletes y Facturación")
st.markdown("Carga tu archivo de Excel para gestionar viajes pendientes de facturar, analizar rutas y descargar reportes organizados.")

# Sidebar - Carga de Archivo
st.sidebar.header("📁 Cargar Archivo")
uploaded_file = st.sidebar.file_uploader("Selecciona tu archivo (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])

# Datos de muestra adaptados
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
        "Completado": ["ASIGNADO"] * 11,
        "Valor de Flete": [0, 950.00, 1650.00, 1650.00, 700.00, 700.00, 1203.75, 205.00, 1458.00, 827.50, 1203.75],
        "Observaciones": ["", "", "", "", "", "", "CON PROFORMA", "", "", "", "CON PROFORMA"],
        "Tipo Viaje": ["VIAJE LOCAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE LOCAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL"],
        "Motorista": ["FERMIN DE JESUS GRACIAS CASTRO", "ROBERTO CARLOS DELGADO NERIO", "JUAN PEREZ", "JUAN PEREZ", "CARLOS GOMEZ", "CARLOS GOMEZ", "MARIO CRUZ", "ROBERTO RIVAS", "ANA LOPEZ", "MARIO CRUZ", "MARIO CRUZ"],
        "Lugar Embarque": ["El Salvador", "Honduras", "Guatemala", "Guatemala", "El Salvador", "El Salvador", "Guatemala", "El Salvador", "Nicaragua", "El Salvador", "Guatemala"],
        "Ciudad Origen": ["Apopa", "Choloma", "Ciudad de Guatemala", "Ciudad de Guatemala", "Apopa", "Apopa", "Ciudad de Guatemala", "Apopa", "Managua", "Apopa", "Ciudad de Guatemala"],
        "Lugar Destino": ["El Salvador", "El Salvador", "El Salvador", "El Salvador", "Honduras", "Honduras", "El Salvador", "El Salvador", "El Salvador", "Honduras", "El Salvador"]
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
    st.sidebar.info("📌 Mostrando datos de demostración. Sube tu archivo Excel para analizar tu información real.")
    df = cargar_datos_muestra()

# Función de búsqueda dinámica de columnas
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
col_completado = buscar_columna(["completado", "completo", "estado", "estatus", "status"], df.columns, "Completado")
col_motorista = buscar_columna(["motorista", "conductor", "chofer"], df.columns, "Motorista")
col_embarque = buscar_columna(["embarque", "lugar_embarque", "origen_embarque"], df.columns, "Lugar Embarque")
col_origen = buscar_columna(["origen", "ciudad_origen", "procedencia"], df.columns, "Ciudad Origen")
col_destino = buscar_columna(["destino", "lugar_destino", "ciudad_destino"], df.columns, "Lugar Destino")

# Limpieza y conversión obligatoria del Flete
if col_flete and col_flete in df.columns:
    df[col_flete] = (
        df[col_flete]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df[col_flete] = pd.to_numeric(df[col_flete], errors="coerce").fillna(0.0)

# Normalización de Fechas
if col_fecha and col_fecha in df.columns:
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
    df["Año-Mes"] = df[col_fecha].dt.to_period("M").astype(str)
else:
    df["Año-Mes"] = "Sin Fecha"

# Asegurar columna de Observaciones
if not col_obs or col_obs not in df.columns:
    df["Observaciones"] = ""
    col_obs = "Observaciones"
else:
    df[col_obs] = df[col_obs].fillna("")

# Funciones de Exportación a Excel con encabezado idéntico al modelo
def generar_excel_reporte_oficial(df_export, titulo_reporte):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name
