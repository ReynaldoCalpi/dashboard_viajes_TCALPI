import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Configuración de la página
st.set_page_config(
    page_title="Dashboard de Control de Viajes y Fletes - Transportes Calpi",
    page_icon="🚚",
    layout="wide"
)

st.title("🚚 Dashboard de Control de Viajes y Fletes - Transportes Calpi")
st.markdown("Análisis oficial de indicadores, filtrado flexible por estado, mes, motorista y ruta.")

# Carga de archivo
st.sidebar.header("📁 Archivo de Excel")
uploaded_file = st.sidebar.file_uploader("Sube tu archivo de Excel", type=["xlsx", "xls", "csv"])

# Función para cargar datos por defecto o el archivo subido
@st.cache_data
def cargar_datos(file):
    if file is not None:
        try:
            if file.name.endswith(".csv"):
                return pd.read_csv(file)
            else:
                xls = pd.ExcelFile(file)
                sheet_name = '1' if '1' in xls.sheet_names else xls.sheet_names[0]
                return pd.read_excel(file, sheet_name=sheet_name)
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")
            return None
    else:
        # Datos de muestra basados en la estructura real si no se sube archivo
        data = {
            "Viaje": [82373, 83637, 83884, 83889, 83915, 83920, 83927, 83932, 83941, 83944],
            "Completado": ["ASIGNADO", "COMPLETADO", "ASIGNADO", "COMPLETADO", "ASIGNADO", "COMPLETADO", "ASIGNADO", "COMPLETADO", "ASIGNADO", "COMPLETADO"],
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
                "MOLINOS MODERNOS, S.A."
            ],
            "Fecha": pd.date_range(start="2026-08-24", periods=10, freq="D"),
            "Valor de Flete": [0, 950.00, 1650.00, 1650.00, 700.00, 700.00, 1203.75, 205.00, 1458.00, 827.50],
            "Motorista": ["Juan Pérez", "Carlos Gómez", "Roberto Rivas", "Mario Cruz", "Ana López"] * 2,
            "Tipo Viaje": ["VIAJE LOCAL", "VIAJE INTERNACIONAL"] * 5
        }
        return pd.DataFrame(data)

df = cargar_datos(uploaded_file)

if df is None:
    st.stop()

if uploaded_file is None:
    st.sidebar.info("📌 Viendo datos de demostración. Sube tu archivo Excel oficial para analizar la información real.")
else:
    st.sidebar.success("¡Archivo cargado con éxito!")

# Normalización de nombres de columnas exactas
col_viaje = next((c for c in df.columns if "viaje" in str(c).lower() and "tipo" not in str(c).lower()), df.columns[0])
col_estado = next((c for c in df.columns if "completado" in str(c).lower() or "estado" in str(c).lower()), "Completado")
col_cliente = next((c for c in df.columns if "cliente" in str(c).lower()), "Cliente")
col_fecha = next((c for c in df.columns if "fecha" in str(c).lower() and "descarga" not in str(c).lower()), "Fecha")
col_flete = next((c for c in df.columns if "flete" in str(c).lower() or "valor" in str(c).lower()), "Valor de Flete")
col_motorista = next((c for c in df.columns if "motorista" in str(c).lower() or "conductor" in str(c).lower()), "Motorista")
col_tipoviaje = next((c for c in df.columns if "tipo viaje" in str(c).lower()), "Tipo Viaje")

# Procesamiento de fechas y meses
if col_fecha in df.columns:
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
    df["Año-Mes"] = df[col_fecha].dt.to_period("M").astype(str)
else:
    df["Año-Mes"] = "Sin Fecha"

# Limpieza de la columna de flete a formato numérico
if col_flete in df.columns:
    df[col_flete] = (
        df[col_flete]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df[col_flete] = pd.to_numeric(df[col_flete], errors="coerce").fillna(0)

# --- FILTROS LATERALES MEJORADOS ---
st.sidebar.header("🔍 Filtros Operativos")

# 1. Filtro por Estado (Completado / Asignado / Anulado)
if col_estado in df.columns:
    estados_disponibles = sorted([str(x) for x in df[col_estado].dropna().unique()])
    selected_estados = st.sidebar.multiselect("Filtrar por Estado:", estados_disponibles, default=estados_disponibles)
else:
    selected_estados = []

# 2. Filtro por Mes (con opción general)
meses = sorted([str(x) for x in df["Año-Mes"].dropna().unique()])
st.sidebar.markdown("---")
st.sidebar.subheader("📅 Filtro por Mes")
todos_meses = st.sidebar.checkbox("Seleccionar Todos los Meses", value=True)
if todos_meses:
    selected_meses = meses
    st.sidebar.multiselect("Meses activos:", meses, default=meses, disabled=True)
else:
    selected_meses = st.sidebar.multiselect("Elige los Meses:", meses, default=meses[:1] if meses else [])

# 3. Filtro por Motorista (con opción general)
if col_motorista in df.columns:
    motoristas = sorted([str(x) for x in df[col_motorista].dropna().unique()])
    st.sidebar.markdown("---")
    st.sidebar.subheader("👨‍✈️ Filtro por Motorista")
    todos_motoristas = st.sidebar.checkbox("Seleccionar Todos los Motoristas", value=True)
    if todos_motoristas:
        selected_motoristas = motoristas
        st.sidebar.multiselect("Motoristas activos:", motoristas, default=motoristas, disabled=True)
    else:
        selected_motoristas = st.sidebar.multiselect("Elige los Motoristas:", motoristas, default=[])
else:
    selected_motoristas = []

# Aplicación de filtros al DataFrame
df_filtered = df.copy()
if col_estado in df.columns and selected_estados:
    df_filtered = df_filtered[df_filtered[col_estado].astype(str).isin(selected_estados)]

if selected_meses:
    df_filtered = df_filtered[df_filtered["Año-Mes"].isin(selected_meses)]

if col_motorista in df.columns and selected_motoristas:
    df_filtered = df_filtered[df_filtered[col_motorista].astype(str).isin(selected_motoristas)]

# --- KPIs ---
st.markdown("### 📊 Indicadores Clave (KPIs)")
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

total_viajes = len(df_filtered)
total_flete = df_filtered[col_flete].sum() if col_flete in df_filtered.columns else 0
flete_promedio = df_filtered[col_flete].mean() if col_flete in df_filtered.columns and total_viajes > 0 else 0
num_motoristas = df_filtered[col_motorista].nunique() if col_motorista in df_filtered.columns else 0

kpi1.metric("Total Viajes", f"{total_viajes:,}")
kpi2.metric("Valor Total Flete", f"${total_flete:,.2f}")
kpi3.metric("Flete Promedio", f"${flete_promedio:,.2f}")
kpi4.metric("Motoristas Activos", f"{num_motoristas}")

st.markdown("---")

# --- GRÁFICOS ---
st.markdown("### 📈 Análisis Gráfico")
col_g1, col_g2 = st.columns(2)

with col_g1:
    if col_flete in df_filtered.columns:
        resumen_mes = df_filtered.groupby("Año-Mes")[col_flete].sum().reset_index()
        fig_mes = px.bar(
            resumen_mes, x="Año-Mes", y=col_flete,
            text_auto="$.2s", title="<b>Flete Total por Mes</b>",
            labels={col_flete: "Monto Flete ($)", "Año-Mes": "Mes"},
            color_discrete_sequence=["#1f77b4"]
        )
        st.plotly_chart(fig_mes, use_container_width=True)

with col_g2:
    if col_motorista in df_filtered.columns and col_flete in df_filtered.columns:
        resumen_moto = df_filtered.groupby(col_motorista)[col_flete].sum().reset_index().sort_values(by=col_flete, ascending=False).head(10)
        fig_moto = px.bar(
            resumen_moto, x=col_motorista, y=col_flete,
            text_auto="$.2s", title="<b>Top Flete Generado por Motorista</b>",
            labels={col_flete: "Monto Flete ($)", col_motorista: "Motorista"},
            color_discrete_sequence=["#2ca02c"]
        )
        st.plotly_chart(fig_moto, use_container_width=True)

st.markdown("---")

# --- REPORTE DETALLADO ---
st.markdown("### 📋 Reporte Detallado de Viajes")
st.dataframe(df_filtered, use_container_width=True)

# --- DESCARGAS ---
@st.cache_data
def convert_to_csv(data):
    return data.to_csv(index=False).encode('utf-8')

@st.cache_data
def convert_to_excel(data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data.to_excel(writer, index=False, sheet_name='Reporte Viajes')
    return output.getvalue()

c_down1, c_down2 = st.columns(2)
c_down1.download_button(
    label="📥 Descargar Reporte en CSV",
    data=convert_to_csv(df_filtered),
    file_name="reporte_viajes_calpi.csv",
    mime="text/csv"
)

c_down2.download_button(
    label="📥 Descargar Reporte en Excel (.xlsx)",
    data=convert_to_excel(df_filtered),
    file_name="reporte_viajes_calpi.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
