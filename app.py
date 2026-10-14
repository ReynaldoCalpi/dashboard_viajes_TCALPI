import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Configuración de la página
st.set_page_config(
    page_title="Dashboard de Control de Viajes y Fletes",
    page_icon="🚚",
    layout="wide"
)

st.title("🚚 Dashboard de Control de Viajes y Fletes")
st.markdown("Carga tu archivo de Excel para analizar indicadores, filtrar rutas y descargar reportes por mes, motorista y destino.")

# Carga de archivo
st.sidebar.header("📁 Cargar Archivo de Excel")
uploaded_file = st.sidebar.file_uploader("Selecciona tu archivo (.xlsx o .xls)", type=["xlsx", "xls", "csv"])

# Datos de muestra en caso de no subir archivo
def cargar_datos_muestra():
    data = {
        "Fecha": pd.date_range(start="2026-08-24", periods=18, freq="D"),
        "# de Viaje": [82373, 83637, 83884, 83889, 83915, 83920, 83927, 83932, 83941, 83944, 83946, 83967, 83968, 83978, 83979, 83980, 83981, 83982],
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
            "HARISA, S.A. DE C.V.", 
            "PRODUCTOS ALIMENTICIOS BOCADELI, S.A DE C.V", 
            "AVICOLA SALVADOREÑA, S.A. DE C.V.", 
            "UNIFERSA-DISAGRO, S.A. DE C.V.", 
            "UNIFERSA-DISAGRO, S.A. DE C.V.", 
            "UNIFERSA-DISAGRO, S.A. DE C.V.", 
            "UNIFERSA-DISAGRO, S.A. DE C.V.", 
            "UNIFERSA-DISAGRO, S.A. DE C.V."
        ],
        "Valor Flete": [0, 950.00, 1650.00, 1650.00, 700.00, 700.00, 1203.75, 205.00, 1458.00, 827.50, 1203.75, 800.00, 0, 1187.50, 1187.50, 1187.50, 1187.50, 1187.50],
        "Observaciones": ["", "", "", "", "", "", "CON PROFORMA", "", "", "", "CON PROFORMA", "", "", "", "", "", "", ""],
        "Tipo Viaje": [
            "VIAJE LOCAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", 
            "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE LOCAL", 
            "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", 
            "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL", 
            "VIAJE INTERNACIONAL", "VIAJE INTERNACIONAL"
        ],
        "Motorista": ["Juan Pérez", "Carlos Gómez", "Roberto Rivas", "Mario Cruz", "Ana López"] * 4,
        "Completado": ["Pendiente", "Completado", "Pendiente", "Completado", "Pendiente", "Completado", "Pendiente", "Completado", "Pendiente", "Completado", "Pendiente", "Completado", "Pendiente", "Completado", "Pendiente", "Completado", "Pendiente", "Completado"]
    }
    for k in data:
        data[k] = data[k][:18]
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
    st.sidebar.info("📌 Viendo datos de demostración basados en el reporte solicitado. Carga tu Excel para analizar datos reales.")
    df = cargar_datos_muestra()

# Detección y normalización de columnas clave
col_fecha = next((c for c in df.columns if "fecha" in str(c).lower() or "date" in str(c).lower()), None)
col_motorista = next((c for c in df.columns if "motorista" in str(c).lower() or "conductor" in str(c).lower()), df.columns[1] if len(df.columns) > 1 else df.columns[0])
col_flete = next((c for c in df.columns if "flete" in str(c).lower() or "valor" in str(c).lower() or "monto" in str(c).lower()), None)
col_completado = next((c for c in df.columns if "completado" in str(c).lower() or "estado" in str(c).lower()), None)

if col_fecha:
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
    df["Año-Mes"] = df[col_fecha].dt.to_period("M").astype(str)
else:
    df["Año-Mes"] = "Sin Fecha"

# Limpieza de la columna de flete para asegurar formato numérico
if col_flete:
    df[col_flete] = (
        df[col_flete]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df[col_flete] = pd.to_numeric(df[col_flete], errors="coerce").fillna(0)

# Filtros Laterales (Seguros contra errores de tipos de datos)
st.sidebar.header("🔍 Filtros Generales")
meses = sorted([str(x) for x in df["Año-Mes"].dropna().unique()])
selected_meses = st.sidebar.multiselect("Filtrar por Mes:", meses, default=meses)

motoristas = sorted([str(x) for x in df[col_motorista].dropna().unique()])
selected_motoristas = st.sidebar.multiselect("Filtrar por Motorista:", motoristas, default=motoristas)

# Selector interactivo para la columna "Completado" / Estado
st.sidebar.header("✅ Selector de Estado")
if col_completado:
    estados_disponibles = sorted([str(x) for x in df[col_completado].dropna().unique()])
    selected_estados = st.sidebar.multiselect("Seleccionar Estado (Completado):", estados_disponibles, default=estados_disponibles)
else:
    col_completado = next((c for c in df.columns if "tipo" in str(c).lower()), df.columns[-1])
    estados_disponibles = sorted([str(x) for x in df[col_completado].dropna().unique()])
    selected_estados = st.sidebar.multiselect("Seleccionar Estado/Tipo:", estados_disponibles, default=estados_disponibles)

# Aplicación de Filtros
df_filtered = df[
    (df["Año-Mes"].isin(selected_meses)) &
    (df[col_motorista].astype(str).isin(selected_motoristas)) &
    (df[col_completado].astype(str).isin(selected_estados))
]

# Indicadores Principales (KPIs)
st.markdown("### 📊 Indicadores Clave (KPIs)")
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

total_viajes = len(df_filtered)
total_flete = df_filtered[col_flete].sum() if col_flete else 0
flete_promedio = df_filtered[col_flete].mean() if col_flete and total_viajes > 0 else 0
num_motoristas = df_filtered[col_motorista].nunique()

kpi1.metric("Total Viajes", f"{total_viajes:,}")
kpi2.metric("Valor Total Flete", f"${total_flete:,.2f}")
kpi3.metric("Flete Promedio", f"${flete_promedio:,.2f}")
kpi4.metric("Motoristas Activos", f"{num_motoristas}")

st.markdown("---")

# Visualizaciones
st.markdown("### 📈 Análisis de Viajes y Fletes")
col_g1, col_g2 = st.columns(2)

with col_g1:
    if col_flete:
        resumen_mes = df_filtered.groupby("Año-Mes")[col_flete].sum().reset_index()
        fig_mes = px.bar(
            resumen_mes, x="Año-Mes", y=col_flete,
            text_auto="$.2s", title="<b>Flete Total por Mes</b>",
            labels={col_flete: "Monto Flete ($)", "Año-Mes": "Mes"},
            color_discrete_sequence=["#1f77b4"]
        )
        st.plotly_chart(fig_mes, use_container_width=True)

with col_g2:
    if col_flete:
        resumen_moto = df_filtered.groupby(col_motorista)[col_flete].sum().reset_index().sort_values(by=col_flete, ascending=False)
        fig_moto = px.bar(
            resumen_moto, x=col_motorista, y=col_flete,
            text_auto="$.2s", title="<b>Flete Generado por Motorista</b>",
            labels={col_flete: "Monto Flete ($)", col_motorista: "Motorista"},
            color_discrete_sequence=["#2ca02c"]
        )
        st.plotly_chart(fig_moto, use_container_width=True)

st.markdown("---")

# Reporte Detallado Estilo Imagen de Referencia
st.markdown("### 📋 REPORTE DE VIAJES PENDIENTES DE FACTURAR")
st.markdown("Visualización ajustada con formato tabular limpio de columnas clave (# de Viaje, Fecha, Cliente, Valor de Flete, Observaciones y Tipo Viaje).")

columnas_reporte = [c for c in ["# de Viaje", "Fecha", "Cliente", "Valor Flete", "Observaciones", "Tipo Viaje"] if c in df_filtered.columns]
if not columnas_reporte:
    columnas_reporte = df_filtered.columns

st.dataframe(df_filtered[columnas_reporte], use_container_width=True)

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
    file_name="reporte_viajes_filtrado.csv",
    mime="text/csv"
)

c_down2.download_button(
    label="📥 Descargar Reporte en Excel (.xlsx)",
    data=convert_to_excel(df_filtered),
    file_name="reporte_viajes_filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
