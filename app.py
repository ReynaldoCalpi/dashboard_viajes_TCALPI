import streamlit as st
import pandas as pd
import plotly.express as px
import io

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
        "Número de Viaje": [f"VJN-{1000 + i}" for i in range(120)],
        "Fecha": pd.date_range(start="2026-01-01", periods=120, freq="D"),
        "Cliente": ["Distribuidora El Sol", "Industrias Lemus", "Comercializadora Alfa", "AgroServicios SV", "Logística Central"] * 24,
        "Tipo Mercancia": ["Granel", "General", "Refrigerada", "Granel", "Especial"] * 24,
        "Completado": ["ASIGNADO", "FACTURADO", "ASIGNADO", "COMPLETADO", "FACTURADO", "ASIGNADO"] * 20,
        "Motorista": ["Juan Pérez", "Carlos Gómez", "Roberto Rivas", "Mario Cruz", "Ana López"] * 24,
        "Lugar Embarque": ["Bodega Apopa", "Terminal Acajutla", "Depósito Norte", "Planta Soyapango"] * 30,
        "Ciudad Origen": ["San Salvador", "Apopa", "Santa Ana", "San Miguel", "Sonsonate"] * 24,
        "Lugar Destino": ["Puerto de Acajutla", "Metapán", "La Libertad", "Ahuachapán", "Usulután"] * 24,
        "Valor de Flete": [350, 420, 500, 280, 600, 450, 380, 520, 310, 490] * 12
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

# Detección inteligente de columnas
def buscar_columna(patrones, df_cols, defecto=None):
    for col in df_cols:
        col_clean = str(col).lower().strip()
        for pat in patrones:
            if pat in col_clean:
                return col
    return defecto

col_num_viaje = buscar_columna(["num", "núm", "numero", "número", "id_viaje", "vjn", "viaje"], df.columns, df.columns[0])
col_fecha = buscar_columna(["fecha", "date"], df.columns, None)
col_cliente = buscar_columna(["cliente", "empresa", "razon_social"], df.columns, "Cliente")
col_tipo_viaje = buscar_columna(["tipo", "tipo_viaje", "mercancia", "mercancía", "servicio"], df.columns, "Tipo Mercancia")
col_completado = buscar_columna(["completado", "completo", "estado", "estatus", "status"], df.columns, "Completado")
col_motorista = buscar_columna(["motorista", "conductor", "chofer"], df.columns, "Motorista")
col_embarque = buscar_columna(["embarque", "lugar_embarque", "origen_embarque"], df.columns, "Lugar Embarque")
col_origen = buscar_columna(["origen", "ciudad_origen", "procedencia"], df.columns, "Ciudad Origen")
col_destino = buscar_columna(["destino", "lugar_destino", "ciudad_destino"], df.columns, "Lugar Destino")
col_flete = buscar_columna(["flete", "valor", "monto", "precio", "tarifa"], df.columns, None)

# === CORRECCIÓN DEL ERROR DE SUMA (TypeError) ===
# Limpieza y conversión obligatoria de la columna de Valor de Flete a número flotante
if col_flete and col_flete in df.columns:
    df[col_flete] = (
        df[col_flete]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df[col_flete] = pd.to_numeric(df[col_flete], errors="coerce").fillna(0.0)

# Normalización de Fechas y Meses
if col_fecha and col_fecha in df.columns:
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
    df["Año-Mes"] = df[col_fecha].dt.to_period("M").astype(str)
else:
    df["Año-Mes"] = "Sin Fecha"

# Funciones auxiliares para descargas
@st.cache_data
def convert_to_csv(data):
    return data.to_csv(index=False).encode('utf-8')

@st.cache_data
def convert_to_excel(data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

# Estructura principal por pestañas
tab_pendientes, tab_dashboard, tab_consulta = st.tabs([
    "📑 Viajes Pendientes de Facturar",
    "📊 Dashboard y Métricas",
    "🔍 Consulta General de Viajes"
])

# ==========================================
# PESTAÑA 1: VIAJES PENDIENTES DE FACTURAR
# ==========================================
with tab_pendientes:
    st.header("📋 Sección de Viajes Pendientes de Facturar")
    st.markdown("Filtrado de viajes según la columna **Completado** (Estado **'ASIGNADO'** u otro valor pendiente).")

    if col_completado in df.columns:
        valores_disponibles = sorted(df[col_completado].dropna().astype(str).unique())
        
        # Seleccionar valor por defecto que contenga 'asignad'
        val_default = [v for v in valores_disponibles if "asignad" in v.lower()]
        if not val_default:
            val_default = [valores_disponibles[0]] if valores_disponibles else []

        filtro_estado = st.multiselect(
            f"Filtrar por estado en la columna '{col_completado}':",
            options=valores_disponibles,
            default=val_default,
            key="filtro_completado_pendientes"
        )
        
        df_asignadas = df[df[col_completado].astype(str).isin(filtro_estado)] if filtro_estado else df.copy()
    else:
        df_asignadas = pd.DataFrame()
        st.warning("⚠️ No se encontró la columna 'Completado' en el archivo cargado.")

    # Filtros adicionales para pendientes
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        if col_cliente in df_asignadas.columns and not df_asignadas.empty:
            clientes_p = sorted(df_asignadas[col_cliente].dropna().astype(str).unique())
            sel_cliente_p = st.multiselect("Filtrar por Cliente:", clientes_p, default=clientes_p, key="p_cliente")
            if sel_cliente_p:
                df_asignadas = df_asignadas[df_asignadas[col_cliente].astype(str).isin(sel_cliente_p)]
    
    with col_f2:
        if col_tipo_viaje in df_asignadas.columns and not df_asignadas.empty:
            tipos_p = sorted(df_asignadas[col_tipo_viaje].dropna().astype(str).unique())
            sel_tipo_p = st.multiselect("Filtrar por Tipo de Viaje:", tipos_p, default=tipos_p, key="p_tipo")
            if sel_tipo_p:
                df_asignadas = df_asignadas[df_asignadas[col_tipo_viaje].astype(str).isin(sel_tipo_p)]

    # Métricas clave de pendientes
    p_kpi1, p_kpi2, p_kpi3 = st.columns(3)
    cant_pendientes = len(df_asignadas)
    monto_pendiente = float(df_asignadas[col_flete].sum()) if col_flete and col_flete in df_asignadas.columns else 0.0
    
    p_kpi1.metric("Pendientes de Facturar", f"{cant_pendientes} Viajes")
    p_kpi2.metric("Monto Total Pendiente", f"${monto_pendiente:,.2f}")
    p_kpi3.metric("Columna Utilizada", f"{col_completado}")

    st.markdown("#### Lista de Viajes Asignados / Pendientes para Facturación")

    # Selección explícita de columnas requeridas
    cols_requeridas = [
        c for c in [col_completado, col_num_viaje, col_fecha, col_cliente, col_tipo_viaje, col_flete, col_motorista, col_embarque, col_origen, col_destino]
        if c and c in df_asignadas.columns
    ]

    if not df_asignadas.empty:
        df_vista_pendientes = df_asignadas[cols_requeridas]
        st.dataframe(df_vista_pendientes, use_container_width=True, hide_index=True)

        st.markdown("##### 📥 Exportar Registro de Viajes Pendientes")
        btn_c1, btn_c2 = st.columns(2)
        btn_c1.download_button(
            label="📥 Descargar Pendientes (Excel)",
            data=convert_to_excel(df_vista_pendientes),
            file_name="viajes_pendientes_por_facturar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="down_p_excel"
        )
        btn_c2.download_button(
            label="📥 Descargar Pendientes (CSV)",
            data=convert_to_csv(df_vista_pendientes),
            file_name="viajes_pendientes_por_facturar.csv",
            mime="text/csv",
            key="down_p_csv"
        )
    else:
        st.info("ℹ️ No se encontraron viajes para los filtros seleccionados.")

# ==========================================
# PESTAÑA 2: DASHBOARD Y MÉTRICAS GENERALES
# ==========================================
with tab_dashboard:
    st.header("📊 Dashboard Estadístico General")

    st.subheader("🔍 Filtros para Métricas")
    f_c1, f_c2, f_c3 = st.columns(3)
    
    with f_c1:
        meses_disp = sorted(df["Año-Mes"].dropna().unique())
        sel_meses_d = st.multiselect("Seleccionar Meses:", meses_disp, default=meses_disp, key="dash_mes")
    
    with f_c2:
        if col_motorista in df.columns:
            motos_disp = sorted(df[col_motorista].dropna().astype(str).unique())
            sel_motos_d = st.multiselect("Seleccionar Motoristas:", motos_disp, default=motos_disp, key="dash_moto")
        else:
            sel_motos_d = []

    with f_c3:
        if col_completado in df.columns:
            est_disp = sorted(df[col_completado].dropna().astype(str).unique())
            sel_est_d = st.multiselect("Seleccionar Estatus (Completado):", est_disp, default=est_disp, key="dash_est")
        else:
            sel_est_d = []

    # Filtrar DF para dashboard
    df_dash = df.copy()
    if sel_meses_d:
        df_dash = df_dash[df_dash["Año-Mes"].isin(sel_meses_d)]
    if sel_motos_d and col_motorista in df_dash.columns:
        df_dash = df_dash[df_dash[col_motorista].astype(str).isin(sel_motos_d)]
    if sel_est_d and col_completado in df_dash.columns:
        df_dash = df_dash[df_dash[col_completado].astype(str).isin(sel_est_d)]

    # KPIs Dashboard
    d_kpi1, d_kpi2, d_kpi3, d_kpi4 = st.columns(4)
    tot_v = len(df_dash)
    tot_f = float(df_dash[col_flete].sum()) if col_flete and col_flete in df_dash.columns else 0.0
    prom_f = float(df_dash[col_flete].mean()) if col_flete and col_flete in df_dash.columns and tot_v > 0 else 0.0
    act_m = df_dash[col_motorista].nunique() if col_motorista in df_dash.columns else 0

    d_kpi1.metric("Total Viajes Filtrados", f"{tot_v:,}")
    d_kpi2.metric("Valor Total Flete", f"${tot_f:,.2f}")
    d_kpi3.metric("Flete Promedio / Viaje", f"${prom_f:,.2f}")
    d_kpi4.metric("Motoristas Activos", f"{act_m}")

    st.markdown("---")

    # Gráficos
    g_col1, g_col2 = st.columns(2)
    
    with g_col1:
        if col_flete and col_flete in df_dash.columns:
            res_mes = df_dash.groupby("Año-Mes")[col_flete].sum().reset_index()
            fig_mes = px.bar(
                res_mes, x="Año-Mes", y=col_flete, text_auto="$.2s",
                title="<b>Monto Total de Flete por Mes</b>",
                labels={col_flete: "Monto Flete ($)", "Año-Mes": "Mes"},
                color_discrete_sequence=["#2b5c8f"]
            )
            st.plotly_chart(fig_mes, use_container_width=True)

    with g_col2:
        if col_flete and col_motorista in df_dash.columns:
            res_moto = df_dash.groupby(col_motorista)[col_flete].sum().reset_index().sort_values(by=col_flete, ascending=False).head(10)
            fig_moto = px.bar(
                res_moto, x=col_motorista, y=col_flete, text_auto="$.2s",
                title="<b>Top 10 Motoristas por Valor de Flete</b>",
                labels={col_flete: "Monto Flete ($)", col_motorista: "Motorista"},
                color_discrete_sequence=["#2ca02c"]
            )
            st.plotly_chart(fig_moto, use_container_width=True)

# ==========================================
# PESTAÑA 3: CONSULTA GENERAL Y DESCARGAS
# ==========================================
with tab_consulta:
    st.header("🔍 Tabla General de Todos los Viajes")
    st.markdown("Visualización completa del archivo con opción de descarga de datos filtrados.")

    st.dataframe(df, use_container_width=True)

    st.markdown("##### 📥 Exportar Todos los Datos")
    c_d1, c_d2 = st.columns(2)
    c_d1.download_button(
        label="📥 Descargar Todo en Excel (.xlsx)",
        data=convert_to_excel(df),
        file_name="consolidado_general_viajes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="down_gen_excel"
    )
    c_d2.download_button(
        label="📥 Descargar Todo en CSV (.csv)",
        data=convert_to_csv(df),
        file_name="consolidado_general_viajes.csv",
        mime="text/csv",
        key="down_gen_csv"
    )
