import streamlit as st
import pandas as pd
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# Diccionario para convertir números de mes a nombres en español
nombres_meses = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}

# Función para cargar datos por defecto o el archivo subido
@st.cache_data
def cargar_datos(file):
    if file is not None:
        try:
            if file.name.endswith(".csv"):
                return pd.read_csv(file)
            else:
                xls = pd.ExcelFile(file)
                # Priorizamos la hoja 2 si existe y está estructurada, o la hoja 1
                sheet_name = '2' if '2' in xls.sheet_names else ('1' if '1' in xls.sheet_names else xls.sheet_names[0])
                if sheet_name == '2':
                    return pd.read_excel(file, sheet_name=sheet_name, header=4)
                else:
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
            "Fecha": pd.date_range(start="2026-10-01", periods=10, freq="D"),
            "Valor de Flete": [0, 950.00, 1650.00, 1650.00, 700.00, 700.00, 1203.75, 205.00, 1458.00, 827.50],
            "Observaciones": ["", "", "", "CON PROFORMA", "", "", "", "", "", ""],
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

# Búsqueda robusta de la columna de flete
col_flete = None
for c in df.columns:
    if any(k in str(c).lower() for k in ["flete", "valor", "monto", "total"]):
        col_flete = c
        break
if not col_flete:
    col_flete = df.columns[-1]

col_observaciones = next((c for c in df.columns if "observacion" in str(c).lower()), "Observaciones")
col_motorista = next((c for c in df.columns if "motorista" in str(c).lower() or "conductor" in str(c).lower()), "Motorista")
col_tipoviaje = next((c for c in df.columns if "tipo viaje" in str(c).lower()), "Tipo Viaje")

# Limpieza estricta e inmediata de la columna de flete a formato numérico real
if col_flete in df.columns:
    df[col_flete] = (
        df[col_flete]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df[col_flete] = pd.to_numeric(df[col_flete], errors="coerce").fillna(0)

# Procesamiento de fechas y conversión a formato legible en español (ej. "Octubre 2026")
if col_fecha in df.columns:
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
    
    def formatear_mes_espanol(dt):
        if pd.isna(dt):
            return "Sin Fecha"
        anio = dt.strftime("%Y")
        num_mes = dt.strftime("%m")
        nombre_mes = nombres_meses.get(num_mes, num_mes)
        return f"{nombre_mes} {anio}"

    df["Mes-Anio"] = df[col_fecha].apply(formatear_mes_espanol)
else:
    df["Mes-Anio"] = "Sin Fecha"

# --- FILTROS LATERALES ---
st.sidebar.header("🔍 Filtros Operativos")

if col_estado in df.columns:
    estados_disponibles = sorted([str(x) for x in df[col_estado].dropna().unique()])
    selected_estados = st.sidebar.multiselect("Filtrar por Estado:", estados_disponibles, default=estados_disponibles)
else:
    selected_estados = []

meses = sorted([str(x) for x in df["Mes-Anio"].dropna().unique()])
st.sidebar.markdown("---")
st.sidebar.subheader("📅 Filtro por Mes")
todos_meses = st.sidebar.checkbox("Seleccionar Todos los Meses", value=True)
if todos_meses:
    selected_meses = meses
    st.sidebar.multiselect("Meses activos:", meses, default=meses, disabled=True)
else:
    selected_meses = st.sidebar.multiselect("Elige los Meses:", meses, default=meses[:1] if meses else [])

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
    df_filtered = df_filtered[df_filtered["Mes-Anio"].isin(selected_meses)]

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
        resumen_mes = df_filtered.groupby("Mes-Anio")[col_flete].sum().reset_index()
        fig_mes = px.bar(
            resumen_mes, x="Mes-Anio", y=col_flete,
            text_auto="$.2s", title="<b>Flete Total por Mes</b>",
            labels={col_flete: "Monto Flete ($)", "Mes-Anio": "Mes"},
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

# --- REPORTE DETALLADO (Vista en Pantalla) ---
st.markdown("### 📋 Reporte Detallado de Viajes")

cols_visibles_map = {
    col_viaje: "# de Viaje",
    col_fecha: "Fecha",
    col_cliente: "Cliente",
    col_flete: "Valor de Flete",
    col_observaciones if col_observaciones in df_filtered.columns else None: "Observaciones",
    col_tipoviaje if col_tipoviaje in df_filtered.columns else None: "Tipo Viaje"
}
cols_visibles_keys = [k for k in cols_visibles_map.keys() if k is not None and k in df_filtered.columns]
df_view = df_filtered[cols_visibles_keys].copy()
df_view.columns = [cols_visibles_map[c] for c in cols_visibles_keys]

df_view.insert(0, "No.", range(1, len(df_view) + 1))
st.dataframe(df_view, use_container_width=True)

# --- GENERADOR DE REPORTE EXCEL CON DISEÑO PROFESIONAL ---
def generar_excel_calpi(data_filtered):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte Viajes"
    
    ws.views.sheetView[0].showGridLines = True
    
    font_titulo = Font(name="Calibri", size=14, bold=True, color="000000")
    font_sub = Font(name="Calibri", size=11, bold=True, color="000000")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="000000")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")
    
    fill_header = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
    fill_total = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
    
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    border_total = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )

    ws.append(["TRANSPORTES CALPI, S.A. DE C.V."])
    ws.append(["REPORTE DE VIAJES PENDIENTES / FILTRADOS"])
    ws.append([])
    
    ws['A1'].font = font_titulo
    ws['A2'].font = font_sub
    
    headers = ["No.", "# de Viaje", "Fecha", "Cliente", "Valor de Flete", "Observaciones", "Tipo Viaje"]
    ws.append(headers)
    
    header_row_idx = 4
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row_idx, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin
        
    current_row = 5
    for idx, row in enumerate(data_filtered.itertuples(), start=1):
        v_viaje = getattr(row, col_viaje.replace(" ", "_"), getattr(row, df.columns[0], ""))
        v_fecha = getattr(row, col_fecha.replace(" ", "_"), "")
        if pd.notna(v_fecha):
            v_fecha = pd.to_datetime(v_fecha).strftime("%d/%m/%Y")
        else:
            v_fecha = ""
            
        v_cliente = str(getattr(row, col_cliente.replace(" ", "_"), ""))
        v_flete = float(getattr(row, col_flete.replace(" ", "_"), 0))
        v_obs = str(getattr(row, col_observaciones.replace(" ", "_"), "")) if col_observaciones in data_filtered.columns else ""
        v_tipo = str(getattr(row, col_tipoviaje.replace(" ", "_"), "")) if col_tipoviaje in data_filtered.columns else ""
        
        row_data = [idx, v_viaje, v_fecha, v_cliente, v_flete, v_obs, v_tipo]
        ws.append(row_data)
        
        for c_idx in range(1, len(row_data) + 1):
            cell = ws.cell(row=current_row, column=c_idx)
            cell.font = font_data
            cell.border = border_thin
            
            if c_idx in [1, 2, 3]:
                cell.alignment = Alignment(horizontal="center")
            elif c_idx == 5:
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")
                
        current_row += 1
        
    total_row_idx = current_row
    ws.cell(row=total_row_idx, column=4, value="TOTAL PENDIENTE DE FACTURAR").font = font_total
    ws.cell(row=total_row_idx, column=4).alignment = Alignment(horizontal="right")
    
    formula_flete = f"=SUM(E5:E{total_row_idx-1})" if total_row_idx > 5 else 0
    cell_total = ws.cell(row=total_row_idx, column=5, value=formula_flete)
    cell_total.font = font_total
    cell_total.number_format = '"$"#,##0.00'
    cell_total.alignment = Alignment(horizontal="right")
    
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=total_row_idx, column=c_idx)
        cell.border = border_total
        cell.fill = fill_total

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row > 2:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# --- BOTONES DE DESCARGA ---
st.markdown("---")
c_down1, c_down2 = st.columns(2)

with c_down1:
    csv_data = df_filtered[cols_visibles_keys].to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Descargar Reporte en CSV",
        data=csv_data,
        file_name="reporte_viajes_calpi.csv",
        mime="text/csv"
    )

with c_down2:
    excel_bytes = generar_excel_calpi(df_filtered)
    st.download_button(
        label="📥 Descargar Reporte Ejecutivo en Excel (.xlsx)",
        data=excel_bytes,
        file_name="Reporte_Viajes_Calpi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
