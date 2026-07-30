import io
import os
import re
import pandas as pd
import streamlit as st

# ==========================================
# 1. PAGE CONFIG & BRANDING CLEANUP
# ==========================================
st.set_page_config(
    page_title="Portal de Operaciones",
    page_icon="📦",
    layout="centered" # Changed to centered for a tighter, cleaner dropdown UI
)

hide_streamlit_branding = """
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    a[href*="github.com"] {display: none !important;}
    
    .stAppViewerFooter, 
    div[class*="stAppDeployButton"],
    div[data-testid="stStatusWidget"],
    .viewerBadge_container__1QSob,
    [data-testid="stDecoration"],
    div[class*="viewerBadge"],
    div[class*="styles_viewerBadge"] {
        display: none !important;
        visibility: hidden !important;
    }
    </style>
"""
st.markdown(hide_streamlit_branding, unsafe_allow_html=True)

# ==========================================
# 2. AUTHENTICATION / PASSWORD PROTECTION
# ==========================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.title("🔒 Portal de Operaciones")
    st.markdown("Por favor, ingresa la contraseña para acceder al sistema.")
    
    password_input = st.text_input("Contraseña", type="password", label_visibility="collapsed")
    
    if st.button("Ingresar", use_container_width=True):
        if password_input == "Operaciones2026!":
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("🔑 Contraseña incorrecta. Intenta nuevamente.")
    return False

if not check_password():
    st.stop()

# ==========================================
# 3. HELPER FUNCTIONS & REPORT LOGIC
# ==========================================
SUPERVISORS_FILE = "SUPERVISORES.xlsx"

def extract_date_from_filename(filename, default_prefix="REPORT"):
    match = re.search(r'\d{4}[-_]?\d{2}[-_]?\d{2}', filename)
    if match:
        clean_date = match.group(0).replace('-', '').replace('_', '')
        return f"{default_prefix}_{clean_date}.xlsx"
    return f"{default_prefix}_PROCESADO.xlsx"

def process_missing_scan(uploaded_file):
    if not os.path.exists(SUPERVISORS_FILE):
        st.error(f"❌ Error: El archivo base '{SUPERVISORS_FILE}' no se encuentra en el servidor.")
        return None
        
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_sup = pd.read_excel(SUPERVISORS_FILE)
        
        df_merged = pd.merge(df_raw, df_sup, on="Estación", how="left")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_merged.to_excel(writer, sheet_name="MISSING SCAN", index=False)
            workbook  = writer.book
            worksheet = writer.sheets["MISSING SCAN"]
            
            header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white'})
            for col_num, col_name in enumerate(df_merged.columns):
                worksheet.write(0, col_num, col_name, header_format)
                worksheet.set_column(col_num, col_num, max(len(str(col_name)) + 3, 12))
        return output.getvalue()
    except Exception as e:
        st.error(f"❌ Error procesando el archivo: {e}")
        return None

def process_r5_metropolitan(uploaded_file):
    if not os.path.exists(SUPERVISORS_FILE):
        st.error(f"❌ Error: El archivo base '{SUPERVISORS_FILE}' no se encuentra en el servidor.")
        return None
        
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_sup = pd.read_excel(SUPERVISORS_FILE)
        
        df_merged = pd.merge(df_raw, df_sup, on="Estación", how="left")
        
        pivot = pd.pivot_table(
            df_merged,
            index=["Supervisor", "Estación"],
            values=["Guía"],
            aggfunc="count",
            fill_value=0
        ).reset_index()

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pivot.to_excel(writer, sheet_name="R5 METROPOLITANO", index=False)
            workbook = writer.book
            worksheet = writer.sheets["R5 METROPOLITANO"]
            
            header_format = workbook.add_format({'bold': True, 'bg_color': '#203764', 'font_color': 'white'})
            for col_num, col_name in enumerate(pivot.columns):
                worksheet.write(0, col_num, col_name, header_format)
                worksheet.set_column(col_num, col_num, max(len(str(col_name)) + 4, 15))
        return output.getvalue()
    except Exception as e:
        st.error(f"❌ Error procesando el archivo: {e}")
        return None

def process_anomalies(uploaded_file):
    if not os.path.exists(SUPERVISORS_FILE):
        st.error(f"❌ Error: El archivo base '{SUPERVISORS_FILE}' no se encuentra en el servidor.")
        return None
        
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheets = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
        df_sup = pd.read_excel(SUPERVISORS_FILE)

        first_sheet_name = list(sheets.keys())[0]
        df_main = sheets[first_sheet_name]
        
        if "Estación" in df_main.columns:
            df_main = pd.merge(df_main, df_sup, on="Estación", how="left")

        summary_pivot = pd.pivot_table(
            df_main,
            index=["Tipo de Anomalía"],
            values=["Guía"],
            aggfunc="count",
            margins=True,
            margins_name="Total General"
        ).reset_index() if "Tipo de Anomalía" in df_main.columns else pd.DataFrame()

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df_sheet in sheets.items():
                df_sheet.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            
            if not summary_pivot.empty:
                summary_pivot.to_excel(writer, sheet_name="RESUMEN ANOMALÍAS", index=False)
                workbook = writer.book
                worksheet = writer.sheets["RESUMEN ANOMALÍAS"]
                header_format = workbook.add_format({'bold': True, 'bg_color': '#375623', 'font_color': 'white'})
                for col_num, col_name in enumerate(summary_pivot.columns):
                    worksheet.write(0, col_num, col_name, header_format)
                    worksheet.set_column(col_num, col_num, 20)
        return output.getvalue()
    except Exception as e:
        st.error(f"❌ Error procesando el archivo: {e}")
        return None

# ==========================================
# 4. MAIN USER INTERFACE
# ==========================================
st.title("📦 Portal de Operaciones")
st.markdown("Automatización de reportes logísticos. Selecciona el módulo y carga tu archivo crudo (raw) para generar el reporte formateado de manera inmediata.")
st.divider()

# Dropdown Menu
report_type = st.selectbox(
    "📋 Selecciona el reporte a procesar:",
    ("Missing Scan", "R5 Metropolitano", "Anomalías"),
    index=0
)

st.markdown("<br>", unsafe_allow_html=True)

# Dynamic UI Context Setup
if report_type == "Missing Scan":
    st.info("💡 **Missing Scan:** Se cruzará la información de las estaciones con el archivo base de supervisores.")
    file_label = "Cargar reporte crudo (Excel) de Missing Scan"
    prefix = "MISSING_SCAN"
    process_func = process_missing_scan

elif report_type == "R5 Metropolitano":
    st.info("💡 **R5 Metropolitano:** Se generará una tabla dinámica resumiendo el conteo de guías por Supervisor y Estación.")
    file_label = "Cargar reporte crudo (Excel) de R5 Metropolitano"
    prefix = "R5_METROPOLITANO"
    process_func = process_r5_metropolitan

else:
    st.info("💡 **Anomalías:** Se conservarán todas las hojas originales y se añadirá una pestaña final con el resumen de anomalías.")
    file_label = "Cargar reporte crudo (Excel) de Anomalías"
    prefix = "ANOMALIAS"
    process_func = process_anomalies

# Fluid Upload & Auto-Process Flow
uploaded_file = st.file_uploader(file_label, type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner(f"Analizando y procesando {report_type}..."):
        processed_bytes = process_func(uploaded_file)
        
        if processed_bytes:
            output_name = extract_date_from_filename(uploaded_file.name, prefix)
            st.success("✅ ¡Reporte generado exitosamente!")
            
            # Centered download button styling
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    label=f"📥 Descargar {output_name}",
                    data=processed_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )