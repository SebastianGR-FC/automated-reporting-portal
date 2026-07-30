import io
import re
import pandas as pd
import streamlit as st

# ==========================================
# 1. PAGE CONFIG & BRANDING CLEANUP (CSS)
# ==========================================
st.set_page_config(
    page_title="Portal de Operaciones",
    page_icon="📦",
    layout="wide"
)

# Hide Streamlit header, top menu, footer, and bottom-right badges
hide_streamlit_branding = """
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    a[href*="github.com"] {display: none !important;}
    
    /* Hide bottom-right Community Cloud badges & widgets */
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
    """Returns `True` if the user enters the correct password."""
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.title("🔒 Portal de Operaciones")
    password_input = st.text_input("Please enter the password to access the portal:", type="password")
    
    if st.button("Login"):
        if password_input == "Operaciones2026!":
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("🔑 Incorrect password. Please try again.")
    return False

if not check_password():
    st.stop()

# ==========================================
# 3. HELPER FUNCTIONS & REPORT LOGIC
# ==========================================
SUPERVISORS_FILE = "SUPERVISORES.xlsx"

def extract_date_from_filename(filename, default_prefix="REPORT"):
    """Extracts date matching YYYYMMDD or YYYY-MM-DD from the filename, or returns default."""
    match = re.search(r'\d{4}[-_]?\d{2}[-_]?\d{2}', filename)
    if match:
        clean_date = match.group(0).replace('-', '')
        return f"{default_prefix}_{clean_date}.xlsx"
    return f"{default_prefix}_PROCESADO.xlsx"

def process_missing_scan(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_sup = pd.read_excel(SUPERVISORS_FILE)
    except Exception as e:
        st.error(f"Error loading required files: {e}")
        return None

    # Normalization / Mapping logic
    df_merged = pd.merge(df_raw, df_sup, on="Estación", how="left")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_merged.to_excel(writer, sheet_name="MISSING SCAN", index=False)
        workbook  = writer.book
        worksheet = writer.sheets["MISSING SCAN"]
        
        # Formatting
        header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white'})
        for col_num, col_name in enumerate(df_merged.columns):
            worksheet.write(0, col_num, col_name, header_format)
            worksheet.set_column(col_num, col_num, max(len(str(col_name)) + 3, 12))

    return output.getvalue()

def process_r5_metropolitan(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_sup = pd.read_excel(SUPERVISORS_FILE)
    except Exception as e:
        st.error(f"Error loading required files: {e}")
        return None

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

def process_anomalies(uploaded_file):
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheets = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
        df_sup = pd.read_excel(SUPERVISORS_FILE)
    except Exception as e:
        st.error(f"Error loading required files: {e}")
        return None

    # Generate Summary Pivot from the main data sheet
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
        # Preserve original raw sheets
        for sheet_name, df_sheet in sheets.items():
            df_sheet.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        
        # Append summary sheet
        if not summary_pivot.empty:
            summary_pivot.to_excel(writer, sheet_name="RESUMEN ANOMALÍAS", index=False)
            
            workbook = writer.book
            worksheet = writer.sheets["RESUMEN ANOMALÍAS"]
            header_format = workbook.add_format({'bold': True, 'bg_color': '#375623', 'font_color': 'white'})
            for col_num, col_name in enumerate(summary_pivot.columns):
                worksheet.write(0, col_num, col_name, header_format)
                worksheet.set_column(col_num, col_num, 20)

    return output.getvalue()

# ==========================================
# 4. MAIN USER INTERFACE
# ==========================================
st.title("📦 Automatización de Reportes de Operaciones")
st.markdown("Selecciona el tipo de reporte que deseas procesar y carga el archivo raw de Excel.")

tab1, tab2, tab3 = st.tabs(["MISSING SCAN", "R5 METROPOLITANO", "ANOMALÍAS"])

# Tab 1: Missing Scan
with tab1:
    st.subheader("Reporte Missing Scan")
    file_ms = st.file_uploader("Cargar Excel raw de Missing Scan", type=["xlsx", "xls"], key="ms_file")
    if file_ms:
        if st.button("Procesar Missing Scan", key="btn_ms"):
            with st.spinner("Procesando datos..."):
                processed_bytes = process_missing_scan(file_ms)
                if processed_bytes:
                    output_name = extract_date_from_filename(file_ms.name, "MISSING_SCAN")
                    st.success("¡Reporte procesado exitosamente!")
                    st.download_button(
                        label="📥 Descargar Reporte Procesado",
                        data=processed_bytes,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

# Tab 2: R5 Metropolitan
with tab2:
    st.subheader("Reporte R5 Metropolitano")
    file_r5 = st.file_uploader("Cargar Excel raw de R5 Metropolitano", type=["xlsx", "xls"], key="r5_file")
    if file_r5:
        if st.button("Procesar R5 Metropolitano", key="btn_r5"):
            with st.spinner("Procesando datos..."):
                processed_bytes = process_r5_metropolitan(file_r5)
                if processed_bytes:
                    output_name = extract_date_from_filename(file_r5.name, "R5_METROPOLITANO")
                    st.success("¡Reporte procesado exitosamente!")
                    st.download_button(
                        label="📥 Descargar Reporte Procesado",
                        data=processed_bytes,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

# Tab 3: Anomalies
with tab3:
    st.subheader("Reporte de Anomalías")
    file_anom = st.file_uploader("Cargar Excel raw de Anomalías", type=["xlsx", "xls"], key="anom_file")
    if file_anom:
        if st.button("Procesar Anomalías", key="btn_anom"):
            with st.spinner("Procesando datos..."):
                processed_bytes = process_anomalies(file_anom)
                if processed_bytes:
                    output_name = extract_date_from_filename(file_anom.name, "ANOMALIAS")
                    st.success("¡Reporte procesado exitosamente!")
                    st.download_button(
                        label="📥 Descargar Reporte Procesado",
                        data=processed_bytes,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
