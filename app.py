import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import xlsxwriter

# ==============================================================================
# PAGE CONFIGURATION & UI CLEANUP
# ==============================================================================
st.set_page_config(page_title="Automated Reporting Portal", page_icon="📊", layout="centered")

hide_streamlit_branding = """
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    a[href*="github.com"] {display: none !important;}
    </style>
"""
st.markdown(hide_streamlit_branding, unsafe_allow_html=True)


# ==============================================================================
# SECURITY: Basic Password Authentication
# ==============================================================================
def check_password():
    def password_entered():
        if st.session_state["password"] == "Operaciones2026!":
            st.session_state["password_correct"] = True
            del st.session_state["password"] 
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("🔒 Please enter the password to access the portal:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("❌ Incorrect password. Please try again:", type="password", on_change=password_entered, key="password")
        return False
    return True


# ==============================================================================
# HELPER FUNCTIONS 
# ==============================================================================
def get_rounded_time_display():
    """
    Calculates time display in CDMX Timezone (UTC-6) according to rounding rules:
    IF MINUTE <= 15 -> :00
    IF MINUTE <= 45 -> :30
    ELSE -> next hour :00
    """
    cdmx_tz = datetime.timezone(datetime.timedelta(hours=-6))
    now = datetime.datetime.now(cdmx_tz)
    minute = now.minute
    if minute <= 15:
        r_dt = now.replace(minute=0, second=0, microsecond=0)
    elif minute <= 45:
        r_dt = now.replace(minute=30, second=0, microsecond=0)
    else:
        r_dt = (now + datetime.timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)
    
    return r_dt.strftime("%I:%M %p").lstrip('0')

def clean_spv_value(val):
    """
    Consolidates CEDIS, NaN, empty strings, and unmapped rows into #N/A
    """
    val_str = str(val).strip() if pd.notna(val) else '#N/A'
    if val_str in ['CEDIS', 'nan', 'NaN', '', 'None', 'NoneType']:
        return '#N/A'
    return val_str

def clean_spv_df(spv_file_bytes):
    """
    Standardizes the SUPERVISORES mapping file to ensure SPV, ZONA, and PDV columns 
    are universally correctly cased and formatted to prevent KeyErrors.
    """
    df_spv = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv.columns = [str(c).strip().upper() for c in df_spv.columns]
    
    if 'SUPERVISOR' in df_spv.columns and 'SPV' not in df_spv.columns:
        df_spv.rename(columns={'SUPERVISOR': 'SPV'}, inplace=True)
        
    if 'ZONA' not in df_spv.columns:
        df_spv['ZONA'] = ''
        
    return df_spv


# ==============================================================================
# REPORT LOGIC WRAPPERS
# ==============================================================================

def generate_tiktok_visits(raw_file_bytes, spv_file_bytes):
    # 1. SAFELY LOAD DATA INTO MEMORY
    xls_raw = pd.ExcelFile(io.BytesIO(raw_file_bytes))
    original_sheet_name = '订单综合信息' if '订单综合信息' in xls_raw.sheet_names else xls_raw.sheet_names[0]
    df_raw = pd.read_excel(xls_raw, sheet_name=original_sheet_name)
    
    # Clean redundant calculated columns from previous runs if present in raw data sheet
    for col in ['SPV', 'SLR']:
        if col in df_raw.columns:
            df_raw = df_raw.drop(columns=[col])
            
    df_raw_copy = df_raw.copy()  # Preserve original raw data to write back
    
    # Load supervisor mapping file
    df_sup = clean_spv_df(spv_file_bytes)
    df_sup.rename(columns={'SPV': 'SUPERVISOR'}, inplace=True)
            
    if 'PDV' in df_sup.columns:
        df_sup['PDV'] = df_sup['PDV'].astype(str).str.strip()
    
    if 'Punto de Recogida' in df_raw.columns:
        df_raw['Punto de Recogida'] = df_raw['Punto de Recogida'].astype(str).str.strip()

    # 2. AUTOMATIC TARGET HOUR DETECTION FROM 'TABLA' SHEET
    try:
        df_tabla = pd.read_excel(xls_raw, sheet_name='TABLA')
        existing_cols = [str(c).strip() for c in df_tabla.columns]
        
        if '9' not in existing_cols: target_hour = 9
        elif '12' not in existing_cols: target_hour = 12
        elif '3' not in existing_cols: target_hour = 3
        else: target_hour = 5
    except Exception:
        df_tabla = pd.DataFrame()
        target_hour = 9

    # 3. DYNAMIC DATE EXTRACTION
    months_en = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun', 
                 7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'}
    
    report_date_str = ""
    if 'Tiempo de registro de la orden' in df_raw.columns:
        dates = pd.to_datetime(df_raw['Tiempo de registro de la orden']).dt.date
        o_date = dates.value_counts().idxmax()
        report_date_str = f"{o_date.day:02d}-{months_en[o_date.month]}"
    else:
        report_date_str = "N/A"

    # 4. DATA CLEANING & MERGING
    df_raw['SLR'] = df_raw['Compañía remitente'].astype(str).apply(
        lambda x: x.split('|')[-1].strip() if '|' in x else x.strip()
    )
    
    df_raw = pd.merge(df_raw, df_sup[['PDV', 'SUPERVISOR']], left_on='Punto de Recogida', right_on='PDV', how='left')
    df_raw.rename(columns={'SUPERVISOR': 'SPV'}, inplace=True)
    
    # Explicitly consolidate CEDIS / NaNs / Empty into '#N/A'
    df_raw['SPV'] = df_raw['SPV'].apply(clean_spv_value)
    
    # Filter out both "CANCELADO" and "EN RECEPCIÓN"
    if 'Estado de la orden' in df_raw.columns:
        df_raw['Estado_upper'] = df_raw['Estado de la orden'].astype(str).str.upper().str.strip()
        df_valid = df_raw[~df_raw['Estado_upper'].isin(['CANCELADO', 'EN RECEPCIÓN', 'EN RECEPCION'])]
    else:
        df_valid = df_raw.copy()

    # 5. TIME LOGIC & PIVOTS
    time_col = str(target_hour)
    df_current_counts = df_valid.groupby('SLR')['Guía de Rastreo'].count().reset_index(name=time_col)

    if target_hour == 9:
        df_mapping = df_valid[['SLR', 'Punto de Recogida', 'SPV']].drop_duplicates(subset=['SLR'])
        df_mapping.rename(columns={'Punto de Recogida': 'PDV'}, inplace=True)
        
        df_tabla = pd.merge(df_mapping, df_current_counts, on='SLR', how='left')
        df_tabla[time_col] = df_tabla[time_col].fillna(0).astype(int)
        df_tabla['SPV'] = df_tabla['SPV'].apply(clean_spv_value)
        
        df_template_source = df_tabla.copy()
        vol_col = time_col
    else:
        df_tabla.columns = [str(c).strip() for c in df_tabla.columns]
        if time_col in df_tabla.columns: df_tabla.drop(columns=[time_col], inplace=True)
        if 'FORMULA' in df_tabla.columns: df_tabla.drop(columns=['FORMULA'], inplace=True)

        df_tabla = pd.merge(df_tabla, df_current_counts, on='SLR', how='left')
        df_tabla['SPV'] = df_tabla['SPV'].apply(clean_spv_value)
        
        # 0s in time slots 12, 3, 5 mean all packages collected -> converted to '#N/A'
        df_tabla[time_col] = df_tabla[time_col].apply(lambda x: '#N/A' if pd.isna(x) or x == 0 else int(x))
        
        if '9' in df_tabla.columns:
            def calculate_formula(row):
                val_9 = row['9']
                val_curr = row[time_col]
                if str(val_curr) == '#N/A' or str(val_9) == '#N/A':
                    return '#N/A'
                try:
                    diff = int(val_9) - int(val_curr)
                    return diff
                except Exception:
                    return '#N/A'

            df_tabla['FORMULA'] = df_tabla.apply(calculate_formula, axis=1)
        else:
            raise ValueError("Error: Columna '9' no encontrada en la hoja TABLA. Procesa primero las 9 AM.")
            
        # Filter for sellers where formula results in 0
        df_template_source = df_tabla[df_tabla['FORMULA'].astype(str) == '0'].copy()
        vol_col = time_col

    # REORDER TABLA COLUMNS: SPV -> PDV -> SLR -> hours (9, 12, 3, 5, FORMULA, etc.)
    desired_order = ['SPV', 'PDV', 'SLR']
    other_columns = [col for col in df_tabla.columns if col not in desired_order]
    final_tabla_columns = [col for col in desired_order if col in df_tabla.columns] + other_columns
    df_tabla = df_tabla[final_tabla_columns]

    # 6. TEMPLATE GENERATION WITH CUSTOM SORTING (#N/A AT THE VERY END)
    unique_spvs = [s for s in df_template_source['SPV'].unique() if str(s).strip() != '#N/A']
    unique_spvs.sort()
    if '#N/A' in df_template_source['SPV'].unique() or '#N/A' in df_tabla['SPV'].unique():
        unique_spvs.append('#N/A')

    template_data = []
    for spv in unique_spvs:
        group = df_template_source[df_template_source['SPV'] == spv]
        
        # Parse column value cleanly for numeric comparisons
        group_vols = pd.to_numeric(group[vol_col], errors='coerce').fillna(0)
        
        template_data.append({
            'SUPERVISOR': spv,
            'TOTAL TT SELLERS': len(group),
            '1-10': len(group[(group_vols >= 1) & (group_vols <= 10)]),
            '11-50': len(group[(group_vols >= 11) & (group_vols <= 50)]),
            '+50': len(group[group_vols > 50])
        })
        
    df_template = pd.DataFrame(template_data)
    
    total_row = {
        'SUPERVISOR': 'TOTAL',
        'TOTAL TT SELLERS': df_template['TOTAL TT SELLERS'].sum() if not df_template.empty else 0,
        '1-10': df_template['1-10'].sum() if not df_template.empty else 0,
        '11-50': df_template['11-50'].sum() if not df_template.empty else 0,
        '+50': df_template['+50'].sum() if not df_template.empty else 0
    }
    df_template = pd.concat([df_template, pd.DataFrame([total_row])], ignore_index=True)

    # 7. EXCEL OUTPUT GENERATION
    output_buffer = io.BytesIO()

    with pd.ExcelWriter(output_buffer, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        workbook = writer.book
        
        # --- Sheet 1: VISITAS ---
        sheet_name = 'VISITAS'
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide_gridlines(0)
        
        font_family = 'Times New Roman'
        
        blue_bg = workbook.add_format({'bg_color': '#5B9BD5', 'font_color': 'black', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_name': font_family, 'font_size': 11})
        purple_bg = workbook.add_format({'bg_color': '#7030A0', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bold': True, 'text_wrap': True, 'font_name': font_family, 'font_size': 11})
        red_bg = workbook.add_format({'bg_color': '#C00000', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_name': font_family, 'font_size': 11})
        yellow_bg = workbook.add_format({'bg_color': '#FFFF00', 'font_color': 'black', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_name': font_family, 'font_size': 11})
        
        title_fmt = workbook.add_format({'bold': True, 'font_size': 13, 'align': 'center', 'valign': 'vcenter', 'font_name': font_family})
        sub_title_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'font_name': font_family})
        time_badge_fmt = workbook.add_format({'bg_color': '#FFFF00', 'bold': True, 'italic': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12, 'font_name': font_family, 'border': 1})
        
        data_fmt = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'font_name': font_family, 'font_size': 11})
        data_num_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_name': font_family, 'font_size': 11})
        
        worksheet.set_column('A:A', 52)
        worksheet.set_column('B:E', 22)

        worksheet.merge_range('B1:D1', 'TT SELLERS TO VISIT BASED ON THE # OF PACKAGES', title_fmt)
        worksheet.merge_range('B2:D2', '按包裹数量需拜访的TT卖家', sub_title_fmt)
        
        formatted_time = get_rounded_time_display()
        worksheet.merge_range('E1:E2', formatted_time, time_badge_fmt)
        worksheet.set_row(0, 22)
        worksheet.set_row(1, 22)
        
        worksheet.write(2, 0, 'SUPERVISOR', blue_bg)
        worksheet.write(2, 1, 'TOTAL TT SELLERS', purple_bg)
        worksheet.write(2, 2, '1-10', red_bg)
        worksheet.write(2, 3, '11-50', red_bg)
        worksheet.write(2, 4, '+50', red_bg)
        worksheet.set_row(2, 28)
        
        worksheet.write(3, 0, '', purple_bg)
        worksheet.merge_range(3, 1, 3, 4, report_date_str, purple_bg)
        worksheet.set_row(3, 22)
        
        current_row = 4
        for _, row in df_template.iterrows():
            is_total_row = (row['SUPERVISOR'] == 'TOTAL')
            row_format = yellow_bg if is_total_row else data_num_fmt
            str_format = yellow_bg if is_total_row else data_fmt
            
            worksheet.write(current_row, 0, str(row['SUPERVISOR']), str_format)
            worksheet.write(current_row, 1, int(row['TOTAL TT SELLERS']), row_format)
            worksheet.write(current_row, 2, int(row['1-10']), row_format)
            worksheet.write(current_row, 3, int(row['11-50']), row_format)
            worksheet.write(current_row, 4, int(row['+50']), row_format)
            worksheet.set_row(current_row, 20)
            current_row += 1

        # --- Sheet 2: TABLA ---
        df_tabla.to_excel(writer, sheet_name='TABLA', index=False)
        worksheet_tabla = writer.sheets['TABLA']
        worksheet_tabla.hide_gridlines(0)
        
        tabla_header = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_name': font_family})
        tabla_data = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_name': font_family})
        
        for col_num, value in enumerate(df_tabla.columns):
            worksheet_tabla.write(0, col_num, value, tabla_header)
            max_len = max(df_tabla[value].astype(str).map(len).max(), len(str(value))) + 4
            worksheet_tabla.set_column(col_num, col_num, max(max_len, 14), tabla_data)
        worksheet_tabla.set_row(0, 24)

        # --- Sheet 3: ORIGINAL RAW DATA ---
        df_raw_copy.to_excel(writer, sheet_name=original_sheet_name, index=False)
        worksheet_raw = writer.sheets[original_sheet_name]
        worksheet_raw.hide_gridlines(0)

    output_filename = "TIKTOK PENDING FROM YESTERDAY.xlsx"
    return output_buffer.getvalue(), output_filename


def generate_aliexpress(raw_file_bytes, spv_file_bytes):
    df_spv = clean_spv_df(spv_file_bytes)
    df_spv['PDV'] = df_spv['PDV'].astype(str).str.strip()
    
    valid_spvs = [str(spv).strip().upper() for spv in df_spv['SPV'].dropna().unique() if 'CEDIS' not in str(spv).strip().upper()]

    try:
        xls = pd.ExcelFile(io.BytesIO(raw_file_bytes))
        original_sheet_name = '订单综合信息' if '订单综合信息' in xls.sheet_names else xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=original_sheet_name)
    except ValueError:
        df_raw = pd.read_excel(io.BytesIO(raw_file_bytes))
        original_sheet_name = 'Sheet1'

    df_raw_original = df_raw.copy()

    if 'Estado de la orden' in df_raw.columns:
        df_raw = df_raw[df_raw['Estado de la orden'] != 'Cancelado']

    seller_pdv_map = df_raw[['Compañía remitente', 'Punto de Recogida']].dropna().drop_duplicates(subset=['Compañía remitente'])
    seller_pdv_map = seller_pdv_map.rename(columns={'Punto de Recogida': 'PDV'})
    seller_pdv_map['PDV'] = seller_pdv_map['PDV'].astype(str).str.strip()

    df_total = df_raw.groupby('Compañía remitente')['Guía de Rastreo'].count().reset_index(name='Total')
    df_abnormal = df_raw[df_raw['Estado de la orden'] == 'La recogida falló'].groupby('Compañía remitente')['Guía de Rastreo'].count().reset_index(name='Abnormal')

    pending_statuses = ['Asignado a PDV', 'Asignado a un mensajero', 'Asignado a mensajero', 'Pendiente', 'Creado']
    df_por_rec = df_raw[df_raw['Estado de la orden'].isin(pending_statuses)].groupby('Compañía remitente')['Guía de Rastreo'].count().reset_index(name='PorRec')

    df_report = df_total.copy()
    df_report = pd.merge(df_report, df_abnormal, on='Compañía remitente', how='left')
    df_report = pd.merge(df_report, df_por_rec, on='Compañía remitente', how='left')

    df_report['Total'] = df_report['Total'].fillna(0).astype(int)
    df_report['Abnormal'] = df_report['Abnormal'].fillna(0).astype(int)
    df_report['PorRec'] = df_report['PorRec'].fillna(0).astype(int)
    df_report['Visita'] = df_report['Total'] - df_report['Abnormal'] - df_report['PorRec']

    df_report = pd.merge(df_report, seller_pdv_map, on='Compañía remitente', how='left')
    df_report = pd.merge(df_report, df_spv[['PDV', 'SPV', 'ZONA']], on='PDV', how='left')

    df_report = df_report.rename(columns={'Compañía remitente': 'SELLER'})
    df_report = df_report.dropna(subset=['SPV']) 
    df_report = df_report[df_report['SPV'].str.strip() != '']

    df_report['SPV'] = df_report['SPV'].astype(str).str.upper()
    df_report['ZONA'] = df_report['ZONA'].fillna('').astype(str)
    df_report['PDV'] = df_report['PDV'].fillna('').astype(str)
    df_report['SELLER'] = df_report['SELLER'].fillna('').astype(str)
    
    df_report = df_report[df_report['SPV'].isin(valid_spvs)]
    df_report = df_report.sort_values(by=['SPV', 'ZONA', 'PDV', 'SELLER'], ascending=[True, True, True, True])

    df_report['Rate'] = np.where(df_report['Total'] > 0, (df_report['Visita'] + df_report['Abnormal']) / df_report['Total'], 0.0)
    df_report['Abnormal_View'] = df_report['Abnormal'].apply(lambda x: '' if x == 0 else int(x))
    df_report['PorRec_View'] = df_report['PorRec'].apply(lambda x: '' if x == 0 else int(x))

    months_es = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'}

    if 'Tiempo de registro de la orden' in df_raw.columns:
        dates = pd.to_datetime(df_raw['Tiempo de registro de la orden']).dt.date
        o_date = dates.value_counts().idxmax() 
        r_date = o_date + datetime.timedelta(days=1)
        
        if r_date.weekday() == 0: 
            o_date_prev = o_date - datetime.timedelta(days=1)
            o_date_zh = f"{o_date_prev.month}月{o_date_prev.day}日 & {o_date.month}月{o_date.day}日"
            o_date_es = f"{months_es[o_date_prev.month]} {o_date_prev.day} & {months_es[o_date.month]} {o_date.day}"
        else:
            o_date_zh = f"{o_date.month}月{o_date.day}日"
            o_date_es = f"{months_es[o_date.month]} {o_date.day}"
            
        r_date_zh = f"{r_date.month}月{r_date.day}日"
        r_date_es = f"{months_es[r_date.month]} {r_date.day}"
    else:
        o_date_zh, o_date_es = "7月29日", "Jul 29"
        r_date_zh, r_date_es = "7月30日", "Jul 30"

    subtitle_str = f"订单日期：{o_date_zh}, 揽收日期 {r_date_zh} Pedidos: {o_date_es} | Reco: {r_date_es}"

    mx_tz = datetime.timezone(datetime.timedelta(hours=-6))
    now = datetime.datetime.now(mx_tz)
    ts = now.timestamp()
    ts_rounded = round(ts / 900) * 900
    now_rounded = datetime.datetime.fromtimestamp(ts_rounded, mx_tz)
    time_str = f"{now_rounded.hour}" if now_rounded.minute == 0 else f"{now_rounded.hour}.{now_rounded.minute:02d}"

    output_filename = f"ALI {r_date_es.upper()} AT {time_str}.xlsx"

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
    workbook = writer.book
    worksheet = workbook.add_worksheet('AliExpress')

    magenta, yellow, red = '#B0005B', '#FFFF00', '#FF0000'
    base_font = 'Calibri'
    base_size = 14

    title_format = workbook.add_format({'bold': True, 'font_size': 36, 'align': 'center', 'valign': 'vcenter', 'bg_color': magenta, 'font_color': 'white', 'font_name': base_font})
    subtitle_format = workbook.add_format({'bold': False, 'font_size': 15, 'align': 'center', 'valign': 'vcenter', 'bg_color': magenta, 'font_color': 'white', 'font_name': base_font})
    header_format = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': magenta, 'font_color': 'white', 'border': 1, 'text_wrap': True, 'font_name': base_font})

    grand_total_label = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': magenta, 'font_color': 'white', 'border': 1, 'font_name': base_font})
    grand_total_num = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': magenta, 'font_color': 'white', 'border': 1, 'num_format': '#,##0', 'font_name': base_font})
    grand_total_pct = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': red, 'font_color': 'white', 'border': 1, 'num_format': '0.00%', 'font_name': base_font})

    subtotal_label_yellow = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': yellow, 'font_color': 'black', 'border': 1, 'font_name': base_font})
    subtotal_num_yellow = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': yellow, 'font_color': 'black', 'border': 1, 'num_format': '#,##0', 'font_name': base_font})
    subtotal_pct_red = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': red, 'font_color': 'white', 'border': 1, 'num_format': '0.00%', 'font_name': base_font})
    subtotal_num_red = workbook.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': red, 'font_color': 'white', 'border': 1, 'num_format': '#,##0', 'font_name': base_font})

    spv_merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'font_size': base_size, 'font_name': base_font})
    data_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': base_size, 'font_name': base_font})
    data_num_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': base_size, 'num_format': '#,##0', 'font_name': base_font})
    data_pct_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': base_size, 'num_format': '0.00%', 'font_name': base_font})

    worksheet.set_column('A:A', 18); worksheet.set_column('B:B', 20); worksheet.set_column('C:C', 34); worksheet.set_column('D:D', 46)
    worksheet.set_column('E:I', 24)

    worksheet.merge_range('A1:I1', 'AliExpress 拜访率 % Visitas AliExpress', title_format)
    worksheet.set_row(0, 70) 
    worksheet.merge_range('A2:I2', subtitle_str, subtitle_format)
    worksheet.set_row(1, 28)

    headers = ['SPV', 'ZONA', '客户归属网点\nPDV', 'SELLER', '待收取总数\nTOTAL a Recolectar', '已记录拜访\nVISITA REGISTRADA', '异常扫描已记录\nABNORMAL SCAN', '待收取包裹数\nPOR RECOLECTAR', '商家拜访率\nRATE %']
    for col, h in enumerate(headers): worksheet.write(2, col, h, header_format)
    worksheet.set_row(2, 54)

    gt_total = int(df_report['Total'].sum())
    gt_abnormal = int(df_report['Abnormal'].sum())
    gt_porrec = int(df_report['PorRec'].sum())
    gt_visita_column = int(df_report['Visita'].sum()) + gt_abnormal
    gt_rate = gt_visita_column / gt_total if gt_total > 0 else 0.0

    worksheet.merge_range('A4:D4', 'GRAND TOTAL 总计', grand_total_label)
    worksheet.write(3, 4, gt_total, grand_total_num)
    worksheet.write(3, 5, gt_visita_column, grand_total_num)
    worksheet.write(3, 6, gt_abnormal, grand_total_num)
    worksheet.write(3, 7, gt_porrec, grand_total_num)
    worksheet.write(3, 8, float(gt_rate), grand_total_pct) 
    worksheet.set_row(3, 32)

    current_row = 4
    for spv_name, group in df_report.groupby('SPV', sort=False):
        spv_start_row = current_row
        
        zona_start_row = current_row
        current_zona = None

        for _, row in group.iterrows():
            if current_zona != row['ZONA']:
                if current_zona is not None:
                    if current_row - 1 > zona_start_row:
                        worksheet.merge_range(zona_start_row, 1, current_row - 1, 1, current_zona, data_format)
                    else:
                        worksheet.write(zona_start_row, 1, current_zona, data_format)
                current_zona = row['ZONA']
                zona_start_row = current_row

            worksheet.write(current_row, 2, row['PDV'], data_format)
            worksheet.write(current_row, 3, row['SELLER'], data_format)
            worksheet.write(current_row, 4, int(row['Total']), data_num_format)
            worksheet.write(current_row, 5, int(row['Visita']), data_num_format)
            worksheet.write(current_row, 6, row['Abnormal_View'], data_num_format)
            worksheet.write(current_row, 7, row['PorRec_View'], data_num_format)
            worksheet.write(current_row, 8, float(row['Rate']), data_pct_format)
            worksheet.set_row(current_row, 24)
            current_row += 1

        if current_zona is not None:
            if current_row - 1 > zona_start_row:
                worksheet.merge_range(zona_start_row, 1, current_row - 1, 1, current_zona, data_format)
            else:
                worksheet.write(zona_start_row, 1, current_zona, data_format)

        if current_row - 1 > spv_start_row: 
            worksheet.merge_range(spv_start_row, 0, current_row - 1, 0, spv_name, spv_merge_format)
        else: 
            worksheet.write(spv_start_row, 0, spv_name, spv_merge_format)

        sub_total, sub_visita, sub_abnormal, sub_porrec = int(group['Total'].sum()), int(group['Visita'].sum()), int(group['Abnormal'].sum()), int(group['PorRec'].sum())
        sub_rate = (sub_visita + sub_abnormal) / sub_total if sub_total > 0 else 0.0

        worksheet.merge_range(f'A{current_row+1}:D{current_row+1}', f'TOTAL {spv_name}', subtotal_label_yellow)
        worksheet.write(current_row, 4, sub_total, subtotal_num_yellow)
        worksheet.write(current_row, 5, sub_visita, subtotal_num_yellow)
        worksheet.write(current_row, 6, sub_abnormal, subtotal_num_yellow)
        worksheet.write(current_row, 7, sub_porrec, subtotal_num_red)
        worksheet.write(current_row, 8, float(sub_rate), subtotal_pct_red)

        worksheet.set_row(current_row, 32)
        current_row += 1

    df_raw_original.to_excel(writer, sheet_name=original_sheet_name, index=False)
    writer.close()
    
    return output.getvalue(), output_filename


def generate_missing_scan(raw_file_bytes, spv_file_bytes):
    df_spv_static = clean_spv_df(spv_file_bytes)

    valid_spvs = [str(spv).strip().upper() for spv in df_spv_static['SPV'].dropna().unique() if 'CEDIS' not in str(spv).strip().upper()]

    spv_map = df_spv_static[['SPV', 'ZONA', 'PDV']].dropna(subset=['PDV']).copy()
    spv_map.columns = ['spv', 'zona', 'pdv']
    spv_map['spv'] = spv_map['spv'].astype(str).str.strip()
    spv_map['zona'] = spv_map['zona'].astype(str).str.strip().replace('nan', 'N/A')
    spv_map['pdv'] = spv_map['pdv'].astype(str).str.strip()
    spv_map.drop_duplicates(subset=['pdv'], inplace=True)

    try:
        df_raw = pd.read_excel(io.BytesIO(raw_file_bytes), sheet_name='sheet0', engine='openpyxl')
    except ValueError:
        df_raw = pd.read_excel(io.BytesIO(raw_file_bytes), sheet_name=0, engine='openpyxl')
    df_raw.columns = df_raw.columns.str.strip()

    raw_date = df_raw['Fecha de estadísticas'].dropna().iloc[0]
    date_obj = pd.to_datetime(raw_date)
    months_en = {1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUN', 7: 'JUL', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'}
    months_es = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}
    
    title_date_str = f"缺失扫描报告 OMISIÓN DE ESCANEO\n揽收日期 {date_obj.month}月{date_obj.day}日 Fecha de Recoleccion: {months_es[date_obj.month]} {date_obj.day}"
    output_filename = f"MISSING SCAN {months_en[date_obj.month]} {date_obj.day}.xlsx"

    val_cols = ['El número de pedidos de escaneo', 'No. de escaneo faltante de recolección', 'Nº de guías con escaneo faltantes de salida']
    pivot_raw = pd.pivot_table(df_raw, index=['Nombre del nodo'], values=val_cols, aggfunc='sum').reset_index()
    pivot_raw.rename(columns={'Nombre del nodo': 'pdv'}, inplace=True)
    pivot_raw['pdv'] = pivot_raw['pdv'].astype(str).str.strip()

    final_df = pd.merge(spv_map, pivot_raw, on='pdv', how='inner')
    for col in val_cols:
        final_df[col] = final_df[col].fillna(0)

    final_df['spv_upper'] = final_df['spv'].astype(str).str.strip().str.upper()
    final_df = final_df[final_df['spv_upper'].isin(valid_spvs)].copy()

    final_df['tasa_recoleccion'] = final_df['No. de escaneo faltante de recolección'] / final_df['El número de pedidos de escaneo']
    final_df['tasa_salida'] = final_df['Nº de guías con escaneo faltantes de salida'] / final_df['El número de pedidos de escaneo']
    final_df.fillna({'tasa_recoleccion': 0, 'tasa_salida': 0}, inplace=True)
    final_df.sort_values(by=['spv', 'zona', 'pdv'], ascending=[True, True, True], inplace=True)

    wb = openpyxl.load_workbook(io.BytesIO(raw_file_bytes))
    if 'Report' in wb.sheetnames:
        del wb['Report']
    ws = wb.create_sheet(title='Report', index=0)

    font_title = Font(name='Calibri', size=36, bold=True, color='FFFFFF')
    font_white_bold = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
    font_black_bold = Font(name='Calibri', size=14, bold=True, color='000000')
    font_regular = Font(name='Calibri', size=14, bold=False, color='000000')
    font_green_bold = Font(name='Calibri', size=14, bold=True, color='00FF00')

    fill_black = PatternFill(start_color='1F1F1F', end_color='1F1F1F', fill_type='solid')
    fill_dark_grey = PatternFill(start_color='4F4F4F', end_color='4F4F4F', fill_type='solid')
    fill_mid_grey = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
    fill_red = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    fill_yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'), top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))

    col_widths = {'A': 18, 'B': 22, 'C': 34, 'D': 30, 'E': 34, 'F': 34, 'G': 30, 'H': 34}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.merge_cells('A1:H1')
    cell_title = ws['A1']
    cell_title.value = title_date_str
    cell_title.font = font_title
    cell_title.fill = fill_black
    cell_title.alignment = align_center
    ws.row_dimensions[1].height = 85
