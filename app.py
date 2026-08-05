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
# REPORT LOGIC WRAPPERS
# ==============================================================================

def generate_aliexpress(raw_file_bytes, spv_file_bytes):
    df_spv = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv.columns = df_spv.columns.astype(str).str.strip().str.lower()
    df_spv['pdv'] = df_spv['pdv'].astype(str).str.strip()

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
    seller_pdv_map = seller_pdv_map.rename(columns={'Punto de Recogida': 'pdv'})
    seller_pdv_map['pdv'] = seller_pdv_map['pdv'].astype(str).str.strip()

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
    df_report = pd.merge(df_report, df_spv[['pdv', 'spv', 'zona']], on='pdv', how='left')

    df_report = df_report.rename(columns={'spv': 'SPV', 'zona': 'ZONA', 'pdv': 'PDV', 'Compañía remitente': 'SELLER'})
    df_report = df_report.dropna(subset=['SPV']) 
    df_report = df_report[df_report['SPV'].str.strip() != '']

    df_report['SPV'] = df_report['SPV'].astype(str).str.upper()
    df_report['ZONA'] = df_report['ZONA'].fillna('').astype(str)
    df_report['PDV'] = df_report['PDV'].fillna('').astype(str)
    df_report['SELLER'] = df_report['SELLER'].fillna('').astype(str)
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

    # Set timezone strictly to Mexico City (UTC-6)
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

    # RESTORED COLORS & INCREASED FONTS
    magenta, yellow, red = '#B0005B', '#FFFF00', '#FF0000'
    base_font = 'Calibri'
    base_size = 14

    title_format = workbook.add_format({'bold': True, 'font_size': 22, 'align': 'center', 'valign': 'vcenter', 'bg_color': magenta, 'font_color': 'white', 'font_name': base_font})
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
    worksheet.set_row(0, 45)
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
    df_spv_static = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv_static.columns = df_spv_static.columns.str.strip()

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

    valid_spvs = ['BONNIE', 'DIANA', 'LUCERO']
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

    font_title = Font(name='Calibri', size=22, bold=True, color='FFFFFF')
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

    # INCREASED SPACING FOR WRAPPED TEXT
    col_widths = {'A': 18, 'B': 22, 'C': 34, 'D': 30, 'E': 34, 'F': 34, 'G': 30, 'H': 34}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.merge_cells('A1:H1')
    cell_title = ws['A1']
    cell_title.value = title_date_str
    cell_title.font = font_title
    cell_title.fill = fill_black
    cell_title.alignment = align_center
    ws.row_dimensions[1].height = 65

    headers = ['负责人\nSPV', '区域\nZONA', '揽收网点\nPDV', '总扫描数据量\nTOTAL A', '未进行入库扫描包裹件数\nSIN ESCANEO DE', '未进行入库扫描包裹的百分比\n% SIN ESCANEO DE', '未进行出库扫描的包裹\nSIN ESCANEO DE', '未进行出库扫描包裹的百分比\n% SIN ESCANEO DE']
    for col_num, h_text in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_num, value=h_text)
        c.font = font_white_bold
        c.fill = fill_dark_grey
        c.alignment = align_center
        c.border = thin_border
    ws.row_dimensions[2].height = 65

    gt_total = final_df['El número de pedidos de escaneo'].sum()
    gt_rec = final_df['No. de escaneo faltante de recolección'].sum()
    gt_rec_pct = gt_rec / gt_total if gt_total else 0
    gt_sal = final_df['Nº de guías con escaneo faltantes de salida'].sum()
    gt_sal_pct = gt_sal / gt_total if gt_total else 0

    ws.merge_cells('A3:C3')
    ws['A3'].value, ws['A3'].font, ws['A3'].fill, ws['A3'].alignment, ws['A3'].border = '总计', font_white_bold, fill_mid_grey, align_center, thin_border
    
    for col_idx, val, num_fmt, f_style, fill_style in [(4, gt_total, '#,##0', font_green_bold, fill_mid_grey), (5, gt_rec, '#,##0', font_green_bold, fill_mid_grey), (6, gt_rec_pct, '0.00%', font_white_bold, fill_red), (7, gt_sal, '#,##0', font_green_bold, fill_mid_grey), (8, gt_sal_pct, '0.00%', font_white_bold, fill_red)]:
        c = ws.cell(row=3, column=col_idx, value=val)
        c.font, c.fill, c.alignment, c.number_format, c.border = f_style, fill_style, align_center, num_fmt, thin_border
    ws.row_dimensions[3].height = 28

    current_row = 4
    for spv_name, group in final_df.groupby('spv', sort=False):
        spv_start_row = current_row
        
        zona_start_row = current_row
        current_zona = None

        for _, row in group.iterrows():
            ws.cell(row=current_row, column=2, value=row['zona']).alignment = align_center
            ws.cell(row=current_row, column=2).border = thin_border
            ws.cell(row=current_row, column=2).font = font_regular
            
            if current_zona != row['zona']:
                if current_zona is not None and current_row - 1 > zona_start_row:
                    ws.merge_cells(start_row=zona_start_row, start_column=2, end_row=current_row - 1, end_column=2)
                current_zona = row['zona']
                zona_start_row = current_row

            ws.cell(row=current_row, column=3, value=row['pdv']).alignment = align_center
            for c_idx, key, fmt in [(4, 'El número de pedidos de escaneo', '#,##0'), (5, 'No. de escaneo faltante de recolección', '#,##0'), (6, 'tasa_recoleccion', '0.00%'), (7, 'Nº de guías con escaneo faltantes de salida', '#,##0'), (8, 'tasa_salida', '0.00%')]:
                c = ws.cell(row=current_row, column=c_idx, value=float(row[key]) if '%' in fmt else row[key])
                c.number_format, c.alignment = fmt, align_center
            for col_idx in [3, 4, 5, 6, 7, 8]:
                ws.cell(row=current_row, column=col_idx).font = font_regular
                ws.cell(row=current_row, column=col_idx).border = thin_border
            ws.row_dimensions[current_row].height = 24
            current_row += 1

        if current_zona is not None and current_row - 1 > zona_start_row:
            ws.merge_cells(start_row=zona_start_row, start_column=2, end_row=current_row - 1, end_column=2)

        if current_row - 1 > spv_start_row: 
            ws.merge_cells(start_row=spv_start_row, start_column=1, end_row=current_row - 1, end_column=1)
            
        ws.cell(row=spv_start_row, column=1, value=spv_name).font = font_black_bold
        ws.cell(row=spv_start_row, column=1).alignment = align_center
        ws.cell(row=spv_start_row, column=1).border = thin_border

        sub_total = group['El número de pedidos de escaneo'].sum()
        sub_rec = group['No. de escaneo faltante de recolección'].sum()
        sub_rec_pct = sub_rec / sub_total if sub_total else 0
        sub_sal = group['Nº de guías con escaneo faltantes de salida'].sum()
        sub_sal_pct = sub_sal / sub_total if sub_total else 0

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        sub = ws.cell(row=current_row, column=1, value=f'TOTAL {spv_name}')
        sub.font, sub.fill, sub.alignment, sub.border = font_black_bold, fill_yellow, align_center, thin_border
        
        ws.cell(row=current_row, column=2).border = thin_border
        ws.cell(row=current_row, column=3).border = thin_border

        for col_idx, val, num_fmt, f_style, fill_style in [(4, sub_total, '#,##0', font_black_bold, fill_yellow), (5, sub_rec, '#,##0', font_black_bold, fill_yellow), (6, sub_rec_pct, '0.00%', font_white_bold, fill_red), (7, sub_sal, '#,##0', font_black_bold, fill_yellow), (8, sub_sal_pct, '0.00%', font_white_bold, fill_red)]:
            c = ws.cell(row=current_row, column=col_idx, value=val)
            c.font, c.fill, c.alignment, c.number_format, c.border = f_style, fill_style, align_center, num_fmt, thin_border
        ws.row_dimensions[current_row].height = 28
        current_row += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), output_filename


def generate_r7_cdmx(raw_file_bytes, spv_file_bytes):
    df_spv = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv.columns = df_spv.columns.astype(str).str.strip()
    df_spv['SPV'] = df_spv['SPV'].astype(str).str.strip().str.upper()
    df_spv['ZONA'] = df_spv['ZONA'].astype(str).str.strip()
    df_spv['PDV'] = df_spv['PDV'].astype(str).str.strip()
    df_spv['PDV_lower'] = df_spv['PDV'].str.lower()
    
    valid_spvs = ['BONNIE', 'DIANA', 'LUCERO']
    spv_map = df_spv[df_spv['SPV'].isin(valid_spvs)].drop_duplicates(subset=['PDV_lower']).copy()

    df_raw = None
    try:
        df_raw = pd.read_excel(io.BytesIO(raw_file_bytes), engine='openpyxl')
    except Exception:
        for enc in ['utf-8-sig', 'latin1', 'cp1252', 'gbk']:
            for sep in [',', '\t', ';', '|']:
                try:
                    df_temp = pd.read_csv(io.BytesIO(raw_file_bytes), sep=sep, encoding=enc, on_bad_lines='skip', low_memory=False)
                    if len(df_temp.columns) > 3:
                        df_raw = df_temp
                        break
                except Exception: continue
            if df_raw is not None: break

    df_raw.columns = df_raw.columns.astype(str).str.strip()
    df_raw['NORMAL_PDV'] = df_raw['Punto de Recogida'].astype(str).str.strip()
    df_raw['PDV_lower'] = df_raw['NORMAL_PDV'].str.lower()
    df_raw['Estado_lower'] = df_raw['Estado de la orden'].astype(str).str.strip().str.lower()

    order_times = pd.to_datetime(df_raw['Tiempo de registro de la orden'], errors='coerce')
    unique_dates = sorted(order_times.dt.date.dropna().unique())
    months_en = {1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUN', 7: 'JUL', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'}
    months_es = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}

    if unique_dates:
        min_o, max_o = unique_dates[0], unique_dates[-1]
        reco_date = max_o + datetime.timedelta(days=1)
        if min_o == max_o:
            pedidos_es = f"{months_es[min_o.month]} {min_o.day}"
            pedidos_zh = f"{min_o.month}月{min_o.day}日"
        elif min_o.month == max_o.month:
            pedidos_es = f"{months_es[min_o.month]} {min_o.day}-{max_o.day}"
            pedidos_zh = f"{min_o.month}月{min_o.day}-{max_o.day}日"
        else:
            pedidos_es = f"{months_es[min_o.month]} {min_o.day}-{months_es[max_o.month]} {max_o.day}"
            pedidos_zh = f"{min_o.month}月{min_o.day}日-{max_o.month}月{max_o.day}日"
        reco_es = f"{months_es[reco_date.month]} {reco_date.day}"
        reco_zh = f"{reco_date.month}月{reco_date.day}日"
    else:
        today = datetime.date.today()
        reco_date = today + datetime.timedelta(days=1)
        pedidos_es, pedidos_zh = f"{months_es[today.month]} {today.day}", f"{today.month}月{today.day}日"
        reco_es, reco_zh = f"{months_es[reco_date.month]} {reco_date.day}", f"{reco_date.month}月{reco_date.day}日"

    subtitle_str = f"订单日期: {pedidos_zh}, 揽收日期 {reco_zh} Pedidos: {pedidos_es} | Reco: {reco_es}"
    
    # FIXED FILENAME
    output_filename = f"R7 CDMX {months_en[reco_date.month]} {reco_date.day}.xlsx"

    df_total = df_raw[~df_raw['Estado_lower'].str.contains('cancelad', na=False)].copy()
    pivot_total = df_total.groupby('PDV_lower').size().reset_index(name='Total de Guias')

    spv_lookup = spv_map.set_index('PDV_lower')
    df_report = pivot_total.copy()
    df_report['SPV'] = df_report['PDV_lower'].map(spv_lookup['SPV'])
    df_report['ZONA'] = df_report['PDV_lower'].map(spv_lookup['ZONA'])
    df_report['PDV'] = df_report['PDV_lower'].map(spv_lookup['PDV'])
    df_report = df_report[df_report['SPV'].notna()].copy()

    mask_cancelado = df_raw['Estado_lower'].str.contains('cancelad', na=False)
    mask_recepcion = df_raw['Estado_lower'].str.contains('recep', na=False)
    df_por_rec = df_raw[~(mask_cancelado | mask_recepcion)].copy()
    pivot_por_rec = df_por_rec.groupby('PDV_lower').size().reset_index(name='Guias por Recolectar')

    por_rec_lookup = pivot_por_rec.set_index('PDV_lower')['Guias por Recolectar']
    df_report['Guias por Recolectar'] = df_report['PDV_lower'].map(por_rec_lookup).fillna(0)
    df_report['Guias Recolectadas'] = df_report['Total de Guias'] - df_report['Guias por Recolectar']
    df_report['Rate %'] = (df_report['Guias Recolectadas'] / df_report['Total de Guias']).fillna(0)
    df_report.sort_values(by=['SPV', 'ZONA', 'PDV'], ascending=[True, True, True], inplace=True)

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    wb = writer.book
    ws = wb.add_worksheet('R7 CDMX')

    # RESTORED COLORS & INCREASED FONTS
    dark_purple, light_purple, yellow, red = '#2E003E', '#480060', '#FFFF00', '#990000'
    base_font = 'Calibri'
    base_size = 14
    
    t_fmt = wb.add_format({'bold': True, 'font_size': 22, 'align': 'center', 'valign': 'vcenter', 'bg_color': dark_purple, 'font_color': 'white', 'font_name': base_font})
    s_fmt = wb.add_format({'bold': False, 'font_size': 15, 'align': 'center', 'valign': 'vcenter', 'bg_color': dark_purple, 'font_color': 'white', 'font_name': base_font})
    h_fmt = wb.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': dark_purple, 'font_color': 'white', 'border': 1, 'text_wrap': True, 'font_name': base_font})
    
    gt_l = wb.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': light_purple, 'font_color': 'white', 'border': 1, 'font_name': base_font})
    gt_v = wb.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': light_purple, 'font_color': 'white', 'border': 1, 'num_format': '#,##0', 'font_name': base_font})
    gt_p = wb.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': red, 'font_color': 'white', 'border': 1, 'num_format': '0.00%', 'font_name': base_font})
    
    sub_l = wb.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': yellow, 'font_color': 'black', 'border': 1, 'font_name': base_font})
    sub_v = wb.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': yellow, 'font_color': 'black', 'border': 1, 'num_format': '#,##0', 'font_name': base_font})
    sub_p = wb.add_format({'bold': True, 'font_size': base_size, 'align': 'center', 'valign': 'vcenter', 'bg_color': red, 'font_color': 'white', 'border': 1, 'num_format': '0.00%', 'font_name': base_font})
    
    spv_f = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'font_size': base_size, 'font_name': base_font})
    d_f = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': base_size, 'num_format': '#,##0', 'font_name': base_font})
    dp_f = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': base_size, 'num_format': '0.00%', 'font_name': base_font})

    ws.set_column('A:A', 18); ws.set_column('B:B', 20); ws.set_column('C:C', 34); ws.set_column('D:F', 24); ws.set_column('G:G', 22)
    ws.merge_range('A1:G1', 'R7 CDMX 所有平台的揽收率', t_fmt)
    ws.set_row(0, 45)
    ws.merge_range('A2:G2', subtitle_str, s_fmt)
    ws.set_row(1, 28)

    headers = ['', 'ZONA', '客户归属网点\nPDV', '当日包裹总数\nTotal de Guias', '已收取的包裹\nGuias Recolectadas', '待收取的包裹\nGuias por Recolectar', '商家拜访率\nRate %']
    for col, h in enumerate(headers): ws.write(2, col, h, h_fmt)
    ws.set_row(2, 54)

    ws.merge_range('A4:C4', 'GRAND TOTAL 总计', gt_l)
    ws.write(3, 3, df_report['Total de Guias'].sum(), gt_v)
    ws.write(3, 4, df_report['Guias Recolectadas'].sum(), gt_v)
    ws.write(3, 5, df_report['Guias por Recolectar'].sum(), gt_v)
    ws.write(3, 6, df_report['Guias Recolectadas'].sum() / df_report['Total de Guias'].sum() if df_report['Total de Guias'].sum() else 0, gt_p)
    ws.set_row(3, 32)

    c_row = 4
    for spv_name, group in df_report.groupby('SPV', sort=False):
        spv_start_row = c_row
        
        zona_start_row = c_row
        current_zona = None

        for _, row in group.iterrows():
            if current_zona != row['ZONA']:
                if current_zona is not None:
                    if c_row - 1 > zona_start_row:
                        ws.merge_range(zona_start_row, 1, c_row - 1, 1, current_zona, d_f)
                    else:
                        ws.write(zona_start_row, 1, current_zona, d_f)
                current_zona = row['ZONA']
                zona_start_row = c_row

            ws.write(c_row, 2, row['PDV'], d_f)
            ws.write(c_row, 3, row['Total de Guias'], d_f)
            ws.write(c_row, 4, row['Guias Recolectadas'], d_f)
            ws.write(c_row, 5, row['Guias por Recolectar'], d_f)
            ws.write(c_row, 6, float(row['Rate %']), dp_f)
            ws.set_row(c_row, 24)
            c_row += 1
            
        if current_zona is not None:
            if c_row - 1 > zona_start_row:
                ws.merge_range(zona_start_row, 1, c_row - 1, 1, current_zona, d_f)
            else:
                ws.write(zona_start_row, 1, current_zona, d_f)

        if c_row - 1 > spv_start_row: 
            ws.merge_range(spv_start_row, 0, c_row - 1, 0, spv_name, spv_f)
        else: 
            ws.write(spv_start_row, 0, spv_name, spv_f)

        s_tot, s_rec, s_por = group['Total de Guias'].sum(), group['Guias Recolectadas'].sum(), group['Guias por Recolectar'].sum()
        ws.merge_range(f'A{c_row+1}:C{c_row+1}', f'TOTAL {spv_name}', sub_l)
        ws.write(c_row, 3, s_tot, sub_v)
        ws.write(c_row, 4, s_rec, sub_v)
        ws.write(c_row, 5, s_por, sub_v)
        ws.write(c_row, 6, s_rec/s_tot if s_tot else 0, sub_p)
        ws.set_row(c_row, 32)
        c_row += 1

    writer.close()
    return output.getvalue(), output_filename


def generate_anomalies(raw_file_bytes, spv_file_bytes, raw_filename):
    df_spv = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv.columns = df_spv.columns.astype(str).str.strip()
    df_spv['SPV'] = df_spv['SPV'].astype(str).str.strip().str.upper()
    df_spv['ZONA'] = df_spv['ZONA'].astype(str).str.strip()
    df_spv['PDV'] = df_spv['PDV'].astype(str).str.strip()
    df_spv['PDV_lower'] = df_spv['PDV'].str.lower()
    
    valid_spvs = ['BONNIE', 'DIANA', 'LUCERO']
    spv_map = df_spv[df_spv['SPV'].isin(valid_spvs)].drop_duplicates(subset=['PDV_lower']).copy()

    xls = pd.ExcelFile(io.BytesIO(raw_file_bytes), engine='openpyxl')
    target_sheet = [s for s in xls.sheet_names if '未取件客户明细' in s or '网点明细' in s]
    sheet_to_load = target_sheet[0] if target_sheet else 0
    df_raw = pd.read_excel(io.BytesIO(raw_file_bytes), sheet_name=sheet_to_load, engine='openpyxl')
    df_raw.columns = df_raw.columns.astype(str).str.strip()

    pdv_col = [c for c in df_raw.columns if '网点' in c][0]
    pendientes_col = [c for c in df_raw.columns if '未取件订单' in c][0]
    registrados_col = [c for c in df_raw.columns if '已登记' in c][0]
    no_registrados_col = [c for c in df_raw.columns if '未登记' in c][0]

    order_col = [c for c in df_raw.columns if '订单日期' in c]
    reco_col = [c for c in df_raw.columns if '考核' in c or '收件' in c]

    order_times = pd.to_datetime(df_raw[order_col[0]], errors='coerce') if order_col else pd.Series(dtype='datetime64[ns]')
    reco_times = pd.to_datetime(df_raw[reco_col[0]], errors='coerce') if reco_col else pd.Series(dtype='datetime64[ns]')
    unique_order = sorted(order_times.dt.date.dropna().unique())
    unique_reco = sorted(reco_times.dt.date.dropna().unique())
    months_es = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}

    if unique_order and unique_reco:
        o_date, r_date = unique_order[-1], unique_reco[-1]
        pedidos_zh, pedidos_es = f"{o_date.month}月{o_date.day}日", f"{months_es[o_date.month]} {o_date.day}"
        reco_zh, reco_es = f"{r_date.month}月{r_date.day}日", f"{months_es[r_date.month]} {r_date.day}"
    else:
        today = datetime.date.today()
        r_date = today + datetime.timedelta(days=1)
        pedidos_zh, pedidos_es = f"{today.month}月{today.day}日", f"{months_es[today.month]} {today.day}"
        reco_zh, reco_es = f"{r_date.month}月{r_date.day}日", f"{months_es[r_date.month]} {r_date.day}"

    subtitle_str = f"订单日期： {pedidos_zh}, 收件日期 {reco_zh} Pedidos: {pedidos_es} | Reco: {reco_es}"
    
    output_filename = raw_filename if raw_filename.endswith('.xlsx') else f"{raw_filename.split('.')[0]}.xlsx"

    for col in [pendientes_col, registrados_col, no_registrados_col]:
        df_raw[col] = pd.to_numeric(df_raw[col], errors='coerce').fillna(0)

    pivot_total = df_raw.groupby(pdv_col).agg({pendientes_col: 'sum', registrados_col: 'sum', no_registrados_col: 'sum'}).reset_index()
    pivot_total.columns = ['PDV', 'Pendientes', 'Registrados', 'No_Registrados']
    pivot_total['PDV_lower'] = pivot_total['PDV'].astype(str).str.strip().str.lower()

    spv_lookup = spv_map.set_index('PDV_lower')
    df_report = pivot_total.copy()
    df_report['SPV'] = df_report['PDV_lower'].map(spv_lookup['SPV'])
    df_report['ZONA'] = df_report['PDV_lower'].map(spv_lookup['ZONA'])
    df_report['PDV_Clean'] = df_report['PDV_lower'].map(spv_lookup['PDV'])

    df_report = df_report[df_report['SPV'].notna()].copy()
    df_report['PDV'] = df_report['PDV_Clean'].combine_first(df_report['PDV'])
    df_report['Rate %'] = (df_report['Registrados'] / df_report['Pendientes']).fillna(0)
    df_report.sort_values(by=['SPV', 'ZONA', 'No_Registrados', 'PDV'], ascending=[True, True, False, True], inplace=True)

    wb = openpyxl.load_workbook(io.BytesIO(raw_file_bytes))
    if 'Anomalias' in wb.sheetnames: del wb['Anomalias']
    ws = wb.create_sheet(title='Anomalias', index=0)

    f_w = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
    f_b = Font(name='Calibri', size=14, bold=True, color='000000')
    f_r = Font(name='Calibri', size=14, bold=False, color='000000')
    
    fil_main = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    fil_dark_grey = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
    fil_y = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    
    a_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    b_t = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'), top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))

    c_w = {'A': 18, 'B': 20, 'C': 34, 'D': 26, 'E': 26, 'F': 26, 'G': 24}
    for col, width in c_w.items(): ws.column_dimensions[col].width = width

    ws.merge_cells('A1:G1')
    ws['A1'].value, ws['A1'].font, ws['A1'].fill, ws['A1'].alignment = '问题件跟进 Seguimiento Paquetes de Anomalia', Font(name='Calibri', size=22, bold=True, color='FFFFFF'), fil_main, a_c
    ws.row_dimensions[1].height = 45

    ws.merge_cells('A2:G2')
    ws['A2'].value, ws['A2'].font, ws['A2'].fill, ws['A2'].alignment = subtitle_str, Font(name='Calibri', size=15, bold=False, color='FFFFFF'), fil_main, a_c
    ws.row_dimensions[2].height = 28

    headers = ['', 'ZONA', '客户归属网点\nPDV', '未取件订单量合计\nPaquetes pendientes de Recoleccion', '已登记问题件量合计\nPaquetes de Anomalia Registrados', '未登记问题件量合计\nPaquetes de Anomalia NO Registrados', '问题件登记率\n% Registro de Paquetes de Anomalia']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = f_w, fil_dark_grey, a_c, b_t
    ws.row_dimensions[3].height = 65

    ws.merge_cells('A4:C4')
    ws['A4'].value, ws['A4'].font, ws['A4'].fill, ws['A4'].alignment, ws['A4'].border = 'GRAND TOTAL 总计', f_w, fil_dark_grey, a_c, b_t

    gt_p, gt_r, gt_n = df_report['Pendientes'].sum(), df_report['Registrados'].sum(), df_report['No_Registrados'].sum()
    for i, val in enumerate([gt_p, gt_r, gt_n], start=4):
        c = ws.cell(row=4, column=i, value=val)
        c.font, c.fill, c.alignment, c.number_format, c.border = f_w, fil_dark_grey, a_c, '#,##0', b_t
    c = ws.cell(row=4, column=7, value=gt_r/gt_p if gt_p else 0)
    c.font, c.fill, c.alignment, c.number_format, c.border = f_w, fil_dark_grey, a_c, '0.00%', b_t
    ws.row_dimensions[4].height = 32

    c_row = 5
    for spv_name, group in df_report.groupby('SPV', sort=False):
        spv_start_row = c_row
        
        zona_start_row = c_row
        current_zona = None

        for _, row in group.iterrows():
            ws.cell(row=c_row, column=2, value=row['ZONA']).alignment = a_c
            ws.cell(row=c_row, column=2).font = f_r
            ws.cell(row=c_row, column=2).border = b_t
            
            if current_zona != row['ZONA']:
                if current_zona is not None and c_row - 1 > zona_start_row:
                    ws.merge_cells(start_row=zona_start_row, start_column=2, end_row=c_row - 1, end_column=2)
                current_zona = row['ZONA']
                zona_start_row = c_row

            ws.cell(row=c_row, column=3, value=row['PDV']).alignment = a_c
            for idx, k, fmt in [(4, 'Pendientes', '#,##0'), (5, 'Registrados', '#,##0'), (6, 'No_Registrados', '#,##0'), (7, 'Rate %', '0.00%')]:
                c = ws.cell(row=c_row, column=idx, value=float(row[k]) if '%' in fmt else row[k])
                c.number_format, c.alignment = fmt, a_c
            for col_idx in [3, 4, 5, 6, 7]:
                ws.cell(row=c_row, column=col_idx).font = f_r
                ws.cell(row=c_row, column=col_idx).border = b_t
            ws.row_dimensions[c_row].height = 24
            c_row += 1

        if current_zona is not None and c_row - 1 > zona_start_row:
            ws.merge_cells(start_row=zona_start_row, start_column=2, end_row=c_row - 1, end_column=2)

        if c_row - 1 > spv_start_row: 
            ws.merge_cells(start_row=spv_start_row, start_column=1, end_row=c_row - 1, end_column=1)
            
        c = ws.cell(row=spv_start_row, column=1, value=spv_name)
        c.font, c.alignment, c.border = f_b, a_c, b_t

        s_p, s_r, s_n = group['Pendientes'].sum(), group['Registrados'].sum(), group['No_Registrados'].sum()
        ws.merge_cells(start_row=c_row, start_column=1, end_row=c_row, end_column=3)
        c = ws.cell(row=c_row, column=1, value=f'TOTAL {spv_name}')
        c.font, c.fill, c.alignment, c.border = f_b, fil_y, a_c, b_t
        
        ws.cell(row=c_row, column=2).border = b_t
        ws.cell(row=c_row, column=3).border = b_t

        for col, val in [(4, s_p), (5, s_r)]:
            c = ws.cell(row=c_row, column=col, value=val)
            c.font, c.fill, c.alignment, c.number_format, c.border = f_b, fil_y, a_c, '#,##0', b_t

        for col, val, fmt in [(6, s_n, '#,##0'), (7, s_r/s_p if s_p else 0, '0.00%')]:
            c = ws.cell(row=c_row, column=col, value=val)
            c.font, c.fill, c.alignment, c.number_format, c.border = f_w, fil_main, a_c, fmt, b_t
            
        ws.row_dimensions[c_row].height = 32
        c_row += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), output_filename


# ==============================================================================
# MAIN WEB UI
# ==============================================================================
if check_password():
    st.title("📊 Automated Reporting Portal")
    st.markdown("Upload your daily raw data below to instantly generate standardized Excel reports.")
    
    try:
        with open("SUPERVISORES.xlsx", "rb") as f:
            spv_file_bytes = f.read()
    except FileNotFoundError:
        st.error(
            "🚨 **CRITICAL ERROR:** `SUPERVISORES.xlsx` was not found in the root directory. "
            "Please make sure the file is saved in the same folder as `app.py` on the server."
        )
        st.stop()
        
    st.divider()

    report_type = st.selectbox(
        "Select the report you want to generate:",
        ("MISSING SCAN", "R7 CDMX", "ANOMALIES (问题件跟进)", "ALIEXPRESS")
    )
    
    raw_data = st.file_uploader("📥 Upload Daily Raw Data (.xlsx, .csv)", type=["xlsx", "csv"])
    
    if st.button("🚀 Generate Report", type="primary", use_container_width=True):
        if raw_data is None:
            st.warning("⚠️ Please upload the Daily Raw Data before proceeding.")
        else:
            with st.spinner(f'Crunching the numbers and formatting your {report_type} report...'):
                try:
                    if report_type == "MISSING SCAN":
                        processed_file, filename = generate_missing_scan(raw_data.getvalue(), spv_file_bytes)
                    elif report_type == "R7 CDMX":
                        processed_file, filename = generate_r7_cdmx(raw_data.getvalue(), spv_file_bytes)
                    elif report_type == "ANOMALIES (问题件跟进)":
                        processed_file, filename = generate_anomalies(raw_data.getvalue(), spv_file_bytes, raw_data.name)
                    elif report_type == "ALIEXPRESS":
                        processed_file, filename = generate_aliexpress(raw_data.getvalue(), spv_file_bytes)
                    
                    st.success(f"✅ Your {report_type} report was generated successfully!")
                    
                    st.download_button(
                        label="📥 Download Finished Excel Report",
                        data=processed_file,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ An error occurred while processing the file: {e}")
