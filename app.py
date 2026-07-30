import streamlit as st
import pandas as pd
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import xlsxwriter

# ==============================================================================
# SECURITY: Basic Password Authentication
# ==============================================================================
def check_password():
    """Returns `True` if the user had the correct password."""
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

def generate_missing_scan(raw_file_bytes, spv_file_bytes):
    # 1. LOAD STATIC MAPPING
    df_spv_static = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv_static.columns = df_spv_static.columns.str.strip()

    spv_map = df_spv_static[['SPV', 'ZONA', 'PDV']].dropna(subset=['PDV']).copy()
    spv_map.columns = ['spv', 'zona', 'pdv']
    spv_map['spv'] = spv_map['spv'].astype(str).str.strip()
    spv_map['zona'] = spv_map['zona'].astype(str).str.strip().replace('nan', 'N/A')
    spv_map['pdv'] = spv_map['pdv'].astype(str).str.strip()
    spv_map.drop_duplicates(subset=['pdv'], inplace=True)

    # 2. LOAD RAW DATA
    try:
        df_raw = pd.read_excel(io.BytesIO(raw_file_bytes), sheet_name='sheet0', engine='openpyxl')
    except ValueError:
        df_raw = pd.read_excel(io.BytesIO(raw_file_bytes), sheet_name=0, engine='openpyxl')
    df_raw.columns = df_raw.columns.str.strip()

    # 3. DATE EXTRACTION
    raw_date = df_raw['Fecha de estadísticas'].dropna().iloc[0]
    date_obj = pd.to_datetime(raw_date)
    months_es = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}
    title_date_str = f"{date_obj.month}月{date_obj.day}日 Fecha de Recoleccion: {months_es[date_obj.month]} {date_obj.day}"
    output_filename = f"MISSING SCAN_{date_obj.strftime('%Y-%m-%d')}.xlsx"

    # 4. PIVOT & MERGE
    val_cols = ['El número de pedidos de escaneo', 'No. de escaneo faltante de recolección', 'Nº de guías con escaneo faltantes de salida']
    pivot_raw = pd.pivot_table(df_raw, index=['Nombre del nodo'], values=val_cols, aggfunc='sum').reset_index()
    pivot_raw.rename(columns={'Nombre del nodo': 'pdv'}, inplace=True)
    pivot_raw['pdv'] = pivot_raw['pdv'].astype(str).str.strip()

    final_df = pd.merge(spv_map, pivot_raw, on='pdv', how='inner')
    for col in val_cols:
        final_df[col] = final_df[col].fillna(0)

    # Strict SPV Filter
    valid_spvs = ['BONNIE', 'DIANA', 'LUCERO']
    final_df['spv_upper'] = final_df['spv'].astype(str).str.strip().str.upper()
    final_df = final_df[final_df['spv_upper'].isin(valid_spvs)].copy()

    final_df['tasa_recoleccion'] = final_df['No. de escaneo faltante de recolección'] / final_df['El número de pedidos de escaneo']
    final_df['tasa_salida'] = final_df['Nº de guías con escaneo faltantes de salida'] / final_df['El número de pedidos de escaneo']
    final_df.fillna({'tasa_recoleccion': 0, 'tasa_salida': 0}, inplace=True)
    final_df.sort_values(by=['spv', 'zona', 'pdv'], ascending=[True, True, True], inplace=True)

    # 5. OPENPYXL FORMATTING
    wb = openpyxl.load_workbook(io.BytesIO(raw_file_bytes))
    if 'Report' in wb.sheetnames:
        del wb['Report']
    ws = wb.create_sheet(title='Report', index=0)

    # Styles Definition
    font_white_bold = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    font_black_bold = Font(name='Calibri', size=11, bold=True, color='000000')
    font_green_bold = Font(name='Calibri', size=11, bold=True, color='00FF00')
    font_regular = Font(name='Calibri', size=11, bold=False, color='000000')

    fill_dark_header = PatternFill(start_color='202020', end_color='202020', fill_type='solid')
    fill_sub_header = PatternFill(start_color='4F4F4F', end_color='4F4F4F', fill_type='solid')
    fill_grand_total = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
    fill_dark_red = PatternFill(start_color='990000', end_color='990000', fill_type='solid')
    fill_yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    col_widths = {'A': 16, 'B': 16, 'C': 28, 'D': 20, 'E': 24, 'F': 24, 'G': 24, 'H': 24}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.merge_cells('A1:H1')
    cell_title = ws['A1']
    cell_title.value = f'缺失扫描报告 OMISIÓN DE ESCANEO\n揽收日期 {title_date_str}'
    cell_title.font = Font(name='Calibri', size=13, bold=True, color='FFFFFF')
    cell_title.fill = fill_dark_header
    cell_title.alignment = align_center
    ws.row_dimensions[1].height = 36

    headers = ['负责人\n\nSPV', '区域\n\nZONA', '揽收网点\n\nPDV', '总扫描数据量\n\nTOTAL A\nESCANEAR', '未进行入库扫描包裹件数\n\nSIN ESCANEO DE\nRECOLECCION', '未进行入库扫描包裹的百分比\n\n% SIN ESCANEO DE\nRECOLECCION', '未进行出库扫描的包裹\n\nSIN ESCANEO DE\nSALIDA', '未进行出库扫描包裹的百分比\n\n% SIN ESCANEO DE\nSALIDA']
    for col_num, h_text in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_num, value=h_text)
        c.font = font_white_bold
        c.fill = fill_sub_header
        c.alignment = align_center
        c.border = thin_border
    ws.row_dimensions[2].height = 55

    gt_total = final_df['El número de pedidos de escaneo'].sum()
    gt_rec = final_df['No. de escaneo faltante de recolección'].sum()
    gt_rec_pct = gt_rec / gt_total if gt_total else 0
    gt_sal = final_df['Nº de guías con escaneo faltantes de salida'].sum()
    gt_sal_pct = gt_sal / gt_total if gt_total else 0

    ws.merge_cells('A3:C3')
    ws['A3'].value, ws['A3'].font, ws['A3'].fill, ws['A3'].alignment = '总计', font_white_bold, fill_grand_total, align_center
    
    for col_idx, val, num_fmt, f_style, fill_style in [(4, gt_total, '#,##0', font_green_bold, fill_grand_total), (5, gt_rec, '#,##0', font_green_bold, fill_grand_total), (6, gt_rec_pct, '0.00%', font_white_bold, fill_dark_red), (7, gt_sal, '#,##0', font_green_bold, fill_grand_total), (8, gt_sal_pct, '0.00%', font_white_bold, fill_dark_red)]:
        c = ws.cell(row=3, column=col_idx, value=val)
        c.font, c.fill, c.alignment, c.number_format, c.border = f_style, fill_style, align_center, num_fmt, thin_border
    ws.row_dimensions[3].height = 28

    current_row = 4
    for spv_name, group in final_df.groupby('spv', sort=False):
        start_row = current_row
        for _, row in group.iterrows():
            ws.cell(row=current_row, column=2, value=row['zona']).alignment = align_center
            ws.cell(row=current_row, column=3, value=row['pdv']).alignment = align_center
            for c_idx, key, fmt in [(4, 'El número de pedidos de escaneo', '#,##0'), (5, 'No. de escaneo faltante de recolección', '#,##0'), (6, 'tasa_recoleccion', '0.00%'), (7, 'Nº de guías con escaneo faltantes de salida', '#,##0'), (8, 'tasa_salida', '0.00%')]:
                c = ws.cell(row=current_row, column=c_idx, value=float(row[key]) if '%' in fmt else row[key])
                c.number_format, c.alignment = fmt, align_center
            for col_idx in range(2, 9):
                ws.cell(row=current_row, column=col_idx).font = font_regular
                ws.cell(row=current_row, column=col_idx).border = thin_border
            current_row += 1

        if current_row - 1 > start_row: ws.merge_cells(start_row=start_row, start_column=1, end_row=current_row - 1, end_column=1)
        ws.cell(row=start_row, column=1, value=spv_name).font = font_black_bold
        ws.cell(row=start_row, column=1).alignment = align_center
        ws.cell(row=start_row, column=1).border = thin_border

        sub_total = group['El número de pedidos de escaneo'].sum()
        sub_rec = group['No. de escaneo faltante de recolección'].sum()
        sub_rec_pct = sub_rec / sub_total if sub_total else 0
        sub_sal = group['Nº de guías con escaneo faltantes de salida'].sum()
        sub_sal_pct = sub_sal / sub_total if sub_total else 0

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        sub = ws.cell(row=current_row, column=1, value=f'TOTAL {spv_name}')
        sub.font, sub.fill, sub.alignment = font_black_bold, fill_yellow, align_center

        for col_idx, val, num_fmt, f_style, fill_style in [(4, sub_total, '#,##0', font_black_bold, fill_yellow), (5, sub_rec, '#,##0', font_black_bold, fill_yellow), (6, sub_rec_pct, '0.00%', font_white_bold, fill_dark_red), (7, sub_sal, '#,##0', font_black_bold, fill_yellow), (8, sub_sal_pct, '0.00%', font_white_bold, fill_dark_red)]:
            c = ws.cell(row=current_row, column=col_idx, value=val)
            c.font, c.fill, c.alignment, c.number_format, c.border = f_style, fill_style, align_center, num_fmt, thin_border
        current_row += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), output_filename

def generate_r5_metropolitan(raw_file_bytes, spv_file_bytes):
    # 1. LOAD STATIC MAPPING
    df_spv = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv.columns = df_spv.columns.astype(str).str.strip()
    df_spv['SPV'] = df_spv['SPV'].astype(str).str.strip().str.upper()
    df_spv['ZONA'] = df_spv['ZONA'].astype(str).str.strip()
    df_spv['PDV'] = df_spv['PDV'].astype(str).str.strip()
    df_spv['PDV_lower'] = df_spv['PDV'].str.lower()
    
    valid_spvs = ['BONNIE', 'DIANA', 'LUCERO']
    spv_map = df_spv[df_spv['SPV'].isin(valid_spvs)].drop_duplicates(subset=['PDV_lower']).copy()

    # 2. LOAD RAW DATA
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

    # 3. DATE EXTRACTION
    order_times = pd.to_datetime(df_raw['Tiempo de registro de la orden'], errors='coerce')
    unique_dates = sorted(order_times.dt.date.dropna().unique())
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
    output_filename = f"R5 Metropolitan_{reco_date.strftime('%Y-%m-%d')}.xlsx"

    # 4. PIVOT & MERGE
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

    # 5. XLSXWRITER FORMATTING
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    wb = writer.book
    ws = wb.add_worksheet('R5 METROPOLITAN')

    dark_purple, light_purple, yellow, red = '#2E003E', '#480060', '#FFFF00', '#990000'
    t_fmt = wb.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'bg_color': dark_purple, 'font_color': 'white'})
    s_fmt = wb.add_format({'bold': False, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'bg_color': dark_purple, 'font_color': 'white'})
    h_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': dark_purple, 'font_color': 'white', 'border': 1, 'text_wrap': True})
    
    gt_l = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': light_purple, 'font_color': 'white', 'border': 1})
    gt_v = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': light_purple, 'font_color': 'white', 'border': 1, 'num_format': '#,##0'})
    gt_p = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': red, 'font_color': 'white', 'border': 1, 'num_format': '0.00%'})
    
    sub_l = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': yellow, 'font_color': 'black', 'border': 1})
    sub_v = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': yellow, 'font_color': 'black', 'border': 1, 'num_format': '#,##0'})
    sub_p = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': red, 'font_color': 'white', 'border': 1, 'num_format': '0.00%'})
    
    spv_f = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True})
    d_f = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
    dp_f = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.00%'})

    ws.set_column('A:A', 16); ws.set_column('B:B', 18); ws.set_column('C:C', 32); ws.set_column('D:F', 20); ws.set_column('G:G', 18)
    ws.merge_range('A1:G1', 'R5 METROPOLITAN 所有平台的揽收率', t_fmt)
    ws.set_row(0, 36)
    ws.merge_range('A2:G2', subtitle_str, s_fmt)
    ws.set_row(1, 22)

    headers = ['', 'ZONA', '客户归属网点\nPDV', '当日包裹总数\nTotal de Guias', '已收取的包裹\nGuias Recolectadas', '待收取的包裹\nGuias por Recolectar', '商家拜访率\nRate %']
    for col, h in enumerate(headers): ws.write(2, col, h, h_fmt)
    ws.set_row(2, 45)

    ws.merge_range('A4:C4', 'GRAND TOTAL 总计', gt_l)
    ws.write(3, 3, df_report['Total de Guias'].sum(), gt_v)
    ws.write(3, 4, df_report['Guias Recolectadas'].sum(), gt_v)
    ws.write(3, 5, df_report['Guias por Recolectar'].sum(), gt_v)
    ws.write(3, 6, df_report['Guias Recolectadas'].sum() / df_report['Total de Guias'].sum() if df_report['Total de Guias'].sum() else 0, gt_p)
    ws.set_row(3, 28)

    c_row = 4
    for spv_name, group in df_report.groupby('SPV', sort=False):
        s_row = c_row
        for _, row in group.iterrows():
            ws.write(c_row, 1, row['ZONA'], d_f)
            ws.write(c_row, 2, row['PDV'], d_f)
            ws.write(c_row, 3, row['Total de Guias'], d_f)
            ws.write(c_row, 4, row['Guias Recolectadas'], d_f)
            ws.write(c_row, 5, row['Guias por Recolectar'], d_f)
            ws.write(c_row, 6, float(row['Rate %']), dp_f)
            c_row += 1

        if c_row - 1 > s_row: ws.merge_range(s_row, 0, c_row - 1, 0, spv_name, spv_f)
        else: ws.write(s_row, 0, spv_name, spv_f)

        s_tot, s_rec, s_por = group['Total de Guias'].sum(), group['Guias Recolectadas'].sum(), group['Guias por Recolectar'].sum()
        ws.merge_range(f'A{c_row+1}:C{c_row+1}', f'TOTAL {spv_name}', sub_l)
        ws.write(c_row, 3, s_tot, sub_v)
        ws.write(c_row, 4, s_rec, sub_v)
        ws.write(c_row, 5, s_por, sub_v)
        ws.write(c_row, 6, s_rec/s_tot if s_tot else 0, sub_p)
        c_row += 1

    writer.close()
    return output.getvalue(), output_filename

def generate_anomalies(raw_file_bytes, spv_file_bytes):
    # 1. LOAD STATIC MAPPING
    df_spv = pd.read_excel(io.BytesIO(spv_file_bytes))
    df_spv.columns = df_spv.columns.astype(str).str.strip()
    df_spv['SPV'] = df_spv['SPV'].astype(str).str.strip().str.upper()
    df_spv['ZONA'] = df_spv['ZONA'].astype(str).str.strip()
    df_spv['PDV'] = df_spv['PDV'].astype(str).str.strip()
    df_spv['PDV_lower'] = df_spv['PDV'].str.lower()
    
    valid_spvs = ['BONNIE', 'DIANA', 'LUCERO']
    spv_map = df_spv[df_spv['SPV'].isin(valid_spvs)].drop_duplicates(subset=['PDV_lower']).copy()

    # 2. LOAD RAW DATA
    xls = pd.ExcelFile(io.BytesIO(raw_file_bytes), engine='openpyxl')
    target_sheet = [s for s in xls.sheet_names if '未取件客户明细' in s or '网点明细' in s]
    sheet_to_load = target_sheet[0] if target_sheet else 0
    df_raw = pd.read_excel(io.BytesIO(raw_file_bytes), sheet_name=sheet_to_load, engine='openpyxl')
    df_raw.columns = df_raw.columns.astype(str).str.strip()

    pdv_col = [c for c in df_raw.columns if '网点' in c][0]
    pendientes_col = [c for c in df_raw.columns if '未取件订单' in c][0]
    registrados_col = [c for c in df_raw.columns if '已登记' in c][0]
    no_registrados_col = [c for c in df_raw.columns if '未登记' in c][0]

    # 3. DATE EXTRACTION
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
    output_filename = f"Automated_Anomalias_Report_{r_date.strftime('%Y-%m-%d')}.xlsx"

    # 4. PIVOT & MERGE
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

    # 5. OPENPYXL FORMATTING
    wb = openpyxl.load_workbook(io.BytesIO(raw_file_bytes))
    if 'Anomalias' in wb.sheetnames: del wb['Anomalias']
    ws = wb.create_sheet(title='Anomalias', index=0)

    f_w = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    f_b = Font(name='Calibri', size=11, bold=True, color='000000')
    f_r = Font(name='Calibri', size=11, bold=False, color='000000')
    fil_r = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    fil_y = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    a_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    b_t = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    c_w = {'A': 16, 'B': 18, 'C': 32, 'D': 24, 'E': 24, 'F': 24, 'G': 20}
    for col, width in c_w.items(): ws.column_dimensions[col].width = width

    ws.merge_cells('A1:G1')
    ws['A1'].value, ws['A1'].font, ws['A1'].fill, ws['A1'].alignment = '问题件跟进 Seguimiento Paquetes de Anomalia', Font(name='Calibri', size=18, bold=True, color='FFFFFF'), fil_r, a_c
    ws.row_dimensions[1].height = 36

    ws.merge_cells('A2:G2')
    ws['A2'].value, ws['A2'].font, ws['A2'].fill, ws['A2'].alignment = subtitle_str, Font(name='Calibri', size=11, bold=False, color='FFFFFF'), fil_r, a_c
    ws.row_dimensions[2].height = 22

    headers = ['', 'ZONA', '客户归属网点\nPDV', '未取件订单量合计\nPaquetes pendientes de\nRecoleccion', '已登记问题件量合计\nPaquetes de Anomalia\nRegistrados', '未登记问题件量合计\nPaquetes de Anomalia NO\nRegistrados', '问题件登记率\n% Registro de\nPaquetes\nde Anomalia']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = f_w, fil_r, a_c, b_t
    ws.row_dimensions[3].height = 65

    ws.merge_cells('A4:C4')
    ws['A4'].value, ws['A4'].font, ws['A4'].fill, ws['A4'].alignment = 'TOTALES               总计', f_w, fil_r, a_c

    gt_p, gt_r, gt_n = df_report['Pendientes'].sum(), df_report['Registrados'].sum(), df_report['No_Registrados'].sum()
    for i, val in enumerate([gt_p, gt_r, gt_n], start=4):
        c = ws.cell(row=4, column=i, value=val)
        c.font, c.fill, c.alignment, c.number_format, c.border = f_w, fil_r, a_c, '#,##0', b_t
    c = ws.cell(row=4, column=7, value=gt_r/gt_p if gt_p else 0)
    c.font, c.fill, c.alignment, c.number_format, c.border = f_w, fil_r, a_c, '0.00%', b_t
    ws.row_dimensions[4].height = 28

    c_row = 5
    for spv_name, group in df_report.groupby('SPV', sort=False):
        s_row = c_row
        for _, row in group.iterrows():
            ws.cell(row=c_row, column=2, value=row['ZONA']).alignment = a_c
            ws.cell(row=c_row, column=3, value=row['PDV']).alignment = a_c
            for idx, k, fmt in [(4, 'Pendientes', '#,##0'), (5, 'Registrados', '#,##0'), (6, 'No_Registrados', '#,##0'), (7, 'Rate %', '0.00%')]:
                c = ws.cell(row=c_row, column=idx, value=float(row[k]) if '%' in fmt else row[k])
                c.number_format, c.alignment = fmt, a_c
            for col_idx in range(2, 8):
                ws.cell(row=c_row, column=col_idx).font = f_r
                ws.cell(row=c_row, column=col_idx).border = b_t
            c_row += 1

        if c_row - 1 > s_row: ws.merge_cells(start_row=s_row, start_column=1, end_row=c_row - 1, end_column=1)
        c = ws.cell(row=s_row, column=1, value=spv_name)
        c.font, c.alignment, c.border = f_b, a_c, b_t

        s_p, s_r, s_n = group['Pendientes'].sum(), group['Registrados'].sum(), group['No_Registrados'].sum()
        ws.merge_cells(start_row=c_row, start_column=1, end_row=c_row, end_column=3)
        c = ws.cell(row=c_row, column=1, value=f'TOTAL {spv_name}')
        c.font, c.fill, c.alignment = f_b, fil_y, a_c

        for col, val in [(4, s_p), (5, s_r)]:
            c = ws.cell(row=c_row, column=col, value=val)
            c.font, c.fill, c.alignment, c.number_format, c.border = f_b, fil_y, a_c, '#,##0', b_t

        for col, val, fmt in [(6, s_n, '#,##0'), (7, s_r/s_p if s_p else 0, '0.00%')]:
            c = ws.cell(row=c_row, column=col, value=val)
            c.font, c.fill, c.alignment, c.number_format, c.border = f_w, fil_r, a_c, fmt, b_t
        c_row += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), output_filename

# ==============================================================================
# MAIN WEB UI
# ==============================================================================
if check_password():
    st.set_page_config(page_title="Automated Reporting Portal", page_icon="📊", layout="centered")
    
    st.title("📊 Automated Reporting Portal")
    st.markdown("Upload your daily raw data below to instantly generate standardized Excel reports.")
    
    # --------------------------------------------------------------------------
    # SILENT STATIC FILE LOADER
    # Reads SUPERVISORES.xlsx directly from the local root directory
    # --------------------------------------------------------------------------
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

    # 1. Report Selection
    report_type = st.selectbox(
        "Select the report you want to generate:",
        ("MISSING SCAN", "R5 METROPOLITAN", "ANOMALIES (问题件跟进)")
    )
    
    # 2. Single Daily File Uploader
    raw_data = st.file_uploader("📥 Upload Daily Raw Data (.xlsx, .csv)", type=["xlsx", "csv"])
    
    # 3. Process Button
    if st.button("🚀 Generate Report", type="primary", use_container_width=True):
        if raw_data is None:
            st.warning("⚠️ Please upload the Daily Raw Data before proceeding.")
        else:
            with st.spinner(f'Crunching the numbers and formatting your {report_type} report...'):
                try:
                    # Pass raw_data.getvalue() and the auto-loaded spv_file_bytes
                    if report_type == "MISSING SCAN":
                        processed_file, filename = generate_missing_scan(raw_data.getvalue(), spv_file_bytes)
                    elif report_type == "R5 METROPOLITAN":
                        processed_file, filename = generate_r5_metropolitan(raw_data.getvalue(), spv_file_bytes)
                    elif report_type == "ANOMALIES (问题件跟进)":
                        processed_file, filename = generate_anomalies(raw_data.getvalue(), spv_file_bytes)
                    
                    st.success(f"✅ Your {report_type} report was generated successfully!")
                    
                    # 4. Download Button
                    st.download_button(
                        label="📥 Download Finished Excel Report",
                        data=processed_file,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ An error occurred while processing the file: {e}")