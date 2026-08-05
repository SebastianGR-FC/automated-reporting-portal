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
