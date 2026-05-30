import streamlit as st
import pandas as pd
import graphviz
import sqlite3
import glob
import os
import re
import math
import json
import openpyxl
from data_pipeline import process_data, load_saved_mappings, save_new_mapping
from db_manager import init_db, insert_data, execute_query, execute_update
from ai_agent import process_ai_chat  # Tích hợp AI từ file ngoài

# ── YARN PARSER ─────────────────────────────────────────────────
_yarn_import_error = None
try:
    from yarn_parser import init_yarn_table, scan_folder, parse_excel_file, upsert_records
    init_yarn_table()
except Exception as _ye:
    _yarn_import_error = str(_ye)
    def scan_folder(*a, **k): return 0, 0, [f"Lỗi: {_ye}"]
    def parse_excel_file(*a, **k): return []
    def upsert_records(*a, **k): return 0
# ─────────────────────────────────────────────────────────────────

# ── BEAM INFO ─────────────────────────────────────────────────────
_beam_import_error = None
try:
    from beam_info import init_beam_tables, import_beam_file, get_beam_on_machine
    init_beam_tables()
except Exception as _be:
    _beam_import_error = str(_be)
    def import_beam_file(*a, **k): return {"xuatkho":0,"yccb":0,"total":0,"errors":[str(_be)]}
    def get_beam_on_machine(*a, **k): return {}
# ─────────────────────────────────────────────────────────────────

# --- HÀM CAN THIỆP LÕI DATABASE CHỐNG LỖI ---
def run_db_command(query, params=()):
    db_files = glob.glob("*.db")
    if not db_files:
        for root, dirs, files in os.walk("."):
            for file in files:
                if file.endswith(".db"):
                    db_files.append(os.path.join(root, file))
    
    success = False
    for db in db_files:
        try:
            with sqlite3.connect(db) as conn:
                conn.execute(query, params)
                conn.commit()
                success = True
        except:
            pass
    return success

# Khởi tạo Không gian ảo và Bảng Cấu hình
run_db_command("CREATE TABLE IF NOT EXISTS Yarn_Dictionary (yarn_type TEXT PRIMARY KEY, coefficient REAL)")
run_db_command("CREATE TABLE IF NOT EXISTS AI_Rules (id INTEGER PRIMARY KEY AUTOINCREMENT, rule_text TEXT)")
run_db_command("CREATE TABLE IF NOT EXISTS AutoSync_Config (id INTEGER PRIMARY KEY, folder_path TEXT, cluster_name TEXT, skip_rows INTEGER, keyword_filter TEXT, mapping_json TEXT, template_name TEXT, sheet_name TEXT)")

try: run_db_command("ALTER TABLE AutoSync_Config ADD COLUMN keyword_filter TEXT")
except: pass
try: run_db_command("ALTER TABLE AutoSync_Config ADD COLUMN mapping_json TEXT")
except: pass
try: run_db_command("ALTER TABLE AutoSync_Config ADD COLUMN template_name TEXT")
except: pass
try: run_db_command("ALTER TABLE AutoSync_Config ADD COLUMN sheet_name TEXT")
except: pass

# --- TÌM KIẾM XUYÊN THẤU + NHẬN DIỆN TỪ KHÓA ---
def get_all_excel_csv_files(base_folder, keyword_filter=""):
    all_files = []
    keywords = [k.strip().lower() for k in keyword_filter.split(",")] if keyword_filter else []
    
    if os.path.exists(base_folder):
        for root, dirs, files in os.walk(base_folder):
            for file in files:
                if file.endswith(('.csv', '.xls', '.xlsx')) and not file.startswith('~$'):
                    full_path = os.path.join(root, file)
                    if keywords:
                        if any(k in full_path.lower() for k in keywords):
                            all_files.append(full_path)
                    else:
                        all_files.append(full_path)
    return all_files

# --- NẠP DỮ LIỆU + ĐỌC MULTI-SHEET + TỰ MÒ NĂM + TRÁM NGÀY ---
def run_multi_file_pipeline(u_files, mapping, cluster, s_name, s_rows, auto_sync, is_local):
    try:
        total_rows = 0
        sheet_zone_map = mapping.get('__SHEET_ZONE_MAP__', {})

        for f in u_files:
            fname = os.path.basename(f) if is_local else f.name
            f_path = f if is_local else ""
            
            if not is_local: f.seek(0)
            
            extracted_year = str(pd.Timestamp.now().year) 
            if is_local:
                year_match = re.search(r'\\(20\d{2})\\', f_path)
                if not year_match: year_match = re.search(r'/(20\d{2})/', f_path)
                if year_match: extracted_year = year_match.group(1)

            dfs_to_process = {} 
            try:
                if fname.endswith('.csv'): 
                    dfs_to_process["CSV"] = pd.read_csv(f, skiprows=s_rows)
                else: 
                    xls = pd.ExcelFile(f)
                    if s_name == "ALL_SHEETS":
                        for sheet in xls.sheet_names:
                            dfs_to_process[sheet] = pd.read_excel(xls, sheet_name=sheet, skiprows=s_rows)
                    else:
                        dfs_to_process[s_name] = pd.read_excel(xls, sheet_name=s_name, skiprows=s_rows)
            except Exception as e: 
                continue 
            
            for current_sheet_name, df in dfs_to_process.items():
                if df.empty: continue

                df = df.loc[:, ~df.columns.duplicated()].copy()

                mapped_zone = sheet_zone_map.get(current_sheet_name, current_sheet_name)
                if mapped_zone == "-- Bỏ qua --":
                    continue

                cols_to_keep = []
                rename_dict = {}

                for src, tgt in mapping.items():
                    if src.startswith('__') or src == "AUTO_SHEET_NAME": continue
                    if src in df.columns:
                        if tgt != "-- Bỏ qua --":
                            cols_to_keep.append(src)
                            if tgt != "Giữ nguyên tên cột gốc (Copy từ trái)":
                                rename_dict[src] = tgt

                if not cols_to_keep:
                    continue 

                df_mapped = df[cols_to_keep].copy()
                df_mapped.rename(columns=rename_dict, inplace=True)
                
                df_mapped = df_mapped.loc[:, ~df_mapped.columns.duplicated()]
                
                df_mapped['sub_location'] = mapped_zone
                
                check_cols = [c for c in df_mapped.columns if c != 'sub_location']
                if not check_cols: continue
                df_mapped = df_mapped.dropna(subset=check_cols, how='all')
                if df_mapped.empty: continue
                
                if 'date_month' in df_mapped.columns and 'date_day' in df_mapped.columns:
                    s_year = df_mapped.get('date_year', pd.Series(extracted_year, index=df_mapped.index)).astype(str).replace(r'(?i)none|nan|^$', extracted_year, regex=True)
                    s_month = df_mapped['date_month'].astype(str).replace(r'(?i)none|nan|^$', pd.NA, regex=True).ffill().bfill()
                    s_day = df_mapped['date_day'].astype(str).replace(r'(?i)none|nan|^$', pd.NA, regex=True).ffill().bfill()

                    s_month = s_month.str.extract(r'(\d+)')[0].str.zfill(2)
                    s_day = s_day.str.extract(r'(\d+)')[0].str.zfill(2)

                    valid_mask = s_month.notna() & s_day.notna()
                    df_mapped['date'] = pd.NaT
                    df_mapped.loc[valid_mask, 'date'] = pd.to_datetime(s_year[valid_mask] + '-' + s_month[valid_mask] + '-' + s_day[valid_mask], errors='coerce')
                    df_mapped['date'] = df_mapped['date'].dt.strftime('%Y-%m-%d')
                    df_mapped.drop(columns=['date_month', 'date_day', 'date_year'], errors='ignore', inplace=True)
                elif 'date' in df_mapped.columns:
                    df_mapped['date'] = df_mapped['date'].astype(str).replace(r'(?i)none|nan|^$', pd.NA, regex=True).ffill().bfill()
                    parsed = pd.to_datetime(df_mapped['date'], errors='coerce', dayfirst=True)
                    if parsed.isna().all():
                        parsed = pd.to_datetime(df_mapped['date'], errors='coerce', dayfirst=False)
                    df_mapped['date'] = parsed.dt.strftime('%Y-%m-%d')

                if 'date' not in df_mapped.columns:
                    df_mapped['date'] = pd.Timestamp.now().strftime('%Y-%m-%d')
                else:
                    df_mapped['date'] = df_mapped['date'].astype(str).replace(r'(?i)none|nan|nat|^$', pd.NA, regex=True)
                    if df_mapped['date'].isna().all():
                        df_mapped['date'] = pd.Timestamp.now().strftime('%Y-%m-%d')
                    else:
                        df_mapped['date'] = df_mapped['date'].fillna(pd.Timestamp.now().strftime('%Y-%m-%d'))
                    
                df_mapped['cluster_name'] = cluster
                has_inbound = 'inbound_date' in df_mapped.columns
                has_outbound = 'outbound_date' in df_mapped.columns
                dfs_to_insert = []
                
                if has_inbound:
                    df_in = df_mapped.dropna(subset=['inbound_date']).copy()
                    if not df_in.empty:
                        df_in['date'] = pd.to_datetime(df_in['inbound_date'], errors='coerce').ffill().bfill().dt.strftime('%Y-%m-%d')
                        df_in['type'] = 'NHAP'
                        dfs_to_insert.append(df_in)
                if has_outbound:
                    df_out = df_mapped.dropna(subset=['outbound_date']).copy()
                    if not df_out.empty:
                        df_out['date'] = pd.to_datetime(df_out['outbound_date'], errors='coerce').ffill().bfill().dt.strftime('%Y-%m-%d')
                        df_out['type'] = 'XUAT'
                        dfs_to_insert.append(df_out)
                if not has_inbound and not has_outbound:
                    if 'date' not in df_mapped.columns or df_mapped['date'].isna().all(): 
                        df_mapped['date'] = pd.Timestamp.now().strftime('%Y-%m-%d')
                    if cluster == "Xưởng Dệt": df_mapped['type'] = 'SAN_XUAT'
                    else: df_mapped['type'] = 'NHAP'
                    dfs_to_insert.append(df_mapped)
                    
                for final_df in dfs_to_insert:
                    clean_df = final_df.drop(columns=['inbound_date', 'outbound_date'], errors='ignore')
                    
                    clean_df = clean_df.loc[:, ~clean_df.columns.duplicated()]
                    
                    for col in clean_df.columns:
                        clean_df[col] = clean_df[col].apply(lambda x: "" if pd.isna(x) else str(x))
                    
                    try:
                        db_files = glob.glob("*.db")
                        if db_files:
                            with sqlite3.connect(db_files[0]) as conn:
                                cursor = conn.execute("PRAGMA table_info(Inventory_Log)")
                                existing_cols = [row[1] for row in cursor.fetchall()]
                                for c in clean_df.columns:
                                    if c not in existing_cols:
                                        conn.execute(f"ALTER TABLE Inventory_Log ADD COLUMN '{c}' TEXT")
                                conn.commit()
                    except: pass
                            
                    insert_data(clean_df)
                    total_rows += len(clean_df)

        st.cache_data.clear() 
        return total_rows, len(u_files)
    except Exception as e: 
        return 0, str(e)

# ==========================================
# HÀM LỌC BẢNG DỮ LIỆU THÔNG MINH
# ==========================================
def get_clean_cluster_data(df, cluster_name):
    if df.empty: return df
    df_clean = df.loc[:, ~df.columns.duplicated()].copy()
    
    try:
        valid_cols = ['id', 'date', 'cluster_name', 'sub_location', 'type']
        has_mapping = False
        
        db_files = glob.glob("*.db")
        if db_files:
            with sqlite3.connect(db_files[0]) as conn:
                df_cfg = pd.read_sql_query(f"SELECT mapping_json FROM AutoSync_Config WHERE cluster_name='{cluster_name}' LIMIT 1", conn)
                if not df_cfg.empty and df_cfg.iloc[0]['mapping_json']:
                    mapping_str = str(df_cfg.iloc[0]['mapping_json'])
                    if mapping_str.strip() not in ["", "{}"]:
                        mapping_data = json.loads(mapping_str)
                        has_mapping = True
                        for k, v in mapping_data.items():
                            if k.startswith('__') or k == "AUTO_SHEET_NAME": continue
                            if isinstance(v, str):
                                if v == "Giữ nguyên tên cột gốc (Copy từ trái)":
                                    valid_cols.append(k)
                                elif v != "-- Bỏ qua --":
                                    valid_cols.append(v)
        
        cols_to_keep = []
        for c in df_clean.columns:
            if has_mapping:
                if c in valid_cols:
                    cols_to_keep.append(c)
            else:
                if c in valid_cols:
                    cols_to_keep.append(c)
                else:
                    col_str = df_clean[c].astype(str).str.strip().str.lower()
                    is_empty = col_str.isin(['', 'nan', 'none', '<na>', 'null'])
                    if not is_empty.all():
                        cols_to_keep.append(c)
                        
        safe_cols = []
        seen = set()
        for c in cols_to_keep:
            c_str = str(c)
            if c_str not in seen:
                seen.add(c_str)
                safe_cols.append(c_str)
        cols_to_keep = safe_cols
        
        cols_to_keep = [c for c in cols_to_keep if c in df_clean.columns]
        
        df_clean = df_clean[cols_to_keep].fillna("")
        
        subset_cols = [c for c in df_clean.columns if c != 'id']
        if subset_cols:
            df_clean = df_clean.drop_duplicates(subset=subset_cols)
            
        return df_clean
    except Exception as e:
        return df_clean.fillna("")

# HÀM RADAR TÌM CỘT TỰ ĐỘNG CHO DASHBOARD
def find_best_col(df, keywords):
    for k in keywords:
        for col in df.columns:
            if k in str(col).lower():
                return col
    return None

# 1. THIẾT LẬP TRANG CHÍNH
st.set_page_config(page_title="TexFlow Systems", page_icon="🌊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
/* ===== FONT IMPORT ===== */
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

/* ===== ROOT VARIABLES ===== */
:root {
    --bg-base:       #F5F7FA;
    --bg-surface:    #FFFFFF;
    --bg-card:       #FFFFFF;
    --bg-hover:      #EEF2F7;
    --border:        #D1DBE8;
    --border-bright: #A8BDD4;
    --accent:        #1A6FBF;
    --accent2:       #5B5FEF;
    --accent3:       #12A068;
    --danger:        #DC3545;
    --text-primary:  #1A2535;
    --text-secondary:#4A6080;
    --text-muted:    #8098B4;
    --glow:          0 2px 12px rgba(26,111,191,0.12);
    --shadow:        0 1px 4px rgba(0,0,0,0.08), 0 4px 16px rgba(0,0,0,0.06);
    --radius:        10px;
    --radius-sm:     7px;
    --font:          "Be Vietnam Pro", sans-serif;
    --mono:          "JetBrains Mono", monospace;
}

/* ===== GLOBAL ===== */
html, body, [class*="css"] {
    font-family: var(--font) !important;
    color: var(--text-primary) !important;
}
.stApp {
    background: var(--bg-base) !important;
}

/* ===== SIDEBAR ===== */
[data-testid="stSidebar"] {
    background: #FFFFFF !important;
    border-right: 1px solid var(--border) !important;
    box-shadow: 2px 0 12px rgba(0,0,0,0.06) !important;
}
[data-testid="stSidebar"] .stMarkdown h1 {
    font-size: 1.3rem !important;
    font-weight: 800 !important;
    color: var(--accent) !important;
    letter-spacing: -0.3px !important;
}
[data-testid="stSidebar"] .stMarkdown p {
    color: var(--text-secondary) !important;
    font-size: 0.78rem !important;
}
[data-testid="stSidebar"] hr { border-color: var(--border) !important; margin: 10px 0 !important; }

/* Nav radio */
[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label {
    background: transparent !important;
    border: 1px solid transparent !important;
    border-radius: var(--radius-sm) !important;
    padding: 9px 13px !important;
    margin: 2px 0 !important;
    transition: all 0.15s !important;
    color: var(--text-secondary) !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
}
[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label:hover {
    background: var(--bg-hover) !important;
    color: var(--accent) !important;
}
[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label[data-checked="true"] {
    background: #EAF2FC !important;
    border-color: #A8CBEC !important;
    color: var(--accent) !important;
    font-weight: 600 !important;
}

/* Sidebar AI settings */
[data-testid="stSidebar"] .stSubheader {
    color: var(--text-secondary) !important;
    font-size: 0.8rem !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] .stTextInput > div > div > input {
    background: #F0F5FB !important;
    border: 1px solid var(--border-bright) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--accent) !important;
    font-family: var(--mono) !important;
    font-size: 0.82rem !important;
}
[data-testid="stSidebar"] .stTextInput label { color: var(--text-muted) !important; font-size: 0.75rem !important; }

/* Sidebar delete button */
[data-testid="stSidebar"] .stButton > button {
    background: #FFF0F0 !important;
    border: 1px solid #FFCCCC !important;
    color: var(--danger) !important;
    font-size: 0.8rem !important;
    border-radius: var(--radius-sm) !important;
    font-weight: 500 !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: #FFE0E0 !important;
    border-color: var(--danger) !important;
}

/* ===== MAIN TITLES ===== */
h1 {
    font-size: 1.6rem !important;
    font-weight: 800 !important;
    color: var(--text-primary) !important;
    padding-bottom: 8px !important;
    border-bottom: 2px solid var(--accent) !important;
    margin-bottom: 1.5rem !important;
}
h2 { font-weight: 700 !important; color: var(--text-primary) !important; }
h3 { font-size: 1rem !important; font-weight: 600 !important; color: var(--text-secondary) !important; }

/* ===== CONTAINERS ===== */
[data-testid="stExpander"] {
    background: var(--bg-card) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    box-shadow: var(--shadow) !important;
    overflow: hidden !important;
}
[data-testid="stExpander"] summary {
    background: #FAFCFF !important;
    padding: 13px 16px !important;
    font-weight: 600 !important;
    font-size: 0.92rem !important;
    color: var(--text-primary) !important;
    border-bottom: 1px solid var(--border) !important;
}
[data-testid="stExpander"] > div:last-child {
    background: var(--bg-surface) !important;
    padding: 16px !important;
}

/* ===== METRICS ===== */
[data-testid="metric-container"] {
    background: var(--bg-card) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 16px !important;
    box-shadow: var(--shadow) !important;
    transition: all 0.15s !important;
}
[data-testid="metric-container"]:hover {
    border-color: var(--accent) !important;
    box-shadow: var(--glow) !important;
    transform: translateY(-1px) !important;
}
[data-testid="metric-container"] [data-testid="stMetricLabel"] {
    color: var(--text-secondary) !important;
    font-size: 0.8rem !important;
    font-weight: 600 !important;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 800 !important;
    color: var(--accent) !important;
}

/* ===== BUTTONS ===== */
.stButton > button {
    background: #FFFFFF !important;
    border: 1.5px solid var(--border-bright) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--accent) !important;
    font-weight: 600 !important;
    font-size: 0.875rem !important;
    padding: 8px 18px !important;
    transition: all 0.15s !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
}
.stButton > button:hover {
    background: #EAF2FC !important;
    border-color: var(--accent) !important;
    box-shadow: var(--glow) !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"] {
    background: var(--accent) !important;
    color: #FFFFFF !important;
    border-color: var(--accent) !important;
    font-weight: 700 !important;
}
.stButton > button[kind="primary"]:hover {
    background: #155DA0 !important;
    border-color: #155DA0 !important;
}

/* ===== INPUTS ===== */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stNumberInput > div > div > input {
    background: #FFFFFF !important;
    border: 1.5px solid var(--border) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--text-primary) !important;
    font-size: 0.875rem !important;
    transition: border-color 0.15s !important;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px rgba(26,111,191,0.1) !important;
    outline: none !important;
}
.stTextInput label, .stTextArea label, .stSelectbox label,
.stNumberInput label, .stFileUploader label {
    color: var(--text-secondary) !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    margin-bottom: 4px !important;
}
.stSelectbox > div > div, .stMultiSelect > div > div {
    background: #FFFFFF !important;
    border: 1.5px solid var(--border) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--text-primary) !important;
}
.stSelectbox > div > div:hover { border-color: var(--accent) !important; }

/* ===== CHAT ===== */
[data-testid="stChatMessage"] {
    background: var(--bg-card) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 12px 16px !important;
    margin: 5px 0 !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important;
}
[data-testid="stChatInput"] > div {
    background: #FFFFFF !important;
    border: 1.5px solid var(--border-bright) !important;
    border-radius: var(--radius) !important;
    box-shadow: var(--shadow) !important;
}
[data-testid="stChatInput"] > div:focus-within {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px rgba(26,111,191,0.1) !important;
}
[data-testid="stChatInput"] textarea {
    background: transparent !important;
    color: var(--text-primary) !important;
    font-size: 0.9rem !important;
}
[data-testid="stChatInput"] button {
    background: var(--accent) !important;
    border-radius: var(--radius-sm) !important;
    color: #FFFFFF !important;
}

/* ===== TABLE ===== */
.stDataFrame, [data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    overflow: hidden !important;
    box-shadow: var(--shadow) !important;
}
.stDataFrame thead tr th {
    background: #F0F5FB !important;
    color: var(--accent) !important;
    font-size: 0.82rem !important;
    font-weight: 700 !important;
    border-bottom: 2px solid var(--border-bright) !important;
    padding: 10px 12px !important;
}
.stDataFrame tbody tr td {
    background: #FFFFFF !important;
    color: var(--text-primary) !important;
    font-size: 0.84rem !important;
    border-bottom: 1px solid #EEF2F7 !important;
    padding: 8px 12px !important;
}
.stDataFrame tbody tr:hover td { background: #F5F9FF !important; }

/* ===== TABS ===== */
.stTabs [data-baseweb="tab-list"] {
    background: #FFFFFF !important;
    border-bottom: 2px solid var(--border) !important;
    padding: 0 4px !important;
    gap: 2px !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--text-secondary) !important;
    font-weight: 600 !important;
    font-size: 0.875rem !important;
    padding: 10px 18px !important;
    border-bottom: 2px solid transparent !important;
    transition: all 0.15s !important;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--accent) !important; }
.stTabs [aria-selected="true"] {
    color: var(--accent) !important;
    border-bottom-color: var(--accent) !important;
    background: transparent !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: var(--bg-surface) !important;
    border: 1px solid var(--border) !important;
    border-top: none !important;
    border-radius: 0 0 var(--radius) var(--radius) !important;
    padding: 18px !important;
}

/* ===== ALERTS ===== */
[data-testid="stAlert"] { border-radius: var(--radius) !important; font-size: 0.875rem !important; }
.stSuccess { background: #F0FBF6 !important; border-color: var(--accent3) !important; color: #0D6E48 !important; }
.stWarning { background: #FFFBF0 !important; border-color: #E8A400 !important; color: #7A5500 !important; }
.stError   { background: #FFF5F5 !important; border-color: var(--danger) !important; color: #9B1C1C !important; }
.stInfo    { background: #F0F6FF !important; border-color: var(--accent) !important; color: #0D4F8C !important; }

/* ===== MISC ===== */
hr { border-color: var(--border) !important; margin: 20px 0 !important; }
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--bg-base); }
::-webkit-scrollbar-thumb { background: #C0CFDE; border-radius: 99px; }
::-webkit-scrollbar-thumb:hover { background: var(--border-bright); }
code, pre {
    background: #F0F5FB !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--accent) !important;
    font-family: var(--mono) !important;
    font-size: 0.82rem !important;
}
#MainMenu { visibility: hidden; }
.stDeployButton { display: none; }
footer { visibility: hidden; }
[data-testid="stFileUploader"] {
    background: #FAFCFF !important;
    border: 2px dashed var(--border) !important;
    border-radius: var(--radius) !important;
}
[data-testid="stFileUploader"]:hover { border-color: var(--accent) !important; }
.stProgress > div > div > div > div {
    background: linear-gradient(90deg, var(--accent), var(--accent2)) !important;
    border-radius: 99px !important;
}
.stMarkdown ul li, .stMarkdown ol li { color: var(--text-secondary) !important; font-size: 0.9rem !important; margin: 4px 0 !important; }
.stMarkdown strong { color: var(--text-primary) !important; font-weight: 700 !important; }
.stMarkdown a { color: var(--accent) !important; text-decoration: none !important; font-weight: 500 !important; }
.stMarkdown h2 { font-size: 1.1rem !important; font-weight: 700 !important; color: var(--accent) !important; }
.stMarkdown h3 { font-size: 0.95rem !important; font-weight: 600 !important; color: var(--text-secondary) !important; }
.stCaption, [data-testid="stCaptionContainer"] p { color: var(--text-muted) !important; font-size: 0.8rem !important; }
.stCheckbox > label, .stRadio > label { color: var(--text-secondary) !important; font-weight: 500 !important; }
.stNumberInput button { background: #F0F5FB !important; border-color: var(--border) !important; color: var(--accent) !important; }
</style>
""", unsafe_allow_html=True)

init_db()

# ── PERSISTENT MEMORY ──────────────────────────────────────────────
_MEMORY_FILE = "chat_memory.json"
_MAX_MEMORY_MSGS = 30

def _load_memory():
    import os
    if os.path.exists(_MEMORY_FILE):
        try:
            with open(_MEMORY_FILE, "r", encoding="utf-8") as _f:
                return json.load(_f).get("messages", [])
        except:
            pass
    return []

def _save_memory(msgs):
    try:
        with open(_MEMORY_FILE, "w", encoding="utf-8") as _f:
            json.dump({"messages": msgs[-_MAX_MEMORY_MSGS:]}, _f, ensure_ascii=False, indent=2)
    except:
        pass
# ───────────────────────────────────────────────────────────────────

# ==========================================
# CƠ CHẾ AUTO-RUN KHI KHỞI ĐỘNG HỆ THỐNG
# ==========================================
# Không tự động sync khi khởi động — người dùng bấm nút thủ công
# ==========================================

# ==========================================
# 2. THANH SIDEBAR
# ==========================================
with st.sidebar:
    st.title("🌊 TexFlow")
    st.caption("Quản lý Chuỗi Cung Ứng Dệt May")
    st.markdown("---")
    menu_selection = st.radio("", ["📊 Dashboard"], label_visibility="collapsed")
    # READ-ONLY banner
    st.markdown(
        "<div style='background:#EAF2FC;border-left:3px solid #1A6FBF;"
        "border-radius:6px;padding:8px 12px;font-size:0.8rem;color:#0D4F8C'>"
        "👁️ <b>Chế độ xem</b><br>Chỉ đọc dữ liệu</div>",
        unsafe_allow_html=True
    )
    # READ-ONLY mode banner
    st.markdown(
        "<div style='background:#EAF2FC;border-left:3px solid #1A6FBF;"
        "border-radius:6px;padding:8px 12px;font-size:0.8rem;color:#0D4F8C;margin-bottom:8px'>"
        "&#128065; <b>Chế độ chỉ đọc</b><br>Không có quyền chỉnh sửa dữ liệu.</div>",
        unsafe_allow_html=True
    )
    st.markdown("---")
    st.subheader("⚙️ Cài đặt Hệ thống AI")
    if "model_name" not in st.session_state:
        st.session_state.model_name = "qwen2.5:3b"
    model_name = st.text_input(
        "Tên Model đang chạy trên Ollama:",
        value=st.session_state.model_name,
        key="model_name_input",
    )
    if model_name != st.session_state.model_name:
        st.session_state.model_name = model_name

# ==========================================
# 3. TRANG DASHBOARD & BÁO CÁO TƯƠNG TÁC
# ==========================================
if menu_selection == "📊 Dashboard":
    st.title("TỔNG QUAN QUY TRÌNH DỆT MAY")
    
    graph = graphviz.Digraph(engine='dot')
    graph.attr(rankdir='TB', size='10,8', bgcolor='#1E1E1E') 
    
    node_attr = {
        'shape': 'box', 'style': 'filled,rounded', 'fontname': 'Helvetica', 
        'fontsize': '12', 'fontcolor': '#FFFFFF', 'color': '#4A90E2', 
        'penwidth': '1.5', 'height': '0.9', 'width': '2.2'
    }
    
    d_nodes = {
        "Kho Sợi Tổng": {"color": "#2C3E50", "label": "📦 [KHO SỢI TỔNG]\nNguồn nguyên liệu"},
        "Xưởng Nhuộm": {"color": "#2C3E50", "label": "🎨 [XƯỞNG NHUỘM]\nXử lý màu"},
        "Kho Sợi Sizing": {"color": "#2C3E50", "label": "🧵 [KHO SỢI SIZING]\nBáo cáo số lượng sợi"},
        "Máy Direct": {"color": "#34495E", "label": "⚙️ [MÁY DIRECT]\nSản xuất beam"},
        "Máy Hồ": {"color": "#34495E", "label": "🧪 [MÁY HỒ]\nHồ sợi"},
        "Máy Sectional": {"color": "#34495E", "label": "⚙️ [MÁY SECTIONAL]\nSản xuất beam"},
        "Kho Beam Sizing": {"color": "#2C3E50", "label": "🧻 [KHO BEAM SIZING]\nLưu trục Sizing"},
        "Kho Sợi Weaving": {"color": "#2C3E50", "label": "🧵 [KHO SỢI WEAVING]\nKho sợi dệt"},
        "Kho Beam Weaving": {"color": "#2C3E50", "label": "🧻 [KHO BEAM WEAVING]\nKho trục dệt"},
        "Xưởng Dệt": {"color": "#2C3E50", "label": "🏭 [XƯỞNG DỆT]\nSản xuất vải"},
        "Kho Thành Phẩm": {"color": "#2C3E50", "label": "🛍️ [KHO THÀNH PHẨM]\nNhập kho vải"}
    }
    for name, props in d_nodes.items():
        graph.node(name, props["label"], fillcolor=props["color"], **node_attr)

    edge_attr = {'color': '#00FFAA', 'penwidth': '1.5', 'arrowsize': '1.0'}
    
    with graph.subgraph() as s:
        s.attr(rank='same')
        s.node('Kho Sợi Tổng')
        s.node('Kho Sợi Sizing')
    with graph.subgraph() as s:
        s.attr(rank='same')
        s.node('Xưởng Nhuộm')
        s.node('Máy Direct')
        s.node('Máy Sectional')
    with graph.subgraph() as s:
        s.attr(rank='same')
        s.node('Máy Hồ')
        s.node('Kho Sợi Weaving')
    with graph.subgraph() as s:
        s.attr(rank='same')
        s.node('Kho Beam Sizing')
        s.node('Kho Beam Weaving')
    with graph.subgraph() as s:
        s.attr(rank='same')
        s.node('Xưởng Dệt')
        s.node('Kho Thành Phẩm')

    edges = [
        ('Kho Sợi Tổng', 'Xưởng Nhuộm'), ('Kho Sợi Tổng', 'Kho Sợi Sizing'), ('Kho Sợi Tổng', 'Kho Sợi Weaving'),
        ('Xưởng Nhuộm', 'Kho Sợi Weaving'), ('Xưởng Nhuộm', 'Kho Sợi Sizing'),
        ('Kho Sợi Sizing', 'Máy Direct'), ('Máy Direct', 'Máy Hồ'), ('Máy Hồ', 'Kho Beam Sizing'),
        ('Kho Sợi Sizing', 'Máy Sectional'), ('Máy Sectional', 'Kho Beam Sizing'),
        ('Kho Beam Sizing', 'Kho Beam Weaving'),
        ('Kho Sợi Weaving', 'Xưởng Dệt'), ('Kho Beam Weaving', 'Xưởng Dệt'), ('Xưởng Dệt', 'Kho Thành Phẩm')
    ]
    for src, dst in edges:
        graph.edge(src, dst, **edge_attr)

    col1, col2, col3 = st.columns([1, 8, 1])
    with col2:
        st.graphviz_chart(graph, use_container_width=True)

    st.divider()

    # --- PHẦN 1: BÁO CÁO CHI TIẾT ---
    st.subheader("🔍 PHÂN TÍCH CHI TIẾT TỪNG CỤM")
    
    selected_node = st.selectbox("👉 Chọn Cụm/Kho trên sơ đồ để xem báo cáo:", list(d_nodes.keys()))
    
    df_cluster = pd.DataFrame() 
    current_view_date = None
    
    with st.container(border=True):
        st.markdown(f"### 📈 Báo cáo: {selected_node}")
        
        try:
            df_cluster_raw = execute_query(f"SELECT * FROM Inventory_Log WHERE cluster_name = '{selected_node}'")
            
            if df_cluster_raw.empty:
                st.info(f"Chưa có dữ liệu nào được tải lên cho cụm [{selected_node}]. Hãy cấu hình Auto-Sync để nạp file.")
            else:
                df_cluster = get_clean_cluster_data(df_cluster_raw, selected_node)

                if selected_node == "Xưởng Dệt":
                    mac_col = find_best_col(df_cluster, ['ten_may', 'số máy', 'máy', 'mac'])
                    loc_col = find_best_col(df_cluster, ['sub_location', 'khu vực', 'xưởng', 'sheet', 'vị trí'])
                    y_col = find_best_col(df_cluster, ['quantity_yard', 'yard', 'mét', 'met'])
                    kg_col = find_best_col(df_cluster, ['quantity_kg', 'tổng', 'total', 'kg', 'khối lượng'])
                    order_col = find_best_col(df_cluster, ['order_id', 'tên hàng', 'sản phẩm', 'mã', 'item_id'])

                    if not mac_col:
                        st.error("🚨 HỆ THỐNG KHÔNG TÌM THẤY CỘT SỐ MÁY. Sếp hãy kiểm tra lại cấu hình Mapping, hãy đảm bảo cột chứa Số Máy được map hoặc có tên chứa chữ 'máy'.")
                    else:
                        if 'date' not in df_cluster.columns: df_cluster['date'] = pd.Timestamp.now().strftime('%Y-%m-%d')
                        # CHỐNG LỖI KHI TÌM UNIQUE TRÊN KIỂU DỮ LIỆU TRỘN LẪN
                        available_dates = sorted([str(d) for d in df_cluster['date'].replace("", pd.NA).dropna().unique()], reverse=True)
                        
                        if not available_dates: 
                            st.info("Chưa có ngày sản xuất nào được ghi nhận hợp lệ.")
                        else:
                            selected_date = st.selectbox("Chọn ngày xem báo cáo Xưởng Dệt:", available_dates)
                            current_view_date = selected_date
                            df_day = df_cluster[df_cluster['date'] == selected_date].copy()

                            if y_col and y_col in df_day.columns: 
                                df_day[y_col] = pd.to_numeric(df_day[y_col], errors='coerce').fillna(0)
                            if kg_col and kg_col in df_day.columns: 
                                df_day[kg_col] = pd.to_numeric(df_day[kg_col], errors='coerce').fillna(0)
                            
                            total_yard = df_day[y_col].sum() if y_col and y_col in df_day.columns and not df_day.empty else 0
                            total_kg = df_day[kg_col].sum() if kg_col and kg_col in df_day.columns and not df_day.empty else 0
                            
                            active_machines_df = df_day[~df_day[order_col].astype(str).str.lower().isin(["nan", "none", "", "0", "chưa xác định", "pd.na"])] if order_col and order_col in df_day.columns and not df_day.empty else pd.DataFrame()
                            active_machines = active_machines_df[mac_col].nunique() if not active_machines_df.empty else 0
                            
                            col1_1, col1_2, col1_3 = st.columns(3)
                            with col1_1: st.metric("Tổng Sản Lượng (Yard)", f"{total_yard:,.0f} Y")
                            with col1_2: st.metric("Tổng Sản Lượng (Kg)", f"{total_kg:,.1f} Kg")
                            with col1_3: st.metric("Số Máy Đang Chạy", f"{active_machines} máy")

                            st.markdown("---")
                            st.markdown("### 🏭 SƠ ĐỒ MẶT BẰNG XƯỞNG DỆT (DIGITAL TWIN)")
                            
                            def assign_zone(loc_str, mac_str):
                                loc_str = str(loc_str).lower().strip()
                                mac_str = str(mac_str).lower().strip()
                                if loc_str in ['1', '1.0', 'wea 1', 'wea1', 'weaving 1', 'w1'] or 'w1' in mac_str or 'wea 1' in loc_str: return 1
                                if loc_str in ['2', '2.0', 'wea 2', 'wea2', 'weaving 2', 'w2'] or 'w2' in mac_str or 'wea 2' in loc_str: return 2
                                if loc_str in ['3', '3.0', 'wea 3', 'wea3', 'weaving 3', 'w3'] or 'w3' in mac_str or 'wea 3' in loc_str: return 3
                                try:
                                    val = int(float(loc_str))
                                    if val in [1, 2, 3]: return val
                                except: pass
                                return 0

                            df_day['zone_id'] = df_day.apply(lambda row: assign_zone(row.get(loc_col, ''), row.get(mac_col, '')), axis=1)

                            unknown_mask = df_day['zone_id'] == 0
                            if not df_day.empty and unknown_mask.any():
                                z_counts = df_day[~unknown_mask].groupby('zone_id')[mac_col].nunique().to_dict()
                                for idx in df_day[unknown_mask].index:
                                    if z_counts.get(1, 0) < 48: 
                                        df_day.at[idx, 'zone_id'] = 1
                                        z_counts[1] = z_counts.get(1, 0) + 1
                                    elif z_counts.get(2, 0) < 56: 
                                        df_day.at[idx, 'zone_id'] = 2
                                        z_counts[2] = z_counts.get(2, 0) + 1
                                    else: 
                                        df_day.at[idx, 'zone_id'] = 3
                                        z_counts[3] = z_counts.get(3, 0) + 1

                            tab_w1, tab_w2, tab_w3 = st.tabs(["🏭 Weaving 1 (48 Máy)", "🏭 Weaving 2 (56 Máy)", "🏭 Weaving 3 (24 Máy)"])
                            
                            def render_zone_dashboard(zone_id, capacity, title, color_hex):
                                df_zone = df_day[df_day['zone_id'] == zone_id].copy() if 'zone_id' in df_day.columns and not df_day.empty else pd.DataFrame()
                                
                                total_y_zone = df_zone[y_col].sum() if y_col and y_col in df_zone.columns and not df_zone.empty else 0
                                total_kg_zone = df_zone[kg_col].sum() if kg_col and kg_col in df_zone.columns and not df_zone.empty else 0
                                
                                active_macs_df_zone = df_zone[~df_zone[order_col].astype(str).str.lower().isin(["nan", "none", "", "0", "chưa xác định", "pd.na"])] if order_col and order_col in df_zone.columns and not df_zone.empty else pd.DataFrame()
                                active_macs_zone = active_macs_df_zone[mac_col].nunique() if mac_col and mac_col in active_macs_df_zone.columns and not active_macs_df_zone.empty else 0
                                
                                c1, c2, c3 = st.columns(3)
                                with c1: st.metric(f"Sản Lượng {title} (Yard)", f"{total_y_zone:,.0f} Y")
                                with c2: st.metric(f"Sản Lượng {title} (Kg)", f"{total_kg_zone:,.1f} Kg")
                                with c3: st.metric("Số Máy Đang Chạy", f"{active_macs_zone} máy")
                                
                                st.markdown("---")
                                
                                zone_mac_list = []
                                z_grouped = pd.DataFrame()
                                
                                if not df_zone.empty:
                                    agg_funcs = {}
                                    if y_col and y_col in df_zone.columns: agg_funcs[y_col] = 'sum'
                                    if kg_col and kg_col in df_zone.columns: agg_funcs[kg_col] = 'sum'
                                    if order_col and order_col in df_zone.columns: agg_funcs[order_col] = 'first'
                                    
                                    group_keys = [mac_col, loc_col] if loc_col and loc_col in df_zone.columns else [mac_col]
                                    
                                    if agg_funcs: z_grouped = df_zone.groupby(group_keys).agg(agg_funcs).reset_index()
                                    else: z_grouped = df_zone[group_keys].drop_duplicates()
                                    
                                    # CHỐNG LỖI "min() arg is an empty sequence" KHI GROUPBY RỖNG
                                    if not z_grouped.empty:
                                        z_grouped['sort_key'] = pd.to_numeric(z_grouped[mac_col].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(999)
                                        z_grouped = z_grouped.sort_values('sort_key')
                                    
                                        for _, row in z_grouped.iterrows():
                                            raw_mac_name = str(row[mac_col]).strip()
                                            if raw_mac_name.endswith('.0'): raw_mac_name = raw_mac_name[:-2]
                                            if raw_mac_name.isdigit(): raw_mac_name = f"Máy {raw_mac_name}"
                                            
                                            zone_mac_list.append({
                                                "mac_name": raw_mac_name,
                                                "yard": row.get(y_col, 0) if y_col else 0,
                                                "kg": row.get(kg_col, 0) if kg_col else 0,
                                                "order": str(row.get(order_col, "N/A")).strip() if order_col else "N/A"
                                            })
                                        
                                running_macs = [m for m in zone_mac_list if str(m['order']).lower() not in ["nan", "none", "", "0", "chưa xác định", "pd.na"]]
                                used_slots = len(running_macs)
                                
                                html = f"<h5 style='margin-top: 5px; margin-bottom: 15px;'>Sơ đồ {title} (Đang chạy: {used_slots} / Tổng máy: {capacity})</h5>"
                                html += "<div style='display: grid; grid-template-columns: repeat(auto-fill, minmax(65px, 1fr)); gap: 6px; margin-bottom: 25px;'>"
                                for i in range(capacity):
                                    if i < len(zone_mac_list):
                                        mac = zone_mac_list[i]
                                        m_order, m_yard, m_kg, m_name = mac['order'], mac['yard'], mac['kg'], mac['mac_name']
                                        if str(m_order).lower() in ["nan", "none", "", "0", "chưa xác định", "pd.na"]:
                                            html += (
                                                f"<div title='{m_name} đang dừng' style='"
                                                f"background-color: #E2E8F0; color: #718096; "
                                                f"padding: 6px 3px; border-radius: 6px; text-align: center; "
                                                f"font-size: 10px; border: 1px dashed #A0AEC0; "
                                                f"min-height: 70px; display:flex; flex-direction:column; "
                                                f"align-items:center; justify-content:center;'>"
                                                f"<b style='font-size:11px; color:#4A5568'>{m_name}</b>"
                                                f"<span style='font-size:9px'>Dừng</span>"
                                                f"</div>"
                                            )
                                        else:
                                            # Rút gọn tên hàng cho vừa ô
                                            _order_short = str(m_order)[:18] + "…" if len(str(m_order)) > 18 else str(m_order)
                                            _kg_str = f"{m_kg:,.0f} Kg" if m_kg else ""
                                            _tooltip = f"Mã hàng: {m_order}\nSản lượng: {m_kg:,.1f} Kg | Yard: {m_yard:,.0f}"
                                            html += (
                                                f"<div title='{_tooltip}' style='"
                                                f"background-color: {color_hex}; color: white; "
                                                f"padding: 6px 3px; border-radius: 6px; text-align: center; "
                                                f"font-size: 10px; cursor: pointer; "
                                                f"border: 1px solid rgba(255,255,255,0.4); "
                                                f"box-shadow: 0 2px 4px rgba(0,0,0,0.2); "
                                                f"line-height: 1.3; min-height: 70px; "
                                                f"display:flex; flex-direction:column; align-items:center; justify-content:center;'>"
                                                f"<b style='font-size:11px'>{m_name}</b>"
                                                f"<span style='font-size:9px; opacity:0.9; word-break:break-word; margin-top:2px'>{_order_short}</span>"
                                                f"<span style='font-size:10px; font-weight:bold; margin-top:2px'>{_kg_str}</span>"
                                                f"</div>"
                                            )
                                    else:
                                        html += (
                                                "<div title='Không có dữ liệu máy' style='"
                                                "background-color: #F7FAFC; color: #CBD5E0; "
                                                "padding: 6px 3px; border-radius: 6px; text-align: center; "
                                                "font-size: 9px; border: 1px dashed #E2E8F0; "
                                                "min-height: 70px; display:flex; align-items:center; justify-content:center;'>"
                                                "Trống</div>"
                                            )
                                html += "</div>"
                                st.markdown(html, unsafe_allow_html=True)
                                
                                with st.expander(f"Xem bảng số liệu chi tiết của riêng {title}"):
                                    # CHỐNG LỖI st.dataframe RỖNG
                                    if not df_zone.empty:
                                        st.dataframe(df_zone.drop(columns=['id', 'cluster_name', 'type', 'zone_id'], errors='ignore'), use_container_width=True)
                                    else:
                                        st.info(f"Không có chi tiết nào cho khu vực {title}.")

                            with tab_w1: render_zone_dashboard(1, 48, "Weaving 1", "#D69E2E")
                            with tab_w2: render_zone_dashboard(2, 56, "Weaving 2", "#3182CE")
                            with tab_w3: render_zone_dashboard(3, 24, "Weaving 3", "#00B517")

                elif "Kho" in selected_node:
                    qty_col = find_best_col(df_cluster, ['quantity', 'số lượng', 'mét', 'met', 'kg', 'total', 'tổng'])
                    item_col = find_best_col(df_cluster, ['item_id', 'tên hàng', 'mã', 'sản phẩm', 'order_id'])
                    loc_col = find_best_col(df_cluster, ['sub_location', 'khu vực', 'vị trí', 'kho'])

                    if not qty_col or not item_col:
                        st.error("🚨 HỆ THỐNG KHÔNG TÌM THẤY CỘT 'MÃ HÀNG' HOẶC 'SỐ LƯỢNG'. Hãy kiểm tra lại Mapping.")
                    else:
                        df_cluster[qty_col] = pd.to_numeric(df_cluster[qty_col], errors='coerce').fillna(0)
                        
                        col1_1, col1_2, col1_3 = st.columns(3)
                        if not loc_col:
                            loc_col = 'sub_location_temp'
                            df_cluster[loc_col] = "Khu vực dưới đất"
                        else:
                            df_cluster[loc_col] = df_cluster[loc_col].replace("", "Khu vực dưới đất").fillna("Khu vực dưới đất")
                        
                        valid_items = df_cluster[df_cluster[item_col].astype(str) != "Chưa xác định"]
                        df_nhap = valid_items[valid_items['type'].isin(['NHAP', 'TON_DAU'])].groupby([item_col, loc_col])[qty_col].sum()
                        df_xuat = valid_items[valid_items['type'] == 'XUAT'].groupby([item_col, loc_col])[qty_col].sum()
                        df_stock = df_nhap.sub(df_xuat, fill_value=0).reset_index(name='quantity')
                        df_stock = df_stock[df_stock['quantity'] > 0] 
                        
                        ton_kho_tong = df_stock['quantity'].sum()
                        so_luong_ma_sp = df_stock[item_col].nunique()
                        
                        with col1_1: st.metric("Tổng Tồn Kho Hiện Tại", f"{ton_kho_tong:,.0f}")
                        with col1_2: st.metric("Số Mã Sản Phẩm Đang Lưu", f"{so_luong_ma_sp}")
                        with col1_3: st.metric("Cảnh báo Sắp Hết Hạn", "0", delta="- Bình thường", delta_color="normal")

                        if selected_node == "Kho Beam Weaving":
                            st.markdown("---")
                            st.markdown("### 🗄️ SƠ ĐỒ LƯU TRỮ BEAM THỰC TẾ (SỐ HÓA)")
                            
                            def extract_location(loc_str):
                                loc_str = str(loc_str).strip()
                                match = re.search(r'(60|95)\D+(\d+)', loc_str)
                                if match: return int(match.group(1)), int(match.group(2))
                                return None, None

                            rack_60_slots, rack_95_slots, floor_beams = [None] * 60, [None] * 95, []
                            for _, row in df_stock.iterrows():
                                beam_data = row.to_dict()
                                rack_type, slot_idx = extract_location(beam_data[loc_col])
                                if rack_type == 60 and 1 <= slot_idx <= 60: rack_60_slots[slot_idx - 1] = beam_data
                                elif rack_type == 95 and 1 <= slot_idx <= 95: rack_95_slots[slot_idx - 1] = beam_data
                                else: floor_beams.append(beam_data) 

                            def draw_fixed_rack(title, capacity, slot_array, color_hex):
                                used_slots = sum(1 for x in slot_array if x is not None)
                                html = f"<h5 style='margin-top: 5px; margin-bottom: 15px;'>{title} (Đang dùng: {used_slots} / Sức chứa: {capacity})</h5>"
                                html += "<div style='display: grid; grid-template-columns: repeat(auto-fill, minmax(60px, 1fr)); gap: 6px; margin-bottom: 25px;'>"
                                for i in range(capacity):
                                    beam = slot_array[i]
                                    if beam is not None:
                                        short_id = str(beam[item_col])[-5:] if len(str(beam[item_col])) > 5 else str(beam[item_col])
                                        b_id, b_qty = beam[item_col], beam['quantity']
                                        html += f"<div title='Mã Beam: {b_id} | Tọa độ: Ô số {i+1} | SL: {b_qty:,.0f}' style='background-color: {color_hex}; color: white; padding: 12px 2px; border-radius: 4px; text-align: center; font-size: 11px; cursor: pointer; border: 1px solid rgba(255,255,255,0.4); font-weight: bold;'>{short_id}</div>"
                                    else:
                                        html += f"<div title='Vị trí Ô {i+1} trống' style='background-color: #1A202C; color: #4A5568; padding: 12px 2px; border-radius: 4px; text-align: center; font-size: 10px; border: 1px dashed #4A5568;'>Trống {i+1}</div>"
                                html += "</div>"
                                return html
                            
                            def draw_floor(beams_list):
                                html = f"<h5 style='margin-top: 5px; margin-bottom: 15px;'>Khu Vực Dưới Đất (Đang để: {len(beams_list)} cuộn)</h5>"
                                if len(beams_list) == 0: return html + "<p style='color:#718096; font-style:italic;'>Kho trống.</p>"
                                html += "<div style='display: grid; grid-template-columns: repeat(auto-fill, minmax(60px, 1fr)); gap: 6px; margin-bottom: 25px;'>"
                                for beam in beams_list:
                                    short_id = str(beam[item_col])[-5:] if len(str(beam[item_col])) > 5 else str(beam[item_col])
                                    b_id, b_loc, b_qty = beam[item_col], beam[loc_col], beam['quantity']
                                    html += f"<div title='Mã: {b_id} | Vị trí: {b_loc} | SL: {b_qty:,.0f}' style='background-color: #D69E2E; color: white; padding: 12px 2px; border-radius: 4px; text-align: center; font-size: 11px; cursor: pointer; border: 1px solid rgba(255,255,255,0.4); font-weight: bold;'>{short_id}</div>"
                                html += "</div>"
                                return html

                            tab_all, tab_60, tab_95, tab_floor = st.tabs(["👁️ Tất cả Kho", "🟢 Giá 60 chỗ", "🔵 Giá 95 chỗ", "🟡 Khu dưới đất"])
                            with tab_all:
                                st.markdown(draw_fixed_rack("1. Giá Treo Beam (Khu Vực 60 Chỗ)", 60, rack_60_slots, "#00B517"), unsafe_allow_html=True)
                                st.markdown(draw_fixed_rack("2. Giá Treo Beam (Khu Vực 95 Chỗ)", 95, rack_95_slots, "#3182CE"), unsafe_allow_html=True)
                                st.markdown(draw_floor(floor_beams), unsafe_allow_html=True)
                            with tab_60: st.markdown(draw_fixed_rack("Giá Treo Beam 60 Chỗ", 60, rack_60_slots, "#00B517"), unsafe_allow_html=True)
                            with tab_95: st.markdown(draw_fixed_rack("Giá Treo Beam 95 Chỗ", 95, rack_95_slots, "#3182CE"), unsafe_allow_html=True)
                            with tab_floor: st.markdown(draw_floor(floor_beams), unsafe_allow_html=True)
                        else:
                            st.markdown("**📊 Biểu đồ Tồn theo Sản Phẩm (Top 10)**")
                            df_stock_total = df_stock.groupby(item_col)['quantity'].sum().reset_index()
                            # CHỐNG LỖI BIỂU ĐỒ RỖNG
                            if not df_stock_total.empty: st.bar_chart(df_stock_total.sort_values(by='quantity', ascending=False).head(10).set_index(item_col)['quantity'])

                with st.expander("Xem bảng số liệu chi tiết (Gốc của Cụm này)"):
                    df_clean = df_cluster.copy()
                    
                    try:
                        db_files = glob.glob("*.db")
                        if db_files and "Kho" in selected_node:
                            with sqlite3.connect(db_files[0]) as conn:
                                df_yarn = pd.read_sql_query("SELECT * FROM Yarn_Dictionary", conn)
                                coeff_dict = {str(k).strip().upper(): float(v) for k, v in zip(df_yarn['yarn_type'], df_yarn['coefficient'])} if not df_yarn.empty else {}
                                
                                qty_col = find_best_col(df_clean, ['quantity', 'mét', 'met', 'số lượng'])
                                yarn_col = find_best_col(df_clean, ['sợi', 'yarn'])
                                beam_size_col = find_best_col(df_clean, ['tổng', 'size'])
                                
                                if qty_col and yarn_col and beam_size_col and coeff_dict:
                                    kg_list = []
                                    for _, row in df_clean.iterrows():
                                        kg = ""
                                        try:
                                            yarn_val = str(row[yarn_col]).upper()
                                            core_match = re.search(r'(\d+[S]/\d+|\d+NTW)', yarn_val)
                                            core_yarn = core_match.group(1) if core_match else yarn_val
                                            coeff = coeff_dict.get(core_yarn)
                                            if not coeff: 
                                                for k, v in coeff_dict.items():
                                                    if k in core_yarn or core_yarn in k: coeff = v; break
                                            if coeff:
                                                m_str = str(row[qty_col]).replace(',', '').strip()
                                                ts_str = str(row[beam_size_col]).replace(',', '').strip()
                                                if m_str and ts_str and m_str != "nan" and ts_str != "nan":
                                                    kg = math.ceil(((float(m_str) / 0.9144) * float(ts_str) * coeff) + 10) / 1000
                                        except: pass
                                        kg_list.append(kg)
                                    df_clean['Kg_TonKho'] = kg_list
                    except Exception: pass
                    
                    # CHỐNG LỖI BẢNG GỐC RỖNG
                    if not df_clean.empty:
                        st.dataframe(df_clean.sort_values('date', ascending=False) if 'date' in df_clean.columns else df_clean, use_container_width=True)
                    else:
                        st.info("Bảng dữ liệu gốc trống.")

        except Exception as e:
            st.error(f"Lỗi tải dữ liệu chi tiết: {e}")

    st.markdown("<br>", unsafe_allow_html=True)

    # --- PHẦN 2: TRỢ LÝ AI ---
    st.subheader("🤖 TRỢ LÝ AI (Động cơ Lật Trang Tự Động)")
    if not model_name:
        st.warning("⚠️ Vui lòng nhập Tên Model Ollama ở thanh bên trái.")
    else:
        with st.expander("📝 1. Không Gian Ảo (Nhập Luật / Công thức cho AI)"):
            st.info("💡 Bạn gõ công thức vào đây. AI sẽ học thuộc vĩnh viễn.")
            col_rule1, col_rule2 = st.columns([8, 2])
            with col_rule1:
                default_rule = """Quy định chung:
1. Khi người dùng hỏi khối lượng (Kg), hệ thống Python bên dưới đã tính sẵn và trả kết quả ở cột 'Kg_TonKho'.
2. AI tuyệt đối KHÔNG cần tự làm toán. Chỉ cần ĐỌC bảng kết quả và trả lời người dùng ngắn gọn, súc tích."""
                new_rule = st.text_area("Quy tắc:", value=default_rule, height=100)
            
            with col_rule2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Lưu Luật", type="primary", use_container_width=True):
                    if new_rule:
                        run_db_command("INSERT INTO AI_Rules (rule_text) VALUES (?)", (new_rule,))
                        st.success("✅ Đã lưu!")
                        st.rerun()
            
            try:
                db_files = glob.glob("*.db")
                if db_files:
                    with sqlite3.connect(db_files[0]) as conn:
                        df_rules = pd.read_sql_query("SELECT * FROM AI_Rules", conn)
                        if not df_rules.empty:
                            st.markdown("**📂 Các tài liệu/luật AI đang ghi nhớ:**")
                            for idx, row in df_rules.iterrows():
                                c_text, c_btn = st.columns([9, 1])
                                with c_text: st.info(row['rule_text'])
                                with c_btn:
                                    if st.button("Xóa", key=f"del_rule_{row['id']}"):
                                        run_db_command(f"DELETE FROM AI_Rules WHERE id = {row['id']}")
                                        st.rerun()
            except: pass

        with st.expander("📚 2. Quản lý Từ Điển Sợi (Xem, Sửa & Xóa hệ số)"):
            st.info("💡 Python sẽ dùng hệ số này để tự tính toán phía sau hậu trường.")
            try:
                db_files = glob.glob("*.db")
                if db_files:
                    with sqlite3.connect(db_files[0]) as conn:
                        df_yarn_view = pd.read_sql_query("SELECT * FROM Yarn_Dictionary", conn)
                    if not df_yarn_view.empty:
                        edited_yarn = st.data_editor(df_yarn_view, num_rows="dynamic", use_container_width=True, key="yarn_editor")
                        if st.button("💾 Lưu thay đổi Từ Điển", type="primary"):
                            run_db_command("DELETE FROM Yarn_Dictionary")
                            for _, r in edited_yarn.iterrows():
                                if pd.notna(r['yarn_type']) and str(r['yarn_type']).strip() != "":
                                    run_db_command("INSERT OR REPLACE INTO Yarn_Dictionary (yarn_type, coefficient) VALUES (?, ?)", 
                                                (str(r['yarn_type']).strip(), float(r['coefficient']) if pd.notna(r['coefficient']) else 0.0))
                            st.rerun()
                    else: st.warning("Từ điển đang trống.")
            except: pass

        st.markdown("<br>", unsafe_allow_html=True)

        col_h1, col_h2 = st.columns([9,1])
        with col_h2:
            if st.button("🗑️ Xóa lịch sử", key="clr_mem", use_container_width=True):
                st.session_state.messages = []
                _save_memory([])
                st.rerun()
        chat_container = st.container(height=400)
        if "messages" not in st.session_state: st.session_state.messages = []
        
        with chat_container:
            for msg in st.session_state.messages:
                with st.chat_message(msg["role"]): st.markdown(msg["content"])
                    
        with st.form(key="ai_chat_form", clear_on_submit=True):
            col_input, col_btn = st.columns([9, 1])
            with col_input: user_prompt = st.text_input("Hỏi AI", label_visibility="collapsed", placeholder="Nhập câu hỏi tại đây...")
            with col_btn: submit_ai = st.form_submit_button("Gửi 🚀")
                
        if submit_ai and user_prompt:
            st.session_state.messages.append({"role": "user", "content": user_prompt})
            with chat_container:
                st.chat_message("user").markdown(user_prompt)
                with st.chat_message("assistant"):
                    message_placeholder = st.empty() 
                    try:
                        # CHẠY ĐỘNG CƠ LẬT TRANG (YIELD DUMMY CHUNK + TEXT CHUNK)
                        response_stream = process_ai_chat(user_prompt, st.session_state.messages, selected_node, model_name, df_cluster, current_view_date)
                        
                        final_text = ""
                        for chunk in response_stream:
                            try:
                                if chunk.text:
                                    final_text += chunk.text
                                    message_placeholder.markdown(final_text + "▌") 
                            except: continue
                        message_placeholder.markdown(final_text)

                    except Exception as e:
                        message_placeholder.markdown(f"Lỗi gọi AI: {str(e)}")

            st.session_state.messages.append({"role": "assistant", "content": final_text})

# ==========================================
# 4. TRANG 2: QUẢN LÝ DỮ LIỆU (CẤU HÌNH TỰ ĐỘNG)
# ==========================================