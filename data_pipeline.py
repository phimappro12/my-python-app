import pandas as pd
import json
import os

CONFIG_FILE = "saved_mappings.json"

def load_saved_mappings():
    """Đọc các mẫu mapping đã lưu từ hệ thống."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_new_mapping(template_name, mapping_dict):
    """Lưu mẫu mapping mới của người dùng vào file JSON."""
    mappings = load_saved_mappings()
    mappings[template_name] = mapping_dict
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(mappings, f, ensure_ascii=False, indent=4)

def process_data(file_obj, mapping_dict, cluster_name, transaction_type, manual_date=None, sheet_name=0, skip_rows=0):
    """Đọc và chuẩn hóa dữ liệu dựa trên mapping linh hoạt."""
    
    # Đọc file
    if file_obj.name.endswith('.csv'):
        df = pd.read_csv(file_obj, skiprows=skip_rows)
    else:
        df = pd.read_excel(file_obj, sheet_name=sheet_name, skiprows=skip_rows)

    # Đổi tên cột
    df_mapped = df.rename(columns=mapping_dict)
    target_cols = list(mapping_dict.values())
    
    # Xử lý ngày tháng
    if 'date' not in target_cols and manual_date:
        df_mapped['date'] = manual_date
        target_cols.append('date')
    elif 'date' not in target_cols and not manual_date:
        raise ValueError("Bạn phải map cột thành 'date' hoặc chọn ngày thủ công!")

    # Lọc cột
    df_final = df_mapped[target_cols].copy()

    # Chuẩn hóa dữ liệu cốt lõi
    df_final['cluster_name'] = cluster_name
    df_final['type'] = transaction_type
    
    if 'unit' not in df_final.columns:
        df_final['unit'] = 'mét'

    if 'item_id' in df_final.columns:
        df_final = df_final.dropna(subset=['item_id'])
        df_final['item_id'] = df_final['item_id'].astype(str).str.replace(r'\.0$', '', regex=True)

    if 'quantity' in df_final.columns:
        df_final['quantity'] = pd.to_numeric(df_final['quantity'], errors='coerce').fillna(0)

    if 'date' in df_final.columns:
        df_final['date'] = pd.to_datetime(df_final['date']).dt.strftime('%Y-%m-%d')

    return df_final