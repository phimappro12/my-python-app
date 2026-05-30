import sqlite3
import pandas as pd

DB_NAME = "inventory.db"

def init_db():
    """Khởi tạo database và tạo bảng nếu chưa có."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Inventory_Log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            cluster_name TEXT,
            item_id TEXT,
            type TEXT,
            quantity REAL,
            unit TEXT
        )
    ''')
    conn.commit()
    conn.close()

def insert_data(df):
    """Ghi DataFrame vào SQLite và TỰ ĐỘNG THÊM CỘT nếu file báo cáo có cột mới."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        # Lấy danh sách các cột đang có trong bảng
        cursor.execute("PRAGMA table_info(Inventory_Log)")
        existing_columns = [row[1] for row in cursor.fetchall()]
        
        # Quét xem dataframe có cột nào lạ không
        for col in df.columns:
            if col not in existing_columns:
                print(f"Phát hiện cột mới: {col}. Đang mở rộng Database...")
                cursor.execute(f"ALTER TABLE Inventory_Log ADD COLUMN '{col}' TEXT")
                
        # Ghi dữ liệu
        df.to_sql('Inventory_Log', conn, if_exists='append', index=False)
        conn.commit()
    except Exception as e:
        raise e
    finally:
        conn.close()

def execute_query(query):
    """Thực thi câu SQL thuần (Dành cho AI lấy dữ liệu)."""
    conn = sqlite3.connect(DB_NAME)
    try:
        if any(keyword in query.upper() for keyword in ['DROP', 'DELETE', 'UPDATE', 'INSERT']):
             raise ValueError("Hệ thống chỉ cho phép AI thực thi câu lệnh SELECT.")
        
        result_df = pd.read_sql_query(query, conn)
        return result_df
    except Exception as e:
        raise e
    finally:
        conn.close()

def execute_update(query, params=()):
    """Thực thi câu SQL cập nhật (CREATE/INSERT/UPDATE/DELETE)."""
    conn = sqlite3.connect(DB_NAME)
    try:
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()
        return True
    except Exception as e:
        print(f"Lỗi DB: {e}")
        return False
    finally:
        conn.close()