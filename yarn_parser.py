"""
yarn_parser.py — Đọc file Excel công thức sợi (Bang mau sx moi)
Trích xuất: Tên hàng, Sợi bông, Sợi ngang, Sợi nền, Sợi Border + tỷ lệ %
"""
import os
import glob
import sqlite3
import pandas as pd

DB_NAME = "inventory.db"

# Row/col constants (0-indexed)
ROW_PRODUCT = 5   # "Tên hàng" → col 5
ROW_YARNS   = 9   # "Các loại sợi" header row
ROW_PCT     = 10  # "Tỉ lệ các loại sợi" row
ROW_MACHINE = 12  # "Loại máy" → col 6

# Column positions of yarn type names and their percentages
YARN_COLS = {
    "soi_bong":   8,
    "soi_ngang":  15,
    "soi_nen":    22,
    "soi_border": 29,
}

def init_yarn_table():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS Yarn_Formula (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name    TEXT,
            sheet_name   TEXT,
            item_name    TEXT,
            ten_may      TEXT,
            soi_bong     TEXT,
            soi_bong_pct REAL,
            soi_ngang    TEXT,
            soi_ngang_pct REAL,
            soi_nen      TEXT,
            soi_nen_pct  REAL,
            soi_border   TEXT,
            soi_border_pct REAL,
            UNIQUE(file_name, sheet_name)
        )
    """)
    # Add ten_may column if table existed before this update
    try:
        conn.execute("ALTER TABLE Yarn_Formula ADD COLUMN ten_may TEXT")
    except:
        pass
    conn.commit()
    conn.close()

def _safe_float(val):
    try:
        f = float(val)
        return round(f * 100, 4) if f <= 1 else round(f, 4)  # convert 0.38 → 38.xx%
    except:
        return None

def parse_one_sheet(df: pd.DataFrame, file_name: str, sheet_name: str) -> dict | None:
    """Trích xuất 1 sheet từ DataFrame (header=None)."""
    try:
        item_name = str(df.iloc[ROW_PRODUCT, 5]).strip()
        if not item_name or item_name in ("nan", "None", ""):
            return None

        row_yarns = df.iloc[ROW_YARNS]
        row_pct   = df.iloc[ROW_PCT]

        # Loại máy: đọc trực tiếp từ file (row 12, col 6)
        # VD: SULZER, VAMATEX, TOYOTA 102, DORNIER...
        ten_may_val = ""
        try:
            raw_may = str(df.iloc[ROW_MACHINE, 6]).strip()
            if raw_may not in ("nan","None",""):
                ten_may_val = raw_may
        except:
            pass

        result = {
            "file_name":  os.path.basename(file_name),
            "sheet_name": sheet_name,
            "item_name":  item_name,
            "ten_may":    ten_may_val,
        }
        for key, col in YARN_COLS.items():
            yarn_type = str(row_yarns.iloc[col]).strip() if col < len(row_yarns) else ""
            yarn_pct  = _safe_float(row_pct.iloc[col])   if col < len(row_pct)   else None
            result[key]          = yarn_type if yarn_type not in ("nan", "None", "") else None
            result[key + "_pct"] = yarn_pct
        return result
    except Exception as e:
        return None

def parse_excel_file(filepath: str) -> list[dict]:
    """Đọc tất cả sheets từ 1 file Excel."""
    results = []
    try:
        xl = pd.ExcelFile(filepath)
        for sheet in xl.sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet, header=None)
            rec = parse_one_sheet(df, filepath, sheet)
            if rec:
                results.append(rec)
    except Exception as e:
        pass
    return results

def upsert_records(records: list[dict]) -> int:
    """Ghi vào DB, bỏ qua nếu đã tồn tại (file+sheet)."""
    if not records:
        return 0
    conn = sqlite3.connect(DB_NAME)
    inserted = 0
    for r in records:
        try:
            conn.execute("""
                INSERT OR REPLACE INTO Yarn_Formula
                (file_name, sheet_name, item_name, ten_may,
                 soi_bong, soi_bong_pct, soi_ngang, soi_ngang_pct,
                 soi_nen, soi_nen_pct, soi_border, soi_border_pct)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                r["file_name"], r["sheet_name"], r["item_name"],
                r.get("ten_may"),
                r.get("soi_bong"), r.get("soi_bong_pct"),
                r.get("soi_ngang"), r.get("soi_ngang_pct"),
                r.get("soi_nen"), r.get("soi_nen_pct"),
                r.get("soi_border"), r.get("soi_border_pct"),
            ))
            inserted += 1
        except:
            pass
    conn.commit()
    conn.close()
    return inserted

def scan_folder(folder: str, keyword: str = "") -> tuple[int, int, list[str]]:
    """
    Quét thư mục, đọc tất cả .xlsx/.xls khớp keyword.
    Returns: (total_records, total_files, errors)
    """
    init_yarn_table()

    # Dùng pathlib.Path.rglob() thay vì glob.glob()
    # vì glob bị bug với đường dẫn có dấu [ ] (VD: [P.M])
    from pathlib import Path
    root = Path(folder)
    all_files = [str(p) for p in root.rglob("*.xlsx")]
    all_files += [str(p) for p in root.rglob("*.xls")]
    # Bỏ file tạm của Excel (bắt đầu bằng ~$)
    all_files = [f for f in all_files if not os.path.basename(f).startswith("~$")]

    if keyword:
        all_files = [f for f in all_files if keyword.lower() in f.lower()]

    total_records = 0
    total_files   = 0
    errors        = []
    for fp in all_files:
        recs = parse_excel_file(fp)
        if recs:
            n = upsert_records(recs)
            total_records += n
            total_files   += 1
        else:
            errors.append(os.path.basename(fp))

    return total_records, total_files, errors