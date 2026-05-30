"""
sizing_pipeline.py — Đọc file FILE_HIỆU_SUẤT_TỔNG.xlsx
Các sheet: MÁY HỒ, MÁY SEC, MÁY QS, SUZUKI, WINDER
Lưu vào bảng Sizing_Log (SQLite)
"""
import os, glob, sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

DB_NAME = "inventory.db"

# ── Column mapping cho từng sheet ──────────────────────────────────
# (col_index, db_field_name)
SHEET_COLS = {
    "MÁY HỒ": {
        "machine_type": "MÁY HỒ",
        "cols": {
            0: "nam", 1: "thang", 2: "ngay", 3: "ten_may", 4: "nguoi_chay",
            5: "tong_soi", 6: "pg", 7: "loai_soi", 8: "pd_yd",
            9: "toc_do_muc_tieu", 10: "toc_do_thuc_te", 11: "ca",
            12: "gio_bat_dau", 13: "gio_ket_thuc", 14: "thoi_gian_phut",
            15: "sl_muc_tieu_mtr", 16: "sl_thuc_te_mtr", 17: "sl_kg",
            18: "hieu_suat_pct", 19: "hs_toc_do_pct",
            20: "dut_soi", 21: "keo_soi", 22: "t_keo_soi",
            23: "chia_luoc", 24: "thay_beam", 25: "dao_tao",
            26: "tap_hop", 27: "an", 28: "ve_sinh",
            29: "cho_qs", 30: "sua_chua", 31: "cup_dien", 32: "khac",
            33: "gd_hoat_dong", 34: "sl_100_khong_dung", 35: "hs_khong_dung",
        }
    },
    "MÁY SEC": {
        "machine_type": "MÁY SEC",
        "cols": {
            0: "nam", 1: "thang", 2: "ngay", 3: "ten_may", 4: "nguoi_chay",
            5: "tong_soi", 6: "pg", 7: "loai_soi", 8: "pd_yd",
            9: "so_section", 10: "chieu_dai_beam",
            11: "toc_do_muc_tieu", 12: "toc_do_thuc_te", 13: "ca",
            14: "gio_bat_dau", 15: "gio_ket_thuc", 16: "thoi_gian_phut",
            17: "sl_muc_tieu_mtr", 18: "sl_thuc_te_mtr", 19: "hieu_suat_pct",
            20: "dut_soi", 21: "cho_bo_soi", 22: "noi_soi",
            23: "beaming_ks", 24: "xo_luoc", 25: "kiem_tra",
            26: "chiet_soi", 27: "sua_chua", 28: "tap_hop",
            29: "an", 30: "ve_sinh", 31: "khac",
            32: "gd_hoat_dong", 33: "sl_100_khong_dung", 34: "hs_khong_dung",
            # Beaming part
            41: "sl_thuc_te_beaming_mtr", 42: "sl_kg_2mtr", 43: "sl_kg_thanh_pham",
            44: "so_beam_ra", 45: "hs_beaming_pct",
            58: "sl_tong_mtr", 60: "hs_tong_pct",
        }
    },
    "MÁY QS": {
        "machine_type": "MÁY QS",
        "cols": {
            0: "nam", 1: "thang", 2: "ngay", 3: "ten_may", 4: "nguoi_chay",
            5: "loai_soi", 6: "tong_soi", 7: "pd_yd",
            8: "toc_do_muc_tieu", 9: "toc_do_thuc_te", 10: "ca",
            11: "gio_bat_dau", 12: "gio_ket_thuc", 13: "thoi_gian_phut",
            14: "sl_muc_tieu_mtr", 15: "sl_thuc_te_mtr", 16: "sl_kg",
            17: "hieu_suat_pct", 18: "hs_toc_do_pct",
            19: "dut_soi", 20: "lot_no", 21: "ncc",
            22: "keo_soi", 23: "thay_beam", 24: "dao_tao",
            25: "tap_hop", 26: "an", 27: "ve_sinh",
            28: "khong_cn", 29: "cho_soi", 30: "cho_bqs",
            31: "sua_chua", 32: "cup_dien", 33: "khac",
            34: "gd_hoat_dong", 35: "sl_100_khong_dung", 36: "hs_khong_dung",
        }
    },
    "SUZUKI": {
        "machine_type": "SUZUKI",
        "cols": {
            0: "nam", 1: "thang", 2: "ngay", 3: "ca", 4: "nguoi_chay",
            5: "order_sample", 6: "tong_soi", 7: "pg",
            8: "loai_soi", 9: "pd_yd", 10: "chieu_dai_beam", 11: "sl_kg",
        }
    },
    "WINDER": {
        "machine_type": "WINDER",
        "cols": {
            0: "nam", 1: "thang", 2: "ngay", 3: "ten_may", 4: "ca",
            5: "nguoi_chay", 6: "phan_loai", 7: "loai_soi",
            8: "so_cone", 9: "cone_mtr", 10: "kg_cone", 11: "tong_kl_kg",
            12: "work_mtr", 13: "sl_kg", 14: "sl_mtr",
            15: "sl_100_mtr", 16: "hieu_suat_pct",
            17: "cho_lay_soi", 18: "kiem_tra", 19: "ve_sinh",
            20: "dao_tao", 21: "an", 22: "khac",
            23: "t_dung_may", 24: "sl_100_khong_dung", 25: "hs_khong_dung",
        }
    },
}


def init_sizing_table():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS Sizing_Log (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name        TEXT,
            machine_type     TEXT,   -- MÁY HỒ / MÁY SEC / MÁY QS / SUZUKI / WINDER
            date             TEXT,   -- YYYY-MM-DD
            nam              INTEGER,
            thang            INTEGER,
            ngay             INTEGER,
            ten_may          TEXT,
            ca               TEXT,
            nguoi_chay       TEXT,
            tong_soi         REAL,
            pg               TEXT,   -- P/G
            loai_soi         TEXT,
            pd_yd            TEXT,
            so_section       REAL,
            chieu_dai_beam   REAL,
            order_sample     TEXT,
            toc_do_muc_tieu  REAL,
            toc_do_thuc_te   REAL,
            thoi_gian_phut   REAL,
            gio_bat_dau      TEXT,   -- HH:MM
            gio_ket_thuc     TEXT,
            sl_muc_tieu_mtr  REAL,
            sl_thuc_te_mtr   REAL,
            sl_kg            REAL,
            hieu_suat_pct    REAL,
            hs_toc_do_pct    REAL,
            hs_khong_dung    REAL,
            gd_hoat_dong     REAL,
            -- Downtime reasons
            dut_soi          REAL,
            keo_soi          REAL,
            t_keo_soi        REAL,
            cho_qs           REAL,
            sua_chua         REAL,
            dao_tao          REAL,
            tap_hop          REAL,
            an               REAL,
            ve_sinh          REAL,
            cup_dien         REAL,
            khac             REAL,
            -- Sectional extra
            sl_thuc_te_beaming_mtr REAL,
            sl_kg_thanh_pham       REAL,
            so_beam_ra             REAL,
            hs_beaming_pct         REAL,
            sl_tong_mtr            REAL,
            hs_tong_pct            REAL,
            -- Winder extra
            so_cone          REAL,
            cone_mtr         REAL,
            kg_cone          REAL,
            phan_loai        TEXT,
            -- Other
            lot_no           TEXT,
            ncc              TEXT,
            UNIQUE(file_name, machine_type, date, ten_may, ca, nguoi_chay, loai_soi)
        )
    """)
    conn.commit()
    conn.close()


def _safe(val, as_type=str):
    """Clean value."""
    if val is None or (isinstance(val, float) and val != val):
        return None
    v = str(val).strip()
    if v in ("", "nan", "NaT", "None"):
        return None
    if as_type == float:
        try:
            return float(v.replace(",", "").replace("%", ""))
        except:
            return None
    if as_type == int:
        try:
            return int(float(v))
        except:
            return None
    return v


def _parse_time(val):
    """Convert time object or string to HH:MM."""
    if val is None:
        return None
    import datetime as dt
    if isinstance(val, dt.time):
        return val.strftime("%H:%M")
    s = str(val).strip()
    if ":" in s:
        parts = s.split(":")[:2]
        return ":".join(p.zfill(2) for p in parts)
    return None


def parse_sheet(filepath: str, sheet_name: str, fname: str) -> list[dict]:
    """Parse one sheet into list of records."""
    cfg = SHEET_COLS.get(sheet_name)
    if not cfg:
        return []

    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
    col_map = cfg["cols"]
    mtype   = cfg["machine_type"]

    records = []
    for _, row in df.iterrows():
        # Filter: col0 must be a year 2020-2030
        yr = _safe(row.iloc[0], int)
        if not yr or not (2020 <= yr <= 2030):
            continue

        r = {"file_name": fname, "machine_type": mtype}
        for col_idx, field in col_map.items():
            if col_idx >= len(row):
                r[field] = None
                continue
            val = row.iloc[col_idx]
            if field in ("gio_bat_dau", "gio_ket_thuc"):
                r[field] = _parse_time(val)
            elif field in ("nam", "thang", "ngay"):
                r[field] = _safe(val, int)
            elif field in ("ten_may", "ca", "nguoi_chay", "loai_soi", "pg",
                           "pd_yd", "order_sample", "lot_no", "ncc", "phan_loai"):
                r[field] = _safe(val)
            else:
                r[field] = _safe(val, float)

        # Build date
        try:
            r["date"] = f"{r['nam']:04d}-{r['thang']:02d}-{r['ngay']:02d}"
        except:
            r["date"] = None

        # ten_may fallback for SUZUKI
        if "ten_may" not in r or not r.get("ten_may"):
            r["ten_may"] = mtype

        records.append(r)
    return records


def upsert_sizing(records: list[dict]) -> int:
    if not records:
        return 0
    conn = sqlite3.connect(DB_NAME)
    n = 0
    FIELDS = [
        "file_name","machine_type","date","nam","thang","ngay",
        "ten_may","ca","nguoi_chay","tong_soi","pg","loai_soi","pd_yd",
        "so_section","chieu_dai_beam","order_sample",
        "toc_do_muc_tieu","toc_do_thuc_te","thoi_gian_phut",
        "gio_bat_dau","gio_ket_thuc",
        "sl_muc_tieu_mtr","sl_thuc_te_mtr","sl_kg",
        "hieu_suat_pct","hs_toc_do_pct","hs_khong_dung","gd_hoat_dong",
        "dut_soi","keo_soi","t_keo_soi","cho_qs","sua_chua",
        "dao_tao","tap_hop","an","ve_sinh","cup_dien","khac",
        "sl_thuc_te_beaming_mtr","sl_kg_thanh_pham","so_beam_ra",
        "hs_beaming_pct","sl_tong_mtr","hs_tong_pct",
        "so_cone","cone_mtr","kg_cone","phan_loai","lot_no","ncc",
    ]
    placeholders = ",".join("?" * len(FIELDS))
    sql = f"INSERT OR REPLACE INTO Sizing_Log ({','.join(FIELDS)}) VALUES ({placeholders})"
    for r in records:
        try:
            vals = [r.get(f) for f in FIELDS]
            conn.execute(sql, vals)
            n += 1
        except Exception as e:
            pass
    conn.commit()
    conn.close()
    return n


def import_sizing_file(filepath: str, sheets: list = None) -> dict:
    """Import all 5 machine sheets from one file."""
    init_sizing_table()
    fname  = os.path.basename(filepath)
    target = sheets or ["MÁY HỒ", "MÁY SEC", "MÁY QS", "SUZUKI", "WINDER"]
    result = {}
    errors = []
    for sh in target:
        try:
            recs = parse_sheet(filepath, sh, fname)
            n    = upsert_sizing(recs)
            result[sh] = n
        except Exception as e:
            result[sh] = 0
            errors.append(f"{sh}: {e}")
    result["total"]  = sum(result.get(s, 0) for s in target)
    result["errors"] = errors
    return result


def list_files_in_folder(folder: str, keyword: str = "") -> list:
    """Liệt kê tất cả file xlsx/xls trong thư mục (đệ quy). Dùng rglob nên hỗ trợ lồng nhau."""
    root = Path(folder)
    files = [str(p) for p in root.rglob("*.xlsx")] + [str(p) for p in root.rglob("*.xls")]
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if keyword:
        files = [f for f in files if keyword.lower() in f.lower()]
    return sorted(files)


def scan_sizing_folder(folder: str, keyword: str = "",
                       progress_cb=None, sheets: list = None) -> dict:
    """
    Quét thư mục đệ quy (năm → tháng → file), đọc tất cả file khớp.
    progress_cb(i, total, filename): callback hiển thị tiến độ (optional).
    """
    init_sizing_table()
    files = list_files_in_folder(folder, keyword)
    total_rows = 0
    total_files = 0
    errors = []
    for i, fp in enumerate(files):
        if progress_cb:
            try:
                progress_cb(i, len(files), os.path.basename(fp))
            except:
                pass
        try:
            r = import_sizing_file(fp, sheets)
            if r["total"] > 0:
                total_rows  += r["total"]
                total_files += 1
            if r.get("errors"):
                errors.extend(r["errors"])
        except Exception as e:
            errors.append(f"{os.path.basename(fp)}: {e}")

    return {"total_rows": total_rows, "total_files": total_files,
            "total_scanned": len(files), "errors": errors}


# ── Parser cho file Preparation_Product_Report (format cũ) ────────
# Sheet: Daily Data
# Columns: Year, Month, Day, Process GR, Machine, Shift, Set no,
#          Total Ends, Yarn Type, Section, Product Length,
#          Product Length Excludes beam, Product Weight Excludes beam,
#          Product Weight, PD&YD, Start time, finish time,
#          Yarn Broken, Rewinder, Lot.no, NCC

# Mapping Process GR → machine_type
PROCESS_GR_MAP = {
    "Sizing":         "MÁY HỒ",
    "Sectional Warp": "MÁY SEC",
    "Direct Warp":    "MÁY QS",
    "Order":          "MÁY QS",    # Direct order
    "Sample Warp":    "MÁY QS",
    "Winder":         "WINDER",
}


def parse_prep_report(filepath: str, fname: str = None) -> list[dict]:
    """
    Đọc sheet 'Daily Data' từ Preparation_Product_Report.
    Map sang cùng cấu trúc Sizing_Log, bỏ trống các cột không có.
    """
    fname = fname or os.path.basename(filepath)
    try:
        df = pd.read_excel(filepath, sheet_name="Daily Data", header=0)
    except Exception:
        return []

    records = []
    for _, row in df.iterrows():
        yr = _safe(row.get("Year"), int)
        if not yr or not (2020 <= yr <= 2030):
            continue

        process_gr = _safe(row.get("Process GR")) or ""
        mtype = PROCESS_GR_MAP.get(process_gr, process_gr or "MÁY HỒ")

        # Parse time
        st_time = _parse_time(row.get("Start time"))
        fin_time = _parse_time(row.get("finish time"))

        # Duration in minutes from start/finish
        thoi_gian = None
        if st_time and fin_time:
            try:
                from datetime import datetime as _dt
                _s = _dt.strptime(st_time, "%H:%M")
                _e = _dt.strptime(fin_time, "%H:%M")
                diff = (_e - _s).seconds // 60
                if diff < 0: diff += 1440  # cross midnight
                thoi_gian = float(diff)
            except:
                pass

        records.append({
            "file_name":       fname,
            "machine_type":    mtype,
            "date":            f"{yr:04d}-{_safe(row.get('Month'),int) or 1:02d}-{_safe(row.get('Day'),int) or 1:02d}",
            "nam":             yr,
            "thang":           _safe(row.get("Month"), int),
            "ngay":            _safe(row.get("Day"), int),
            "ten_may":         _safe(row.get("Machine")) or mtype,
            "ca":              _safe(row.get("Shift")),
            "nguoi_chay":      None,
            "tong_soi":        _safe(row.get("Total Ends"), float),
            "pg":              None,
            "loai_soi":        _safe(row.get("Yarn Type")),
            "pd_yd":           _safe(row.get("PD&YD")),
            "so_section":      _safe(row.get("Section"), float),
            "chieu_dai_beam":  None,
            "order_sample":    _safe(row.get("Set no")),
            "toc_do_muc_tieu": None,
            "toc_do_thuc_te":  None,
            "thoi_gian_phut":  thoi_gian,
            "gio_bat_dau":     st_time,
            "gio_ket_thuc":    fin_time,
            "sl_muc_tieu_mtr": None,
            "sl_thuc_te_mtr":  _safe(row.get("Product Length\nExcludes beam 2mtr"), float) or
                               _safe(row.get("Product Length"), float),
            "sl_kg":           _safe(row.get("Product Weight\nExcludes beam 2mtr"), float) or
                               _safe(row.get("Product Weight"), float),
            "hieu_suat_pct":   None,
            "hs_toc_do_pct":   None,
            "hs_khong_dung":   None,
            "gd_hoat_dong":    None,
            "dut_soi":         _safe(row.get("Yarn Broken"), float),
            "keo_soi":         None, "t_keo_soi": None, "cho_qs": None,
            "sua_chua":        None, "dao_tao": None, "tap_hop": None,
            "an":              None, "ve_sinh": None, "cup_dien": None, "khac": None,
            "sl_thuc_te_beaming_mtr": None, "sl_kg_thanh_pham": None,
            "so_beam_ra":      None, "hs_beaming_pct": None,
            "sl_tong_mtr":     None, "hs_tong_pct": None,
            "so_cone":         None, "cone_mtr": None, "kg_cone": None,
            "phan_loai":       None,
            "lot_no":          _safe(row.get("Lot.no")),
            "ncc":             _safe(row.get("NCC")),
        })
    return records


def import_prep_report(filepath: str) -> dict:
    """Import file Preparation_Product_Report vào Sizing_Log."""
    init_sizing_table()
    fname = os.path.basename(filepath)
    recs  = parse_prep_report(filepath, fname)
    n     = upsert_sizing(recs)
    return {"total": n, "rows_read": len(recs), "errors": []}


def detect_and_import(filepath: str) -> dict:
    """
    Tự động detect loại file và import đúng cách:
    - FILE_HIỆU_SUẤT_TỔNG.xlsx → import_sizing_file (5 sheets riêng)
    - Preparation_Product_Report*.xlsx → import_prep_report (Daily Data)
    """
    xl = pd.ExcelFile(filepath)
    sheets = xl.sheet_names
    if "Daily Data" in sheets:
        r = import_prep_report(filepath)
        r["file_type"] = "prep_report"
        return r
    elif any(s in sheets for s in ["MÁY HỒ","MÁY SEC","MÁY QS","SUZUKI","WINDER"]):
        r = import_sizing_file(filepath)
        r["file_type"] = "hieu_suat_tong"
        return r
    else:
        return {"total": 0, "errors": [f"Không nhận ra định dạng file: {os.path.basename(filepath)}"], "file_type": "unknown"}


def scan_sizing_folder_v2(folder: str, keyword: str = "", progress_cb=None) -> dict:
    """
    Quét thư mục đệ quy, tự detect loại file và import.
    Hỗ trợ cả 2 định dạng: FILE_HIỆU_SUẤT_TỔNG và Preparation_Product_Report.
    """
    init_sizing_table()
    files = list_files_in_folder(folder, keyword)
    total_rows = 0; total_files = 0
    prep_count = 0; hieu_suat_count = 0
    errors = []
    for i, fp in enumerate(files):
        if progress_cb:
            try: progress_cb(i, len(files), os.path.basename(fp))
            except: pass
        try:
            r = detect_and_import(fp)
            if r["total"] > 0:
                total_rows  += r["total"]
                total_files += 1
                if r.get("file_type") == "prep_report":
                    prep_count += r["total"]
                else:
                    hieu_suat_count += r["total"]
        except Exception as e:
            errors.append(f"{os.path.basename(fp)}: {e}")
    return {
        "total_rows": total_rows, "total_files": total_files,
        "total_scanned": len(files),
        "prep_report": prep_count, "hieu_suat": hieu_suat_count,
        "errors": errors,
    }