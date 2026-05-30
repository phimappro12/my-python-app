"""
weaving_pipeline.py — Đọc file TOTAL_weaving_I_II_III.xlsx
3 sheets: wea 1, wea 2, wea 3 → Inventory_Log
Dùng positional access để tránh vấn đề tên cột.
"""
import os, re, sqlite3
import pandas as pd
from pathlib import Path

DB_NAME = "inventory.db"
DEFAULT_YEAR = "2026"

# Sheet → sub_location mapping
SHEET_LOC = {"wea 1": "Weaving 1", "wea 2": "Weaving 2", "wea 3": "Weaving 3"}


def _s(val):
    if val is None: return None
    v = str(val).strip()
    return v if v not in ("", "nan", "None", "NaT") else None


def _f(val):
    try: return float(str(val).replace(",","").strip())
    except: return None


def _date(yr, m, d):
    try: return f"{int(yr):04d}-{int(float(m)):02d}-{int(float(d)):02d}"
    except: return None


def init_weaving_table():
    conn = sqlite3.connect(DB_NAME)
    for col, typ in [
        ("color","TEXT"),("p_beam_yarn","TEXT"),("rpm","REAL"),
        ("kl_ca_a","REAL"),("kl_ca_b","REAL"),("kl_ca_c","REAL"),
        ("tieu_chuan_kg","REAL"),("hs_ca_a","REAL"),("hs_ca_b","REAL"),
        ("hs_ca_c","REAL"),("hieu_suat_2ca","REAL"),("hieu_suat_3ca","REAL"),
        ("one_pcs","REAL"),("kind_of_dyeing","TEXT"),("gsm","REAL"),
        ("dty","TEXT"),("filament","TEXT"),("phan_loai","TEXT"),("ghi_chu","TEXT"),
        ("file_name","TEXT"),
    ]:
        try: conn.execute(f"ALTER TABLE Inventory_Log ADD COLUMN {col} {typ}")
        except: pass

    # ✅ FIX Bug2: Tạo UNIQUE index để chặn import trùng lặp về sau
    # (date + sub_location + ten_may) là khoá tự nhiên của 1 dòng sản xuất
    try:
        conn.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS idx_weaving_no_dup
            ON Inventory_Log(date, sub_location, ten_may)
            WHERE date IS NOT NULL AND sub_location IS NOT NULL AND ten_may IS NOT NULL
        """)
    except Exception:
        pass  # index có thể đã tồn tại

    # ✅ FIX Bug2: Xoá sạch duplicate cũ đã lỡ import (giữ row có id lớn nhất)
    try:
        conn.execute("""
            DELETE FROM Inventory_Log
            WHERE id NOT IN (
                SELECT MAX(id)
                FROM Inventory_Log
                WHERE date IS NOT NULL AND sub_location IS NOT NULL AND ten_may IS NOT NULL
                GROUP BY date, sub_location, ten_may
            )
            AND date IS NOT NULL AND sub_location IS NOT NULL AND ten_may IS NOT NULL
        """)
    except Exception:
        pass

    conn.commit(); conn.close()


def parse_weaving_sheet(filepath, sheet_name, year=DEFAULT_YEAR, fname=None):
    fname = fname or os.path.basename(filepath)
    sub_loc = SHEET_LOC.get(sheet_name)
    if not sub_loc: return []

    df = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
    cols = list(df.columns)

    # Positional: Month=0, Date=1, Máy=2, Tên hàng=3, Color=4, Beam yarn=5, RPM=6
    # wea 1: KL_A=7, unnamed=8(ignore), KL_B=9, std_kg=10, total=11, HS_A=12, unnamed=13, HS_B=14, eff_2ca=15, ghi_chu=16, pcs=17, dyeing=18, phan_loai=19, dty=20, fil=21
    # wea 2/3: KL_A=7, KL_B=8, KL_C=9, std=10, total=11, HS_A=12, HS_B=13, HS_C=14, eff_3ca=15, ghi_chu=16, pcs=17, dyeing=18, gsm/phan_loai=19, dty=20, fil=21

    n_cols = len(cols)
    is_2ca = "2 ca" in str(cols[15]) if n_cols > 15 else False

    records = []
    for _, row in df.iterrows():
        vals = list(row)
        month = _s(vals[0]) if len(vals) > 0 else None
        day   = _s(vals[1]) if len(vals) > 1 else None
        mac   = _s(vals[2]) if len(vals) > 2 else None

        if not month or not day or not mac: continue
        try: int(float(month)); int(float(day)); float(mac)
        except: continue

        date_str = _date(year, month, day)
        if not date_str: continue

        if is_2ca:  # wea 1 (2-shift)
            r = {
                "date": date_str, "sub_location": sub_loc,
                "cluster_name": "Xưởng Dệt", "file_name": fname,
                "ten_may":      mac,
                "item_id":      _s(vals[3])  if n_cols>3  else None,
                "color":        _s(vals[4])  if n_cols>4  else None,
                "p_beam_yarn":  _s(vals[5])  if n_cols>5  else None,
                "rpm":          _f(vals[6])  if n_cols>6  else None,
                "kl_ca_a":      _f(vals[7])  if n_cols>7  else None,
                "kl_ca_b":      _f(vals[9])  if n_cols>9  else None,
                "tieu_chuan_kg":_f(vals[10]) if n_cols>10 else None,
                "quantity_kg":  _f(vals[11]) if n_cols>11 else None,
                "hs_ca_a":      _f(vals[12]) if n_cols>12 else None,
                "hs_ca_b":      _f(vals[14]) if n_cols>14 else None,
                "hieu_suat_2ca":_f(vals[15]) if n_cols>15 else None,
                "ghi_chu":      _s(vals[16]) if n_cols>16 else None,
                "one_pcs":      _f(vals[17]) if n_cols>17 else None,
                "kind_of_dyeing":_s(vals[18])if n_cols>18 else None,
                "phan_loai":    _s(vals[19]) if n_cols>19 else None,
                "dty":          _s(vals[20]) if n_cols>20 else None,
                "filament":     _s(vals[21]) if n_cols>21 else None,
            }
            # Map hiệu suất tổng cho AI agent
            # (hieu_suat stored above)
        else:  # wea 2, wea 3 (3-shift)
            r = {
                "date": date_str, "sub_location": sub_loc,
                "cluster_name": "Xưởng Dệt", "file_name": fname,
                "ten_may":      mac,
                "item_id":      _s(vals[3])  if n_cols>3  else None,
                "color":        _s(vals[4])  if n_cols>4  else None,
                "p_beam_yarn":  _s(vals[5])  if n_cols>5  else None,
                "rpm":          _f(vals[6])  if n_cols>6  else None,
                "kl_ca_a":      _f(vals[7])  if n_cols>7  else None,
                "kl_ca_b":      _f(vals[8])  if n_cols>8  else None,
                "kl_ca_c":      _f(vals[9])  if n_cols>9  else None,
                "tieu_chuan_kg":_f(vals[10]) if n_cols>10 else None,
                "quantity_kg":  _f(vals[11]) if n_cols>11 else None,
                "hs_ca_a":      _f(vals[12]) if n_cols>12 else None,
                "hs_ca_b":      _f(vals[13]) if n_cols>13 else None,
                "hs_ca_c":      _f(vals[14]) if n_cols>14 else None,
                "hieu_suat_3ca":_f(vals[15]) if n_cols>15 else None,
                "ghi_chu":      _s(vals[16]) if n_cols>16 else None,
                "one_pcs":      _f(vals[17]) if n_cols>17 else None,
                "kind_of_dyeing":_s(vals[18])if n_cols>18 else None,
                "gsm":          _f(vals[19]) if n_cols>19 else None,
                "phan_loai":    _s(vals[19]) if n_cols>19 else None,  # wea3
                "dty":          _s(vals[20]) if n_cols>20 else None,
                "filament":     _s(vals[21]) if n_cols>21 else None,
            }
            # (hieu_suat stored above)

        records.append(r)
    return records


def upsert_weaving(records):
    if not records: return 0
    conn = sqlite3.connect(DB_NAME)
    n = 0
    all_fields = list(records[0].keys())
    cols = ",".join(f'[{f}]' for f in all_fields)
    ph   = ",".join("?" * len(all_fields))
    # ✅ FIX Bug2: INSERT OR IGNORE — bỏ qua nếu (date, sub_location, ten_may) đã tồn tại
    sql  = f"INSERT OR IGNORE INTO Inventory_Log ({cols}) VALUES ({ph})"
    for r in records:
        try:
            conn.execute(sql, [r.get(f) for f in all_fields])
            n += 1
        except: pass
    conn.commit(); conn.close()
    return n


def import_weaving_total(filepath, sheets=None, year=DEFAULT_YEAR):
    init_weaving_table()
    fname  = os.path.basename(filepath)
    target = sheets or ["wea 1", "wea 2", "wea 3"]
    result = {}; errors = []
    for sh in target:
        try:
            recs = parse_weaving_sheet(filepath, sh, year, fname)
            n = upsert_weaving(recs)
            result[sh] = n
        except Exception as e:
            result[sh] = 0; errors.append(f"{sh}: {e}")
    result["total"]  = sum(v for k,v in result.items() if k not in ("total","errors"))
    result["errors"] = errors
    return result


def scan_weaving_folder(folder, year=DEFAULT_YEAR, keyword="TOTAL", progress_cb=None):
    files = [str(p) for p in Path(folder).rglob("*.xlsx")]
    files += [str(p) for p in Path(folder).rglob("*.xls")]
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if keyword:
        files = [f for f in files if keyword.lower() in f.lower()]
    total_rows = 0; total_files = 0; errors = []
    for i, fp in enumerate(files):
        if progress_cb:
            try: progress_cb(i, len(files), os.path.basename(fp))
            except: pass
        try:
            yr_m = re.search(r'20(\d{2})', fp)
            yr   = f"20{yr_m.group(1)}" if yr_m else year
            r = import_weaving_total(fp, year=yr)
            if r["total"] > 0: total_rows += r["total"]; total_files += 1
            errors.extend(r.get("errors",[]))
        except Exception as e:
            errors.append(f"{os.path.basename(fp)}: {e}")
    return {"total_rows": total_rows, "total_files": total_files,
            "total_scanned": len(files), "errors": errors}