"""
beam_info.py — Đọc file thông tin Beam (BEAM_DAT) và lưu vào SQLite
====================================================================
Sheet chính:
  - XUATKHO: danh sách beam đã xuất kho lên máy (có số mét, số kg thực tế)
  - YCCB   : danh sách beam đã yêu cầu xuất (kế hoạch)
  - 19-5   : phiếu yêu cầu beam theo ngày

Bảng DB:
  - Beam_Info (XUATKHO): thông tin beam đang/đã trên máy — nguồn chính để tính còn lại
  - Beam_Request (YCCB) : yêu cầu xuất beam
"""

import os
import sqlite3
import pandas as pd
from pathlib import Path

DB_NAME = "inventory.db"


# ── Mapping cột Excel → tên cột DB ────────────────────────────────
XUATKHO_COLS = {
    "MÃ BEAM":          "ma_beam",
    "1\n2":             "weaving",       # 1=W1, 2=W2, 3=W3
    "Số \nmáy":         "so_may",
    "LOẠI SỢI":         "loai_may",      # TOY, SUL, DOR, VAMA...
    "MÃ SỢI":           "ma_soi",
    "TÊN HÀNG":         "ten_hang",
    "TỔNG SỢI":         "tong_soi",
    "LOẠI SỢI.1":       "loai_soi",      # CD 30S/2 (10)...
    "P":                "beam_p",         # P beam
    "G":                "beam_g",         # G beam
    "SỐ\nMÉT":          "so_met",         # số mét ban đầu
    "PD/\nYD":          "phan_loai",
    "NGÀY":             "ngay_len_may",
    "GIỜ":              "gio_len_may",
    "SỐ GIÀN":          "so_gian",
    "SỐ MÉT THỰC TẾ":   "so_met_thuc_te",
    "SỐ KG\nTHỰC TẾ":  "so_kg_thuc_te",
    "SỐ CHỈ THỊ BEAM":  "chi_thi_beam",
    "SỐ CM THỤT VÀO":   "cm_thut_vao",
    "LOẠI MÂM BEAM":    "loai_mam",
    "NGƯỜI XUẤT":       "nguoi_xuat",
    "GHI CHÚ":          "ghi_chu",
}

YCCB_COLS = {
    "MÃ YÊU CẦU\n빔번호":     "ma_beam",
    "WEA\n공장구분":            "weaving",
    "MÁY\n호수":               "so_may",
    "LOẠI MÁY\n직기":          "loai_may",
    "MÃ SỢI \n본수":           "ma_soi",
    "TÊN HÀNG\n제품명":        "ten_hang",
    "TỔNG SỐ SỢI\n총본수":    "tong_soi",
    "LOẠI SỢI\n사종":          "loai_soi",
    "P":                       "beam_p",
    "G":                       "beam_g",
    "SỐ MÉT\n빔m":             "so_met",
    "PD/YD":                   "phan_loai",
    "NGÀY GIAO\n출고요청일":   "ngay_giao",
    "GIỜ GIAO\n출고요청시간":  "gio_giao",
    "GHI CHÚ\n빔출고시간":     "ghi_chu",
}


def init_beam_tables():
    """Tạo bảng Beam_Info và Beam_Request nếu chưa có."""
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS Beam_Info (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name       TEXT,
            ma_beam         TEXT,
            po_code         TEXT,
            ngay_yc         TEXT,
            ngay_giao       TEXT,
            weaving         TEXT,
            so_may          TEXT,
            loai_may        TEXT,
            ma_soi          TEXT,
            ten_hang        TEXT,
            tong_soi        REAL,
            loai_soi        TEXT,
            beam_p          TEXT,
            beam_g          TEXT,
            so_met          REAL,
            phan_loai       TEXT,
            ngay_len_may    TEXT,
            gio_len_may     TEXT,
            so_gian         TEXT,
            so_met_thuc_te  REAL,
            so_kg_thuc_te   REAL,
            chi_thi_beam    TEXT,
            cm_thut_vao     TEXT,
            loai_mam        TEXT,
            nguoi_xuat      TEXT,
            ghi_chu         TEXT,
            UNIQUE(ma_beam, weaving, so_may)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS Beam_Request (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name   TEXT,
            ma_beam     TEXT,
            weaving     TEXT,
            so_may      TEXT,
            loai_may    TEXT,
            ma_soi      TEXT,
            ten_hang    TEXT,
            tong_soi    REAL,
            loai_soi    TEXT,
            beam_p      TEXT,
            beam_g      TEXT,
            so_met      REAL,
            phan_loai   TEXT,
            ngay_giao   TEXT,
            gio_giao    TEXT,
            ghi_chu     TEXT,
            UNIQUE(ma_beam)
        )
    """)
    # Thêm cột mới nếu bảng đã tồn tại từ phiên bản cũ
    _new_cols = {
        "po_code":   "TEXT",
        "ngay_yc":   "TEXT",
        "ngay_giao": "TEXT",
    }
    for _col, _typ in _new_cols.items():
        try:
            conn.execute(f"ALTER TABLE Beam_Info ADD COLUMN {_col} {_typ}")
        except:
            pass  # Cột đã tồn tại → bỏ qua

    conn.commit()
    conn.close()


def _norm(val) -> str:
    """Chuẩn hóa giá trị về string sạch."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "NaT", "None"):
        return None
    v = str(val).strip()
    # Bỏ tab/newline trong ô Excel
    v = v.replace("\t", "").replace("\r", "")
    return v if v else None


def _norm_float(val) -> float:
    """Chuyển về float, trả về None nếu không hợp lệ."""
    try:
        f = float(str(val).replace(",", "").strip())
        return f if not pd.isna(f) else None
    except:
        return None


def _parse_weaving(val) -> str:
    """Chuyển 1/2/3 → 'Weaving 1/2/3'."""
    v = _norm(val)
    if v and v.isdigit():
        return f"Weaving {v}"
    return v


def read_xuatkho(filepath: str, file_name: str = None) -> list[dict]:
    """Đọc sheet XUATKHO từ file Excel."""
    fname = file_name or os.path.basename(filepath)
    try:
        df = pd.read_excel(filepath, sheet_name="XUATKHO", header=1)
    except Exception as e:
        return []

    records = []
    for _, row in df.iterrows():
        ma_beam = _norm(row.get("MÃ BEAM"))
        if not ma_beam or ma_beam in ("nan", "MÃ BEAM"):
            continue

        records.append({
            "file_name":       fname,
            "ma_beam":         ma_beam,
            "weaving":         _parse_weaving(row.get("1\n2")),
            "so_may":          _norm(row.get("Số \nmáy")),
            "loai_may":        _norm(row.get("LOẠI SỢI")),
            "ma_soi":          _norm(row.get("MÃ SỢI")),
            "ten_hang":        _norm(row.get("TÊN HÀNG")),
            "tong_soi":        _norm_float(row.get("TỔNG SỢI")),
            "loai_soi":        _norm(row.get("LOẠI SỢI.1")),
            "beam_p":          _norm(row.get("P")),
            "beam_g":          _norm(row.get("G")),
            "so_met":          _norm_float(row.get("SỐ\nMÉT")),
            "phan_loai":       _norm(row.get("PD/\nYD")),
            "ngay_len_may":    str(row.get("NGÀY", ""))[:10] if pd.notna(row.get("NGÀY")) else None,
            "gio_len_may":     _norm(row.get("GIỜ")),
            "so_gian":         _norm(row.get("SỐ GIÀN")),
            "so_met_thuc_te":  _norm_float(row.get("SỐ MÉT THỰC TẾ")),
            "so_kg_thuc_te":   _norm_float(row.get("SỐ KG\nTHỰC TẾ")),
            "chi_thi_beam":    _norm(row.get("SỐ CHỈ THỊ BEAM")),
            "cm_thut_vao":     _norm(row.get("SỐ CM THỤT VÀO")),
            "loai_mam":        _norm(row.get("LOẠI MÂM BEAM")),
            "nguoi_xuat":      _norm(row.get("NGƯỜI XUẤT")),
            "ghi_chu":         _norm(row.get("GHI CHÚ")),
        })
    return records


def read_yccb(filepath: str, file_name: str = None) -> list[dict]:
    """Đọc sheet YCCB (beam yêu cầu) từ file Excel."""
    fname = file_name or os.path.basename(filepath)
    try:
        df = pd.read_excel(filepath, sheet_name="YCCB", header=5)
    except Exception:
        return []

    records = []
    for _, row in df.iterrows():
        ma_beam = _norm(row.get("MÃ YÊU CẦU\n빔번호"))
        if not ma_beam:
            continue
        ngay = row.get("NGÀY GIAO\n출고요청일")
        records.append({
            "file_name":   fname,
            "ma_beam":     ma_beam,
            "weaving":     _parse_weaving(row.get("WEA\n공장구분")),
            "so_may":      _norm(row.get("MÁY\n호수")),
            "loai_may":    _norm(row.get("LOẠI MÁY\n직기")),
            "ma_soi":      _norm(row.get("MÃ SỢI \n본수")),
            "ten_hang":    _norm(row.get("TÊN HÀNG\n제품명")),
            "tong_soi":    _norm_float(row.get("TỔNG SỐ SỢI\n총본수")),
            "loai_soi":    _norm(row.get("LOẠI SỢI\n사종")),
            "beam_p":      _norm(row.get("P")),
            "beam_g":      _norm(row.get("G")),
            "so_met":      _norm_float(row.get("SỐ MÉT\n빔m")),
            "phan_loai":   _norm(row.get("PD/YD")),
            "ngay_giao":   str(ngay)[:10] if pd.notna(ngay) else None,
            "gio_giao":    _norm(row.get("GIỜ GIAO\n출고요청시간")),
            "ghi_chu":     _norm(row.get("GHI CHÚ\n빔출고시간")),
        })
    return records


def upsert_beam_info(records: list[dict]) -> int:
    if not records: return 0
    conn = sqlite3.connect(DB_NAME)
    n = 0
    for r in records:
        try:
            conn.execute("""
                INSERT OR REPLACE INTO Beam_Info
                (file_name,ma_beam,po_code,ngay_yc,ngay_giao,weaving,so_may,loai_may,
                 ma_soi,ten_hang,tong_soi,loai_soi,beam_p,beam_g,so_met,phan_loai,
                 ngay_len_may,gio_len_may,so_gian,so_met_thuc_te,so_kg_thuc_te,
                 chi_thi_beam,cm_thut_vao,loai_mam,nguoi_xuat,ghi_chu)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                r["file_name"], r["ma_beam"],
                r.get("po_code"), r.get("ngay_yc"), r.get("ngay_giao"),
                r["weaving"], r["so_may"], r["loai_may"],
                r["ma_soi"], r["ten_hang"], r["tong_soi"],
                r["loai_soi"], r["beam_p"], r["beam_g"], r["so_met"],
                r["phan_loai"], r["ngay_len_may"], r["gio_len_may"],
                r["so_gian"], r["so_met_thuc_te"], r["so_kg_thuc_te"],
                r["chi_thi_beam"], r["cm_thut_vao"], r["loai_mam"],
                r["nguoi_xuat"], r["ghi_chu"],
            ))
            n += 1
        except Exception:
            pass
    conn.commit(); conn.close()
    return n


def upsert_beam_request(records: list[dict]) -> int:
    if not records: return 0
    conn = sqlite3.connect(DB_NAME)
    n = 0
    for r in records:
        try:
            conn.execute("""
                INSERT OR REPLACE INTO Beam_Request
                (file_name,ma_beam,weaving,so_may,loai_may,ma_soi,ten_hang,
                 tong_soi,loai_soi,beam_p,beam_g,so_met,phan_loai,
                 ngay_giao,gio_giao,ghi_chu)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                r["file_name"], r["ma_beam"], r["weaving"], r["so_may"],
                r["loai_may"], r["ma_soi"], r["ten_hang"], r["tong_soi"],
                r["loai_soi"], r["beam_p"], r["beam_g"], r["so_met"],
                r["phan_loai"], r["ngay_giao"], r["gio_giao"], r["ghi_chu"],
            ))
            n += 1
        except Exception:
            pass
    conn.commit(); conn.close()
    return n


def import_beam_file(filepath: str) -> dict:
    """
    Đọc 1 file Excel và lưu cả XUATKHO + YCCB vào DB.
    Returns: {xuatkho: n, yccb: n, errors: []}
    """
    init_beam_tables()
    fname = os.path.basename(filepath)
    errors = []

    recs_xuat  = read_xuatkho(filepath, fname)
    recs_yccb  = read_yccb(filepath, fname)
    recs_total = read_total(filepath, fname)

    n_xuat  = upsert_beam_info(recs_xuat)
    n_yccb  = upsert_beam_request(recs_yccb)
    n_total = upsert_beam_info(recs_total)  # TOTAL → Beam_Info

    return {
        "xuatkho": n_xuat,
        "yccb":    n_yccb,
        "total_sheet": n_total,
        "total":   n_xuat + n_yccb + n_total,
        "errors":  errors,
    }


def get_beam_on_machine(weaving: str, so_may: str) -> dict:
    """
    Lấy thông tin beam đang trên máy từ Beam_Info.
    Dùng cho tính toán beam còn lại.
    so_may: '12', '12.0' đều được
    """
    conn = sqlite3.connect(DB_NAME)
    try:
        try:
            mac_int = str(int(float(so_may)))
        except:
            mac_int = so_may
        df = pd.read_sql_query(f"""
            SELECT * FROM Beam_Info
            WHERE weaving = '{weaving}'
              AND (so_may = '{so_may}' OR so_may = '{mac_int}')
            ORDER BY ngay_len_may DESC, id DESC
            LIMIT 1
        """, conn)
        if df.empty:
            return {}
        return df.iloc[0].to_dict()
    finally:
        conn.close()


def calc_beam_remaining_v2(weaving: str, so_may: str,
                           kg_used_since_install: float,
                           soi_bong_pct: float = 0,
                           soi_nen_pct: float = 0) -> dict:
    """
    Tính beam còn lại dựa trên Beam_Info (số mét ban đầu) + kg đã dệt.
    
    so_met ban đầu = so_met_thuc_te (nếu có) hoặc so_met
    Mật độ (kg/m):
      - Nếu có so_kg_thuc_te và so_met_thuc_te → dùng trực tiếp
      - Nếu không → dùng loại sợi từ bảng 'full' sheet (MTR-KG)
    
    Beam trên (sợi bông) = soi_bong_pct % tổng kg dệt
    Beam dưới (sợi nền)  = soi_nen_pct  % tổng kg dệt
    """
    b = get_beam_on_machine(weaving, so_may)
    if not b:
        return {"error": f"Không có dữ liệu beam cho {weaving} Máy {so_may}"}

    initial_m   = b.get("so_met_thuc_te") or b.get("so_met") or 0
    initial_kg  = b.get("so_kg_thuc_te")
    ma_beam     = b.get("ma_beam", "")
    ten_hang    = b.get("ten_hang", "")
    ngay_len    = b.get("ngay_len_may", "")
    loai_soi    = b.get("loai_soi", "")
    phan_loai   = b.get("phan_loai", "")

    # kg/m = initial_kg / initial_m (nếu có đủ dữ liệu)
    kg_per_m = None
    if initial_kg and initial_m and float(initial_m) > 0:
        kg_per_m = float(initial_kg) / float(initial_m)

    def _calc_one(pct, label):
        if not initial_m:
            return {"initial_m": None, "used_m": None, "remaining_m": None,
                    "remaining_pct": None, "label": label}
        kg_beam = kg_used_since_install * (pct / 100)
        if kg_per_m and kg_per_m > 0:
            used_m = round(kg_beam / kg_per_m, 1)
            rem_m  = round(max(float(initial_m) - used_m, 0), 1)
            rem_pct= round(rem_m / float(initial_m) * 100, 1)
        else:
            used_m = None; rem_m = None; rem_pct = None
        return {
            "initial_m":     round(float(initial_m), 0),
            "used_kg":       round(kg_beam, 1),
            "used_m":        used_m,
            "remaining_m":   rem_m,
            "remaining_pct": rem_pct,
            "label":         label,
        }

    return {
        "machine":           f"{weaving} — Máy {so_may}",
        "ma_beam":           ma_beam,
        "ten_hang":          ten_hang,
        "loai_soi":          loai_soi,
        "phan_loai":         phan_loai,
        "ngay_len_may":      ngay_len,
        "initial_m":         initial_m,
        "initial_kg":        initial_kg,
        "kg_per_meter":      round(kg_per_m, 5) if kg_per_m else None,
        "kg_used_total":     round(kg_used_since_install, 1),
        "beam_tren":         _calc_one(soi_bong_pct, "Beam trên (sợi bông)"),
        "beam_duoi":         _calc_one(soi_nen_pct,  "Beam dưới (sợi nền)"),
    }

def read_total(filepath: str, file_name: str = None) -> list[dict]:
    """
    Đọc sheet TOTAL — bảng tổng hợp toàn bộ lịch sử beam.
    Header ở row 8, data bắt đầu từ row 9.
    """
    fname = file_name or os.path.basename(filepath)
    try:
        df = pd.read_excel(filepath, sheet_name="TOTAL", header=8)
    except Exception:
        return []

    # Rename columns to standard names
    col_map = {
        "MÃ BEAM":               "ma_beam",
        "요청일자\nNgày Y/C":      "ngay_yc",
        "납기일\nNgày giao":       "ngay_giao",
        "1동\n2동":               "weaving_raw",
        "PO CODE":                "po_code",
        "Mã Sợi":                 "ma_soi",
        "MC\nNO":                "so_may",
        "MC TYPLE":               "loai_may",
        "Tên hàng":               "ten_hang",
        "총본수\nTổng số sợi":     "tong_soi",
        "구분 \nPhân loại":       "beam_phan_loai",   # P/G
        "사종\nLoại Sợi":         "loai_soi",
        "빔크기\nKích thước beam": "kich_thuoc_beam",
        "Khối lượng beam    KG":  "beam_wgt_kg",
        "빔수량\nSố Lượng":        "so_luong",
        "PD\nYD":                "phan_loai",
        "Ngày yêu cầu beam":     "ngay_yc_beam",
        "SỐ MÉT NHẬN":           "so_met_nhan",
    }
    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    records = []
    for _, row in df.iterrows():
        ma_beam = _norm(row.get("ma_beam"))
        if not ma_beam or not str(ma_beam).strip():
            continue
        # Parse weaving: 1/1.0/2.0 → Weaving 1/2/3
        wv_raw = _norm(row.get("weaving_raw"))
        if wv_raw:
            import re as _re
            _wm = _re.match(r"^(\d+)(\.0)?$", str(wv_raw).strip())
            weaving = f"Weaving {int(_wm.group(1))}" if _wm else wv_raw
        else:
            weaving = wv_raw

        records.append({
            "file_name":      fname,
            "ma_beam":        ma_beam,
            "weaving":        weaving,
            "so_may":         _norm(row.get("so_may")),
            "loai_may":       _norm(row.get("loai_may")),
            "ma_soi":         _norm(row.get("ma_soi")),
            "ten_hang":       _norm(row.get("ten_hang")),
            "tong_soi":       _norm_float(row.get("tong_soi")),
            "loai_soi":       _norm(row.get("loai_soi")),
            "beam_p":         "P" if str(row.get("beam_phan_loai","")).strip().upper() == "P" else None,
            "beam_g":         "G" if str(row.get("beam_phan_loai","")).strip().upper() == "G" else None,
            "so_met":         _norm_float(row.get("kich_thuoc_beam")),
            "so_kg_thuc_te":  _norm_float(row.get("beam_wgt_kg")),
            "phan_loai":      _norm(row.get("phan_loai")),
            "ngay_len_may":   str(row.get("ngay_yc_beam", ""))[:10] if pd.notna(row.get("ngay_yc_beam")) else None,
            "po_code":        _norm(row.get("po_code")),
            "ngay_yc":        str(row.get("ngay_yc",""))[:10] if pd.notna(row.get("ngay_yc")) else None,
            "ngay_giao":      str(row.get("ngay_giao",""))[:10] if pd.notna(row.get("ngay_giao")) else None,
            "so_met_thuc_te": _norm_float(row.get("so_met_nhan")),
            "ghi_chu":        None,
            "gio_len_may":    None,
            "so_gian":        None,
            "chi_thi_beam":   None,
            "cm_thut_vao":    None,
            "loai_mam":       None,
            "nguoi_xuat":     None,
        })
    return records