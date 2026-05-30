"""
beam_calculator.py — Tính beam còn lại trên máy
================================================================
Logic:
  - Beam_Info (TOTAL sheet): ngay_len_may MỚI NHẤT ≤ hôm nay = beam hiện tại
  - Khi có beam mới lên (ngay_len_may sau hơn), beam cũ bị thay thế
  - kg/m = so_kg_thuc_te / so_met
  - mét dùng = kg_dệt_từ_ngày_lắp / kg_per_m
  - còn lại = so_met − mét dùng
"""
import sqlite3
import pandas as pd
from datetime import date

DB_NAME = "inventory.db"


def get_current_beam(weaving: str, so_may: str, as_of_date: str = None) -> dict:
    if not as_of_date:
        as_of_date = str(date.today())
    try:
        mac_int = str(int(float(so_may)))
    except:
        mac_int = so_may
    conn = sqlite3.connect(DB_NAME)
    try:
        df = pd.read_sql_query(f"""
            SELECT * FROM Beam_Info
            WHERE weaving = '{weaving}'
              AND (so_may = '{so_may}' OR so_may = '{mac_int}')
              AND ngay_len_may IS NOT NULL
              AND ngay_len_may <= '{as_of_date}'
            ORDER BY ngay_len_may DESC LIMIT 1
        """, conn)
        return df.iloc[0].to_dict() if not df.empty else {}
    finally:
        conn.close()


def get_all_current_beams(weaving: str = None, as_of_date: str = None) -> pd.DataFrame:
    if not as_of_date:
        as_of_date = str(date.today())
    conn = sqlite3.connect(DB_NAME)
    try:
        where = f"ngay_len_may <= '{as_of_date}' AND ngay_len_may IS NOT NULL AND weaving IS NOT NULL"
        if weaving:
            where += f" AND weaving = '{weaving}'"
        df = pd.read_sql_query(f"SELECT * FROM Beam_Info WHERE {where}", conn)
        if df.empty:
            return df
        df = df.sort_values("ngay_len_may", ascending=False)
        df = df.drop_duplicates(subset=["weaving", "so_may"], keep="first")
        return df.reset_index(drop=True)
    finally:
        conn.close()


def get_kg_used(weaving: str, so_may: str, from_date: str, to_date: str = None) -> float:
    if not to_date:
        to_date = str(date.today())
    try:
        mac_int = str(int(float(so_may)))
    except:
        mac_int = so_may
    conn = sqlite3.connect(DB_NAME)
    try:
        df = pd.read_sql_query(f"""
            SELECT COALESCE(
                SUM(CAST(REPLACE(REPLACE(quantity_kg,',',''),' ','') AS REAL)),0) AS total_kg
            FROM Inventory_Log
            WHERE cluster_name = 'Xưởng Dệt'
              AND sub_location = '{weaving}'
              AND (ten_may = '{so_may}' OR ten_may = '{mac_int}' OR ten_may = '{mac_int}.0')
              AND date >= '{from_date}' AND date <= '{to_date}'
        """, conn)
        return float(df.iloc[0, 0] or 0)
    finally:
        conn.close()


def get_yarn_formula(item_name: str) -> dict:
    conn = sqlite3.connect(DB_NAME)
    try:
        clean = str(item_name).replace("'", "''")
        df = pd.read_sql_query(f"""
            SELECT * FROM Yarn_Formula
            WHERE UPPER(item_name) = UPPER('{clean}')
            ORDER BY id DESC LIMIT 1
        """, conn)
        if df.empty:
            return {}
        r = df.iloc[0]
        return {
            "ten_may":       str(r.get("ten_may") or ""),
            "soi_bong_type": str(r.get("soi_bong") or "Sợi bông"),
            "soi_bong_pct":  float(r.get("soi_bong_pct") or 0),
            "soi_nen_type":  str(r.get("soi_nen") or "Sợi nền"),
            "soi_nen_pct":   float(r.get("soi_nen_pct") or 0),
        }
    except:
        return {}
    finally:
        conn.close()


def calc_beam_remaining(weaving: str, so_may: str, as_of_date: str = None) -> dict:
    if not as_of_date:
        as_of_date = str(date.today())

    beam = get_current_beam(weaving, so_may, as_of_date)
    if not beam:
        return {"error": f"Không tìm thấy beam cho {weaving} Máy {so_may}", "machine": f"{weaving} — Máy {so_may}"}

    initial_m  = float(beam.get("so_met") or 0)
    initial_kg = float(beam.get("so_kg_thuc_te") or 0)
    install_dt = str(beam.get("ngay_len_may") or "")[:10]
    item_name  = str(beam.get("ten_hang") or "").strip()
    ma_beam    = str(beam.get("ma_beam") or "")
    loai_may   = str(beam.get("loai_may") or "")
    phan_loai  = str(beam.get("phan_loai") or "")

    kg_per_m  = round(initial_kg / initial_m, 5) if (initial_m > 0 and initial_kg > 0) else None
    tong_soi  = float(beam.get("tong_soi") or 0)
    kg_used  = get_kg_used(weaving, so_may, install_dt, as_of_date) if install_dt else 0.0
    formula  = get_yarn_formula(item_name)
    pct_bong = (formula.get("soi_bong_pct", 0) / 100) if formula else 0
    pct_nen  = (formula.get("soi_nen_pct",  0) / 100) if formula else 0

    # Lấy loai_soi và tong_soi trực tiếp từ Beam_Info để dùng công thức hệ số
    _loai_soi_beam = str(beam.get("loai_soi") or "").strip()
    _he_so = None
    _he_so_formula = ""
    if tong_soi and tong_soi > 0 and _loai_soi_beam:
        try:
            from mtr_kg import find_he_so as _fhs
            _he_so = _fhs(_loai_soi_beam)
            if _he_so:
                _he_so_formula = f"{_loai_soi_beam} → hệ số {_he_so}"
        except Exception:
            pass

    def _calc(pct, soi_type):
        if not initial_m:
            return {"initial_m": None, "used_kg": 0, "used_m": None,
                    "remaining_m": None, "remaining_pct": None,
                    "soi_type": soi_type, "method": "no_data"}
        kg_beam = round(kg_used * pct, 2)
        used_m  = None; rem_m = None; rem_pct = None
        method  = "unknown"

        # Ưu tiên: dùng hệ số sợi chính thức (mtr_kg.py)
        # Công thức: mét = kg × 1000 / (tong_soi × he_so) × 0.9144
        if _he_so and tong_soi and tong_soi > 0:
            try:
                from mtr_kg import kg_to_mtr as _k2m
                _r = _k2m(kg_beam, tong_soi, _loai_soi_beam)
                if "mtr" in _r and _r["mtr"] is not None:
                    used_m  = round(_r["mtr"], 0)
                    rem_m   = round(max(initial_m - used_m, 0), 0)
                    rem_pct = round(rem_m / initial_m * 100, 1)
                    method  = f"hệ_số ({_he_so_formula})"
            except Exception:
                pass

        # Fallback: kg/m từ số liệu thực tế của beam (so_kg/so_met)
        if used_m is None and kg_per_m and kg_per_m > 0:
            used_m  = round(kg_beam / kg_per_m, 0)
            rem_m   = round(max(initial_m - used_m, 0), 0)
            rem_pct = round(rem_m / initial_m * 100, 1)
            method  = f"kg/m_beam ({kg_per_m} kg/m)"

        return {"initial_m": round(initial_m,0), "used_kg": kg_beam,
                "used_m": used_m, "remaining_m": rem_m,
                "remaining_pct": rem_pct, "soi_type": soi_type,
                "method": method}

    # Loại sợi hiển thị: từ Yarn_Formula nếu có, fallback về loai_soi của beam
    _soi_bong_type = formula.get("soi_bong_type", _loai_soi_beam or "Sợi bông") if formula else _loai_soi_beam or "Sợi bông"
    _soi_nen_type  = formula.get("soi_nen_type",  "Sợi nền")  if formula else "Sợi nền"

    beam_tren = _calc(pct_bong, _soi_bong_type)
    beam_duoi = _calc(pct_nen,  _soi_nen_type)

    note = ""
    if not formula:
        note = f"Chưa có CT sợi cho '{item_name}' — dùng loai_soi beam để tính"
    if not _he_so:
        note += f" | Không tìm thấy hệ số cho '{_loai_soi_beam}'"
    elif _he_so and note == "":
        note = f"Dùng hệ số sợi: {_he_so_formula}"

    return {
        "machine": f"{weaving} — Máy {so_may}", "ma_beam": ma_beam,
        "ten_hang": item_name, "loai_may": loai_may, "phan_loai": phan_loai,
        "ngay_len_may": install_dt, "as_of_date": as_of_date,
        "initial_m": round(initial_m,0), "initial_kg": round(initial_kg,2),
        "kg_per_meter": round(kg_per_m,5) if kg_per_m else None,
        "kg_used_total": round(kg_used,2),
        "beam_tren": beam_tren, "beam_duoi": beam_duoi,
        "has_formula": bool(formula), "note": note.strip(" |"),
        # Thêm để ai_agent tính trực tiếp
        "loai_soi": _loai_soi_beam,
        "tong_soi": tong_soi,
        "he_so": _he_so,
        "he_so_formula": _he_so_formula,
        "pct_bong": pct_bong,
        "pct_nen": pct_nen,
    }


def get_beam_status_table(weaving: str = None, as_of_date: str = None) -> pd.DataFrame:
    """Bảng tóm tắt tất cả beam trên máy — dùng cho UI."""
    if not as_of_date:
        as_of_date = str(date.today())
    beams = get_all_current_beams(weaving, as_of_date)
    if beams.empty:
        return pd.DataFrame()
    rows = []
    for _, b in beams.iterrows():
        m = calc_beam_remaining(str(b["weaving"]), str(b["so_may"]), as_of_date)
        if "error" in m:
            continue
        bt, bd = m["beam_tren"], m["beam_duoi"]
        rows.append({
            "Xưởng": m["machine"].split(" — ")[0],
            "Máy":   m["machine"].split("Máy ")[-1],
            "Mã Beam": m["ma_beam"], "Loại máy": m["loai_may"],
            "Tên hàng": m["ten_hang"], "Ngày lên máy": m["ngay_len_may"],
            "Tổng mét": m["initial_m"], "kg/m": m["kg_per_meter"],
            "KG dệt": m["kg_used_total"],
            "Beam trên còn (m)": bt.get("remaining_m"),
            "Beam trên còn (%)": bt.get("remaining_pct"),
            "Sợi bông": bt.get("soi_type"),
            "Beam dưới còn (m)": bd.get("remaining_m"),
            "Beam dưới còn (%)": bd.get("remaining_pct"),
            "Sợi nền": bd.get("soi_type"),
            "PD/YD": m["phan_loai"],
            "CT sợi": "✅" if m["has_formula"] else "❌",
            "Ghi chú": m["note"],
        })
    return pd.DataFrame(rows)


def calc_all_beams(weaving: str = None, as_of_date: str = None) -> list:
    beams = get_all_current_beams(weaving, as_of_date)
    return [calc_beam_remaining(str(r["weaving"]), str(r["so_may"]), as_of_date)
            for _, r in beams.iterrows()]