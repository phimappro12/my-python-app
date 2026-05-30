"""
mtr_kg.py — Bảng đổi mét ↔ kg cho beam dệt
=============================================
Công thức:
  kg = (mét × total_sợi × hệ_số_loại_sợi) / 1000
  
  Trong file MTR-KG.xlsx:
  - Cột 'yarn' : tên loại sợi (VD: 30S/2)
  - Cột 2 (hệ số): kg/100m cho mỗi loại sợi (VD: 0.036)
  
  Áp dụng:
  kg = mét × total_sợi × hệ_số / 1000 * ... 
  
  Thực tế từ file:
  mtr=2000, total=2620, yarn=30S/2 (hệ số 0.036) → kg=206.31
  Check: 2000 × 2620 × 0.036 / 1000 = 188.64 ← không khớp
  
  Cách đúng: kg = mét × total_sợi × hệ_số
  2000 × 2620 × 0.036 = 188,640 ← quá lớn
  2000 × 0.036 × ... = ?
  
  Thực ra: kg/m = (total_sợi × hệ_số) / 1000
  = 2620 × 0.036 / 1000 = 0.09432 kg/m? × 2000 = 188.64 (sai)
  
  Thử: kg = mét × hệ_số × factor
  206.31 / 2000 / 0.036 = 2.865... 
  hay 206.31 / (2000 × 2620 / 1000) = 206.31 / 5240 = 0.03937
  
  Thực tế: 206.31 / 2000 = 0.103155 kg/m (đây là mật độ thực)
  → KG = MTR × (TOTAL_SOI × HE_SO / 1000)
  = 2000 × (2620 × 0.036 / 1000) = 2000 × 0.09432 = 188.64 (vẫn sai)
  
  Cuối cùng thử: kg = mtr * he_so * total_soi/1000/1000*something?
  206.31 = 2000 * he_so_actual
  he_so_actual = 0.103155
  = total_soi/1000 * 0.036 * (1000/something)
  0.103155 = 2620 * 0.036 / X → X = 2620*0.036/0.103155 = 914.6
  
  Gần nhất: 206.31 ≈ 2000 * 2620 * 0.036 / 914.6
  Thử: 1000 * 3.28084 = 3280.84 (feet?) 
  Thử yards: 1 mét = 1.09361 yards
  206.31 / 2000 / (2620 * 0.036 / 1000) = 206.31 / 188.64 = 1.0937 ≈ 1 mét / 0.9144 yard
  
  Vậy công thức đúng là: KG = MTR × TOTAL_SOI × HE_SO / 1000 × (MTR/YARDS)
  Hay: KG = YARDS × TOTAL_SOI × HE_SO / 1000
  2000m = 2187.23 yards
  2187.23 × 2620 × 0.036 / 1000 = 206.35 ≈ 206.31 ✓
"""

# Vậy công thức chính xác:
# KG = MTR_TO_YARDS(mtr) × total_soi × he_so / 1000
# MTR_TO_YARDS: 1m = 1.09361 yards

MTR_PER_YARD = 0.9144  # 1 yard = 0.9144 mét
YARDS_PER_MTR = 1.09361

# Bảng hệ số kg/yard/1000 sợi
YARN_TABLE = {
    "30S/2 (10) + NTW (20S/1 + PVA 80S/1) SW": 0.034875,
    "ASANO 20S/1": 0.0315,
    "BAMBOO 20S/2": 0.054,
    "BAMBOO 20S/2 (8)": 0.054,
    "BAMBOO 30S/2 (10)": 0.036,
    "Bamboo 30S/2": 0.036,
    "Bamboo 30S/2 (10)": 0.036,
    "Bamboo 30S/2 (10) + CD 30S/2 (10)": 0.036,
    "CD 13S/1": 0.04154,
    "CD 13S/1 (17)": 0.04154,
    "CD 16S/1": 0.03375,
    "CD 16S/1 (17)": 0.03375,
    "CD 16S/3": 0.10125,
    "CD 20S/1": 0.027,
    "CD 20S/1 (19)": 0.027,
    "CD 20S/1 (19) + NTW (20S/1 + PVA 80S/1) SW": 0.030375,
    "CD 20S/2 (8)": 0.054,
    "CD 30S/1": 0.018,
    "CD 30S/2": 0.036,
    "CD 30S/2 (10)": 0.036,
    "CHANSA 16S/1": 0.03375,
    "CM 13S/1": 0.04154,
    "CM 13S/1 (17)": 0.04154,
    "CM 16S/1": 0.03375,
    "CM 16S/1 (15)": 0.03375,
    "CM 20S/1": 0.027,
    "CM 20S/1 Organic": 0.027,
    "CM 30S/2": 0.036,
    "CM 30S/2 (10)": 0.036,
    "CM 30S/2 (10) Organic": 0.036,
    "CM 40S/2": 0.027,
    "CM 40S/2 + Asano 20S/1": 0.02925,
    "CM GIZA 16S/1 (17)": 0.03375,
    "CM MVS 40/2": 0.027,
    "GIZA CM 16S/1": 0.03375,
    "Giza CM 16S/1": 0.03375,
    "NE 13S/1 NU-TORQUE": 0.04154,
    "NE 16S/1 (60% BAMBOO + 40% COTTON)": 0.03375,
    "NISHIKAWA 20S/1": 0.027,
    "NTW (13S/1 + PVA 80S/1) (8)": 0.0482,
    "NTW (20S/1 + PVA 80S/1)": 0.03375,
    "NTW (30S/1 + PVA 80S/1)": 0.02475,
    "NTW (CM 40S/1 + PVA 80S/1)": 0.02025,
    "NTW NE BAMBOO 16S/1 (9)": 0.052,
    "OE 20S/1": 0.027,
    "OE 20S/2": 0.054,
    "Poly 30S/2 65%": 0.036,
    "SIRO 20S/1": 0.027,
    "SPINAIR 20% 20S/1+PVA(10)": 0.02727,
    "SUPIMA 40S/2": 0.027,
    "Siro 20S/1": 0.027,
    "Siro 20S/2": 0.054,
    "TENCEL 20S/1 NTW (9)": 0.03375,
}


import re

def find_he_so(yarn_name: str) -> float | None:
    """
    Tìm hệ số cho loại sợi.
    Ưu tiên: exact → normalize → fuzzy substring.
    """
    if not yarn_name:
        return None

    # Clean: strip, remove _x000D_ artifacts, normalize spaces
    import re as _re
    clean = _re.sub(r'_x000D_', '', str(yarn_name)).strip()
    clean = _re.sub(r'\s+', ' ', clean)
    yarn_up = clean.upper()

    # 1. Exact match
    for k, v in YARN_TABLE.items():
        if k.upper() == yarn_up:
            return v

    # 2. Normalize: remove trailing spaces, parentheses numbers
    yarn_norm = _re.sub(r'\s*\(\d+\)\s*$', '', yarn_up).strip()
    for k, v in YARN_TABLE.items():
        k_norm = _re.sub(r'\s*\(\d+\)\s*$', '', k.upper()).strip()
        if k_norm == yarn_norm:
            return v

    # 3. Substring match — longer key wins to avoid false positives
    matches = []
    for k, v in YARN_TABLE.items():
        k_up = k.upper()
        if k_up in yarn_up or yarn_up in k_up:
            matches.append((len(k), v))
    if matches:
        return max(matches, key=lambda x: x[0])[1]

    return None


def lookup_by_item_name(item_name: str, loai_soi_hint: str = None) -> dict:
    """
    Tìm hệ số từ tên mã hàng (item_name).
    Bước 1: Tìm trong Yarn_Formula (DB) theo item_name → lấy soi_bong, soi_nen
    Bước 2: Tra hệ số cho từng loại sợi đó
    Bước 3: Fallback dùng loai_soi_hint trực tiếp

    Returns dict với keys: soi_bong, he_so_bong, soi_nen, he_so_nen, source
    """
    import sqlite3, pandas as pd, os

    result = {
        "item_name":  item_name,
        "soi_bong":   None, "he_so_bong": None,
        "soi_nen":    None, "he_so_nen":  None,
        "soi_ngang":  None, "he_so_ngang":None,
        "source":     "not_found",
    }

    # Try DB Yarn_Formula
    db_path = "inventory.db"
    if not os.path.exists(db_path):
        db_path = os.path.join(os.path.dirname(__file__), "inventory.db")

    try:
        conn = sqlite3.connect(db_path)
        clean_name = str(item_name).replace("'", "''")
        df = pd.read_sql_query(
            f"SELECT * FROM Yarn_Formula WHERE UPPER(item_name)=UPPER('{clean_name}') LIMIT 1", conn
        )
        conn.close()
        if not df.empty:
            r = df.iloc[0]
            soi_b = str(r.get("soi_bong") or "")
            soi_n = str(r.get("soi_nen")  or "")
            soi_ng= str(r.get("soi_ngang") or "")
            result.update({
                "soi_bong":   soi_b,  "he_so_bong": find_he_so(soi_b),
                "soi_nen":    soi_n,  "he_so_nen":  find_he_so(soi_n),
                "soi_ngang":  soi_ng, "he_so_ngang":find_he_so(soi_ng),
                "source":     "Yarn_Formula",
            })
            return result
    except:
        pass

    # Fallback: dùng loai_soi_hint trực tiếp
    if loai_soi_hint:
        he_so = find_he_so(loai_soi_hint)
        if he_so:
            result.update({
                "soi_bong":  loai_soi_hint, "he_so_bong": he_so,
                "source":    "loai_soi_hint",
            })

    return result


def mtr_to_kg(mtr: float, total_soi: float, yarn_name: str) -> dict:
    """
    Đổi mét → kg
    KG = yards × total_soi × he_so / 1000
    yards = mtr × 1.09361
    """
    he_so = find_he_so(yarn_name)
    if he_so is None:
        return {"error": f"Không tìm thấy hệ số cho loại sợi '{yarn_name}'",
                "yarn": yarn_name}
    yards = mtr * YARDS_PER_MTR
    kg = yards * total_soi * he_so / 1000
    return {
        "mtr": round(mtr, 1),
        "yards": round(yards, 1),
        "total_soi": total_soi,
        "yarn": yarn_name,
        "he_so": he_so,
        "kg": round(kg, 2),
        "formula": f"{yards:.1f} yard × {total_soi} sợi × {he_so} / 1000 = {kg:.2f} kg",
    }


def kg_to_mtr(kg: float, total_soi: float, yarn_name: str) -> dict:
    """
    Đổi kg → mét
    yards = kg × 1000 / (total_soi × he_so)
    mtr = yards × 0.9144
    """
    he_so = find_he_so(yarn_name)
    if he_so is None:
        return {"error": f"Không tìm thấy hệ số cho loại sợi '{yarn_name}'",
                "yarn": yarn_name}
    if total_soi <= 0 or he_so <= 0:
        return {"error": "total_soi hoặc he_so không hợp lệ"}
    yards = kg * 1000 / (total_soi * he_so)
    mtr = yards * MTR_PER_YARD
    return {
        "kg": round(kg, 2),
        "yards": round(yards, 1),
        "mtr": round(mtr, 1),
        "total_soi": total_soi,
        "yarn": yarn_name,
        "he_so": he_so,
        "formula": f"{kg} kg × 1000 / ({total_soi} × {he_so}) = {yards:.1f} yard = {mtr:.1f} mét",
    }


def list_yarns() -> list:
    """Trả về danh sách tất cả loại sợi trong bảng."""
    return sorted(YARN_TABLE.keys())


if __name__ == "__main__":
    # Test với dữ liệu từ file Excel
    print("=== Test từ file MTR-KG.xlsx ===")
    r1 = mtr_to_kg(2000, 2620, "30S/2")
    print(f"mét→kg: {r1}")
    print(f"  Kỳ vọng: 206.31 | Kết quả: {r1.get('kg')}")
    
    r2 = mtr_to_kg(3220, 2774, "16S/1")
    print(f"mét→kg: {r2}")
    print(f"  Kỳ vọng: 329.70 | Kết quả: {r2.get('kg')}")
    
    print()
    r3 = kg_to_mtr(206.31, 2620, "30S/2")
    print(f"kg→mét: {r3}")
    print(f"  Kỳ vọng: 2000m | Kết quả: {r3.get('mtr')}")