"""
kho_beam_pipeline.py — Nhập snapshot Kho Beam Weaving theo phương pháp DIFF
============================================================================
Logic:
  - Mỗi ngày user upload 1 snapshot (danh sách beam đang trong kho hôm nay)
  - Hệ thống so sánh vs snapshot ngày hôm trước → tự detect NHAP / XUAT
  - NHAP: beam CÓ trong hôm nay nhưng KHÔNG có hôm qua
  - XUAT: beam CÓ hôm qua nhưng KHÔNG có hôm nay
  - Snapshot được lưu dưới type='SNAPSHOT' để dùng làm base so sánh
"""
import os, re, sqlite3
import pandas as pd
from datetime import date, timedelta

DB_NAME = "inventory.db"
CLUSTER = "Kho Beam Weaving"


def _init_table():
    conn = sqlite3.connect(DB_NAME)
    for col, typ in [
        ("sub_location", "TEXT"), ("xuong", "TEXT"), ("po_code", "TEXT"),
        ("beam_size", "REAL"), ("ten_may", "TEXT"), ("loai_may", "TEXT"),
        ("ten_hang", "TEXT"), ("dyeing_type", "TEXT"), ("total_yarn", "REAL"),
        ("beam_type", "TEXT"), ("ten_soi", "TEXT"),
    ]:
        try:
            conn.execute(f"ALTER TABLE Inventory_Log ADD COLUMN {col} TEXT")
        except Exception:
            pass
    try:
        conn.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS idx_kbw_snapshot
            ON Inventory_Log(cluster_name, date, item_id, sub_location)
            WHERE cluster_name = 'Kho Beam Weaving'
        """)
    except Exception:
        pass
    conn.commit()
    conn.close()


def _parse_erp_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Chuẩn hóa DataFrame từ file ERP (Peace ERP hoặc Excel export).
    Cột quan trọng: Vị trí, Request No / Mã Beam, Xưởng, Tên hàng, Số mét...
    """
    df.columns = [str(c).strip() for c in df.columns]

    def _find(keywords):
        for kw in keywords:
            for c in df.columns:
                if kw.lower() in c.lower():
                    return c
        return None

    col_slot    = _find(["vị trí", "vi tri", "position", "stock", "giàn", "gian"])
    col_beam    = _find(["request no", "mã beam", "ma beam", "beam id", "request"])
    col_xuong   = _find(["xưởng", "xuong", "w1", "weaving"])
    col_po      = _find(["po#", "po code", "po_code", "mã đơn"])
    col_size    = _find(["beam size", "size"])
    col_may     = _find(["tên máy", "ten may", "machine"])
    col_loai    = _find(["loại máy", "loai may", "machine type"])
    col_hang    = _find(["tên hàng", "ten hang", "item"])
    col_dyeing  = _find(["dyeing", "dyeing type"])
    col_yarn    = _find(["total yarn", "tổng sợi", "tong soi"])
    col_btype   = _find(["beam type", "loại beam"])
    col_soi     = _find(["tên sợi", "ten soi", "yarn name"])

    rows = []
    for _, r in df.iterrows():
        beam_id = str(r.get(col_beam, "") or "").strip() if col_beam else ""
        if not beam_id or beam_id in ("nan", "", "None"):
            continue
        slot = str(r.get(col_slot, "") or "").strip() if col_slot else ""
        rows.append({
            "item_id":      beam_id,
            "sub_location": slot,
            "xuong":        str(r.get(col_xuong, "") or "") if col_xuong else "",
            "po_code":      str(r.get(col_po, "")    or "") if col_po    else "",
            "beam_size":    pd.to_numeric(r.get(col_size, None), errors="coerce") if col_size else None,
            "ten_may":      str(r.get(col_may, "")   or "") if col_may   else "",
            "loai_may":     str(r.get(col_loai, "")  or "") if col_loai  else "",
            "ten_hang":     str(r.get(col_hang, "")  or "") if col_hang  else "",
            "dyeing_type":  str(r.get(col_dyeing, "") or "") if col_dyeing else "",
            "total_yarn":   pd.to_numeric(r.get(col_yarn, None), errors="coerce") if col_yarn else None,
            "beam_type":    str(r.get(col_btype, "") or "") if col_btype else "",
            "ten_soi":      str(r.get(col_soi, "")   or "") if col_soi   else "",
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def get_latest_snapshot(as_of_date: str = None) -> pd.DataFrame:
    """Lấy snapshot gần nhất (type=SNAPSHOT) trước ngày as_of_date."""
    if not as_of_date:
        as_of_date = str(date.today())
    conn = sqlite3.connect(DB_NAME)
    try:
        df = pd.read_sql_query(
            f"SELECT * FROM Inventory_Log "
            f"WHERE cluster_name='{CLUSTER}' AND type='SNAPSHOT' "
            f"AND date < '{as_of_date}' "
            f"ORDER BY date DESC LIMIT 5000",
            conn
        )
        if df.empty:
            return df
        # Lấy đúng ngày snapshot mới nhất
        latest_date = df["date"].max()
        return df[df["date"] == latest_date].copy()
    finally:
        conn.close()


def import_snapshot(df_raw: pd.DataFrame, snapshot_date: str = None,
                    file_name: str = "") -> dict:
    """
    Nhập 1 snapshot mới → tính diff vs snapshot cũ → insert NHAP/XUAT/SNAPSHOT.
    Returns: {nhap: int, xuat: int, snapshot: int, date: str}
    """
    _init_table()
    if not snapshot_date:
        snapshot_date = str(date.today())

    df_today = _parse_erp_df(df_raw)
    if df_today.empty:
        return {"nhap": 0, "xuat": 0, "snapshot": 0, "date": snapshot_date,
                "error": "Không đọc được cột Mã Beam từ file"}

    today_set = set(df_today["item_id"].unique())

    # Lấy snapshot hôm trước
    df_prev = get_latest_snapshot(as_of_date=snapshot_date)
    prev_set = set(df_prev["item_id"].unique()) if not df_prev.empty else set()

    nhap_set  = today_set - prev_set   # Có hôm nay, không có hôm qua → NHẬP
    xuat_set  = prev_set - today_set   # Có hôm qua, không có hôm nay → XUẤT

    conn = sqlite3.connect(DB_NAME)

    def _make_rows(beam_ids, tx_type, df_source):
        rows = []
        for bid in beam_ids:
            _subset = df_source[df_source["item_id"] == bid]
            r = _subset.iloc[0].to_dict() if not _subset.empty else {"item_id": bid}
            rows.append({
                "date":         snapshot_date,
                "cluster_name": CLUSTER,
                "item_id":      bid,
                "type":         tx_type,
                "quantity":     1,
                "unit":         "beam",
                "sub_location": r.get("sub_location", ""),
                "xuong":        r.get("xuong", ""),
                "po_code":      r.get("po_code", ""),
                "beam_size":    r.get("beam_size"),
                "ten_may":      r.get("ten_may", ""),
                "loai_may":     r.get("loai_may", ""),
                "ten_hang":     r.get("ten_hang", ""),
                "dyeing_type":  r.get("dyeing_type", ""),
                "total_yarn":   r.get("total_yarn"),
                "beam_type":    r.get("beam_type", ""),
                "ten_soi":      r.get("ten_soi", ""),
                "file_name":    file_name,
            })
        return rows

    # Xóa transaction cũ cùng ngày (re-import)
    conn.execute(
        f"DELETE FROM Inventory_Log WHERE cluster_name='{CLUSTER}' AND date='{snapshot_date}'"
    )

    # Insert NHAP
    nhap_rows = _make_rows(nhap_set,  "NHAP",     df_today)
    # Insert XUAT (dùng df_prev để lấy metadata beam đã ra)
    xuat_rows = _make_rows(xuat_set,  "XUAT",     df_prev)
    # Insert SNAPSHOT (toàn bộ hôm nay để làm base so sánh ngày mai)
    snap_rows = _make_rows(today_set, "SNAPSHOT",  df_today)

    for rows in [nhap_rows, xuat_rows, snap_rows]:
        if rows:
            pd.DataFrame(rows).to_sql("Inventory_Log", conn, if_exists="append", index=False)

    conn.commit()
    conn.close()

    return {
        "nhap":     len(nhap_set),
        "xuat":     len(xuat_set),
        "snapshot": len(snap_rows),
        "date":     snapshot_date,
        "nhap_list": sorted(nhap_set)[:20],
        "xuat_list": sorted(xuat_set)[:20],
        "prev_date": df_prev["date"].max() if not df_prev.empty else None,
    }


def get_kho_summary(as_of_date: str = None) -> dict:
    """Tồn kho beam tại ngày as_of_date = snapshot mới nhất."""
    if not as_of_date:
        as_of_date = str(date.today())
    df = get_latest_snapshot(as_of_date=as_of_date)
    if df.empty:
        return {"total_beam": 0, "date": as_of_date, "by_xuong": {}}
    by_xuong = {}
    for _, r in df.iterrows():
        xg = str(r.get("xuong") or "Khác")
        by_xuong.setdefault(xg, 0)
        by_xuong[xg] += 1
    return {
        "total_beam": len(df),
        "date":       df["date"].max(),
        "by_xuong":   by_xuong,
        "by_slot_type": {
            "Giá 60": len(df[df["sub_location"].str.contains("60", na=False)]),
            "Giá 95": len(df[df["sub_location"].str.contains("95", na=False)]),
            "Khác":   len(df[~df["sub_location"].str.contains("60|95", na=False)]),
        }
    }