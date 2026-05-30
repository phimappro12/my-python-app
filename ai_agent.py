"""
ai_agent.py - Phiên bản 2.0 (Kiến trúc Python-First)
======================================================
KIẾN TRÚC MỚI:
  1. Python phát hiện Intent (không dùng AI)
  2. Python tự động đọc Schema từ DB
  3. Python xây dựng SQL hoàn toàn (không phụ thuộc AI nhỏ)
  4. Python thực thi nhiều query song song (so sánh, cross-source)
  5. AI chỉ viết báo cáo cuối từ dữ liệu có sẵn

Hỗ trợ:
  - Tháng 1,2,3,4 xưởng 1 máy 1 sản lượng bao nhiêu
  - So sánh máy 1 vs máy 3 (hay 2 xưởng)
  - Ngày 4/5 máy 3 xưởng 3 đang chạy hàng gì
  - So sánh 10 ngày liên tiếp xem xu hướng hiệu suất
  - Kết hợp kho beam + xưởng dệt tính tổng hợp
  - Hỏi theo khoảng ngày, từng máy, tìm mã hàng
"""

from openai import OpenAI
import pandas as pd
import re
import json
from db_manager import execute_query

# =====================================================================
# CẤU HÌNH KẾT NỐI AI SERVER (TAILSCALE)
# =====================================================================
AI_BASE_URL = "http://100.104.121.57:11434/v1"
AI_API_KEY  = "ollama"
AI_MODEL    = "qwen2.5:3b"
DEFAULT_YEAR = "2026"


class DummyChunk:
    def __init__(self, text):
        self.text = text


# =====================================================================
# MODULE 1: SCHEMA DISCOVERY
# =====================================================================
def get_db_schema() -> dict:
    """
    Tự động đọc cấu trúc và metadata thực tế từ database.
    Trả về dict chứa: columns, qty_col, mac_col, item_col,
                      eff_cols, clusters, locations, date_range
    """
    schema = {
        "columns": [],
        "qty_col": None,
        "mac_col": "ten_may",
        "item_col": "item_id",
        "eff_2ca_col": None,
        "eff_a_col": None,
        "eff_b_col": None,
        "clusters": [],
        "locations": [],
        "date_range": {"min": None, "max": None},
    }

    try:
        # Lấy danh sách cột Sizing_Log
        sz_col_df = execute_query("PRAGMA table_info(Sizing_Log)")
        if not sz_col_df.empty:
            schema["sizing_cols"] = sz_col_df["name"].tolist()
        else:
            schema["sizing_cols"] = []
        
        # Lấy danh sách cột thực
        col_df = execute_query("PRAGMA table_info(Inventory_Log)")
        if not col_df.empty:
            schema["columns"] = col_df["name"].tolist()

        cols = schema["columns"]

        # --- Tìm cột sản lượng ---
        for kw in ["quantity_kg", "total\n(kg)", "total_kg"]:
            hit = next((c for c in cols if kw in c.lower()), None)
            if hit:
                schema["qty_col"] = hit
                break
        if not schema["qty_col"]:
            schema["qty_col"] = next(
                (c for c in cols if "quantity" in c.lower()), None
            )

        # --- Tìm cột tên máy ---
        schema["mac_col"] = next(
            (c for c in cols if "ten_may" in c.lower()), "ten_may"
        )

        # --- Tìm cột mã hàng ---
        schema["item_col"] = next(
            (c for c in cols if "item_id" in c.lower()), "item_id"
        )

        # --- Tìm cột màu sắc ---
        schema["color_col"] = next(
            (c for c in cols if c.upper() == "COLOR" or "colour" in c.lower()
             or c.lower() == "mau" or "màu" in c.lower()), None
        )

        # --- Tìm cột hiệu suất ---
        for c in cols:
            cl = c.lower()
            is_eff = "hiệu suất" in cl or "효율" in cl or "efficiency" in cl
            is_a   = ("ca a" in cl or "a반" in cl or "a 반" in cl) and "hs" in cl
            is_b   = ("ca b" in cl or "b반" in cl or "b 반" in cl) and "hs" in cl
            is_2ca = ("2 ca" in cl or "2ca" in cl or "hiệu suất \n2 ca" in cl)
            if is_eff and is_2ca:
                schema["eff_2ca_col"] = c
            elif is_eff and is_a:
                schema["eff_a_col"] = c
            elif is_eff and is_b:
                schema["eff_b_col"] = c
            elif is_eff and not schema["eff_2ca_col"] and not is_a and not is_b:
                schema["eff_2ca_col"] = c

        # --- Metadata ---
        df_meta = execute_query(
            "SELECT DISTINCT cluster_name FROM Inventory_Log "
            "WHERE cluster_name IS NOT NULL"
        )
        schema["clusters"] = df_meta["cluster_name"].tolist() if not df_meta.empty else []

        df_loc = execute_query(
            "SELECT DISTINCT sub_location FROM Inventory_Log "
            "WHERE sub_location IS NOT NULL LIMIT 20"
        )
        schema["locations"] = df_loc["sub_location"].tolist() if not df_loc.empty else []

        df_date = execute_query(
            "SELECT MIN(date) as mn, MAX(date) as mx FROM Inventory_Log "
            "WHERE date IS NOT NULL"
        )
        if not df_date.empty:
            schema["date_range"]["min"] = str(df_date["mn"].iloc[0])
            schema["date_range"]["max"] = str(df_date["mx"].iloc[0])

    except Exception as e:
        schema["error"] = str(e)

    return schema


# =====================================================================
# MODULE 2: INTENT DETECTION (thuần Python, không cần AI)
# =====================================================================
def detect_intent(prompt: str) -> dict:
    """
    Phân tích câu hỏi thành các tín hiệu cụ thể để xây SQL.
    
    Trả về:
      query_type  : EXPLAIN | SIMPLE | MULTI_PERIOD | COMPARISON | CROSS_SOURCE
      periods     : danh sách tháng [1,2,3,4]
      machines    : danh sách số máy ['1','3']
      locations   : ['Weaving 1','Weaving 3']
      date_range  : {'from': 'YYYY-MM-DD', 'to': 'YYYY-MM-DD'} hoặc None
      exact_date  : 'YYYY-MM-DD' hoặc None
      item_kw     : mã hàng tìm kiếm hoặc None
      clusters    : ['xưởng dệt', 'kho beam yarn', ...]
      flags       : dict các cờ boolean
    """
    p = prompt.lower().strip()

    # ----------------------------------------------------------------
    # 0. Giải thích / tại sao
    # ----------------------------------------------------------------
    explain_kw = [
        "tính thế nào", "tính sao", "tại sao", "vì sao", "sao lại",
        "nguyên nhân", "lý do", "từ đâu ra", "cách tính",
        "số này ở đâu", "vậy sao", "công thức",
        # Thêm các biến thể phổ biến
        "tính toán thế nào", "tính toán sao", "ra được", "con số này",
        "từ đâu", "tính ra sao", "giải thích", "tại sao lại",
        "đúng không", "có chính xác", "tin được không", "chính xác không",
        "kiểm tra lại", "số liệu có đúng",
    ]
    if any(k in p for k in explain_kw):
        return {"query_type": "EXPLAIN"}

    intent = {
        "query_type": "SIMPLE",
        "periods": [],
        "machines": [],
        "locations": [],
        "date_range": None,
        "exact_date": None,
        "item_kw": None,
        "clusters": [],
        "flags": {
            "is_comparison":  False,
            "is_cross_source": False,
            "is_find_machine": False,
            "is_whole_factory": False,
        },
    }

    # ----------------------------------------------------------------
    # 1. Tháng — hỗ trợ "tháng 1,2,3,4" / "tháng 1, 2, 3" / "từ tháng X đến Y"
    # ----------------------------------------------------------------
    _months = set()
    # A: "tháng 1,2,3,4" hoặc "tháng 1, 2, 3"
    _m_block = re.search(r"tháng\s*([\d][,\s\d]*)", p)
    if _m_block:
        for _n in re.findall(r"\d{1,2}", _m_block.group(1)):
            if 1 <= int(_n) <= 12:
                _months.add(int(_n))
    # B: "tháng X" rải rác trong câu
    for _n in re.findall(r"tháng\s*(\d{1,2})", p):
        if 1 <= int(_n) <= 12:
            _months.add(int(_n))
    # C: "từ tháng X đến tháng Y" → range
    _rng = re.search(r"từ\s+tháng\s*(\d{1,2})\s+(?:đến|tới|->)\s+tháng\s*(\d{1,2})", p)
    if _rng:
        _a, _b = int(_rng.group(1)), int(_rng.group(2))
        _months.update(range(min(_a, _b), max(_a, _b) + 1))
    # ✅ FIX: "tháng này/nay" → lấy tháng hiện tại, dùng date_range để filter đúng tháng+năm
    if "tháng này" in p or "tháng nay" in p:
        from datetime import date as _d
        _now = _d.today()
        # Dùng date_range thay vì periods để filter đúng năm (tránh lấy cả tháng 5/2025)
        import calendar as _cal
        _last_day = _cal.monthrange(_now.year, _now.month)[1]
        intent["date_range"] = {
            "from": f"{_now.year}-{_now.month:02d}-01",
            "to":   f"{_now.year}-{_now.month:02d}-{_last_day:02d}",
        }
        intent["flags"]["is_current_month"] = True
    elif "tháng trước" in p:
        from datetime import date as _d
        import calendar as _cal
        _now = _d.today()
        _prev_month = _now.month - 1 if _now.month > 1 else 12
        _prev_year  = _now.year if _now.month > 1 else _now.year - 1
        _last_day = _cal.monthrange(_prev_year, _prev_month)[1]
        intent["date_range"] = {
            "from": f"{_prev_year}-{_prev_month:02d}-01",
            "to":   f"{_prev_year}-{_prev_month:02d}-{_last_day:02d}",
        }

    # ✅ FIX: Thêm xử lý "tuần này/trước/qua/vừa rồi"
    _tuan_nay_kw  = ["tuần này", "tuần nay"]
    _tuan_truoc_kw = ["tuần trước", "tuần qua", "tuần vừa rồi", "tuần vừa qua",
                      "7 ngày qua", "7 ngày trước", "7 ngày vừa qua"]
    if any(k in p for k in _tuan_nay_kw) and not intent.get("date_range"):
        from datetime import date as _d, timedelta as _td
        _now = _d.today()
        _mon = _now - _td(days=_now.weekday())          # Thứ 2 tuần này
        _sun = _mon + _td(days=6)                        # Chủ nhật tuần này
        intent["date_range"] = {"from": str(_mon), "to": str(_sun)}
        intent["flags"]["is_current_week"] = True
    elif any(k in p for k in _tuan_truoc_kw) and not intent.get("date_range"):
        from datetime import date as _d, timedelta as _td
        _now = _d.today()
        _this_mon = _now - _td(days=_now.weekday())
        _prev_mon = _this_mon - _td(days=7)              # Thứ 2 tuần trước
        _prev_sun = _prev_mon + _td(days=6)              # Chủ nhật tuần trước
        intent["date_range"] = {"from": str(_prev_mon), "to": str(_prev_sun)}
        intent["flags"]["is_prev_week"] = True

    intent["periods"] = sorted(_months)
    if len(intent["periods"]) > 1 or any(k in p for k in ["các tháng", "từng tháng", "cả năm", "nhiều tháng"]):
        intent["flags"]["is_multi_period"] = True
        intent["query_type"] = "MULTI_PERIOD"
    elif len(intent["periods"]) == 1:
        pass  # 1 tháng đơn → SIMPLE

    # ----------------------------------------------------------------
    # 2. Máy + Xưởng — detect cặp "máy X xưởng Y" trước, rồi mới extract riêng
    # ----------------------------------------------------------------
    # Cặp ghép: "máy 1,2,3 của xưởng 1 với máy 2,3,4 của xưởng 2"
    # hoặc    : "máy 17 xưởng 1 và máy 23 xưởng 2"
    _raw_pairs = re.findall(r"máy\s*([\d,\s]+?)(?:\s+của\s+|\s+)xưởng\s*(\d+)", p)
    _pairs = []
    for _macs_str, _loc in _raw_pairs:
        for _m in re.findall(r"\d+", _macs_str):
            _pairs.append((_m, f"Weaving {_loc}"))

    if len(_pairs) >= 2:
        intent["machine_location_pairs"] = _pairs
        intent["machines"]  = [mac for mac, _ in _pairs]      # giữ trùng lặp (máy 1 ở 2 xưởng)
        intent["locations"] = list(dict.fromkeys(loc for _, loc in _pairs))
    else:
        # Không có cặp → extract riêng như cũ
        intent["machine_location_pairs"] = []
        machines_found = re.findall(r"máy\s*(\d+)", p)
        intent["machines"] = sorted(set(machines_found), key=lambda x: int(x))

        workshops_found = re.findall(r"xưởng\s*(\d+)", p)
        intent["locations"] = [f"Weaving {w}" for w in sorted(set(workshops_found), key=int)]
        # Khi câu hỏi đề cập xưởng cụ thể → luôn query Xưởng Dệt
        if workshops_found and not intent.get("clusters"):
            intent["clusters"] = ["Xưởng Dệt"]

    # ----------------------------------------------------------------
    # 3. Xưởng / khu vực (chỉ bổ sung nếu chưa có từ pairs)
    # ----------------------------------------------------------------
    if not intent.get("machine_location_pairs"):
        pass  # đã set ở trên

    # ----------------------------------------------------------------
    # 4. Cờ toàn xưởng
    # ----------------------------------------------------------------
    whole_kw = ["nguyên xưởng", "cả xưởng", "toàn xưởng", "tổng xưởng"]
    if any(k in p for k in whole_kw):
        intent["flags"]["is_whole_factory"] = True

    # Hỏi tên xưởng mà không đề cập máy cụ thể -> toàn xưởng
    if intent["locations"] and not intent["machines"]:
        intent["flags"]["is_whole_factory"] = True

    # ----------------------------------------------------------------
    # 4a. Màu sắc / thông tin cột đặc biệt
    # ----------------------------------------------------------------
    _color_kw = ["màu gì", "màu sắc", "màu của", "color", "màu nào",
                 "đang chạy màu", "chạy màu gì", "màu hàng"]
    if any(k in p for k in _color_kw):
        intent["flags"]["need_color"] = True

    # ----------------------------------------------------------------
    # 4a-0. SIZING_QUERY — câu hỏi về máy hồ/sectional/direct/winder
    # ----------------------------------------------------------------
    _sizing_machines = ["máy hồ", "máy sec", "máy qs", "máy direct", "máy sectional",
                        "winder", "suzuki", "bng", "benninger", "karlmayer", "honghwa",
                        "sizing", "hồ sợi", "máy hồ sợi"]
    _sizing_kw = ["hiệu suất sizing", "sản lượng sizing", "tốc độ máy hồ",
                  "thời gian chạy", "tốc độ thực tế", "tốc độ mục tiêu",
                  "mét máy qs", "mét máy sec", "mét máy hồ"]
    if (any(k in p for k in _sizing_machines) or any(k in p for k in _sizing_kw)):
        intent["query_type"] = "SIZING_QUERY"
        intent["flags"]["is_sizing"] = True
        # Detect machine type from question
        _sz_machine_map = {
            "máy hồ": "MÁY HỒ", "hồ sợi": "MÁY HỒ", "bng": "MÁY HỒ",
            "benninger": "MÁY HỒ", "karlmayer": "MÁY HỒ", "honghwa": "MÁY HỒ",
            "máy sec": "MÁY SEC", "sectional": "MÁY SEC",
            "máy qs": "MÁY QS", "máy direct": "MÁY QS", "direct": "MÁY QS",
            "winder": "WINDER", "suzuki": "SUZUKI",
        }
        for kw, mtype in _sz_machine_map.items():
            if kw in p:
                intent["sz_machine_type"] = mtype
                break

    # ----------------------------------------------------------------
    # 4a-1. BEAM_STATUS — beam trên máy còn bao nhiêu mét/kg
    # ----------------------------------------------------------------
    _beam_kw = [
        "beam còn", "beam trên máy", "beam máy", "còn bao nhiêu mét",
        "còn bao nhiêu yard", "beam còn lại", "beam hết chưa",
        "beam đang còn", "còn lại bao nhiêu", "hết beam chưa",
        "mét còn lại", "yard còn lại", "beam trên", "beam dưới",
        "thay beam", "beam mới", "lên beam",
    ]
    # ✅ FIX: Câu hỏi về beam lên máy gần nhất / mới nhất → BEAM_RECENT
    _beam_recent_kw = [
        "lên máy gần nhất", "lên máy mới nhất", "beam gần nhất",
        "beam mới nhất", "beam nào lên", "beam nào đang",
        "beam đang chạy", "beam hiện tại", "beam đang dùng",
        "lên máy ngày", "ngày lên máy", "lên máy lúc",
        "beam nào weaving", "beam nào xưởng",
    ]
    if any(k in p for k in _beam_recent_kw):
        intent["query_type"] = "BEAM_RECENT"
        intent["flags"]["is_beam_recent"] = True
    elif any(k in p for k in _beam_kw):
        intent["query_type"] = "BEAM_STATUS"
        intent["flags"]["is_beam_status"] = True

    # ----------------------------------------------------------------
    # 4a-2. ITEM_SCHEDULE — lịch chạy của 1 mã hàng (ngày nào + máy nào)
    # ----------------------------------------------------------------
    _sched_kw = [
        "có những ngày nào", "chạy những ngày nào", "ngày nào chạy",
        "ngày nào và máy nào", "máy nào và ngày nào", "lịch chạy",
        "đang chạy ngày nào", "chạy ngày mấy", "những ngày chạy",
        "ngày nào đang chạy", "hôm nào chạy", "khi nào chạy",
        "chạy vào ngày nào", "ngày nào có", "các ngày chạy",
    ]
    if any(k in p for k in _sched_kw):
        intent["query_type"] = "ITEM_SCHEDULE"
        intent["flags"]["is_item_schedule"] = True

    # ----------------------------------------------------------------
    # 4a-3. Phân loại PD / YD
    # ----------------------------------------------------------------
    _phan_loai = None
    # Detect "PD" / "YD" trong câu hỏi (không bị nhầm với mã hàng)
    _p_upper = p.upper()
    if re.search(r'(?<![A-Z0-9])PD(?![A-Z0-9])', _p_upper):
        _phan_loai = "PD"
    elif re.search(r'(?<![A-Z0-9])YD(?![A-Z0-9])', _p_upper):
        _phan_loai = "YD"
    if _phan_loai:
        intent["phan_loai"] = _phan_loai

    # ----------------------------------------------------------------
    # ----------------------------------------------------------------
    # 4a-3. Phân loại PD / YD
    # ----------------------------------------------------------------
    _phan_loai = None
    _p_upper = p.upper()
    _has_pd = bool(re.search(r'(?<![A-Z0-9])PD(?![A-Z0-9])', _p_upper))
    _has_yd = bool(re.search(r'(?<![A-Z0-9])YD(?![A-Z0-9])', _p_upper))
    if _has_pd and not _has_yd:
        _phan_loai = "PD"
    elif _has_yd and not _has_pd:
        _phan_loai = "YD"
    # ✅ FIX: Có cả YD và PD → so sánh 2 nhóm (không lọc riêng, query 2 lần)
    if _has_pd and _has_yd:
        intent["flags"]["is_pdyd_compare"] = True
        intent["query_type"] = "PDYD_COMPARE"
    elif _phan_loai:
        intent["phan_loai"] = _phan_loai

    # 4b. CATALOG — tổng hợp danh sách mã hàng đang chạy
    # ----------------------------------------------------------------
    _catalog_kw = [
        "tổng hợp mã", "danh sách mã", "mã nào đang", "các mã hàng",
        "mã sản phẩm", "mã hàng nào", "hàng nào đang chạy", "mã đang chạy",
        "sản phẩm đang chạy", "mặt hàng đang chạy", "hàng đang chạy",
        "tổng hợp các mã", "những mã hàng", "các sản phẩm đang",
        "danh sách hàng", "hàng gì đang",
        # "bao nhiêu mã hàng / sản phẩm"
        "bao nhiêu mã", "mấy mã", "bao nhiêu sản phẩm",
        "bao nhiêu loại hàng", "mấy loại hàng", "bao nhiêu loại mã",
        "chạy bao nhiêu", "chạy mấy", "số lượng mã",
    ]
    _ranking_kw = ["nhiều nhất", "ít nhất", "cao nhất", "thấp nhất",
                   "phổ biến nhất", "chạy nhiều", "sản lượng cao",
                   "chủ yếu", "chính", "đứng đầu", "top", "hàng đầu"]
    _has_catalog_kw = any(k in p for k in _catalog_kw)
    _has_ranking_kw = any(k in p for k in _ranking_kw)

    # "mã hàng nào chạy nhiều nhất?" → TOP_CATALOG (1 câu trả lời trực tiếp)
    # "mã hàng nào đang chạy?" → CATALOG (liệt kê đầy đủ)
    if _has_catalog_kw or _has_ranking_kw:
        intent["query_type"] = "CATALOG"
        intent["flags"]["is_catalog"] = True
        if _has_ranking_kw:
            intent["flags"]["is_top_catalog"] = True
            intent["flags"]["is_top_asc"] = any(k in p for k in ["ít nhất","thấp nhất","kém nhất","ít hơn"])
            # Số top N: "top 5 mã hàng" → n=5, mặc định 5
            import re as _re2
            _nm = _re2.search(r"top\s*(\d+)", p)
            intent["flags"]["top_n"] = int(_nm.group(1)) if _nm else 5

    # ----------------------------------------------------------------
    # 5. So sánh
    # ----------------------------------------------------------------
    compare_kw = ["so sánh", "vs ", "đối chiếu", "bên nào", "cái nào", "tốt hơn"]
    multi_entity = len(intent["machines"]) >= 2 or len(intent["locations"]) >= 2
    # "so sánh xưởng 1 và xưởng 2" = 2 locations → multi_entity = True
    if not multi_entity and len(workshops_found if "workshops_found" in dir() else []) >= 2:
        multi_entity = True
    has_compare_kw = any(k in p for k in compare_kw)
    # Chỉ set COMPARISON khi thực sự có 2+ máy hoặc 2+ xưởng để so sánh.
    # "so sánh tháng 1,2,3,4" → MULTI_PERIOD (so sánh các tháng, không phải máy/xưởng)
    if multi_entity and (has_compare_kw or any(k in p for k in ["hay", "và", "với"])):
        intent["flags"]["is_comparison"] = True
        intent["query_type"] = "COMPARISON"

    # ----------------------------------------------------------------
    # 6. Cross-source (kho beam + xưởng dệt)
    # ----------------------------------------------------------------
    beam_kw = ["kho beam", "beam yarn", "kho sợi beam", "p beam"]
    weave_kw = ["xưởng dệt", "xưởng", "weaving"]
    _has_beam_kw  = any(k in p for k in beam_kw)
    _has_weave_kw = any(k in p for k in weave_kw)
    # Từ khóa hỏi tồn kho / hiện tại
    _ton_kw = ["tổng kg", "tồn kho", "hiện tại", "còn trong kho", "kho còn",
               "tổng tồn", "tổng trong kho", "bao nhiêu kg", "bao nhiêu mét",
               "còn bao nhiêu", "trong kho beam", "trong kho"]
    _has_ton_kw = any(k in p for k in _ton_kw)

    if _has_beam_kw:
        if _has_weave_kw:
            # Có cả kho beam + xưởng dệt → cross-source thật sự
            intent["flags"]["is_cross_source"] = True
            intent["query_type"] = "CROSS_SOURCE"
            intent["clusters"] = ["xuong det", "xưởng dệt"]
        elif _has_ton_kw or "hiện tại" in p or "tồn" in p:
            # ✅ FIX: Hỏi tồn kho / hiện tại → dùng Beam_Info trực tiếp, không query Inventory_Log
            intent["flags"]["is_beam_warehouse_status"] = True
            intent["query_type"] = "BEAM_WAREHOUSE_STATUS"
        else:
            # Hỏi lịch sử nhập/xuất kho beam
            intent["flags"]["is_beam_warehouse"] = True
            intent["query_type"] = "CROSS_SOURCE"
            intent["clusters"] = ["beam", "kho beam"]

    # ----------------------------------------------------------------
    # 7. Khoảng ngày: 10/4 -> 16/4
    # ----------------------------------------------------------------
    rng = re.search(
        r"(\d{1,2})/(\d{1,2})\s*(?:->|-|đến|~|tới)\s*(?:ngày\s*)?(\d{1,2})/(\d{1,2})",
        p,
    )
    if rng:
        d1, m1, d2, m2 = rng.groups()
        yr = DEFAULT_YEAR
        intent["date_range"] = {
            "from": f"{yr}-{m1.zfill(2)}-{d1.zfill(2)}",
            "to":   f"{yr}-{m2.zfill(2)}-{d2.zfill(2)}",
    # ----------------------------------------------------------------
        }
    # 8. Ngày chính xác: 4/5 hoặc 04/03
    # ----------------------------------------------------------------
    # ✅ FIX: "tháng 3/2026" → date_range tháng 3 năm 2026 (TRƯỚC khi parse d/m)
    _thang_nam = re.search(r"tháng\s*(\d{1,2})[/\-](20\d{2})", p)
    if _thang_nam and not intent.get("date_range") and not intent.get("exact_date"):
        import calendar as _cal
        _mm = int(_thang_nam.group(1))
        _yy = int(_thang_nam.group(2))
        if 1 <= _mm <= 12:
            _last = _cal.monthrange(_yy, _mm)[1]
            intent["date_range"] = {

                "from": f"{_yy}-{_mm:02d}-01",
                "to":   f"{_yy}-{_mm:02d}-{_last:02d}",
            }
            if _mm not in intent["periods"]:
                intent["periods"] = sorted(set(intent["periods"]) | {_mm})

    if not intent["date_range"]:
        ex = re.search(r"(?<!\d)(\d{1,2})/(\d{1,2})(?![/\d])", p)
        if ex:
            d, m = ex.group(1), ex.group(2)
            # ✅ FIX: validate tháng hợp lệ (1-12) trước khi set exact_date
            if 1 <= int(m) <= 12:
                intent["exact_date"] = f"{DEFAULT_YEAR}-{m.zfill(2)}-{d.zfill(2)}"

    # ----------------------------------------------------------------
    # 8b. Năm cụ thể: "năm 2026" → date_range cả năm
    # ----------------------------------------------------------------
    yr_match = re.search(r"năm\s*(20\d{2})", p)
    if yr_match and not intent.get("exact_date") and not intent.get("date_range"):
        yr = yr_match.group(1)
        intent["date_range"] = {"from": f"{yr}-01-01", "to": f"{yr}-12-31"}
        intent["flags"]["is_full_year"] = True

    # ----------------------------------------------------------------
    # 8c. Quý: "quý 1/2/3/4"
    # ----------------------------------------------------------------
    q_match = re.search(r"qu[yý]\s*([1-4])", p)
    if q_match and not intent.get("exact_date") and not intent.get("date_range"):
        q = int(q_match.group(1))
        qm = {1: ("01","03"), 2: ("04","06"), 3: ("07","09"), 4: ("10","12")}[q]
        yr = DEFAULT_YEAR
        # Lấy năm nếu có
        yr_m2 = re.search(r"20\d{2}", p)
        if yr_m2: yr = yr_m2.group(0)
        intent["date_range"] = {"from": f"{yr}-{qm[0]}-01", "to": f"{yr}-{qm[1]}-30"}
        intent["flags"]["quarter"] = q

    # ----------------------------------------------------------------
    # 9. Tìm mã hàng - 3 patterns kết hợp để bắt các cách nói khác nhau
    # ----------------------------------------------------------------
    _STOP = (r'(?=\s+(?:có|nào|này|chạy|đang|trong|của|ở|tại|là|và|với|hay|'
             r'hoặc|không|ko|nhé|sao|thì|đó|tháng|ngày|năm|xưởng|máy|kho|'
             r'bên|cái|loại|bao|nhiêu|ra|nữa|thế|đây|kia|hết|xong|rồi)|\s*$)')
    # Từ câu hỏi / filler — nếu result chứa bất kỳ từ này thì bỏ qua (không phải mã hàng)
    _Q_WORDS = {"gì", "nào", "đó", "không", "ko", "bao", "nhiêu", "thế", "hàng", "mã"}
    _NO_Q    = r'(?!(?:gì|nào|đó|không|ko|bao|nhiêu)\s)'  # block sau keyword
    # Mở rộng STOP để không nuốt theo từ sau mã hàng
    _STOP = (r'(?=\s+(?:có|nào|này|chạy|đang|trong|của|ở|tại|là|và|với|hay|'
             r'hoặc|không|ko|nhé|sao|thì|đó|tháng|ngày|năm|xưởng|máy|kho|'
             r'bên|cái|loại|bao|nhiêu|ra|nữa|thế|đây|kia|hết|xong|rồi|'
             r'sản|lượng|hiệu|suất|so|sánh)|}|\s*$)')
    _ITEM_PATTERNS = [
        # P1: sau "mã hàng / mã / code / item" — tối đa 80 ký tự (đủ cho tên dài)
        r'(?:mã\s+(?:hàng\s+)?|code\s+|item\s+)' + _NO_Q + r'([\w\s/()\.\-]{2,80}?)' + _STOP,
        # P2: sau "chạy hàng X"
        r'(?:chạy\s+hàng\s+|hàng\s+(?:này\s+là\s+)?)' + _NO_Q + r'([\w\s/()\.\-]{2,80}?)' + _STOP,
        # P3: mã hàng dạng "CS 32", "NV100", "PD-300" xuất hiện tự do
        r'(?<!\w)([a-z]{1,5}[\s-]?\d{1,4})(?=\s+(?:có|nào|trong|tháng|ngày|xưởng|máy)|\s*$)',
    ]
    _item_found = None
    for _pat in _ITEM_PATTERNS:
        _m = re.search(_pat, p)
        if _m:
            _candidate = _m.group(1).strip()
            # Bỏ qua nếu bất kỳ token nào là từ hỏi / filler
            if not (set(_candidate.lower().split()) & _Q_WORDS):
                _item_found = _candidate
                break
    if _item_found:
        intent["item_kw"] = _item_found

    # ----------------------------------------------------------------
    # 10. Tìm máy nào
    # ----------------------------------------------------------------
    find_mac_kw = ["máy nào", "những máy nào", "còn máy nào", "máy gì", "máy khác"]
    # Không set is_find_machine nếu đã có machine_location_pairs (người dùng đang so sánh máy cụ thể)
    # VD: "máy 1 xưởng 2 và máy 1 xưởng 3 máy nào hiệu quả hơn" → COMPARISON, không phải find
    _has_pairs = len(intent.get("machine_location_pairs", [])) >= 2
    _has_specific_machines = len(intent.get("machines", [])) >= 2 or (
        len(intent.get("machines", [])) >= 1 and len(intent.get("locations", [])) >= 1
    )
    if any(k in p for k in find_mac_kw) and not _has_pairs and not _has_specific_machines:
        intent["flags"]["is_find_machine"] = True

    # ----------------------------------------------------------------
    # 11. Trend N ngày liên tiếp / gần đây
    #     VD: "10 ngày liên tiếp", "7 ngày gần nhất", "5 ngày qua"
    #     Kết hợp với exact_date → query khoảng [exact_date - N+1 .. exact_date]
    # ----------------------------------------------------------------
    trend_patterns = [
        r"(\d+)\s*ngày\s+(?:liên tiếp|liên tục|gần nhất|gần đây|qua|vừa qua|trước)",
        r"(?:xem|so sánh|kiểm tra|nhìn lại)\s+(\d+)\s*ngày",
    ]
    trend_m = None
    for pat in trend_patterns:
        trend_m = re.search(pat, p)
        if trend_m:
            break
    # Bắt thêm: "N ngày" khi đã có ngày cụ thể + từ khoá hiệu suất/xu hướng
    if not trend_m and intent.get("exact_date"):
        eff_kw = ["hiệu quả", "hiệu suất", "xu hướng", "trend", "chạy tốt", "ổn định", "so sánh"]
        if any(k in p for k in eff_kw):
            trend_m = re.search(r"(\d+)\s*ngày", p)

    if trend_m:
        n = int(trend_m.group(1))
        if 2 <= n <= 90:
            intent["n_days_trend"] = n
            intent["query_type"]   = "TREND"
            intent["flags"]["is_trend"] = True

    return intent




# =====================================================================
# MODULE: SUGGESTION GENERATOR
# =====================================================================
def suggest_questions(intent: dict, schema: dict) -> str:
    """
    Tạo danh sách câu hỏi gợi ý khi query thất bại hoặc không ra dữ liệu.
    Dựa trên intent đã detect được để gợi ý câu hỏi đúng cú pháp.
    """
    qt      = intent.get("query_type", "SIMPLE")
    macs    = intent.get("machines", [])
    locs    = intent.get("locations", [])
    periods = intent.get("periods", [])
    item_kw = intent.get("item_kw")
    clusters = schema.get("clusters", [])
    locations_db = schema.get("locations", [])
    date_min = schema.get("date_range", {}).get("min", "")
    date_max = schema.get("date_range", {}).get("max", "")

    lines = ["\n💡 **Gợi ý câu hỏi đúng cú pháp:**\n"]

    weaving_locs = [l for l in locations_db if "Weaving" in l]
    sample_locs  = weaving_locs[:3] if weaving_locs else ["Weaving 1", "Weaving 2", "Weaving 3"]

    # Xác định tháng mẫu từ dữ liệu thực
    sample_month = "1"
    if date_min:
        try: sample_month = str(int(date_min[5:7]))
        except: pass

    # --- Gợi ý theo loại intent ---
    if qt in ("COMPARISON",) or (len(macs) >= 1 and len(locs) >= 2):
        # Gợi ý so sánh máy giữa 2 xưởng
        mac = macs[0] if macs else "27"
        l0 = sample_locs[0].replace("Weaving ", "") if sample_locs else "1"
        l1 = sample_locs[1].replace("Weaving ", "") if len(sample_locs) > 1 else "2"
        mth = periods[0] if periods else sample_month
        lines.append(f'- **So sánh máy + tháng:** "so sánh máy {mac} xưởng {l0} và xưởng {l1} tháng {mth}"')
        lines.append(f'- **Tất cả tháng:** "so sánh máy {mac} xưởng {l0} và xưởng {l1}"')
        lines.append(f'- **Theo hiệu suất:** "hiệu suất máy {mac} xưởng {l0} và xưởng {l1} tháng {mth}"')

    if qt in ("MULTI_PERIOD",) or len(periods) > 1:
        item = item_kw or "CS 32"
        mths = ", ".join(str(m) for m in periods[:4]) if periods else "1, 2, 3, 4"
        lines.append(f'- **Sản lượng theo tháng:** "mã hàng {item} tháng {mths} sản lượng bao nhiêu"')
        if macs:
            lines.append(f'- **Máy theo tháng:** "máy {macs[0]} tháng {mths} sản lượng bao nhiêu"')

    if item_kw or qt == "SIMPLE":
        item = item_kw or "CS 32"
        loc  = locs[0] if locs else sample_locs[0] if sample_locs else "Weaving 1"
        loc_n = loc.replace("Weaving ", "")
        mac  = macs[0] if macs else "1"
        lines.append(f'- **Mã hàng + xưởng:** "mã hàng {item} xưởng {loc_n} tháng {sample_month} sản lượng bao nhiêu"')
        if not macs:
            lines.append(f'- **Tất cả máy xưởng:** "xưởng {loc_n} tháng {sample_month} sản lượng bao nhiêu"')

    # Luôn thêm ví dụ catalog vào suggestion
    lines.append(f'- **Danh sách mã hàng:** "mã sản phẩm đang chạy xưởng 1"')
    lines.append(f'- **Danh sách mã hàng:** "tổng hợp mã hàng xưởng 1 tháng {sample_month}"')

    if len(lines) <= 3:
        # Fallback generic
        lines.append(f'- "mã hàng CS 32 tháng {sample_month} sản lượng bao nhiêu"')
        lines.append(f'- "so sánh máy 17 xưởng 1 và máy 23 xưởng 2 tháng {sample_month}"')
        lines.append(f'- "máy nào xưởng 1 tháng {sample_month} hiệu suất cao nhất"')

    lines.append("\n> *(Lưu ý: tên xưởng dùng số: xưởng 1, xưởng 2, xưởng 3 | tên máy: máy 1..50 | tháng: tháng 1..12)*")
    return "\n".join(lines)

# =====================================================================
# MODULE 3: SQL BUILDER (thuần Python, không cần AI)
# =====================================================================

def _escape_col(col: str) -> str:
    """Bọc tên cột trong dấu ngoặc vuông."""
    return f'[{col}]'


def _mac_in_clause(machines: list) -> str:
    """Tạo điều kiện IN cho tên máy (handle cả '1' và '1.0')."""
    vals = []
    for m in machines:
        vals.append(f"'{m}'")
        if "." not in m:
            vals.append(f"'{m}.0'")
    return ", ".join(vals)


def _month_conditions(periods: list) -> str:
    """Tạo điều kiện OR cho danh sách tháng."""
    parts = [f"[date] LIKE '%-{str(m).zfill(2)}-%'" for m in periods]
    return f"({' OR '.join(parts)})"


def _build_where(intent: dict, schema: dict, machine_filter: str = None) -> str:
    """
    Xây dựng mệnh đề WHERE từ intent.
    machine_filter: nếu truyền vào thì override intent['machines']\n"""
    parts = []
    mac = machine_filter or None
    macs = [mac] if mac else intent.get("machines", [])
    locs = intent.get("locations", [])
    periods = intent.get("periods", [])
    date_range = intent.get("date_range")
    exact_date = intent.get("exact_date")
    item_kw = intent.get("item_kw")
    flags = intent.get("flags", {})
    schema_mac = schema.get("mac_col", "ten_may")
    schema_item = schema.get("item_col", "item_id")

    # Xưởng / sub_location
    if locs and not flags.get("is_whole_factory"):
        loc_list = ", ".join([f"'{l}'" for l in locs])
        parts.append(f"[sub_location] IN ({loc_list})")
    elif locs and flags.get("is_whole_factory"):
        # Lọc xưởng nhưng lấy tất cả máy
        loc_list = ", ".join([f"'{l}'" for l in locs])
        parts.append(f"[sub_location] IN ({loc_list})")

    # Máy
    if macs and not flags.get("is_whole_factory") and not flags.get("is_find_machine"):
        parts.append(f"[{schema_mac}] IN ({_mac_in_clause(macs)})")

    # Thời gian
    if date_range:
        parts.append(f"[date] BETWEEN '{date_range['from']}' AND '{date_range['to']}'")
    elif exact_date:
        parts.append(f"[date] = '{exact_date}'")
    elif periods:
        parts.append(_month_conditions(periods))

    # Mã hàng - UPPER() để case-insensitive (CS 32 = cs 32 = Cs 32)
    if item_kw:
        safe_kw = item_kw.strip().replace("'", "''").upper()
        # Ưu tiên exact match (=) — chỉ dùng LIKE khi mã hàng có thể là substring
        # Exact match ngăn "SW NEW COLOR MUJI 40" khớp với "ANTI NEW COLOR MUJI 40"
        parts.append(f"UPPER([{schema_item}]) = '{safe_kw}' OR UPPER([{schema_item}]) LIKE UPPER('{safe_kw} %') OR UPPER([{schema_item}]) LIKE UPPER('% {safe_kw}')")

    # Cluster
    clusters = intent.get("clusters", [])
    if clusters:
        cl_list = ", ".join([f"'{c}'" for c in clusters])
        parts.append(f"[cluster_name] IN ({cl_list})")

    # Phân loại PD / YD
    # ✅ FIX: Tên cột trong DB là 'phan_loai' (không dấu, từ weaving_pipeline.py)
    phan_loai = intent.get("phan_loai")
    if phan_loai:
        parts.append(
            f"(UPPER(TRIM([phan_loai])) = '{phan_loai}' "
            f"OR UPPER(TRIM([phan_loai])) LIKE '{phan_loai}%')"
        )

    return " AND ".join(parts) if parts else "1=1"



def _eff_expr(col: str) -> str:
    """SQL expression: đọc giá trị hiệu suất từ cột, trả về NULL nếu ≤ 0."""
    return f"NULLIF(CAST(REPLACE(REPLACE([{col}],'%',''),' ','') AS REAL), 0)"


def _make_eff_sql(schema: dict) -> str | None:
    """
    Tạo SQL AVG expression tính hiệu suất theo thứ tự ưu tiên cho từng dòng:
      1. Hiệu suất 2 ca (nếu > 0)
      2. (ca A + ca B) / 2 (nếu cả 2 đều > 0)
      3. Ca A hoặc Ca B (lấy cái có số liệu)
    Trả về None nếu không có cột hiệu suất nào trong schema.
    """
    e2 = schema.get("eff_2ca_col")
    ea = schema.get("eff_a_col")
    eb = schema.get("eff_b_col")

    if not e2 and not ea and not eb:
        return None

    parts = []
    if e2:
        parts.append(f"WHEN {_eff_expr(e2)} IS NOT NULL AND {_eff_expr(e2)} > 0 THEN {_eff_expr(e2)}")
    if ea and eb:
        parts.append(
            f"WHEN {_eff_expr(ea)} IS NOT NULL AND {_eff_expr(ea)} > 0 "
            f"AND {_eff_expr(eb)} IS NOT NULL AND {_eff_expr(eb)} > 0 "
            f"THEN ({_eff_expr(ea)} + {_eff_expr(eb)}) / 2"
        )
    if ea:
        parts.append(f"WHEN {_eff_expr(ea)} IS NOT NULL AND {_eff_expr(ea)} > 0 THEN {_eff_expr(ea)}")
    if eb:
        parts.append(f"WHEN {_eff_expr(eb)} IS NOT NULL AND {_eff_expr(eb)} > 0 THEN {_eff_expr(eb)}")

    case_expr = "CASE " + " ".join(parts) + " ELSE NULL END"
    return f"ROUND(AVG({case_expr}), 2) as avg_eff"


def build_queries(intent: dict, schema: dict, user_prompt_lower: str) -> list:
    """
    Trả về list[(label, sql)] dựa trên loại intent.
    Có thể trả về nhiều query để so sánh.
    """
    qt = intent.get("query_type", "SIMPLE")
    schema_qty  = schema.get("qty_col")
    schema_mac  = schema.get("mac_col", "ten_may")
    schema_item = schema.get("item_col", "item_id")
    schema_eff  = schema.get("eff_2ca_col") or schema.get("eff_a_col")

    # Cột select detail (raw)
    def detail_select(include_color=False):
        cols = ["[date]", "[sub_location]", f"[{schema_mac}]", f"[{schema_item}]"]
        # Thêm cột COLOR nếu cần hoặc user hỏi về màu
        _color = schema.get("color_col")
        if _color and (include_color or intent.get("flags", {}).get("need_color")):
            cols.append(f"[{_color}]")
        if schema_qty:
            cols.append(f"[{schema_qty}]")
        if schema_eff:
            cols.append(f"[{schema_eff}]")
        # Thêm cột eff_a, eff_b nếu có
        if schema.get("eff_a_col") and schema.get("eff_a_col") != schema_eff:
            cols.append(f"[{schema.get('eff_a_col')}]")
        if schema.get("eff_b_col"):
            cols.append(f"[{schema.get('eff_b_col')}]")
        return ", ".join(cols)

    # Cột select aggregated
    def agg_select(group_by_month=False):
        cols = [f"[sub_location]", f"[{schema_mac}]"]
        if group_by_month:
            cols.append("strftime('%Y-%m', [date]) as month")
        else:
            cols.append("[date]")
        if schema_qty:
            cols.append(
                f"ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)), 2) as total_kg"
            )
        _eff = _make_eff_sql(schema)
        if _eff:
            cols.append(_eff)
        return ", ".join(cols)

    # ==============================================================
    # EXPLAIN → trả về rỗng, xử lý ở tầng trên
    # ==============================================================
    if qt == "EXPLAIN":
        return []

    # ==============================================================
    # TREND: N ngày liên tiếp kết thúc tại anchor_date
    # Nếu câu hỏi có cả ngày cụ thể (VD: 3/3) lẫn "10 ngày liên tiếp",
    # trả về 2 query:
    #   [0] Chi tiết đúng ngày đó (mã hàng, HS ngay ngày 3/3)
    #   [1] Bảng xu hướng 10 ngày liên tiếp kết thúc tại 3/3
    # ==============================================================
    if qt == "TREND":
        from datetime import datetime, timedelta

        n = intent.get("n_days_trend", 10)
        anchor_str = intent.get("exact_date")
        if not anchor_str:
            # ✅ FIX: Tự động dùng ngày có data mới nhất trong DB thay vì today
            # → Tránh tình huống today=29/5 nhưng data chỉ đến 22/5 → 7 ngày trống
            try:
                import sqlite3 as _sqt
                _where_anchor = _build_where(
                    dict(intent, exact_date=None, date_range=None,
                         flags=dict(intent.get("flags",{}), is_trend=False)),
                    schema
                )
                _conn_a = _sqt.connect("inventory.db")
                _row_a = _conn_a.execute(
                    f"SELECT MAX([date]) FROM Inventory_Log WHERE {_where_anchor}"
                ).fetchone()
                _conn_a.close()
                _max_date = (_row_a[0] or "")[:10] if _row_a else ""
                anchor_str = _max_date if _max_date else datetime.today().strftime("%Y-%m-%d")
            except Exception:
                anchor_str = datetime.today().strftime("%Y-%m-%d")

        try:
            anchor_dt = datetime.strptime(anchor_str[:10], "%Y-%m-%d")
        except ValueError:
            anchor_dt = datetime.today()

        start_dt  = anchor_dt - timedelta(days=n - 1)
        start_str = start_dt.strftime("%Y-%m-%d")

        detail_cols = ["[date]", "[sub_location]", f"[{schema_mac}]", f"[{schema_item}]"]
        if schema_qty:
            detail_cols.append(f"[{schema_qty}]")
        if schema_eff:
            detail_cols.append(f"[{schema_eff}]")
        if schema.get("eff_a_col") and schema.get("eff_a_col") != schema_eff:
            detail_cols.append(f"[{schema.get('eff_a_col')}]")
        if schema.get("eff_b_col"):
            detail_cols.append(f"[{schema.get('eff_b_col')}]")

        queries_out = []

        # --- Query 1: Đúng ngày anchor ---
        if intent.get("exact_date"):
            day_intent = dict(
                intent,
                query_type = "SIMPLE",
                date_range = None,
                flags      = dict(intent.get("flags", {}), is_trend=False),
            )
            where_day = _build_where(day_intent, schema)
            sql_day = (
                f"SELECT {', '.join(detail_cols)} FROM Inventory_Log "
                f"WHERE {where_day} "
                f"ORDER BY [sub_location], [{schema_mac}]"
            )
            queries_out.append((f"Ngay_{anchor_str}", sql_day))

        # --- Query 2: Bảng xu hướng N ngày ---
        trend_intent = dict(
            intent,
            query_type  = "SIMPLE",
            exact_date  = None,
            date_range  = {"from": start_str, "to": anchor_str},
            flags       = dict(intent.get("flags", {}), is_trend=True),
        )
        where_trend = _build_where(trend_intent, schema)
        sql_trend = (
            f"SELECT {', '.join(detail_cols)} FROM Inventory_Log "
            f"WHERE {where_trend} "
            f"ORDER BY [date] ASC, [sub_location], [{schema_mac}]"
        )
        queries_out.append((f"Trend_{n}ngay_{start_str}_{anchor_str}", sql_trend))

        return queries_out

    # ==============================================================
    # COMPARISON: chạy query riêng cho từng máy / từng xưởng
    # ==============================================================
    if qt == "COMPARISON":
        queries = []

        # Nếu không có 2 máy / 2 xưởng thực sự → fallback về MULTI_PERIOD
        if len(intent["machines"]) < 2 and len(intent["locations"]) < 2:
            intent = dict(intent, query_type="MULTI_PERIOD",
                          flags=dict(intent.get("flags", {}), is_comparison=False))
            return build_queries(intent, schema, user_prompt_lower)

        if len(intent["machines"]) >= 2:
            # Ưu tiên dùng machine_location_pairs nếu có (VD: "máy 17 xưởng 1 và máy 23 xưởng 2")
            pairs = intent.get("machine_location_pairs", [])
            if pairs:
                iter_list = pairs  # [(mac, loc), ...]
            else:
                # Không có pairs → dùng tất cả machines với location đầu tiên (nếu có)
                default_loc = intent["locations"][0] if intent["locations"] else None
                iter_list = [(mac, default_loc) for mac in intent["machines"]]

            for mac, loc in iter_list:
                single = dict(
                    intent,
                    machines=[mac],
                    locations=[loc] if loc else [],
                    flags=dict(intent["flags"], is_comparison=False, is_whole_factory=False),
                )
                where = _build_where(single, schema)
                loc_n = loc.replace("Weaving ", "Xưởng ") if loc else ""
                label = f"{loc_n} — Máy {mac}" if loc_n else f"Máy {mac}"

                group_month = bool(intent["periods"] and len(intent["periods"]) > 1)
                sel = agg_select(group_by_month=group_month)
                grp_cols = f"[sub_location], [{schema_mac}]"
                if group_month:
                    grp_cols += ", strftime('%Y-%m', [date])"
                order = "month" if group_month else "[sub_location]"
                sql = (
                    f"SELECT {sel} FROM Inventory_Log "
                    f"WHERE {where} "
                    f"GROUP BY {grp_cols} "
                    f"ORDER BY {order}"
                )
                queries.append((label, sql))


        elif len(intent["locations"]) >= 2:
            # So sánh theo xưởng — GIỮ NGUYÊN bộ lọc máy nếu có
            # VD: "máy 27 xưởng 1 vs xưởng 2" → lọc ten_may=27 ở cả 2 xưởng
            has_machine_filter = bool(intent.get("machines"))
            for loc in intent["locations"]:
                single = dict(
                    intent,
                    locations=[loc],
                    machines=intent.get("machines", []),  # Giữ máy, không reset
                    flags=dict(
                        intent["flags"],
                        is_comparison=False,
                        is_whole_factory=(not has_machine_filter),
                    ),
                )
                where = _build_where(single, schema)
                if has_machine_filter:
                    mac_label = ", ".join(f"Máy {m}" for m in intent["machines"])
                    label = f"{loc} — {mac_label}"
                else:
                    label = loc

                group_month = bool(intent["periods"] and len(intent["periods"]) > 1)
                # Khi lọc theo máy cụ thể: GROUP BY phải bao gồm ten_may
                # để tránh SQLite gộp nhầm toàn xưởng
                if has_machine_filter:
                    grp_cols = f"[sub_location], [{schema_mac}]"
                    if group_month:
                        grp_cols += ", strftime('%Y-%m', [date])"
                    # SELECT: dùng month hoặc aggregate rõ ràng, bỏ [date] raw
                    if group_month:
                        sel = agg_select(group_by_month=True)
                    else:
                        sel = (
                            f"[sub_location], [{schema_mac}]"
                        )
                        if schema_qty:
                            sel += f", ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)), 2) as total_kg"
                        _eff = _make_eff_sql(schema)
                        if _eff:
                            sel += f", {_eff}"
                else:
                    grp_cols = f"[sub_location]"
                    if group_month:
                        grp_cols += ", strftime('%Y-%m', [date])"
                    sel = agg_select(group_by_month=group_month)

                sql = (
                    f"SELECT {sel} FROM Inventory_Log "
                    f"WHERE {where} "
                    f"GROUP BY {grp_cols} "
                    f"ORDER BY {('[sub_location]' if not group_month else 'month')}"
                )
                queries.append((label, sql))

        return queries if queries else _single_query(intent, schema, detail_select, agg_select)



    # ==============================================================
    # ITEM_SCHEDULE: lịch chạy của 1 mã hàng — ngày nào + máy nào
    # ==============================================================
    if qt == "ITEM_SCHEDULE" or intent.get("flags", {}).get("is_item_schedule"):
        schema_qty  = schema.get("qty_col", "quantity_kg")
        schema_item = schema.get("item_col", "item_id")
        mac_norm    = f"CAST(CAST([{schema_mac}] AS INTEGER) AS TEXT)"

        where_parts = []
        # Lọc mã hàng — bắt buộc phải có
        _sched_item = intent.get("item_kw", "")
        if not _sched_item:
            # Không có mã hàng → không thể chạy ITEM_SCHEDULE
            return [("_ERROR_NO_ITEM", "SELECT 'Vui lòng chỉ định mã hàng, VD: CS 32 chạy những ngày nào' as msg")]
        _si = _sched_item.strip().replace("'", "''").upper()
        where_parts.append(
            f"(UPPER([{schema_item}]) = '{_si}' "
            f"OR UPPER([{schema_item}]) LIKE UPPER('{_si} %') "
            f"OR UPPER([{schema_item}]) LIKE UPPER('% {_si}'))"
        )
        # Lọc xưởng nếu có
        locs = intent.get("locations", [])
        if locs:
            loc_vals = ", ".join(f"'{l}'" for l in locs)
            where_parts.append(f"[sub_location] IN ({loc_vals})")
        # Lọc thời gian nếu có
        _dr = intent.get("date_range")
        _ex = intent.get("exact_date")
        _pr = intent.get("periods", [])
        if _dr:
            where_parts.append(f"[date] BETWEEN '{_dr['from']}' AND '{_dr['to']}'")
        elif _pr:
            mc = " OR ".join(f"[date] LIKE '%-{str(m).zfill(2)}-%'" for m in _pr)
            where_parts.append(f"({mc})")

        where_parts.append(f"[{schema_item}] IS NOT NULL")
        where_str = " AND ".join(where_parts)

        _eff = _make_eff_sql(schema)
        eff_sel = f", {_eff}" if _eff else ""
        sched_sql = (
            f"SELECT [date], [sub_location], {mac_norm} as ten_may_norm, "
            f"ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)),2) as total_kg"
            f"{eff_sel} "
            f"FROM Inventory_Log "
            f"WHERE {where_str} "
            f"GROUP BY [date], [sub_location], {mac_norm} "
            f"HAVING total_kg > 0 "
            f"ORDER BY [date], [sub_location], CAST({mac_norm} AS INTEGER)"
        )
        return [("Lịch chạy", sched_sql)]

    # ==============================================================
    # CATALOG: liệt kê mã hàng đang chạy
    # ==============================================================
    if qt == "CATALOG":
        where_parts = []
        locs = intent.get("locations", [])
        periods = intent.get("periods", [])
        schema_item = schema.get("item_col", "item_id")
        schema_qty  = schema.get("qty_col", "quantity_kg")

        if locs:
            # Người dùng chỉ định xưởng cụ thể → dùng đúng đó
            loc_vals = ", ".join(f"'{l}'" for l in locs)
            where_parts.append(f"[sub_location] IN ({loc_vals})")
        else:
            # Không chỉ định → mặc định chỉ lấy Weaving (lọc bỏ Stock)
            where_parts.append("[sub_location] LIKE 'Weaving%'")

        # Bộ lọc thời gian theo thứ tự ưu tiên: date_range > exact_date > periods
        _cat_date_range = intent.get("date_range")
        _cat_exact_date = intent.get("exact_date")
        if _cat_date_range:
            where_parts.append(
                f"[date] BETWEEN '{_cat_date_range['from']}' AND '{_cat_date_range['to']}'"
            )
        elif _cat_exact_date:
            where_parts.append(f"[date] = '{_cat_exact_date}'")
        elif periods:
            month_conds = " OR ".join(f"[date] LIKE '%-{str(m).zfill(2)}-%'" for m in periods)
            where_parts.append(f"({month_conds})")
        # Không có bộ lọc thời gian → lấy toàn bộ (đã có Weaving filter)

        # Lọc bỏ item_id rỗng / số 0
        where_parts.append(f"[{schema_item}] IS NOT NULL AND [{schema_item}] != '' AND [{schema_item}] != '0'")

        where_str = " AND ".join(where_parts) if where_parts else "1=1"
        catalog_sql = (
            f"SELECT [sub_location], [{schema_item}] as item_id, "
            f"ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)),2) as total_kg, "
            f"COUNT(DISTINCT [{schema.get('machine_col','ten_may')}]) as n_machines "
            f"FROM Inventory_Log "
            f"WHERE {where_str} "
            f"GROUP BY [sub_location], [{schema_item}] "
            f"ORDER BY [sub_location], total_kg DESC"
        )
        return [("Danh sách mã hàng", catalog_sql)]

    # ==============================================================
    # MULTI_PERIOD: tổng hợp theo tháng
    # ==============================================================
    if qt == "MULTI_PERIOD":
        where = _build_where(intent, schema)
        group_cols = f"[sub_location], [{schema_mac}], strftime('%Y-%m', [date]) as month"
        sel = (
            f"[sub_location], [{schema_mac}], "
            f"strftime('%Y-%m', [date]) as month"
        )
        if schema_qty:
            sel += (
                f", ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)), 2) as total_kg"
            )
        _eff = _make_eff_sql(schema)
        if _eff:
            sel += f", {_eff}"
        sql = (
            f"SELECT {sel} FROM Inventory_Log "
            f"WHERE {where} "
            f"GROUP BY [sub_location], [{schema_mac}], strftime('%Y-%m', [date]) "
            f"ORDER BY month, [sub_location], [{schema_mac}]"
        )
        return [("Tổng hợp đa tháng", sql)]

    # ==============================================================
    # CROSS_SOURCE: query cho từng cluster rồi tổng hợp
    # ==============================================================
    if qt == "CROSS_SOURCE":
        queries = []
        all_clusters = schema.get("clusters", [])

        # Tìm cluster tên giống xưởng dệt
        weave_clusters = [
            c for c in all_clusters
            if any(k in c.lower() for k in ["xưởng", "det", "weav", "dệt"])
        ]
        beam_clusters = [
            c for c in all_clusters
            if any(k in c.lower() for k in ["beam", "sợi", "yarn", "kho"])
        ]

        # ✅ FIX: Nếu user chỉ hỏi kho beam (is_beam_warehouse) → bỏ qua Xưởng Dệt
        _only_beam = intent.get("flags", {}).get("is_beam_warehouse", False)
        _pairs = []
        if not _only_beam:
            _pairs.append((weave_clusters, "Xưởng Dệt"))
        _pairs.append((beam_clusters, "Kho Beam/Sợi"))

        for clusters_group, label in _pairs:
            if not clusters_group:
                continue
            tmp_intent = dict(intent, clusters=clusters_group,
                              flags=dict(intent["flags"], is_cross_source=False))
            where = _build_where(tmp_intent, schema)

            grp = f"[sub_location], [{schema_mac}], strftime('%Y-%m', [date]) as month"
            sel = f"[sub_location], [{schema_mac}], strftime('%Y-%m', [date]) as month, [cluster_name]"
            if schema_qty:
                sel += (
                    f", ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)), 2) as total_kg"
                )
            sql = (
                f"SELECT {sel} FROM Inventory_Log "
                f"WHERE {where} "
                f"GROUP BY [sub_location], [{schema_mac}], strftime('%Y-%m', [date]) "
                f"ORDER BY month"
            )
            queries.append((label, sql))

        return queries if queries else _single_query(intent, schema, detail_select, agg_select)

    # ==============================================================
    # SIMPLE: query chi tiết (default)
    # ==============================================================
    return _single_query(intent, schema, detail_select, agg_select)


def _single_query(intent, schema, detail_select_fn, agg_select_fn):
    """Xây query đơn cho các trường hợp SIMPLE."""
    schema_mac  = schema.get("mac_col", "ten_may")
    where = _build_where(intent, schema)

    flags = intent.get("flags", {})
    periods = intent.get("periods", [])
    exact_date = intent.get("exact_date")

    if flags.get("is_find_machine"):
        # Tìm máy nào -> GROUP BY xưởng + máy (chuẩn hoá tên máy: "46.0" -> "46")
        schema_qty = schema.get("qty_col")
        schema_eff = schema.get("eff_2ca_col") or schema.get("eff_a_col")
        # mac_norm: loại bỏ phần ".0" cuối để gom chung "46" và "46.0"
        mac_norm = f"CAST(CAST([{schema_mac}] AS INTEGER) AS TEXT)"
        sel = f"[sub_location], {mac_norm} as ten_may_norm"
        if schema_qty:
            sel += f", ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)), 2) as total_kg"
        _eff = _make_eff_sql(schema)
        if _eff:
            sel += f", {_eff}"
        sql = (
            f"SELECT {sel} FROM Inventory_Log "
            f"WHERE {where} "
            f"GROUP BY [sub_location], {mac_norm} "
            f"HAVING total_kg > 0 "
            f"ORDER BY [sub_location], total_kg DESC"
        )
        return [("MachineList", sql)]

    elif len(periods) == 1:
        # 1 tháng duy nhất → chi tiết từng ngày
        sel = detail_select_fn()
        sql = (
            f"SELECT {sel} FROM Inventory_Log "
            f"WHERE {where} "
            f"ORDER BY [date], [sub_location], [{schema_mac}]"
        )
        return [("Chi tiết tháng", sql)]

    elif exact_date:
        # Ngày chính xác → chi tiết
        sel = detail_select_fn()
        sql = (
            f"SELECT {sel} FROM Inventory_Log "
            f"WHERE {where} "
            f"ORDER BY [sub_location], [{schema_mac}]"
        )
        return [("Chi tiết ngày", sql)]

    elif intent.get("date_range"):
        # Khoảng ngày → chi tiết
        sel = detail_select_fn()
        sql = (
            f"SELECT {sel} FROM Inventory_Log "
            f"WHERE {where} "
            f"ORDER BY [date], [sub_location], [{schema_mac}]"
        )
        return [("Khoảng ngày", sql)]

    else:
        # Không có lọc thời gian → GROUP BY tháng để tổng hợp
        schema_qty = schema.get("qty_col")
        schema_eff = schema.get("eff_2ca_col") or schema.get("eff_a_col")
        sel = f"[sub_location], [{schema_mac}], strftime('%Y-%m', [date]) as month"
        if schema_qty:
            sel += f", ROUND(SUM(CAST(REPLACE(REPLACE([{schema_qty}],',',''),' ','') AS REAL)), 2) as total_kg"
        _eff = _make_eff_sql(schema)
        if _eff:
            sel += f", {_eff}"
        sql = (
            f"SELECT {sel} FROM Inventory_Log "
            f"WHERE {where} "
            f"GROUP BY [sub_location], [{schema_mac}], strftime('%Y-%m', [date]) "
            f"ORDER BY month, [sub_location], [{schema_mac}]"
        )
        return [("Tổng hợp", sql)]


# =====================================================================
# MODULE 4: AGGREGATION & FORMATTING
# =====================================================================

def _to_float(val) -> float:
    try:
        return float(str(val).replace(",", "").replace("%", "").strip())
    except Exception:
        return 0.0


def aggregate_df(df: pd.DataFrame, schema: dict) -> dict:
    """Tổng hợp một DataFrame thành dict số liệu."""
    cols = df.columns.tolist()
    qty_col  = next((c for c in cols if c.lower() == "total_kg"), None)
    if not qty_col:
        qty_col = next((c for c in cols if schema.get("qty_col","").lower() in c.lower()), None)
    eff_col  = next((c for c in cols if "avg_eff" in c.lower()), None)
    if not eff_col:
        eff_col = schema.get("eff_2ca_col") or schema.get("eff_a_col")
        if eff_col and eff_col not in cols:
            eff_col = None
    mac_col  = next((c for c in cols if c.lower() in ("ten_may_norm", "ten_may")), None)
    if not mac_col:
        mac_col = next((c for c in cols if "ten_may" in c.lower()), None)
    item_col = next((c for c in cols if "item_id" in c.lower()), None)

    # Làm sạch cột số
    if qty_col:
        df["_qty"] = df[qty_col].apply(_to_float)
    else:
        df["_qty"] = 0.0

    # ----------------------------------------------------------------
    # Tính hiệu suất theo thứ tự ưu tiên từng dòng:
    #   1. Hiệu suất 2 ca (nếu > 0)
    #   2. Trung bình ca A và ca B (nếu cả 2 đều > 0)
    #   3. Ca A hoặc ca B (lấy cái nào có số)
    #   4. 0 nếu không có dữ liệu
    # ----------------------------------------------------------------
    ea_col = schema.get("eff_a_col")
    eb_col = schema.get("eff_b_col")

    # Đọc từng cột vào df (nếu có)
    df["_eff_2ca"] = df[eff_col].apply(_to_float) if eff_col else 0.0
    df["_ea"]      = df[ea_col].apply(_to_float)  if (ea_col and ea_col in cols) else 0.0
    df["_eb"]      = df[eb_col].apply(_to_float)  if (eb_col and eb_col in cols) else 0.0

    def _resolve_eff(row):
        v2ca = row["_eff_2ca"]
        if v2ca > 0:
            return v2ca                              # Ưu tiên 1: 2ca đã có
        ea = row["_ea"]
        eb = row["_eb"]
        if ea > 0 and eb > 0:
            return round((ea + eb) / 2, 4)          # Ưu tiên 2: trung bình 2 ca
        if ea > 0:
            return ea                                # Ưu tiên 3: chỉ có ca A
        if eb > 0:
            return eb                                # Ưu tiên 4: chỉ có ca B
        return 0.0

    df["_eff"] = df.apply(_resolve_eff, axis=1)

    total_kg = round(df["_qty"].sum(), 2)
    valid_eff = df[df["_eff"] > 0]["_eff"]
    avg_eff  = round(valid_eff.mean(), 2) if not valid_eff.empty else 0.0

    # Theo tháng
    by_month = {}
    if "month" in cols:
        for mth, grp in df.groupby("month"):
            by_month[str(mth)] = {
                "kg": round(grp["_qty"].sum(), 2),
                "eff": round(grp[grp["_eff"] > 0]["_eff"].mean(), 2)
                       if not grp[grp["_eff"] > 0].empty else 0.0,
            }
    elif "date" in cols:
        df["_mth"] = df["date"].astype(str).str[:7]
        for mth, grp in df.groupby("_mth"):
            by_month[str(mth)] = {
                "kg": round(grp["_qty"].sum(), 2),
                "eff": round(grp[grp["_eff"] > 0]["_eff"].mean(), 2)
                       if not grp[grp["_eff"] > 0].empty else 0.0,
            }

    # Theo máy — hỗ trợ cả ten_may và ten_may_norm (MachineList query)
    by_machine = {}
    _mac_groupby = "ten_may_norm" if "ten_may_norm" in cols else mac_col
    if _mac_groupby and _mac_groupby in cols:
        for mac, grp in df.groupby(_mac_groupby):
            by_machine[str(mac)] = {
                "kg": round(grp["_qty"].sum(), 2),
                "eff": round(grp[grp["_eff"] > 0]["_eff"].mean(), 2)
                       if not grp[grp["_eff"] > 0].empty else 0.0,
            }
    # Fallback: MachineList query đã aggregate sẵn (có total_kg, avg_eff)
    elif "ten_may_norm" in cols and "total_kg" in cols:
        for _, row in df.iterrows():
            mac = str(row["ten_may_norm"])
            kg  = float(row.get("total_kg") or 0)
            eff = float(row.get("avg_eff") or 0)
            if kg > 0:
                by_machine[mac] = {"kg": kg, "eff": eff}

    # Theo xưởng
    by_location = {}
    if "sub_location" in cols:
        for loc, grp in df.groupby("sub_location"):
            by_location[str(loc)] = {
                "kg": round(grp["_qty"].sum(), 2),
                "eff": round(grp[grp["_eff"] > 0]["_eff"].mean(), 2)
                       if not grp[grp["_eff"] > 0].empty else 0.0,
            }

    # Mã hàng đang chạy
    items = []
    if item_col:
        items = df[item_col].dropna().astype(str).unique().tolist()[:20]

    # Chi tiết raw — 200 dòng để đủ show tất cả máy
    raw_rows = []
    display_cols = [c for c in ["date", "sub_location", mac_col, item_col, qty_col, eff_col] if c and c in cols]
    # Nếu có cột ten_may_norm (MachineList query), bổ sung vào
    if "ten_may_norm" in cols:
        display_cols = ["sub_location", "ten_may_norm", "total_kg", "avg_eff"] + [
            c for c in display_cols if c not in ["sub_location", "ten_may_norm", "total_kg", "avg_eff"]
        ]
    for _, row in df.head(200).iterrows():
        raw_rows.append({c: row.get(c) for c in display_cols if c})

    return {
        "total_rows": len(df),
        "total_kg": total_kg,
        "avg_eff": avg_eff,
        "by_month": by_month,
        "by_machine": by_machine,
        "by_location": by_location,
        "items": items,
        "raw_rows": raw_rows,
    }


def format_agg_for_display(label: str, agg: dict) -> str:
    """Tạo chuỗi hiển thị chi tiết từ agg dict."""
    lines = []
    if label != "Kết quả":
        lines.append(f"### 📊 {label}")

    if agg["total_kg"] > 0:
        lines.append(f"- **Tổng sản lượng:** {agg['total_kg']:,.2f} Kg")
    if agg["avg_eff"] > 0:
        lines.append(f"- **Hiệu suất trung bình:** {agg['avg_eff']}%")

    if agg["by_month"]:
        lines.append("- **Theo tháng:**")
        for mth, d in sorted(agg["by_month"].items()):
            m_label = f"{mth[5:7]}/{mth[:4]}"
            row_str = f"  + Tháng {m_label}: {d['kg']:,.2f} Kg"
            if d["eff"] > 0:
                row_str += f" | HS: {d['eff']}%"
            lines.append(row_str)

    elif agg["by_machine"] and not agg["by_month"]:
        lines.append("- **Theo máy:**")
        for mac, d in sorted(agg["by_machine"].items(), key=lambda x: x[1]["kg"], reverse=True):
            row_str = f"  + Máy {mac}: {d['kg']:,.2f} Kg"
            if d["eff"] > 0:
                row_str += f" | HS: {d['eff']}%"
            lines.append(row_str)

    active_locs = {loc: d for loc, d in agg["by_location"].items() if d["kg"] > 0}
    if active_locs and len(active_locs) > 1:
        lines.append("- **Theo xưởng:**")
        for loc, d in active_locs.items():
            row_str = f"  + {loc}: {d['kg']:,.2f} Kg"
            if d["eff"] > 0:
                row_str += f" | HS: {d['eff']}%"
            lines.append(row_str)

    if agg["items"]:
        lines.append(f"- **Mã hàng:** {', '.join(str(i) for i in agg['items'][:15])}")

    return "\n".join(lines)


def format_raw_table(label: str, agg: dict) -> str:
    """Tạo bảng chi tiết từng dòng (dùng cho SIMPLE/ngày) — kèm tổng ngày."""
    if not agg["raw_rows"]:
        return ""

    # Tính tổng ngày trước
    _day_total_kg  = agg.get("total_kg", 0)
    _day_avg_eff   = agg.get("avg_eff", 0)
    _n_machines    = len({row.get("ten_may","") for row in agg["raw_rows"]
                          if row.get("ten_may","") not in ("","nan","None")})

    header = (
        f"📅 **{label}**\n"
        f"  - Tổng sản lượng ngày: **{_day_total_kg:,.2f} Kg**"
        + (f" | Hiệu suất TB: **{_day_avg_eff}%**" if _day_avg_eff > 0 else "")
        + (f" | {_n_machines} máy" if _n_machines > 0 else "")
        + "\n"
    )

    # Chỉ show máy có SL > 0 để tránh dòng rỗng
    lines = [header, "\n**Chi tiết từng máy:**\n"]
    shown = 0
    for row in agg["raw_rows"]:
        qty_v = next(
            (v for k, v in row.items()
             if ("qty" in k.lower() or "quantity" in k.lower()) and _to_float(v) > 0),
            None,
        )
        if not qty_v:
            continue  # bỏ qua máy không có sản lượng

        mac = row.get("ten_may", "?")
        parts = [f"Máy {mac}"]

        item_v = next(
            (str(v) for k, v in row.items()
             if "item" in k.lower() and str(v) not in ["nan","","None"]),
            None,
        )
        if item_v:
            parts.append(f"Mã {item_v}")

        # Hiển thị màu nếu có
        color_v = next(
            (str(v) for k, v in row.items()
             if k.upper() == "COLOR" and str(v) not in ["nan","","None"]),
            None,
        )
        if color_v:
            parts.append(f"🎨 Màu: **{color_v}**")

        parts.append(f"SL: {_to_float(qty_v):,.2f} Kg")

        eff_v = next(
            (v for k, v in row.items() if "eff" in k.lower() or "hiệu suất" in k.lower()),
            None,
        )
        if eff_v and _to_float(eff_v) > 0:
            parts.append(f"HS: {_to_float(eff_v)}%")

        lines.append("- " + " | ".join(parts))
        shown += 1

    if shown == 0:
        lines.append("_(Không có máy nào có sản lượng > 0 trong ngày này)_")
    return "\n".join(lines)




def format_trend_table(label: str, agg: dict, n_days: int = 10) -> str:
    """
    Tạo bảng xu hướng từng ngày dạng text từ raw_rows.
    Gom theo ngày, tính tổng SL và HS trung bình mỗi ngày.
    """
    rows = agg.get("raw_rows", [])
    if not rows:
        return ""

    from collections import defaultdict
    by_day = defaultdict(lambda: {"kg": 0.0, "eff_sum": 0.0, "eff_n": 0, "items": set()})

    for r in rows:
        d = str(r.get("date", ""))[:10]
        if not d:
            continue
        for k, v in r.items():
            kl = k.lower()
            if any(x in kl for x in ["qty", "quantity", "total_kg"]):
                by_day[d]["kg"] += _to_float(v)
            if any(x in kl for x in ["eff", "hiệu suất", "효율"]):
                fv = _to_float(v)
                if fv > 0:
                    by_day[d]["eff_sum"] += fv
                    by_day[d]["eff_n"]   += 1
        item_k = next((k for k in r if "item" in k.lower()), None)
        if item_k and str(r.get(item_k, "")) not in ("nan", "", "None"):
            by_day[d]["items"].add(str(r[item_k]))

    if not by_day:
        return ""

    sorted_days = sorted(by_day.keys())
    kg_vals  = [by_day[d]["kg"]  for d in sorted_days]
    eff_vals = [by_day[d]["eff_sum"] / by_day[d]["eff_n"]
                if by_day[d]["eff_n"] > 0 else 0
                for d in sorted_days]

    max_kg  = max(kg_vals)  if kg_vals  else 0
    max_eff = max(eff_vals) if eff_vals else 0

    _actual_days = len(sorted_days)
    _header_note = f" (có {_actual_days}/{n_days} ngày dữ liệu)" if _actual_days < n_days else ""
    lines = [f"**📈 BẢNG XU HƯỚNG {n_days} NGÀY{_header_note} (từng ngày theo thứ tự):**\n"]
    lines.append("| Ngày | Sản lượng (Kg) | Hiệu suất | Mã hàng |")
    lines.append("|------|---------------|-----------|---------|")

    for d in sorted_days:
        v    = by_day[d]
        kg   = round(v["kg"], 2)
        eff_v = round(v["eff_sum"] / v["eff_n"], 1) if v["eff_n"] > 0 else 0
        it   = ", ".join(list(v["items"])[:2]) if v["items"] else "—"

        kg_str  = f"{kg:,.2f}" if kg > 0 else "—"
        eff_str = f"{eff_v}%"  if eff_v > 0 else "—"

        # Đánh dấu ngày đỉnh
        star = ""
        if kg > 0 and kg == max_kg:
            star += " 🏆"
        if eff_v > 0 and eff_v == max_eff:
            star += " ⭐"

        lines.append(f"| {d} | {kg_str}{star} | {eff_str} | {it} |")

    # Tổng kết nhanh
    total_kg_all = sum(kg_vals)
    avg_eff_all  = round(sum(e for e in eff_vals if e > 0) / max(1, len([e for e in eff_vals if e > 0])), 1)
    days_with_data = sum(1 for k in kg_vals if k > 0)
    # Phân tích xu hướng: tăng/giảm/không đủ dữ liệu
    _trend_arrow = ""
    if len(kg_vals) >= 2:
        _first_half = sum(kg_vals[:len(kg_vals)//2]) / max(1, len(kg_vals)//2)
        _second_half = sum(kg_vals[len(kg_vals)//2:]) / max(1, len(kg_vals) - len(kg_vals)//2)
        if _second_half > _first_half * 1.05: _trend_arrow = "📈 đang tăng"
        elif _second_half < _first_half * 0.95: _trend_arrow = "📉 đang giảm"
        else: _trend_arrow = "↔️ ổn định"

    _no_data_days = len(sorted_days) - days_with_data
    _summary_parts = [f"**Tổng: {total_kg_all:,.0f} Kg** | {days_with_data} ngày có dữ liệu"]
    if _no_data_days > 0:
        _summary_parts.append(f"{_no_data_days} ngày không có dữ liệu (nghỉ/ngừng máy)")
    if avg_eff_all: _summary_parts.append(f"HS trung bình: {avg_eff_all}%")
    if _trend_arrow: _summary_parts.append(f"Xu hướng: {_trend_arrow}")
    summary_line = "\n> " + " | ".join(_summary_parts)
    lines.append(summary_line)

    return "\n".join(lines)


def format_machine_list(agg: dict) -> str:
    """
    Hiển thị danh sách máy gom theo xưởng - dùng cho câu hỏi 'máy nào'.
    Dữ liệu đến từ GROUP BY query (có sub_location, ten_may_norm, total_kg, avg_eff).
    """
    rows = agg.get("raw_rows", [])
    if not rows:
        return "📭 Không có dữ liệu."

    # Gom theo sub_location
    from collections import defaultdict
    by_loc = defaultdict(list)
    for r in rows:
        loc = str(r.get("sub_location", "—"))
        # Tên máy: thử các key khác nhau
        mac = str(r.get("ten_may_norm") or r.get("ten_may") or "?")
        kg  = _to_float(r.get("total_kg", 0))
        eff_v = _to_float(r.get("avg_eff", 0))
        if kg <= 0:
            continue
        by_loc[loc].append({"mac": mac, "kg": kg, "eff": eff_v})

    if not by_loc:
        return "📭 Không có máy nào có sản lượng > 0."

    lines = []
    grand_total = 0
    for loc in sorted(by_loc.keys()):
        machines = sorted(by_loc[loc], key=lambda x: x["kg"], reverse=True)
        loc_total = sum(m["kg"] for m in machines)
        grand_total += loc_total
        lines.append(f"\n**🏭 {loc}** — {len(machines)} máy | Tổng: {loc_total:,.0f} Kg")
        for m in machines:
            eff_str = f" | HS: {m['eff']}%" if m["eff"] > 0 else ""
            lines.append(f"  - Máy **{m['mac']}**: {m['kg']:,.0f} Kg{eff_str}")

    lines.append(f"\n> **Tổng cộng:** {len([r for rows in by_loc.values() for r in rows])} máy | {grand_total:,.0f} Kg")
    return "\n".join(lines)


def format_comparison_table(results: list) -> str:
    """Bảng so sánh markdown cho nhiều query."""
    if len(results) < 2:
        return ""
    lines = ["\n**⚖️ BẢNG SO SÁNH:**\n"]
    lines.append("| Đơn vị | Sản lượng (Kg) | Hiệu suất TB |")
    lines.append("|--------|----------------|--------------|")
    for label, agg in results:
        kg  = f"{agg['total_kg']:,.2f}" if agg["total_kg"] > 0 else "—"
        eff = f"{agg['avg_eff']}%"       if agg["avg_eff"]  > 0 else "—"
        lines.append(f"| {label} | {kg} | {eff} |")
    return "\n".join(lines)


# =====================================================================
# MODULE 5: MAIN ENTRY POINT
# =====================================================================


def _handle_pdyd_compare(intent, user_prompt):
    """
    So sánh YD vs PD: tổng kg + hiệu suất trung bình cho từng nhóm.
    Query 2 lần với phan_loai = PD / YD.
    """
    import sqlite3 as _sq, pandas as _pd
    from datetime import date as _date

    as_of = str(_date.today())
    try:
        conn = _sq.connect("inventory.db")

        def _build_filter(pl):
            parts = [
                f"cluster_name = 'Xưởng Dệt'",
                f"UPPER(TRIM([phan_loai])) = '{pl}'"
            ]
            locs = intent.get("locations", [])
            if locs:
                loc_conds = " OR ".join(
                    f"LOWER([sub_location]) LIKE '%weaving {l}%'"
                    for l in locs
                )
                parts.append(f"({loc_conds})")

            dr = intent.get("date_range")
            periods = intent.get("periods", [])
            if dr:
                parts.append(f"[date] >= '{dr['from']}' AND [date] <= '{dr['to']}'")
            elif periods:
                m_conds = " OR ".join(f"[date] LIKE '%-{str(m).zfill(2)}-%'" for m in periods)
                parts.append(f"({m_conds})")
            return " AND ".join(parts)

        def _query(pl):
            where = _build_filter(pl)
            # Cột hiệu suất
            eff_cols = []
            try:
                cols_df = _pd.read_sql_query("PRAGMA table_info(Inventory_Log)", conn)
                cols = cols_df["name"].tolist()
                for c in cols:
                    cl = c.lower()
                    if ("hiệu suất" in cl or "효율" in cl or "eff" in cl) and "2 ca" in cl:
                        eff_cols.append(c); break
                if not eff_cols:
                    for c in cols:
                        cl = c.lower()
                        if "hiệu suất" in cl or "효율" in cl or "eff" in cl:
                            eff_cols.append(c); break
            except Exception:
                pass

            eff_expr = (
                f"ROUND(AVG(CAST(NULLIF(TRIM([{eff_cols[0]}]),'') AS REAL)),2) as avg_eff"
                if eff_cols else "NULL as avg_eff"
            )
            sql = (
                f"SELECT COUNT(DISTINCT [ten_may]) as n_may, "
                f"ROUND(SUM(CAST(REPLACE(REPLACE(COALESCE([quantity_kg],'0'),',',''),' ','') AS REAL)),2) as total_kg, "
                f"{eff_expr} "
                f"FROM Inventory_Log WHERE {where}"
            )
            try:
                df = _pd.read_sql_query(sql, conn)
                r = df.iloc[0]
                return {
                    "n_may":    int(r.get("n_may") or 0),
                    "total_kg": float(r.get("total_kg") or 0),
                    "avg_eff":  float(r.get("avg_eff") or 0),
                }
            except Exception as e:
                return {"n_may": 0, "total_kg": 0, "avg_eff": 0, "err": str(e)}

        yd = _query("YD")
        pd_r = _query("PD")
        conn.close()

        # Build time label
        dr = intent.get("date_range")
        periods = intent.get("periods", [])
        t_lbl = ""
        if dr:
            t_lbl = f" ({dr['from']} → {dr['to']})"
        elif periods:
            t_lbl = " tháng " + ", ".join(str(m) for m in sorted(periods))

        locs = intent.get("locations", [])
        loc_lbl = " — ".join(f"Weaving {l}" for l in locs) if locs else "Toàn xưởng"

        out = [f"## So Sanh YD vs PD — {loc_lbl}{t_lbl}\n\n"]
        out.append("| | **YD** | **PD** | Chênh lệch |\n")
        out.append("|---|---|---|---|\n")

        # Sản lượng
        diff_kg = yd["total_kg"] - pd_r["total_kg"]
        winner_kg = "YD ↑" if diff_kg > 0 else ("PD ↑" if diff_kg < 0 else "Bằng nhau")
        out.append(
            f"| Sản lượng (Kg) | **{yd['total_kg']:,.0f}** | **{pd_r['total_kg']:,.0f}** "
            f"| {winner_kg} ({abs(diff_kg):,.0f} Kg) |\n"
        )

        # Hiệu suất
        if yd["avg_eff"] > 0 or pd_r["avg_eff"] > 0:
            diff_eff = round(yd["avg_eff"] - pd_r["avg_eff"], 1)
            winner_eff = "YD ↑" if diff_eff > 0 else ("PD ↑" if diff_eff < 0 else "Bằng nhau")
            out.append(
                f"| Hiệu suất TB (%) | **{yd['avg_eff']}%** | **{pd_r['avg_eff']}%** "
                f"| {winner_eff} ({abs(diff_eff)}%) |\n"
            )
        else:
            out.append("| Hiệu suất TB (%) | — | — | *(chưa có dữ liệu eff)* |\n")

        # Số máy
        out.append(
            f"| Số máy | {yd['n_may']} máy | {pd_r['n_may']} máy | — |\n"
        )

        # Tỷ lệ
        total = yd["total_kg"] + pd_r["total_kg"]
        if total > 0:
            pct_yd = round(yd["total_kg"] / total * 100, 1)
            pct_pd = round(pd_r["total_kg"] / total * 100, 1)
            out.append(f"\n**Ty le:** YD {pct_yd}% | PD {pct_pd}% (tong {total:,.0f} Kg)\n")

        yield DummyChunk("".join(out))

        # Fallback comment nếu AI timeout
        _ask_eff = "hiệu suất" in user_prompt.lower() or "hiệu quả" in user_prompt.lower()
        if yd["avg_eff"] > 0 or pd_r["avg_eff"] > 0:
            if yd["avg_eff"] > pd_r["avg_eff"]:
                _comment = f"> 💬 YD hiệu suất cao hơn PD ({yd['avg_eff']}% vs {pd_r['avg_eff']}%)."
            elif pd_r["avg_eff"] > yd["avg_eff"]:
                _comment = f"> 💬 PD hiệu suất cao hơn YD ({pd_r['avg_eff']}% vs {yd['avg_eff']}%)."
            else:
                _comment = "> 💬 YD và PD có hiệu suất tương đương nhau."
        elif yd["total_kg"] > pd_r["total_kg"]:
            _comment = f"> 💬 YD sản lượng cao hơn PD ({yd['total_kg']:,.0f} vs {pd_r['total_kg']:,.0f} Kg)."
        else:
            _comment = f"> 💬 PD sản lượng cao hơn YD ({pd_r['total_kg']:,.0f} vs {yd['total_kg']:,.0f} Kg)."

        # Try AI first
        try:
            _prompt = (
                f"Cau hoi: '{user_prompt}'\n"
                f"YD: {yd['total_kg']:,.0f}Kg | HS:{yd['avg_eff']}%\n"
                f"PD: {pd_r['total_kg']:,.0f}Kg | HS:{pd_r['avg_eff']}%\n"
                "Viet 1-2 cau nhan xet ngan. Tieng Viet. Toi da 40 tu."
            )
            _cl = OpenAI(base_url=AI_BASE_URL, api_key=AI_API_KEY)
            _s = _cl.chat.completions.create(
                model=AI_MODEL,
                messages=[{"role":"user","content":_prompt}],
                stream=True, temperature=0.3, max_tokens=80, timeout=8,
            )
            _ai_txt = ""
            yield DummyChunk("\n\n> 💬 ")
            for _c in _s:
                if _c.choices[0].delta.content:
                    _ai_txt += _c.choices[0].delta.content
                    yield DummyChunk(_c.choices[0].delta.content)
            if not _ai_txt.strip():
                yield DummyChunk(_comment[5:])  # strip "> 💬 "
        except Exception:
            yield DummyChunk("\n\n" + _comment)

    except Exception as e:
        yield DummyChunk(f"Loi so sanh YD/PD: {e}\n")


def _generate_fallback_comment(intent: dict, results: list, user_prompt: str) -> str:
    """
    Rule-based commentary khi AI timeout.
    Luôn trả về 1-2 câu nhận xét dựa trên dữ liệu thực tế — không cần AI.
    """
    try:
        p = (user_prompt or "").lower()
        qt = intent.get("query_type", "")
        is_trend  = intent.get("flags", {}).get("is_trend", False)
        is_month  = bool(intent.get("periods"))
        date_lbl  = ""
        dr = intent.get("date_range")
        ex = intent.get("exact_date")
        if ex:
            date_lbl = f"ngày {ex[8:]}/{ex[5:7]}"
        elif dr:
            date_lbl = f"từ {dr['from']} đến {dr['to']}"
        elif is_month:
            date_lbl = "tháng " + ", ".join(str(m) for m in sorted(intent.get("periods", [])))

        # Gom dữ liệu tổng từ results
        total_kg  = 0.0
        avg_eff   = 0.0
        eff_count = 0
        all_kg_days = []  # cho trend
        top_item  = ""
        top_kg    = 0.0

        for _lbl, _agg, _ in (results or []):
            total_kg += _agg.get("total_kg", 0) or 0
            _e = _agg.get("avg_eff", 0) or 0
            if _e > 0:
                avg_eff   += _e
                eff_count += 1
            # by_machine
            for _m, _d in (_agg.get("by_machine", {}) or {}).items():
                if (_d.get("kg") or 0) > top_kg:
                    top_kg   = _d["kg"]
                    top_item = _m
            # by_day cho trend
            if is_trend:
                for _d, _dv in (_agg.get("by_day", {}) or {}).items():
                    all_kg_days.append((_d, _dv.get("kg", 0)))

        avg_eff = round(avg_eff / max(1, eff_count), 1)
        total_kg = round(total_kg, 1)

        lines = []

        # 1. TREND
        if is_trend and all_kg_days:
            all_kg_days.sort()
            vals = [v for _, v in all_kg_days if v > 0]
            if len(vals) >= 2:
                half = len(vals) // 2
                first_avg  = sum(vals[:half]) / half
                second_avg = sum(vals[half:]) / max(1, len(vals) - half)
                if second_avg > first_avg * 1.05:
                    trend_word = "đang tăng dần"
                elif second_avg < first_avg * 0.95:
                    trend_word = "có xu hướng giảm"
                else:
                    trend_word = "ổn định"
                best_day, best_kg = max(all_kg_days, key=lambda x: x[1])
                lines.append(
                    f"> 💬 Sản lượng {trend_word} trong khoảng này. "
                    f"Ngày cao nhất: **{best_day}** đạt **{best_kg:,.0f} Kg**."
                )
            return "\n".join(lines)

        # 2. Dữ liệu ngày/tháng thông thường
        if total_kg <= 0:
            return "> 💬 Không có dữ liệu sản xuất trong khoảng thời gian này."

        # Đánh giá mức sản lượng
        if total_kg >= 50000:
            sl_label = "rất cao"
        elif total_kg >= 30000:
            sl_label = "tốt"
        elif total_kg >= 15000:
            sl_label = "bình thường"
        elif total_kg >= 5000:
            sl_label = "thấp"
        else:
            sl_label = "rất thấp"

        # Câu 1: tổng sản lượng
        c1 = f"> 💬 Tổng sản lượng {date_lbl} đạt **{total_kg:,.0f} Kg** — mức {sl_label}."
        if avg_eff > 0:
            eff_label = "tốt" if avg_eff >= 70 else ("chấp nhận" if avg_eff >= 55 else "cần cải thiện")
            c1 += f" Hiệu suất trung bình **{avg_eff}%** ({eff_label})."
        lines.append(c1)

        # Câu 2: máy/điểm đáng chú ý
        if "hiệu suất" in p or "hs" in p:
            if avg_eff >= 70:
                lines.append("> Hiệu suất duy trì ổn định, tiếp tục theo dõi để giữ phong độ.")
            elif avg_eff < 55:
                lines.append("> Hiệu suất còn thấp — cần kiểm tra máy và công thức sợi.")
        elif top_item and top_kg > 0:
            lines.append(f"> Máy **{top_item}** đóng góp sản lượng cao nhất ({top_kg:,.0f} Kg).")

        return "\\n".join(lines)

    except Exception:
        return "> 💬 Dữ liệu đã hiển thị đầy đủ ở trên."


def _handle_beam_recent(intent, user_prompt):
    """
    Câu hỏi: "Beam nào lên máy gần nhất trên weaving 3?"
    → Query Beam_Info ORDER BY ngay_len_may DESC, lọc theo weaving/so_may nếu có.
    """
    try:
        import sqlite3 as _sq
        conn = _sq.connect("inventory.db")

        # Xây filter
        where_parts = ["ngay_len_may IS NOT NULL"]

        locs = intent.get("locations", [])
        macs = intent.get("machines", [])

        # Weaving filter
        _weaving_filter = None
        for loc in locs:
            loc_s = str(loc).strip()
            if loc_s.isdigit():
                _weaving_filter = f"Weaving {loc_s}"
            elif "weaving" in loc_s.lower():
                _weaving_filter = loc_s.title()
            elif loc_s.lower().startswith("w") and loc_s[1:].isdigit():
                _weaving_filter = f"Weaving {loc_s[1:]}"
        # Bắt thêm từ câu hỏi: "xưởng 3", "weaving 3"
        import re as _re3
        _wm = _re3.search(r"(?:xưởng|weaving)\s*(\d+)", user_prompt.lower())
        if _wm and not _weaving_filter:
            _weaving_filter = f"Weaving {_wm.group(1)}"

        if _weaving_filter:
            where_parts.append(f"LOWER(weaving) = LOWER('{_weaving_filter}')")

        # Machine filter
        if macs:
            mac_list = ", ".join(f"'{m}'" for m in macs)
            where_parts.append(f"(so_may IN ({mac_list}) OR CAST(so_may AS INTEGER) IN ({mac_list}))")

        where_sql = " AND ".join(where_parts)
        limit = 1 if (not _weaving_filter and not macs) else 10
        if _weaving_filter and not macs:
            limit = 10  # Tất cả máy của xưởng đó

        df = pd.read_sql_query(
            f"SELECT ma_beam, weaving, so_may, loai_may, ten_hang, loai_soi, "
            f"ngay_len_may, so_met, so_kg_thuc_te, phan_loai "
            f"FROM Beam_Info WHERE {where_sql} "
            f"ORDER BY ngay_len_may DESC "
            f"LIMIT {limit}",
            conn
        )
        conn.close()

        if df.empty:
            msg = f"Không tìm thấy beam nào"
            if _weaving_filter:
                msg += f" trên {_weaving_filter}"
            yield DummyChunk(msg + " trong Beam_Info.\n")
            return

        # Format output
        scope = f"**{_weaving_filter}**" if _weaving_filter else "tất cả xưởng"
        if macs:
            scope += f" Máy {', '.join(str(m) for m in macs)}"

        out = [f"## Beam lên máy gần nhất — {scope}\n"]

        for i, (_, r) in enumerate(df.iterrows()):
            ngay = str(r.get("ngay_len_may", ""))[:10]
            may  = r.get("so_may", "?")
            weav = r.get("weaving", "")
            ma   = r.get("ma_beam", "?")
            hang = str(r.get("ten_hang", "") or "—")[:40]
            soi  = str(r.get("loai_soi", "") or "—")
            met  = r.get("so_met", 0) or 0
            kg   = r.get("so_kg_thuc_te", 0) or 0
            pl   = str(r.get("phan_loai", "") or "")
            loai = str(r.get("loai_may", "") or "")

            medal = ["🥇","🥈","🥉"] + ["▪️"]*20
            out.append(
                f"{medal[i]} **Máy {may}** ({weav}) — lên ngày **{ngay}**\n"
                f"  Mã beam: `{ma}` | {hang}\n"
                f"  Loại máy: {loai} | Sợi: {soi} | "
                f"{met:,.0f}m / {kg:,.1f}kg"
                + (f" | {pl}" if pl else "") + "\n\n"
            )

        yield DummyChunk("".join(out))

        # AI nhận xét
        try:
            _prompt = (
                f"Cau hoi: '{user_prompt}'\n"
                + "".join(out[:3]) + "\n\n"
                + "1 cau nhan xet ngan ve beam vua len may. Tieng Viet. Toi da 30 tu."
            )
            _cl = OpenAI(base_url=AI_BASE_URL, api_key=AI_API_KEY)
            _s = _cl.chat.completions.create(
                model=AI_MODEL,
                messages=[{"role":"user","content":_prompt}],
                stream=True, temperature=0.3, max_tokens=60, timeout=8,
            )
            yield DummyChunk("\n\n> 💬 ")
            for _c in _s:
                if _c.choices[0].delta.content:
                    yield DummyChunk(_c.choices[0].delta.content)
        except Exception:
            pass

    except Exception as e:
        yield DummyChunk(f"Lỗi đọc Beam_Info: {e}\n")


def _handle_beam_warehouse_status(intent, user_prompt):
    """
    Tồn kho Beam Weaving = Inventory_Log với cluster_name = 'Kho Beam Weaving'
    Tính: SUM(NHAP + TON_DAU) - SUM(XUAT) theo item_id + sub_location
    Đây là nguồn dữ liệu giống dashboard sơ đồ giá treo beam.
    """
    try:
        import sqlite3 as _sq
        conn = _sq.connect("inventory.db")

        # Lấy đúng cluster name cho kho beam weaving
        df_clusters = pd.read_sql_query(
            "SELECT DISTINCT cluster_name FROM Inventory_Log "
            "WHERE LOWER(cluster_name) LIKE '%beam%' OR LOWER(cluster_name) LIKE '%kho%' "
            "ORDER BY cluster_name",
            conn
        )
        beam_clusters = df_clusters["cluster_name"].tolist()

        # Tìm cluster khớp nhất với "Kho Beam Weaving"
        target_cluster = None
        for c in beam_clusters:
            cl = c.lower()
            if "beam weaving" in cl or "beam wea" in cl:
                target_cluster = c
                break
        if not target_cluster and beam_clusters:
            target_cluster = beam_clusters[0]

        if not target_cluster:
            yield DummyChunk("Khong tim thay cluster Kho Beam Weaving trong database.\n"
                             f"Cac cluster hien co: {beam_clusters}")
            conn.close()
            return

        # Tính tồn kho: nhập - xuất (giống logic trong app.py dòng 860-863)
        df_nhap = pd.read_sql_query(
            f"SELECT item_id, sub_location, "
            f"SUM(CAST(REPLACE(REPLACE(COALESCE(quantity,0),',',''),' ','') AS REAL)) AS sl_nhap "
            f"FROM Inventory_Log "
            f"WHERE cluster_name = '{target_cluster}' "
            f"AND type IN ('NHAP','TON_DAU') "
            f"AND item_id IS NOT NULL "
            f"GROUP BY item_id, sub_location",
            conn
        )
        df_xuat = pd.read_sql_query(
            f"SELECT item_id, sub_location, "
            f"SUM(CAST(REPLACE(REPLACE(COALESCE(quantity,0),',',''),' ','') AS REAL)) AS sl_xuat "
            f"FROM Inventory_Log "
            f"WHERE cluster_name = '{target_cluster}' "
            f"AND type = 'XUAT' "
            f"AND item_id IS NOT NULL "
            f"GROUP BY item_id, sub_location",
            conn
        )
        # Cũng thử với quantity_kg nếu quantity không có
        df_nhap_kg = pd.read_sql_query(
            f"SELECT item_id, sub_location, "
            f"SUM(CAST(REPLACE(REPLACE(COALESCE(quantity_kg,0),',',''),' ','') AS REAL)) AS sl_nhap "
            f"FROM Inventory_Log "
            f"WHERE cluster_name = '{target_cluster}' "
            f"AND type IN ('NHAP','TON_DAU') "
            f"AND item_id IS NOT NULL "
            f"GROUP BY item_id, sub_location",
            conn
        )
        conn.close()

        # Chọn cột có dữ liệu
        use_df_nhap = df_nhap if df_nhap["sl_nhap"].sum() > 0 else df_nhap_kg

        if use_df_nhap.empty or use_df_nhap["sl_nhap"].sum() == 0:
            yield DummyChunk(
                f"Cluster '{target_cluster}' chua co du lieu nhap kho.\n"
                f"Tat ca cac cluster beam hien co: {beam_clusters}\n"
                "Hay import file kho beam truoc."
            )
            return

        # Merge nhập - xuất
        df_stock = use_df_nhap.merge(df_xuat, on=["item_id","sub_location"], how="left")
        df_stock["sl_xuat"] = df_stock["sl_xuat"].fillna(0)
        df_stock["ton_kho"] = df_stock["sl_nhap"] - df_stock["sl_xuat"]
        df_stock = df_stock[df_stock["ton_kho"] > 0]

        lines_out = [f"## Ton Kho: {target_cluster}\n\n"]

        if df_stock.empty:
            lines_out.append("> Khong con beam nao trong kho (da xuat het).\n")
        else:
            tong_sl    = df_stock["ton_kho"].sum()
            so_ma      = df_stock["item_id"].nunique()
            so_vi_tri  = df_stock["sub_location"].nunique()

            lines_out.append(f"**Tong ton kho: {tong_sl:,.0f} | {so_ma} ma hang | {so_vi_tri} vi tri**\n\n")

            # Tóm tắt theo khu vực giá
            import re as _re
            def get_gia(loc):
                m = _re.search(r"(60|95)", str(loc))
                return f"Gia {m.group(1)} cho" if m else "Khu khac/Duoi Dat"

            df_stock["khu_vuc"] = df_stock["sub_location"].apply(get_gia)
            df_by_khu = df_stock.groupby("khu_vuc")["ton_kho"].agg(["sum","count"]).reset_index()
            df_by_khu.columns = ["khu_vuc","tong","so_beam"]

            lines_out.append("| Khu Vuc | So Beam | Tong SL |\n")
            lines_out.append("|---------|---------|---------|\n")
            for _, r in df_by_khu.iterrows():
                lines_out.append(f"| {r['khu_vuc']} | {int(r['so_beam'])} | {r['tong']:,.0f} |\n")

            lines_out.append(f"\n**Tong tat ca: {tong_sl:,.0f}**\n")

            # Top 10 mã hàng nhiều nhất
            df_top = df_stock.groupby("item_id")["ton_kho"].sum().sort_values(ascending=False).head(10)
            if not df_top.empty:
                lines_out.append("\n**Top ma hang nhieu nhat:**\n")
                for ma, qty in df_top.items():
                    lines_out.append(f"- {ma}: {qty:,.0f}\n")

        out = "".join(lines_out)
        yield DummyChunk(out)

        try:
            _prompt = (
                "Cau hoi: '" + user_prompt + "'\n"
                "Du lieu ton kho beam:\n" + out[:500] + "\n\n"
                "Viet 1-2 cau nhan xet ngan. Tieng Viet. Toi da 40 tu."
            )
            _cl = OpenAI(base_url=AI_BASE_URL, api_key=AI_API_KEY)
            _stream = _cl.chat.completions.create(
                model=AI_MODEL,
                messages=[{"role": "user", "content": _prompt}],
                stream=True, temperature=0.3, max_tokens=80, timeout=12,
            )
            yield DummyChunk("\n\n> 💬 ")
            for _c in _stream:
                if _c.choices[0].delta.content:
                    yield DummyChunk(_c.choices[0].delta.content)
        except Exception:
            pass

    except Exception as e:
        yield DummyChunk(f"Loi doc ton kho beam: {e}\n")


def _handle_beam_status(intent, user_prompt, current_view_date=None):
    from datetime import date as _date
    try:
        from beam_calculator import calc_beam_remaining, get_beam_status_table
    except ImportError as e:
        yield DummyChunk("beam_calculator.py chua tim thay: " + str(e))
        return

    as_of = (current_view_date or str(_date.today()))[:10]
    macs  = intent.get("machines", [])
    locs  = intent.get("locations", [])
    pairs = intent.get("machine_location_pairs", [])

    if pairs:
        targets = [(p["weaving"], str(p["mac"])) for p in pairs]
    elif macs and locs:
        targets = [(locs[0], str(m)) for m in macs]
    elif macs and not locs:
        targets = [(("Weaving " + str(w)), str(m)) for w in [1,2,3] for m in macs]
    else:
        targets = None

    if targets is None:
        weaving_filter = locs[0] if locs else None
        df_st = get_beam_status_table(weaving_filter, as_of)
        if df_st.empty:
            yield DummyChunk("Chua co du lieu beam. Hay nhap file BEAM_DAT truoc.\n")
            return
        label = weaving_filter or "Tat ca xuong"
        total = len(df_st)
        out = "**Trang thai beam - " + label + "** (tinh den " + as_of + ")\n"
        out += str(total) + " may co beam\n\n"
        for _, row in df_st.iterrows():
            bt_rm  = row.get("Beam trên còn (m)")
            bt_pct = row.get("Beam trên còn (%)")
            bd_rm  = row.get("Beam dưới còn (m)")
            ok = (row.get("CT sợi") == "OK")
            xuong = str(row.get("Xưởng", ""))
            may   = str(row.get("Máy", ""))
            mb    = str(row.get("Mã Beam", ""))
            th    = str(row.get("Tên hàng", ""))[:30]
            tm    = float(row.get("Tổng mét", 0) or 0)
            ln  = "- **" + xuong + " May " + may + "** | " + mb + " | " + th + " | " + f"{tm:,.0f}" + "m"
            if ok and bt_rm is not None:
                ln += " -> tren: " + f"{bt_rm:,.0f}" + "m (" + str(bt_pct) + "%) | duoi: " + f"{bd_rm:,.0f}" + "m"
            else:
                ln += " | Chua co CT soi"
            out += ln + "\n"
        yield DummyChunk(out)
        return

    parts = []
    checked = set()
    for weaving, mac in targets:
        key = (weaving, mac)
        if key in checked:
            continue
        r = calc_beam_remaining(weaving, mac, as_of)
        if "error" in r:
            for w in [1, 2, 3]:
                alt = "Weaving " + str(w)
                if alt == weaving or (alt, mac) in checked:
                    continue
                r2 = calc_beam_remaining(alt, mac, as_of)
                if "error" not in r2:
                    r = r2
                    break
        if "error" in r:
            parts.append("- " + weaving + " May " + mac + ": Khong tim thay beam\n")
            continue
        checked.add(key)

        bt    = r["beam_tren"]
        bd    = r["beam_duoi"]
        has_f = r["has_formula"]
        kpm   = r["kg_per_meter"]
        kg_u  = float(r["kg_used_total"] or 0)
        init_m  = float(r["initial_m"] or 0)
        init_kg = float(r["initial_kg"] or 0)
        mach  = str(r["machine"])
        mb    = str(r["ma_beam"])
        th    = str(r["ten_hang"])
        lmay  = str(r["loai_may"])
        pl    = str(r["phan_loai"])
        ngay  = str(r["ngay_len_may"])

        # --- Format đẹp cho từng máy ---
        pct_bar = lambda pct: ("█" * int((pct or 0)/10) + "░" * (10 - int((pct or 0)/10)))[:10] if pct is not None else "░░░░░░░░░░"

        blk  = "\n---\n"
        blk += "### " + mach + "\n"
        blk += "| | |\n|---|---|\n"
        blk += "| **Mã Beam** | `" + mb + "` |\n"
        blk += "| **Mã hàng** | " + th + " |\n"
        blk += "| **Loại máy** | " + lmay + " |\n"
        blk += "| **Ngày lên** | " + ngay + " |\n"
        blk += "| **Tổng ban đầu** | " + f"{init_m:,.0f}" + "m / " + f"{init_kg:,.0f}" + "kg |\n"
        blk += "| **Đã dệt** | " + f"{kg_u:,.0f}" + "kg |\n\n"

        # ── Lấy hệ số sợi từ beam_calculator result (đã tính sẵn) ──────
        _loai_soi_b = str(r.get("loai_soi") or "").strip()
        _tong_soi_b = float(r.get("tong_soi") or 0)
        _he_so_b    = r.get("he_so")
        # Nếu beam_calculator cũ chưa trả he_so → tự tính
        if not _he_so_b and _loai_soi_b:
            try:
                from mtr_kg import find_he_so as _fhs3
                _he_so_b = _fhs3(_loai_soi_b)
            except Exception:
                _he_so_b = None
        # Import kg_to_mtr
        try:
            from mtr_kg import kg_to_mtr as _k2m3
        except Exception:
            _k2m3 = None
        # Nếu tong_soi = 0 → thử lấy từ DB trực tiếp
        if (not _tong_soi_b or not _loai_soi_b):
            try:
                from beam_calculator import get_current_beam as _gcb2
                _bi2 = _gcb2(weaving, mac, as_of)
                if not _loai_soi_b: _loai_soi_b = str((_bi2 or {}).get("loai_soi") or "").strip()
                if not _tong_soi_b: _tong_soi_b = float((_bi2 or {}).get("tong_soi") or 0)
                if not _he_so_b and _loai_soi_b:
                    from mtr_kg import find_he_so as _fhs4
                    _he_so_b = _fhs4(_loai_soi_b)
            except Exception:
                pass

        _bong_pct = float(r.get("pct_bong") or 0)
        _nen_pct  = float(r.get("pct_nen")  or 0)
        if _bong_pct + _nen_pct < 1:
            from beam_calculator import get_yarn_formula as _gyf
            _fm = _gyf(th)
            _bong_pct = float(_fm.get("soi_bong_pct") or 0) if _fm else 0
            _nen_pct  = float(_fm.get("soi_nen_pct")  or 0) if _fm else 0
        if _bong_pct + _nen_pct < 1:
            _bong_pct = 50; _nen_pct = 50

        # Hiển thị công thức đang dùng
        if _he_so_b and _tong_soi_b > 0:
            blk += "> 📐 **Công thức hệ số sợi:** `" + _loai_soi_b + " | tong_soi=" + str(int(_tong_soi_b)) + " | he_so=" + str(_he_so_b) + "`\n\n"
        elif kpm:
            blk += "> 📐 **Công thức kg/m (fallback):** `" + f"{init_kg:,.0f}" + "/" + f"{init_m:,.0f}" + " = " + str(kpm) + "`\n\n"

        for _pct_b, icon_b, lv_b in [
            (_bong_pct, "🔼", "Beam TRÊN (sợi bông)"),
            (_nen_pct,  "🔽", "Beam DƯỚI (sợi nền)")
        ]:
            used_kg_b = round(kg_u * _pct_b / 100, 1)
            # Ưu tiên hệ số sợi, fallback kg/m
            if _he_so_b and _tong_soi_b > 0 and used_kg_b > 0 and _k2m3 is not None:
                try:
                    _rr = _k2m3(used_kg_b, _tong_soi_b, _loai_soi_b)
                    used_m_b  = round(_rr.get("mtr", 0), 0)
                    _cn = f"hệ số {_he_so_b}"
                except Exception as _k2m_err:
                    # fallback nếu lỗi
                    used_m_b = round(used_kg_b / kpm, 0) if kpm and kpm > 0 else 0
                    _cn = f"kg/m (err: {_k2m_err})"
            elif kpm and kpm > 0:
                used_m_b = round(used_kg_b / kpm, 0) if used_kg_b > 0 else 0
                _cn = f"÷ {kpm}"
            else:
                used_m_b = 0; _cn = "?"
            rem_m_b   = round(max(init_m - used_m_b, 0), 0)
            pct_rem_b = round(rem_m_b / init_m * 100, 1) if init_m > 0 else 0
            bar_b     = pct_bar(pct_rem_b)
            stat_b    = "✅" if pct_rem_b > 20 else ("⚠️" if pct_rem_b > 5 else "🔴")
            st_b      = _loai_soi_b or ""
            blk += icon_b + " **" + lv_b + "** (" + st_b + ") — tỷ lệ **" + str(round(_pct_b, 1)) + "%**\n"
            calc_b = (
                "  > " + f"{kg_u:,.1f}" + "kg × " + str(round(_pct_b,1)) + "% = **"
                + f"{used_kg_b:,.1f}" + "kg** → **" + f"{used_m_b:,.0f}" + "m dùng** (" + _cn + ")"
            )
            blk += calc_b + "\n"
            # kg còn lại ước tính
            if used_m_b > 0:
                rem_kg_b = round(rem_m_b * used_kg_b / used_m_b, 1)
                kg_str_b = " / **" + f"{rem_kg_b:,.1f}" + "kg**"
            else:
                kg_str_b = ""
            blk += "  " + bar_b + " còn **" + f"{rem_m_b:,.0f}" + "m**" + kg_str_b + " (" + str(pct_rem_b) + "%) " + stat_b + "\n"
        parts.append(blk)  # ✅ FIX: append blk vào parts sau khi xây xong

    yield DummyChunk("".join(parts))  # ✅ yield SAU khi duyệt hết targets



def _handle_sizing_query(intent, user_prompt, current_view_date=None):
    """
    Xử lý câu hỏi về Sizing / Sectional / Direct / Winder.
    Query trực tiếp Sizing_Log — không dùng Inventory_Log.
    """
    from datetime import date as _date
    import sqlite3, pandas as pd

    as_of = (current_view_date or str(_date.today()))[:10]
    p = user_prompt.lower()

    # Xác định machine_type
    mtype = intent.get("sz_machine_type")
    if not mtype:
        # Default từ câu hỏi
        for kw, mt in [("bng","MÁY HỒ"),("benninger","MÁY HỒ"),("karlmayer","MÁY HỒ"),
                        ("honghwa","MÁY HỒ"),("máy hồ","MÁY HỒ"),
                        ("sectional","MÁY SEC"),("máy sec","MÁY SEC"),
                        ("direct","MÁY QS"),("máy qs","MÁY QS"),
                        ("winder","WINDER"),("suzuki","SUZUKI")]:
            if kw in p: mtype = mt; break
        if not mtype: mtype = None  # Tất cả

    # Build base WHERE
    mtype_filter = f"machine_type = '{mtype}'" if mtype else "1=1"

    # Detect tháng
    import re as _re
    month_m = _re.search(r'tháng\s*(\d+)', p)
    year_m  = _re.search(r'(\d{4})', p)
    month_filter = ""
    if month_m:
        mm = int(month_m.group(1))
        yy = int(year_m.group(1)) if year_m else int(as_of[:4])
        month_filter = f"AND strftime('%Y-%m', date) = '{yy:04d}-{mm:02d}'"
    elif "tháng này" in p or "tháng nay" in p:
        month_filter = f"AND strftime('%Y-%m', date) = '{as_of[:7]}'"

    # Detect "hôm nay"
    if "hôm nay" in p or "ngày nay" in p:
        month_filter = f"AND date = '{as_of}'"

    # Detect machine name (Bng, Karlmayer...)
    machine_name_filter = ""
    for mac_name in ["bng","karlmayer","honghwa","benninger"]:
        if mac_name in p:
            machine_name_filter = f"AND LOWER(ten_may) LIKE '%{mac_name}%'"
            break

    try:
        conn = sqlite3.connect("inventory.db")

        # ── Câu hỏi: so sánh tốc độ thực tế và mục tiêu ──
        if "tốc độ" in p and ("mục tiêu" in p or "so sánh" in p or "thực tế" in p):
            df = pd.read_sql_query(f"""
                SELECT ten_may, date,
                    AVG(toc_do_muc_tieu) as toc_do_mt,
                    AVG(toc_do_thuc_te) as toc_do_tt,
                    COUNT(*) as so_ca
                FROM Sizing_Log
                WHERE {mtype_filter} {month_filter} {machine_name_filter}
                GROUP BY ten_may
                ORDER BY toc_do_tt DESC
            """, conn)
            if df.empty:
                yield DummyChunk("Không có dữ liệu tốc độ.\n"); conn.close(); return
            out = f"**So sánh tốc độ — {mtype or 'Sizing'}** {month_filter.replace('AND','').strip()}\n\n"
            for _, r in df.iterrows():
                mt = float(r['toc_do_mt'] or 0); tt = float(r['toc_do_tt'] or 0)
                pct = round(tt/mt*100,1) if mt>0 else 0
                bar_n = int(pct/10); bar = "█"*bar_n + "░"*(10-bar_n)
                out += f"**{r['ten_may']}** ({int(r['so_ca'])} ca)\n"
                out += f"  Mục tiêu: {mt:,.0f} m/min | Thực tế: {tt:,.0f} m/min\n"
                out += f"  {bar} **{pct}%** đạt\n\n"
            yield DummyChunk(out); conn.close(); return

        # ── Câu hỏi: thời gian chạy trung bình ──
        if "thời gian" in p and ("trung bình" in p or "bao lâu" in p):
            df = pd.read_sql_query(f"""
                SELECT ten_may,
                    AVG(thoi_gian_phut) as avg_phut,
                    MIN(thoi_gian_phut) as min_phut,
                    MAX(thoi_gian_phut) as max_phut,
                    COUNT(*) as so_lan
                FROM Sizing_Log
                WHERE {mtype_filter} {month_filter} {machine_name_filter}
                  AND thoi_gian_phut > 0
                GROUP BY ten_may ORDER BY avg_phut DESC
            """, conn)
            if df.empty:
                yield DummyChunk("Không có dữ liệu thời gian.\n"); conn.close(); return
            out = f"**Thời gian chạy — {mtype or 'Sizing'}**\n\n"
            for _, r in df.iterrows():
                avg = float(r['avg_phut'] or 0)
                out += f"**{r['ten_may']}**: TB {avg:,.0f} phút ({avg/60:.1f}h) | Min: {r['min_phut']:.0f} | Max: {r['max_phut']:.0f} | {int(r['so_lan'])} lần\n"
            yield DummyChunk(out); conn.close(); return

        # ── Câu hỏi: hiệu suất ──
        if any(k in p for k in ["hiệu suất","hs","efficiency"]):
            df = pd.read_sql_query(f"""
                SELECT ten_may,
                    AVG(hieu_suat_pct) as avg_hs,
                    MIN(hieu_suat_pct) as min_hs,
                    MAX(hieu_suat_pct) as max_hs,
                    COUNT(*) as so_ca,
                    SUM(sl_thuc_te_mtr) as tong_mtr,
                    SUM(sl_kg) as tong_kg
                FROM Sizing_Log
                WHERE {mtype_filter} {month_filter} {machine_name_filter}
                GROUP BY ten_may ORDER BY avg_hs DESC
            """, conn)
            if df.empty:
                yield DummyChunk("Không có dữ liệu hiệu suất.\n"); conn.close(); return
            label = month_filter.replace("AND strftime('%Y-%m', date) = ","tháng ").replace("'","").strip() if month_filter else "toàn bộ"
            out = f"**Hiệu suất {mtype or 'Sizing'} — {label}**\n\n"
            for _, r in df.iterrows():
                hs = float(r['avg_hs'] or 0)
                color = "✅" if hs > 0.6 else ("⚠️" if hs > 0.4 else "🔴")
                out += f"{color} **{r['ten_may']}**: {hs:.1%} TB | Min {r['min_hs']:.1%} | Max {r['max_hs']:.1%} | {int(r['so_ca'])} ca\n"
                if r['tong_mtr']: out += f"   → {float(r['tong_mtr']):,.0f}m / {float(r['tong_kg'] or 0):,.1f}kg\n"
            yield DummyChunk(out); conn.close(); return

        # ── Câu hỏi: tổng mét / sản lượng ──
        if any(k in p for k in ["tổng mét","sản lượng","bao nhiêu mét","bao nhiêu kg"]):
            df = pd.read_sql_query(f"""
                SELECT ten_may,
                    SUM(sl_thuc_te_mtr) as tong_mtr,
                    SUM(sl_kg) as tong_kg,
                    AVG(hieu_suat_pct) as avg_hs,
                    COUNT(*) as so_ca
                FROM Sizing_Log
                WHERE {mtype_filter} {month_filter} {machine_name_filter}
                GROUP BY ten_may ORDER BY tong_mtr DESC
            """, conn)
            if df.empty:
                yield DummyChunk("Không có dữ liệu sản lượng.\n"); conn.close(); return
            label = month_filter.replace("AND strftime('%Y-%m', date) = ","tháng ").replace("'","").strip() if month_filter else "toàn bộ"
            out = f"**Sản lượng {mtype or 'Sizing'} — {label}**\n\n"
            tot_m = df['tong_mtr'].sum(); tot_k = df['tong_kg'].sum()
            out += f"Tổng: **{tot_m:,.0f}m** / **{tot_k:,.1f}kg**\n\n"
            for _, r in df.iterrows():
                hs = float(r['avg_hs'] or 0)
                out += f"**{r['ten_may']}**: {float(r['tong_mtr'] or 0):,.0f}m — {float(r['tong_kg'] or 0):,.1f}kg | HS {hs:.1%} | {int(r['so_ca'])} ca\n"
            yield DummyChunk(out); conn.close(); return

        # ── Default: tóm tắt ──
        df = pd.read_sql_query(f"""
            SELECT ten_may, date, ca, loai_soi, sl_thuc_te_mtr, sl_kg, hieu_suat_pct
            FROM Sizing_Log
            WHERE {mtype_filter} {month_filter} {machine_name_filter}
            ORDER BY date DESC LIMIT 20
        """, conn)
        if df.empty:
            yield DummyChunk("Không có dữ liệu phù hợp.\n")
        else:
            yield DummyChunk(df.to_markdown(index=False) + "\n")
        conn.close()
    except Exception as e:
        yield DummyChunk(f"Lỗi query Sizing_Log: {e}\n")


def process_ai_chat(
    user_prompt, messages, selected_node, model_name,
    df_clean=None, current_view_date=None
):
    """
    Generator: yield DummyChunk(text) liên tục để Streamlit stream ra.
    """
    model_name = AI_MODEL

    client = OpenAI(base_url=AI_BASE_URL, api_key=AI_API_KEY)

    try:
        # ----------------------------------------------------------------
        # 1. Phát hiện ý định
        # ----------------------------------------------------------------
        intent = detect_intent(user_prompt)

        # ----------------------------------------------------------------
        # 1b. BEAM_STATUS — tính beam còn lại (trước mọi query khác)
        # ----------------------------------------------------------------
        if intent.get("query_type") == "SIZING_QUERY" or intent.get("flags", {}).get("is_sizing"):
            yield from _handle_sizing_query(intent, user_prompt, current_view_date)
            return

        if intent.get("query_type") == "SIZING_QUERY" or intent.get("flags", {}).get("is_sizing"):
            yield from _handle_sizing_query(intent, user_prompt, current_view_date)
            return

        # ── PDYD_COMPARE: so sánh YD vs PD ──────────────────────────────────
        if intent.get("query_type") == "PDYD_COMPARE" or intent.get("flags", {}).get("is_pdyd_compare"):
            yield from _handle_pdyd_compare(intent, user_prompt)
            return

        # ── BEAM_RECENT: beam nào lên máy gần nhất ─────────────────────────
        if intent.get("query_type") == "BEAM_RECENT" or intent.get("flags", {}).get("is_beam_recent"):
            yield from _handle_beam_recent(intent, user_prompt)
            return

        # ── BEAM_WAREHOUSE_STATUS: tồn kho beam (chưa lên máy) ───────────────
        if intent.get("query_type") == "BEAM_WAREHOUSE_STATUS" or intent.get("flags", {}).get("is_beam_warehouse_status"):
            yield from _handle_beam_warehouse_status(intent, user_prompt)
            return

        if intent.get("query_type") == "BEAM_STATUS" or intent.get("flags", {}).get("is_beam_status"):
            yield from _handle_beam_status(intent, user_prompt, current_view_date)
            return

        # ----------------------------------------------------------------
        # 2. Chế độ giải thích → AI đọc lịch sử và giải thích
        # ----------------------------------------------------------------
        if intent.get("query_type") == "EXPLAIN":
            # Chỉ lấy 2 message gần nhất (câu hỏi + trả lời ngay trước)
            # để tránh nhầm lẫn với context CS 32 hay câu hỏi cũ
            recent = messages[-3:-1] if len(messages) >= 3 else messages[:-1]
            # Trích xuất phần số liệu từ assistant message gần nhất
            last_assistant = ""
            for m in reversed(recent):
                if m.get("role") == "assistant":
                    last_assistant = m.get("content", "")[:1500]
                    break
            last_user_q = ""
            for m in reversed(recent):
                if m.get("role") == "user":
                    last_user_q = m.get("content", "")[:300]
                    break

            explain_prompt = (
                "Bạn là chuyên gia dữ liệu nhà máy dệt, đang GIẢI THÍCH cách tính toán cho Sếp.\n\n"
                f"CÂU HỎI TRƯỚC ĐÓ: \"{last_user_q}\"\n\n"
                f"KẾT QUẢ ĐÃ BÁO CÁO (trả lời ngay trước):\n{last_assistant}\n\n"
                f"CÂU HỎI MỚI: \"{user_prompt}\"\n\n"
                "NHIỆM VỤ:\n"
                "1. Tìm đúng con số được hỏi trong KẾT QUẢ ĐÃ BÁO CÁO ở trên.\n"
                "2. Giải thích: số đó = SUM(quantity_kg) của máy/xưởng/mã hàng nào, lọc theo điều kiện gì.\n"
                "3. Nêu điều kiện WHERE thực tế: lọc sub_location + ten_may + mã hàng + tháng.\n"
                "4. Nếu số có vẻ bất thường (1 máy mà vài trăm nghìn Kg) → cảnh báo thẳng.\n"
                "NGUYÊN TẮC: Chỉ giải thích dựa trên KẾT QUẢ ĐÃ BÁO CÁO, không bịa. "
                "Tiếng Việt rõ ràng. Tối đa 120 từ."
            )
            stream = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": explain_prompt}],
                stream=True,
                temperature=0.2,
                max_tokens=400,
            )
            for chunk in stream:
                if chunk.choices[0].delta.content:
                    yield DummyChunk(chunk.choices[0].delta.content)
            return


        # ----------------------------------------------------------------
        # 2b. Gắn current_view_date vào intent nếu user nói "ngày này/hôm nay"
        # ----------------------------------------------------------------
        _day_ref_kw = ["ngày này", "hôm nay", "ngày đang xem", "ngày hôm nay",
                       "ngày đó", "ngày vừa rồi", "ngày hiện tại"]
        if current_view_date and any(k in user_prompt.lower() for k in _day_ref_kw):
            intent["exact_date"] = current_view_date
            # Nếu chưa có periods (tháng) thì extract từ date
            if not intent.get("periods"):
                try:
                    _m = int(current_view_date[5:7])
                    if 1 <= _m <= 12:
                        intent["periods"] = [_m]
                except: pass

        # ── Phát hiện "ngày X VÀ tháng Y" → cần 2 queries riêng ──
        _has_day  = intent.get("exact_date") or intent.get("date_range")
        _has_month = bool(intent.get("periods"))
        if _has_day and _has_month and "tháng" in user_prompt.lower():
            intent["flags"]["is_day_and_month"] = True

        # ----------------------------------------------------------------
        # ----------------------------------------------------------------
        # 2b. Gắn current_view_date vào intent nếu user nói "ngày này/hôm nay"
        # ----------------------------------------------------------------
        _day_ref_kw = ["ngày này", "hôm nay", "ngày đang xem", "ngày hôm nay",
                       "ngày đó", "ngày vừa rồi", "ngày hiện tại"]
        if current_view_date and any(k in user_prompt.lower() for k in _day_ref_kw):
            intent["exact_date"] = current_view_date
            if not intent.get("periods"):
                try:
                    _m = int(current_view_date[5:7])
                    if 1 <= _m <= 12:
                        intent["periods"] = [_m]
                except: pass
        # Phát hiện "ngày X VÀ tháng Y" → đánh dấu is_day_and_month
        _has_day   = bool(intent.get("exact_date") or intent.get("date_range"))
        _has_month = bool(intent.get("periods"))
        if _has_day and _has_month and "tháng" in user_prompt.lower():
            intent["flags"]["is_day_and_month"] = True


        # 3. Đọc Schema thực tế từ DB
        # ----------------------------------------------------------------
        schema = get_db_schema()
        if "error" in schema:
            yield DummyChunk(f"⚠️ Không đọc được Schema DB: {schema['error']}")
            return

        # ── Gắn selected_node (cụm đang chọn) vào intent nếu chưa có cluster ──
        if selected_node and selected_node not in ("", "Tat ca", "T\u1ea5t c\u1ea3"):
            if not intent.get("clusters") and intent.get("query_type") not in ("CROSS_SOURCE","BEAM_STATUS","CATALOG","ITEM_SCHEDULE"):
                intent["clusters"] = [selected_node]
                intent["_auto_cluster"] = selected_node


        # ── Gắn selected_node (cụm đang chọn) vào intent nếu chưa có cluster ──
        # selected_node = "Xưởng Dệt", "Kho Beam Weaving", "Kho Sợi Tổng"...
        if selected_node and selected_node not in ("", "Tất cả"):
            # Chuẩn hóa tên cụm để khớp DB
            _node_map = {
                "Xưởng Dệt":         "Xưởng Dệt",
                "Kho Beam Weaving":   "Kho Beam Weaving",
                "Kho Sợi Tổng":       "Kho Sợi Tổng",
                "Kho Sợi Sizing":     "Kho Sợi Sizing",
                "Kho Sợi Weaving":    "Kho Sợi Weaving",
                "Kho Beam Sizing":    "Kho Beam Sizing",
                "Máy Sectional":      "Máy Sectional",
                "Máy Direct":         "Máy Direct",
                "Máy Hồ":             "Máy Hồ",
                "Xưởng Nhuộm":        "Xưởng Nhuộm",
            }
            _node_clean = _node_map.get(selected_node, selected_node)
            # Override: khi câu hỏi đề cập rõ cluster khác → bỏ auto_cluster
            _p_lower = user_prompt.lower()
            _mention_other = any(k in _p_lower for k in [
                "kho beam", "beam weaving", "xưởng dệt",
                "kho sợi", "máy hồ", "sizing"
            ])
            # "xưởng 1/2/3" trong câu hỏi → đây là câu về Xưởng Dệt
            # Detect xưởng 1/2/3 bằng substring (tránh vấn đề Unicode regex)
            _has_xuong = any(f"xưởng {w}" in _p_lower for w in ["1","2","3","dệt"])
            if _has_xuong:
                _mention_other = True
                if not intent.get("clusters"):
                    intent["clusters"] = ["Xưởng Dệt"]
            if _mention_other:
                _node_clean = None  # câu hỏi đã tự xác định cluster rồi
            # Nếu câu hỏi không đề cập cụm khác → tự động lọc theo cụm đang chọn
            if (_node_clean and not intent.get("clusters")
                    and intent.get("query_type") not in ("CROSS_SOURCE", "BEAM_STATUS")):
                intent["clusters"] = [_node_clean]
                intent["_auto_cluster"] = _node_clean

        # ----------------------------------------------------------------
        # 4. Xây dựng SQL queries (thuần Python)
        # ----------------------------------------------------------------
        queries = build_queries(intent, schema, user_prompt.lower())

        if not queries:
            yield DummyChunk("📭 Không xác định được câu hỏi.\n\n" + suggest_questions(intent, schema))
            return

        # ----------------------------------------------------------------
        # 5. Thực thi từng query
        # ----------------------------------------------------------------
        results = []  # list of (label, agg_dict, df_raw)
        qt_now = intent.get("query_type", "SIMPLE")

        for label, sql in queries:
            # Hiển thị cluster đang query nếu tự động detect
            if intent.get("_auto_cluster") and not getattr(process_ai_chat, "_cluster_shown", False):
                yield DummyChunk(f"🏭 **Cụm: {intent['_auto_cluster']}**\n\n")

            # Show cluster context
            _ac = intent.get("_auto_cluster")
            if _ac and label == (queries[0][0] if queries else ""):
                yield DummyChunk(f"\U0001f3ed **C\u1ee5m: {_ac}**\n\n")

            # Hiển thị SQL debug
            sql_preview = sql[:120] + "…" if len(sql) > 120 else sql
            yield DummyChunk(f"> 🔍 `{sql_preview}`\n\n")

            try:
                df_res = execute_query(sql)
            except Exception as sql_err:
                yield DummyChunk(f"❌ Lỗi SQL [{label}]: {str(sql_err)}\n\n")
                continue

            if df_res.empty:
                yield DummyChunk(f"📭 Không có dữ liệu khớp cho **{label}**.\n\n" + suggest_questions(intent, schema) + "\n\n")
                continue

            agg = aggregate_df(df_res.copy(), schema)
            results.append((label, agg, df_res))

            # --- Quyết định cách hiển thị theo loại query ---
            is_machine_list = label == "MachineList"
            is_day_detail = (
                not is_machine_list
                and (
                    label.startswith("Ngay_")
                    or (qt_now == "SIMPLE" and intent.get("exact_date")
                        and not intent.get("flags", {}).get("is_find_machine"))
                    or (qt_now == "SIMPLE"
                        and len(intent.get("periods", [])) == 1
                        and not intent.get("flags", {}).get("is_find_machine"))
                )
            )

            if is_machine_list:
                # Danh sách máy gom theo xưởng
                yield DummyChunk(format_machine_list(agg) + "\n\n")

            elif is_day_detail:
                yield DummyChunk(format_raw_table(label.replace("Ngay_", "📅 Ngày "), agg) + "\n\n")

            # In bảng xu hướng từng ngày cho TREND query
            elif label.startswith("Trend_"):
                n_days_label = intent.get("n_days_trend", 10)
                yield DummyChunk(format_trend_table(label, agg, n_days_label) + "\n\n")

            elif label == "Danh sách mã hàng":
                _df = df_res
                _is_top = intent.get("flags", {}).get("is_top_catalog", False)
                _top_n  = intent.get("flags", {}).get("top_n", 5)
                _asc    = intent.get("flags", {}).get("is_top_asc", False)

                if _df is not None and not _df.empty:
                    _by_loc = {}
                    for _, row in _df.iterrows():
                        _loc  = str(row.get("sub_location", ""))
                        _item = str(row.get("item_id", "")).strip()
                        _kg   = float(row.get("total_kg", 0) or 0)
                        _nm   = int(row.get("n_machines", 0) or 0)
                        if _item and _item not in ("", "0", "nan"):
                            _by_loc.setdefault(_loc, []).append((_item, _kg, _nm))

                    # Tạo label thời gian
                    _t_dr = intent.get("date_range")
                    _t_ex = intent.get("exact_date")
                    _t_pr = intent.get("periods", [])
                    _t_lbl = ""
                    if _t_ex:
                        _t_lbl = f" ngày {_t_ex[8:]}/{_t_ex[5:7]}/{_t_ex[:4]}"
                    elif _t_dr:
                        if intent.get("flags", {}).get("is_full_year"):
                            _t_lbl = f" năm {_t_dr['from'][:4]}"
                        elif intent.get("flags", {}).get("quarter"):
                            _t_lbl = f" quý {intent['flags']['quarter']}/{_t_dr['from'][:4]}"
                        else:
                            _t_lbl = f" từ {_t_dr['from']} đến {_t_dr['to']}"
                    elif _t_pr:
                        _t_lbl = " tháng " + ", ".join(str(m) for m in sorted(_t_pr))

                    if _is_top and _by_loc:
                        # ✅ TOP CATALOG: chỉ trả lời trực tiếp, không liệt kê hết
                        _rank_word = "ít nhất" if _asc else "nhiều nhất"
                        out_top = []
                        for _loc in sorted(_by_loc):
                            _sorted = sorted(_by_loc[_loc], key=lambda x: x[1], reverse=not _asc)
                            _top = _sorted[:_top_n]
                            _total_ma = len(_by_loc[_loc])
                            if _top:
                                # Câu trả lời trực tiếp cho mã #1
                                _best = _top[0]
                                _medal = "🥇"
                                out_top.append(
                                    f"{_medal} **{_loc}** — mã hàng chạy **{_rank_word}** "
                                    f"trong {_total_ma} mã{_t_lbl}:\n\n"
                                    f"**{_best[0]}** — {_best[1]:,.0f} Kg | {_best[2]} máy\n\n"
                                )
                                if len(_top) > 1:
                                    out_top.append(f"Top {len(_top)} mã {_rank_word}:\n")
                                    medals = ["🥇","🥈","🥉"] + ["4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"]
                                    for _ri, (_item, _kg, _nm) in enumerate(_top):
                                        _med = medals[_ri] if _ri < len(medals) else f"{_ri+1}."
                                        out_top.append(f"{_med} **{_item}**: {_kg:,.0f} Kg | {_nm} máy\n")
                        if out_top:
                            yield DummyChunk("".join(out_top) + "\n")
                            # Trả lời trực tiếp xong → skip commentary thừa
                            intent["flags"]["_catalog_answered"] = True
                            return

                    else:
                        # CATALOG đầy đủ: liệt kê theo xưởng
                        out = []
                        for _loc in sorted(_by_loc):
                            out.append(f"**{_loc}** — {len(_by_loc[_loc])} mã hàng\n")
                            for _item, _kg, _nm in _by_loc[_loc][:50]:
                                _kg_str = f"{_kg:,.0f} Kg" if _kg > 0 else ""
                                _nm_str = f" | {_nm} máy" if _nm > 0 else ""
                                out.append(f"  - {_item}{(': ' + _kg_str) if _kg_str else ''}{_nm_str}\n")
                        if out:
                            _total_items = sum(len(v) for v in _by_loc.values())
                            _summary_parts = [f"{_loc}: **{len(items)} mã**" for _loc, items in sorted(_by_loc.items())]
                            if len(_by_loc) == 1:
                                _loc1 = list(_by_loc.keys())[0]
                                _n1   = len(_by_loc[_loc1])
                                _summary = f"✅ **{_loc1} chạy {_n1} mã hàng{_t_lbl}**\n"
                            else:
                                _summary = f"📋 **Tổng: {_total_items} mã hàng{_t_lbl}** ({' | '.join(_summary_parts)})\n"
                            yield DummyChunk(_summary + "".join(out) + "\n")

            elif label == "Lịch chạy":
                # ITEM_SCHEDULE: lịch chạy theo ngày + máy
                _sched_df = df_res
                if _sched_df is not None and not _sched_df.empty:
                    _item_name = (intent.get("item_kw") or "").upper()
                    _dates = sorted(_sched_df["date"].unique())
                    _n_dates = len(_dates)
                    # Đếm lượt máy-ngày (1 máy chạy 1 ngày = 1 lượt)
                    _n_turns = len(_sched_df)
                    _total_kg = float(_sched_df["total_kg"].sum())
                    # Group by date
                    _by_date = {}
                    for _, row in _sched_df.iterrows():
                        _d   = str(row["date"])
                        _loc = str(row.get("sub_location",""))
                        _mac = str(row.get("ten_may_norm",""))
                        _kg  = float(row.get("total_kg", 0) or 0)
                        _ef  = float(row.get("avg_eff",   0) or 0)
                        _by_date.setdefault(_d, []).append((_loc, _mac, _kg, _ef))
                    out = [
                        f"📅 **{_item_name}** — chạy **{_n_dates} ngày** | "
                        f"**{_n_turns} lượt máy** | Tổng: **{_total_kg:,.0f} Kg**\n\n"
                    ]
                    for _d in _dates:
                        _d_fmt = f"{_d[8:]}/{_d[5:7]}/{_d[:4]}"
                        _entries = _by_date[_d]
                        _day_kg  = sum(e[2] for e in _entries)
                        # Group entries by workshop for compact display
                        _by_loc2 = {}
                        for _loc, _mac, _kg, _ef in _entries:
                            _by_loc2.setdefault(_loc, []).append((_mac, _kg, _ef))
                        _locs_str = " | ".join(
                            f"{_l.replace('Weaving ','W')}: máy {', '.join(m for m,_,__ in sorted(_by_loc2[_l], key=lambda x: int(x[0]) if x[0].isdigit() else 0))}"
                            for _l in sorted(_by_loc2)
                        )
                        out.append(f"- **{_d_fmt}** ({_day_kg:,.0f} Kg) — {_locs_str}\n")
                    yield DummyChunk("".join(out))
                    return  # không cần thêm gì nữa

        if not results:
            yield DummyChunk("📭 *Không tìm thấy dữ liệu nào khớp yêu cầu.*\n\n" + suggest_questions(intent, schema))
            return

        # ----------------------------------------------------------------
        # 6. Hiển thị số liệu tổng hợp
        # ----------------------------------------------------------------
        # results = list of (label, agg, df_raw)
        # - MachineList: đã được in bởi format_machine_list ở trên → bỏ qua
        # - Trend_: đã được in bởi format_trend_table → bỏ qua
        # - Ngay_: đã được in bởi format_raw_table → chỉ in summary agg
        for label, agg, _ in results:
            skip = label.startswith("Trend_") or label == "MachineList" or label == "Danh sách mã hàng" or label == "Lịch chạy"  # Lịch chạy rendered below
            if not skip:
                block = format_agg_for_display(label.replace("Ngay_", ""), agg)
                if block.strip():
                    yield DummyChunk(f"\n{block}\n")

        # ================================================================
        # 7. KIẾN TRÚC MỚI: Python viết FACT, AI chỉ viết COMMENTARY
        #    → Đảm bảo số liệu 100% chính xác, AI không thể bịa số
        # ================================================================
        # Các query type đã hiển thị dữ liệu đầy đủ → không cần AI commentary
        _no_commentary_types = {"CATALOG", "ITEM_SCHEDULE"}
        _no_commentary_flags = {"is_catalog", "is_item_schedule"}
        if (intent.get("query_type") in _no_commentary_types
                or any(intent.get("flags", {}).get(f) for f in _no_commentary_flags)):
            return

        # MachineList: chỉ cần commentary khi user hỏi xếp hạng (cao nhất/thấp nhất)
        # Nếu chỉ hỏi "có những máy nào" → data đã đủ, không cần AI
        _is_mlist = any(lbl == "MachineList" for lbl, _, __ in results)
        if _is_mlist:
            _ask_rank_now = any(k in user_prompt.lower() for k in [
                "cao nhất","thấp nhất","tốt nhất","kém nhất","nhiều nhất",
                "ít nhất","xếp hạng","so sánh","hơn","vượt"
            ])
            if not _ask_rank_now:
                return  # dữ liệu đã đủ, không bịa thêm

        is_comparison   = intent.get("flags", {}).get("is_comparison") or len(results) > 1
        is_trend        = any(lbl.startswith("Trend_") for lbl, _, __ in results)
        is_multi_period = intent.get("query_type") == "MULTI_PERIOD"
        is_cross        = intent.get("query_type") == "CROSS_SOURCE"
        is_machine_query = any(lbl == "MachineList" for lbl, _, __ in results)

        # --- 7a. Python tạo FACT BLOCK (số liệu chính xác) ---
        fact_lines = []
        commentary_context = []   # context ngắn gọn cho AI làm analysis

        for lbl, agg, _ in results:
            if lbl == "MachineList":
                continue
            if lbl.startswith("Trend_"):
                continue  # Trend đã hiện qua format_trend_table, không thêm vào fact_lines

            total_kg  = agg.get("total_kg", 0)
            avg_eff   = agg.get("avg_eff", 0)
            by_month  = agg.get("by_month", {})
            by_loc    = agg.get("by_location", {})
            items     = agg.get("items", [])
            by_mac    = agg.get("by_machine", {})

            prefix = f"**{lbl}** — " if lbl not in ("Kết quả", "Chi tiết tháng",
                                                      "Chi tiết ngày", "Tổng hợp", "Tổng hợp đa tháng") else ""
            _ask_eff_q = any(k in user_prompt.lower() for k in
                              ["hiệu suất", "hiệu quả", "efficiency", " hs "])
            if _ask_eff_q:
                # ✅ FIX: User hỏi hiệu suất → đưa eff lên đầu, sản lượng phụ
                if avg_eff > 0:
                    fact_lines.append(f"{prefix}**Hiệu suất trung bình: {avg_eff}%**")
                else:
                    fact_lines.append(f"{prefix}⚠️ Không tìm thấy dữ liệu hiệu suất trong khoảng này.")
                if total_kg > 0:
                    fact_lines.append(f"Sản lượng kèm theo: {total_kg:,.0f} Kg")
            else:
                if total_kg > 0:
                    fact_lines.append(f"{prefix}**Tổng sản lượng: {total_kg:,.2f} Kg**")
                if avg_eff > 0:
                    fact_lines.append(f"Hiệu suất trung bình: **{avg_eff}%**")

            # Tháng đỉnh / đáy — nếu hỏi hiệu suất thì sort theo eff, không phải kg
            _ask_eff_q2 = any(k in user_prompt.lower() for k in
                              ["hiệu suất", "hiệu quả", "efficiency", " hs "])
            if by_month:
                months_with_data = [(k, v) for k, v in by_month.items() if v["kg"] > 0]
                months_with_eff  = [(k, v) for k, v in by_month.items() if v.get("eff", 0) > 0]
                if _ask_eff_q2 and months_with_eff:
                    # Hiệu suất focus: sort by eff
                    best_e  = max(months_with_eff, key=lambda x: x[1]["eff"])
                    worst_e = min(months_with_eff, key=lambda x: x[1]["eff"])
                    m_b = f"{best_e[0][5:7]}/{best_e[0][:4]}"
                    m_w = f"{worst_e[0][5:7]}/{worst_e[0][:4]}"
                    fact_lines.append(f"Tháng HS cao nhất: **{m_b}** — {best_e[1]['eff']}%"
                                      + (f" ({best_e[1]['kg']:,.0f} Kg)" if best_e[1]["kg"] > 0 else ""))
                    if worst_e[0] != best_e[0]:
                        fact_lines.append(f"Tháng HS thấp nhất: **{m_w}** — {worst_e[1]['eff']}%"
                                          + (f" ({worst_e[1]['kg']:,.0f} Kg)" if worst_e[1]["kg"] > 0 else ""))
                elif months_with_data:
                    best  = max(months_with_data, key=lambda x: x[1]["kg"])
                    worst = min(months_with_data, key=lambda x: x[1]["kg"])
                    m_b   = f"{best[0][5:7]}/{best[0][:4]}"
                    m_w   = f"{worst[0][5:7]}/{worst[0][:4]}"
                    fact_lines.append(f"Tháng cao nhất: **{m_b}** — {best[1]['kg']:,.2f} Kg"
                                      + (f" | HS: {best[1]['eff']}%" if best[1]["eff"] > 0 else ""))
                    if worst[0] != best[0]:
                        fact_lines.append(f"Tháng thấp nhất: **{m_w}** — {worst[1]['kg']:,.2f} Kg"
                                          + (f" | HS: {worst[1]['eff']}%" if worst[1]["eff"] > 0 else ""))
                    # Context ngắn cho AI
                    commentary_context.append(
                        f"SL tháng: " + ", ".join(
                            f"{k[5:7]}/{k[:4]}={v['kg']:,.0f}Kg(HS:{v['eff']}%)"
                            for k, v in sorted(months_with_data)
                        )
                    )

            # Theo xưởng (chỉ hiện xưởng có sản lượng > 0)
            if by_loc and len(by_loc) > 1:
                loc_parts = [f"{loc}={d['kg']:,.0f}Kg" for loc, d in by_loc.items() if d["kg"] > 0]
                if loc_parts:
                    fact_lines.append("Theo xưởng: " + " | ".join(loc_parts))
                    commentary_context.append("Xưởng: " + ", ".join(loc_parts))

            # Top 3 máy cho machine queries
            if by_mac and is_machine_query:
                top3 = sorted([(k, v) for k, v in by_mac.items() if v["kg"] > 0],
                              key=lambda x: x[1]["kg"], reverse=True)[:3]
                if top3:
                    t3_str = ", ".join(f"Máy {k}={v['kg']:,.0f}Kg" for k, v in top3)
                    fact_lines.append(f"Top 3 máy: {t3_str}")
                    commentary_context.append(f"Top3: {t3_str}")

            # Mã hàng
            if items:
                fact_lines.append(f"Mã hàng: {', '.join(str(i) for i in items[:5])}")

        # Trend table text for AI (chỉ dùng trong mode TREND)
        trend_table = ""
        if is_trend:
            trend_label_obj = next(
                ((l, a) for l, a, _ in results if l.startswith("Trend_")), None
            )
            if trend_label_obj:
                _tl, trend_agg = trend_label_obj
                rows = trend_agg.get("raw_rows", [])
                if rows:
                    from collections import defaultdict
                    by_day2 = defaultdict(lambda: {"kg": 0.0, "eff_sum": 0.0, "eff_n": 0, "items": set()})
                    for r in rows:
                        d = str(r.get("date", ""))[:10]
                        if not d:
                            continue
                        for k, v in r.items():
                            kl = k.lower()
                            if any(x in kl for x in ["qty", "quantity", "total_kg"]):
                                by_day2[d]["kg"] += _to_float(v)
                            if any(x in kl for x in ["eff", "hiệu suất", "효율"]):
                                fv = _to_float(v)
                                if fv > 0:
                                    by_day2[d]["eff_sum"] += fv
                                    by_day2[d]["eff_n"]   += 1
                        item_k = next((k for k in r if "item" in k.lower()), None)
                        if item_k and str(r.get(item_k, "")) not in ("nan", "", "None"):
                            by_day2[d]["items"].add(str(r[item_k]))
                    lines2 = ["Dữ liệu từng ngày:"]
                    for d in sorted(by_day2.keys()):
                        v2 = by_day2[d]
                        kg2 = f"{v2['kg']:,.2f}Kg" if v2["kg"] > 0 else "—"
                        ef2 = f"{round(v2['eff_sum']/v2['eff_n'],1)}%" if v2["eff_n"] > 0 else "—"
                        it2 = ",".join(list(v2["items"])[:2]) if v2["items"] else "—"
                        lines2.append(f"  {d}: SL={kg2}|HS={ef2}|Hàng={it2}")
                    trend_table = "\n".join(lines2)
                    # Also store by_day in trend_agg for commentary
                    trend_agg["by_day"] = dict(by_day2)

        # --- 7b. Yield FACT BLOCK (Python đảm bảo đúng) ---
        # For trend queries, fact_lines are already shown in format_trend_table — skip
        if fact_lines and not is_trend:
            yield DummyChunk("\n\n---\n\n" + "\n".join(fact_lines) + "\n\n")

        # --- 7c. AI chỉ viết COMMENTARY (không được nhắc lại số tổng) ---
        context_str = "\n".join(commentary_context) if commentary_context else "Xem dữ liệu ở trên."

        if is_machine_query:
            # Python tự tính trả lời — không để AI đoán từ danh sách 30+ máy
            _p = user_prompt.lower()
            # Chỉ show ranking khi user hỏi rõ cao nhất/thấp nhất
            _ask_rank = any(k in _p for k in ["cao nhất","tốt nhất","thấp nhất","kém nhất",
                                               "ít nhất","nhiều nhất","lớn nhất","yếu nhất",
                                               "xếp hạng","so sánh"])
            _ask_hs  = _ask_rank and any(k in _p for k in ["hiệu suất","hiệu quả","hs","efficiency","tốt nhất"])
            _ask_sl  = _ask_rank and any(k in _p for k in ["sản lượng","sản xuất","nhiều nhất","kg","lớn nhất"])
            _ask_low = any(k in _p for k in ["thấp nhất","kém nhất","ít nhất","yếu nhất"])

            # Gom by_machine — KEY = (xưởng, máy) để tránh cộng nhầm máy 38 xưởng 1 + xưởng 2
            _mac_data = {}  # key: "Weaving 1 — Máy 38", value: {kg, eff}
            for _lbl, _agg, _ in results:
                # Xác định xưởng từ label hoặc by_location
                _loc_name = ""
                _by_loc = _agg.get("by_location", {})
                if len(_by_loc) == 1:
                    _loc_name = list(_by_loc.keys())[0]
                for _mac, _d in _agg.get("by_machine", {}).items():
                    if _d["kg"] > 0:
                        _key = f"{_loc_name} — Máy {_mac}" if _loc_name else f"Máy {_mac}"
                        _mac_data[_key] = _d

            _direct_answer = ""
            if _mac_data and _ask_rank:
                _with_eff = [(m, d) for m, d in _mac_data.items() if d["eff"] > 0]
                _with_kg  = [(m, d) for m, d in _mac_data.items() if d["kg"] > 0]

                if _ask_hs and _with_eff:
                    _best_hs  = max(_with_eff, key=lambda x: x[1]["eff"])
                    _worst_hs = min(_with_eff, key=lambda x: x[1]["eff"])
                    if _ask_low:
                        _direct_answer = (
                            f"🏆 **{_worst_hs[0]}** hiệu suất **thấp nhất**: {_worst_hs[1]['eff']}% "
                            f"({_worst_hs[1]['kg']:,.0f} Kg)\n"
                            f"_(Cao nhất: {_best_hs[0]} — {_best_hs[1]['eff']}%)_"
                        )
                    else:
                        _direct_answer = (
                            f"🏆 **{_best_hs[0]}** hiệu suất **cao nhất**: {_best_hs[1]['eff']}% "
                            f"({_best_hs[1]['kg']:,.0f} Kg)\n"
                            f"_(Thấp nhất: {_worst_hs[0]} — {_worst_hs[1]['eff']}%)_"
                        )
                elif _with_kg:
                    _best_kg  = max(_with_kg,  key=lambda x: x[1]["kg"])
                    _worst_kg = min(_with_kg,  key=lambda x: x[1]["kg"])
                    if _ask_low:
                        _direct_answer = (
                            f"🏆 **{_worst_kg[0]}** sản lượng **thấp nhất**: {_worst_kg[1]['kg']:,.0f} Kg\n"
                            f"_(Cao nhất: {_best_kg[0]} — {_best_kg[1]['kg']:,.0f} Kg)_"
                        )
                    else:
                        _direct_answer = (
                            f"🏆 **{_best_kg[0]}** sản lượng **cao nhất**: {_best_kg[1]['kg']:,.0f} Kg\n"
                            f"_(Thấp nhất: {_worst_kg[0]} — {_worst_kg[1]['kg']:,.0f} Kg)_"
                        )

            if _direct_answer:
                yield DummyChunk(_direct_answer)
                return  # Không cần AI — câu trả lời đã chính xác 100%

            # Fallback: AI nhận xét tổng quát nếu không detect được trọng tâm
            commentary_prompt = f"""Bạn là quản đốc nhà máy dệt, trả lời ngắn gọn đúng câu hỏi.

CÂU HỎI: "{user_prompt}"
BỐI CẢNH: {context_str}

VIẾT 1-2 CÂU trả lời thẳng vào câu hỏi. KHÔNG lan man. Tiếng Việt."""

        elif is_trend:
            n_days = intent.get("n_days_trend", 10)
            # Build a compact facts string for AI
            _trend_facts = ""
            if trend_agg:
                _days_with = [d for d,v in trend_agg.get("by_day",{}).items() if v.get("kg",0)>0]
                _kgs = [trend_agg["by_day"][d]["kg"] for d in _days_with]
                if _kgs:
                    _trend_facts = (
                        f"Có {len(_days_with)} ngày SL trong {n_days} ngày. "
                        f"Ngày cao nhất: {max(_kgs):,.0f}kg. Ngày thấp nhất: {min(_kgs):,.0f}kg. "
                        f"Tổng: {sum(_kgs):,.0f}kg."
                    )

            commentary_prompt = (
                f"Bạn là quản đốc xưởng dệt. Câu hỏi: \"{user_prompt}\"\n\n"
                f"Số liệu thực tế: {_trend_facts}\n\n"
                f"Viết 2 câu nhận xét tự nhiên như đang báo cáo miệng cho sếp: "
                f"(1) nhận xét xu hướng tăng/giảm/ổn, "
                f"(2) kết luận ngắn gọn cần làm gì hoặc điểm đáng chú ý. "
                f"Không liệt kê lại số. Tiếng Việt. Tối đa 60 từ."
            )

        elif is_comparison:
            # Python tính cứng toàn bộ bảng xếp hạng — hỗ trợ N máy (không chỉ 2)
            _all_machines = [
                (lbl, agg["total_kg"], agg["avg_eff"])
                for lbl, agg, _ in results
                if agg["total_kg"] > 0
            ]
            _cmp_facts = []
            if _all_machines:
                # Xếp hạng hiệu suất
                _ranked_eff = sorted(
                    [(lbl, ef) for lbl, _, ef in _all_machines if ef > 0],
                    key=lambda x: x[1], reverse=True
                )
                # Xếp hạng sản lượng
                _ranked_kg = sorted(_all_machines, key=lambda x: x[1], reverse=True)

                if _ranked_eff:
                    _best_eff_lbl, _best_eff_val = _ranked_eff[0]
                    _worst_eff_lbl, _worst_eff_val = _ranked_eff[-1]
                    _cmp_facts.append(
                        f"- HIỆU SUẤT CAO NHẤT: {_best_eff_lbl} = {_best_eff_val}%"
                    )
                    _cmp_facts.append(
                        f"- HIỆU SUẤT THẤP NHẤT: {_worst_eff_lbl} = {_worst_eff_val}%"
                    )
                    if len(_ranked_eff) > 2:
                        _rank_str = " > ".join(f"{l}({e}%)" for l, e in _ranked_eff)
                        _cmp_facts.append(f"- Xếp hạng HS đầy đủ: {_rank_str}")

                if _ranked_kg:
                    _best_kg_lbl, _best_kg_val, _ = _ranked_kg[0]
                    _cmp_facts.append(
                        f"- SẢN LƯỢNG CAO NHẤT: {_best_kg_lbl} = {_best_kg_val:,.2f} Kg"
                    )

            _cmp_facts_str = "\n".join(_cmp_facts) if _cmp_facts else context_str
            _n = len(_all_machines)
            _ask_sl  = any(k in user_prompt.lower() for k in ["sản lượng","sản xuất","cao nhất","nhiều nhất","kg","tổng"])
            _ask_eff = any(k in user_prompt.lower() for k in ["hiệu suất","hiệu quả","efficiency"])
            _focus   = "SẢN LƯỢNG" if (_ask_sl and not _ask_eff) else "HIỆU SUẤT" if _ask_eff else "SẢN LƯỢNG và HIỆU SUẤT"
            commentary_prompt = (
                f"Bạn là quản đốc tóm tắt kết quả so sánh {_n} máy cho Sếp.\n\n"
                f"KẾT QUẢ ĐÃ TÍNH CHÍNH XÁC (bắt buộc dùng đúng, KHÔNG đảo ngược):\n{_cmp_facts_str}\n\n"
                f'CÂU HỎI: "{user_prompt}"\n'
                f"NGƯỜI DÙNG HỎI VỀ: {_focus}\n\n"
                f"VIẾT {'2' if _n <= 2 else '3'} CÂU nhận xét tập trung vào {_focus}:\n"
                "1. May nao cao nhat va thap nhat. 2. Nhan xet chenh lech. 3. Goi y cai thien.\n"
                + ("2. Nhan xet chenh lech va goi y.\n" if _n <= 2 else
                   "2. Xep hang day du.\n" + "3. Goi y cai thien may kem nhat.\n")
                + "TUYET DOI khong bia, khong dao nguoc ket qua. Tieng Viet. Toi da 80 tu."
            )
        elif is_multi_period or by_month:
            # --- Python tính FACTS cứng trước khi truyền cho AI ---
            _month_facts = []
            if by_month:
                _months_with_data = [(k, v) for k, v in by_month.items() if v["kg"] > 0]
                if _months_with_data:
                    _best_sl  = max(_months_with_data, key=lambda x: x[1]["kg"])
                    _worst_sl = min(_months_with_data, key=lambda x: x[1]["kg"])
                    _months_with_eff = [(k, v) for k, v in _months_with_data if v["eff"] > 0]
                    _best_eff  = max(_months_with_eff, key=lambda x: x[1]["eff"]) if _months_with_eff else None
                    _worst_eff = min(_months_with_eff, key=lambda x: x[1]["eff"]) if _months_with_eff else None

                    # Xu hướng SL: so sánh nửa đầu và nửa cuối
                    sorted_m = sorted(_months_with_data)
                    if len(sorted_m) >= 3:
                        mid = len(sorted_m) // 2
                        avg_first = sum(v["kg"] for _, v in sorted_m[:mid]) / mid
                        avg_last  = sum(v["kg"] for _, v in sorted_m[mid:]) / max(len(sorted_m) - mid, 1)
                        if avg_last > avg_first * 1.05:
                            trend_str = "tăng dần"
                        elif avg_last < avg_first * 0.95:
                            trend_str = "giảm dần"
                        else:
                            trend_str = "tương đối ổn định"
                    else:
                        trend_str = "chưa đủ dữ liệu để xác định xu hướng"

                    m_best_sl_label  = f"{_best_sl[0][5:7]}/{_best_sl[0][:4]}"
                    m_worst_sl_label = f"{_worst_sl[0][5:7]}/{_worst_sl[0][:4]}"
                    _month_facts.append(f"- Xu hướng sản lượng: {trend_str}")
                    _month_facts.append(f"- Tháng SL cao nhất: {m_best_sl_label} ({_best_sl[1]['kg']:,.0f} Kg)")
                    _month_facts.append(f"- Tháng SL thấp nhất: {m_worst_sl_label} ({_worst_sl[1]['kg']:,.0f} Kg)")
                    if _best_eff:
                        m_best_eff_label = f"{_best_eff[0][5:7]}/{_best_eff[0][:4]}"
                        _month_facts.append(f"- Tháng HS cao nhất: {m_best_eff_label} ({_best_eff[1]['eff']}%)")
                    if _worst_eff:
                        m_worst_eff_label = f"{_worst_eff[0][5:7]}/{_worst_eff[0][:4]}"
                        _month_facts.append(f"- Tháng HS thấp nhất: {m_worst_eff_label} ({_worst_eff[1]['eff']}%)")

            facts_block = "\n".join(_month_facts) if _month_facts else context_str

            commentary_prompt = f"""Bạn là quản đốc phân xưởng dệt đang báo cáo cho sếp.

SỰ KIỆN THỰC TẾ (không được thay đổi hay bịa thêm):
{facts_block}

Câu hỏi: "{user_prompt}"

Viết 2-3 câu nhận xét tự nhiên như người trong nghề nói chuyện:
- Không liệt kê lại số liệu, chỉ diễn giải ý nghĩa.
- Chỉ ra điểm đáng chú ý nhất (tháng bất thường, xu hướng tăng/giảm rõ).
- Kết bằng nhận xét chung hoặc lưu ý thực tế.
Viết liền mạch, không dùng gạch đầu dòng. Tiếng Việt tự nhiên. Tối đa 80 từ."""

        else:
            commentary_prompt = f"""Bạn là quản đốc phân xưởng dệt. Dưới đây là dữ liệu sản xuất thực tế.

DỮ LIỆU: {context_str}

Câu hỏi của sếp: "{user_prompt}"

Viết đúng 2 câu nhận xét như người trong nghề:
- Câu 1: đánh giá kết quả (tốt / chấp nhận / cần cải thiện) kèm lý do cụ thể.
- Câu 2: gợi ý hoặc lưu ý ngắn nếu cần (nếu kết quả tốt thì nhận xét điểm đáng chú ý).
KHÔNG đọc lại số liệu đã hiển thị. Viết tự nhiên như nói chuyện, không dùng gạch đầu dòng. Tiếng Việt. Tối đa 50 từ."""

        # Lấy biến by_month từ result đầu tiên nếu cần
        by_month = results[0][1].get("by_month", {}) if results else {}

        # ✅ FIX TIMEOUT: Thử AI trước (8 giây), nếu lỗi → rule-based fallback
        _ai_comment = ""
        try:
            final_stream = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": commentary_prompt}],
                stream=True,
                temperature=0.4,
                max_tokens=120,
                timeout=8,   # giảm xuống 8 giây
            )
            for chunk in final_stream:
                _c = chunk.choices[0].delta.content
                if _c:
                    _ai_comment += _c
                    yield DummyChunk(_c)
        except Exception:
            _ai_comment = ""  # sẽ dùng fallback bên dưới

        # ── Fallback rule-based: luôn có nhận xét kể cả khi AI timeout ──
        if not _ai_comment.strip():
            _fb = _generate_fallback_comment(intent, results, user_prompt)
            if _fb:
                yield DummyChunk(_fb)

    except Exception as e:
        err_str = str(e)
        if "timed out" not in err_str.lower() and "timeout" not in err_str.lower():
            yield DummyChunk(f"\n⚠️ Lỗi: {err_str[:120]}")
        # Fallback ngay cả khi outer exception
        try:
            _fb2 = _generate_fallback_comment(intent, results, user_prompt)
            if _fb2:
                yield DummyChunk(_fb2)
        except Exception:
            pass