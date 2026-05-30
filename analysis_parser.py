"""
analysis_parser.py — Đọc sheet 'analysis', parse ghi chú lỗi → điền thời gian vào bảng
Công thức:
  - 'JQ 30' → 30 phút
  - '14H35-15H50' → 15h50 - 14h35 = 75 phút
  - '6H-8H10+60' → 130 phút (130min + 60min)
"""

import re
import pandas as pd
import openpyxl
from copy import copy
from datetime import datetime

# ── Tên cột trong bảng (col index → tên field) ─────────────────────────────
COL_IDX = {
    0:  'so_may',
    1:  'no_order',
    2:  'mc_stop',
    3:  'thay_go',
    4:  'loi_go_jq',          # MACHINE REPAIR: lỗi go+JQ
    5:  'ccsn',
    6:  'brake',
    7:  'cutter',
    8:  'loi',
    9:  'kiem_dai',
    10: 'sua_may_khac',
    11: 'tyming_tren',         # TYMING UP
    12: 'tyming_duoi',         # TYMING DOWN
    13: 'pull_beam_tren',      # PULL BEAM UP
    14: 'pull_beam_duoi',      # PULL BEAM DOWN
    15: 'cho',                 # CHỜ
    16: 'no_beam_wea',         # NO BEAM
    17: 'no_beam_other',
    18: 'beam_error',          # BEAM/BEAM YARN
    19: 'beam_xac_nhan',
    20: 'beam_roi',
    21: 'beam_broke',
    22: 'clean_mc',            # CLEAN
    23: 'clean_feeder',
    24: 'clean_tb',
    25: 'clean_other',
    26: 'sample_cms',          # SAMPLE
    27: 'sample_tb',
    28: 'sample_sample',
    29: 'changed_cp',          # CHANGED PRODUCT
    30: 'changed_kfile',
    31: 'changed_pm',
    32: 'changed_broke',
    33: 'k_co_soi',
}

# ── Bảng phân loại từ khóa → col ───────────────────────────────────────────
# Nguồn: NOTE_GHI_CHÚ_DOWNTIMES.xlsx - sheet analysis
# Dạng đầy đủ VÀ viết tắt chữ cái đầu đều được nhận dạng
# Từ khóa dài hơn được match trước (sorted by length desc)

KEYWORD_MAP = [
    # ── Col 0: Lỗi vải/sợi ────────────────────────────────────────
    ('sợi ngang bắn ngược', 'loi'),
    ('border nổi p',        'loi'),
    ('nổi p ngược',         'loi'),
    ('lỗi sợi ngang',       'loi'),
    ('noi soi ngang',       'loi'),
    ('rách khăn',           'loi'),
    ('rach khan',           'loi'),
    ('hở khăn',             'loi'),
    ('ho khan',             'loi'),
    ('sọc ngang',           'loi'),
    ('soc ngang',           'loi'),
    ('thiếu nền',           'loi'),
    ('nổi sợi dty',         'loi'),
    ('dùn p',               'loi'),
    ('dun p',               'loi'),
    ('nổi p',               'loi'),
    ('noi p',               'loi'),
    ('lỗi',                 'loi'),
    ('loi',                 'loi'),
    # ── Col 1: Sửa máy ────────────────────────────────────────────
    ('quấn lăn gai',        'sua_may_khac'),
    ('quan lan gai',        'sua_may_khac'),
    ('biên phế test board', 'sua_may_khac'),
    ('bien phe test',       'sua_may_khac'),
    ('đóng máy sửa',        'sua_may_khac'),
    ('dong may sua',        'sua_may_khac'),
    ('chuyên gia dừng',     'sua_may_khac'),
    ('chuyen gia',          'sua_may_khac'),
    ('lực căng p',          'sua_may_khac'),
    ('luc cang',            'sua_may_khac'),
    ('đóng máy chờ chỉ thị','sua_may_khac'),
    ('lỗi máy quấn vải',    'sua_may_khac'),
    ('lỗi máy cấp sợi',     'sua_may_khac'),
    ('gãy thang ngang',     'sua_may_khac'),
    ('chờ sợi ngang',       'sua_may_khac'),
    ('thay lò xo',          'sua_may_khac'),
    ('thay lo xo',          'sua_may_khac'),
    ('nổ tụ điện',          'sua_may_khac'),
    ('no tu dien',          'sua_may_khac'),
    ('dùn p lăn gai',       'sua_may_khac'),
    ('leno không đan',      'sua_may_khac'),
    ('lỗi sập nguồn',       'sua_may_khac'),
    ('sap nguon',           'sua_may_khac'),
    ('bể ống dầu',          'sua_may_khac'),
    ('be ong dau',          'sua_may_khac'),
    ('dùn lamen',           'sua_may_khac'),
    ('dun lamen',           'sua_may_khac'),
    ('sọc lăn gai',         'sua_may_khac'),
    ('dm kt mcs',           'sua_may_khac'),
    ('thay sứt g',          'sua_may_khac'),
    ('thay board',          'sua_may_khac'),
    ('test board',          'sua_may_khac'),
    ('chạy nền',            'sua_may_khac'),
    ('chay nen',            'sua_may_khac'),
    ('móc p',               'sua_may_khac'),
    ('moc p',               'sua_may_khac'),
    ('gãy đũa',             'sua_may_khac'),
    ('gay dua',             'sua_may_khac'),
    ('cảm biến',            'sua_may_khac'),
    ('cam bien',            'sua_may_khac'),
    ('chỉ may',             'sua_may_khac'),
    ('chi may',             'sua_may_khac'),
    ('chi thi',             'sua_may_khac'),
    ('lỗi module',          'sua_may_khac'),
    ('loi module',          'sua_may_khac'),
    ('hư motor',            'sua_may_khac'),
    ('hu motor',            'sua_may_khac'),
    ('xổ beam',             'pull_beam_tren'),
    ('xì hơi',              'sua_may_khac'),
    ('xi hoi',              'sua_may_khac'),
    ('lăn gai',             'sua_may_khac'),
    ('lan gai',             'sua_may_khac'),
    ('leno',                'sua_may_khac'),
    ('haness',              'sua_may_khac'),
    ('sửa',                 'sua_may_khac'),
    ('sua',                 'sua_may_khac'),
    # ── Col 2: PM Kiểm / Changed Product ─────────────────────────
    ('chờ kiểm tra size',   'changed_pm'),
    ('cho kiem size',       'changed_pm'),
    ('pm kiem',             'changed_pm'),
    ('trả đơn hàng',        'changed_cp'),
    ('tra don hang',        'changed_cp'),
    ('chuyển pp+gg',        'changed_cp'),
    ('chờ file',            'changed_kfile'),
    ('cho file',            'changed_kfile'),
    ('ktkm',                'changed_pm'),     # FIX: KTKM = PM kiểm tra
    ('đsp',                 'changed_cp'),
    ('dsp',                 'changed_cp'),
    # ── Col 3: Rối Beam ─────────────────────────────────────────
    ('beam p dư thiếu sợi', 'beam_error'),
    # FIX: Thêm biến thể dấu và tổ hợp GO+LỖI
    ('go + lỗi',             'loi'),       # "GO + LỖI XhYY-ZhWW" → phân loại lỗi
    ('go + loi',             'loi'),
    ('rói',                  'beam_roi'),  # biến thể dấu của 'rối'
    ('roi ',                 'beam_roi'),
    ('mat p',               'beam_roi'),
    ('mất p',               'beam_roi'),
    ('rối',                 'beam_roi'),
    # ── Col 4: Đứt / Broke ──────────────────────────────────────
    ('đứt sợi ngang',       'changed_broke'),  # FIX: đứt sợi ngang → đsn BROKE
    ('dut soi ngang',       'changed_broke'),
    ('xử lí sợi đứt',       'changed_broke'),
    ('xu li soi',           'changed_broke'),
    ('đứt p lăn gai',       'changed_broke'),
    ('đứt liên tục',        'changed_broke'),
    ('dut lien tuc',        'changed_broke'),
    ('đứt p',               'changed_broke'),
    ('dut p',               'changed_broke'),
    ('đứt g',               'changed_broke'),
    ('dsn',                 'changed_broke'),
    ('đứt',                 'changed_broke'),
    ('dut',                 'changed_broke'),
    # ── Col 5: Clean / Vệ sinh ───────────────────────────────────
    ('vệ sinh bụi p',       'clean_other'),
    ('ve sinh bui',         'clean_other'),
    ('chang sizw',          'sample_sample'),
    ('vệ sinh',             'clean_mc'),
    ('ve sinh',             'clean_mc'),
    ('cms',                 'sample_cms'),     # FIX: CMS trong SAMPLE section
    ('feeder',              'clean_feeder'),
    ('tb-cbb',              'clean_tb'),
    # ── Go + JQ ─────────────────────────────────────────────────
    ('lỗi go',              'loi_go_jq'),
    ('loi go',              'loi_go_jq'),
    ('go ',                 'loi_go_jq'),
    (' go',                 'loi_go_jq'),
    ('jq ',                 'loi_go_jq'),
    (' jq',                 'loi_go_jq'),
    # ── Cutter ──────────────────────────────────────────────────
    ('dao cắt',             'cutter'),
    ('dao cat',             'cutter'),
    ('cutter',              'cutter'),
    # ── Tyming / nối sợi ────────────────────────────────────────
    ('tyming',              'tyming_tren'),
    ('nối sợi',             'tyming_tren'),
    # ── Pull Beam / Kéo nền ─────────────────────────────────────
    ('pull beam',           'pull_beam_tren'),
    ('kéo nền',             'pull_beam_tren'),
    ('keo nen',             'pull_beam_tren'),
    ('lên beam',            'pull_beam_tren'),
    ('len beam',            'pull_beam_tren'),
    ('thay beam',           'pull_beam_tren'),
    ('beam p+g',            'pull_beam_tren'),
    ('hbp+g',               'pull_beam_tren'),
    ('hbg',                 'pull_beam_tren'),
    ('hbp ',                'pull_beam_tren'),
    ('kéo beam',            'pull_beam_tren'),
    ('keo beam',            'pull_beam_tren'),
    ('kb',                  'pull_beam_tren'),
    # ── Kiểm đai ────────────────────────────────────────────────
    ('kiểm đai',            'kiem_dai'),
    ('kiếm đai',            'kiem_dai'),
    ('kiem dai',            'kiem_dai'),
    # ── Brake ───────────────────────────────────────────────────
    ('brake',               'brake'),
    ('phanh',               'brake'),
    # ── MC Stop ─────────────────────────────────────────────────
    ('mc stop',             'mc_stop'),
    # ── Thay go ─────────────────────────────────────────────────
    ('thay go',             'thay_go'),
    # ── CCSN ────────────────────────────────────────────────────
    ('máy ccsn',            'ccsn'),
    ('ccsn',                'ccsn'),
    ('feeder-weft',         'ccsn'),
    # ── Chờ ─────────────────────────────────────────────────────
    ('chờ nối',             'cho'),
    ('cho noi',             'cho'),
    ('no beam',             'no_beam_wea'),
    # ── Sample ──────────────────────────────────────────────────
    ('chạy mẫu',            'sample_sample'),
    ('chay mau',            'sample_sample'),
    ('sample',              'sample_sample'),
]

# Sort by keyword length descending (longer match first)
KEYWORD_MAP.sort(key=lambda x: len(x[0]), reverse=True)



def parse_minutes(token: str) -> int:
    """
    Tính số phút từ chuỗi:
    - '14H35-15H50' → 75
    - '6H-8H10' → 130
    - '30' → 30
    - '60+30' → 90
    - '6H-8H10+60' → 130+60 = 190 (chỉ cộng số sau dấu +)
    """
    token = token.strip()
    # 1. Time range: XH[MM]-XH[MM]
    m = re.search(r'(\d+)H(\d+)?\s*[-→>]+\s*(\d+)H(\d+)?', token, re.IGNORECASE)
    if m:
        h1 = int(m.group(1)); m1 = int(m.group(2) or 0)
        h2 = int(m.group(3)); m2 = int(m.group(4) or 0)
        diff = (h2 * 60 + m2) - (h1 * 60 + m1)
        if diff < 0: diff += 1440
        # FIX: chỉ cộng số đứng SAU dấu + (không lấy hết số trong chuỗi)
        # VD: "6H-8H10+60" → +60 ✓ | "HBG 12H25-14H45 LẦN 3" → +0 ✓
        rest = token[m.end():]
        extras = sum(int(n) for n in re.findall(r'(?<=\+)\s*(\d+)', rest))
        return diff + extras
    # 2. Phép cộng rõ ràng: '60+30', '90+20+10'
    if '+' in token:
        nums = [int(n) for n in re.findall(r'\d+', token) if int(n) < 600]
        if nums:
            return sum(nums)
    # 3. Số đơn: 'RỐI 30', 'KTKM 60', 'GO 20'
    # FIX: chỉ lấy số CUỐI CÙNG hợp lệ, không cộng tất cả
    # VD: "ĐỨT SỢI NGANG 25" → 25 ✓, không phải sum của các số khác
    nums = re.findall(r'\d+', token)
    if nums:
        valid = [int(n) for n in nums if 1 <= int(n) < 600]
        if valid:
            return valid[-1]
    return 0


def classify_token(token: str) -> str:
    """Phân loại 1 đoạn lỗi → tên field."""
    t_lower = token.lower().strip()
    for kw, field in KEYWORD_MAP:
        if kw in t_lower:
            return field
    return 'sua_may_khac'  # default: sửa máy khác


# Các token đặc biệt không có số nhưng có thời gian mặc định (phút)
SPECIAL_TOKENS = {
    'mc stop': ('mc_stop', 1440),   # Máy dừng cả ngày = 24h = 1440 phút
}


def parse_note(note: str) -> dict:
    """
    Parse toàn bộ ghi chú → {field: total_minutes}
    Split by ',' and '/'
    """
    result = {}
    parts = re.split(r'[,/]', str(note))
    for part in parts:
        part = part.strip()
        if not part:
            continue

        # ✅ FIX: Kiểm tra token đặc biệt trước (VD: "MC STOP" → mc_stop=1440)
        part_lower = part.lower().strip()
        special_hit = None
        for kw, (field, default_mins) in SPECIAL_TOKENS.items():
            if kw in part_lower:
                special_hit = (field, default_mins)
                break

        mins = parse_minutes(part)

        if special_hit and mins == 0:
            # Token đặc biệt không có số → dùng giá trị mặc định
            field, mins = special_hit
            result[field] = result.get(field, 0) + mins
        elif mins == 0:
            continue
        else:
            field = classify_token(part)
            result[field] = result.get(field, 0) + mins
    return result


def read_analysis_sheet(filepath: str) -> list[dict]:
    """
    Đọc sheet 'analysis' → list of machine records.
    Mỗi record: {so_may, note, parsed: {field: minutes}, existing: {col: value}}

    ✅ FIX: Note của máy N được lưu ở hàng máy N+1 (cột 37 = 'máy N', cột 39 = nội dung note).
    Phải đọc nhãn 'máy N' trước để xây bảng {so_may: note}, sau đó mới gán cho đúng máy.
    """
    df = pd.read_excel(filepath, sheet_name='analysis', header=None)

    # ── BƯỚC 1: Xây bảng note từ cột nhãn 'máy N' (pandas col 36) + nội dung (pandas col 38) ──
    note_map = {}  # {so_may: note_text}
    for i in range(5, df.shape[0]):
        row = df.iloc[i]
        # Cột 37 (openpyxl) = pandas 36 = nhãn 'máy N'
        label = row.iloc[36] if df.shape[1] > 36 else None
        if pd.isna(label) or not str(label).strip():
            continue
        import re as _re
        m = _re.search(r'\d+', str(label))
        if not m:
            continue
        mac_num = int(m.group())
        # Cột 39 (openpyxl) = pandas 38 = nội dung note
        note_text = ''
        for nc in [38, 39, 37, 40]:
            if nc < df.shape[1]:
                v = row.iloc[nc]
                v_str = str(v).strip() if pd.notna(v) else ''
                # Bỏ qua nếu chỉ là số (không phải note thực sự)
                if v_str and not v_str.lstrip('-').replace('.','').isdigit():
                    note_text = v_str
                    break
        if note_text:
            note_map[mac_num] = note_text

    # ── BƯỚC 2: Xây records từ cột so_may bên trái, gán note đúng từ note_map ──
    records = []
    for i in range(5, df.shape[0]):
        row = df.iloc[i]
        so_may_raw = row.iloc[0]
        if pd.isna(so_may_raw):
            continue
        try:
            so_may = int(float(so_may_raw))
        except:
            continue

        # Lấy note đúng cho máy này từ bảng đã xây
        note_raw = note_map.get(so_may, '')

        # Get existing values from data cols 1-33
        existing = {}
        for col_idx in range(1, 34):
            if col_idx < df.shape[1]:
                v = row.iloc[col_idx]
                if pd.notna(v) and str(v).strip() and str(v) != '0':
                    try:
                        existing[col_idx] = float(v)
                    except:
                        existing[col_idx] = str(v)

        parsed = parse_note(note_raw) if note_raw else {}

        records.append({
            'excel_row':  i + 1,  # 1-indexed for openpyxl
            'so_may':     so_may,
            'note':       note_raw,
            'parsed':     parsed,
            'existing':   existing,
        })

    return records


def fill_excel(filepath: str, records: list[dict], output_path: str) -> dict:
    """
    Điền thời gian đã parse vào file Excel gốc → lưu file mới.
    Chỉ điền vào ô trống (không ghi đè dữ liệu đã có).
    """
    # Build field → col index reverse map
    field_to_col = {v: k for k, v in COL_IDX.items()}

    wb = openpyxl.load_workbook(filepath)
    ws = wb['analysis']

    filled = 0
    skipped = 0
    log = []

    for rec in records:
        excel_row = rec['excel_row']  # ✅ FIX: excel_row đã là 1-indexed từ read_analysis_sheet, không cộng thêm nữa
        so_may = rec['so_may']
        parsed = rec['parsed']

        for field, minutes in parsed.items():
            col_idx = field_to_col.get(field)
            if col_idx is None:
                continue

            excel_col = col_idx + 1  # openpyxl 1-indexed

            cell = ws.cell(row=excel_row, column=excel_col)
            current = cell.value

            # Only fill if empty or 0
            if current is None or current == 0 or current == '':
                cell.value = minutes
                filled += 1
                log.append(f"Máy {so_may}: {field} = {minutes}min (từ: {rec['note'][:40]})")
            else:
                skipped += 1

    wb.save(output_path)
    return {'filled': filled, 'skipped': skipped, 'log': log}


def generate_report(filepath: str) -> pd.DataFrame:
    """
    Tạo báo cáo dạng DataFrame từ file gốc + parsed notes.
    Để preview trước khi fill.
    """
    records = read_analysis_sheet(filepath)
    rows = []
    for r in records:
        row = {'Máy': r['so_may'], 'Ghi chú': r['note'][:60] if r['note'] else '—'}
        for field, mins in r['parsed'].items():
            row[field] = mins
        rows.append(row)
    return pd.DataFrame(rows)


if __name__ == '__main__':
    import sys
    fp = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/25-05-2026_제직_3동_RPM자동_변경_완료_analysis_ok.xlsx'

    print('Đọc file:', fp)
    records = read_analysis_sheet(fp)
    print(f'Tìm thấy {len(records)} máy')
    for r in records:
        if r['note']:
            print(f"\nMáy {r['so_may']}: {r['note'][:70]}")
            for field, mins in r['parsed'].items():
                print(f"  → {field}: {mins} phút")