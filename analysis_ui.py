"""
analysis_ui.py — Giao diện nhập thời gian tự động từ sheet analysis
Tích hợp vào app.py hoặc chạy standalone: streamlit run analysis_ui.py
"""

import streamlit as st
import pandas as pd
import os
import tempfile
from datetime import datetime

try:
    from analysis_parser import (
        read_analysis_sheet, fill_excel, generate_report,
        COL_IDX, KEYWORD_MAP
    )
    _parser_ok = True
except ImportError as e:
    _parser_ok = False
    _parser_err = str(e)


def render_analysis_ui():
    """Render phần UI phân tích và điền thời gian."""
    st.markdown("### 🕒 Tự Động Nhập Thời Gian (Sheet Analysis)")

    if not _parser_ok:
        st.error(f"analysis_parser.py chưa tìm thấy: `{_parser_err}`")
        return

    with st.container(border=True):
        st.caption("Đọc file báo cáo ngày → parse ghi chú → điền thời gian vào đúng ô → xuất file Excel")

        # ── Upload ──
        col_u1, col_u2 = st.columns([3, 2])
        with col_u1:
            up_file = st.file_uploader(
                "📎 Tải file báo cáo ngày (.xlsx):",
                type=["xlsx", "xls"],
                key="analysis_upload",
                help="File có sheet 'analysis' với cột ghi chú lỗi bên phải"
            )
        with col_u2:
            if "analysis_file_path" not in st.session_state:
                st.session_state.analysis_file_path = ""
            _saved = st.text_input(
                "Hoặc đường dẫn thư mục tự động quét:",
                value=st.session_state.analysis_file_path,
                key="analysis_folder_input",
                placeholder="VD: Z:\\WEAVING\\2026\\05.2026"
            )
            if _saved != st.session_state.analysis_file_path:
                st.session_state.analysis_file_path = _saved

        if up_file is None:
            st.info("Tải file báo cáo lên để bắt đầu.")
            return

        # Save to temp
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as _tmp:
            _tmp.write(up_file.read())
            _tmp_path = _tmp.name

        # ── Parse preview ──
        try:
            records = read_analysis_sheet(_tmp_path)
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")
            os.unlink(_tmp_path)
            return

        if not records:
            st.warning("Không tìm thấy dữ liệu trong sheet 'analysis'.")
            os.unlink(_tmp_path)
            return

        st.success(f"✅ Tìm thấy **{len(records)} máy** trong sheet analysis")

        # Preview table
        preview_rows = []
        for r in records:
            row = {
                "Máy": r['so_may'],
                "Ghi chú": r['note'][:70] if r['note'] else "—",
            }
            for field, mins in r['parsed'].items():
                row[field] = f"{mins}'"
            preview_rows.append(row)

        df_preview = pd.DataFrame(preview_rows)

        # Highlight parsed columns
        _parsed_cols = [c for c in df_preview.columns if c not in ("Máy", "Ghi chú")]

        st.markdown("#### 📋 Kết quả phân tích (xem trước khi điền)")

        # Summary stats
        total_filled_cells = sum(len(r['parsed']) for r in records)
        machines_with_notes = sum(1 for r in records if r['parsed'])
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Máy có ghi chú", machines_with_notes)
        with c2: st.metric("Ô sẽ điền", total_filled_cells)
        with c3: st.metric("Tổng máy", len(records))

        # Filter show only machines with data
        show_all = st.checkbox("Hiện tất cả máy (kể cả không có ghi chú)", value=False, key="show_all_machines")
        df_show = df_preview if show_all else df_preview[df_preview['Ghi chú'] != '—']

        st.dataframe(df_show, use_container_width=True, hide_index=True)

        # Detail per machine
        with st.expander("📖 Chi tiết từng máy"):
            for r in records:
                if not r['note'] or r['note'] in ('0', ''):
                    continue
                st.markdown(f"**Máy {r['so_may']}:** `{r['note'][:80]}`")
                if r['parsed']:
                    for field, mins in r['parsed'].items():
                        st.markdown(f"  → **{field}**: {mins} phút")
                else:
                    st.caption("  (không parse được thời gian)")
                st.divider()

        # ── Fill options ──
        st.markdown("#### ⚙️ Tùy chọn điền")
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            overwrite = st.checkbox(
                "Ghi đè ô đã có dữ liệu", value=False, key="overwrite_check",
                help="Mặc định chỉ điền ô trống. Bật để ghi đè."
            )
        with col_o2:
            preview_only = st.checkbox("Chỉ xem trước, không điền", value=False, key="preview_only")

        # ── Export button ──
        btn_fill = st.button(
            "📥 Điền thời gian & Xuất Excel",
            key="btn_fill_analysis",
            type="primary",
            use_container_width=True,
            disabled=preview_only
        )

        if btn_fill:
            out_name = f"analysis_filled_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            out_path = os.path.join(tempfile.gettempdir(), out_name)

            with st.spinner("Đang điền thời gian vào Excel..."):
                result = fill_excel(_tmp_path, records, out_path)

            st.success(
                f"✅ Điền xong! **{result['filled']} ô** được cập nhật "
                f"| {result['skipped']} ô bỏ qua (đã có dữ liệu)"
            )

            if result['log']:
                with st.expander(f"📋 Log chi tiết ({len(result['log'])} thao tác)"):
                    for line in result['log']:
                        st.text(f"  ✓ {line}")

            # Download
            with open(out_path, "rb") as f:
                st.download_button(
                    label=f"⬇️ Tải xuống {out_name}",
                    data=f.read(),
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_analysis"
                )
            try:
                os.unlink(out_path)
            except:
                pass

        try:
            os.unlink(_tmp_path)
        except:
            pass


if __name__ == "__main__":
    st.set_page_config(page_title="Analysis Parser", page_icon="🕒", layout="wide")
    st.title("🕒 Tự Động Nhập Thời Gian Dừng Máy")
    render_analysis_ui()