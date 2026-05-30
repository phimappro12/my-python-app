"""
mtr_kg_ui.py — Giao diện tính mét ↔ kg cho beam dệt
Chạy riêng: streamlit run mtr_kg_ui.py
Hoặc nhúng vào app.py
"""
import streamlit as st
import pandas as pd
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from mtr_kg import mtr_to_kg, kg_to_mtr, list_yarns, YARN_TABLE

st.set_page_config(page_title="MTR ↔ KG Converter", page_icon="⚖️", layout="wide")

st.title("⚖️ Bảng Đổi Mét ↔ Kg (Beam Dệt)")
st.caption("Công thức: KG = Yards × Tổng sợi × Hệ số / 1000 | 1m = 1.09361 yards")

col1, col2 = st.columns(2, gap="large")

# ── CỘT TRÁI: Mét → Kg ──
with col1:
    st.markdown("### 📏 Mét → Kg")
    with st.container(border=True):
        m_mtr     = st.number_input("Số mét:", min_value=0.0, value=2000.0, step=100.0, key="m_mtr")
        m_total   = st.number_input("Tổng sợi:", min_value=1.0, value=2620.0, step=50.0, key="m_total",
                                    help="Số sợi trong beam (VD: 2620)")
        m_yarn    = st.selectbox("Loại sợi:", options=["-- Chọn --"] + list_yarns(),
                                  key="m_yarn")
        m_custom  = st.text_input("Hoặc nhập loại sợi khác:", key="m_custom",
                                   placeholder="VD: CD 30S/2 (10)")

        if st.button("Tính KG", key="btn_mtr", type="primary", use_container_width=True):
            yarn = m_custom.strip() if m_custom.strip() else (m_yarn if m_yarn != "-- Chọn --" else "")
            if not yarn:
                st.warning("Vui lòng chọn hoặc nhập loại sợi")
            else:
                r = mtr_to_kg(m_mtr, m_total, yarn)
                if "error" in r:
                    st.error(r["error"])
                else:
                    st.success(f"**{r['kg']:,.2f} kg**")
                    st.markdown(
                        f"| | |\n|---|---|\n"
                        f"| Mét ban đầu | **{r['mtr']:,.0f} m** |\n"
                        f"| Quy yards | {r['yards']:,.1f} yards |\n"
                        f"| Tổng sợi | {r['total_soi']:,} |\n"
                        f"| Loại sợi | {r['yarn']} |\n"
                        f"| Hệ số | {r['he_so']} |\n"
                        f"| **Kết quả** | **{r['kg']:,.2f} kg** |"
                    )
                    st.caption(f"📐 {r['formula']}")

# ── CỘT PHẢI: Kg → Mét ──
with col2:
    st.markdown("### ⚖️ Kg → Mét")
    with st.container(border=True):
        k_kg      = st.number_input("Số kg:", min_value=0.0, value=206.31, step=10.0, key="k_kg")
        k_total   = st.number_input("Tổng sợi:", min_value=1.0, value=2620.0, step=50.0, key="k_total")
        k_yarn    = st.selectbox("Loại sợi:", options=["-- Chọn --"] + list_yarns(),
                                  key="k_yarn")
        k_custom  = st.text_input("Hoặc nhập loại sợi khác:", key="k_custom",
                                   placeholder="VD: CD 30S/2 (10)")

        if st.button("Tính Mét", key="btn_kg", type="primary", use_container_width=True):
            yarn = k_custom.strip() if k_custom.strip() else (k_yarn if k_yarn != "-- Chọn --" else "")
            if not yarn:
                st.warning("Vui lòng chọn hoặc nhập loại sợi")
            else:
                r = kg_to_mtr(k_kg, k_total, yarn)
                if "error" in r:
                    st.error(r["error"])
                else:
                    st.success(f"**{r['mtr']:,.1f} mét** ({r['yards']:,.1f} yards)")
                    st.markdown(
                        f"| | |\n|---|---|\n"
                        f"| Kg ban đầu | **{r['kg']:,.2f} kg** |\n"
                        f"| Tổng sợi | {r['total_soi']:,} |\n"
                        f"| Loại sợi | {r['yarn']} |\n"
                        f"| Hệ số | {r['he_so']} |\n"
                        f"| Yards | {r['yards']:,.1f} yards |\n"
                        f"| **Kết quả** | **{r['mtr']:,.1f} mét** |"
                    )
                    st.caption(f"📐 {r['formula']}")

st.markdown("---")
st.markdown("### 📋 Bảng hệ số tất cả loại sợi")
df_table = pd.DataFrame([
    {"Loại sợi": k, "Hệ số (kg/yard/1000sợi)": v,
     "100m/2620 sợi (kg)": round(100 * 1.09361 * 2620 * v / 1000, 2)}
    for k, v in sorted(YARN_TABLE.items(), key=lambda x: x[1])
])
st.dataframe(df_table, use_container_width=True, hide_index=True)