
# 🧠 Chatbot lập kế hoạch tuần & báo cáo tháng từ file Excel
# ⚠️ Yêu cầu: pip install streamlit pandas openpyxl

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Lập kế hoạch & Báo cáo công việc", layout="wide")
st.title("🗂️ Lập kế hoạch tuần & Tổng hợp báo cáo tháng")

# --- Session state init ---
if "weekly_plan_data" not in st.session_state:
    st.session_state.weekly_plan_data = []

# --- Upload form templates ---
st.sidebar.header("📤 Tải lên biểu mẫu")
weekly_template_file = st.sidebar.file_uploader("📁 Mẫu kế hoạch tuần (.xlsx)", type=["xlsx"])
monthly_template_file = st.sidebar.file_uploader("📁 Mẫu báo cáo tháng (.xlsx)", type=["xlsx"])
executed_weekly_files = st.sidebar.file_uploader("📚 Kế hoạch tuần đã thực hiện (có thể nhiều)", type=["xlsx"], accept_multiple_files=True)

# --- Nhập kế hoạch tuần mới ---
st.subheader("📅 Nhập kế hoạch tuần mới")
if weekly_template_file:
    weekly_template = pd.read_excel(weekly_template_file)
    if not weekly_template.empty:
        with st.form("weekly_form"):
            st.write("⏱️ Nhập dữ liệu cho từng dòng trong biểu mẫu:")
            weekly_inputs = []
            for i, row in weekly_template.iterrows():
                st.markdown(f"### ➤ {row[0]}")
                row_input = {}
                for col in weekly_template.columns[1:]:
                    val = st.text_input(f"{col} ({row[0]})", key=f"{col}_{i}")
                    row_input[col] = val
                row_input[weekly_template.columns[0]] = row[0]
                weekly_inputs.append(row_input)
            submitted = st.form_submit_button("✅ Hoàn tất kế hoạch tuần")
            if submitted:
                st.session_state.weekly_plan_data = weekly_inputs
                st.success("✅ Đã lưu kế hoạch tuần.")

# --- Xuất kế hoạch tuần vừa hoàn thành ---
if st.session_state.weekly_plan_data:
    st.subheader("📤 Xuất file kế hoạch tuần")
    df_week = pd.DataFrame(st.session_state.weekly_plan_data)
    st.dataframe(df_week)

    buffer = BytesIO()
    df_week.to_excel(buffer, index=False, engine='openpyxl')
    st.download_button("⬇️ Tải file kế hoạch tuần", data=buffer.getvalue(),
                       file_name="ke_hoach_tuan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Tổng hợp báo cáo tháng ---
st.subheader("📊 Tổng hợp báo cáo tháng")
if executed_weekly_files and monthly_template_file:
    combined_df = pd.DataFrame()
    for file in executed_weekly_files:
        df = pd.read_excel(file)
        combined_df = pd.concat([combined_df, df], ignore_index=True)

    st.write("✅ Dữ liệu tổng hợp từ các kế hoạch tuần:")
    st.dataframe(combined_df)

    monthly_template = pd.read_excel(monthly_template_file)
    if not monthly_template.empty:
        # Giả định báo cáo tháng có cột 'Thời gian' là Tuần 1, 2, 3, 4...
        result_df = monthly_template.copy()
        for i, row in result_df.iterrows():
            week_label = str(row[0]).strip().lower()
            matched = combined_df[combined_df.iloc[:, 0].astype(str).str.strip().str.lower().str.contains(week_label)]
            for col in result_df.columns[1:]:
                result_df.at[i, col] = "; ".join(matched[col].dropna().astype(str)) if col in matched.columns else ""

        st.subheader("📄 Xem trước báo cáo tháng")
        st.dataframe(result_df)

        buffer = BytesIO()
        result_df.to_excel(buffer, index=False, engine='openpyxl')
        st.download_button("⬇️ Tải báo cáo tháng", data=buffer.getvalue(),
                           file_name="bao_cao_thang.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("📌 Vui lòng tải lên cả mẫu báo cáo tháng và các file kế hoạch tuần đã thực hiện để tổng hợp.")
