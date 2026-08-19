import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import openpyxl

st.set_page_config(page_title="Restaurant Inventory & Chain Management System", layout="wide")

# --- CSS ---
st.markdown("""
    <style>
    .stApp { background-color: #f7f9fa; }
    [data-testid="stSidebar"] { background-color: #b4cfdc; border-right: 1px solid #90bcd5; }
    div.stButton > button:first-child, div.stFormSubmitButton > button:first-child {
        background-color: #bed650 !important; color: #1e293b !important; font-weight: bold; border-radius: 8px;
    }
    </style>
""", unsafe_allow_html=True)

# --- CONFIG ---
DATA_FILE = "restaurant_inventory_data.xlsx"
BRANCH_DATA_FILE = "branch_inventory_data.xlsx"
EXPORT_FILE = "branch_export_history.csv"
TRANSFER_FILE = "branch_transfer_history.csv"
ORDER_FILE = "branch_order_requests.csv"
BRANCH_PROC_FILE = "branch_processing_history.csv"
SECRET_ACTION_PWD = "264221"
BRANCH_LIST = ["Shibuya", "Little Geisha Baross", "Little Geisha Corvin", "URBN.Station", "Matchy"]
UNIT_LIST = ["Kg", "g", "L", "ml", "Can", "Chai", "Thùng", "Gói", "Hộp", "Cái"]

# --- DỮ LIỆU NGÔN NGỮ ---
LANG = {
    "vi": {
        "m_history": "Xem Lịch Sử Giao Dịch & Sơ Chế",
        "tip_history": "💡 Xem lại nhật ký hoạt động và tải báo cáo Excel.",
        "m_order": "Chi Nhánh Đặt Hàng Kho Tổng",
        "m_overview": "Tổng Quan & Cảnh Báo Kho",
        "m_import": "Nhập Hàng Kho Tổng",
        "m_edit": "Sửa Tồn Kho Đầu Kỳ",
        "m_add": "Thêm Sản Phẩm Mới",
        "m_process": "Sơ Chế & Hao Hụt",
        "m_distribute": "Cấp Hàng Cho Chi Nhánh",
        "m_branch_inv": "Kho Riêng Chi Nhánh",
        "m_transfer": "Chuyển Hàng Giữa Chi Nhánh",
        "m_guide": "Hướng Dẫn Sử Dụng",
        "title": "Hệ Thống Quản Lý Kho & Sơ Chế Chuỗi Nhà Hàng",
        "menu": "Chức Năng Hệ Thống"
    }
}
# (Các phần khác giữ nguyên như yêu cầu của người dùng...)

# --- HÀM LOAD/SAVE ---
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_excel(DATA_FILE, sheet_name="MainStock"), pd.read_excel(DATA_FILE, sheet_name="ProcessingLog")
    return pd.DataFrame(), pd.DataFrame()

# --- MAIN APP LOGIC ---
if "lang_key" not in st.session_state: st.session_state.lang_key = "vi"
T = LANG[st.session_state.lang_key]

# ... [PHẦN LOGIN & CÁC MENU KHÁC GIỮ NGUYÊN] ...

# --- PHẦN LỊCH SỬ (HISTORY) ĐÃ TÍCH HỢP TẢI EXCEL ---
if choice == "history":
    st.subheader(T["m_history"])
    st.info(T["tip_history"])
    tab1, tab2, tab3 = st.tabs(["Cấp hàng", "Sơ chế", "Chuyển nội bộ"])
    
    with tab1:
        if os.path.exists(EXPORT_FILE):
            df = pd.read_csv(EXPORT_FILE)
            st.dataframe(df, use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as w: df.to_excel(w, index=False)
            st.download_button("📥 Tải Lịch Sử Cấp Hàng (Excel)", output.getvalue(), "Lich_Su_Cap_Hang.xlsx")
    
    with tab2:
        processing_df, _ = load_data()
        if not processing_df.empty:
            st.dataframe(processing_df, use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as w: processing_df.to_excel(w, index=False)
            st.download_button("📥 Tải Nhật Ký Sơ Chế (Excel)", output.getvalue(), "Nhat_Ky_So_Che.xlsx")
            
    with tab3:
        if os.path.exists(TRANSFER_FILE):
            df = pd.read_csv(TRANSFER_FILE)
            st.dataframe(df, use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as w: df.to_excel(w, index=False)
            st.download_button("📥 Tải Lịch Sử Chuyển (Excel)", output.getvalue(), "Lich_Su_Chuyen.xlsx")
