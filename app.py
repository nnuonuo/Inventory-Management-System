import streamlit as st
import pandas as pd
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Restaurant Inventory & Chain Management System", layout="wide")

# --- TÙY CHỈNH GIAO DIỆN BẰNG CSS ---
st.markdown("""
    <style>
    .stApp {
        background-color: #f7f9fa;
    }
    [data-testid="stSidebar"] {
        background-color: #b4cfdc;
        border-right: 1px solid #90bcd5;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label {
        color: #1e293b !important;
    }
    div.stButton > button:first-child, div.stFormSubmitButton > button:first-child {
        background-color: #bed650 !important;
        color: #1e293b !important;
        font-weight: bold;
        border: 1px solid #a8c238;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover, div.stFormSubmitButton > button:first-child:hover {
        background-color: #a8c238 !important;
        border-color: #96af29;
    }
    .stTextInput input, .stSelectbox select, .stNumberInput input {
        border-color: #90bcd5 !important;
        border-radius: 6px;
    }
    h1 {
        color: #2c3e50;
    }
    .stAlert {
        border-radius: 8px;
    }
    </style>
""", unsafe_allow_html=True)

DATA_FILE = "restaurant_inventory_data.xlsx"
BRANCH_DATA_FILE = "branch_inventory_data.xlsx"
EXPORT_FILE = "branch_export_history.csv"
TRANSFER_FILE = "branch_transfer_history.csv"
ORDER_FILE = "branch_order_requests.csv"
BRANCH_PROC_FILE = "branch_processing_history.csv"
SECRET_ACTION_PWD = "264221"

BRANCH_LIST = ["Shibuya", "Little Geisha Baross", "Little Geisha Corvin", "URBN.Station", "Matchy"]
UNIT_LIST = ["Kg", "g", "L", "ml", "Can", "Chai", "Thùng", "Gói", "Hộp", "Cái"]

if "passwords" not in st.session_state:
    st.session_state.passwords = {
        "nuonuo": "264221",
        "heni": "Heni2026",
        "admin": "budapest2026",
        "shibuya": "shibuya123",
        "geisha.baross": "geisha2023",
        "geisha.corvin": "Corvin2026",
        "urbn": "ub2026",
        "matchy": "matchy2026"
    }

LANG = {
    "vi": {
        "login_title": "Đăng Nhập Hệ Thống Quản Lý Kho & Chuỗi Nhà Hàng",
        "login_desc": "Vui lòng chọn ngôn ngữ, nhập **ID tài khoản** và **Mật khẩu** của bạn để tiếp tục.",
        "id_label": "ID tài khoản (Admin hoặc Chi nhánh):",
        "pwd_label": "Mật khẩu:",
        "btn_login": "Đăng Nhập",
        "title": "Hệ Thống Quản Lý Kho & Sơ Chế Chuỗi Nhà Hàng",
        "menu": "Chức Năng Hệ Thống",
        "m_overview": "Tổng Quan & Cảnh Báo Kho",
        "m_import": "Nhập Hàng Kho Tổng",
        "m_edit": "Sửa Tồn Kho Đầu Kỳ",
        "m_add": "Thêm Sản Phẩm Mới",
        "m_process": "Sơ Chế & Hao Hụt",
        "m_distribute": "Cấp Hàng Cho Chi Nhánh",
        "m_branch_inv": "Kho Riêng Chi Nhánh",
        "m_transfer": "Chuyển Hàng Giữa Chi Nhánh",
        "m_order": "Chi Nhánh Đặt Hàng Kho Tổng",
        "m_guide": "Hướng Dẫn Sử Dụng",
        "total_items": "Tổng số mặt hàng",
        "low_stock_warn": "Cảnh báo tồn kho thấp",
        "total_branches": "Tổng số chi nhánh hoạt động",
        "main_stock_table": "Bảng Tồn Kho Kho Tổng (Main Stock)",
        "guide_content": """### Hướng Dẫn Sử Dụng Hệ Thống
* **Quản trị viên (Admin):** Toàn quyền quản lý kho tổng, kho chi nhánh, nhập hàng, thêm sản phẩm, duyệt đơn và phân bổ cấp hàng.
* **Chi nhánh (Branch):** Quản lý kho nội bộ riêng, tự thêm sản phẩm mua ngoài, sơ chế hao hụt, chuyển hàng nội bộ và đặt hàng từ kho tổng."""
    },
    "en": {
        "login_title": "Restaurant Inventory & Chain Management Login",
        "login_desc": "Please select your language, enter your **Account ID** and **Password** to continue.",
        "id_label": "Account ID (Admin or Branch):",
        "pwd_label": "Password:",
        "btn_login": "Log In",
        "title": "Restaurant Inventory & Chain Management System",
        "menu": "System Menu",
        "m_overview": "Overview & Stock Alerts",
        "m_import": "Main Stock Import",
        "m_edit": "Edit Opening Stock",
        "m_add": "Add New Item",
        "m_process": "Processing Log",
        "m_distribute": "Branch Distribution",
        "m_branch_inv": "Branch Local Inventory",
        "m_transfer": "Inter-branch Transfer",
        "m_order": "Branch Order Request",
        "m_guide": "User Guide",
        "total_items": "Total Items",
        "low_stock_warn": "Low Stock Alerts",
        "total_branches": "Active Branches",
        "main_stock_table": "Main Inventory Stock Table",
        "guide_content": """### User Guide
* **Administrator (Admin):** Full control over main stock, branch inventories, imports, and distribution.
* **Branch:** Manage local inventory, add local items, branch processing/waste log, internal transfers, and place orders."""
    },
    "hu": {
        "login_title": "Éttermi Készletkezelő Bejelentkezés",
        "login_desc": "Kérjük, adja meg a fiókazonosítót és a jelszót.",
        "id_label": "Fiók azonosító:",
        "pwd_label": "Jelszó:",
        "btn_login": "Bejelentkezés",
        "title": "Éttermi Készletkezelő és Lánc Rendszer",
        "menu": "Rendszer Menü",
        "m_overview": "Áttekintés és Készletriasztások",
        "m_import": "Központi Készlet Bevételezés",
        "m_edit": "Nyitókészlet Szerkesztése",
        "m_add": "Új Termék Hozzáadása",
        "m_process": "Feldolgozási Napló",
        "m_distribute": "Kiosztás Egységeknek",
        "m_branch_inv": "Egységek Saját Készlete",
        "m_transfer": "Egységek közötti átadás",
        "m_order": "Egységek Rendelése a Központból",
        "m_guide": "Használati Útmutató",
        "total_items": "Összes termék",
        "low_stock_warn": "Alacsony készlet riasztás",
        "total_branches": "Aktív egységek",
        "main_stock_table": "Központi Készlet Táblázat",
        "guide_content": """### Használati Útmutató
* **Rendszergazda (Admin):** Teljes körű vezérlés a központi raktár, fiókkészletek és kiosztások felett.
* **Egység (Branch):** Saját helyi készlet kezelése, saját tételek, feldolgozási és hulladék napló, belső átadások és rendelések."""
    }
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "role" not in st.session_state:
    st.session_state.role = ""
if "branch_name" not in st.session_state:
    st.session_state.branch_name = ""
if "lang_key" not in st.session_state:
    st.session_state.lang_key = "vi"
if "dist_rows_count" not in st.session_state:
    st.session_state.dist_rows_count = 1
if "transfer_rows_count" not in st.session_state:
    st.session_state.transfer_rows_count = 1

if not st.session_state.logged_in:
    login_lang_choice = st.selectbox("Chọn ngôn ngữ / Language / Nyelv:", ["Tiếng Việt", "English", "Magyar"], index=0)
    login_lang_key = "vi" if login_lang_choice == "Tiếng Việt" else ("en" if login_lang_choice == "English" else "hu")
    TL = LANG[login_lang_key]

    st.title(TL["login_title"])
    st.markdown(TL["login_desc"])
    
    with st.form("login_form"):
        input_id = st.text_input(TL["id_label"])
        input_pwd = st.text_input(TL["pwd_label"], type="password")
        submitted_login = st.form_submit_button(TL["btn_login"])
        
        if submitted_login:
            clean_id = input_id.strip().lower()
            if clean_id in st.session_state.passwords and input_pwd == st.session_state.passwords[clean_id]:
                st.session_state.logged_in = True
                if clean_id in ["nuonuo", "heni", "admin"]:
                    st.session_state.role = "Admin"
                    st.session_state.branch_name = ""
                else:
                    st.session_state.role = "Branch"
                    branch_names_map = {
                        "shibuya": "Shibuya",
                        "geisha.baross": "Little Geisha Baross",
                        "geisha.corvin": "Little Geisha Corvin",
                        "urbn": "URBN.Station",
                        "matchy": "Matchy"
                    }
                    st.session_state.branch_name = branch_names_map.get(clean_id, clean_id)
                st.session_state.lang_key = login_lang_key
                st.success("Đăng nhập thành công / Login successful!")
                st.rerun()
            else:
                st.error("Sai ID hoặc mật khẩu! / Incorrect ID or Password!")
    st.stop()

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            main_stock = pd.read_excel(DATA_FILE, sheet_name="MainStock")
            processing = pd.read_excel(DATA_FILE, sheet_name="ProcessingLog")
        except:
            main_stock = pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "OpeningStock", "Import", "BranchExport", "ProcessingExport", "Source", "SubUnit", "SubQuantity", "TotalConverted"])
            processing = pd.DataFrame(columns=["Date", "BatchID", "RawMaterial", "UsedQuantity", "FinishedProduct", "ProducedQuantity", "WasteLoss", "Note"])
        
        for col in ["OpeningStock", "Import", "BranchExport", "ProcessingExport"]:
            if col in main_stock.columns:
                main_stock[col] = pd.to_numeric(main_stock[col], errors="coerce").fillna(0.0)
        return main_stock, processing
    else:
        main_stock = pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "OpeningStock", "Import", "BranchExport", "ProcessingExport", "Source", "SubUnit", "SubQuantity", "TotalConverted"])
        processing = pd.DataFrame(columns=["Date", "BatchID", "RawMaterial", "UsedQuantity", "FinishedProduct", "ProducedQuantity", "WasteLoss", "Note"])
        save_data(main_stock, processing)
        return main_stock, processing

def save_data(main_stock, processing):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        main_stock.to_excel(writer, sheet_name="MainStock", index=False)
        processing.to_excel(writer, sheet_name="ProcessingLog", index=False)

def load_branch_data():
    branch_data = {}
    if os.path.exists(BRANCH_DATA_FILE):
        for b in BRANCH_LIST:
            try:
                branch_data[b] = pd.read_excel(BRANCH_DATA_FILE, sheet_name=b)
            except:
                branch_data[b] = pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "StockQty", "ImportedQty", "UsedQty", "Note"])
    else:
        for b in BRANCH_LIST:
            branch_data[b] = pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "StockQty", "ImportedQty", "UsedQty", "Note"])
        save_branch_data(branch_data)
    return branch_data

def save_branch_data(branch_data):
    with pd.ExcelWriter(BRANCH_DATA_FILE, engine="openpyxl") as writer:
        for b, df in branch_data.items():
            df.to_excel(writer, sheet_name=b, index=False)

if not os.path.exists(EXPORT_FILE):
    pd.DataFrame(columns=["ExportDate", "Branch", "ItemName", "Unit", "Quantity", "Sender", "Receiver"]).to_csv(EXPORT_FILE, index=False)

if not os.path.exists(TRANSFER_FILE):
    pd.DataFrame(columns=["TransferDate", "FromBranch", "ToBranch", "ItemName", "Unit", "Quantity", "Staff"]).to_csv(TRANSFER_FILE, index=False)

if not os.path.exists(ORDER_FILE):
    pd.DataFrame(columns=["OrderDate", "Branch", "ItemName", "Unit", "Quantity", "Status", "Note"]).to_csv(ORDER_FILE, index=False)

if not os.path.exists(BRANCH_PROC_FILE):
    pd.DataFrame(columns=["Date", "Branch", "RawMaterial", "UsedQuantity", "FinishedProduct", "ProducedQuantity", "WasteLoss", "Note"]).to_csv(BRANCH_PROC_FILE, index=False)

def calculate_closing_stock(df):
    if df.empty:
        df["ClosingStock"] = []
        return df
    for col in ["OpeningStock", "Import", "BranchExport", "ProcessingExport"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["ClosingStock"] = df["OpeningStock"] + df["Import"] - df["BranchExport"] - df["ProcessingExport"]
    return df

current_lang_key = st.session_state.get("lang_key", "vi")
T = LANG[current_lang_key]

st.sidebar.header("Ngôn Ngữ / Language")
selected_lang_ui = st.sidebar.selectbox(
    "Chọn ngôn ngữ hiển thị:", 
    ["Tiếng Việt", "English", "Magyar"], 
    index=0 if current_lang_key == "vi" else (1 if current_lang_key == "en" else 2),
    key="ui_lang_select"
)
new_lang_key = "vi" if selected_lang_ui == "Tiếng Việt" else ("en" if selected_lang_ui == "English" else "hu")
if new_lang_key != st.session_state.lang_key:
    st.session_state.lang_key = new_lang_key
    st.rerun()

T = LANG[st.session_state.lang_key]

st.sidebar.markdown("---")
st.sidebar.markdown(f"Tài khoản: `{st.session_state.role}`")
if st.session_state.role == "Branch":
    st.sidebar.markdown(f"Chi nhánh: `{st.session_state.branch_name}`")

if st.sidebar.button("Đăng Xuất / Log Out", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.role = ""
    st.session_state.branch_name = ""
    st.rerun()

st.sidebar.markdown("---")

pending_order_count = 0
if os.path.exists(ORDER_FILE):
    try:
        df_check_order = pd.read_csv(ORDER_FILE)
        if not df_check_order.empty and "Status" in df_check_order.columns:
            pending_order_count = len(df_check_order[df_check_order["Status"].str.contains("chờ", case=False, na=False)])
    except:
        pass

st.sidebar.title(T["menu"])

order_menu_label = T["m_order"]
if st.session_state.role == "Admin" and pending_order_count > 0:
    order_menu_label = f"🔴 {T['m_order']} ({pending_order_count} chờ duyệt)"

if st.session_state.role == "Admin":
    menu_options = {
        T["m_overview"]: "overview",
        T["m_import"]: "import",
        T["m_edit"]: "edit",
        T["m_add"]: "add",
        T["m_process"]: "process",
        T["m_distribute"]: "distribute",
        T["m_branch_inv"]: "branch_inv",
        T["m_transfer"]: "transfer",
        order_menu_label: "order",
        T["m_guide"]: "guide"
    }
else:
    menu_options = {
        T["m_overview"]: "overview",
        T["m_branch_inv"]: "branch_inv",
        T["m_transfer"]: "transfer",
        T["m_order"]: "order",
        T["m_guide"]: "guide"
    }

selected_menu_label = st.sidebar.radio("Chọn chức năng:", list(menu_options.keys()))
choice = menu_options[selected_menu_label]

st.title(T["title"])

main_stock_df, processing_df = load_data()
main_stock_df = calculate_closing_stock(main_stock_df)
branch_data_dict = load_branch_data()

if choice == "overview":
    st.subheader(T["m_overview"])
    total_items = len(main_stock_df)
    low_stock_items = len(main_stock_df[main_stock_df["ClosingStock"] <= 5])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label=T["total_items"], value=total_items)
    with col2:
        st.metric(label=T["low_stock_warn"], value=low_stock_items, delta="Ổn định" if low_stock_items == 0 else "Chú ý", delta_color="inverse")
    with col3:
        st.metric(label=T["total_branches"], value=len(BRANCH_LIST))
        
    st.markdown("---")
    st.markdown(f"### {T['main_stock_table']}")
    st.dataframe(main_stock_df, use_container_width=True)

elif choice == "import":
    st.subheader(T["m_import"])
    with st.form("import_form"):
        item_list = main_stock_df["ItemName"].tolist() if not main_stock_df.empty else []
        selected_item = st.selectbox("Chọn sản phẩm cần nhập:", item_list)
        import_qty = st.number_input("Số lượng nhập thêm:", min_value=0.0, step=1.0)
        import_source = st.text_input("Nhà cung cấp / Nguồn gốc:", value="Nhà cung cấp chính")
        submitted_import = st.form_submit_button("Xác Nhận Nhập Hàng")
        if submitted_import and selected_item:
            idx = main_stock_df[main_stock_df["ItemName"] == selected_item].index
            if not idx.empty:
                main_stock_df.loc[idx, "Import"] += import_qty
                main_stock_df.loc[idx, "Source"] = import_source
                save_data(main_stock_df, processing_df)
                st.success("Thành công!")
                st.rerun()

elif choice == "edit":
    st.subheader(T["m_edit"])
    with st.form("edit_stock_form"):
        item_list = main_stock_df["ItemName"].tolist() if not main_stock_df.empty else []
        selected_item = st.selectbox("Chọn sản phẩm cần sửa:", item_list)
        new_opening = st.number_input("Tồn kho đầu kỳ mới:", min_value=0.0, step=1.0)
        pwd_input = st.text_input("Nhập mật khẩu bảo mật:", type="password")
        submitted_edit = st.form_submit_button("Cập Nhật Tồn Kho")
        if submitted_edit:
            if pwd_input != SECRET_ACTION_PWD:
                st.error("Sai mật khẩu xác nhận!")
            elif selected_item:
                idx = main_stock_df[main_stock_df["ItemName"] == selected_item].index
                if not idx.empty:
                    main_stock_df.loc[idx, "OpeningStock"] = new_opening
                    save_data(main_stock_df, processing_df)
                    st.success("Cập nhật thành công!")
                    st.rerun()

elif choice == "add":
    st.subheader(T["m_add"])
    auto_id = f"SP{len(main_stock_df)+1:03d}" if not main_stock_df.empty else "SP001"
    with st.form("add_item_form"):
        new_id = st.text_input("Mã sản phẩm (Tự động):", value=auto_id, disabled=True)
        new_name = st.text_input("Tên sản phẩm:")
        new_unit = st.selectbox("Đơn vị tính:", UNIT_LIST)
        new_opening = st.number_input("Tồn kho đầu kỳ:", min_value=0.0, step=1.0)
        new_source = st.text_input("Nguồn cung cấp:", value="Nhà cung cấp chính")
        submitted_add = st.form_submit_button("Thêm Sản Phẩm")
        if submitted_add and new_name.strip():
            new_row = {
                "ItemID": auto_id, "ItemName": new_name, "Unit": new_unit,
                "OpeningStock": new_opening, "Import": 0.0, "BranchExport": 0.0,
                "ProcessingExport": 0.0, "Source": new_source, "SubUnit": "", "SubQuantity": 0.0, "TotalConverted": ""
            }
            main_stock_df = pd.concat([main_stock_df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(main_stock_df, processing_df)
            st.success("Thêm sản phẩm thành công!")
            st.rerun()

elif choice == "process":
    st.subheader(T["m_process"])
    with st.form("process_form"):
        p_date = st.date_input("Ngày sơ chế:", datetime.now())
        p_batch = st.text_input("Mã lô:", value=f"BATCH-{datetime.now().strftime('%Y%m%d')}")
        p_raw = st.selectbox("Nguyên liệu thô:", main_stock_df["ItemName"].tolist() if not main_stock_df.empty else [])
        p_used_qty = st.number_input("Số lượng dùng:", min_value=0.0, step=1.0)
        p_finished = st.text_input("Tên thành phẩm thu được:")
        p_produced_qty = st.number_input("Số lượng thành phẩm:", min_value=0.0, step=1.0)
        p_waste = st.number_input("Hao hụt / Phế phẩm:", min_value=0.0, step=1.0)
        p_note = st.text_area("Ghi chú:")
        submitted_process = st.form_submit_button("Lưu Sơ Chế")
        if submitted_process:
            idx = main_stock_df[main_stock_df["ItemName"] == p_raw].index
            if not idx.empty:
                main_stock_df.loc[idx, "ProcessingExport"] += p_used_qty
                new_log = {
                    "Date": p_date.strftime("%Y-%m-%d"), "BatchID": p_batch,
                    "RawMaterial": p_raw, "UsedQuantity": p_used_qty,
                    "FinishedProduct": p_finished, "ProducedQuantity": p_produced_qty,
                    "WasteLoss": p_waste, "Note": p_note
                }
                processing_df = pd.concat([processing_df, pd.DataFrame([new_log])], ignore_index=True)
                save_data(main_stock_df, processing_df)
                st.success("Đã ghi nhận sơ chế!")
                st.rerun()

elif choice == "distribute":
    st.subheader(T["m_distribute"])
    with st.form("distribute_form"):
        d_branch = st.selectbox("Chọn chi nhánh nhận hàng:", BRANCH_LIST)
        item_list_dist = main_stock_df["ItemName"].tolist() if not main_stock_df.empty else []
        
        selected_items_data = []
        for i in range(st.session_state.dist_rows_count):
            st.markdown(f"**Sản phẩm #{i+1}**")
            col_i1, col_i2 = st.columns([2, 1])
            with col_i1:
                d_item = st.selectbox("Sản phẩm:", item_list_dist, key=f"dist_item_{i}") if item_list_dist else None
            with col_i2:
                d_qty = st.number_input("Số lượng:", min_value=0.0, step=1.0, key=f"dist_qty_{i}")
            selected_items_data.append((d_item, d_qty))
            
        submitted_add_row = st.form_submit_button("➕ Thêm sản phẩm khác")
        submitted_dist = st.form_submit_button("Xác Nhận Cấp Hàng")
        
        if submitted_add_row:
            st.session_state.dist_rows_count += 1
            st.rerun()
            
        if submitted_dist:
            valid_any = False
            exp_df = pd.read_csv(EXPORT_FILE)
            target_branch_df = branch_data_dict[d_branch]
            
            for d_item, d_qty in selected_items_data:
                if d_item and d_qty > 0:
                    valid_any = True
                    idx = main_stock_df[main_stock_df["ItemName"] == d_item].index
                    if not idx.empty:
                        main_stock_df.loc[idx, "BranchExport"] += d_qty
                        
                    b_idx = target_branch_df[target_branch_df["ItemName"] == d_item].index
                    unit_val = "Kg"
                    if not idx.empty:
                        unit_val = main_stock_df.loc[idx, "Unit"].values[0]

                    if not b_idx.empty:
                        target_branch_df.loc[b_idx, "StockQty"] += d_qty
                        target_branch_df.loc[b_idx, "ImportedQty"] += d_qty
                    else:
                        match_row = main_stock_df[main_stock_df["ItemName"] == d_item]
                        item_id_val = match_row["ItemID"].values[0] if not match_row.empty else "SP000"
                        new_b_row = {
                            "ItemID": item_id_val, "ItemName": d_item, "Unit": unit_val,
                            "StockQty": d_qty, "ImportedQty": d_qty, "UsedQty": 0.0, "Note": "Nhận từ Kho Tổng"
                        }
                        target_branch_df = pd.concat([target_branch_df, pd.DataFrame([new_b_row])], ignore_index=True)
                    
                    branch_data_dict[d_branch] = target_branch_df

                    export_record = {
                        "ExportDate": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Branch": d_branch, "ItemName": d_item, "Unit": unit_val,
                        "Quantity": d_qty, "Sender": "Kho Tổng", "Receiver": d_branch
                    }
                    exp_df = pd.concat([exp_df, pd.DataFrame([export_record])], ignore_index=True)
            
            if valid_any:
                exp_df.to_csv(EXPORT_FILE, index=False)
                save_data(main_stock_df, processing_df)
                save_branch_data(branch_data_dict)
                st.session_state.dist_rows_count = 1
                st.success(f"Đã cấp hàng và tự động cập nhật vào kho riêng của chi nhánh **{d_branch}**!")
                st.rerun()

elif choice == "branch_inv":
    st.subheader(T["m_branch_inv"])
    if st.session_state.role == "Admin":
        active_branch = st.selectbox("Chọn chi nhánh để quản lý kho:", BRANCH_LIST)
    else:
        active_branch = st.session_state.branch_name
        st.markdown(f"Đang xem kho riêng của chi nhánh: **{active_branch}**")

    current_b_df = branch_data_dict.get(active_branch, pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "StockQty", "ImportedQty", "UsedQty", "Note"]))
    
    st.markdown(f"### Kho Tồn Của Chi Nhánh: {active_branch}")
    st.dataframe(current_b_df, use_container_width=True)

    st.markdown("---")
    
    # 1. Thêm sản phẩm mới vào kho chi nhánh (mua ngoài)
    with st.expander("➕ Tự thêm sản phẩm mới vào kho chi nhánh này (Mua ngoài / Khác)"):
        with st.form(f"add_branch_item_form_{active_branch}"):
            b_new_name = st.text_input("Tên sản phẩm mới:")
            b_new_unit = st.selectbox("Đơn vị tính:", UNIT_LIST, key=f"b_unit_{active_branch}")
            b_new_stock = st.number_input("Số lượng tồn kho ban đầu:", min_value=0.0, step=1.0, key=f"b_stock_{active_branch}")
            b_new_note = st.text_input("Ghi chú nguồn hàng:", value="Chi nhánh tự mua ngoài", key=f"b_note_{active_branch}")
            btn_add_b_item = st.form_submit_button("Thêm Sản Phẩm Vào Kho Chi Nhánh")
            
            if btn_add_b_item and b_new_name.strip():
                b_auto_id = f"BS-{len(current_b_df)+1:03d}"
                new_b_row = {
                    "ItemID": b_auto_id,
                    "ItemName": b_new_name,
                    "Unit": b_new_unit,
                    "StockQty": b_new_stock,
                    "ImportedQty": b_new_stock,
                    "UsedQty": 0.0,
                    "Note": b_new_note
                }
                current_b_df = pd.concat([current_b_df, pd.DataFrame([new_b_row])], ignore_index=True)
                branch_data_dict[active_branch] = current_b_df
                save_branch_data(branch_data_dict)
                st.success(f"Đã thêm sản phẩm **{b_new_name}** vào kho chi nhánh **{active_branch}** thành công!")
                st.rerun()

    # 2. Ghi nhận sơ chế & hao hụt tại chi nhánh
    with st.expander("🔪 Sơ chế & Ghi nhận hao hụt tại chi nhánh"):
        with st.form(f"branch_processing_form_{active_branch}"):
            bp_date = st.date_input("Ngày sơ chế:", datetime.now())
            bp_raw = st.selectbox("Chọn nguyên liệu trong kho chi nhánh:", current_b_df["ItemName"].tolist() if not current_b_df.empty else [])
            bp_used = st.number_input("Số lượng nguyên liệu mang ra sơ chế:", min_value=0.0, step=1.0)
            bp_finished_name = st.text_input("Tên thành phẩm thu được sau sơ chế:")
            bp_prod_qty = st.number_input("Số lượng thành phẩm thu được:", min_value=0.0, step=1.0)
            bp_waste = st.number_input("Hao hụt / Phế phẩm bỏ đi:", min_value=0.0, step=1.0)
            bp_note = st.text_input("Ghi chú chi tiết:")
            btn_save_bp = st.form_submit_button("Lưu Sơ Chế & Trừ Kho")
            
            if btn_save_bp and bp_raw:
                b_idx = current_b_df[current_b_df["ItemName"] == bp_raw].index
                if not b_idx.empty:
                    current_b_df.loc[b_idx, "UsedQty"] += bp_used
                    current_b_df.loc[b_idx, "StockQty"] = max(0.0, current_b_df.loc[b_idx, "StockQty"] - bp_used)
                    branch_data_dict[active_branch] = current_b_df
                    save_branch_data(branch_data_dict)
                    
                    bp_record = {
                        "Date": bp_date.strftime("%Y-%m-%d"),
                        "Branch": active_branch,
                        "RawMaterial": bp_raw,
                        "UsedQuantity": bp_used,
                        "FinishedProduct": bp_finished_name,
                        "ProducedQuantity": bp_prod_qty,
                        "WasteLoss": bp_waste,
                        "Note": bp_note
                    }
                    b_proc_df = pd.read_csv(BRANCH_PROC_FILE)
                    b_proc_df = pd.concat([b_proc_df, pd.DataFrame([bp_record])], ignore_index=True)
                    b_proc_df.to_csv(BRANCH_PROC_FILE, index=False)
                    
                    st.success(f"Đã ghi nhận sơ chế và cập nhật trừ kho chi nhánh **{active_branch}** thành công!")
                    st.rerun()

    # 3. Cập nhật xuất dùng thông thường tại chi nhánh
    st.markdown("#### Cập nhật xuất dùng hàng ngày tại chi nhánh")
    with st.form(f"update_branch_stock_{active_branch}"):
        b_items = current_b_df["ItemName"].tolist() if not current_b_df.empty else []
        sel_b_item = st.selectbox("Chọn mặt hàng xuất dùng:", b_items) if b_items else st.selectbox("Không có nguyên liệu", [])
        used_amount = st.number_input("Số lượng xuất dùng:", min_value=0.0, step=1.0)
        b_note = st.text_input("Ghi chú xuất dùng:")
        btn_save_b = st.form_submit_button("Cập Nhật Xuất Dùng (Trừ Tồn Kho)")
        
        if btn_save_b and sel_b_item:
            b_idx = current_b_df[current_b_df["ItemName"] == sel_b_item].index
            if not b_idx.empty:
                current_b_df.loc[b_idx, "UsedQty"] += used_amount
                current_b_df.loc[b_idx, "StockQty"] = max(0.0, current_b_df.loc[b_idx, "StockQty"] - used_amount)
                branch_data_dict[active_branch] = current_b_df
                save_branch_data(branch_data_dict)
                st.success("Cập nhật kho chi nhánh thành công!")
                st.rerun()

elif choice == "transfer":
    st.subheader(T["m_transfer"])
    transfer_units = BRANCH_LIST + ["Kho Tổng"]
    with st.form("transfer_form"):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            t_from = st.selectbox("Từ đơn vị gửi:", transfer_units)
        with col_f2:
            t_to = st.selectbox("Đến đơn vị nhận:", transfer_units, index=1 if len(transfer_units) > 1 else 0)
            
        t_staff = st.text_input("Nhân viên thực hiện:")
        item_list_trans = main_stock_df["ItemName"].tolist() if not main_stock_df.empty else []
        selected_trans_data = []
        
        for i in range(st.session_state.transfer_rows_count):
            st.markdown(f"**Sản phẩm chuyển #{i+1}**")
            col_t1, col_t2, col_t3 = st.columns([2, 1, 1])
            with col_t1:
                t_item = st.selectbox("Sản phẩm:", item_list_trans, key=f"trans_item_{i}") if item_list_trans else None
            with col_t2:
                t_qty = st.number_input("Số lượng:", min_value=0.0, step=1.0, key=f"trans_qty_{i}")
            with col_t3:
                t_unit = st.selectbox("Đơn vị:", UNIT_LIST, key=f"trans_unit_{i}")
            selected_trans_data.append((t_item, t_qty, t_unit))
            
        submitted_add_trans_row = st.form_submit_button("➕ Thêm sản phẩm khác")
        submitted_trans = st.form_submit_button("Xác Nhận Chuyển Hàng")
        
        if submitted_add_trans_row:
            st.session_state.transfer_rows_count += 1
            st.rerun()
            
        if submitted_trans:
            if t_from == t_to:
                st.error("Đơn vị gửi và nhận không được trùng nhau!")
            else:
                valid_any_trans = False
                trans_df = pd.read_csv(TRANSFER_FILE)
                for t_item, t_qty, t_unit in selected_trans_data:
                    if t_item and t_qty > 0:
                        valid_any_trans = True
                        trans_record = {
                            "TransferDate": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "FromBranch": t_from, "ToBranch": t_to,
                            "ItemName": t_item, "Unit": t_unit,
                            "Quantity": t_qty, "Staff": t_staff
                        }
                        trans_df = pd.concat([trans_df, pd.DataFrame([trans_record])], ignore_index=True)
                
                if valid_any_trans:
                    trans_df.to_csv(TRANSFER_FILE, index=False)
                    st.session_state.transfer_rows_count = 1
                    st.success("Chuyển hàng nội bộ thành công!")
                    st.rerun()

elif choice.startswith("🔴") or choice == "order":
    st.subheader(T["m_order"])
    if st.session_state.role == "Branch":
        st.markdown(f"Giao diện đặt hàng cho chi nhánh: **{st.session_state.branch_name}**")
        with st.form("branch_order_form"):
            order_item = st.selectbox("Chọn sản phẩm muốn đặt:", main_stock_df["ItemName"].tolist() if not main_stock_df.empty else [])
            col_o1, col_o2 = st.columns([2, 1])
            with col_o1:
                order_qty = st.number_input("Số lượng đặt:", min_value=1.0, step=1.0)
            with col_o2:
                order_unit = st.selectbox("Đơn vị:", UNIT_LIST)
            order_note = st.text_area("Ghi chú thêm cho Kho Tổng:")
            submitted_order = st.form_submit_button("Gửi Yêu Cầu Đặt Hàng")
            if submitted_order:
                new_order = {
                    "OrderDate": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Branch": st.session_state.branch_name,
                    "ItemName": order_item, "Unit": order_unit,
                    "Quantity": order_qty, "Status": "Đang chờ duyệt", "Note": order_note
                }
                order_df = pd.read_csv(ORDER_FILE)
                order_df = pd.concat([order_df, pd.DataFrame([new_order])], ignore_index=True)
                order_df.to_csv(ORDER_FILE, index=False)
                st.success("Gửi yêu cầu đặt hàng thành công!")
                st.rerun()
    else:
        order_df = pd.read_csv(ORDER_FILE)
        if order_df.empty:
            st.info("Chưa có đơn hàng nào.")
        else:
            st.dataframe(order_df, use_container_width=True)
            with st.form("update_order_form"):
                order_idx = st.number_input("Nhập số thứ tự dòng đơn hàng cần xử lý:", min_value=0, max_value=max(0, len(order_df)-1), step=1)
                new_status = st.selectbox("Đổi trạng thái thành:", ["Đang chờ duyệt", "Đã duyệt", "Từ chối"])
                submitted_update = st.form_submit_button("Cập Nhật Trạng Thái Đơn")
                if submitted_update:
                    order_df.loc[order_idx, "Status"] = new_status
                    order_df.to_csv(ORDER_FILE, index=False)
                    st.success("Cập nhật trạng thái thành công!")
                    st.rerun()

elif choice == "guide":
    st.subheader(T["m_guide"])
    st.markdown(T["guide_content"])
