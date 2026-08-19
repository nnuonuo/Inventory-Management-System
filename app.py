import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import openpyxl

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
UNIT_LIST = ["Kg", "g", "L", "ml", "Can", "Chai", "Thùng", "Gói", "Hộp", "Cái", "db"]

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
        "m_branch_import": "Nhập Hàng Chi Nhánh",
        "m_edit": "Sửa Tồn Kho Đầu Kỳ",
        "m_add": "Thêm Sản Phẩm Mới",
        "m_process": "Sơ Chế Kho Tổng & Hao Hụt",
        "m_branch_process": "Sơ Chế & Hao Hụt Chi Nhánh",
        "m_distribute": "Cấp Hàng Cho Chi Nhánh",
        "m_branch_inv": "Kho Riêng Chi Nhánh",
        "m_transfer": "Chuyển Hàng Giữa Chi Nhánh",
        "m_order": "Chi Nhánh Đặt Hàng Kho Tổng",
        "m_history": "Xem Lịch Sử Giao Dịch & Sơ Chế",
        "m_guide": "Hướng Dẫn Sử Dụng",
        "total_items": "Tổng số mặt hàng",
        "low_stock_warn": "Cảnh báo tồn kho thấp",
        "total_branches": "Tổng số chi nhánh hoạt động",
        "main_stock_table": "Bảng Tồn Kho Kho Tổng (Main Stock)",
        "tip_overview": "💡 **Mục này làm gì?** Xem nhanh tổng số lượng mặt hàng, các sản phẩm tồn kho thấp cần chú ý và thông tin chuỗi.",
        "tip_import": "💡 **Mục này làm gì?** Ghi nhận số lượng nguyên vật liệu mới nhập từ nhà cung cấp vào kho tổng.",
        "tip_branch_import": "💡 **Mục này làm gì?** Ghi nhận số lượng nguyên vật liệu mới nhập trực tiếp tại kho của chi nhánh.",
        "tip_edit": "💡 **Mục này làm gì?** Điều chỉnh lại tồn kho đầu kỳ của sản phẩm (Dành riêng cho Admin, yêu cầu mật khẩu bảo mật).",
        "tip_add": "💡 **Mục này làm gì?** Khai báo sản phẩm mới vào danh mục quản lý của kho tổng.",
        "tip_process": "💡 **Mục này làm gì?** Ghi nhận quá trình sơ chế nguyên liệu thô thành thành phẩm và hao hụt phát sinh tại kho tổng.",
        "tip_branch_process": "💡 **Mục này làm gì?** Ghi nhận quá trình sơ chế nguyên liệu tại chi nhánh, tự động trừ tồn kho nguyên liệu thô và tăng lượng thành phẩm.",
        "tip_distribute": "💡 **Mục này làm gì?** Phân bổ và cấp phát hàng hóa từ kho tổng xuống các chi nhánh trong hệ thống.",
        "tip_branch_inv": "💡 **Mục này làm gì?** Quản lý số lượng tồn kho nội bộ tại từng chi nhánh cụ thể.",
        "tip_transfer": "💡 **Mục này làm gì?** Ghi nhận hoạt động điều chuyển hàng hóa qua lại giữa các đơn vị nội bộ.",
        "tip_order": "💡 **Mục này làm gì?** Quản lý các phiếu đặt hàng từ chi nhánh gửi về kho tổng để xét duyệt.",
        "tip_history": "💡 **Mục này làm gì?** Xem lại toàn bộ nhật ký giao dịch, cấp hàng, sơ chế và chuyển nội bộ kèm theo tùy chọn tải Excel.",
        "guide_content": """### 📖 Sổ Tay Hướng Dẫn Sử Dụng Hệ Thống Kho & Chuỗi

#### 1. 📊 Tổng Quan & Cảnh Báo Kho (Overview)
* **Mục đích:** Xem nhanh tổng số lượng mặt hàng, các sản phẩm đang có nguy cơ hết hàng ($\le 5$) và số lượng chi nhánh.
* **Cần làm gì:** Kiểm tra bảng tồn kho chính và bấm nút **Tải xuống bảng tồn kho tổng** để xuất file Excel báo cáo.

#### 2. 📥 Nhập Hàng Kho Tổng (Main Stock Import)
* **Mục đích:** Ghi nhận số lượng nguyên vật liệu mới nhập từ nhà cung cấp vào kho tổng.
* **Cần làm gì:** Chọn đúng sản phẩm cần nhập, nhập số lượng thêm và điền tên nhà cung cấp, sau đó bấm **Xác Nhận Nhập Hàng**.

#### 3. 📥 Nhập Hàng Chi Nhánh (Branch Local Import)
* **Mục đích:** Ghi nhận số lượng nguyên vật liệu mới nhập trực tiếp tại kho của chi nhánh.
* **Cần làm gì:** Chọn chi nhánh, chọn sản phẩm, nhập số lượng và nhà cung cấp địa phương.

#### 4. 📝 Sửa Tồn Kho Đầu Kỳ (Edit Opening Stock)
* **Mục đích:** Điều chỉnh lại số lượng gốc ban đầu của sản phẩm.
* **Cần làm gì:** Chỉ dành cho Admin khi kiểm kê kho. Bắt buộc nhập mật khẩu bảo mật (`264221`) để xác nhận thay đổi.

#### 5. ➕ Thêm Sản Phẩm Mới (Add New Item)
* **Mục đích:** Khai báo một mặt hàng mới vào hệ thống quản lý.
* **Cần làm gì:** Nhập tên sản phẩm, chọn đơn vị tính, điền số lượng tồn đầu kỳ và nguồn cung cấp.

#### 6. 🔪 Sơ Chế & Hao Hụt Kho Tổng / Chi Nhánh (Processing Log)
* **Mục đích:** Quản lý quy trình chế biến nguyên liệu thô thành thành phẩm và ghi nhận phần hao hụt/phế phẩm.
* **Cần làm gì:** Chọn nguyên liệu thô bị trừ kho, nhập số lượng dùng, tên thành phẩm, số lượng thu về và lượng hao hụt.

#### 7. 🚚 Cấp Hàng Cho Chi Nhánh (Branch Distribution)
* **Mục đích:** Phân bổ hàng hóa từ kho tổng xuống các chi nhánh.
* **Cần làm gì:** Chọn chi nhánh nhận, chọn sản phẩm và số lượng tương ứng, hệ thống sẽ tự động trừ kho tổng và cộng vào kho chi nhánh.

#### 8. 🏪 Kho Riêng Chi Nhánh (Branch Local Inventory)
* **Mục đích:** Theo dõi số lượng tồn kho nội bộ tại từng chi nhánh.
* **Cần làm gì:** Kiểm tra hàng tồn và bấm nút tải Excel tương ứng. Có thể tự thêm sản phẩm mua ngoài bằng cách mở phần mở rộng bên dưới.

#### 9. 🔄 Chuyển Hàng Giữa Chi Nhánh (Inter-branch Transfer)
* **Mục đích:** Ghi nhận việc điều chuyển hàng hóa nội bộ giữa các chi nhánh hoặc giữa chi nhánh với kho tổng.
* **Cần làm gì:** Chọn đơn vị gửi, đơn vị nhận, điền tên nhân viên thực hiện và danh sách sản phẩm cần chuyển.

#### 10. 📋 Chi Nhánh Đặt Hàng Kho Tổng (Order Request)
* **Mục đích:** Cho phép chi nhánh gửi yêu cầu đặt hàng về kho tổng hoặc để Admin duyệt các đơn hàng.
* **Cần làm gì:** Chi nhánh điền đơn đặt hàng gửi đi; Admin vào mục này để kiểm tra danh sách chờ duyệt từ các chi nhánh.

#### 11. 📜 Xem Lịch Sử Giao Dịch & Sơ Chế (History)
* **Mục đích:** Tra cứu lại toàn bộ nhật ký hoạt động cũ của hệ thống.
* **Cần làm gì:** Chuyển qua lại giữa các tab để xem báo cáo và bấm nút tải Excel tương ứng để lưu trữ."""
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
        "m_branch_import": "Branch Local Import",
        "m_edit": "Edit Opening Stock",
        "m_add": "Add New Item",
        "m_process": "Main Stock Processing Log",
        "m_branch_process": "Branch Processing & Waste Log",
        "m_distribute": "Branch Distribution",
        "m_branch_inv": "Branch Local Inventory",
        "m_transfer": "Inter-branch Transfer",
        "m_order": "Branch Order Request",
        "m_history": "View Transaction & Processing History",
        "m_guide": "User Guide",
        "total_items": "Total Items",
        "low_stock_warn": "Low Stock Alerts",
        "total_branches": "Active Branches",
        "main_stock_table": "Main Inventory Stock Table",
        "tip_overview": "💡 **What does this do?** Quickly view total items, low stock alerts, and branch status.",
        "tip_import": "💡 **What does this do?** Record newly imported raw materials from suppliers into the main stock.",
        "tip_branch_import": "💡 **What does this do?** Record newly imported raw materials directly into the branch inventory.",
        "tip_edit": "💡 **What does this do?** Adjust the opening stock quantity of items (Admin only, requires security password).",
        "tip_add": "💡 **What does this do?** Register a new product into the main inventory catalog.",
        "tip_process": "💡 **What does this do?** Record the processing of raw materials into finished goods and waste/loss tracking in main stock.",
        "tip_branch_process": "💡 **What does this do?** Record raw material processing at the branch, automatically deducting raw stock and adding finished goods.",
        "tip_distribute": "💡 **What does this do?** Allocate and distribute goods from the main stock down to branch locations.",
        "tip_branch_inv": "💡 **What does this do?** Manage internal stock levels at each specific branch location.",
        "tip_transfer": "💡 **What does this do?** Record internal stock transfers between branches or between branches and the main stock.",
        "tip_order": "💡 **What does this do?** Manage order requests sent from branches to the main stock for approval.",
        "tip_history": "💡 **What does this do?** Review full activity history, distribution logs, processing logs, and internal transfers with Excel export options.",
        "guide_content": """### 📖 System User Guide & Manual

#### 1. 📊 Overview & Stock Alerts
* **Purpose:** Quickly check total items, low stock warnings ($\le 5$), and branch status.
* **Action:** Review the main table and click **Download Main Stock (Excel)** to export reports.

#### 2. 📥 Main Stock Import
* **Purpose:** Record new raw materials imported from suppliers into the central warehouse.
* **Action:** Select the product, enter the quantity, specify the supplier, and click **Confirm Import**.

#### 3. 📥 Branch Local Import
* **Purpose:** Record newly imported raw materials directly into the branch inventory.
* **Action:** Select branch, item, quantity, and local supplier.

#### 4. 📝 Edit Opening Stock
* **Purpose:** Adjust the baseline initial opening stock of products.
* **Action:** Admin only, requires security password (`264221`) to confirm changes.

#### 5. ➕ Add New Item
* **Purpose:** Register a brand new product into the inventory system.
* **Action:** Enter product name, select unit, set opening stock, and supply source.

#### 6. 🔪 Processing & Waste Log
* **Purpose:** Manage transformation of raw materials into finished goods and track waste/loss.
* **Action:** Select date, batch ID, raw material used, output product, and waste quantity.

#### 7. 🚚 Branch Distribution
* **Purpose:** Allocate items from the main stock to individual branches.
* **Action:** Choose target branch, select items and quantities. Stock is automatically deducted from main and added to branch.

#### 8. 🏪 Branch Local Inventory
* **Purpose:** Track internal stock levels at each branch location.
* **Action:** Check stock or download Excel reports. Branch staff can add local purchased items using the expander below.

#### 9. 🔄 Inter-branch Transfer
* **Purpose:** Record internal stock movements between branches or between branch and central warehouse.
* **Action:** Select sender, receiver, staff name, and list of items to transfer.

#### 10. 📋 Branch Order Request
* **Purpose:** Allow branches to request items from central stock or for admins to review orders.
* **Action:** Branches submit requests; Admin reviews and processes pending orders here.

#### 11. 📜 Transaction & Processing History
* **Purpose:** Look up historical logs of all system activities.
* **Action:** Switch between tabs and download respective Excel reports."""
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
        "m_branch_import": "Egységek Bevételezése",
        "m_edit": "Nyitókészlet Szerkesztése",
        "m_add": "Új Termék Hozzáadása",
        "m_process": "Központi Feldolgozási Napló",
        "m_branch_process": "Egység Feldolgozási és Veszteség Napló",
        "m_distribute": "Kiosztás Egységeknek",
        "m_branch_inv": "Egységek Saját Készlete",
        "m_transfer": "Egységek közötti átadás",
        "m_order": "Egységek Rendelése a Központból",
        "m_history": "Tranzakció és Feldolgozási Előzmények",
        "m_guide": "Használati Útmutató",
        "total_items": "Összes termék",
        "low_stock_warn": "Alacsony készlet riasztás",
        "total_branches": "Aktív egységek",
        "main_stock_table": "Központi Készlet Táblázat",
        "tip_overview": "💡 **Mire való ez a menüpont?** Gyors áttekintés a termékek számáról, az alacsony készlet riasztásokról és az egységek állapotáról.",
        "tip_import": "💡 **Mire való ez a menüpont?** Beszállítóktól érkező alapanyagok bevételezése a központi készletbe.",
        "tip_branch_import": "💡 **Mire való ez a menüpont?** Új alapanyagok bevételezésének rögzítése közvetlenül az egység raktárába.",
        "tip_edit": "💡 **Mire való ez a menüpont?** A nyitókészlet mennyiségének módosítása (Csak adminisztrátornak, biztonsági jelszó szükséges).",
        "tip_add": "💡 **Mire való ez a menüpont?** Új termék felvétele a központi készlet katalógusába.",
        "tip_process": "💡 **Mire való ez a menüpont?** Nyersanyagok feldolgozása késztermékké a központban, valamint a veszteség rögzítése.",
        "tip_branch_process": "💡 **Mire való ez a menüpont?** Nyersanyagok feldolgozásának rögzítése az egységben, alapanyag levonással és késztermék hozzáadásával.",
        "tip_distribute": "💡 **Mire való ez a menüpont?** Áruk kiosztása és átszállítása a központi raktárból az egységekbe.",
        "tip_branch_inv": "💡 **Mire való ez a menüpont?** Az egyes egységek belső készletszintjének kezelése.",
        "tip_transfer": "💡 **Mire való ez a menüpont?** Belső árumozgások és átadások rögzítése az egységek között.",
        "tip_order": "💡 **Mire való ez a menüpont?** Az egységek által a központ felé küldött rendelési kérelmek kezelése és jóváhagyása.",
        "tip_history": "💡 **Mire való ez a menüpont?** Teljes tranzakciós előzmények, kiosztási naplók, feldolgozási naplók és belső átadások megtekintése Excel export lehetőséggel.",
        "guide_content": """### 📖 Rendszer Használati Útmutató

#### 1. 📊 Áttekintés és Készletriasztások (Overview)
* **Cél:** Gyorsan ellenőrizheti az összes tételt, az alacsony készletű termékeket ($\le 5$) és az aktív egységeket.
* **Teendő:** Ellenőrizze a táblázatot, és kattintson a **Központi készlet letöltése** gombra a jelentés exportálásához.

#### 2. 📥 Központi Készlet Bevételezés (Main Import)
* **Cél:** Beszállítótól érkező új alapanyagok rögzítése a központi raktárba.
* **Teendő:** Válassza ki a terméket, adja meg a mennyiséget és a szállítót, majd kattintson a megerősítésre.

#### 3. 📥 Egységek Bevételezése (Branch Local Import)
* **Cél:** Új alapanyagok bevételezésének rögzítése közvetlenül az egység raktárába.
* **Teendő:** Válassza ki az egységet, terméket, mennyiséget és a helyi beszállítót.

#### 4. 📝 Nyitókészlet Szerkesztése (Edit Opening Stock)
* **Cél:** A termékek eredeti alapmennyiségének módosítása.
* **Teendő:** Csak Admin számára, biztonsági jelszó (`264221`) szükséges a módosításhoz.

#### 5. ➕ Új Termék Hozzáadása (Add New Item)
* **Cél:** Új árucikk felvétele a nyilvántartási rendszerbe.
* **Teendő:** Adja meg a termék nevét, mértékegységét, nyitómennyiségét és a forrást.

#### 6. 🔪 Feldolgozási Napló & Veszteség (Processing Log)
* **Cél:** Nyersanyagok késztermékké alakításának és a hulladéknak a nyomon követése.
* **Teendő:** Válassza ki a dátumot, tételszámot, felhasznált alapanyagot, előállított terméket és a veszteséget.

#### 7. 🚚 Kiosztás Egységeknek (Branch Distribution)
* **Cél:** Áruk elosztása a központi raktárból az egyes egységek felé.
* **Teendő:** Válassza ki az egységet, a termékeket és a mennyiséget. A rendszer automatikusan vonja le a központból.

#### 8. 🏪 Egységek Saját Készlete (Branch Local Inventory)
* **Cél:** Az egyes egységek belső készletének nyomon követése.
* **Teendő:** Ellenőrizze a készletet és töltse le az Excel jelentést.

#### 9. 🔄 Egységek Közötti Átadás (Inter-branch Transfer)
* **Cél:** Belső árumozgások rögzítése az egységek vagy a központ között.
* **Teendő:** Válassza ki a küldőt, fogadót, a dolgozót és az átadandó tételeket.

#### 10. 📋 Egységek Rendelése (Order Request)
* **Cél:** Egységek rendelési kérelmeinek küldése és adminisztrátori jóváhagyása.
* **Teendő:** Az egységek leadják a rendelést; az Admin itt tudja jóváhagyni azokat.

#### 11. 📜 Tranzakciós Előzmények (History)
* **Cél:** Korábbi tevékenységek és naplók áttekintése.
* **Teendő:** Váltson a fülek között és töltse le az Excel jelentéseket."""
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
            pending_order_count = len(df_check_order[df_check_order["Status"].str.contains("chờ|pending|vár", case=False, na=False)])
    except:
        pass

st.sidebar.title(T["menu"])

order_menu_label = T["m_order"]
if st.session_state.role == "Admin" and pending_order_count > 0:
    order_menu_label = f"🔴 {T['m_order']} ({pending_order_count})"

if st.session_state.role == "Admin":
    menu_options = {
        T["m_overview"]: "overview",
        T["m_import"]: "import",
        T["m_branch_import"]: "branch_import",
        T["m_edit"]: "edit",
        T["m_add"]: "add",
        T["m_process"]: "process",
        T["m_branch_process"]: "branch_process",
        T["m_distribute"]: "distribute",
        T["m_branch_inv"]: "branch_inv",
        T["m_transfer"]: "transfer",
        order_menu_label: "order",
        T["m_history"]: "history",
        T["m_guide"]: "guide"
    }
else:
    menu_options = {
        T["m_branch_inv"]: "branch_inv",
        T["m_branch_import"]: "branch_import",
        T["m_branch_process"]: "branch_process",
        T["m_transfer"]: "transfer",
        T["m_order"]: "order",
        T["m_guide"]: "guide"
    }

selected_menu_label = st.sidebar.radio(
    "",
    list(menu_options.keys()),
    label_visibility="collapsed"
)
choice = menu_options[selected_menu_label]

st.title(T["title"])

main_stock_df, processing_df = load_data()
main_stock_df = calculate_closing_stock(main_stock_df)
branch_data_dict = load_branch_data()

if choice == "overview":
    st.subheader(T["m_overview"])
    st.info(T["tip_overview"])
    
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

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        main_stock_df.to_excel(writer, sheet_name='Tong_Ket_Kho_Tong', index=False)
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Tải xuống bảng tồn kho tổng (Excel)",
        data=excel_data,
        file_name=f"Tong_Ket_Ton_Kho_Tong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

elif choice == "import":
    st.subheader(T["m_import"])
    st.info(T["tip_import"])
    
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

elif choice == "branch_import":
    st.subheader(T["m_branch_import"])
    st.info(T["tip_branch_import"])
    
    active_branch = st.session_state.branch_name if st.session_state.role == "Branch" else st.selectbox("Chọn chi nhánh nhập hàng:", BRANCH_LIST)
    current_b_df = branch_data_dict.get(active_branch, pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "StockQty", "ImportedQty", "UsedQty", "Note"]))
    
    with st.form(f"branch_import_form_{active_branch}"):
        b_item_list = current_b_df["ItemName"].tolist() if not current_b_df.empty else []
        selected_b_item = st.selectbox("Chọn sản phẩm cần nhập tại chi nhánh:", b_item_list) if b_item_list else None
        b_import_qty = st.number_input("Số lượng nhập thêm:", min_value=0.0, step=1.0)
        b_import_source = st.text_input("Nhà cung cấp / Nguồn gốc:", value="Nhà cung cấp địa phương")
        submitted_b_import = st.form_submit_button("Xác Nhận Nhập Hàng Chi Nhánh")
        
        if submitted_b_import and selected_b_item:
            b_idx = current_b_df[current_b_df["ItemName"] == selected_b_item].index
            if not b_idx.empty:
                current_b_df.loc[b_idx, "StockQty"] += b_import_qty
                current_b_df.loc[b_idx, "ImportedQty"] += b_import_qty
                current_b_df.loc[b_idx, "Note"] = f"Nhập hàng từ: {b_import_source}"
                
                branch_data_dict[active_branch] = current_b_df
                save_branch_data(branch_data_dict)
                
                st.success(f"Nhập hàng thành công cho chi nhánh **{active_branch}**!")
                st.rerun()

elif choice == "edit":
    st.subheader(T["m_edit"])
    st.info(T["tip_edit"])
    
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
    st.info(T["tip_add"])
    
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
    st.info(T["tip_process"])
    
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

elif choice == "branch_process":
    st.subheader(T["m_branch_process"])
    st.info(T["tip_branch_process"])
    
    active_branch = st.session_state.branch_name if st.session_state.role == "Branch" else st.selectbox("Chọn chi nhánh thực hiện sơ chế:", BRANCH_LIST)
    current_b_df = branch_data_dict.get(active_branch, pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "StockQty", "ImportedQty", "UsedQty", "Note"]))
    
    with st.form("branch_process_form"):
        bp_date = st.date_input("Ngày sơ chế:", datetime.now())
        bp_raw_list = current_b_df["ItemName"].tolist() if not current_b_df.empty else []
        bp_raw = st.selectbox("Chọn nguyên liệu thô tại kho chi nhánh:", bp_raw_list)
        bp_used_qty = st.number_input("Số lượng nguyên liệu sử dụng:", min_value=0.0, step=1.0)
        bp_finished = st.text_input("Tên thành phẩm thu được:")
        bp_produced_qty = st.number_input("Số lượng thành phẩm thu về:", min_value=0.0, step=1.0)
        bp_waste = st.number_input("Hao hụt / Phế phẩm:", min_value=0.0, step=1.0)
        bp_note = st.text_area("Ghi chú quy trình:")
        
        submitted_bp = st.form_submit_button("Xác Nhận Sơ Chế Tại Chi Nhánh")
        
        if submitted_bp:
            if not bp_raw or bp_used_qty <= 0:
                st.error("Vui lòng chọn nguyên liệu hợp lệ và nhập số lượng sử dụng lớn hơn 0!")
            else:
                raw_idx = current_b_df[current_b_df["ItemName"] == bp_raw].index
                if raw_idx.empty:
                    st.error("Không tìm thấy nguyên liệu trong kho chi nhánh!")
                else:
                    current_stock = current_b_df.loc[raw_idx, "StockQty"].values[0]
                    if bp_used_qty > current_stock:
                        st.error(f"Số lượng dùng ({bp_used_qty}) vượt quá tồn kho hiện tại của chi nhánh ({current_stock})!")
                    else:
                        current_b_df.loc[raw_idx, "StockQty"] -= bp_used_qty
                        current_b_df.loc[raw_idx, "UsedQty"] += bp_used_qty
                        
                        fin_idx = current_b_df[current_b_df["ItemName"] == bp_finished].index
                        if not fin_idx.empty and bp_finished.strip():
                            current_b_df.loc[fin_idx, "StockQty"] += bp_produced_qty
                            current_b_df.loc[fin_idx, "ImportedQty"] += bp_produced_qty
                        elif bp_finished.strip():
                            fin_auto_id = f"BP-{len(current_b_df)+1:03d}"
                            unit_val = current_b_df.loc[raw_idx, "Unit"].values[0]
                            new_fin_row = {
                                "ItemID": fin_auto_id, "ItemName": bp_finished, "Unit": unit_val,
                                "StockQty": bp_produced_qty, "ImportedQty": bp_produced_qty, "UsedQty": 0.0, "Note": f"Sơ chế từ {bp_raw}"
                            }
                            current_b_df = pd.concat([current_b_df, pd.DataFrame([new_fin_row])], ignore_index=True)
                        
                        branch_data_dict[active_branch] = current_b_df
                        save_branch_data(branch_data_dict)
                        
                        bp_record = {
                            "Date": bp_date.strftime("%Y-%m-%d"),
                            "Branch": active_branch,
                            "RawMaterial": bp_raw,
                            "UsedQuantity": bp_used_qty,
                            "FinishedProduct": bp_finished,
                            "ProducedQuantity": bp_produced_qty,
                            "WasteLoss": bp_waste,
                            "Note": bp_note
                        }
                        df_bp_hist = pd.read_csv(BRANCH_PROC_FILE)
                        df_bp_hist = pd.concat([df_bp_hist, pd.DataFrame([bp_record])], ignore_index=True)
                        df_bp_hist.to_csv(BRANCH_PROC_FILE, index=False)
                        
                        st.success(f"Đã ghi nhận sơ chế thành công cho chi nhánh **{active_branch}**!")
                        st.rerun()

elif choice == "distribute":
    st.subheader(T["m_distribute"])
    st.info(T["tip_distribute"])
    
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
    st.info(T["tip_branch_inv"])
    
    if st.session_state.role == "Admin":
        active_branch = st.selectbox("Chọn chi nhánh để quản lý kho:", BRANCH_LIST)
    else:
        active_branch = st.session_state.branch_name
        st.markdown(f"Đang xem kho riêng của chi nhánh: **{active_branch}**")

    current_b_df = branch_data_dict.get(active_branch, pd.DataFrame(columns=["ItemID", "ItemName", "Unit", "StockQty", "ImportedQty", "UsedQty", "Note"]))
    
    st.markdown(f"### Kho Tồn Của Chi Nhánh: {active_branch}")
    st.dataframe(current_b_df, use_container_width=True)

    output_b = io.BytesIO()
    with pd.ExcelWriter(output_b, engine='openpyxl') as writer:
        current_b_df.to_excel(writer, sheet_name=active_branch, index=False)
    excel_b_data = output_b.getvalue()
    st.download_button(
        label=f"📥 Tải xuống kho chi nhánh {active_branch} (Excel)",
        data=excel_b_data,
        file_name=f"Kho_Chi_Nhanh_{active_branch}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("---")
    
    # CHỈ CHO PHÉP ADMIN SỬA TỒN KHO CHI NHÁNH KHI NHẬP MẬT KHẨU HOẶC ĐĂNG NHẬP ADMIN
    if st.session_state.role == "Admin":
        with st.expander("🛠️ Sửa Tồn Kho / Đầu Kỳ Tại Chi Nhánh Này (Dành cho Admin)"):
            with st.form(f"edit_branch_stock_form_{active_branch}"):
                b_item_list_edit = current_b_df["ItemName"].tolist() if not current_b_df.empty else []
                selected_edit_b_item = st.selectbox("Chọn sản phẩm cần sửa tồn kho:", b_item_list_edit) if b_item_list_edit else None
                new_b_stock_val = st.number_input("Số lượng tồn kho mới:", min_value=0.0, step=1.0)
                edit_b_pwd = st.text_input("Nhập mật khẩu bảo mật Admin:", type="password")
                submitted_edit_b = st.form_submit_button("Xác Nhận Sửa Tồn Kho Chi Nhánh")
                
                if submitted_edit_b:
                    if edit_b_pwd != SECRET_ACTION_PWD:
                        st.error("Sai mật khẩu xác nhận!")
                    elif selected_edit_b_item:
                        b_idx = current_b_df[current_b_df["ItemName"] == selected_edit_b_item].index
                        if not b_idx.empty:
                            current_b_df.loc[b_idx, "StockQty"] = new_b_stock_val
                            current_b_df.loc[b_idx, "Note"] = "Admin điều chỉnh tồn kho"
                            branch_data_dict[active_branch] = current_b_df
                            save_branch_data(branch_data_dict)
                            st.success(f"Đã cập nhật tồn kho chi nhánh **{active_branch}** thành công!")
                            st.rerun()
    
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
                    "ItemID": b_auto_id, "ItemName": b_new_name, "Unit": b_new_unit,
                    "StockQty": b_new_stock, "ImportedQty": b_new_stock, "UsedQty": 0.0, "Note": b_new_note
                }
                current_b_df = pd.concat([current_b_df, pd.DataFrame([new_b_row])], ignore_index=True)
                branch_data_dict[active_branch] = current_b_df
                save_branch_data(branch_data_dict)
                st.success(f"Đã thêm sản phẩm thành công!")
                st.rerun()

elif choice == "transfer":
    st.subheader(T["m_transfer"])
    st.info(T["tip_transfer"])
    
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
    st.info(T["tip_order"])
    
    if st.session_state.role == "Branch":
        st.markdown(f"Giao diện đặt hàng cho chi nhánh: **{st.session_state.branch_name}**")
        with st.form("branch_order_form"):
            o_item = st.selectbox("Chọn mặt hàng cần đặt:", main_stock_df["ItemName"].tolist() if not main_stock_df.empty else [])
            o_qty = st.number_input("Số lượng đặt:", min_value=0.0, step=1.0)
            o_unit = st.selectbox("Đơn vị:", UNIT_LIST)
            o_note = st.text_input("Ghi chú / Yêu cầu thêm:")
            btn_submit_order = st.form_submit_button("Gửi Đơn Đặt Hàng")
            
            if btn_submit_order and o_qty > 0:
                ord_df = pd.read_csv(ORDER_FILE)
                new_ord = {
                    "OrderDate": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Branch": st.session_state.branch_name,
                    "ItemName": o_item, "Unit": o_unit,
                    "Quantity": o_qty, "Status": "Chờ duyệt", "Note": o_note
                }
                ord_df = pd.concat([ord_df, pd.DataFrame([new_ord])], ignore_index=True)
                ord_df.to_csv(ORDER_FILE, index=False)
                st.success("Đã gửi đơn đặt hàng tới kho tổng thành công!")
                st.rerun()
    else:
        st.markdown("### Danh Sách Đơn Đặt Hàng Từ Các Chi Nhánh (Admin Quản Lý)")
        if os.path.exists(ORDER_FILE):
            ord_df = pd.read_csv(ORDER_FILE)
            st.dataframe(ord_df, use_container_width=True)

elif choice == "history":
    if st.session_state.role != "Admin":
        st.error("🚫 Bạn không có quyền truy cập vào mục lịch sử này. Khu vực này chỉ dành cho Admin.")
    else:
        st.subheader(T["m_history"])
        st.info(T["tip_history"])

        tab1, tab2, tab3, tab4 = st.tabs([
            "📦 Cấp Hàng Kho Tổng",
            "🔪 Sơ Chế Kho Tổng",
            "🔪 Sơ Chế Chi Nhánh",
            "🔄 Chuyển Nội Bộ"
        ])

        with tab1:
            st.markdown("### Lịch sử cấp hàng cho các chi nhánh")
            if os.path.exists(EXPORT_FILE):
                df_exp = pd.read_csv(EXPORT_FILE)
                st.dataframe(df_exp, use_container_width=True)
                if not df_exp.empty:
                    out_exp = io.BytesIO()
                    with pd.ExcelWriter(out_exp, engine='openpyxl') as writer:
                        df_exp.to_excel(writer, sheet_name='LichSuCapHang', index=False)
                    st.download_button(
                        label="📥 Tải xuống lịch sử cấp hàng (Excel)",
                        data=out_exp.getvalue(),
                        file_name=f"Lich_Su_Cap_Hang_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("Chưa có dữ liệu.")

        with tab2:
            st.markdown("### Lịch sử sơ chế kho tổng")
            if os.path.exists(DATA_FILE):
                try:
                    df_main_proc = pd.read_excel(DATA_FILE, sheet_name="ProcessingLog")
                    if not df_main_proc.empty:
                        st.dataframe(df_main_proc, use_container_width=True)
                        out_proc = io.BytesIO()
                        with pd.ExcelWriter(out_proc, engine='openpyxl') as proc_writer:
                            df_main_proc.to_excel(proc_writer, sheet_name='SoChe_KhoTong', index=False)
                        st.download_button(
                            label="📥 Tải xuống lịch sử sơ chế kho tổng (Excel)",
                            data=out_proc.getvalue(),
                            file_name=f"Lich_Su_So_Che_Kho_Tong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.info("Chưa có dữ liệu.")
                except:
                    st.info("Chưa có dữ liệu.")

        with tab3:
            st.markdown("### Lịch sử sơ chế hao hụt tại các chi nhánh")
            if os.path.exists(BRANCH_PROC_FILE):
                df_bp_hist = pd.read_csv(BRANCH_PROC_FILE)
                if not df_bp_hist.empty:
                    st.dataframe(df_bp_hist, use_container_width=True)
                    out_bp = io.BytesIO()
                    with pd.ExcelWriter(out_bp, engine='openpyxl') as bp_writer:
                        df_bp_hist.to_excel(bp_writer, sheet_name='SoChe_ChiNhanh', index=False)
                    st.download_button(
                        label="📥 Tải xuống lịch sử sơ chế chi nhánh (Excel)",
                        data=out_bp.getvalue(),
                        file_name=f"Lich_Su_So_Che_Chi_Nhanh_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("Chưa có dữ liệu sơ chế từ chi nhánh nào.")
            else:
                st.info("Chưa có dữ liệu sơ chế chi nhánh.")

        with tab4:
            st.markdown("### Lịch sử chuyển hàng nội bộ")
            if os.path.exists(TRANSFER_FILE):
                df_trans = pd.read_csv(TRANSFER_FILE)
                st.dataframe(df_trans, use_container_width=True)
                if not df_trans.empty:
                    out_trans = io.BytesIO()
                    with pd.ExcelWriter(out_trans, engine='openpyxl') as writer:
                        df_trans.to_excel(writer, sheet_name='LichSuChuyenNoiBo', index=False)
                    st.download_button(
                        label="📥 Tải xuống lịch sử chuyển hàng (Excel)",
                        data=out_trans.getvalue(),
                        file_name=f"Lich_Su_Chuyen_Hang_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("Chưa có dữ liệu.")

elif choice == "guide":
    st.markdown(T["guide_content"])
