import os
import pandas as pd
import re
from openpyxl import load_workbook, Workbook


#ROOT_PATH = r"\\10.147.32.1\MA_Div\Data_Link"
#ROOT_PATH = r"Documents"
ROOT_PATH = r"c:\Users\vnintern09\Documents"

#cac nam FY
#ACTUAL_ROOT = os.path.join(ROOT_PATH, "CostDX-BM-TH-ACT", "XUẤT KHO FY25")
#tong ket
#BUDGET_PATH = os.path.join(ROOT_PATH, "BUDGET FY25")
#luu tam
#TEMP_PATH = os.path.join(ROOT_PATH, "Temp_Result")

try:
    # Khởi tạo đường dẫn
    ACTUAL_ROOT = os.path.join(ROOT_PATH, "CostDX-BM-TH-ACT", "XUẤT KHO FY25")
    BUDGET_PATH = os.path.join(ROOT_PATH, "BUDGET FY25")
    TEMP_PATH   = os.path.join(ROOT_PATH, "Temp_Result") 

    # Kiểm tra tồn tại từng thư mục
    if not os.path.exists(ACTUAL_ROOT):
        raise FileNotFoundError(f"Không tìm thấy thư mục Actual: {ACTUAL_ROOT}")
    if not os.path.exists(BUDGET_PATH):
        raise FileNotFoundError(f"Không tìm thấy thư mục Budget: {BUDGET_PATH}")
    # Temp có thể chưa tồn tại → tự tạo
    if not os.path.exists(TEMP_PATH):
        os.makedirs(TEMP_PATH, exist_ok=True)

except Exception as e:
    print(f"Lỗi cấu hình đường dẫn dữ liệu: {e}")


#danh sach phan loai san pham
#CATEGORIES = ["BM", "PM", "CM", "TH", "Tiêu hao", "TIÊU HAO"]
CATEGORIES = ["BM", "PM", "CM", "TH"]
#danh sach cac thang nam tai chinh
MONTH_MAPPING = {
    "Tháng 04.2025": "APR",
    "Tháng 05.2025": "MAY",
    "Tháng 06.2025": "JUN",
    "Tháng 07.2025": "JUL",
    "Tháng 08.2025": "AUG",
    "Tháng 09.2025": "SEP",
    "Tháng 10.2025": "OCT",
    "Tháng 11.2025": "NOV",
    "Tháng 12.2025": "DEC",
    "Tháng 01.2026": "JAN",
    "Tháng 02.2026": "FEB",
    "Tháng 03.2026": "MAR"
}

def normalize_category(raw_value):
    """'Tiêu hao' → 'TH'"""
    if pd.isna(raw_value):
        return None
    val = str(raw_value).strip().upper()
    if "TIÊU HAO" in val:
        return "TH"
    return val if val in CATEGORIES else None
#-----------------------------
# P1: Tinh tong + ghi file tam
#-----------------------------
                  
def process_actual(fiscal_year = "FY25"):
    print("phan 1 bat dau")
    #actual_folder = os.path.join(ACTUAL_PATH, f"XUẤT T{fiscal_year}")
    if not os.path.exists(ACTUAL_ROOT):
        #khong tim thay thu muc
        print(f"khong co {ACTUAL_ROOT}")
        return
    
    os.makedirs(TEMP_PATH, exist_ok=True)

    #danh sách theo thứ tự năm tài chính
    months = sorted (
        [f for f in os.listdir(ACTUAL_ROOT) if f in MONTH_MAPPING], 
        key=lambda x: list(MONTH_MAPPING.keys()).index(x)
    )
    #months = sorted ([f for f in os.listdir(ACTUAL_ROOT) if re.match(r"Tháng\s+\d{2}\.2025", f)])

    #yearly_data = {}  # {product: {month: {cat: value}}}

    #quét tất cả các tháng từ 312
    all_products = set()

    for month_folder in months: 
        month_path = os.path.join(ACTUAL_ROOT, month_folder)
        #tìm theo file "theo dõi xuất kho"
        #excel_file = [f for f in os.listdir(month_path) if re.match(r"1\.\s*Theo\s*dõi\s*Xuất\s*kho\s*T\d+\.xlsx", f)]
        #if not excel_file:
        #    continue

        #excel_file = excel_file[0]
        #file_path = os.path.join(month_path, excel_file)
        excel_file = None
        for f in os.listdir(month_path):
            #dieu kien file .xlsx co chu "theo doi xuat kho" trong name
            if f.endswith(".xlsx") and "Theo dõi Xuất kho" in f:
                excel_file = f
                break
        if not excel_file:
            continue
        file_path = os.path.join(month_path, excel_file)

        #tim sheet chứa XUẤT T
        #month_num = re.search(r'T(\d+)', excel_file).group(1)
        #sheet_name = f"XUẤT T{month_num}" #chon sheet chu xuat
        excel = pd.ExcelFile(file_path)
        sheet_name = None
        for sheet in excel.sheet_names:
            #sheet co chua "XUAT T"
            if "XUẤT T" in sheet.upper():
                sheet_name = sheet
                break
        if not sheet_name:
            continue

        try:
            if sheet_name in pd.ExcelFile(file_path).sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="M,J,W")
                df.columns = ["Mã Line", "Tổng tiền", "Phân loại"]
                df_valid = df[df["Mã Line"] != 312].dropna(subset=["Mã Line"])
                #all_products.update(df_valid["Mã Line"].astype(str).str.strip().unique())
                all_products_list = []
                for value in df_valid["Mã Line"].unique():
                    try:
                        code = int(float(value))
                        if code != 312:
                            all_products_list.append(str(code))
                    except:
                        continue
                all_products.update(set(all_products_list))
        except:
            continue
    print(f"tim {len(all_products)} ma san pham")
    
    yearly_data = {}
    for product_idx, product in enumerate(sorted(all_products), 1):
        print(f"[{product_idx}/{len(all_products)}] ma {product}")

        for month_folder in months:
            month_path = os.path.join(ACTUAL_ROOT, month_folder)
            #excel_files = [f for f in os.listdir(month_path) if re.match(r"1\.\s*Theo\s*dõi\s*Xuất\s*kho\s*T\d+\.xlsx", f)]
            excel_file = None
            for f in os.listdir(month_path):
                if f.endswith(".xlsx") and "Theo dõi Xuất kho" in f:
                    excel_file = f
                    break
            if not excel_file:
                yearly_data.setdefault(product, {}).setdefault(
                    month_folder, {c: 0 for c in CATEGORIES}
                )
                continue
            file_path = os.path.join(month_path, excel_file)

            excel = pd.ExcelFile(file_path)
            sheet_name = None
            for sheet in excel.sheet_names:
                if "XUẤT T" in sheet.upper():
                    sheet_name = sheet
                    break
            if not sheet_name:
                yearly_data.setdefault(product, {}).setdefault(
                    month_folder, {c: 0 for c in CATEGORIES}
                )
                continue
           
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="M,J,W")
                df.columns = ["Mã Line", "Tổng tiền", "Phân loại"]
                
                # CHỈ LẤY DÒNG CỦA MÃ NÀY
                df["Ma_str"] = df["Mã Line"].astype(str).str.strip()
                df_product = df[df["Ma_str"] == product]
                print("===="*10)
                print(df_product)
                print("===="*10)
                df_product = df_product[df_product["Mã Line"] != 312]
                
                df_product["CatKey"] = df_product["Phân loại"].apply(normalize_category)
                df_product = df_product.dropna(subset=["CatKey"])
                
                # Tính tổng riêng từng phân loại
                month_totals = {c: 0 for c in CATEGORIES}
                if not df_product.empty:
                    grouped = df_product.groupby("CatKey")["Tổng tiền"].sum()
                    #grouped = df_product.groupby("CatKey")["Tổng tiền"].sum().reset_index()
                    #for _, row in grouped.iterrows():
                    #    cat = row["CatKey"]
                    #    if cat in CATEGORIES:
                    #        month_totals[cat] = float(row["Tổng tiền"])
                    for cat in CATEGORIES:
                        if cat in grouped.index:
                            month_totals[cat] = float(grouped[cat])
                    print(
                        f"  Ma {product}, {month_folder}: "
                        f"BM={month_totals['BM']:,.0f}, "
                        f"PM={month_totals['PM']:,.0f}, "
                        f"CM={month_totals['CM']:,.0f}, "
                        f"TH={month_totals['TH']:,.0f}"
                    )
                yearly_data.setdefault(product, {}).setdefault(month_folder, month_totals)
                
            except Exception as e:
                yearly_data.setdefault(product, {}).setdefault(
                    month_folder, {c: 0 for c in CATEGORIES}
                )

        month_map = yearly_data[product]
        wb = Workbook()
        ws = wb.active
        ws.title = f"Dữ liệu {fiscal_year}"
        #header: Month | BM | PM | CM | TH
        ws.append(["Month"] + CATEGORIES)
        
        for month_folder in sorted(months):
            month_data = month_map.get(month_folder, {c: 0 for c in CATEGORIES})

            #tao dong: [Thang, bm_tong, pm_tong, cm_tong, th_tong]
            row_data = [
                month_folder,
                month_data.get("BM", 0),
                month_data.get("PM", 0),
                month_data.get("CM", 0),
                month_data.get("TH", 0),
            ]
            ws.append(row_data)

        #luu file row.<ma>.FY25.xlsx
        filename = os.path.join(TEMP_PATH, f"row.{product}.{fiscal_year}.xlsx")
        wb.save(filename)
        print(f"  OK Luu row.{product}.FY25.xlsx")
    print("phan 1 hoan thanh")

#-------------------------------------
# P2: cap nhat du lieu vao file budget
#-------------------------------------

def update_budget(fiscal_year="FY25"): 
    #budget_folder = os.path.join(BUDGET_PATH, fiscal_year)
    temp_files = [f for f in os.listdir(TEMP_PATH) if f.endswith(f".{fiscal_year}.xlsx")]
    print(f"PHAN 2: Cap nhat {len(temp_files)} file tam")

    if not os.path.exists(BUDGET_PATH):
        print(f"Khong co {BUDGET_PATH}")
        return
    
    #temp_files = [f for f in os.listdir(TEMP_PATH) if f.endswith(f".{fiscal_year}.xlsx")]

    for temp_file in temp_files:
        product = temp_file.split(".")[1]
        temp_path = os.path.join(TEMP_PATH, temp_file)
        df_temp = pd.read_excel(temp_path)

        budget_file = os.path.join(BUDGET_PATH, f"{product}.25.xlsx")
        if not os.path.exists(budget_file):
            print(f"  Khong co {budget_file}")
            continue

        wb = load_workbook(budget_file)

        for idx, row in df_temp.iterrows():
            month_folder = str(row["Month"])
            budget_sheet = MONTH_MAPPING.get(month_folder)

            if budget_sheet and budget_sheet in wb.sheetnames:
                ws = wb[budget_sheet]

                for r in range(2, ws.max_row + 1):
                    item = (
                        str(ws[f"F{r}"].value).strip().lower()
                        if ws[f"F{r}"].value
                        else ""
                    )
                    remark_raw = (
                        str(ws[f"N{r}"].value).strip()
                        if ws[f"N{r}"].value
                        else ""
                    )

                    if item == "xuất kho":
                        remark_key = normalize_category(remark_raw)
                        if remark_key and remark_key in CATEGORIES:
                            value = float(row[remark_key])
                            ws[f"I{r}"].value = value
        
        wb.save(budget_file)
        print(f"  OK cap nhat {product}.25.xlsx")
        
    print("PHAN 2 HOAN THANH")


def main():
    process_actual("FY25")
    update_budget("FY25")

if __name__ == "__main__":
    main()       
