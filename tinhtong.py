import os
import pandas as pd
from openpyxl import load_workbook, Workbook

ROOT_PATH = r"\\10.147.32.1\MA_Div\Data_Link"

#cac nam FY
ACTUAL_ROOT = os.path.join(ROOT_PATH, "CostDX-BM-TH-ACT", "XUẤT KHO FY25")
#tong ket
BUDGET_PATH = os.path.join(ROOT_PATH, "BUDGET FY25")
#luu tam
TEMP_PATH = os.path.join(ROOT_PATH, "Temp_Result")

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
#----------------------------------------------
# P1: Tinh tong + ghi file tam
#----------------------------------------------
                  
def process_actual(fiscal_year = "FY25"):
    #actual_folder = os.path.join(ACTUAL_PATH, f"XUẤT T{fiscal_year}")
    if not os.path.exists(ACTUAL_ROOT):
        #khong tim thay thu muc
        return
    
    os.makedirs(TEMP_PATH, exist_ok=True)
    months = sorted ([f for f in os.listdir(ACTUAL_ROOT) if f in MONTH_MAPPING], key=lambda x: list(MONTH_MAPPING.keys()).index(x))
    #months = sorted ([f for f in os.listdir(ACTUAL_ROOT) if re.match(r"Tháng\s+\d{2}\.2025", f)])

    #yearly_data = {}  # {product: {month: {cat: value}}}
    all_products = set()

    for month_folder in months:
        month_path = os.path.join(ACTUAL_ROOT, month_folder)
        excel_file = [f for f in os.listdir(month_path) if re.match(r"1\.\s*Theo\s*dõi\s*Xuất\s*kho\s*T\d+\.xlsx", f)]
        if not excel_file:
            continue

        excel_file = excel_file[0]
        file_path = os.path.join(month_path, excel_file)
        #excel = pd.ExcelFile(file_path)
        month_num = re.search(r'T(\d+)', excel_file).group(1)
        sheet_name = f"XUẤT T{month_num}" #chon sheet chu xuat

        try:
            if sheet_name in pd.ExcelFile(file_path).sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="M,J,W")
                df.columns = ["Mã Line", "Tổng tiền", "Phân loại"]
                df_valid = df[df["Mã Line"] != 312].dropna(subset=["Mã Line"])
                all_products.update(df_valid["Mã Line"].astype(str).str.strip().unique())
        except:
            continue
    
    yearly_data = {}
    for product_idx, product in enumerate(sorted(all_products), 1):
        for mounth_folder in months:
            month_path = os.path.join(ACTUAL_ROOT, month_folder)
            excel_files = [f for f in os.listdir(month_path) if re.match(r"1\.\s*Theo\s*dõi\s*Xuất\s*kho\s*T\d+\.xlsx", f)]

            if not excel_files:
                yearly_data.setdefault(product, {}).setdefault(month_folder, {c: 0 for c in CATEGORIES})
                continue

            excel_file = excel_files[0]
            file_path = os.path.join(month_path, excel_file)
            month_name = f"XUẤT T{month_num}"

            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="M,J,W")
                df.columns = ["Mã Line", "Tổng tiền", "Phân loại"]
                
                # CHỈ LẤY DÒNG CỦA MÃ NÀY
                df_product = df[df["Mã Line"] == float(product)]
                df_product = df_product[df_product["Mã Line"] != 312]
                
                df_product["CatKey"] = df_product["Phân loại"].apply(normalize_category)
                df_product = df_product.dropna(subset=["CatKey"])
                
                # Tính tổng riêng từng phân loại
                month_totals = {c: 0 for c in CATEGORIES}
                if not df_product.empty:
                    grouped = df_product.groupby("CatKey")["Tổng tiền"].sum().reset_index()
                    for _, row in grouped.iterrows():
                        cat = row["CatKey"]
                        if cat in CATEGORIES:
                            month_totals[cat] = float(row["Tổng tiền"])
                
                yearly_data.setdefault(product, {}).setdefault(month_folder, month_totals)
                
            except Exception as e:
                yearly_data.setdefault(product, {}).setdefault(month_folder, {c: 0 for c in CATEGORIES})

        month_map = yearly_data[product]
        wb = Workbook()
        ws = wb.active
        ws.title = f"Dữ liệu {fiscal_year}"
        #header: Month | BM | PM | CM | TH
        ws.append(["Month"] + CATEGORIES)
        
        for month_folder in sorted(months):
            month_data = month_map.get(month_folder, {c: 0 for c in CATEGORIES})
            #lay rieng tung phan loai
            bm_value = month_data.get("BM", 0)
            pm_value = month_data.get("PM", 0)
            cm_value = month_data.get("CM", 0)
            th_value = month_data.get("TH", 0)

            #tao dong: [Thang, bm_tong, pm_tong, cm_tong, th_tong]
            row_data = [month_folder, bm_value, pm_value, cm_value, th_value]
            ws.append(row_data)

        #luu file row.<ma>.FY25.xlsx
        filename = os.path.join(TEMP_PATH, f"row.{product}.{fiscal_year}.xlsx")
        wb.save(filename)

#----------------------------------------------
# P2: cap nhat du lieu vao file budget
#----------------------------------------------

def update_budget(fiscal_year="FY25"): 
    #budget_folder = os.path.join(BUDGET_PATH, fiscal_year)
    if not os.path.exists(BUDGET_PATH):
        return
    
    temp_files = [f for f in os.listdir(TEMP_PATH) if f.endswith(f".{fiscal_year}.xlsx")]

    for temp_file in temp_files:
        product = temp_file.split(".")[1]
        temp_path = os.path.join(TEMP_PATH, temp_file)
        df_temp = pd.read_excel(temp_path)

        budget_file = os.path.join(BUDGET_PATH, f"{product}.25.xlsx")
        if not os.path.exists(budget_file):
            continue

        wb = load_workbook(budget_file)

        for idx, row in df_temp.iterrows():
            month_folder = str(row["Month"])
            budget_sheet = MONTH_MAPPING.get(month_folder)

            if budget_sheet and budget_sheet in wb.sheetnames:
                ws = wb[budget_sheet]

                for r in range(2, ws.max_row + 1):
                    item = str(ws[f"F{r}"].value).strip().lower() if ws[f"F{r}"].value else ""
                    remark_raw = str(ws[f"N{r}"].value).strip() if ws[f"N{r}"].value else ""

                    if item == "Xuất kho":
                        remark_key = normalize_category(remark_raw)
                        if remark_key and remark_key in CATEGORIES:
                            value = float(row[remark_key])
                            ws[f"I{r}"].value = value
        
        wb.save(budget_file)

def main():
    process_actual("FY25")

    update_budget("FY25")

if __name__ == "__main__":
    main()
