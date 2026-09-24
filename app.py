import streamlit as st
import pandas as pd
import datetime
import os
import io
import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
# 1. Page Configuration
st.set_page_config(
    page_title="CITIZEN | CMA Integrated Logistics & Gate Pass System",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Industrial Theme CSS
st.markdown("""
<style>
    .stApp { background-color: #f4f6f9; }
    .top-header {
        background: linear-gradient(135deg, #102a43 0%, #243b53 100%);
        padding: 16px 22px;
        border-radius: 8px;
        color: white;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 18px;
    }
    .top-header h2 { margin: 0; font-size: 1.35rem; font-weight: 700; color: #ffffff; }
    .metric-card {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 12px 16px;
        border-left: 5px solid #2b6cb0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        margin-bottom: 12px;
    }
    .metric-title { font-size: 0.75rem; font-weight: 600; color: #627d98; text-transform: uppercase; }
    .metric-value { font-size: 1.5rem; font-weight: 700; color: #102a43; margin-top: 2px; }
    .stButton>button {
        background-color: #102a43 !important;
        color: #ffffff !important;
        font-weight: 600;
        border-radius: 6px;
        padding: 6px 16px;
    }
</style>
""", unsafe_allow_html=True)
# --- JSON DATABASE ENGINE ---
HISTORY_FILE = "history_log.json"

# ตรวจสอบว่ามีไฟล์ประวัติหรือยัง ถ้ายังไม่มีให้สร้างไฟล์ว่างๆ ขึ้นมา
if not os.path.exists(HISTORY_FILE):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump([], f)

def save_to_history_json(df, inv_no, inv_date):
    # เปิดอ่านไฟล์ประวัติเดิม
    with open(HISTORY_FILE, "r", encoding="utf-8") as f:
        history_data = json.load(f)
    
    # ดึงเวลาปัจจุบัน
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # วนลูปนำรายการที่แบ่งยอดแล้วมาต่อท้าย
    for _, r in df.iterrows():
        supp = str(r.get("Supplier", "")).strip()
        qty_str = str(r.get("Quantity (Allocated)", "")).strip()
        
        if supp and supp != "nan" and qty_str and qty_str != "nan":
            try:
                qty = int(float(qty_str))
                if qty > 0:
                    history_data.append({
                        "Timestamp": now_str,
                        "Invoice_No": inv_no,
                        "Invoice_Date": inv_date,
                        "Part_No": str(r["Part No."]),
                        "Description": str(r["Description of goods"]),
                        "Supplier": supp,
                        "Allocated_Qty": qty
                    })
            except ValueError:
                pass
                
    # เซฟกลับลงไฟล์
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history_data, f, ensure_ascii=False, indent=4)
#3. Backend Engine: Load Master Data
@st.cache_data
def load_backend_master():
    master_path = "Master_Data_Casting.xlsx"
    if os.path.exists(master_path):
        xls = pd.ExcelFile(master_path)
        df_items = pd.read_excel(xls, "Item_Master")
        df_alloc = pd.read_excel(xls, "Supplier_Allocation")
        
        # แปลงเป็นตัวอักษรและตัดช่องว่างเฉยๆ (ไม่ลบ 0 แล้ว)
        if 'Code_CMA' in df_items.columns:
            df_items['Code_CMA'] = df_items['Code_CMA'].astype(str).str.strip()
            
        return df_items, df_alloc
    return None, None

master_items, master_alloc = load_backend_master()
VENDOR_MAP = {
    "PLTHE01": "TMY",
    "PLPAI02": "PLM",
    "PLALP01": "ALPS",
    "PLYAG01": "YGT"
}
#4. Parse Full Items from CMV Invoice
def parse_full_invoice(file):
    xls = pd.ExcelFile(file)
    target_sheet = "IV" if "IV" in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(xls, target_sheet, header=None)
    
    iv_no = "CMV/INV26-106"
    iv_date = "2026-08-25"
    
    for r in range(min(15, len(df_raw))):
        row_vals = [str(x) for x in df_raw.iloc[r].dropna().tolist()]
        for idx, val in enumerate(row_vals):
            if "Invoice No" in val and idx + 1 < len(row_vals):
                iv_no = row_vals[idx + 1]
            if "Invoice date" in val and idx + 1 < len(row_vals):
                iv_date = str(row_vals[idx + 1])[:10]
                
    items = []
    allowed_vendors = ["TMY", "PLM", "ALPS", "YGT"] # 4 ซัพพลายเออร์หลัก
    
    for r in range(20, len(df_raw)):
        row = df_raw.iloc[r]
        val_0 = row[0]
        if pd.notna(val_0) and str(val_0).strip() != "TOTAL":
            try:
                item_no = int(val_0)
                part_no = str(row[1]).strip() if pd.notna(row[1]) else ""
                desc = str(row[2]).strip() if pd.notna(row[2]) else ""
                po = str(row[4]).strip() if pd.notna(row[4]) else ""
                qty = int(row[5]) if pd.notna(row[5]) else 0
                unit_price = row[7] if pd.notna(row[7]) else ""
                amt = row[8] if pd.notna(row[8]) else ""
                
                default_vendor = ""
                default_allocated_qty = ""
                
                # ==========================================
                # ตัวแปรเช็กเงื่อนไข (ถ้าเป็น False คือไม่หาซัพพลายเออร์ แต่ยังโชว์ข้อมูล)
                # ==========================================
                should_find_supplier = True
                
                # 1. ต้องขึ้นต้นด้วยเลข 0 เท่านั้น
                if not part_no.startswith('0'):
                    should_find_supplier = False
                    
                # 2. ต้องไม่ลงท้ายด้วย R หรือ F
                if part_no.upper().endswith(('R', 'F')):
                    should_find_supplier = False
                    
                # 3. ต้องไม่ใช่ชิ้นงาน BED
                if "BED" in desc.upper():
                    should_find_supplier = False
                
                # ==========================================
                
                # ค้นหาซัพพลายเออร์ด้วยระบบ "สะพานเชื่อม (Item_Code)"
                # ==========================================
                if should_find_supplier and (master_items is not None):
                    # Step 1: ค้นหาบรรทัดจาก Code_CMA ในใบ Invoice
                    m = master_items[master_items["Code_CMA"] == part_no]
                    
                    if not m.empty:
                        # ข้ามสะพานไปหา Item_Code หลัก
                        it_code = m.iloc[0]["Item_Code"]
                        
                        # Step 2: เหมาดึงซัพพลายเออร์ทั้งหมดที่มี Item_Code เดียวกัน (รวมงาน FINISH ด้วย)
                        related_rows = master_items[master_items["Item_Code"] == it_code]
                        
                        if "Route_Vendor" in related_rows.columns:
                            raw_vendors = related_rows["Route_Vendor"].dropna().unique()
                            mapped_vendors = [VENDOR_MAP.get(str(v).strip(), str(v).strip()) for v in raw_vendors]
                            
                            # คัดเฉพาะ 4 เจ้าหลัก และลบตัวซ้ำออก
                            valid_vendors = []
                            for v in mapped_vendors:
                                if v in allowed_vendors and v not in valid_vendors:
                                    valid_vendors.append(v)
                            
                            if valid_vendors:
                                default_vendor = " / ".join(valid_vendors)
                                default_allocated_qty = str(qty)

                if not default_vendor.strip():
                    default_allocated_qty = ""
                    
               # *** นำข้อมูล "ทุกบรรทัด" มาต่อท้ายในตาราง ***
                items.append({
                    "No": item_no,
                    "Part No.": part_no,
                    "Description of goods": desc,
                    "P.O No.": po,
                    "Quantity": qty,
                    "Unit Price": unit_price,
                    "Amount (JPY)": amt,
                    "Supplier": default_vendor,
                    "Quantity (Allocated)": str(default_allocated_qty)
                })
            except (ValueError, TypeError):
                continue
                
    df_res = pd.DataFrame(items)
    if not df_res.empty:
        df_res["Quantity (Allocated)"] = df_res["Quantity (Allocated)"].astype(str)
    return df_res, iv_no, iv_date
# 5. Export Exact Formatted CMV Invoice Excel (openpyxl)
def create_annotated_invoice_excel(df_table, iv_no, iv_date):
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    import io
    import pandas as pd

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IV"
    
    # 1. 🔥 บังคับบีบให้อยู่ในหน้ากระดาษเดียวเป๊ะๆ
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1  # บังคับความสูงให้จบใน 1 หน้า
    
    # ขยาย Margin ให้มีพื้นที่เพิ่มขึ้นอีก
    ws.page_margins.left = 0.15
    ws.page_margins.right = 0.15
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.1

    thin_border = Border(
        left=Side(style='thin', color='A6B0BA'),
        right=Side(style='thin', color='A6B0BA'),
        top=Side(style='thin', color='A6B0BA'),
        bottom=Side(style='thin', color='A6B0BA')
    )

    font_bold = Font(name='Arial', size=10, bold=True)
    font_title = Font(name='Arial', size=13, bold=True, color='102A43')
    font_sub = Font(name='Arial', size=9, color='555555')
    font_cell = Font(name='Arial', size=9)
    font_annot = Font(name='Arial', size=9, bold=True, color='002060')

    fill_orig_hdr = PatternFill(start_color='102A43', end_color='102A43', fill_type='solid')
    fill_annot_hdr = PatternFill(start_color='C55A11', end_color='C55A11', fill_type='solid')
    fill_annot_cell = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_tot = PatternFill(start_color='EAEAEA', end_color='EAEAEA', fill_type='solid')
    
    # ย่อความสูงบรรทัดหัวๆ ลงเพื่อประหยัดพื้นที่
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[6].height = 15
    ws.row_dimensions[7].height = 15

    ws['A1'] = "CITIZEN MACHINERY VIETNAM CO., LTD"
    ws['A1'].font = font_title
    ws['A2'] = "Land plot J2, J3, J4, Japan Hai Phong IZ, Hong An ward, Hai Phong, Vietnam"
    ws['A2'].font = font_sub

    ws.merge_cells('A4:I4')
    ws['A4'] = f"INVOICE (WITH STORE ALLOCATION) — {iv_no}"
    ws['A4'].font = Font(name='Arial', size=12, bold=True, color='102A43')
    ws['A4'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A4'].fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    ws['A6'] = f"Invoice No.: {iv_no}"
    ws['A6'].font = font_bold
    ws['A7'] = f"Invoice Date: {iv_date}"
    ws['A7'].font = font_bold
    
    # 2. 🔥 จัดข้อความ Consignee กับ Purpose ให้ชิดซ้าย (left)
    ws['E6'] = "Consignee: CITIZEN MACHINERY ASIA CO., LTD."
    ws['E6'].font = font_bold
    ws['E6'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws['E7'] = "Purpose: Store Receiving & Allocation Record"
    ws['E7'].font = font_bold
    ws['E7'].alignment = Alignment(horizontal='left', vertical='center')

    headers = ['No', 'Part No.', 'Description of goods', 'P.O No.', 'Quantity', 'Unit Price', 'Amount (JPY)', 'Supplier', 'Quantity (Allocated)']
    
    ws.row_dimensions[9].height = 20
    
    for col_idx, h_name in enumerate(headers, start=1):
        c = ws.cell(row=9, column=col_idx, value=h_name)
        c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF')
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.fill = fill_orig_hdr if col_idx <= 7 else fill_annot_hdr
        c.border = thin_border

    current_row = 10
    tot_qty = 0
    tot_alloc = 0
    for _, r in df_table.iterrows():
        # ย่อความสูงของแถวข้อมูลลงนิดนึง เพื่อให้พอยัดเข้า 1 หน้าได้ง่ายขึ้น
        ws.row_dimensions[current_row].height = 14.5
        
        supp_val = str(r.get('Supplier', '')).strip()
        alloc_val = str(r.get('Quantity (Allocated)', '')).strip()

        ws.cell(row=current_row, column=1, value=r.get('No', '')).alignment = Alignment(horizontal='center')
        ws.cell(row=current_row, column=2, value=r.get('Part No.', '')).alignment = Alignment(horizontal='left')
        ws.cell(row=current_row, column=3, value=r.get('Description of goods', '')).alignment = Alignment(horizontal='left')
        ws.cell(row=current_row, column=4, value=r.get('P.O No.', '')).alignment = Alignment(horizontal='left')

        c_qty = ws.cell(row=current_row, column=5, value=r.get('Quantity', ''))
        c_qty.alignment = Alignment(horizontal='right')
        c_qty.number_format = '#,##0'

        c_up = ws.cell(row=current_row, column=6, value=r.get('Unit Price', ''))
        c_up.alignment = Alignment(horizontal='right')

        c_amt = ws.cell(row=current_row, column=7, value=r.get('Amount (JPY)', ''))
        c_amt.alignment = Alignment(horizontal='right')
        c_amt.number_format = '#,##0'

        c_supp = ws.cell(row=current_row, column=8, value=supp_val)
        c_supp.alignment = Alignment(horizontal='center')
        c_supp.font = font_annot
        c_supp.fill = fill_annot_cell

        c_alloc = ws.cell(row=current_row, column=9, value=alloc_val)
        c_alloc.alignment = Alignment(horizontal='center')
        c_alloc.font = font_annot
        c_alloc.fill = fill_annot_cell

        for col_idx in range(1, 10):
            ws.cell(row=current_row, column=col_idx).border = thin_border
            if col_idx < 8:
                ws.cell(row=current_row, column=col_idx).font = font_cell

        if pd.notna(r.get('Quantity')):
            try:
                tot_qty += int(r['Quantity'])
            except:
                pass
        if alloc_val.replace('.0', '').isdigit():
            tot_alloc += int(float(alloc_val))

        current_row += 1

    ws.row_dimensions[current_row].height = 18
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
    ws.cell(row=current_row, column=1, value="TOTAL").alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row=current_row, column=5, value=tot_qty).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(row=current_row, column=5).number_format = '#,##0'
    ws.cell(row=current_row, column=8, value="TOTAL ALLOCATED").alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(row=current_row, column=9, value=tot_alloc if tot_alloc > 0 else "").alignment = Alignment(horizontal='center', vertical='center')

    for col_idx in range(1, 10):
        cell_t = ws.cell(row=current_row, column=col_idx)
        cell_t.font = font_bold
        cell_t.fill = fill_tot
        cell_t.border = thin_border

    col_widths = [8, 18, 28, 20, 12, 12, 16, 20, 20]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
# 6. Export Outward Delivery Note / Gate Pass Excel
def create_gate_pass_excel(df_records, vendor_name, gp_no, doc_date):
    import io
    import math
    import re
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.worksheet.pagebreak import Break
    import pandas as pd
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "OUTWARD DELIVERY NOTE"
    
    ws.views.sheetView[0].showGridLines = False
    
    # --- 🖨️ ตั้งค่าหน้ากระดาษ (บังคับเต็ม A4) ---
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0 
    
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2

    # --- Styles ---
    thin = Side(style='thin', color='000000')
    border_box = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_bottom = Border(bottom=thin)
    
    font_logo = Font(name='Cambria', size=45, bold=True) 
    font_title = Font(name='Calibri', size=20, bold=True)
    font_company = Font(name='Calibri', size=16, bold=True)
    font_address = Font(name='Calibri', size=8.5, bold=True) 
    font_bold_lg = Font(name='Calibri', size=14, bold=True)
    font_bold_md = Font(name='Calibri', size=12, bold=True)
    font_normal_md = Font(name='Calibri', size=12)
    font_red_no = Font(name='Calibri', size=16, color='FF0000', bold=True)

    def get_page_gp_no(base_no, page_index):
        if page_index == 0:
            return base_no
        match = re.search(r'(\d+)$', str(base_no))
        if match:
            num_str = match.group(1)
            new_num = str(int(num_str) + page_index).zfill(len(num_str))
            return str(base_no)[:match.start()] + new_num
        return f"{base_no}-{page_index+1}"

    records_count = len(df_records)
    total_pages = max(1, math.ceil(records_count / 20))
    
    for p in range(total_pages):
        offset = p * 45  
        current_gp_no = get_page_gp_no(gp_no, p)
        
        # --- Header ---
        ws.merge_cells(f'A{offset+1}:B{offset+4}')
        ws[f'A{offset+1}'] = "CITIZEN"
        ws[f'A{offset+1}'].font = font_logo
        ws[f'A{offset+1}'].alignment = Alignment(horizontal='left', vertical='center')
        
        ws.merge_cells(f'C{offset+1}:E{offset+3}')
        ws[f'C{offset+1}'] = "CITIZEN MACHINERY ASIA CO., LTD."
        ws[f'C{offset+1}'].font = font_company 
        ws[f'C{offset+1}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        ws.merge_cells(f'F{offset+1}:G{offset+4}')
        addr = "199 Moo 1 Phaholyotin Road, Sanaptube,\nWang Noi, Ayutthaya 13170\nTel: 66 (0)35 902-604-2 Fax: 66 (0)35 902-644\nTEX ID 0105544056802"
        ws[f'F{offset+1}'] = addr
        ws[f'F{offset+1}'].font = font_address 
        ws[f'F{offset+1}'].alignment = Alignment(horizontal='right', vertical='top', wrap_text=True)
        
        # --- Title ---
        ws.merge_cells(f'A{offset+6}:G{offset+6}')
        ws.row_dimensions[offset+6].height = 25 
        ws[f'A{offset+6}'] = "ใบนำของออกนอกโรงงาน"
        ws[f'A{offset+6}'].font = font_title 
        ws[f'A{offset+6}'].alignment = Alignment(horizontal='center', vertical='center')
        
        # --- Info ---
        ws.row_dimensions[offset+8].height = 20
        ws[f'A{offset+8}'] = "Date (วันที่)"
        ws[f'A{offset+8}'].font = font_bold_md
        ws[f'B{offset+8}'] = ""
        ws[f'B{offset+8}'].border = border_bottom
        
        ws[f'F{offset+8}'] = "No."
        ws[f'G{offset+8}'] = current_gp_no 
        ws[f'F{offset+8}'].font = font_bold_lg; ws[f'F{offset+8}'].alignment = Alignment(horizontal='right')
        ws[f'G{offset+8}'].font = font_red_no; ws[f'G{offset+8}'].alignment = Alignment(horizontal='left')
        
        ws.row_dimensions[offset+10].height = 20
        ws[f'A{offset+10}'] = "Send to(ส่ง)"
        ws[f'A{offset+10}'].font = font_bold_md
        ws[f'B{offset+10}'] = f"{vendor_name}"
        ws[f'B{offset+10}'].font = font_bold_lg
        ws[f'B{offset+10}'].border = border_bottom
        ws[f'B{offset+10}'].alignment = Alignment(horizontal='center', vertical='bottom')
        
        ws.row_dimensions[offset+12].height = 20
        ws[f'A{offset+12}'] = "The purpose (วัตถุประสงค์)"
        ws[f'A{offset+12}'].font = font_bold_md
        
        ws.merge_cells(f'A{offset+13}:G{offset+13}')
        ws.row_dimensions[offset+13].height = 25
        
        # 🔥 วิธีที่ชัวร์ที่สุด ไม่พังแน่นอน
        ws[f'A{offset+13}'] = "     [   ] ส่งซ่อม (Send to repair)          [ ✓ ] จ้างกัด Casting          [   ] อื่น ๆ (Other)________________________"
        ws[f'A{offset+13}'].font = font_bold_md
        ws[f'A{offset+13}'].alignment = Alignment(vertical='center')
        
        # --- Table Headers ---
        headers = [
            ("A", "Item\nลำดับ"),
            ("B", "Description\nรายการ"),
            ("C", "Q'ty\nจำนวน"),
            ("D", "Delivery Date\nวันส่งมอบ"),
            ("E", "PO.No.\nเลขที่ใบสั่งซื้อ"),
            ("F", "Invoice No."),
            ("G", "Remark\nคำอธิบาย")
        ]
        for col, text in headers:
            c = ws[f'{col}{offset+15}']
            c.value = text
            c.font = font_bold_md
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border = border_box
            ws.row_dimensions[offset+15].height = 30
            
        # --- Table Data (20 Rows) ---
        current_r = offset + 16
        for slot in range(20):
            row_num = current_r + slot
            ws.row_dimensions[row_num].height = 29 
            
            data_idx = (p * 20) + slot
            if data_idx < records_count:
                r = df_records.iloc[data_idx]
                ws[f'A{row_num}'] = slot + 1 
                ws[f'B{row_num}'] = str(r.get('Part No.', '')).strip()
                ws[f'C{row_num}'] = r.get('Assigned_Qty', 0)
                ws[f'D{row_num}'] = ""
                ws[f'E{row_num}'] = ""
                ws[f'F{row_num}'] = r.get('Invoice No.', '')
                ws[f'G{row_num}'] = str(r.get('Remark', '')).strip() 
            else:
                for col in ["A", "B", "C", "D", "E", "F", "G"]:
                    ws[f'{col}{row_num}'] = ""
                
            ws[f'A{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'B{row_num}'].alignment = Alignment(horizontal='left', vertical='center')
            ws[f'C{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'D{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'E{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'F{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'G{row_num}'].alignment = Alignment(horizontal='left', vertical='center')
            
            for col_l in ["A", "B", "C", "D", "E", "F", "G"]:
                ws[f'{col_l}{row_num}'].border = border_box
                ws[f'{col_l}{row_num}'].font = font_normal_md

        # --- Footer ---
        footer_start = offset + 37 
        
        ws.row_dimensions[footer_start].height = 30 
        ws[f'A{footer_start}'] = "Return By (ผู้ส่ง) ........................................................"
        ws[f'E{footer_start}'] = "Receive by (ผู้รับ) ........................................................"
        ws[f'A{footer_start}'].font = font_bold_md
        ws[f'E{footer_start}'].font = font_bold_md
        
        ws.row_dimensions[footer_start+2].height = 30 
        ws[f'A{footer_start+2}'] = "Date (วันที่)          ........................................................"
        ws[f'E{footer_start+2}'] = "Date (วันที่)          ........................................................"
        ws[f'A{footer_start+2}'].font = font_bold_md
        ws[f'E{footer_start+2}'].font = font_bold_md
        
        ws.row_dimensions[footer_start+3].height = 20 
        
        box_start = footer_start + 4 
        
        ws.merge_cells(f'A{box_start}:B{box_start}')
        ws.row_dimensions[box_start].height = 20
        ws[f'A{box_start}'] = "Expect return date"
        ws[f'A{box_start}'].font = font_bold_md
        ws[f'A{box_start}'].alignment = Alignment(horizontal='left', vertical='top')
        
        ws.merge_cells(f'A{box_start+1}:B{box_start+1}')
        ws.row_dimensions[box_start+1].height = 15 
        
        ws.merge_cells(f'A{box_start+2}:B{box_start+2}')
        ws.row_dimensions[box_start+2].height = 20
        ws[f'A{box_start+2}'] = "(วันส่งคืน)"
        ws[f'A{box_start+2}'].font = font_bold_md
        ws[f'A{box_start+2}'].alignment = Alignment(horizontal='left', vertical='bottom')
        
        ws.merge_cells(f'C{box_start}:G{box_start+2}')
        ws[f'C{box_start}'] = ""
        
        for row in range(box_start, box_start+3):
            for col in ['A', 'B']:
                ws[f'{col}{row}'].border = Border(left=thin if col=='A' else None, right=thin if col=='B' else None, top=thin if row==box_start else None, bottom=thin if row==box_start+2 else None)
            for col in ['C', 'D', 'E', 'F', 'G']:
                ws[f'{col}{row}'].border = Border(left=thin if col=='C' else None, right=thin if col=='G' else None, top=thin if row==box_start else None, bottom=thin if row==box_start+2 else None)
        
        doc_code_row = box_start + 3
        ws.row_dimensions[doc_code_row].height = 20
        ws[f'A{doc_code_row}'] = "CMA-FR-STS-01-02 (01/09/25)"
        ws[f'A{doc_code_row}'].font = font_address
        
        if p < total_pages - 1:
            page_break = Break(id=offset+45)
            ws.row_breaks.append(page_break)

    # --- Set Column Widths ---
    ws.column_dimensions['A'].width = 13 
    ws.column_dimensions['B'].width = 23 
    ws.column_dimensions['C'].width = 9  
    ws.column_dimensions['D'].width = 16 
    ws.column_dimensions['E'].width = 24 
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 24

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
# 7. Sidebar
st.sidebar.markdown("### 📁 Data Sources")
if master_items is not None:
    st.sidebar.success("✓ โหลด `Master_Data_Casting.xlsx` เรียบร้อย")
else:
    st.sidebar.error("⚠️ ไม่พบไฟล์ Master_Data_Casting.xlsx")

uploaded_file = st.sidebar.file_uploader("📥 อัปโหลด Invoice ขาเข้า (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    if "full_invoice_df" not in st.session_state or st.sidebar.button("🔄 โหลดข้อมูลใหม่"):
        df_full, iv_num, iv_dt = parse_full_invoice(uploaded_file)
        st.session_state.full_invoice_df = df_full
        st.session_state.iv_number = iv_num
        st.session_state.iv_date = iv_dt

# 8. Top Header Banner
st.markdown("""
<div class="top-header">
    <div>
        <h2>CITIZEN MACHINERY ASIA | Integrated Inbound & Outward Management Hub</h2>
        <span>ระบบจัดการใบแจ้งหนี้สโตร์, ใบส่งของออก (Gate Pass) และส่งออก MO สำหรับ MC Frame</span>
    </div>
    <div style="font-weight:600; background:#2b6cb0; padding:4px 12px; border-radius:4px;">Factory CMATH</div>
</div>
""", unsafe_allow_html=True)

# 9. Main Workflow Tabs
if "full_invoice_df" in st.session_state and not st.session_state.full_invoice_df.empty:
    df_show = st.session_state.full_invoice_df.copy()
    df_show["Quantity (Allocated)"] = df_show["Quantity (Allocated)"].astype(str)
    
    total_count = len(df_show)
    cutting_cnt = len(df_show[df_show["Supplier"].str.strip() != ""])
    total_pcs = df_show["Quantity"].sum()
    
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="metric-card"><div class="metric-title">รายการสินค้าทั้งหมด</div><div class="metric-value">{total_count} <span style="font-size:0.85rem; color:#627d98;">Items</span></div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="metric-card" style="border-left-color:#dd6b20;"><div class="metric-title">รายการที่ระบุส่งกัด</div><div class="metric-value">{cutting_cnt} <span style="font-size:0.85rem; color:#627d98;">Items</span></div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="metric-card" style="border-left-color:#319795;"><div class="metric-title">จำนวนชิ้นงานรวม</div><div class="metric-value">{total_pcs} <span style="font-size:0.85rem; color:#627d98;">Pcs</span></div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="metric-card" style="border-left-color:#805ad5;"><div class="metric-title">Invoice No.</div><div class="metric-value" style="font-size:1.1rem; padding-top:4px;">{st.session_state.iv_number}</div></div>', unsafe_allow_html=True)
    tab1, tab2, tab3, tab4 = st.tabs([
    "📋 1. ใบแจ้งหนี้พร้อมจัดสรร", 
    "🚚 2. ใบนำของออก (Gate Pass)", 
    "📄 3. ส่งออก (MC Frame MO)",
    "🔍 4. ค้นหาประวัติ (Tracking History)"
])
    # --- TAB 1 ---
    with tab1:
        st.subheader("ขั้นตอนที่ 1: ตรวจสอบและระบุซัพพลายเออร์")
        edited_df = st.data_editor(
            df_show,
            column_config={
                "No": st.column_config.NumberColumn("No", disabled=True, width="small"),
                "Part No.": st.column_config.TextColumn("Part No.", disabled=True),
                "Description of goods": st.column_config.TextColumn("Description of goods", disabled=True),
                "P.O No.": st.column_config.TextColumn("P.O No.", disabled=True),
                "Quantity": st.column_config.NumberColumn("Quantity", disabled=True, width="small"),
                "Unit Price": st.column_config.NumberColumn("Unit Price", disabled=True),
                "Amount (JPY)": st.column_config.NumberColumn("Amount (JPY)", disabled=True),
                "Supplier": st.column_config.TextColumn("Supplier"),
                "Quantity (Allocated)": st.column_config.TextColumn("Quantity (Allocated)")
            },
            hide_index=True,
            use_container_width=True
        )
        
        if st.button("💾 บันทึกตารางด้านบน"):
            for idx, r in edited_df.iterrows():
                if not str(r["Supplier"]).strip():
                    edited_df.at[idx, "Quantity (Allocated)"] = ""
            st.session_state.full_invoice_df = edited_df
            st.success("บันทึกข้อมูลเรียบร้อย!")
            
        
       # --NEW SPLIT LOGIC
        st.markdown("### 🔀 จัดสรรสัดส่วนซัพพลายเออร์ (Multi-Supplier Split)")
        df_to_split = st.session_state.full_invoice_df.copy()
        
        final_rows = []
        split_parts_for_highlight = [] # เก็บชื่อพาร์ทที่โดนแบ่งยอดไว้ทำไฮไลต์สีฟ้า
        
        # วนลูปอ่านข้อมูลทีละบรรทัดจากบนลงล่าง เพื่อรักษาลำดับเดิมเป๊ะๆ
        for i, (idx, row) in enumerate(df_to_split.iterrows()):
            suppliers_str = str(row['Supplier'])
            
            # เช็กว่าบรรทัดนี้ต้องแบ่งยอดหรือไม่ (มีเครื่องหมาย /)
            if '/' in suppliers_str:
                suppliers = [s.strip() for s in suppliers_str.split('/')]
                try:
                    total_qty = float(row['Quantity (Allocated)'])
                except ValueError:
                    total_qty = 0
                    
                part_no = row.get('Part No.', 'Unknown Code')
                part_desc = row.get('Description of goods', 'Unknown Part')
                
                if len(suppliers) >= 2 and total_qty > 0:
                    split_parts_for_highlight.append(part_no)
                    st.markdown(f"**📦 [{part_no}] {part_desc} (ทั้งหมด {int(total_qty)} ชิ้น)**")
                    
                    cols = st.columns(len(suppliers))
                    allocated_qtys = []
                    
                    for j, sup in enumerate(suppliers):
                        with cols[j]:
                            default_val = int(total_qty) if j == 0 else 0
                            q = st.number_input(f"ส่งให้ {sup} (pcs)", min_value=0, max_value=int(total_qty), value=default_val, step=1, key=f"split_{i}_{j}")
                            allocated_qtys.append((sup, q))
                    
                    current_sum = sum(q for _, q in allocated_qtys)
                    if current_sum != int(total_qty):
                        st.error(f"⚠️ ยอดรวมที่จัดสรร ({current_sum}) ยังไม่เท่ากับยอดเต็ม ({int(total_qty)}) กรุณาปรับตัวเลข")
                    else:
                        is_first_row = True
                        for sup, q in allocated_qtys:
                            if q > 0:
                                row_new = row.copy()
                                row_new['Supplier'] = sup
                                row_new['Quantity (Allocated)'] = str(q)
                                
                                # ถ้าไม่ใช่บรรทัดแรกที่ถูกแบ่ง ให้ซ่อนตัวเลขช่อง No.
                                if not is_first_row:
                                    row_new['No'] = ""
                                
                                try:
                                    row_new['Amount (JPY)'] = float(row_new['Unit Price']) * q
                                except:
                                    pass
                                    
                                final_rows.append(row_new)
                                is_first_row = False # เปลี่ยนสถานะว่าผ่านบรรทัดแรกไปแล้ว
                    st.markdown("---")
                else:
                    final_rows.append(row)
            else:
                # บรรทัดปกติที่ไม่มีการแบ่งยอด ก็ต่อท้ายตามลำดับเดิม
                final_rows.append(row)
                
        # สร้างตารางและรีเซ็ตดัชนีให้ถูกต้อง
        df_final = pd.DataFrame(final_rows)
        df_final.reset_index(drop=True, inplace=True) # <-- บรรทัดนี้สำคัญมาก! ป้องกัน KeyError
        
        # จัดการค่าว่างให้สวยงาม
        df_final = df_final.fillna("")
        
        st.session_state.final_split_df = df_final

        st.markdown("### 📋 ตารางสรุปข้อมูลหลังแบ่งจำนวน (ข้อมูลที่แท้จริงที่จะนำไปออกเอกสาร)")
        
        # ปรับฟังก์ชันไฮไลต์ให้ปลอดภัยขึ้น
        def highlight_split(row):
            part = row.get('Part No.', '')
            if part in split_parts_for_highlight:
                return ['background-color: #e0f2fe'] * len(row)
            return [''] * len(row)

        # ใช้ dataframe ธรรมดาก่อน ถ้าแบบมีสีพัง ให้ลบ .style... ออกเหลือแค่ df_final
        st.dataframe(df_final.style.apply(highlight_split, axis=1), use_container_width=True, hide_index=True)
        
        st.success(f"📌 ข้อมูลพร้อมสำหรับการพิมพ์แล้ว (รวม {len(df_final)} รายการ)")
        st.markdown("---")
        
        excel_annotated = create_annotated_invoice_excel(
            st.session_state.final_split_df, 
            st.session_state.iv_number, 
            st.session_state.iv_date
        )
            
        st.download_button(
            label="📄 ดาวน์โหลดไฟล์ Invoice สำหรับส่งสโตร์ (Excel)",
            data=excel_annotated,
            file_name=f"Invoice_Store_{st.session_state.iv_number.replace('/', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# ปุ่มสำหรับบันทึกประวัติลงฐานข้อมูล JSON
        st.markdown("---")
        if st.button("💾 ยืนยันการจัดสรรและบันทึกประวัติ"):
            save_to_history_json(st.session_state.final_split_df, st.session_state.iv_number, st.session_state.iv_date)
            st.success("✅ บันทึกประวัติการจัดสรรลงระบบเรียบร้อยแล้ว! สามารถตรวจสอบได้ที่ Tab 4")
    # --- TAB 2 ---
    with tab2:
        st.subheader("ขั้นตอนที่ 2: พรีวิวและดาวน์โหลดใบนำของออก (Outward Delivery Note)")
        col_s1, col_s2, col_s3 = st.columns([1.5, 1.5, 1])
        with col_s1:
            target_vendor = st.selectbox("เลือกซัพพลายเออร์ที่ต้องการออกใบนำของ:", ["TMY", "PLM", "ALPS", "YGT"])
        with col_s2:
            gate_pass_no = st.text_input("เลขที่ใบนำของออก", value="CMA2608001")
        with col_s3:
            delivery_date = st.date_input("วันที่ส่งของ", value=datetime.date(2026, 9, 2))
            
        gatepass_items = []
            # ใช้ final_split_df เพื่อสร้างรายการใบนำของออก
        for i, r in st.session_state.final_split_df.iterrows():
                supp_str = str(r.get("Supplier", ""))
                try:
                    qty_val = float(str(r.get("Quantity (Allocated)", 0)).strip())
                except:
                    qty_val = 0

                # กรองเอาเฉพาะรายการของซัพพลายเออร์ที่เลือก และมียอดส่งจริง
                if target_vendor in supp_str and qty_val > 0:
                    part = str(r.get("Part No.", ""))
                    c_code = part
                    
                    # 1. ดึง Casting_Group จาก Master Data ตรงๆ
                    if master_items is not None:
                        m = master_items[(master_items["Code_CMA"].astype(str).str.strip() == part) | 
                                         (master_items["Code_CMA"].astype(str).str.lstrip('0') == part.lstrip('0'))]
                        if not m.empty:
                            if "Route_Vendor" in m.columns:
                                m_vendor = m[m["Route_Vendor"].astype(str).str.contains(target_vendor, na=False)]
                                if not m_vendor.empty:
                                    c_code = str(m_vendor.iloc[0]["Casting_Group"])
                                else:
                                    c_code = str(m.iloc[0]["Casting_Group"])
                            else:
                                c_code = str(m.iloc[0]["Casting_Group"])
                    
                    # 2. ดักเคสพิเศษ TMY 615-1034
                    part_base = part.lstrip('0')
                    if target_vendor == "TMY" and part_base.startswith("615-1034"):
                        st.warning(f"⚠️ พบรายการพิเศษ TMY: {part}")
                        c_code = st.selectbox(
                            f"กรุณาเลือกรหัส Casting สำหรับ {part}:",
                            options=["0615-1061F", "0615-1034-2F"], 
                            key=f"tmy_choice_{i}"
                        )
                    
                    # 3. 🔥 ลอจิกซ่อนคำซ้ำ (สำหรับหน้า Dashboard) และตัดเลข 0
                    def get_base_ui(s):
                        s_clean = str(s).strip().lstrip('0')
                        pts = s_clean.split('-')
                        if len(pts) >= 2:
                            return f"{pts[0]}-{pts[1]}"
                        return s_clean

                    # เทียบรหัสฐาน ถ้าตรงกันให้เว้นว่าง
                    if get_base_ui(part) == get_base_ui(c_code):
                        final_remark = ""
                    else:
                        # ถ้าไม่ตรงกัน ให้ดึงค่ามาโชว์ และปอกเลข 0 ข้างหน้าออก
                        final_remark = c_code.lstrip('0') 
                    
                    # 4. เก็บลงตารางส่งให้ Dashboard และ Excel
                    gatepass_items.append({
                        "Part No.": part,
                        "Assigned_Qty": str(int(qty_val)),
                        "Invoice No.": st.session_state.iv_number,
                        "Remark": final_remark 
                    })
        df_gp = pd.DataFrame(gatepass_items)
        with st.container(border=True):
            # (ด้านล่างจะเป็นโค้ดเดิม with st.container(border=True): ปล่อยไว้เหมือนเดิมครับ)
            st.markdown(f"### **CITIZEN MACHINERY ASIA CO., LTD.**")
            st.caption("199, Mu 1 Phahon Yothin Road, Sanap Tuep Sub-district, Wang Noi, Ayutthaya 13170")
            st.markdown(f"#### **ใบส่งของออก / OUTWARD DELIVERY NOTE (งานจ้างกัด / CUTTING SERVICE) — {target_vendor}**")
            
            if not df_gp.empty:
                st.dataframe(df_gp, hide_index=True, use_container_width=True)
                valid_qtys = [int(float(x)) for x in df_gp["Assigned_Qty"] if str(x).replace('.0', '').isdigit()]
                tot_gp_qty = sum(valid_qtys)
                st.markdown(f"<h4 style='text-align:right;'>รวมจำนวนทั้งสิ้น: {tot_gp_qty} ชิ้น</h4>", unsafe_allow_html=True)
                
                gp_excel = create_gate_pass_excel(df_gp, target_vendor, gate_pass_no, delivery_date)
                st.download_button(
                    label=f"📄 ดาวน์โหลดใบนำของออก ({target_vendor}) เป็นไฟล์ Excel",
                    data=gp_excel,
                    file_name=f"GatePass_{target_vendor}_{gate_pass_no}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning(f"ยังไม่มีรายการที่ระบุส่งไปยัง {target_vendor}")
# --- TAB 3 ---
    with tab3:
        st.subheader("ขั้นตอนที่ 3: ส่งออกชุดข้อมูล MO สำหรับอัปโหลดเข้า MC Frame (แยกไฟล์ตามซัพพลายเออร์)")
        st.caption("(รูปแบบข้อมูลอิงตามไฟล์แม่แบบ PUS (กรอกเฉพาะคอลัมน์ที่จำเป็น))")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            default_start_date = pd.Timestamp.now().strftime("%d/%m/%Y 0:00")
            mfg_start_date = st.text_input("Actual manufacturing start date", value=default_start_date)
        with col_d2:
            mfg_finish_date = st.text_input("Sched. manufacturing finish date", value="", help="เว้นว่างไว้ตามเงื่อนไขใหม่")

        mc_rows = []
        for _, r in st.session_state.final_split_df.iterrows():
            supp = str(r["Supplier"]).strip()
            if supp and str(r["Quantity (Allocated)"]).strip():
                part = str(r["Part No."]).strip()
                
                # --- 🔥 LOGIC การดึง Casting Code ที่ถูกต้อง 100% ---
                c_code = part 
                
                # 1. เช็คก่อนว่ามีรหัสพิเศษที่เลือกจากดรอปดาวน์ (เก็บใน Remark) ไหม?
                remark_val = str(r.get("Remark", "")).strip() 
                if remark_val != "" and remark_val.lower() != "nan":
                    # ถ้ามีการเลือกจากดรอปดาวน์ ให้ใช้ตัวนี้เลย!
                    c_code = remark_val 
                else:
                    # 2. ถ้าไม่มี ค่อยไปหา Casting Group จาก Master Data
                    if master_items is not None:
                        m = master_items[(master_items["Code_CMA"] == part) | (master_items["Item_Code"] == part)]
                        if not m.empty:
                            it_code = m.iloc[0]["Item_Code"]
                            alloc = master_alloc[master_alloc["Item_Code"] == it_code]
                            
                            vendor_alloc = alloc[alloc["Supplier_Code"].map(VENDOR_MAP).fillna(alloc["Supplier_Code"]) == supp]
                            if not vendor_alloc.empty:
                                c_code = vendor_alloc.iloc[0]["Casting_Group"] # ใช้ Casting_Group แล้ว!
                            elif not alloc.empty:
                                c_code = alloc.iloc[0]["Casting_Group"] # ใช้ Casting_Group แล้ว!
                
                # 3. ตัด 0 ข้างหน้าทิ้ง
                c_code = str(c_code).strip()
                if c_code.startswith("0"):
                    c_code = c_code[1:] # ตัด 0 ตัวแรกทิ้ง
                # -------------------------------------------------------------

                c_code_upper = str(c_code).upper()
                if 'F' in c_code_upper:
                    storage_loc = "MP03"
                elif 'R' in c_code_upper:
                    storage_loc = "MP02"
                else:
                    storage_loc = "MP02"

                try:
                    qty = int(float(r["Quantity (Allocated)"]))
                except ValueError:
                    qty = 0

                if qty > 0:
                    mc_rows.append({
                        "_Supplier": supp,
                        "Item CD": c_code, # ดึงโค้ดตามเงื่อนไขเป๊ะๆ
                        "Manufacturing loc. CD": "OS01",
                        "BOM pattern": 1,
                        "Lot No.": "*",
                        "SERIAL No.": "",
                        "MFG No.": "",
                        "Sched. manufacturing finish date": "",
                        "Actual manufacturing start date": mfg_start_date,
                        "Actual manufacturing finish date": "",
                        "Posting date": "",
                        "Sched. manufacturing qty.": qty,
                        "Actual manufacturing qty.": qty,
                        "Completed": "",
                        "Storage loc. CD": storage_loc,
                        "Operation dept.": "PUS",
                        "Responsible PIC": "",
                        "Manufacturing note": "",
                        "Line CD": "",
                        "Defective reason CD": "",
                        "Defective qty.": "",
                        "Defective item yard": "",
                        "Defective item rack No.": "",
                        "Mold branch No.": "",
                        "Number of cavities": "",
                        "Shot count": "",
                        "Shot wt.": "",
                        "Spec. CD": ""
                    })

        df_all_mc = pd.DataFrame(mc_rows)

        if not df_all_mc.empty:
            mo_suppliers = df_all_mc["_Supplier"].unique()
            st.markdown("### 📦 เลือกดาวน์โหลดไฟล์ MO ตามซัพพลายเออร์")

            dl_cols = st.columns(len(mo_suppliers))
            for i, supp_name in enumerate(mo_suppliers):
                df_supp = df_all_mc[df_all_mc["_Supplier"] == supp_name].drop(columns=["_Supplier"])

                with dl_cols[i]:
                    st.info(f"**{supp_name}** (รวม {len(df_supp)} รายการ)")
                    st.download_button(
                        label=f"🚀 ดาวน์โหลด MO ของ {supp_name}",
                        data=df_supp.to_csv(index=False).encode('utf-8-sig'),
                        file_name=f"MO_Upload_{supp_name}_{st.session_state.iv_number.replace('/', '_')}.csv",
                        mime="text/csv",
                        key=f"dl_mo_{supp_name}"
                    )

            st.markdown("---")
            st.markdown("### 👀 พรีวิวข้อมูลก่อนดาวน์โหลด")
            preview_vendor = st.selectbox("เลือกดูตัวอย่างข้อมูลของซัพพลายเออร์:", mo_suppliers)
            df_preview = df_all_mc[df_all_mc["_Supplier"] == preview_vendor].drop(columns=["_Supplier"])
            st.dataframe(df_preview, hide_index=True, use_container_width=True)
        else:
            st.warning("ยังไม่มีข้อมูลสำหรับออกไฟล์ MO")
# --- TAB 4 ---
    with tab4:
        st.subheader("🔍 ค้นหาประวัติการทำงานย้อนหลัง")
        st.caption("ข้อมูลทั้งหมดถูกบันทึกไว้ในรูปแบบไฟล์ JSON")
        
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                hist_data = json.load(f)
                
            if hist_data:
                df_hist = pd.DataFrame(hist_data)
                
                # สร้างช่องค้นหา
                col_s1, col_s2, col_s3 = st.columns(3)
                with col_s1:
                    search_part = st.text_input("🔍 ค้นหาด้วย Part No.")
                with col_s2:
                    search_inv = st.text_input("🧾 ค้นหาด้วย เลข Invoice")
                with col_s3:
                    search_supp = st.selectbox("🏢 กรองตามซัพพลายเออร์", ["ทั้งหมด"] + list(df_hist["Supplier"].unique()))
                
                # ระบบกรองข้อมูล
                df_show = df_hist.copy()
                if search_part:
                    df_show = df_show[df_show["Part_No"].str.contains(search_part, case=False, na=False)]
                if search_inv:
                    df_show = df_show[df_show["Invoice_No"].str.contains(search_inv, case=False, na=False)]
                if search_supp != "ทั้งหมด":
                    df_show = df_show[df_show["Supplier"] == search_supp]
                    
                # แสดงผลตาราง
                st.dataframe(df_show, hide_index=True, use_container_width=True)
                st.info(f"📊 พบข้อมูลทั้งหมด {len(df_show)} รายการ")
            else:
                st.warning("📭 ยังไม่มีประวัติการจัดสรรข้อมูล")
        else:
            st.error("⚠️ ไม่พบไฟล์ฐานข้อมูล (history_log.json)")
else:
    st.info("👈 กรุณาอัปโหลดไฟล์ Invoice ขาเข้า (.xlsx) ที่แถบด้านซ้าย เพื่อเริ่มใช้งาน")
    