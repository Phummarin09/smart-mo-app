import streamlit as st
import pandas as pd
import numpy as np

# ตั้งค่าหน้าจอแบบกว้าง
st.set_page_config(page_title="Smart Procurement Dashboard", layout="wide")

# --- Custom CSS สไตล์ละมุน โค้งมน ไล่เฉดสีพาสเทล (Soft Modern UI) ---
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .soft-card {
        background: #ffffff; padding: 24px; border-radius: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.05); margin-bottom: 20px;
        border: 1px solid rgba(255, 255, 255, 0.8);
    }
    .dashboard-header {
        background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%); color: white; padding: 24px 30px; border-radius: 20px; margin-bottom: 25px;
        box-shadow: 0 10px 20px -5px rgba(79, 70, 229, 0.3);
    }
    .metric-gradient-1 { background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%); color: white; padding: 20px; border-radius: 16px; box-shadow: 0 8px 16px -4px rgba(99, 102, 241, 0.3); }
    .metric-gradient-2 { background: linear-gradient(135deg, #f43f5e 0%, #fb7185 100%); color: white; padding: 20px; border-radius: 16px; box-shadow: 0 8px 16px -4px rgba(244, 63, 94, 0.3); }
    .metric-gradient-3 { background: linear-gradient(135deg, #3b82f6 0%, #60a5fa 100%); color: white; padding: 20px; border-radius: 16px; box-shadow: 0 8px 16px -4px rgba(59, 130, 246, 0.3); }
    .metric-gradient-4 { background: linear-gradient(135deg, #f59e0b 0%, #fbbf24 100%); color: white; padding: 20px; border-radius: 16px; box-shadow: 0 8px 16px -4px rgba(245, 158, 11, 0.3); }
    .metric-label { font-size: 0.85rem; font-weight: 600; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-number { font-size: 2rem; font-weight: 800; margin-top: 8px; }
    .stButton>button { border-radius: 12px; font-weight: 600; padding: 0.5rem 1rem; border: none; }
    </style>
""", unsafe_allow_html=True)

# รายชื่อ 12 เดือนสำหรับใช้ในระบบ
months_list = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

try:
    # 1. โหลด Master Data
    df_master = pd.read_excel("Master_Item_Data.xlsx")
    df_master.columns = df_master.columns.str.strip()
    df_master.rename(columns={'New Code': 'Part_No', 'Name': 'Part_Name', 'Mode': 'Model', 'SUPPLIER': 'Supplier', 'SAFTY': 'Safety_Stock'}, inplace=True)
    
    total_items = len(df_master)

    # --- TOP HEADER ---
    st.markdown(f"""
        <div class="dashboard-header">
            <h1 style="margin:0; font-size:1.75rem; font-weight:800; color:white;">✨ Smart Procurement: Advance Planning System</h1>
            <p style="margin:5px 0 0 0; font-size:0.9rem; opacity:0.9;">ระบบคำนวณและวางแผนการสั่งซื้อล่วงหน้า (Annual Plan) เพื่อรักษาระดับสต็อกอัตโนมัติ • ทั้งหมด {total_items} รายการ</p>
        </div>
    """, unsafe_allow_html=True)

    # --- CONTROL PANEL ---
    st.markdown('<div class="soft-card">', unsafe_allow_html=True)
    c_dl, c_ul = st.columns(2)
    
    with c_dl:
        st.markdown("### 📥 1. โหลดฟอร์มแผนจัดซื้อ (Master Template)")
        st.markdown("<p style='color:#64748b; font-size:0.85rem;'>ดาวน์โหลดไฟล์มาตรฐานสำหรับกรอกข้อมูลสต็อกและแผนล่วงหน้า 12 เดือน</p>", unsafe_allow_html=True)
        
        template_df = df_master[['Part_No', 'Model', 'Supplier', 'Part_Name']].copy()
        template_df['Safety_Stock'] = df_master['Safety_Stock'].fillna(30) if 'Safety_Stock' in df_master.columns else 30
        
        # สร้างคอลัมน์ 12 เดือน
        for m in months_list:
            template_df[f'Stock_{m}'] = 0
            template_df[f'PO_{m}'] = 0
            template_df[f'Plan_{m}'] = 0

        csv_template = template_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 ดาวน์โหลดแบบฟอร์ม 12 เดือน (CSV)",
            data=csv_template,
            file_name="Smart_Procurement_Template.csv",
            mime="text/csv",
            type="secondary"
        )

    with c_ul:
        st.markdown("### 📤 2. อัปโหลดเพื่อประมวลผล (Upload Data)")
        st.markdown("<p style='color:#64748b; font-size:0.85rem;'>นำเข้าไฟล์ที่อัปเดตข้อมูลแล้ว เพื่อให้ระบบคำนวณยอดสั่งซื้อ (Delivery)</p>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("เลือกไฟล์ Excel หรือ CSV", type=['csv', 'xlsx'], label_visibility="collapsed")
    
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        df_upload = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        df_upload.columns = df_upload.columns.str.strip()
        
        # เช็กว่าไฟล์ถูกต้องไหม (ต้องมีคอลัมน์ของเดือน Jan เป็นอย่างน้อย)
        if 'Plan_Jan' not in df_upload.columns:
            st.error("⚠️ รูปแบบไฟล์ไม่ถูกต้อง! กรุณาดาวน์โหลด Template แบบ 12 เดือนใหม่จากด้านบนครับ")
            st.stop()

        # แปลงข้อมูลเป็นตัวเลขให้ครบทุกเดือน
        all_numeric_cols = ['Safety_Stock']
        for m in months_list:
            all_numeric_cols.extend([f'Stock_{m}', f'PO_{m}', f'Plan_{m}'])
            
        for col in all_numeric_cols:
            if col in df_upload.columns:
                df_upload[col] = pd.to_numeric(df_upload[col], errors='coerce').fillna(0)

        # แผงเลือกเดือน
        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        
        selected_month = st.selectbox(
            "📅 เลือกรอบเดือนที่ต้องการคำนวณยอดสั่งซื้อ (ระบบจะดึงข้อมูลอัตโนมัติจากเดือนที่เลือก):",
            months_list,
            index=8 # ค่าเริ่มต้นที่เดือน 9 (Sep)
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # ⚙️ ดึงข้อมูลตามเดือนที่เลือกมาจากชื่อคอลัมน์แบบไดนามิก
        df_upload['Target_Plan'] = df_upload[f'Plan_{selected_month}']
        df_upload['Current_Stock_Active'] = df_upload[f'Stock_{selected_month}']
        df_upload['PO_Balance_Active'] = df_upload[f'PO_{selected_month}']

        # สูตรคำนวณ Delivery: (แผน + Safety Stock) - สต็อกที่มี
        df_upload['Required_Order'] = (df_upload['Target_Plan'] + df_upload['Safety_Stock']) - df_upload['Current_Stock_Active']
        df_upload['Required_Order'] = df_upload['Required_Order'].apply(lambda x: max(x, 0))
        
        def check_po(row):
            if row['PO_Balance_Active'] == 0: return "⚠️ ยังไม่มี PO"
            elif row['Required_Order'] > row['PO_Balance_Active']: return "⚠️ PO ไม่พอ"
            return "✅ PO เพียงพอ"
        df_upload['สถานะ PO'] = df_upload.apply(check_po, axis=1)
        
        def assign_status(row):
            if row['Required_Order'] > 0: return "⚡ ORDER NOW"
            else: return "⏳ BALANCED"
        df_upload['สถานะการตัดสินใจ'] = df_upload.apply(assign_status, axis=1)
        
        # สรุป Metrics
        total_processed = len(df_upload)
        order_now_count = len(df_upload[df_upload['สถานะการตัดสินใจ'] == "⚡ ORDER NOW"])
        balanced_count = len(df_upload[df_upload['สถานะการตัดสินใจ'] == "⏳ BALANCED"])
        po_warning_count = len(df_upload[df_upload['สถานะ PO'].str.contains("⚠️")])

        # --- METRIC CARDS ---
        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            st.markdown(f"""<div class="metric-gradient-1"><div class="metric-label">Total Items</div><div class="metric-number">{total_processed}</div></div>""", unsafe_allow_html=True)
        with mc2:
            st.markdown(f"""<div class="metric-gradient-2"><div class="metric-label">⚡ Order Now</div><div class="metric-number">{order_now_count}</div></div>""", unsafe_allow_html=True)
        with mc3:
            st.markdown(f"""<div class="metric-gradient-3"><div class="metric-label">⏳ Balanced</div><div class="metric-number">{balanced_count}</div></div>""", unsafe_allow_html=True)
        with mc4:
            st.markdown(f"""<div class="metric-gradient-4"><div class="metric-label">⚠️ PO Warning</div><div class="metric-number">{po_warning_count}</div></div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # แปลงเป็นจำนวนเต็ม
        cols_to_int = ['Safety_Stock', 'Current_Stock_Active', 'PO_Balance_Active', 'Target_Plan', 'Required_Order']
        df_upload[cols_to_int] = df_upload[cols_to_int].astype(int)
        
        display_cols = ['Part_No', 'Model', 'Supplier', 'Part_Name', 'Safety_Stock', 'Current_Stock_Active', 'PO_Balance_Active', 'Target_Plan', 'Required_Order', 'สถานะ PO', 'สถานะการตัดสินใจ']
        
        # --- FILTER & TABLE ---
        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        f_col1, f_col2 = st.columns([2, 4])
        with f_col1:
            selected_supplier = st.selectbox("กรองตาม Supplier:", ["ทั้งหมด"] + list(df_upload['Supplier'].dropna().unique()))
        with f_col2:
            search_query = st.text_input("🔍 ค้นหารหัสพาร์ท หรือ ชื่อชิ้นงาน:", placeholder="พิมพ์ค้นหาด่วน...")
        
        filtered_df = df_upload.copy()
        if selected_supplier != "ทั้งหมด":
            filtered_df = filtered_df[filtered_df['Supplier'] == selected_supplier]
        if search_query:
            filtered_df = filtered_df[
                filtered_df['Part_No'].astype(str).str.contains(search_query, case=False, na=False) |
                filtered_df['Part_Name'].astype(str).str.contains(search_query, case=False, na=False)
            ]

        st.markdown(f"#### 📋 ตารางผลลัพธ์รอบเดือน: **{selected_month}** (แสดง {len(filtered_df)} รายการ)")
        
        def style_table(val):
            if "ORDER" in str(val): return "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
            elif "BALANCED" in str(val): return "background-color: #eff6ff; color: #1e40af; font-weight: 700;"
            elif "⚠️" in str(val): return "color: #ef4444; font-weight: 700;"
            elif "✅" in str(val): return "color: #10b981; font-weight: 700;"
            return ""

        st.dataframe(
            filtered_df[display_cols].rename(columns={
                'Current_Stock_Active': f'สต็อก (Stock {selected_month})', 
                'PO_Balance_Active': f'ยอด PO (PO {selected_month})', 
                'Target_Plan': f'แผนผลิต (Plan {selected_month})', 
                'Required_Order': 'ยอดที่ต้องสั่ง (Delivery)'
            }).style.map(style_table),
            use_container_width=True, hide_index=True
        )
        
        # ปุ่ม Export
        st.markdown("<br>", unsafe_allow_html=True)
        csv_result = filtered_df[display_cols].to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label=f"💾 Export สรุปแผนจัดซื้อเดือน {selected_month} (CSV)",
            data=csv_result,
            file_name=f"Smart_Procurement_Report_{selected_month}.csv",
            mime="text/csv",
            type="primary"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
    else:
        st.markdown('<div class="soft-card" style="text-align: center; padding: 40px;">', unsafe_allow_html=True)
        st.markdown("💡 **คำแนะนำเบื้องต้น:** เริ่มต้นโดยการดาวน์โหลด Template ไปอัปเดตข้อมูลสต็อกและแผนการใช้ จากนั้นอัปโหลดกลับเข้าระบบเพื่อวิเคราะห์ยอดสั่งซื้อ", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

except FileNotFoundError:
    st.error("⚠️ ไม่พบไฟล์ Master_Item_Data.xlsx กรุณาวางไฟล์ไว้ในโฟลเดอร์เดียวกับโปรแกรมครับ")
except Exception as e:
    st.error(f"⚠️ เกิดข้อผิดพลาด: {e}")