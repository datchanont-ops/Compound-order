import streamlit as st
import pandas as pd
from datetime import datetime
import os
from io import BytesIO
import re

st.set_page_config(page_title="ระบบสั่งยาง (Subcontractor)", layout="wide")
st.title("📝 ระบบสั่งยาง (Order Form)")

# -------------------------------------------------------------
# 1. ดึงข้อมูล Master Data จากไฟล์ Excel
# -------------------------------------------------------------
@st.cache_data
def load_data():
    file_path = "data base batch size.xlsx"
    if os.path.exists(file_path):
        df = pd.read_excel(file_path, sheet_name='batch size')
        df['SUB'] = df['SUB'].ffill() 
        df = df.dropna(subset=['Compound SUB', 'SUB']) 
        return df
    else:
        st.error(f"ไม่พบไฟล์ {file_path} ในระบบ")
        return pd.DataFrame()

master_data = load_data()
sub_list_fixed = ["CRM", "KSA", "MCR", "PKPS", "PS", "SCC", "SCR", "TEP", "TPS", "TPT", "VK", "WCL01"]

columns_schema = ["เลือกเพื่อลบ", "Plan type", "Recipe name", "Quantity", "ประเภทยาง", "Unit", "Delivery Date", "Remark", "Total Kg"]

if 'order_df' not in st.session_state:
    st.session_state.order_df = pd.DataFrame(columns=columns_schema)

if 'last_delivery_date' not in st.session_state:
    st.session_state.last_delivery_date = datetime.today()

if 'order_confirmed' not in st.session_state:
    st.session_state.order_confirmed = False

def calculate_total_kg(df, master):
    for i, row in df.iterrows():
        try:
            qty = float(row['Quantity'])
        except:
            qty = 0.0
        
        if row['Unit'] == 'Batch':
            recipe = str(row['Recipe name'])
            if not master.empty and recipe in master['Compound SUB'].values:
                batch_size = master[master['Compound SUB'] == recipe]['Batch size (kg)'].iloc[0]
                df.at[i, 'Total Kg'] = round(qty * float(batch_size), 3)
            else:
                df.at[i, 'Total Kg'] = 0.0
        else: 
            df.at[i, 'Total Kg'] = round(qty, 3)
    return df

# -------------------------------------------------------------
# 3. ส่วนหัว: เลือกบริษัท SUB
# -------------------------------------------------------------
st.subheader("ข้อมูลลูกค้า")
col1, col2 = st.columns(2)
with col1:
    selected_sub = st.selectbox("ชื่อลูกค้า (Customer)", ["กรุณาเลือก..."] + sub_list_fixed)
with col2:
    order_date = st.date_input("วันที่ทำรายการ (วันที่วางแผนสั่งยาง)", datetime.today())

st.divider()

if selected_sub != "กรุณาเลือก...":
    
    # -------------------------------------------------------------
    # 4. เมนูจัดการข้อมูล (เพิ่มมือ / อัปโหลด)
    # -------------------------------------------------------------
    tab1, tab2 = st.tabs(["✍️ เพิ่มรายการด้วยตัวเอง", "📥 อัปโหลดไฟล์มาตรฐาน (ตัวอย่าง ใบสั่งยาง.xlsx)"])
    
    # --- Tab 1: เพิ่มรายการด้วยตัวเอง ---
    with tab1:
        if not master_data.empty:
            filtered_compounds = master_data[master_data['SUB'] == selected_sub]['Compound SUB'].tolist()
            filtered_compounds = sorted(list(set(filtered_compounds)))
        else:
            filtered_compounds = []
            
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        c_plan, c5, c6, c7 = st.columns([1, 1.5, 2.5, 1]) 
        
        with c1:
            recipe = st.selectbox("Recipe name (ชื่อยาง)", ["เลือกหรือพิมพ์ค้นหา..."] + filtered_compounds)
        
        default_rubber_idx = 0 
        if recipe != "เลือกหรือพิมพ์ค้นหา...":
            match = re.search(r'\d', recipe) 
            if match and match.group() == '5':
                default_rubber_idx = 1 

        with c2:
            rubber_type = st.selectbox("ประเภทยาง", ["ยางแผ่น", "ยาง Ribbon", "ยางเป่า", "ยาง ext", "อื่นๆ"], index=default_rubber_idx)
        with c3:
            qty = st.number_input("Quantity (จำนวน)", min_value=1, step=1)
        with c4:
            unit = st.selectbox("Unit (หน่วย)", ["Batch", "Kg"])
            
        with c_plan:
            plan_type = st.selectbox("Plan type (แผน)", ["แผนปกติ", "ปรับเพิ่มยาง", "ลดยาง"])
        with c5:
            delivery_date = st.date_input("Delivery Date", value=st.session_state.last_delivery_date)
        with c6:
            remark = st.text_input("Remark (หมายเหตุเพิ่มเติม)")
                
        with c7:
            st.write("") 
            if st.button("➕ เพิ่มรายการ", use_container_width=True):
                if recipe == "เลือกหรือพิมพ์ค้นหา...":
                    st.error("กรุณาเลือกหรือพิมพ์ชื่อ Recipe name!")
                else:
                    new_row = {
                        "เลือกเพื่อลบ": False, "Plan type": plan_type, "Recipe name": recipe, 
                        "Quantity": qty, "ประเภทยาง": rubber_type, "Unit": unit, 
                        "Delivery Date": delivery_date.strftime("%d.%m.%Y"), "Remark": remark, "Total Kg": 0.0
                    }
                    df_new = pd.DataFrame([new_row])
                    df_new = calculate_total_kg(df_new, master_data)
                    st.session_state.order_df = pd.concat([st.session_state.order_df, df_new], ignore_index=True)
                    st.session_state.last_delivery_date = delivery_date
                    st.session_state.order_confirmed = False
                    st.rerun()

    # --- Tab 2: อัปโหลด Template มาตรฐาน ---
    with tab2:
        col_down, col_up = st.columns([1, 1])
        with col_down:
            st.write("**1. ดาวน์โหลดไฟล์มาตรฐานเพื่อนำไปกรอก**")
            template_path = "ตัวอย่าง ใบสั่งยาง.xlsx"
            if os.path.exists(template_path):
                with open(template_path, "rb") as file:
                    btn = st.download_button(
                        label="⬇️ ดาวน์โหลด 'ตัวอย่าง ใบสั่งยาง.xlsx'",
                        data=file,
                        file_name="แบบฟอร์มสั่งยาง_มาตรฐาน.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("ไม่พบไฟล์ 'ตัวอย่าง ใบสั่งยาง.xlsx' ในระบบ")
            
        with col_up:
            st.write("**2. อัปโหลดไฟล์มาตรฐานที่กรอกเสร็จแล้ว**")
            uploaded_file = st.file_uploader("ลากไฟล์มาวาง หรือ กดเลือกไฟล์", type=["xlsx", "xls"])
            if uploaded_file and st.button("📥 นำเข้าข้อมูลเข้าสู่ตาราง"):
                try:
                    # อ่านไฟล์มาตรฐาน ข้าม 3 บรรทัดแรก
                    df_upload = pd.read_excel(uploaded_file, sheet_name='form', skiprows=3)
                    # ตัดแถวแรกทิ้ง (เพราะเป็นหัวข้อภาษาไทย) และลบแถวที่ว่างออก
                    df_upload = df_upload.iloc[1:].dropna(subset=['Recipe name'])
                    
                    # จัดเรียงข้อมูลให้ตรงกับโครงสร้างของระบบ
                    formatted_data = []
                    for _, row in df_upload.iterrows():
                        formatted_data.append({
                            "เลือกเพื่อลบ": False,
                            "Plan type": "แผนปกติ", # ตั้งเป็นค่าเริ่มต้น
                            "Recipe name": row['Recipe name'],
                            "Quantity": row['Quantity'],
                            "ประเภทยาง": row['Unnamed: 3'], # คอลัมน์ประเภทยางในไฟล์เดิม
                            "Unit": row['Unit'],
                            "Delivery Date": pd.to_datetime(row['Delivery Date']).strftime("%d.%m.%Y") if pd.notna(row['Delivery Date']) else "",
                            "Remark": row['Unnamed: 7'] if pd.notna(row['Unnamed: 7']) else "",
                            "Total Kg": 0.0
                        })
                        
                    df_formatted = pd.DataFrame(formatted_data)
                    df_formatted = calculate_total_kg(df_formatted, master_data)
                    
                    st.session_state.order_df = pd.concat([st.session_state.order_df, df_formatted], ignore_index=True)
                    st.session_state.order_confirmed = False
                    st.success(f"นำเข้าข้อมูลจากไฟล์สำเร็จ จำนวน {len(df_formatted)} รายการ!")
                    st.rerun()
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ กรุณาใช้ไฟล์รูปแบบมาตรฐานเท่านั้น (Error: {e})")

    st.divider()

    # -------------------------------------------------------------
    # 5. สรุปรายการในใบสั่งยาง
    # -------------------------------------------------------------
    st.subheader("📋 สรุปรายการในใบสั่งยาง")
    
    if not st.session_state.order_df.empty:
        # [แก้ไขที่ 1] ปรับกลับไปใช้ตารางที่ไม่สามารถแก้ไขข้อความได้โดยตรง แต่ยังติ๊กลบได้
        edited_df = st.data_editor(
            st.session_state.order_df,
            use_container_width=True,
            num_rows="fixed", # ปิดการเพิ่มบรรทัดว่าง
            hide_index=True,
            disabled=["Plan type", "Recipe name", "Quantity", "ประเภทยาง", "Unit", "Delivery Date", "Remark", "Total Kg"], # ล็อกไม่ให้พิมพ์แก้
            column_config={
                "เลือกเพื่อลบ": st.column_config.CheckboxColumn("เลือก", default=False)
            }
        )
        
        st.session_state.order_df = edited_df

        col_btn1, col_btn2 = st.columns([2, 8])
        with col_btn1:
            if st.button("❌ ลบรายการที่เลือก"):
                remaining_rows = st.session_state.order_df[~st.session_state.order_df['เลือกเพื่อลบ']]
                st.session_state.order_df = remaining_rows.reset_index(drop=True)
                st.rerun()
        with col_btn2:
            if st.button("🗑️ ล้างรายการทั้งหมด"):
                st.session_state.order_df = pd.DataFrame(columns=columns_schema)
                st.rerun()
    else:
        st.info("ยังไม่มีรายการสั่งยางในใบสั่งนี้")

    st.divider()
    
    if st.button("🚀 ยืนยันการส่งใบสั่งยางไปยังโรงงาน", type="primary"):
        if not st.session_state.order_df.empty:
            st.session_state.order_confirmed = True
            st.success("✅ ระบบบันทึกข้อมูลเรียบร้อยแล้ว คุณสามารถดาวน์โหลดไฟล์ Excel ได้ด้านล่าง")
        else:
            st.error("ไม่มีรายการสั่งยางในตาราง!")

    # -------------------------------------------------------------
    # 6. ฟังก์ชัน Export Excel
    # -------------------------------------------------------------
    if st.session_state.order_confirmed and not st.session_state.order_df.empty:
        df_export = st.session_state.order_df.copy()
        df_export = df_export.drop(columns=['เลือกเพื่อลบ'])
        df_export.insert(0, 'Customer', selected_sub) 
        df_export.insert(0, 'No.', range(1, len(df_export) + 1))
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Order_Data')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 ดาวน์โหลดหลักฐานใบสั่งยาง (Excel)",
            data=processed_data,
            file_name=f"Order_{selected_sub}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )