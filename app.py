import streamlit as st
import pandas as pd
from datetime import datetime
import os
from io import BytesIO
import re # นำเข้ามาเพื่อใช้เช็คตัวเลขในชื่อสูตรยาง

# ตั้งค่าหน้าจอ
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

# -------------------------------------------------------------
# 2. ระบบตะกร้าเก็บรายการ (Session State)
# -------------------------------------------------------------
if 'order_list' not in st.session_state:
    st.session_state.order_list = []

if 'last_delivery_date' not in st.session_state:
    st.session_state.last_delivery_date = datetime.today()

if 'order_confirmed' not in st.session_state:
    st.session_state.order_confirmed = False

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

# -------------------------------------------------------------
# 4. ฟอร์มเพิ่มรายการสั่งยาง (Real-time)
# -------------------------------------------------------------
if selected_sub != "กรุณาเลือก...":
    st.subheader("เพิ่มรายการสั่งยาง")
    
    if not master_data.empty:
        filtered_compounds = master_data[master_data['SUB'] == selected_sub]['Compound SUB'].tolist()
        filtered_compounds = sorted(list(set(filtered_compounds)))
    else:
        filtered_compounds = []
    
    # แบ่งคอลัมน์
    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
    c_plan, c5, c6, c7 = st.columns([1, 1.5, 2.5, 1]) 
    
    with c1:
        recipe = st.selectbox("Recipe name (ชื่อยาง)", ["เลือกหรือพิมพ์ค้นหา..."] + filtered_compounds)
    
    # [แก้ไขที่ 1] Logic เช็คชื่อยาง: ถ้าตัวเลขแรกของชื่อยางคือ '5' ให้สลับไปเป็น "ยาง Ribbon"
    default_rubber_idx = 0 # ค่าตั้งต้น: 0 = "ยางแผ่น"
    if recipe != "เลือกหรือพิมพ์ค้นหา...":
        match = re.search(r'\d', recipe) # ค้นหาตัวเลขตัวแรกที่เจอในชื่อ
        if match and match.group() == '5':
            default_rubber_idx = 1 # เปลี่ยนค่าเป็น: 1 = "ยาง Ribbon"

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
        add_button = st.button("➕ เพิ่มรายการ", use_container_width=True)
        
    if add_button:
        if recipe == "เลือกหรือพิมพ์ค้นหา...":
            st.error("กรุณาเลือกหรือพิมพ์ชื่อ Recipe name!")
        else:
            batch_size_per_unit = 0
            if not master_data.empty and recipe in filtered_compounds:
                recipe_info = master_data[master_data['Compound SUB'] == recipe].iloc[0]
                batch_size_per_unit = recipe_info['Batch size (kg)']
            
            if unit == "Batch":
                total_kg = qty * batch_size_per_unit
            else:
                total_kg = qty
            
            st.session_state.order_list.append({
                "Customer": selected_sub,
                "Plan type": plan_type,
                "Recipe name": recipe,
                "Quantity": qty,
                "ประเภทยาง": rubber_type,
                "Unit": unit,
                "Delivery Date": delivery_date.strftime("%d.%m.%Y"),
                "Remark": remark,
                "Total Kg": round(total_kg, 3)
            })
            
            st.session_state.last_delivery_date = delivery_date
            st.session_state.order_confirmed = False
            
            st.success(f"เพิ่มรายการ {recipe} ลงในใบสั่งแล้ว!")
            st.rerun()

# -------------------------------------------------------------
# 5. แสดงตารางรายการที่สั่ง [แก้ไขที่ 2: ระบบลบรายการที่เลือก]
# -------------------------------------------------------------
if len(st.session_state.order_list) > 0:
    st.subheader("📋 สรุปรายการในใบสั่งยาง")
    
    # แปลงตะกร้าเป็น DataFrame และเพิ่มคอลัมน์ 'เลือก' (Checkbox)
    df_order = pd.DataFrame(st.session_state.order_list)
    df_order.insert(0, 'เลือกเพื่อลบ', False)
    
    st.write("💡 *คำแนะนำ: ติ๊กถูกที่หน้ารายการที่ต้องการลบ แล้วกดปุ่มลบรายการด้านล่าง*")
    edited_df = st.data_editor(
        df_order,
        use_container_width=True,
        hide_index=True,
        column_config={"เลือกเพื่อลบ": st.column_config.CheckboxColumn("เลือก", default=False)}
    )
    
    # สร้างปุ่มสำหรับลบรายการ และ ล้างทั้งหมด
    col_btn1, col_btn2, col_btn3 = st.columns([2, 2, 6])
    with col_btn1:
        if st.button("❌ ลบรายการที่เลือก"):
            # กรองเก็บเฉพาะแถวที่ "ไม่ได้ติ๊กเลือก"
            remaining_rows = edited_df[~edited_df['เลือกเพื่อลบ']].drop(columns=['เลือกเพื่อลบ'])
            st.session_state.order_list = remaining_rows.to_dict('records')
            st.rerun()
    with col_btn2:
        if st.button("🗑️ ล้างรายการทั้งหมด"):
            st.session_state.order_list = []
            st.rerun()
            
    st.divider()
    
    if st.button("🚀 ยืนยันการส่งใบสั่งยางไปยังโรงงาน", type="primary"):
        st.session_state.order_confirmed = True
        st.success("✅ ระบบบันทึกข้อมูลเรียบร้อยแล้ว คุณสามารถดาวน์โหลดไฟล์ Excel ได้ด้านล่าง")

# -------------------------------------------------------------
# 6. ฟังก์ชัน Export Excel
# -------------------------------------------------------------
if st.session_state.order_confirmed and len(st.session_state.order_list) > 0:
    df_export = pd.DataFrame(st.session_state.order_list)
    df_export.insert(0, 'No.', range(1, len(df_export) + 1))
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Order_Data')
    processed_data = output.getvalue()
    
    st.download_button(
        label="📥 ดาวน์โหลดใบสั่งยาง (Excel)",
        data=processed_data,
        file_name=f"Order_{selected_sub}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )