import streamlit as st
import pandas as pd
import io

# ตั้งค่าหน้าตาของเว็บเบื้องต้น
st.set_page_config(page_title="ระบบคำนวณค่าคอมมิชชั่น", page_icon="📊", layout="wide")

# ฟังก์ชันล้างข้อมูลตัวเลขเดิมของคุณ
def clean_num(val):
    if pd.isna(val) or val is None:
        return 0.0
    try:
        return float(str(val).strip().replace(',', ''))
    except:
        return 0.0

# ส่วนหัวของหน้าเว็บ
st.title("📊 ระบบคำนวณค่าคอมมิชชั่นพนักงานขาย")
st.subheader("อัปโหลดไฟล์ Excel เพื่อคำนวณและสรุปยอดอัตโนมัติ")
st.markdown("---")

# สร้างส่วนอัปโหลดไฟล์ (แบ่งเป็น 2 คอลัมน์ซ้าย-ขวา)
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📁 1. ไฟล์ราคากลาง")
    cp_file = st.file_uploader("เลือกไฟล์ราคากลาง (center_price.xlsx)", type=["xlsx"])

with col2:
    st.markdown("### 🧾 2. ไฟล์ใบเสร็จ")
    rc_file = st.file_uploader("เลือกไฟล์ใบเสร็จ (receipt_data.xlsx)", type=["xlsx"])

# เมื่อผู้ใช้งานอัปโหลดไฟล์ครบทั้งคู่แล้ว ระบบจะทำงานทันที
if cp_file is not None and rc_file is not None:
    try:
        # โหลดไฟล์จากข้อมูลที่อัปโหลดเข้ามา
        df_cp = pd.read_excel(cp_file)
        df_rc = pd.read_excel(rc_file)

        # ล็อกตำแหน่งคอลัมน์ตามลอจิกของคุณ
        cp_id_name = df_cp.columns[0]
        cp_price_name = df_cp.columns[1]
        cp_rate_name = df_cp.columns[2]
        
        price_map = {}
        for _, row in df_cp.iterrows():
            prod_id = str(row[cp_id_name]).strip()
            std_price = clean_num(row[cp_price_name])
            comm_rate = clean_num(row[cp_rate_name])
            price_map[prod_id] = [std_price, comm_rate]

        rc_sale_name = df_rc.columns[0]   # คอลัมน์ที่ 1 (A): รหัสพนักงานขาย
        rc_bill_name = df_rc.columns[1]   # คอลัมน์ที่ 2 (B): เลขที่ใบเสร็จ
        rc_id_name = df_rc.columns[2]     # คอลัมน์ที่ 3 (C): รหัสสินค้า
        rc_qty_name = df_rc.columns[3]    # คอลัมน์ที่ 4 (D): จำนวนขายจริง

        # ประมวลผลคำนวณค่าคอมมิชชั่น
        sale_list, receipt_list, product_list, qty_list = [], [], [], []
        std_price_list, base_sales_list, rate_list, commission_list = [], [], [], []

        for _, row in df_rc.iterrows():
            sale_id = str(row[rc_sale_name]).strip() 
            bill_no = str(row[rc_bill_name]).strip()
            prod_id = str(row[rc_id_name]).strip()
            qty = clean_num(row[rc_qty_name]) 
            
            if prod_id in price_map:
                standard_price = price_map[prod_id][0]
                raw_rate = price_map[prod_id][1]
            else:
                standard_price = 0.0
                raw_rate = 0.0
                
            commission_rate = raw_rate / 100 
            base_sales = standard_price * qty
            commission_amount = base_sales * commission_rate
            
            sale_list.append(sale_id)
            receipt_list.append(bill_no)
            product_list.append(prod_id)
            qty_list.append(qty)
            std_price_list.append(standard_price)
            base_sales_list.append(base_sales)
            rate_list.append(commission_rate)
            commission_list.append(commission_amount)

        df_final = pd.DataFrame({
            'รหัสพนักงานขาย': sale_list,
            'เลขที่ใบเสร็จ': receipt_list,
            'รหัสสินค้า': product_list,
            'ราคากลาง': std_price_list,
            'จำนวนขายจริง': qty_list,
            'ยอดรวมตามราคากลาง': base_sales_list,
            'เปอร์เซ็นต์ค่าคอม': rate_list,
            'ยอดคอมมิชชั่น': commission_list
        })

        df_final = df_final.sort_values(by=['รหัสพนักงานขาย', 'เลขที่ใบเสร็จ'])

        # ดึงยอดรวมใหญ่ทั้งหมด
        total_commission_all = df_final['ยอดคอมมิชชั่น'].sum()

        # จัดทำตารางสรุปเพื่อเตรียมทำไฟล์ดาวน์โหลด
        df_summary_sale = df_final.groupby('รหัสพนักงานขาย')['ยอดคอมมิชชั่น'].sum().reset_index()
        df_summary_sale.columns = ['รหัสพนักงานขาย', 'รวมค่าคอมมิชชั่นสุทธิ']
        total_row_sale = pd.DataFrame([['รวมค่าคอมมิชชั่นทั้งหมดทั้งสิ้น', total_commission_all]], columns=['รหัสพนักงานขาย', 'รวมค่าคอมมิชชั่นสุทธิ'])
        df_summary_sale_final = pd.concat([df_summary_sale, total_row_sale], ignore_index=True)

        df_summary_receipt = df_final.groupby(['รหัสพนักงานขาย', 'เลขที่ใบเสร็จ'])['ยอดคอมมิชชั่น'].sum().reset_index()
        df_summary_receipt.columns = ['รหัสพนักงานขาย', 'เลขที่ใบเสร็จ', 'รวมค่าคอมมิชชั่นสุทธิ']

        # สร้างไฟล์ Excel ในหน่วยความจำ (Memory Buffer) เพื่อเตรียมปุ่มดาวน์โหลดบนเว็บ
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_summary_sale_final.to_excel(writer, sheet_name='สรุปตามพนักงานขาย', index=False)
            df_summary_receipt.to_excel(writer, sheet_name='สรุปตามใบเสร็จ', index=False)
            df_final.to_excel(writer, sheet_name='รายละเอียดทั้งหมด', index=False)
        buffer.seek(0)

        # ------------------ ส่วนแสดงผลบนหน้าเว็บ (UI) ------------------
        st.success("🎉 คำนวณข้อมูลสำเร็จเรียบร้อย!")
        
        # กล่องแสดงยอดรวมใหญ่โดดเด่น (Metric Card)
        st.metric(label="💰 ยอดรวมค่าคอมมิชชั่นทั้งหมดของบริษัท", value=f"{total_commission_all:,.2f} บาท")
        
        # ปุ่มดาวน์โหลดไฟล์รายงาน Excel สรุป
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์รายงานสรุป (Excel)",
            data=buffer,
            file_name="final_commission_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        # ทำ Tab แยกแสดงผลบนหน้าจอให้ดูง่าย สวยงาม
        tab1, tab2, tab3 = st.tabs(["👤 สรุปตามพนักงานขาย", "🧾 สรุปตามใบเสร็จ", "📄 รายละเอียดทั้งหมด"])
        
        with tab1:
            st.dataframe(df_summary_sale, use_container_width=True)
            
        with tab2:
            st.dataframe(df_summary_receipt, use_container_width=True)
            
        with tab3:
            st.dataframe(df_final, use_container_width=True)

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลข้อมูล: {e}")
        st.info("กรุณาตรวจสอบว่าไฟล์ Excel ที่อัปโหลดมีโครงสร้างตารางและหัวคอลัมน์ที่ถูกต้อง")
else:
    st.info("💡 คำแนะนำ: กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ไฟล์เพื่อเริ่มกระบวนการคำนวณครับ")
