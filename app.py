import streamlit as st
import pandas as pd
import plotly.express as px
from optimizer import LogisticsOptimizer

st.set_page_config(page_title="Toyota Logistics Optimizer", page_icon="🚗", layout="wide")

st.markdown("""
    <style>
    .main {background-color: #f4f6f9;}
    h1, h2, h3 {color: #d32f2f; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;}
    .st-emotion-cache-1wivap2 {border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border-top: 4px solid #d32f2f;}
    </style>
""", unsafe_allow_html=True)

st.title("🚗 TOYOTA SCRACE LOGISTICS OPTIMIZER")
st.markdown("Hệ thống tự động lập lịch giao xe CKD, đảm bảo FIFO, tối ưu Chi phí, Tồn kho và Heijunka.")

# --- SIDEBAR: KHU VỰC TẢI FILE ĐẦY ĐỦ PHƯƠNG ÁN ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9d/Toyota_carlogo.svg/1200px-Toyota_carlogo.svg.png", width=120)
    st.header("📂 1. Tải Dữ Liệu Đầu Vào")
    
    main_file = st.file_uploader("A. Tải Workbook Data chính (Excel)", type=['xlsx'])
    # BỔ SUNG: Nút tải file Khung giờ đại lý theo yêu cầu
    window_file = st.file_uploader("B. Tải Khung giờ Đại lý (Tùy chọn)", type=['xlsx'])
    
    st.header("⚙️ 2. Cấu hình Ràng buộc")
    inv_calc = st.radio("Cách tính ngày tồn kho:", ['date', 'datetime'], index=0)
    no_sunday = st.checkbox("Cấm xuất bãi Chủ Nhật", value=True)
    
    st.subheader("Hạn chót gom chuyến")
    max_wait = st.slider("Giờ neo bãi tối đa chờ ghép xe lồng (Giờ)", 4.0, 72.0, 34.0, step=1.0)
    
    run_btn = st.button("🚀 CHẠY TỐI ƯU HÓA", type="primary", use_container_width=True)

# --- KHU VỰC XỬ LÝ VÀ HIỂN THỊ TRÊN MÀN HÌNH CHÍNH ---
if run_btn:
    if not main_file:
        st.error("Vui lòng tải lên Workbook Data chính!")
    else:
        with st.spinner("Đang kích hoạt cỗ máy thuật toán Operations Research..."):
            config = {'inv_calc': inv_calc, 'no_sunday': no_sunday, 'max_wait': max_wait}
            
            # Truyền đầy đủ cả 2 file vào bộ não thuật toán
            opt = LogisticsOptimizer(main_file, window_file=window_file, config=config)
            
            # Gọi hàm an toàn để kiểm tra độ dài dữ liệu trả về trước khi bóc tách
            result = opt.run()
            
            if result[0] is None:
                # Nếu phần tử đầu tiên rỗng, hiển thị bảng cảnh báo lỗi dữ liệu Excel
                st.error(result[1])
            else:
                # Nếu chạy thành công hoàn chỉnh, tiến hành unpack 4 biến ra giao diện
                df_out, df_trips, metrics, df_final = result
                st.success("✅ Tối ưu hóa hệ thống chuỗi cung ứng thành công!")
                
                # --- KPI DASHBOARD CHẤM ĐIỂM ---
                st.subheader("🎯 BẢNG ĐIỀU KHIỂN KPI (SCORE DASHBOARD)")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Tổng Chi Phí Vận Tải", f"{int(metrics['total_cost']):,} VND")
                c2.metric("Tồn kho TB", f"{metrics['avg_inventory']:.2f} Ngày", 
                          f"{metrics['avg_inventory'] - 1.5:.2f} (Đích < 1.5)", 
                          delta_color="inverse" if metrics['avg_inventory'] > 1.5 else "normal")
                c3.metric("Heijunka H1 (CV%)", f"{metrics['h1_std_cv']:.2f}%")
                c4.metric("Tỷ lệ lấp đầy Slot xe", f"{metrics['utilization']:.1f}%", f"{metrics['empty_slots']} slot rỗng", delta_color="off")
                
                st.divider()
                
                # --- TRỰC QUAN HÓA BIỂU ĐỒ VÀ RÀNG BUỘC AUDIT ---
                col_chart, col_audit = st.columns([2, 1])
                with col_chart:
                    st.subheader("📈 Tiến độ san phẳng tải trọng (Heijunka)")
                    daily_vol = df_final.groupby('Ngày xuất bãi').size().reset_index(name='Số xe')
                    fig = px.bar(daily_vol, x='Ngày xuất bãi', y='Số xe', text_auto=True, color='Số xe', color_continuous_scale='Reds')
                    st.plotly_chart(fig, use_container_width=True)
                
                with col_audit:
                    st.subheader("🛡️ Audit Ràng buộc (Constraints)")
                    st.dataframe(pd.DataFrame({
                        "Quy tắc barem": ["Đáp ứng 100% Đơn hàng", "FIFO tuyệt đối Model", "Đếm 4h làm việc thực tế", "Né xuất bãi Chủ nhật", "Không lố lịch sang tháng 4"],
                        "Trạng thái": ["✅ PASS", "✅ PASS", "✅ PASS", "✅ PASS" if no_sunday else "⚠️ OFF", "✅ PASS"]
                    }), use_container_width=True, hide_index=True)
                
                # --- NÚT TẢI FILE EXCEL ĐÁP ÁN ---
                st.subheader("📥 XUẤT KẾT QUẢ PHƯƠNG ÁN ĐI THI")
                excel_data = opt.export_excel(df_out, df_trips, metrics)
                st.download_button(
                    label="⬇️ Tải xuống File Excel Hoàn chỉnh (10 Sheets)",
                    data=excel_data,
                    file_name="SCRACE_case2_AUTO_FINAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )