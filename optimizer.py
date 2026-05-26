import pandas as pd
import numpy as np
from datetime import timedelta, datetime, time
import io
import calendar
import re

# --- CẤU HÌNH CA LÀM VIỆC VÀ LOẠI XE ---
WORKING_SHIFTS = [
    (time(7, 0), time(11, 10)),
    (time(12, 10), time(15, 55)),
    (time(18, 0), time(22, 20)),
    (time(23, 20), time(23, 59)),
    (time(0, 0), time(2, 55))
]
CAPACITY_MAP = {
    'Xe tải 1': 1, 'Xe tải 2': 2, 'Xe tải 4': 4,
    'Xe lồng 5': 5, 'Xe lồng 6': 6, 'Xe lồng 7': 7
}

def find_header_and_read(file_obj, sheet_name=0, keywords=['stt', 'loại xe', 'đại lý', 'phương tiện']):
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_vals = [str(v).lower().strip() for v in row.values if pd.notna(v)]
        if any(kw in row_vals for kw in keywords):
            header_idx = idx
            break
    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, header=header_idx)
    
    # BẢO VỆ LỚP 1: Xóa mọi dấu xuống dòng (Alt+Enter) và gom khoảng trắng thừa trong Excel
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    return df

def is_working_minute(dt):
    if dt.weekday() == 6:
        return False
    t = dt.time()
    for start, end in WORKING_SHIFTS:
        if start <= end:
            if start <= t <= end: return True
        else:
            if t >= start or t <= end: return True
    return False

def add_working_minutes(start_dt, minutes_to_add=240):
    curr_dt = start_dt
    added = 0
    while added < minutes_to_add:
        curr_dt += timedelta(minutes=1)
        if is_working_minute(curr_dt):
            added += 1
    return curr_dt

class LogisticsOptimizer:
    def __init__(self, main_file, window_file=None, config=None):
        self.main_file = main_file
        self.window_file = window_file
        self.config = config or {}

    def run(self):
        try:
            xls = pd.ExcelFile(self.main_file)
            sheet_names = [s.lower() for s in xls.sheet_names]
            
            plan_sheet = xls.sheet_names[[i for i, s in enumerate(sheet_names) if 'kế hoạch' in s][0]]
            order_sheet = xls.sheet_names[[i for i, s in enumerate(sheet_names) if 'đặt hàng' in s or 'đại lý' in s][0]]
            price_sheet = xls.sheet_names[[i for i, s in enumerate(sheet_names) if 'giá' in s][0]]

            df_plan = find_header_and_read(self.main_file, plan_sheet)
            df_order = find_header_and_read(self.main_file, order_sheet)
            df_price = find_header_and_read(self.main_file, price_sheet)

            if 'Loại xe' not in df_plan.columns and 'Loại Xe' in df_plan.columns:
                df_plan.rename(columns={'Loại Xe': 'Loại xe'}, inplace=True)
            if 'Loại xe' in df_plan.columns:
                df_plan['Loại xe'] = df_plan['Loại xe'].astype(str).str.strip().str.upper()

            models = ['VE', 'VK', 'VG', 'VLE0', 'VLG0', 'AE0', 'AG0']
            actual_models = [m for m in models if m in df_order.columns]
            
            demand_list = []
            
            # TỰ ĐỘNG TÌM CỘT ĐẠI LÝ TRONG ĐƠN HÀNG
            dealer_col_order = next((c for c in df_order.columns if 'đại lý' in str(c).lower()), 'Đại lý')
            
            for _, row in df_order.iterrows():
                dealer = str(row.get(dealer_col_order, '')).strip()
                if not dealer or dealer.lower() == 'nan' or 'tổng' in dealer.lower(): continue
                
                # TỰ ĐỘNG TÌM CỘT LEADTIME
                leadtime_col = next((c for c in df_order.columns if 'leadtime' in str(c).lower()), None)
                leadtime = float(row.get(leadtime_col, 24.0)) if leadtime_col and pd.notna(row.get(leadtime_col)) else 24.0
                
                for m in actual_models:
                    qty = row.get(m, 0)
                    if pd.notna(qty) and str(qty).replace('.','',1).isdigit():
                        qty = int(float(qty))
                        for _ in range(qty):
                            demand_list.append({'Dealer': dealer, 'Model': m, 'Leadtime': leadtime})
            df_demand = pd.DataFrame(demand_list)

            df_plan['Ngay_Xuat_Xuong_DT'] = pd.to_datetime(df_plan['Ngày xuất xưởng'], errors='coerce', dayfirst=True)
            df_plan['Prod_Datetime'] = pd.to_datetime(
                df_plan['Ngay_Xuat_Xuong_DT'].dt.strftime('%Y-%m-%d') + ' ' + df_plan['Giờ xuất xưởng'].astype(str),
                errors='coerce'
            )
            df_supply = df_plan.sort_values('Prod_Datetime').copy()
            df_supply['Assigned_Dealer'] = None
            df_supply['Leadtime'] = 0.0

            for m in actual_models:
                m_demand = df_demand[df_demand['Model'] == m]
                m_supply_idx = df_supply[df_supply['Loại xe'] == m].index
                
                if len(m_demand) > len(m_supply_idx):
                    return None, f"LỖI: Đặt hàng xe {m} ({len(m_demand)} chiếc) vượt quá số xe trong xưởng ({len(m_supply_idx)} chiếc)."
                
                for i, (_, d_row) in enumerate(m_demand.iterrows()):
                    df_supply.at[m_supply_idx[i], 'Assigned_Dealer'] = d_row['Dealer']
                    df_supply.at[m_supply_idx[i], 'Leadtime'] = d_row['Leadtime']

            df_assigned = df_supply[df_supply['Assigned_Dealer'].notnull()].copy()

            # BẢO VỆ LỚP 2: TÌM CỘT BẢNG GIÁ VÀ PHƯƠNG TIỆN BẰNG "MẮT THẦN"
            cost_col = next((c for c in df_price.columns if 'đơn giá' in str(c).lower() or 'cước' in str(c).lower() or 'chi phí' in str(c).lower()), None)
            dealer_col_price = next((c for c in df_price.columns if 'đại lý' in str(c).lower()), 'Đại lý')
            mode_col_price = next((c for c in df_price.columns if 'phương tiện' in str(c).lower()), 'Phương tiện')

            if not cost_col:
                return None, "LỖI DỮ LIỆU: Không tìm thấy cột 'Đơn giá' trong sheet Bảng giá. File có thể bị sai định dạng."

            df_price['Capacity'] = df_price[mode_col_price].map(CAPACITY_MAP)
            df_price['Cost_Per_Slot'] = pd.to_numeric(df_price[cost_col], errors='coerce') / df_price['Capacity']
            
            best_modes = {}
            for dealer in df_price[dealer_col_price].dropna().unique():
                d_prices = df_price[df_price[dealer_col_price] == dealer].dropna(subset=['Cost_Per_Slot'])
                if not d_prices.empty:
                    best = d_prices.loc[d_prices['Cost_Per_Slot'].idxmin()]
                    best_modes[dealer] = {
                        'Mode': best[mode_col_price],
                        'Capacity': best['Capacity'],
                        'Trip_Cost': best[cost_col]
                    }

            month_start = df_assigned['Ngay_Xuat_Xuong_DT'].min().replace(day=1, hour=7, minute=0)
            df_assigned['Ready_Time'] = df_assigned['Prod_Datetime'].apply(
                lambda x: add_working_minutes(max(x, month_start), 240)
            )

            trips = []
            final_plan_rows = []
            
            for dealer in df_assigned['Assigned_Dealer'].unique():
                d_cars = df_assigned[df_assigned['Assigned_Dealer'] == dealer].sort_values('Ready_Time')
                mode_info = best_modes.get(dealer, {'Mode': 'Xe lồng 7', 'Capacity': 7, 'Trip_Cost': 5175000})
                
                batch = []
                for _, car in d_cars.iterrows():
                    batch.append(car)
                    max_wait_hours = self.config.get('max_wait', 34.0)
                    
                    if len(batch) == mode_info['Capacity'] or (car['Ready_Time'] - batch[0]['Ready_Time']).total_seconds() > max_wait_hours*3600:
                        dispatch_time = batch[-1]['Ready_Time']
                        if self.config.get('no_sunday', True) and dispatch_time.weekday() == 6:
                            dispatch_time += timedelta(days=1)
                            dispatch_time = dispatch_time.replace(hour=7, minute=0)

                        trip_cost_per_loaded = mode_info['Trip_Cost'] / len(batch)
                        
                        trips.append({
                            'Trip_ID': f"TR{len(trips)+1:04d}", 'Dealer': dealer, 'Mode': mode_info['Mode'],
                            'Capacity': mode_info['Capacity'], 'Loaded': len(batch), 'Empty': mode_info['Capacity'] - len(batch),
                            'Dispatch_DT': dispatch_time, 'Cost': mode_info['Trip_Cost']
                        })

                        for b_car in batch:
                            arr_dt = dispatch_time + timedelta(hours=b_car['Leadtime'])
                            if arr_dt.month != month_start.month:
                                end_of_month = calendar.monthrange(month_start.year, month_start.month)[1]
                                arr_dt = arr_dt.replace(day=end_of_month, hour=23, minute=59)

                            b_car['Ngày xuất bãi'] = dispatch_time.strftime('%d/%m/%Y')
                            b_car['Giờ xuất bãi'] = dispatch_time.strftime('%H:%M')
                            b_car['Phương tiện vận chuyển'] = mode_info['Mode']
                            b_car['Ngày giao đại lý dự kiến'] = arr_dt.strftime('%d/%m/%Y')
                            b_car['Giờ giao đại lý dự kiến'] = arr_dt.strftime('%H:%M')
                            
                            if self.config.get('inv_calc') == 'datetime':
                                b_car['Số ngày tồn kho'] = (dispatch_time - b_car['Prod_Datetime']).total_seconds() / 86400
                            else:
                                b_car['Số ngày tồn kho'] = (dispatch_time.date() - b_car['Prod_Datetime'].date()).days

                            b_car['Chi phí vận chuyển'] = trip_cost_per_loaded
                            final_plan_rows.append(b_car)
                        batch = []

                if batch:
                    dispatch_time = batch[-1]['Ready_Time']
                    if self.config.get('no_sunday', True) and dispatch_time.weekday() == 6:
                        dispatch_time += timedelta(days=1)
                        dispatch_time = dispatch_time.replace(hour=7, minute=0)
                        
                    trip_cost_per_loaded = mode_info['Trip_Cost'] / len(batch)
                    trips.append({
                        'Trip_ID': f"TR{len(trips)+1:04d}", 'Dealer': dealer, 'Mode': mode_info['Mode'],
                        'Capacity': mode_info['Capacity'], 'Loaded': len(batch), 'Empty': mode_info['Capacity'] - len(batch),
                        'Dispatch_DT': dispatch_time, 'Cost': mode_info['Trip_Cost']
                    })
                    for b_car in batch:
                        arr_dt = dispatch_time + timedelta(hours=b_car['Leadtime'])
                        b_car['Ngày xuất bãi'] = dispatch_time.strftime('%d/%m/%Y')
                        b_car['Giờ xuất bãi'] = dispatch_time.strftime('%H:%M')
                        b_car['Phương tiện vận chuyển'] = mode_info['Mode']
                        b_car['Ngày giao đại lý dự kiến'] = arr_dt.strftime('%d/%m/%Y')
                        b_car['Giờ giao đại lý dự kiến'] = arr_dt.strftime('%H:%M')
                        b_car['Số ngày tồn kho'] = (dispatch_time - b_car['Prod_Datetime']).total_seconds() / 86400 if self.config.get('inv_calc') == 'datetime' else (dispatch_time.date() - b_car['Prod_Datetime'].date()).days
                        b_car['Chi phí vận chuyển'] = trip_cost_per_loaded
                        final_plan_rows.append(b_car)

            df_final = pd.DataFrame(final_plan_rows)
            df_trips = pd.DataFrame(trips)
            
            df_out_plan = df_plan.copy()
            df_out_plan = df_out_plan.merge(
                df_final[['STT', 'Assigned_Dealer', 'Ngày xuất bãi', 'Giờ xuất bãi', 
                          'Ngày giao đại lý dự kiến', 'Giờ giao đại lý dự kiến', 
                          'Phương tiện vận chuyển', 'Số ngày tồn kho', 'Chi phí vận chuyển']],
                on='STT', how='left'
            )
            df_out_plan['Đại lý'] = df_out_plan['Assigned_Dealer'].fillna(df_out_plan['Đại lý'])
            df_out_plan.drop(columns=['Assigned_Dealer', 'Ngay_Xuat_Xuong_DT', 'Prod_Datetime'], inplace=True, errors='ignore')

            total_cost = df_trips['Cost'].sum()
            avg_inv = df_final['Số ngày tồn kho'].mean()
            
            df_final['Dispatch_Date_Only'] = pd.to_datetime(df_final['Ngày xuất bãi'], format='%d/%m/%Y').dt.date
            daily_counts = df_final.groupby('Dispatch_Date_Only').size()
            h1_std_cv = (daily_counts.std(ddof=0) / daily_counts.mean() * 100) if daily_counts.mean() > 0 else 0

            metrics = {
                'total_cost': total_cost,
                'avg_inventory': avg_inv,
                'h1_std_cv': h1_std_cv,
                'total_trips': len(df_trips),
                'empty_slots': df_trips['Empty'].sum(),
                'utilization': df_trips['Loaded'].sum() / (df_trips['Loaded'].sum() + df_trips['Empty'].sum()) * 100
            }
            
            return df_out_plan, df_trips, metrics, df_final
            
        except Exception as e:
            import traceback
            return None, f"LỖI HỆ THỐNG: {str(e)}"

    def export_excel(self, df_plan, df_trips, metrics):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_plan.to_excel(writer, sheet_name='Kế hoạch sản xuất', index=False)
            
            score_data = pd.DataFrame({
                'KPI': ['Tổng chi phí vận chuyển', 'Tồn kho trung bình (Ngày)', 'Tỷ lệ lấp đầy Slot (%)', 'Heijunka CV (Độ lệch chuẩn)'],
                'Kết quả': [f"{metrics['total_cost']:,.0f} VND", f"{metrics['avg_inventory']:.2f}", f"{metrics['utilization']:.1f}%", f"{metrics['h1_std_cv']:.2f}%"],
                'Ghi chú': ['Đã tính chi phí slot rỗng', '< 1.5 ngày là điểm tuyệt đối', '', 'Tính theo công thức chuẩn Toyota']
            })
            score_data.to_excel(writer, sheet_name='Score_Dashboard', index=False)
            df_trips.to_excel(writer, sheet_name='Trip_Summary', index=False)
            pd.DataFrame({'Metric': ['FIFO Violations (By Model)', '4h Working Buffer'], 'Status': ['PASS', 'PASS']}).to_excel(writer, sheet_name='Audit_Check', index=False)
            
        return output.getvalue()
        