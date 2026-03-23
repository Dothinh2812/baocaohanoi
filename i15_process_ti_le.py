# -*- coding: utf-8 -*-
"""
Module tính toán các tỉ lệ cho báo cáo I1.5 (Suy hao cao K1/K2)
"""

import pandas as pd
import os


def process_i15_ti_le_shc_k1():
    """
    Tính toán các tỉ lệ Số TB Suy hao cao K1:
    
    Đọc sheet 'TH_SHC_I15' từ file I1.5 report.xlsx và tính:
    1. Tỉ lệ Số TB SHC của từng cá nhân/tổng TB SHC của tổ
    2. Tổng TB SHC của từng tổ/ tổng TB SHC của toàn trung tâm
    3. Tỉ lệ Số TB SHC của từng cá nhân/tổng TB SHC của toàn trung tâm
    4. Tỉ lệ SHC/số TB quản lý (nếu có)
    
    Kết quả được ghi vào sheet mới 'Ti_Le_SHC_K1' trong cùng file I1.5 report.xlsx
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TÍNH TOÁN TỈ LỆ SỐ TB SUY HAO CAO K1")
        print("="*80)
        
        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "I1.5 report.xlsx")
        
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False
        
        print(f"\n✓ Đang đọc sheet 'TH_SHC_I15' từ file: {input_file}")
        
        # Đọc sheet TH_SHC_I15
        try:
            df = pd.read_excel(input_file, sheet_name='TH_SHC_I15')
            print(f"✅ Đã đọc sheet, tổng số dòng: {len(df)}")
        except Exception as e:
            print(f"❌ Không thể đọc sheet 'TH_SHC_I15': {e}")
            print("⚠️ Vui lòng chạy hàm process_I15_report_with_tracking() trước")
            return False
        
        # Kiểm tra các cột cần thiết
        required_columns = ['NVKT_DB', 'Đơn vị', 'Số TB Suy hao cao K1']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"❌ Không tìm thấy các cột: {', '.join(missing_columns)}")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False
        
        print("\n✓ Đang tính toán các tỉ lệ SHC K1...")
        
        # Tính tổng số TB SHC của toàn trung tâm
        tong_shc_trung_tam = df['Số TB Suy hao cao K1'].sum()
        print(f"  - Tổng số TB Suy hao cao toàn trung tâm: {tong_shc_trung_tam}")
        
        # Tính tổng số TB SHC của từng tổ
        df_to = df.groupby('Đơn vị')['Số TB Suy hao cao K1'].sum().reset_index()
        df_to.columns = ['Đơn vị', 'Tổng SHC tổ']
        
        print(f"  - Số lượng tổ: {len(df_to)}")
        
        # Merge để thêm cột "Tổng SHC tổ" vào dataframe gốc
        df_result = pd.merge(df, df_to, on='Đơn vị', how='left')
        
        # Tính các tỉ lệ
        # 1. Tỉ lệ cá nhân/tổ (%)
        df_result['Tỉ lệ cá nhân/tổ (%)'] = df_result.apply(
            lambda row: round((row['Số TB Suy hao cao K1'] / row['Tổng SHC tổ'] * 100), 2) 
            if row['Tổng SHC tổ'] > 0 else 0,
            axis=1
        )
        
        # 2. Tỉ lệ tổ/trung tâm (%)
        df_result['Tỉ lệ tổ/trung tâm (%)'] = df_result.apply(
            lambda row: round((row['Tổng SHC tổ'] / tong_shc_trung_tam * 100), 2)
            if tong_shc_trung_tam > 0 else 0,
            axis=1
        )
        
        # 3. Tỉ lệ cá nhân/trung tâm (%)
        df_result['Tỉ lệ cá nhân/trung tâm (%)'] = df_result.apply(
            lambda row: round((row['Số TB Suy hao cao K1'] / tong_shc_trung_tam * 100), 2)
            if tong_shc_trung_tam > 0 else 0,
            axis=1
        )
        
        print(f"✅ Đã tính toán các tỉ lệ cho {len(df_result)} NVKT")
        
        # Chọn các cột cần hiển thị
        output_columns = [
            'Đơn vị',
            'NVKT_DB',
            'Số TB Suy hao cao K1',
            'Tổng SHC tổ',
            'Tỉ lệ cá nhân/tổ (%)',
            'Tỉ lệ tổ/trung tâm (%)',
            'Tỉ lệ cá nhân/trung tâm (%)'
        ]
        
        # Thêm các cột khác nếu có  
        if 'Số TB quản lý' in df_result.columns:
            output_columns.insert(3, 'Số TB quản lý')
        
        if 'Tỉ lệ SHC (%)' in df_result.columns:
            output_columns.insert(4 if 'Số TB quản lý' in output_columns else 3, 'Tỉ lệ SHC (%)')
        
        df_output = df_result[output_columns].copy()
        
        # Sắp xếp theo Đơn vị và NVKT_DB
        df_output = df_output.sort_values(['Đơn vị', 'NVKT_DB']).reset_index(drop=True)
        
        # Thêm cột TT (Thứ tự)
        df_output.insert(0, 'TT', range(1, len(df_output) + 1))
        
        # Ghi vào sheet mới 'Ti_Le_SHC_K1'
        print("\n✓ Đang ghi vào sheet mới 'Ti_Le_SHC_K1'...")
        
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_output.to_excel(writer, sheet_name='Ti_Le_SHC_K1', index=False)
        
        print(f"✅ Đã ghi dữ liệu vào sheet 'Ti_Le_SHC_K1'")
        
        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN:")
        print(f"  - Tổng số tổ: {df_output['Đơn vị'].nunique()}")
        print(f"  - Tổng số NVKT: {len(df_output) - 1}")  # Trừ dòng header
        print(f"  - Tổng số TB Suy hao cao toàn trung tâm: {tong_shc_trung_tam}")
        if 'Số TB quản lý' in df_output.columns:
            print(f"  - Tổng số TB quản lý: {df_output['Số TB quản lý'].sum()}")
        print(f"  - Tỉ lệ cá nhân/tổ trung bình: {df_output['Tỉ lệ cá nhân/tổ (%)'].mean():.2f}%")
        print(f"  - Tỉ lệ cá nhân/trung tâm trung bình: {df_output['Tỉ lệ cá nhân/trung tâm (%)'].mean():.2f}%")
        print("-"*80)
        
        # In một vài dòng mẫu
        print("\n📊 Một số dòng dữ liệu mẫu:")
        print(df_output.head(10).to_string(index=False))
        
        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TÍNH TOÁN TỈ LỆ SỐ TB SUY HAO CAO K1")
        print("="*80)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Lỗi khi tính toán tỉ lệ SHC K1: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_i15_ti_le_shc_k2():
    """
    Tính toán các tỉ lệ Số TB Suy hao cao K2:
    Tương tự như K1 nhưng cho file I1.5_k2 report.xlsx
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TÍNH TOÁN TỈ LỆ SỐ TB SUY HAO CAO K2")
        print("="*80)
        
        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "I1.5_k2 report.xlsx")
        
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False
        
        print(f"\n✓ Đang đọc sheet 'TH_SHC_I15' từ file: {input_file}")
        
        # Đọc sheet TH_SHC_I15
        try:
            df = pd.read_excel(input_file, sheet_name='TH_SHC_I15')
            print(f"✅ Đã đọc sheet, tổng số dòng: {len(df)}")
        except Exception as e:
            print(f"❌ Không thể đọc sheet 'TH_SHC_I15': {e}")
            print("⚠️ Vui lòng chạy hàm process_I15_k2_report_with_tracking() trước")
            return False
        
        # Kiểm tra các cột cần thiết
        required_columns = ['NVKT_DB', 'Đơn vị', 'Số TB Suy hao cao K2']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"❌ Không tìm thấy các cột: {', '.join(missing_columns)}")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False
        
        print("\n✓ Đang tính toán các tỉ lệ SHC K2...")
        
        # Tính tổng số TB SHC của toàn trung tâm
        tong_shc_trung_tam = df['Số TB Suy hao cao K2'].sum()
        print(f"  - Tổng số TB Suy hao cao toàn trung tâm: {tong_shc_trung_tam}")
        
        # Tính tổng số TB SHC của từng tổ
        df_to = df.groupby('Đơn vị')['Số TB Suy hao cao K2'].sum().reset_index()
        df_to.columns = ['Đơn vị', 'Tổng SHC tổ']
        
        print(f"  - Số lượng tổ: {len(df_to)}")
        
        # Merge để thêm cột "Tổng SHC tổ" vào dataframe gốc
        df_result = pd.merge(df, df_to, on='Đơn vị', how='left')
        
        # Tính các tỉ lệ
        # 1. Tỉ lệ cá nhân/tổ (%)
        df_result['Tỉ lệ cá nhân/tổ (%)'] = df_result.apply(
            lambda row: round((row['Số TB Suy hao cao K2'] / row['Tổng SHC tổ'] * 100), 2) 
            if row['Tổng SHC tổ'] > 0 else 0,
            axis=1
        )
        
        # 2. Tỉ lệ tổ/trung tâm (%)
        df_result['Tỉ lệ tổ/trung tâm (%)'] = df_result.apply(
            lambda row: round((row['Tổng SHC tổ'] / tong_shc_trung_tam * 100), 2)
            if tong_shc_trung_tam > 0 else 0,
            axis=1
        )
        
        # 3. Tỉ lệ cá nhân/trung tâm (%)
        df_result['Tỉ lệ cá nhân/trung tâm (%)'] = df_result.apply(
            lambda row: round((row['Số TB Suy hao cao K2'] / tong_shc_trung_tam * 100), 2)
            if tong_shc_trung_tam > 0 else 0,
            axis=1
        )
        
        print(f"✅ Đã tính toán các tỉ lệ cho {len(df_result)} NVKT")
        
        # Chọn các cột cần hiển thị
        output_columns = [
            'Đơn vị',
            'NVKT_DB',
            'Số TB Suy hao cao K2',
            'Tổng SHC tổ',
            'Tỉ lệ cá nhân/tổ (%)',
            'Tỉ lệ tổ/trung tâm (%)',
            'Tỉ lệ cá nhân/trung tâm (%)'
        ]
        
        # Thêm các cột khác nếu có
        if 'Số TB quản lý' in df_result.columns:
            output_columns.insert(3, 'Số TB quản lý')
        
        if 'Tỉ lệ SHC (%)' in df_result.columns:
            output_columns.insert(4 if 'Số TB quản lý' in output_columns else 3, 'Tỉ lệ SHC (%)')
        
        df_output = df_result[output_columns].copy()
        
        # Sắp xếp theo Đơn vị và NVKT_DB
        df_output = df_output.sort_values(['Đơn vị', 'NVKT_DB']).reset_index(drop=True)
        
        # Thêm cột TT (Thứ tự)
        df_output.insert(0, 'TT', range(1, len(df_output) + 1))
        
        # Ghi vào sheet mới 'Ti_Le_SHC_K2'
        print("\n✓ Đang ghi vào sheet mới 'Ti_Le_SHC_K2'...")
        
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_output.to_excel(writer, sheet_name='Ti_Le_SHC_K2', index=False)
        
        print(f"✅ Đã ghi dữ liệu vào sheet 'Ti_Le_SHC_K2'")
        
        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN:")
        print(f"  - Tổng số tổ: {df_output['Đơn vị'].nunique()}")
        print(f"  - Tổng số NVKT: {len(df_output) - 1}")
        print(f"  - Tổng số TB Suy hao cao toàn trung tâm: {tong_shc_trung_tam}")
        if 'Số TB quản lý' in df_output.columns:
            print(f"  - Tổng số TB quản lý: {df_output['Số TB quản lý'].sum()}")
        print(f"  - Tỉ lệ cá nhân/tổ trung bình: {df_output['Tỉ lệ cá nhân/tổ (%)'].mean():.2f}%")
        print(f"  - Tỉ lệ cá nhân/trung tâm trung bình: {df_output['Tỉ lệ cá nhân/trung tâm (%)'].mean():.2f}%")
        print("-"*80)
        
        # In một vài dòng mẫu
        print("\n📊 Một số dòng dữ liệu mẫu:")
        print(df_output.head(10).to_string(index=False))
        
        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TÍNH TOÁN TỈ LỆ SỐ TB SUY HAO CAO K2")
        print("="*80)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Lỗi khi tính toán tỉ lệ SHC K2: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # Test K1
    process_i15_ti_le_shc_k1()
    
    # Test K2
    # process_i15_ti_le_shc_k2()
