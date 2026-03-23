#!/usr/bin/env python3
"""
Script để bổ sung cột Điểm BSC vào file So_sanh_C11_SM2.xlsx
Áp dụng cho cả 2 sheet: So_sanh_chi_tiet và Thong_ke_theo_don_vi
"""

import pandas as pd
import numpy as np


def tinh_diem_C11_TP1(kq):
    """
    Tính điểm C1.1 Thành phần 1 (30%): Tỷ lệ sửa chữa phiếu chất lượng chủ động
    
    Args:
        kq: Tỷ lệ sửa chữa chủ động (dạng phần trăm, vd: 98 = 98%)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có phiếu cần sửa = tốt
    
    # Chuyển về dạng thập phân (98% -> 0.98)
    kq_decimal = kq / 100 if kq > 1 else kq
    
    if kq_decimal >= 0.99:
        return 5
    elif kq_decimal > 0.96:
        return 1 + 4 * (kq_decimal - 0.96) / 0.03
    else:
        return 1


def add_bsc_score_columns(file_path):
    """
    Thêm cột Điểm BSC vào file So_sanh_C11_SM2.xlsx
    """
    print(f"📖 Đang đọc file: {file_path}")
    
    # Đọc tất cả các sheet
    excel_file = pd.ExcelFile(file_path)
    
    # Dictionary để lưu tất cả các sheet (đã xử lý và chưa xử lý)
    all_sheets = {}
    
    # ========== XỬ LÝ SHEET 1: So_sanh_chi_tiet ==========
    if 'So_sanh_chi_tiet' in excel_file.sheet_names:
        print("\n📊 Xử lý sheet: So_sanh_chi_tiet")
        df_chitiet = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
        
        # Tính điểm BSC cho cột Thô
        print("   ✓ Tính Điểm BSC (Thô)...")
        df_chitiet['Điểm BSC (Thô)'] = df_chitiet['Tỷ lệ % (Thô)'].apply(tinh_diem_C11_TP1)
        
        # Tính điểm BSC cho cột Sau GT
        print("   ✓ Tính Điểm BSC (Sau GT)...")
        df_chitiet['Điểm BSC (Sau GT)'] = df_chitiet['Tỷ lệ % (Sau GT)'].apply(tinh_diem_C11_TP1)
        
        # Tính chênh lệch điểm
        print("   ✓ Tính Chênh lệch Điểm BSC...")
        df_chitiet['Chênh lệch Điểm'] = (
            df_chitiet['Điểm BSC (Sau GT)'] - df_chitiet['Điểm BSC (Thô)']
        ).round(2)
        
        # Làm tròn điểm BSC
        df_chitiet['Điểm BSC (Thô)'] = df_chitiet['Điểm BSC (Thô)'].round(2)
        df_chitiet['Điểm BSC (Sau GT)'] = df_chitiet['Điểm BSC (Sau GT)'].round(2)
        
        # Sắp xếp lại các cột
        cols = df_chitiet.columns.tolist()
        # Tìm vị trí của 'Chênh lệch %'
        if 'Chênh lệch %' in cols:
            idx = cols.index('Chênh lệch %')
            # Chèn các cột điểm BSC sau 'Chênh lệch %'
            new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
            # Thêm các cột còn lại (nếu có)
            remaining = [c for c in cols if c not in new_cols]
            new_cols = new_cols + remaining
            df_chitiet = df_chitiet[new_cols]
        
        all_sheets['So_sanh_chi_tiet'] = df_chitiet
        print(f"   ✅ Đã thêm 3 cột: Điểm BSC (Thô), Điểm BSC (Sau GT), Chênh lệch Điểm")
    
    # ========== XỬ LÝ SHEET 2: Thong_ke_theo_don_vi ==========
    if 'Thong_ke_theo_don_vi' in excel_file.sheet_names:
        print("\n📊 Xử lý sheet: Thong_ke_theo_don_vi")
        df_donvi = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
        
        # Tính điểm BSC cho cột Thô
        print("   ✓ Tính Điểm BSC (Thô)...")
        df_donvi['Điểm BSC (Thô)'] = df_donvi['Tỷ lệ % (Thô)'].apply(tinh_diem_C11_TP1)
        
        # Tính điểm BSC cho cột Sau GT
        print("   ✓ Tính Điểm BSC (Sau GT)...")
        df_donvi['Điểm BSC (Sau GT)'] = df_donvi['Tỷ lệ % (Sau GT)'].apply(tinh_diem_C11_TP1)
        
        # Tính chênh lệch điểm
        print("   ✓ Tính Chênh lệch Điểm BSC...")
        df_donvi['Chênh lệch Điểm'] = (
            df_donvi['Điểm BSC (Sau GT)'] - df_donvi['Điểm BSC (Thô)']
        ).round(2)
        
        # Làm tròn điểm BSC
        df_donvi['Điểm BSC (Thô)'] = df_donvi['Điểm BSC (Thô)'].round(2)
        df_donvi['Điểm BSC (Sau GT)'] = df_donvi['Điểm BSC (Sau GT)'].round(2)
        
        # Sắp xếp lại các cột
        cols = df_donvi.columns.tolist()
        # Tìm vị trí của 'Thay đổi %'
        if 'Thay đổi %' in cols:
            idx = cols.index('Thay đổi %')
            # Chèn các cột điểm BSC sau 'Thay đổi %'
            new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
            # Thêm các cột còn lại (nếu có)
            remaining = [c for c in cols if c not in new_cols]
            new_cols = new_cols + remaining
            df_donvi = df_donvi[new_cols]
        
        all_sheets['Thong_ke_theo_don_vi'] = df_donvi
        print(f"   ✅ Đã thêm 3 cột: Điểm BSC (Thô), Điểm BSC (Sau GT), Chênh lệch Điểm")
    
    # ========== GIỮ NGUYÊN CÁC SHEET KHÁC ==========
    for sheet_name in excel_file.sheet_names:
        if sheet_name not in all_sheets:
            all_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"   • Giữ nguyên sheet: {sheet_name}")
    
    # ========== GHI LẠI FILE ==========
    print(f"\n💾 Đang ghi lại file: {file_path}")
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("✅ Hoàn thành!")
    
    # Thống kê
    print("\n" + "="*60)
    print("📊 THỐNG KÊ:")
    print("="*60)
    
    if 'So_sanh_chi_tiet' in all_sheets:
        df = all_sheets['So_sanh_chi_tiet']
        print(f"\n📌 So_sanh_chi_tiet:")
        print(f"   - Số NVKT: {len(df)}")
        print(f"   - Điểm BSC (Thô) TB: {df['Điểm BSC (Thô)'].mean():.2f}")
        print(f"   - Điểm BSC (Sau GT) TB: {df['Điểm BSC (Sau GT)'].mean():.2f}")
        print(f"   - Chênh lệch Điểm TB: {df['Chênh lệch Điểm'].mean():.2f}")
    
    if 'Thong_ke_theo_don_vi' in all_sheets:
        df = all_sheets['Thong_ke_theo_don_vi']
        print(f"\n📌 Thong_ke_theo_don_vi:")
        print(f"   - Số đơn vị: {len(df)}")
        print(f"   - Điểm BSC (Thô) TB: {df['Điểm BSC (Thô)'].mean():.2f}")
        print(f"   - Điểm BSC (Sau GT) TB: {df['Điểm BSC (Sau GT)'].mean():.2f}")
        print(f"   - Chênh lệch Điểm TB: {df['Chênh lệch Điểm'].mean():.2f}")


if __name__ == "__main__":
    file_path = "downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx"
    add_bsc_score_columns(file_path)
