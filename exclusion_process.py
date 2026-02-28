# -*- coding: utf-8 -*-
"""
Module xử lý giảm trừ phiếu báo hỏng và tạo báo cáo so sánh trước/sau giảm trừ

Chức năng:
- Đọc danh sách BAOHONG_ID cần loại trừ từ file ds_phieu_loai_tru.xlsx
- Tạo báo cáo so sánh trước/sau giảm trừ cho các báo cáo chi tiết C1.1, C1.2
- Xuất kết quả vào thư mục downloads/kq_sau_giam_tru/
"""

import pandas as pd
import os
from datetime import datetime
import re
from pathlib import Path

import kpi_calculator
from kpi_calculator import (
    tinh_diem_kpi_nvkt, 
    tinh_diem_kpi_nvkt_sau_giam_tru,
    tinh_diem_C11_TP1, tinh_diem_C11_TP2,
    tinh_diem_C12_TP1, tinh_diem_C12_TP2,
    tinh_diem_C14, tinh_diem_C15
)
import report_generator


def normalize_id(id_val):
    """
    Chuẩn hóa BAOHONG_ID:
    - Chuyển về string
    - Loại bỏ ký tự lạ _x000d_, _x000D_ (carriage return từ Excel), khoảng trắng, xuống dòng
    - Loại bỏ hậu tố .0 nếu có (do pandas đọc nhầm thành float)
    """
    if pd.isna(id_val):
        return ""
    
    # Chuyển về string và loại bỏ khoảng trắng
    s = str(id_val).strip()
    
    # Loại bỏ _x000d_ và _x000D_ (case-insensitive) - đây là ký tự carriage return từ Excel
    s = re.sub(r'_x000[dD]_', '', s).strip()
    
    # Loại bỏ các ký tự điều khiển khác
    s = re.sub(r'[\r\n\t]', '', s)
    
    # Loại bỏ .0 nếu là số
    if s.endswith('.0'):
        s = s[:-2]
        
    return s


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


def tinh_diem_C11_TP2(kq):
    """
    Tính điểm C1.1 Thành phần 2 (70%): Tỷ lệ sửa chữa báo hỏng đúng quy định (không tính hẹn)
    
    Args:
        kq: Tỷ lệ sửa chữa báo hỏng đúng quy định (dạng phần trăm, vd: 85 = 85%)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có phiếu báo hỏng = tốt
    
    # Chuyển về dạng thập phân (85% -> 0.85)
    kq_decimal = kq / 100 if kq > 1 else kq
    
    if kq_decimal >= 0.85:
        return 5
    elif kq_decimal >= 0.82:
        return 4 + (kq_decimal - 0.82) / 0.03
    elif kq_decimal >= 0.79:
        return 3 + (kq_decimal - 0.79) / 0.03
    elif kq_decimal >= 0.76:
        return 2
    else:
        return 1


def tinh_diem_C12_TP1(kq):
    """
    Tính điểm C1.2 Thành phần 1 (50%): Tỷ lệ thuê bao báo hỏng lặp lại
    LƯU Ý: Càng thấp càng tốt
    
    Args:
        kq: Tỷ lệ báo hỏng lặp lại (dạng phần trăm, vd: 2.5 = 2.5%)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có hỏng lặp lại = tốt
    
    # Chuyển về dạng thập phân (2.5% -> 0.025)
    kq_decimal = kq / 100 if kq > 1 else kq
    
    if kq_decimal <= 0.025:
        return 5
    elif kq_decimal < 0.04:
        return 5 - 4 * (kq_decimal - 0.025) / 0.015
    else:
        return 1


def tinh_diem_C12_TP2(kq):
    """
    Tính điểm C1.2 Thành phần 2 (50%): Tỷ lệ sự cố dịch vụ BRCĐ
    LƯU Ý: Càng thấp càng tốt
    
    Args:
        kq: Tỷ lệ sự cố dịch vụ BRCĐ (dạng phần trăm, vd: 2.0 = 2.0%)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có sự cố = tốt
    
    # Chuyển về dạng thập phân (2.0% -> 0.02)
    kq_decimal = kq / 100 if kq > 1 else kq
    
    if kq_decimal <= 0.02:
        return 5
    elif kq_decimal < 0.03:
        return 5 - 4 * (kq_decimal - 0.02) / 0.01
    else:
        return 1


def add_bsc_scores_to_c12_tp2(file_path):
    """
    Thêm cột Điểm BSC vào file SM4-C12-ti-le-su-co-dv-brcd.xlsx
    Áp dụng cho 2 sheet: So_sanh_chi_tiet và Thong_ke_theo_don_vi
    
    Args:
        file_path: Đường dẫn đầy đủ đến file SM4-C12-ti-le-su-co-dv-brcd.xlsx
    """
    try:
        print("\n" + "-"*80)
        print("BỔ SUNG CỘT ĐIỂM BSC VÀO BÁO CÁO C1.2 TP2 (SM4-C12 BRCĐ)")
        print("-"*80)
        
        if not os.path.exists(file_path):
            print(f"⚠️ Không tìm thấy file: {file_path}")
            return
        
        # Đọc tất cả các sheet
        excel_file = pd.ExcelFile(file_path)
        all_sheets = {}
        
        # XỬ LÝ SHEET 1: So_sanh_chi_tiet
        if 'So_sanh_chi_tiet' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: So_sanh_chi_tiet")
            df = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
            
            # Tính điểm BSC - LƯU Ý: Dùng tỷ lệ báo hỏng
            df['Điểm BSC (Thô)'] = df['Tỷ lệ báo hỏng (%) (Thô)'].apply(tinh_diem_C12_TP2).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ báo hỏng (%) (Sau GT)'].apply(tinh_diem_C12_TP2).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Chênh lệch %' in cols:
                idx = cols.index('Chênh lệch %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['So_sanh_chi_tiet'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # XỬ LÝ SHEET 2: Thong_ke_theo_don_vi
        if 'Thong_ke_theo_don_vi' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: Thong_ke_theo_don_vi")
            df = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
            
            # Tính điểm BSC - LƯU Ý: Dùng tỷ lệ báo hỏng
            df['Điểm BSC (Thô)'] = df['Tỷ lệ báo hỏng % (Thô)'].apply(tinh_diem_C12_TP2).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ báo hỏng % (Sau GT)'].apply(tinh_diem_C12_TP2).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Chênh lệch %' in cols:
                idx = cols.index('Chênh lệch %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['Thong_ke_theo_don_vi'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # Giữ nguyên các sheet khác
        for sheet_name in excel_file.sheet_names:
            if sheet_name not in all_sheets:
                all_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Ghi lại file
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("✅ Đã bổ sung điểm BSC vào báo cáo C1.2 TP2 (SM4-C12 BRCĐ)")
        
    except Exception as e:
        print(f"⚠️ Lỗi khi bổ sung điểm BSC: {e}")
        import traceback
        traceback.print_exc()


def tinh_diem_C14(kq):
    """
    Tính điểm C1.4: Độ hài lòng của khách hàng sau lắp đặt và sửa chữa
    
    Args:
        kq: Độ hài lòng khách hàng (dạng phần trăm, vd: 99.5 = 99.5%)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return pd.NA  # Không có dữ liệu
    
    # Chuyển về dạng thập phân (99.5% -> 0.995)
    kq_decimal = kq / 100 if kq > 1 else kq
    
    if kq_decimal >= 0.995:
        return 5
    elif kq_decimal > 0.95:
        return 1 + 4 * (kq_decimal - 0.95) / 0.045
    else:
        return 1


def tinh_diem_C15(kq):
    """
    Tính điểm C1.5: Tỉ lệ thiết lập dịch vụ đạt thời gian quy định
    
    Args:
        kq: Tỉ lệ thiết lập dịch vụ đạt (dạng phần trăm, vd: 99.5 = 99.5%)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return pd.NA  # Không có dữ liệu
    
    # Chuyển về dạng thập phân (99.5% -> 0.995)
    kq_decimal = kq / 100 if kq > 1 else kq
    
    if kq_decimal >= 0.995:
        return 5
    elif kq_decimal > 0.895:
        return 1 + 4 * (kq_decimal - 0.895) / 0.10
    else:
        return 1


def add_bsc_scores_to_c15(file_path):
    """
    Thêm cột Điểm BSC vào file So_sanh_C15.xlsx
    Áp dụng cho 2 sheet: So_sanh_chi_tiet và Thong_ke_theo_don_vi
    
    Args:
        file_path: Đường dẫn đầy đủ đến file So_sanh_C15.xlsx
    """
    try:
        print("\n" + "-"*80)
        print("BỔ SUNG CỘT ĐIỂM BSC VÀO BÁO CÁO C1.5")
        print("-"*80)
        
        if not os.path.exists(file_path):
            print(f"⚠️ Không tìm thấy file: {file_path}")
            return
        
        # Đọc tất cả các sheet
        excel_file = pd.ExcelFile(file_path)
        all_sheets = {}
        
        # XỬ LÝ SHEET 1: So_sanh_chi_tiet
        if 'So_sanh_chi_tiet' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: So_sanh_chi_tiet")
            df = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
            
            # Tính điểm BSC
            df['Điểm BSC (Thô)'] = df['Tỷ lệ đạt (%) (Thô)'].apply(tinh_diem_C15).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ đạt (%) (Sau GT)'].apply(tinh_diem_C15).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Chênh lệch %' in cols:
                idx = cols.index('Chênh lệch %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['So_sanh_chi_tiet'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # XỬ LÝ SHEET 2: Thong_ke_theo_don_vi
        if 'Thong_ke_theo_don_vi' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: Thong_ke_theo_don_vi")
            df = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
            
            # Tính điểm BSC
            df['Điểm BSC (Thô)'] = df['Tỷ lệ đạt % (Thô)'].apply(tinh_diem_C15).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ đạt % (Sau GT)'].apply(tinh_diem_C15).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Thay đổi %' in cols:
                idx = cols.index('Thay đổi %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['Thong_ke_theo_don_vi'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # Giữ nguyên các sheet khác
        for sheet_name in excel_file.sheet_names:
            if sheet_name not in all_sheets:
                all_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Ghi lại file
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("✅ Đã bổ sung điểm BSC vào báo cáo C1.5")
        
    except Exception as e:
        print(f"⚠️ Lỗi khi bổ sung điểm BSC: {e}")
        import traceback
        traceback.print_exc()


def add_bsc_scores_to_c14(file_path):
    """
    Thêm cột Điểm BSC vào file So_sanh_C14.xlsx
    Áp dụng cho 2 sheet: So_sanh_chi_tiet và Thong_ke_theo_don_vi
    
    Args:
        file_path: Đường dẫn đầy đủ đến file So_sanh_C14.xlsx
    """
    try:
        print("\n" + "-"*80)
        print("BỔ SUNG CỘT ĐIỂM BSC VÀO BÁO CÁO C1.4")
        print("-"*80)
        
        if not os.path.exists(file_path):
            print(f"⚠️ Không tìm thấy file: {file_path}")
            return
        
        # Đọc tất cả các sheet
        excel_file = pd.ExcelFile(file_path)
        all_sheets = {}
        
        # XỬ LÝ SHEET 1: So_sanh_chi_tiet
        if 'So_sanh_chi_tiet' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: So_sanh_chi_tiet")
            df = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
            
            # Tính điểm BSC - LƯU Ý: Dùng tỷ lệ hài lòng
            df['Điểm BSC (Thô)'] = df['Tỷ lệ HL (%) (Thô)'].apply(tinh_diem_C14).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ HL (%) (Sau GT)'].apply(tinh_diem_C14).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Chênh lệch %' in cols:
                idx = cols.index('Chênh lệch %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['So_sanh_chi_tiet'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # XỬ LÝ SHEET 2: Thong_ke_theo_don_vi
        if 'Thong_ke_theo_don_vi' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: Thong_ke_theo_don_vi")
            df = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
            
            # Tính điểm BSC - LƯU Ý: Dùng tỷ lệ hài lòng
            df['Điểm BSC (Thô)'] = df['Tỷ lệ HL % (Thô)'].apply(tinh_diem_C14).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ HL % (Sau GT)'].apply(tinh_diem_C14).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Thay đổi HL %' in cols:
                idx = cols.index('Thay đổi HL %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['Thong_ke_theo_don_vi'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # Giữ nguyên các sheet khác
        for sheet_name in excel_file.sheet_names:
            if sheet_name not in all_sheets:
                all_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Ghi lại file
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("✅ Đã bổ sung điểm BSC vào báo cáo C1.4")
        
    except Exception as e:
        print(f"⚠️ Lỗi khi bổ sung điểm BSC: {e}")
        import traceback
        traceback.print_exc()


def add_bsc_scores_to_c12_sm1(file_path):
    """
    Thêm cột Điểm BSC vào file So_sanh_C12_SM1.xlsx
    Áp dụng cho 2 sheet: So_sanh_chi_tiet và Thong_ke_theo_don_vi
    
    Args:
        file_path: Đường dẫn đầy đủ đến file So_sanh_C12_SM1.xlsx
    """
    try:
        print("\n" + "-"*80)
        print("BỔ SUNG CỘT ĐIỂM BSC VÀO BÁO CÁO C1.2 SM1")
        print("-"*80)
        
        if not os.path.exists(file_path):
            print(f"⚠️ Không tìm thấy file: {file_path}")
            return
        
        # Đọc tất cả các sheet
        excel_file = pd.ExcelFile(file_path)
        all_sheets = {}
        
        # XỬ LÝ SHEET 1: So_sanh_chi_tiet
        if 'So_sanh_chi_tiet' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: So_sanh_chi_tiet")
            df = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
            
            # Tính điểm BSC - LƯU Ý: Dùng tỷ lệ HLL
            df['Điểm BSC (Thô)'] = df['Tỷ lệ HLL % (Thô)'].apply(tinh_diem_C12_TP1).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ HLL % (Sau GT)'].apply(tinh_diem_C12_TP1).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Chênh lệch %' in cols:
                idx = cols.index('Chênh lệch %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['So_sanh_chi_tiet'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # XỬ LÝ SHEET 2: Thong_ke_theo_don_vi
        if 'Thong_ke_theo_don_vi' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: Thong_ke_theo_don_vi")
            df = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
            
            # Tính điểm BSC - LƯU Ý: Dùng tỷ lệ HLL
            df['Điểm BSC (Thô)'] = df['Tỷ lệ HLL % (Thô)'].apply(tinh_diem_C12_TP1).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ HLL % (Sau GT)'].apply(tinh_diem_C12_TP1).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Thay đổi %' in cols:
                idx = cols.index('Thay đổi %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['Thong_ke_theo_don_vi'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # Giữ nguyên các sheet khác
        for sheet_name in excel_file.sheet_names:
            if sheet_name not in all_sheets:
                all_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Ghi lại file
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("✅ Đã bổ sung điểm BSC vào báo cáo C1.2 SM1")
        
    except Exception as e:
        print(f"⚠️ Lỗi khi bổ sung điểm BSC: {e}")
        import traceback
        traceback.print_exc()


def add_bsc_scores_to_c11_sm4(file_path):
    """
    Thêm cột Điểm BSC vào file So_sanh_C11_SM4.xlsx
    Áp dụng cho 2 sheet: So_sanh_chi_tiet và Thong_ke_theo_don_vi
    
    Args:
        file_path: Đường dẫn đầy đủ đến file So_sanh_C11_SM4.xlsx
    """
    try:
        print("\n" + "-"*80)
        print("BỔ SUNG CỘT ĐIỂM BSC VÀO BÁO CÁO C1.1 SM4")
        print("-"*80)
        
        if not os.path.exists(file_path):
            print(f"⚠️ Không tìm thấy file: {file_path}")
            return
        
        # Đọc tất cả các sheet
        excel_file = pd.ExcelFile(file_path)
        all_sheets = {}
        
        # XỬ LÝ SHEET 1: So_sanh_chi_tiet
        if 'So_sanh_chi_tiet' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: So_sanh_chi_tiet")
            df = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
            
            # Tính điểm BSC
            df['Điểm BSC (Thô)'] = df['Tỷ lệ % (Thô)'].apply(tinh_diem_C11_TP2).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ % (Sau GT)'].apply(tinh_diem_C11_TP2).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Chênh lệch %' in cols:
                idx = cols.index('Chênh lệch %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['So_sanh_chi_tiet'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # XỬ LÝ SHEET 2: Thong_ke_theo_don_vi
        if 'Thong_ke_theo_don_vi' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: Thong_ke_theo_don_vi")
            df = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
            
            # Tính điểm BSC
            df['Điểm BSC (Thô)'] = df['Tỷ lệ % (Thô)'].apply(tinh_diem_C11_TP2).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ % (Sau GT)'].apply(tinh_diem_C11_TP2).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Thay đổi %' in cols:
                idx = cols.index('Thay đổi %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['Thong_ke_theo_don_vi'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # Giữ nguyên các sheet khác
        for sheet_name in excel_file.sheet_names:
            if sheet_name not in all_sheets:
                all_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Ghi lại file
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("✅ Đã bổ sung điểm BSC vào báo cáo C1.1 SM4")
        
    except Exception as e:
        print(f"⚠️ Lỗi khi bổ sung điểm BSC: {e}")
        import traceback
        traceback.print_exc()


def add_bsc_scores_to_c11_sm2(file_path):
    """
    Thêm cột Điểm BSC vào file So_sanh_C11_SM2.xlsx
    Áp dụng cho 2 sheet: So_sanh_chi_tiet và Thong_ke_theo_don_vi
    
    Args:
        file_path: Đường dẫn đầy đủ đến file So_sanh_C11_SM2.xlsx
    """
    try:
        print("\n" + "-"*80)
        print("BỔ SUNG CỘT ĐIỂM BSC VÀO BÁO CÁO C1.1 SM2")
        print("-"*80)
        
        if not os.path.exists(file_path):
            print(f"⚠️ Không tìm thấy file: {file_path}")
            return
        
        # Đọc tất cả các sheet
        excel_file = pd.ExcelFile(file_path)
        all_sheets = {}
        
        # XỬ LÝ SHEET 1: So_sanh_chi_tiet
        if 'So_sanh_chi_tiet' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: So_sanh_chi_tiet")
            df = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
            
            # Tính điểm BSC
            df['Điểm BSC (Thô)'] = df['Tỷ lệ % (Thô)'].apply(tinh_diem_C11_TP1).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ % (Sau GT)'].apply(tinh_diem_C11_TP1).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Chênh lệch %' in cols:
                idx = cols.index('Chênh lệch %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['So_sanh_chi_tiet'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # XỬ LÝ SHEET 2: Thong_ke_theo_don_vi
        if 'Thong_ke_theo_don_vi' in excel_file.sheet_names:
            print("✓ Đang xử lý sheet: Thong_ke_theo_don_vi")
            df = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
            
            # Tính điểm BSC
            df['Điểm BSC (Thô)'] = df['Tỷ lệ % (Thô)'].apply(tinh_diem_C11_TP1).round(2)
            df['Điểm BSC (Sau GT)'] = df['Tỷ lệ % (Sau GT)'].apply(tinh_diem_C11_TP1).round(2)
            df['Chênh lệch Điểm'] = (df['Điểm BSC (Sau GT)'] - df['Điểm BSC (Thô)']).round(2)
            
            # Sắp xếp lại cột
            cols = df.columns.tolist()
            if 'Thay đổi %' in cols:
                idx = cols.index('Thay đổi %')
                new_cols = cols[:idx+1] + ['Điểm BSC (Thô)', 'Điểm BSC (Sau GT)', 'Chênh lệch Điểm']
                remaining = [c for c in cols if c not in new_cols]
                df = df[new_cols + remaining]
            
            all_sheets['Thong_ke_theo_don_vi'] = df
            print(f"  → Đã thêm 3 cột điểm BSC (TB: {df['Điểm BSC (Sau GT)'].mean():.2f})")
        
        # Giữ nguyên các sheet khác
        for sheet_name in excel_file.sheet_names:
            if sheet_name not in all_sheets:
                all_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Ghi lại file
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("✅ Đã bổ sung điểm BSC vào báo cáo C1.1 SM2")
        
    except Exception as e:
        print(f"⚠️ Lỗi khi bổ sung điểm BSC: {e}")
        import traceback
        traceback.print_exc()


def load_exclusion_list(exclusion_file="du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx"):
    """
    Đọc danh sách BAOHONG_ID cần loại trừ từ file Excel
    
    Args:
        exclusion_file: Đường dẫn file chứa danh sách loại trừ
        
    Returns:
        set: Tập hợp các BAOHONG_ID cần loại trừ (dạng string)
    """
    try:
        if not os.path.exists(exclusion_file):
            print(f"⚠️ Không tìm thấy file loại trừ: {exclusion_file}")
            return set()
        
        df = pd.read_excel(exclusion_file)
        
        if 'BAOHONG_ID' not in df.columns:
            print(f"⚠️ Không tìm thấy cột 'BAOHONG_ID' trong file {exclusion_file}")
            return set()
        
        # Chuẩn hóa tất cả ID
        exclusion_ids = {normalize_id(idx) for idx in df['BAOHONG_ID'].tolist() if pd.notna(idx)}
        exclusion_ids.discard("") # Loại bỏ chuỗi rỗng nếu có
        
        print(f"✅ Đã đọc {len(exclusion_ids)} mã BAOHONG_ID sau khi chuẩn hóa")
        return exclusion_ids
        
    except Exception as e:
        print(f"❌ Lỗi khi đọc file loại trừ: {e}")
        return set()


def extract_nvkt_name(ten_kv):
    """
    Trích xuất tên NVKT từ cột TEN_KV
    Ví dụ:
    - Sơn Lộc 1 - Nguyễn Thành Sơn -> Nguyễn Thành Sơn
    - VNM3-Khuất Anh Chiến( VXN) -> Khuất Anh Chiến
    """
    if pd.isna(ten_kv):
        return None
    
    ten_kv = str(ten_kv).strip()
    
    # Trường hợp có dấu "-"
    if '-' in ten_kv:
        parts = ten_kv.split('-')
        nvkt_name = parts[-1].strip()
    else:
        nvkt_name = ten_kv
    
    # Loại bỏ phần trong ngoặc đơn
    if '(' in nvkt_name:
        nvkt_name = nvkt_name.split('(')[0].strip()
    
    return nvkt_name


def calculate_statistics(df, has_ten_doi=True, dat_column='DAT_TT_KO_HEN', dat_value=1):
    """
    Tính toán thống kê theo TEN_DOI và NVKT
    
    Args:
        df: DataFrame đã được xử lý với cột NVKT
        has_ten_doi: Có cột TEN_DOI hay không
        dat_column: Tên cột để xác định phiếu đạt
        dat_value: Giá trị để xác định phiếu đạt
        
    Returns:
        DataFrame với thống kê
    """
    report_data = []
    
    if has_ten_doi:
        group_columns = ['TEN_DOI', 'NVKT']
    else:
        group_columns = ['NVKT']
    
    for group_key, group_df in df.groupby(group_columns):
        tong_phieu = len(group_df)
        so_phieu_dat = len(group_df[group_df[dat_column] == dat_value])
        ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0
        
        if has_ten_doi:
            ten_doi, nvkt = group_key
            report_data.append({
                'TEN_DOI': ten_doi,
                'NVKT': nvkt,
                'Tổng phiếu': tong_phieu,
                'Số phiếu đạt': so_phieu_dat,
                'Tỷ lệ %': ty_le
            })
        else:
            report_data.append({
                'NVKT': group_key,
                'Tổng phiếu': tong_phieu,
                'Số phiếu đạt': so_phieu_dat,
                'Tỷ lệ %': ty_le
            })
    
    return pd.DataFrame(report_data)


def calculate_statistics_keep_denominator(df_before_stats, df_raw, exclusion_ids,
                                          has_ten_doi=True, dat_column='DAT_TT_KO_HEN', dat_value=1):
    """
    Tính toán thống kê SAU giảm trừ với LOGIC:
    - Mẫu số (Tổng phiếu): GIỮ NGUYÊN từ df_before_stats
    - Tử số (Phiếu đạt): TĂNG thêm số phiếu KHÔNG ĐẠT bị loại trừ (chuyển thành đạt)

    Args:
        df_before_stats: DataFrame thống kê TRƯỚC giảm trừ (từ calculate_statistics)
        df_raw: DataFrame dữ liệu thô đầy đủ
        exclusion_ids: Set các BAOHONG_ID cần loại trừ
        has_ten_doi: Có cột TEN_DOI hay không
        dat_column: Tên cột để xác định phiếu đạt
        dat_value: Giá trị để xác định phiếu đạt

    Returns:
        DataFrame với thống kê SAU giảm trừ (mẫu giữ nguyên, tử tăng)
    """
    report_data = []

    if has_ten_doi:
        group_columns = ['TEN_DOI', 'NVKT']
    else:
        group_columns = ['NVKT']

    # Xác định các phiếu bị loại trừ
    df_raw['BAOHONG_ID_STR'] = df_raw['BAOHONG_ID'].apply(normalize_id)
    df_excluded_items = df_raw[df_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()

    for group_key, group_df_before_row in df_before_stats.groupby(group_columns):
        if len(group_df_before_row) == 0:
            continue

        # Lấy thống kê TRƯỚC từ df_before_stats
        before_row = group_df_before_row.iloc[0]
        tong_phieu_truoc = before_row['Tổng phiếu']
        so_phieu_dat_truoc = before_row['Số phiếu đạt']

        # Đếm số phiếu KHÔNG ĐẠT bị loại trừ cho nhóm này (sẽ chuyển thành đạt)
        if has_ten_doi:
            ten_doi, nvkt = group_key
            excluded_not_dat_count = len(df_excluded_items[
                (df_excluded_items['TEN_DOI'] == ten_doi) &
                (df_excluded_items['NVKT'] == nvkt) &
                (df_excluded_items[dat_column] != dat_value)
            ])
        else:
            nvkt = group_key
            excluded_not_dat_count = len(df_excluded_items[
                (df_excluded_items['NVKT'] == nvkt) &
                (df_excluded_items[dat_column] != dat_value)
            ])

        # Tính toán SAU giảm trừ
        tong_phieu_sau = tong_phieu_truoc  # GIỮ NGUYÊN MẪU SỐ
        so_phieu_dat_sau = so_phieu_dat_truoc + excluded_not_dat_count  # TĂNG TỬ SỐ (chuyển không đạt -> đạt)
        so_phieu_dat_sau = min(so_phieu_dat_sau, tong_phieu_sau)  # Đảm bảo không vượt quá tổng

        ty_le = round((so_phieu_dat_sau / tong_phieu_sau * 100), 2) if tong_phieu_sau > 0 else 0

        if has_ten_doi:
            report_data.append({
                'TEN_DOI': ten_doi,
                'NVKT': nvkt,
                'Tổng phiếu': tong_phieu_sau,
                'Số phiếu đạt': so_phieu_dat_sau,
                'Tỷ lệ %': ty_le
            })
        else:
            report_data.append({
                'NVKT': nvkt,
                'Tổng phiếu': tong_phieu_sau,
                'Số phiếu đạt': so_phieu_dat_sau,
                'Tỷ lệ %': ty_le
            })

    return pd.DataFrame(report_data)


def calculate_unit_stats(df_before, df_after, ten_doi_col='TEN_DOI', 
                         tong_col='Tổng phiếu', dat_col='Số phiếu đạt'):
    """
    Tính thống kê theo từng đơn vị (Tổ) trước và sau giảm trừ
    
    Args:
        df_before: DataFrame thống kê TRƯỚC giảm trừ (nhóm theo NVKT)
        df_after: DataFrame thống kê SAU giảm trừ (nhóm theo NVKT)
        ten_doi_col: Tên cột đơn vị
        tong_col: Tên cột tổng phiếu
        dat_col: Tên cột số phiếu đạt
        
    Returns:
        DataFrame với thống kê theo từng đơn vị + dòng Tổng TTVT
    """
    # Mapping tên đội ngắn
    TEAM_SHORT_NAMES = {
        'Tổ Kỹ thuật địa bàn Phúc Thọ': 'Phúc Thọ',
        'Tổ Kỹ thuật địa bàn Quảng Oai': 'Quảng Oai',
        'Tổ Kỹ thuật địa bàn Suối Hai': 'Suối Hai',
        'Tổ Kỹ thuật địa bàn Sơn Tây': 'Sơn Tây',
    }
    
    unit_stats = []
    
    # Nếu có cột TEN_DOI, nhóm theo đơn vị
    if ten_doi_col in df_before.columns:
        # Nhóm theo đơn vị TRƯỚC giảm trừ
        unit_before = df_before.groupby(ten_doi_col).agg({
            tong_col: 'sum',
            dat_col: 'sum'
        }).reset_index()
        
        # Nhóm theo đơn vị SAU giảm trừ
        unit_after = df_after.groupby(ten_doi_col).agg({
            tong_col: 'sum',
            dat_col: 'sum'
        }).reset_index()
        
        # Merge để có cả 2
        unit_merged = pd.merge(
            unit_before, unit_after,
            on=ten_doi_col, how='outer',
            suffixes=(' (Thô)', ' (Sau GT)')
        )
        
        for _, row in unit_merged.iterrows():
            ten_doi = row[ten_doi_col]
            short_name = TEAM_SHORT_NAMES.get(ten_doi, ten_doi)
            
            tong_tho = row.get(f'{tong_col} (Thô)', 0) or 0
            tong_sau = row.get(f'{tong_col} (Sau GT)', 0) or 0
            dat_tho = row.get(f'{dat_col} (Thô)', 0) or 0
            dat_sau = row.get(f'{dat_col} (Sau GT)', 0) or 0
            
            tyle_tho = round((dat_tho / tong_tho * 100), 2) if tong_tho > 0 else 0
            tyle_sau = round((dat_sau / tong_sau * 100), 2) if tong_sau > 0 else 0
            
            unit_stats.append({
                'Đơn vị': short_name,
                'Tổng phiếu (Thô)': int(tong_tho),
                'Phiếu loại trừ': int(tong_tho - tong_sau),
                'Tổng phiếu (Sau GT)': int(tong_sau),
                'Phiếu đạt (Thô)': int(dat_tho),
                'Phiếu đạt (Sau GT)': int(dat_sau),
                'Tỷ lệ % (Thô)': tyle_tho,
                'Tỷ lệ % (Sau GT)': tyle_sau,
                'Thay đổi %': round(tyle_sau - tyle_tho, 2)
            })
    
    # Thêm dòng TỔNG TTVT
    tong_tho_all = df_before[tong_col].sum()
    tong_sau_all = df_after[tong_col].sum()
    dat_tho_all = df_before[dat_col].sum()
    dat_sau_all = df_after[dat_col].sum()
    
    tyle_tho_all = round((dat_tho_all / tong_tho_all * 100), 2) if tong_tho_all > 0 else 0
    tyle_sau_all = round((dat_sau_all / tong_sau_all * 100), 2) if tong_sau_all > 0 else 0
    
    unit_stats.append({
        'Đơn vị': 'TTVT Sơn Tây',
        'Tổng phiếu (Thô)': int(tong_tho_all),
        'Phiếu loại trừ': int(tong_tho_all - tong_sau_all),
        'Tổng phiếu (Sau GT)': int(tong_sau_all),
        'Phiếu đạt (Thô)': int(dat_tho_all),
        'Phiếu đạt (Sau GT)': int(dat_sau_all),
        'Tỷ lệ % (Thô)': tyle_tho_all,
        'Tỷ lệ % (Sau GT)': tyle_sau_all,
        'Thay đổi %': round(tyle_sau_all - tyle_tho_all, 2)
    })
    
    return pd.DataFrame(unit_stats)


def create_c11_comparison_report(exclusion_ids, output_dir):
    """
    Tạo báo cáo so sánh C1.1 (SM4-C11) trước/sau giảm trừ
    
    Args:
        exclusion_ids: Set các BAOHONG_ID cần loại trừ
        output_dir: Thư mục xuất kết quả
        
    Returns:
        dict: Thống kê tổng hợp
    """
    try:
        print("\n" + "="*80)
        print("TẠO BÁO CÁO SO SÁNH C1.1 (SM4-C11) TRƯỚC/SAU GIẢM TRỪ")
        print("="*80)
        
        input_file = os.path.join("downloads", "baocao_hanoi", "SM4-C11.xlsx")
        
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return None
        
        # Đọc dữ liệu thô
        df_raw = pd.read_excel(input_file, sheet_name='Sheet1')
        print(f"✅ Đã đọc file, tổng số dòng thô: {len(df_raw)}")
        
        # Kiểm tra cột cần thiết
        if 'BAOHONG_ID' not in df_raw.columns:
            print("❌ Không tìm thấy cột BAOHONG_ID")
            return None
        
        if 'TEN_KV' not in df_raw.columns or 'DAT_TT_KO_HEN' not in df_raw.columns:
            print("❌ Không tìm thấy cột TEN_KV hoặc DAT_TT_KO_HEN")
            return None
        
        has_ten_doi = 'TEN_DOI' in df_raw.columns
        
        # Chuẩn hóa cột NVKT
        df_raw['NVKT'] = df_raw['TEN_KV'].apply(extract_nvkt_name)
        df_raw = df_raw[df_raw['NVKT'].notna()].copy()
        
        # Lọc dữ liệu sau giảm trừ
        df_raw['BAOHONG_ID_STR'] = df_raw['BAOHONG_ID'].apply(normalize_id)
        df_excluded = df_raw[~df_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        
        num_excluded = len(df_raw) - len(df_excluded)
        print(f"✅ Đã loại trừ {num_excluded} phiếu, còn lại {len(df_excluded)} phiếu")
        
        # Tính thống kê TRƯỚC giảm trừ
        print("\n✓ Đang tính thống kê TRƯỚC giảm trừ...")
        df_stats_before = calculate_statistics(df_raw, has_ten_doi)

        # Tính thống kê SAU giảm trừ - LOGIC MỚI: GIỮ NGUYÊN MẪU SỐ
        print("✓ Đang tính thống kê SAU giảm trừ (giữ nguyên mẫu số)...")
        df_stats_after = calculate_statistics_keep_denominator(
            df_stats_before, df_raw, exclusion_ids, has_ten_doi
        )
        
        # Merge kết quả để so sánh
        print("✓ Đang tạo báo cáo so sánh...")
        
        if has_ten_doi:
            merge_columns = ['TEN_DOI', 'NVKT']
        else:
            merge_columns = ['NVKT']
        
        df_comparison = pd.merge(
            df_stats_before,
            df_stats_after,
            on=merge_columns,
            how='outer',
            suffixes=(' (Thô)', ' (Sau GT)')
        )
        
        # Tính chênh lệch tỷ lệ
        df_comparison['Chênh lệch %'] = (
            df_comparison['Tỷ lệ % (Sau GT)'].fillna(0) - 
            df_comparison['Tỷ lệ % (Thô)'].fillna(0)
        ).round(2)
        
        # Sắp xếp cột theo thứ tự mong muốn
        if has_ten_doi:
            column_order = [
                'TEN_DOI', 'NVKT',
                'Tổng phiếu (Thô)', 'Tổng phiếu (Sau GT)',
                'Số phiếu đạt (Thô)', 'Số phiếu đạt (Sau GT)',
                'Tỷ lệ % (Thô)', 'Tỷ lệ % (Sau GT)',
                'Chênh lệch %'
            ]
        else:
            column_order = [
                'NVKT',
                'Tổng phiếu (Thô)', 'Tổng phiếu (Sau GT)',
                'Số phiếu đạt (Thô)', 'Số phiếu đạt (Sau GT)',
                'Tỷ lệ % (Thô)', 'Tỷ lệ % (Sau GT)',
                'Chênh lệch %'
            ]
        
        df_comparison = df_comparison[column_order]
        
        # Sắp xếp theo TEN_DOI và NVKT
        if has_ten_doi:
            df_comparison = df_comparison.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            df_comparison = df_comparison.sort_values('NVKT').reset_index(drop=True)
        
        # Tạo sheet tổng hợp
        tong_before = df_stats_before['Tổng phiếu'].sum()
        tong_after = df_stats_after['Tổng phiếu'].sum()
        dat_before = df_stats_before['Số phiếu đạt'].sum()
        dat_after = df_stats_after['Số phiếu đạt'].sum()
        tyle_before = round((dat_before / tong_before * 100), 2) if tong_before > 0 else 0
        tyle_after = round((dat_after / tong_after * 100), 2) if tong_after > 0 else 0
        
        df_tongke = pd.DataFrame([{
            'Chỉ tiêu': 'C1.1 - Tỷ lệ phiếu sửa chữa BRCD đúng quy định (không tính hẹn)',
            'Tổng phiếu (Thô)': tong_before,
            'Phiếu loại trừ': num_excluded,
            'Tổng phiếu (Sau GT)': tong_after,
            'Phiếu đạt (Thô)': dat_before,
            'Phiếu đạt (Sau GT)': dat_after,
            'Tỷ lệ % (Thô)': tyle_before,
            'Tỷ lệ % (Sau GT)': tyle_after,
            'Thay đổi %': round(tyle_after - tyle_before, 2)
        }])
        
        # Lấy danh sách phiếu bị loại trừ với chi tiết
        df_loai_tru = df_raw[df_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        # Chỉ giữ một số cột quan trọng
        cols_to_keep = ['BAOHONG_ID', 'MA_TB', 'TEN_TB', 'TEN_KV', 'TEN_DOI', 'NGAY_BAO_HONG', 'NGAY_NGHIEM_THU', 'DAT_TT_KO_HEN']
        cols_available = [c for c in cols_to_keep if c in df_loai_tru.columns]
        df_loai_tru = df_loai_tru[cols_available]
        
        # Tính thống kê theo đơn vị (Tổ)
        df_unit_stats = calculate_unit_stats(df_stats_before, df_stats_after)
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C11_SM4.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_unit_stats.to_excel(writer, sheet_name='Thong_ke_theo_don_vi', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"✅ Đã tạo báo cáo so sánh C1.1 (SM4-C11)")
        print(f"   - Tổng phiếu thô: {tong_before}")
        print(f"   - Phiếu loại trừ: {num_excluded}")
        print(f"   - Tổng phiếu sau GT: {tong_after}")
        print(f"   - Tỷ lệ thô: {tyle_before}% -> Sau GT: {tyle_after}%")
        
        # Bổ sung cột Điểm BSC vào file vừa tạo
        add_bsc_scores_to_c11_sm4(output_file)
        
        return {
            'chi_tieu': 'C1.1 SM4',
            'tong_tho': tong_before,
            'loai_tru': num_excluded,
            'tong_sau_gt': tong_after,
            'tyle_tho': tyle_before,
            'tyle_sau_gt': tyle_after
        }
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo C1.1: {e}")
        import traceback
        traceback.print_exc()
        return None


def create_c11_sm2_comparison_report(exclusion_ids, output_dir):
    """
    Tạo báo cáo so sánh C1.1 SM2 trước/sau giảm trừ
    
    Args:
        exclusion_ids: Set các BAOHONG_ID cần loại trừ
        output_dir: Thư mục xuất kết quả
        
    Returns:
        dict: Thống kê tổng hợp
    """
    try:
        print("\n" + "="*80)
        print("TẠO BÁO CÁO SO SÁNH C1.1 SM2 TRƯỚC/SAU GIẢM TRỪ")
        print("="*80)
        
        input_file = os.path.join("downloads", "baocao_hanoi", "SM2-C11.xlsx")
        
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return None
        
        # Đọc dữ liệu thô
        df_raw = pd.read_excel(input_file, sheet_name='Sheet1')
        print(f"✅ Đã đọc file, tổng số dòng thô: {len(df_raw)}")
        
        # Kiểm tra cột cần thiết
        if 'BAOHONG_ID' not in df_raw.columns:
            print("❌ Không tìm thấy cột BAOHONG_ID")
            return None
        
        if 'TEN_KV' not in df_raw.columns:
            print("❌ Không tìm thấy cột TEN_KV")
            return None
        
        has_ten_doi = 'TEN_DOI' in df_raw.columns
        
        # Chuẩn hóa cột NVKT
        df_raw['NVKT'] = df_raw['TEN_KV'].apply(extract_nvkt_name)
        df_raw = df_raw[df_raw['NVKT'].notna()].copy()
        
        # Tính cột TG (thời gian xử lý) nếu chưa có
        if 'TG' not in df_raw.columns:
            if 'NGAY_NGHIEM_THU' in df_raw.columns and 'NGAY_BAO_HONG' in df_raw.columns:
                df_raw['NGAY_NGHIEM_THU'] = pd.to_datetime(df_raw['NGAY_NGHIEM_THU'], errors='coerce')
                df_raw['NGAY_BAO_HONG'] = pd.to_datetime(df_raw['NGAY_BAO_HONG'], errors='coerce')
                df_raw['TG'] = (df_raw['NGAY_NGHIEM_THU'] - df_raw['NGAY_BAO_HONG']).dt.total_seconds() / 3600
        
        # Lọc bỏ các phiếu có TG = NULL (giống logic trong c1_process.py)
        print(f"\n✓ Số dòng ban đầu: {len(df_raw)}")
        df_raw = df_raw[df_raw['TG'].notna()].copy()
        print(f"✓ Số dòng sau khi loại bỏ TG = NULL: {len(df_raw)}")
        
        # Xác định phiếu đạt: TG <= 72 giờ
        df_raw['PHIEU_DAT'] = (df_raw['TG'] <= 72).astype(int)
        
        # Lọc dữ liệu sau giảm trừ
        df_raw['BAOHONG_ID_STR'] = df_raw['BAOHONG_ID'].apply(normalize_id)
        df_excluded = df_raw[~df_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        
        num_excluded = len(df_raw) - len(df_excluded)
        print(f"✅ Đã loại trừ {num_excluded} phiếu, còn lại {len(df_excluded)} phiếu")
        
        # Tính thống kê TRƯỚC giảm trừ
        print("\n✓ Đang tính thống kê TRƯỚC giảm trừ...")
        df_stats_before = calculate_statistics(df_raw, has_ten_doi, dat_column='PHIEU_DAT', dat_value=1)

        # Tính thống kê SAU giảm trừ - LOGIC MỚI: GIỮ NGUYÊN MẪU SỐ
        print("✓ Đang tính thống kê SAU giảm trừ (giữ nguyên mẫu số)...")
        df_stats_after = calculate_statistics_keep_denominator(
            df_stats_before, df_raw, exclusion_ids, has_ten_doi,
            dat_column='PHIEU_DAT', dat_value=1
        )
        
        # Merge kết quả để so sánh
        print("✓ Đang tạo báo cáo so sánh...")
        
        if has_ten_doi:
            merge_columns = ['TEN_DOI', 'NVKT']
        else:
            merge_columns = ['NVKT']
        
        df_comparison = pd.merge(
            df_stats_before,
            df_stats_after,
            on=merge_columns,
            how='outer',
            suffixes=(' (Thô)', ' (Sau GT)')
        )
        
        # Tính chênh lệch tỷ lệ
        df_comparison['Chênh lệch %'] = (
            df_comparison['Tỷ lệ % (Sau GT)'].fillna(0) - 
            df_comparison['Tỷ lệ % (Thô)'].fillna(0)
        ).round(2)
        
        # Sắp xếp cột theo thứ tự mong muốn
        if has_ten_doi:
            column_order = [
                'TEN_DOI', 'NVKT',
                'Tổng phiếu (Thô)', 'Tổng phiếu (Sau GT)',
                'Số phiếu đạt (Thô)', 'Số phiếu đạt (Sau GT)',
                'Tỷ lệ % (Thô)', 'Tỷ lệ % (Sau GT)',
                'Chênh lệch %'
            ]
        else:
            column_order = [
                'NVKT',
                'Tổng phiếu (Thô)', 'Tổng phiếu (Sau GT)',
                'Số phiếu đạt (Thô)', 'Số phiếu đạt (Sau GT)',
                'Tỷ lệ % (Thô)', 'Tỷ lệ % (Sau GT)',
                'Chênh lệch %'
            ]
        
        df_comparison = df_comparison[column_order]
        
        # Sắp xếp
        if has_ten_doi:
            df_comparison = df_comparison.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            df_comparison = df_comparison.sort_values('NVKT').reset_index(drop=True)
        
        # Tạo sheet tổng hợp
        tong_before = df_stats_before['Tổng phiếu'].sum()
        tong_after = df_stats_after['Tổng phiếu'].sum()
        dat_before = df_stats_before['Số phiếu đạt'].sum()
        dat_after = df_stats_after['Số phiếu đạt'].sum()
        tyle_before = round((dat_before / tong_before * 100), 2) if tong_before > 0 else 0
        tyle_after = round((dat_after / tong_after * 100), 2) if tong_after > 0 else 0
        
        df_tongke = pd.DataFrame([{
            'Chỉ tiêu': 'C1.1 SM2 - Tỷ lệ phiếu sửa chữa BRCD trong 72h',
            'Tổng phiếu (Thô)': tong_before,
            'Phiếu loại trừ': num_excluded,
            'Tổng phiếu (Sau GT)': tong_after,
            'Phiếu đạt (Thô)': dat_before,
            'Phiếu đạt (Sau GT)': dat_after,
            'Tỷ lệ % (Thô)': tyle_before,
            'Tỷ lệ % (Sau GT)': tyle_after,
            'Thay đổi %': round(tyle_after - tyle_before, 2)
        }])
        
        # Lấy danh sách phiếu bị loại trừ
        df_loai_tru = df_raw[df_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        cols_to_keep = ['BAOHONG_ID', 'MA_TB', 'TEN_TB', 'TEN_KV', 'TEN_DOI', 'NGAY_BAO_HONG', 'NGAY_NGHIEM_THU', 'TG', 'PHIEU_DAT']
        cols_available = [c for c in cols_to_keep if c in df_loai_tru.columns]
        df_loai_tru = df_loai_tru[cols_available]
        
        # Tính thống kê theo đơn vị (Tổ)
        df_unit_stats = calculate_unit_stats(df_stats_before, df_stats_after)
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C11_SM2.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_unit_stats.to_excel(writer, sheet_name='Thong_ke_theo_don_vi', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"✅ Đã tạo báo cáo so sánh C1.1 SM2")
        print(f"   - Tổng phiếu thô: {tong_before}")
        print(f"   - Phiếu loại trừ: {num_excluded}")
        print(f"   - Tổng phiếu sau GT: {tong_after}")
        print(f"   - Tỷ lệ thô: {tyle_before}% -> Sau GT: {tyle_after}%")
        
        # Bổ sung cột Điểm BSC vào file vừa tạo
        add_bsc_scores_to_c11_sm2(output_file)
        
        return {
            'chi_tieu': 'C1.1 SM2',
            'tong_tho': tong_before,
            'loai_tru': num_excluded,
            'tong_sau_gt': tong_after,
            'tyle_tho': tyle_before,
            'tyle_sau_gt': tyle_after
        }
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo C1.1 SM2: {e}")
        import traceback
        traceback.print_exc()
        return None


def create_c12_comparison_report(exclusion_ids, output_dir):
    """
    Tạo báo cáo so sánh C1.2 trước/sau giảm trừ
    Chỉ tiêu: Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại
    
    Công thức: Tỷ lệ HLL = (Số phiếu HLL từ SM1-C12) / (Số phiếu báo hỏng từ SM2-C12) * 100
    
    QUAN TRỌNG: Sử dụng sheet TH_SM1C12_HLL_Thang có sẵn thay vì tính lại từ Sheet1
    
    Args:
        exclusion_ids: Set các BAOHONG_ID cần loại trừ
        output_dir: Thư mục xuất kết quả
        
    Returns:
        dict: Thống kê tổng hợp
    """
    import math
    
    try:
        print("\n" + "="*80)
        print("TẠO BÁO CÁO SO SÁNH C1.2 (SM1-C12 + SM2-C12) TRƯỚC/SAU GIẢM TRỪ")
        print("="*80)
        
        input_file_sm1 = os.path.join("downloads", "baocao_hanoi", "SM1-C12.xlsx")
        input_file_sm2 = os.path.join("downloads", "baocao_hanoi", "SM2-C12.xlsx")
        
        if not os.path.exists(input_file_sm1):
            print(f"❌ Không tìm thấy file: {input_file_sm1}")
            return None
        
        if not os.path.exists(input_file_sm2):
            print(f"❌ Không tìm thấy file: {input_file_sm2}")
            return None
        
        # ========== ĐỌC DỮ LIỆU TRƯỚC GIẢM TRỪ ==========
        print("\n" + "-"*40)
        print("ĐỌC DỮ LIỆU TRƯỚC GIẢM TRỪ")
        print("-"*40)
        
        # Đọc sheet TH_SM1C12_HLL_Thang (dữ liệu đã tổng hợp sẵn)
        df_before = pd.read_excel(input_file_sm1, sheet_name='TH_SM1C12_HLL_Thang')
        print(f"✅ Đã đọc TH_SM1C12_HLL_Thang (trước GT): {len(df_before)} dòng")
        print(f"   - Tổng HLL: {df_before['Số phiếu HLL'].sum()}")
        print(f"   - Tổng BH: {df_before['Số phiếu báo hỏng'].sum()}")
        
        has_ten_doi = 'TEN_DOI' in df_before.columns
        
        # ========== XỬ LÝ DỮ LIỆU SAU GIẢM TRỪ ==========
        print("\n" + "-"*40)
        print("XỬ LÝ DỮ LIỆU SAU GIẢM TRỪ")
        print("-"*40)
        
        # ĐỌC VÀ LỌC SM1 (dữ liệu thô)
        df_sm1_raw = pd.read_excel(input_file_sm1, sheet_name='Sheet1')
        print(f"✅ Đã đọc SM1-C12 Sheet1: {len(df_sm1_raw)} dòng")
        
        df_sm1_raw['NVKT'] = df_sm1_raw['TEN_KV'].apply(extract_nvkt_name)
        df_sm1_raw = df_sm1_raw[df_sm1_raw['NVKT'].notna()].copy()
        
        df_sm1_raw['BAOHONG_ID_STR'] = df_sm1_raw['BAOHONG_ID'].apply(normalize_id)
        df_sm1_excluded = df_sm1_raw[~df_sm1_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        num_excluded_sm1 = len(df_sm1_raw) - len(df_sm1_excluded)
        print(f"✅ Loại trừ SM1: {num_excluded_sm1} phiếu, còn lại {len(df_sm1_excluded)} phiếu")
        
        # Tính số phiếu HLL (đếm MA_TB duy nhất)
        def calculate_hll_by_nvkt(df, has_ten_doi):
            report_data = []
            if has_ten_doi:
                for (ten_doi, nvkt), group_df in df.groupby(['TEN_DOI', 'NVKT']):
                    if pd.isna(nvkt):
                        continue
                    so_phieu_hll = group_df['MA_TB'].nunique()
                    report_data.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Số phiếu HLL': so_phieu_hll
                    })
            else:
                for nvkt, group_df in df.groupby('NVKT'):
                    if pd.isna(nvkt):
                        continue
                    so_phieu_hll = group_df['MA_TB'].nunique()
                    report_data.append({
                        'NVKT': nvkt,
                        'Số phiếu HLL': so_phieu_hll
                    })
            return pd.DataFrame(report_data)
        
        df_hll_after = calculate_hll_by_nvkt(df_sm1_excluded, has_ten_doi)
        print(f"✅ HLL sau GT: {df_hll_after['Số phiếu HLL'].sum()} phiếu")
        
        # ========== ĐỌC VÀ XỬ LÝ SM2-C12 ==========
        print("\n" + "-"*40)
        print("XỬ LÝ SM2-C12 (Tổng phiếu báo hỏng)")
        print("-"*40)
        
        df_sm2_raw = pd.read_excel(input_file_sm2, sheet_name='Sheet1')
        print(f"✅ Đã đọc SM2-C12 thô: {len(df_sm2_raw)} dòng")
        
        # Chuẩn hóa cột NVKT
        df_sm2_raw['NVKT'] = df_sm2_raw['TEN_KV'].apply(extract_nvkt_name)
        df_sm2_raw = df_sm2_raw[df_sm2_raw['NVKT'].notna()].copy()

        # KHÔNG ÁP DỤNG GIẢM TRỪ CHO SM2 (Mẫu số giữ nguyên)
        print(f"✅ Mẫu số (SM2) GIỮ NGUYÊN: {len(df_sm2_raw)} phiếu (không áp dụng giảm trừ)")

        has_ten_doi_sm2 = 'TEN_DOI' in df_sm2_raw.columns

        # Tính số phiếu báo hỏng theo NVKT (Mẫu số - không thay đổi trước/sau GT)
        def calculate_bh_by_nvkt(df, has_ten_doi):
            report_data = []
            if has_ten_doi:
                for (ten_doi, nvkt), group_df in df.groupby(['TEN_DOI', 'NVKT']):
                    if pd.isna(nvkt):
                        continue
                    report_data.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': group_df['MA_TB'].nunique()
                    })
            else:
                for nvkt, group_df in df.groupby('NVKT'):
                    if pd.isna(nvkt):
                        continue
                    report_data.append({
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': group_df['MA_TB'].nunique()
                    })
            return pd.DataFrame(report_data)

        # Mẫu số giữ nguyên trước và sau giảm trừ
        df_bh_before = calculate_bh_by_nvkt(df_sm2_raw, has_ten_doi_sm2)
        df_bh_after = df_bh_before.copy()  # Mẫu số sau GT = Mẫu số trước GT

        print(f"✅ Báo hỏng (Mẫu số): {df_bh_before['Số phiếu báo hỏng'].sum()} phiếu (giữ nguyên trước/sau GT)")
        
        
        # ========== KẾT HỢP VÀ TÍNH TỶ LỆ ==========
        print("\n" + "-"*40)
        print("TẠO BÁO CÁO SO SÁNH")
        print("-"*40)        
        
        # Merge HLL sau GT với báo hỏng sau GT
        if has_ten_doi and has_ten_doi_sm2:
            merge_cols = ['TEN_DOI', 'NVKT']
        else:
            merge_cols = ['NVKT']
        
        # Tính tỷ lệ cho dữ liệu TRƯỚC giảm trừ (đã có sẵn trong df_before)
        if 'Tỉ lệ HLL tháng (2.5%)' in df_before.columns:
            df_before = df_before.rename(columns={'Tỉ lệ HLL tháng (2.5%)': 'Tỷ lệ HLL %'})
        else:
            df_before['Tỷ lệ HLL %'] = df_before.apply(
                lambda r: round((r['Số phiếu HLL'] / r['Số phiếu báo hỏng'] * 100), 2) if r['Số phiếu báo hỏng'] > 0 else 0,
                axis=1
            )
        
        # Merge HLL với báo hỏng (SAU giảm trừ)
        df_after = pd.merge(df_hll_after, df_bh_after, on=merge_cols, how='outer')
        df_after['Số phiếu HLL'] = df_after['Số phiếu HLL'].fillna(0).astype(int)
        df_after['Số phiếu báo hỏng'] = df_after['Số phiếu báo hỏng'].fillna(0).astype(int)
        df_after['Tỷ lệ HLL %'] = df_after.apply(
            lambda r: round((r['Số phiếu HLL'] / r['Số phiếu báo hỏng'] * 100), 2) if r['Số phiếu báo hỏng'] > 0 else 0,
            axis=1
        )
        
        # Merge để so sánh trước/sau
        df_comparison = pd.merge(
            df_before[merge_cols + ['Số phiếu HLL', 'Số phiếu báo hỏng', 'Tỷ lệ HLL %']],
            df_after[merge_cols + ['Số phiếu HLL', 'Số phiếu báo hỏng', 'Tỷ lệ HLL %']],
            on=merge_cols,
            how='outer',
            suffixes=(' (Thô)', ' (Sau GT)')
        )
        
        # Tính chênh lệch
        df_comparison['Chênh lệch %'] = (
            df_comparison['Tỷ lệ HLL % (Sau GT)'].fillna(0) - 
            df_comparison['Tỷ lệ HLL % (Thô)'].fillna(0)
        ).round(2)
        
        # Sắp xếp cột
        if has_ten_doi:
            column_order = [
                'TEN_DOI', 'NVKT',
                'Số phiếu HLL (Thô)', 'Số phiếu HLL (Sau GT)',
                'Số phiếu báo hỏng (Thô)', 'Số phiếu báo hỏng (Sau GT)',
                'Tỷ lệ HLL % (Thô)', 'Tỷ lệ HLL % (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            column_order = [
                'NVKT',
                'Số phiếu HLL (Thô)', 'Số phiếu HLL (Sau GT)',
                'Số phiếu báo hỏng (Thô)', 'Số phiếu báo hỏng (Sau GT)',
                'Tỷ lệ HLL % (Thô)', 'Tỷ lệ HLL % (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values('NVKT').reset_index(drop=True)
        
        df_comparison = df_comparison[column_order]
        
        # Tạo sheet tổng hợp
        hll_tho = df_before['Số phiếu HLL'].sum()
        hll_sau = df_after['Số phiếu HLL'].sum()
        bh_tho = df_before['Số phiếu báo hỏng'].sum()
        bh_sau = df_after['Số phiếu báo hỏng'].sum()
        tyle_tho = round((hll_tho / bh_tho * 100), 2) if bh_tho > 0 else 0
        tyle_sau = round((hll_sau / bh_sau * 100), 2) if bh_sau > 0 else 0
        
        df_tongke = pd.DataFrame([{
            'Chỉ tiêu': 'C1.2 - Tỷ lệ thuê bao báo hỏng BRCĐ lặp lại',
            'Phiếu HLL (Thô)': hll_tho,
            'Phiếu HLL (Sau GT)': hll_sau,
            'Loại trừ SM1': num_excluded_sm1,
            'Phiếu báo hỏng (Thô)': bh_tho,
            'Phiếu báo hỏng (Sau GT)': bh_sau,
            'Ghi chú mẫu số': 'Giữ nguyên (không loại trừ SM2)',
            'Tỷ lệ HLL % (Thô)': tyle_tho,
            'Tỷ lệ HLL % (Sau GT)': tyle_sau,
            'Thay đổi %': round(tyle_sau - tyle_tho, 2)
        }])
        
        # Lấy danh sách phiếu bị loại trừ từ SM1 (thực sự loại trừ)
        df_loai_tru_sm1 = df_sm1_raw[df_sm1_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()

        # Lấy danh sách phiếu trong exclusion_ids từ SM2 (CHỈ ĐỂ THAM KHẢO, không loại trừ khỏi mẫu số)
        df_sm2_raw['BAOHONG_ID_STR'] = df_sm2_raw['BAOHONG_ID'].apply(normalize_id)
        df_loai_tru_sm2 = df_sm2_raw[df_sm2_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        df_loai_tru_sm2['Ghi chú'] = 'Chỉ tham khảo - Không loại trừ khỏi mẫu số'
        
        cols_to_keep = ['BAOHONG_ID', 'MA_TB', 'TEN_TB', 'TEN_KV', 'TEN_DOI', 'NGAY_BAO_HONG']
        cols_sm1 = [c for c in cols_to_keep if c in df_loai_tru_sm1.columns]
        cols_sm2 = [c for c in cols_to_keep if c in df_loai_tru_sm2.columns]
        
        df_loai_tru_sm1 = df_loai_tru_sm1[cols_sm1]
        df_loai_tru_sm1['Nguồn'] = 'SM1-C12'
        df_loai_tru_sm2 = df_loai_tru_sm2[cols_sm2]
        df_loai_tru_sm2['Nguồn'] = 'SM2-C12'
        
        # Tính thống kê theo đơn vị (Tổ) cho C1.2
        # Chuẩn bị df_before_unit và df_after_unit để tính theo đội
        # Tính lại từ df_before và df_after nhưng chỉ dùng 2 cột quan trọng
        df_before_unit = df_before[merge_cols + ['Số phiếu HLL', 'Số phiếu báo hỏng']].copy()
        df_before_unit = df_before_unit.rename(columns={
            'Số phiếu HLL': 'Số phiếu đạt',
            'Số phiếu báo hỏng': 'Tổng phiếu'
        })
        df_after_unit = df_after[merge_cols + ['Số phiếu HLL', 'Số phiếu báo hỏng']].copy()
        df_after_unit = df_after_unit.rename(columns={
            'Số phiếu HLL': 'Số phiếu đạt',
            'Số phiếu báo hỏng': 'Tổng phiếu'
        })
        
        df_unit_stats = calculate_unit_stats(
            df_before_unit, df_after_unit, 
            tong_col='Tổng phiếu', 
            dat_col='Số phiếu đạt'
        )
        # Đổi tên cột cho phù hợp với C1.2
        df_unit_stats = df_unit_stats.rename(columns={
            'Phiếu đạt (Thô)': 'Phiếu HLL (Thô)',
            'Phiếu đạt (Sau GT)': 'Phiếu HLL (Sau GT)',
            'Tỷ lệ % (Thô)': 'Tỷ lệ HLL % (Thô)',
            'Tỷ lệ % (Sau GT)': 'Tỷ lệ HLL % (Sau GT)',
            'Tổng phiếu (Thô)': 'Phiếu báo hỏng (Thô)',
            'Tổng phiếu (Sau GT)': 'Phiếu báo hỏng (Sau GT)'
        })
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C12_SM1.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_unit_stats.to_excel(writer, sheet_name='Thong_ke_theo_don_vi', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru_sm1.to_excel(writer, sheet_name='DS_loai_tru_SM1', index=False)
            df_loai_tru_sm2.to_excel(writer, sheet_name='DS_loai_tru_SM2', index=False)
        
        print(f"\n✅ Đã tạo báo cáo so sánh C1.2")
        print(f"   - Phiếu HLL (Tử số): {hll_tho} -> {hll_sau} (loại trừ SM1: {num_excluded_sm1})")
        print(f"   - Phiếu báo hỏng (Mẫu số): {bh_tho} -> {bh_sau} (GIỮ NGUYÊN - không loại trừ SM2)")
        print(f"   - Tỷ lệ HLL: {tyle_tho}% -> {tyle_sau}% (Δ: {round(tyle_sau - tyle_tho, 2)}%)")
        
        # Bổ sung cột Điểm BSC vào file vừa tạo
        add_bsc_scores_to_c12_sm1(output_file)
        
        return {
            'chi_tieu': 'C1.2',
            'tong_tho': bh_tho,
            'loai_tru': num_excluded_sm1,  # Chỉ tử số (SM1) loại trừ, mẫu số (SM2) giữ nguyên
            'tong_sau_gt': bh_sau,
            'tyle_tho': tyle_tho,
            'tyle_sau_gt': tyle_sau
        }
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo C1.2: {e}")
        import traceback
        traceback.print_exc()
        return None


def create_sm1_c12_excluded_file(exclusion_ids, output_dir):
    """
    Tạo file SM1-C12 sau giảm trừ với đúng cấu trúc như file gốc:
    - Sheet1: Dữ liệu thô đã loại trừ
    - TH_phieu_hong_lai_7_ngay: Tổng hợp phiếu HLL theo NVKT
    - TH_SM1C12_HLL_Thang: Tỷ lệ HLL tháng kết hợp với SM2-C12
    
    Args:
        exclusion_ids: Set các BAOHONG_ID cần loại trừ
        output_dir: Thư mục xuất kết quả
        
    Returns:
        bool: True nếu thành công
    """
    import math
    
    try:
        print("\n" + "="*80)
        print("TẠO FILE SM1-C12 SAU GIẢM TRỪ (CẤU TRÚC GỐC)")
        print("="*80)
        
        # Đường dẫn files
        input_file_sm1 = os.path.join("downloads", "baocao_hanoi", "SM1-C12.xlsx")
        input_file_sm2 = os.path.join("downloads", "baocao_hanoi", "SM2-C12.xlsx")
        output_file = os.path.join(output_dir, "SM1-C12_sau_giam_tru.xlsx")
        
        if not os.path.exists(input_file_sm1):
            print(f"❌ Không tìm thấy file: {input_file_sm1}")
            return False
        
        # ========== BƯỚC 1: Đọc và lọc dữ liệu Sheet1 ==========
        print("\n" + "-"*40)
        print("BƯỚC 1: Xử lý dữ liệu Sheet1")
        print("-"*40)
        
        df_sm1 = pd.read_excel(input_file_sm1, sheet_name='Sheet1')
        print(f"✅ Đã đọc SM1-C12, tổng số dòng thô: {len(df_sm1)}")
        
        # Lọc dữ liệu sau giảm trừ
        df_sm1['BAOHONG_ID_STR'] = df_sm1['BAOHONG_ID'].apply(normalize_id)
        df_sm1_excluded = df_sm1[~df_sm1['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        
        # Xóa cột tạm
        df_sm1_excluded = df_sm1_excluded.drop(columns=['BAOHONG_ID_STR'])
        
        num_excluded = len(df_sm1) - len(df_sm1_excluded)
        print(f"✅ Đã loại trừ {num_excluded} phiếu, còn lại {len(df_sm1_excluded)} phiếu")
        
        # ========== BƯỚC 2: Tạo sheet TH_phieu_hong_lai_7_ngay ==========
        print("\n" + "-"*40)
        print("BƯỚC 2: Tạo sheet TH_phieu_hong_lai_7_ngay")
        print("-"*40)
        
        # Chuẩn hóa cột NVKT
        df_sm1_excluded['NVKT'] = df_sm1_excluded['TEN_KV'].apply(extract_nvkt_name)
        
        has_ten_doi = 'TEN_DOI' in df_sm1_excluded.columns
        
        # Nhóm theo TEN_DOI và NVKT, tính số phiếu HLL
        report_data_hll = []

        if has_ten_doi:
            for (ten_doi, nvkt), group_df in df_sm1_excluded.groupby(['TEN_DOI', 'NVKT']):
                if pd.isna(nvkt):
                    continue
                # Số phiếu HLL = Đếm số MA_TB duy nhất
                so_phieu_hll = group_df['MA_TB'].nunique()
                report_data_hll.append({
                    'TEN_DOI': ten_doi,
                    'NVKT': nvkt,
                    'Số phiếu HLL': so_phieu_hll
                })
        else:
            for nvkt, group_df in df_sm1_excluded.groupby('NVKT'):
                if pd.isna(nvkt):
                    continue
                so_phieu_hll = group_df['MA_TB'].nunique()
                report_data_hll.append({
                    'NVKT': nvkt,
                    'Số phiếu HLL': so_phieu_hll
                })
        
        df_hll_7ngay = pd.DataFrame(report_data_hll)
        if has_ten_doi:
            df_hll_7ngay = df_hll_7ngay.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            df_hll_7ngay = df_hll_7ngay.sort_values('NVKT').reset_index(drop=True)
        
        print(f"✅ Đã tạo sheet TH_phieu_hong_lai_7_ngay với {len(df_hll_7ngay)} dòng")
        
        # ========== BƯỚC 3: Tạo sheet TH_SM1C12_HLL_Thang ==========
        print("\n" + "-"*40)
        print("BƯỚC 3: Tạo sheet TH_SM1C12_HLL_Thang")
        print("-"*40)
        
        # Đọc dữ liệu SM2-C12 để lấy tổng phiếu báo hỏng
        if os.path.exists(input_file_sm2):
            df_sm2 = pd.read_excel(input_file_sm2, sheet_name='Sheet1')
            print(f"✅ Đã đọc SM2-C12, tổng số dòng thô: {len(df_sm2)}")

            # KHÔNG ÁP DỤNG GIẢM TRỪ CHO SM2-C12 (Mẫu số giữ nguyên)
            # Chuẩn hóa cột NVKT cho SM2 (dữ liệu gốc)
            df_sm2['NVKT'] = df_sm2['TEN_KV'].apply(extract_nvkt_name)

            has_ten_doi_sm2 = 'TEN_DOI' in df_sm2.columns

            # Tính tổng phiếu báo hỏng theo NVKT (KHÔNG giảm trừ - mẫu số giữ nguyên)
            print(f"✅ Mẫu số (Số phiếu báo hỏng) GIỮ NGUYÊN: {len(df_sm2)} phiếu (không áp dụng giảm trừ)")
            report_data_bh = []

            if has_ten_doi_sm2:
                for (ten_doi, nvkt), group_df in df_sm2.groupby(['TEN_DOI', 'NVKT']):
                    if pd.isna(nvkt):
                        continue
                    report_data_bh.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': group_df['MA_TB'].nunique()
                    })
            else:
                for nvkt, group_df in df_sm2.groupby('NVKT'):
                    if pd.isna(nvkt):
                        continue
                    report_data_bh.append({
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': group_df['MA_TB'].nunique()
                    })
            
            df_bao_hong = pd.DataFrame(report_data_bh)
            
            # Merge dữ liệu HLL với tổng phiếu báo hỏng
            if has_ten_doi and has_ten_doi_sm2:
                df_merged = pd.merge(
                    df_hll_7ngay[['TEN_DOI', 'NVKT', 'Số phiếu HLL']],
                    df_bao_hong[['TEN_DOI', 'NVKT', 'Số phiếu báo hỏng']],
                    on=['TEN_DOI', 'NVKT'],
                    how='outer'
                )
            else:
                df_merged = pd.merge(
                    df_hll_7ngay[['NVKT', 'Số phiếu HLL']] if 'NVKT' in df_hll_7ngay.columns else df_hll_7ngay,
                    df_bao_hong[['NVKT', 'Số phiếu báo hỏng']] if 'NVKT' in df_bao_hong.columns else df_bao_hong,
                    on='NVKT',
                    how='outer'
                )
            
            # Điền 0 cho các giá trị NaN
            df_merged['Số phiếu HLL'] = df_merged['Số phiếu HLL'].fillna(0).astype(int)
            df_merged['Số phiếu báo hỏng'] = df_merged['Số phiếu báo hỏng'].fillna(0).astype(int)
            
            # Tính Tỉ lệ HLL tháng
            def calculate_ty_le_hll(row):
                if row['Số phiếu báo hỏng'] == 0:
                    return 0
                return round((row['Số phiếu HLL'] / row['Số phiếu báo hỏng']) * 100, 2)
            
            df_merged['Tỉ lệ HLL tháng (2.5%)'] = df_merged.apply(calculate_ty_le_hll, axis=1)
            
            # Sắp xếp cột
            if 'TEN_DOI' in df_merged.columns:
                df_hll_thang = df_merged[['TEN_DOI', 'NVKT', 'Số phiếu HLL', 'Số phiếu báo hỏng', 'Tỉ lệ HLL tháng (2.5%)']].copy()
                df_hll_thang = df_hll_thang.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
            else:
                df_hll_thang = df_merged[['NVKT', 'Số phiếu HLL', 'Số phiếu báo hỏng', 'Tỉ lệ HLL tháng (2.5%)']].copy()
                df_hll_thang = df_hll_thang.sort_values('NVKT').reset_index(drop=True)
            
            print(f"✅ Đã tạo sheet TH_SM1C12_HLL_Thang với {len(df_hll_thang)} dòng")
        else:
            print(f"⚠️ Không tìm thấy file SM2-C12, sẽ tạo TH_SM1C12_HLL_Thang chỉ với dữ liệu SM1")
            df_hll_thang = df_hll_7ngay.copy()
            df_hll_thang['Số phiếu báo hỏng'] = 0
            df_hll_thang['Tỉ lệ HLL tháng (2.5%)'] = 0
        
        # ========== BƯỚC 4: Ghi vào file Excel ==========
        print("\n" + "-"*40)
        print("BƯỚC 4: Ghi file Excel")
        print("-"*40)
        
        # Xóa cột NVKT tạm khỏi Sheet1 (giữ nguyên cấu trúc gốc)
        if 'NVKT' in df_sm1_excluded.columns:
            df_sheet1 = df_sm1_excluded.drop(columns=['NVKT'])
        else:
            df_sheet1 = df_sm1_excluded
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
            df_hll_7ngay.to_excel(writer, sheet_name='TH_phieu_hong_lai_7_ngay', index=False)
            df_hll_thang.to_excel(writer, sheet_name='TH_SM1C12_HLL_Thang', index=False)
        
        print(f"\n✅ Đã tạo file: {output_file}")
        print(f"   - Sheet1: {len(df_sheet1)} dòng (dữ liệu thô sau giảm trừ)")
        print(f"   - TH_phieu_hong_lai_7_ngay: {len(df_hll_7ngay)} dòng")
        print(f"   - TH_SM1C12_HLL_Thang: {len(df_hll_thang)} dòng")
        
        # In thống kê so sánh
        print("\n" + "-"*40)
        print("SO SÁNH TRƯỚC/SAU GIẢM TRỪ:")
        print("-"*40)
        
        # Tính tỷ lệ HLL tổng
        tong_hll_sau = df_hll_thang['Số phiếu HLL'].sum()
        tong_bh_sau = df_hll_thang['Số phiếu báo hỏng'].sum()
        tyle_sau = round((tong_hll_sau / tong_bh_sau * 100), 2) if tong_bh_sau > 0 else 0
        
        # Đọc dữ liệu gốc để so sánh
        try:
            df_goc = pd.read_excel(input_file_sm1, sheet_name='TH_SM1C12_HLL_Thang')
            tong_hll_tho = df_goc['Số phiếu HLL'].sum()
            tong_bh_tho = df_goc['Số phiếu báo hỏng'].sum()
            tyle_tho = round((tong_hll_tho / tong_bh_tho * 100), 2) if tong_bh_tho > 0 else 0
            
            print(f"  Tổng phiếu HLL: {tong_hll_tho} -> {tong_hll_sau} (Δ: {tong_hll_sau - tong_hll_tho})")
            print(f"  Tổng phiếu báo hỏng: {tong_bh_tho} -> {tong_bh_sau} (Δ: {tong_bh_sau - tong_bh_tho})")
            print(f"  Tỷ lệ HLL: {tyle_tho}% -> {tyle_sau}% (Δ: {round(tyle_sau - tyle_tho, 2)}%)")
        except Exception as e:
            print(f"  Tổng phiếu HLL sau GT: {tong_hll_sau}")
            print(f"  Tổng phiếu báo hỏng: {tong_bh_sau}")
            print(f"  Tỷ lệ HLL sau GT: {tyle_sau}%")
        
        print("-"*40)
        
        return True
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo file SM1-C12 sau giảm trừ: {e}")
        import traceback
        traceback.print_exc()
        return False


def create_c12_ti_le_bao_hong_comparison_report(exclusion_ids, output_dir):
    """
    Tạo báo cáo so sánh C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng trước/sau giảm trừ
    
    Công thức: Tỷ lệ báo hỏng = (Số phiếu báo hỏng BRCĐ / Tổng thuê bao) * 100
    
    Dữ liệu:
    - Số phiếu báo hỏng: Đếm từ SM4-C11.xlsx Sheet1, lọc các dịch vụ BRCĐ
    - Tổng thuê bao: Từ file tham chiếu Tonghop_thuebao_NVKT_DB_C12.xlsx
    
    Args:
        exclusion_ids: Set các BAOHONG_ID cần loại trừ
        output_dir: Thư mục xuất kết quả
        
    Returns:
        dict: Thống kê tổng hợp
    """
    try:
        print("\n" + "="*80)
        print("TẠO BÁO CÁO SO SÁNH C1.2 - TỶ LỆ THUÊ BAO BRCĐ BÁO HỎNG TRƯỚC/SAU GIẢM TRỪ")
        print("="*80)
        
        input_file = os.path.join("downloads", "baocao_hanoi", "SM4-C11.xlsx")
        ref_file = os.path.join("du_lieu_tham_chieu", "Tonghop_thuebao_NVKT_DB_C12.xlsx")
        
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return None
        
        if not os.path.exists(ref_file):
            print(f"❌ Không tìm thấy file tham chiếu: {ref_file}")
            return None
        
        # ========== ĐỌC VÀ XỬ LÝ SM4-C11 ==========
        print("\n" + "-"*40)
        print("XỬ LÝ SM4-C11 (Phiếu báo hỏng dịch vụ BRCĐ)")
        print("-"*40)
        
        df_raw = pd.read_excel(input_file, sheet_name=0)  # Sheet 1
        print(f"✅ Đã đọc SM4-C11 thô: {len(df_raw)} dòng")
        
        # Kiểm tra cột cần thiết
        if 'BAOHONG_ID' not in df_raw.columns:
            print("❌ Không tìm thấy cột BAOHONG_ID")
            return None
        
        if 'TEN_DICH_VU' not in df_raw.columns:
            print("❌ Không tìm thấy cột TEN_DICH_VU")
            return None
        
        if 'TEN_KV' not in df_raw.columns:
            print("❌ Không tìm thấy cột TEN_KV")
            return None
        
        # Lọc các dịch vụ BRCĐ
        dich_vu_brcd = ['Băng rộng cố định', 'Cố định', 'IMS']
        df_brcd = df_raw[df_raw['TEN_DICH_VU'].isin(dich_vu_brcd)].copy()
        print(f"✅ Đã lọc {len(df_brcd)} bản ghi dịch vụ BRCĐ")
        
        # Chuẩn hóa cột NVKT
        df_brcd['NVKT'] = df_brcd['TEN_KV'].apply(extract_nvkt_name)
        df_brcd = df_brcd[df_brcd['NVKT'].notna()].copy()
        
        # Chuẩn hóa tên về Title Case để đồng nhất
        df_brcd['NVKT'] = df_brcd['NVKT'].str.strip().str.title()
        
        has_ten_doi = 'TEN_DOI' in df_brcd.columns
        
        # Áp dụng giảm trừ
        df_brcd['BAOHONG_ID_STR'] = df_brcd['BAOHONG_ID'].apply(normalize_id)
        df_excluded = df_brcd[~df_brcd['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        num_excluded = len(df_brcd) - len(df_excluded)
        print(f"✅ Loại trừ: {num_excluded} phiếu, còn lại {len(df_excluded)} phiếu")
        
        # Đếm số phiếu báo hỏng theo NVKT (TRƯỚC giảm trừ)
        if has_ten_doi:
            df_count_before = df_brcd.groupby(['TEN_DOI', 'NVKT']).size().reset_index(name='Số phiếu báo hỏng')
            df_count_after = df_excluded.groupby(['TEN_DOI', 'NVKT']).size().reset_index(name='Số phiếu báo hỏng')
        else:
            df_count_before = df_brcd.groupby('NVKT').size().reset_index(name='Số phiếu báo hỏng')
            df_count_after = df_excluded.groupby('NVKT').size().reset_index(name='Số phiếu báo hỏng')
        
        print(f"✅ Số phiếu trước GT: {df_count_before['Số phiếu báo hỏng'].sum()}")
        print(f"✅ Số phiếu sau GT: {df_count_after['Số phiếu báo hỏng'].sum()}")
        
        # ========== ĐỌC FILE THAM CHIẾU THUÊ BAO ==========
        print("\n" + "-"*40)
        print("XỬ LÝ FILE THAM CHIẾU THUÊ BAO")
        print("-"*40)
        
        df_ref = pd.read_excel(ref_file)
        print(f"✅ Đã đọc file tham chiếu: {len(df_ref)} dòng")
        
        # Tìm cột tên NVKT
        nvkt_col = None
        for col in df_ref.columns:
            if 'NVKT' in col.upper() or 'TÊN' in col.upper():
                nvkt_col = col
                break
        
        if nvkt_col is None or 'Tổng TB' not in df_ref.columns:
            print("❌ Không tìm thấy cột NVKT hoặc Tổng TB trong file tham chiếu")
            return None
        
        df_ref_clean = df_ref[[nvkt_col, 'Tổng TB']].copy()
        df_ref_clean.columns = ['NVKT', 'Tổng TB']
        
        # Chuẩn hóa tên NVKT để khớp với dữ liệu từ df_brcd
        df_ref_clean['NVKT'] = df_ref_clean['NVKT'].apply(extract_nvkt_name)
        df_ref_clean = df_ref_clean.dropna(subset=['NVKT', 'Tổng TB'])
        
        # Chuẩn hóa tên về Title Case để đồng nhất
        df_ref_clean['NVKT'] = df_ref_clean['NVKT'].str.strip().str.title()
        
        # Gộp các bản sao (nếu có) - cộng tổng TB
        df_ref_clean = df_ref_clean.groupby('NVKT', as_index=False).agg({'Tổng TB': 'sum'})
        
        print(f"✅ Đã xử lý file tham chiếu: {len(df_ref_clean)} NVKT")
        
        # ========== MERGE VÀ TÍNH TỶ LỆ ==========
        print("\n" + "-"*40)
        print("TẠO BÁO CÁO SO SÁNH")
        print("-"*40)
        
        # Merge dữ liệu TRƯỚC giảm trừ
        df_before = pd.merge(df_count_before, df_ref_clean, on='NVKT', how='outer')
        df_before['Số phiếu báo hỏng'] = df_before['Số phiếu báo hỏng'].fillna(0).astype(int)
        df_before['Tổng TB'] = df_before['Tổng TB'].fillna(0).astype(int)
        df_before['Tỷ lệ báo hỏng (%)'] = df_before.apply(
            lambda r: round((r['Số phiếu báo hỏng'] / r['Tổng TB'] * 100), 2) if r['Tổng TB'] > 0 else 0, 
            axis=1
        )
        
        # Merge dữ liệu SAU giảm trừ
        df_after = pd.merge(df_count_after, df_ref_clean, on='NVKT', how='outer')
        df_after['Số phiếu báo hỏng'] = df_after['Số phiếu báo hỏng'].fillna(0).astype(int)
        df_after['Tổng TB'] = df_after['Tổng TB'].fillna(0).astype(int)
        df_after['Tỷ lệ báo hỏng (%)'] = df_after.apply(
            lambda r: round((r['Số phiếu báo hỏng'] / r['Tổng TB'] * 100), 2) if r['Tổng TB'] > 0 else 0, 
            axis=1
        )
        
        # Merge để so sánh trước/sau
        merge_cols = ['TEN_DOI', 'NVKT'] if has_ten_doi else ['NVKT']
        
        # Nếu không có TEN_DOI, chỉ dùng NVKT
        if 'TEN_DOI' not in df_before.columns:
            merge_cols = ['NVKT']
        
        df_comparison = pd.merge(
            df_before, df_after,
            on=merge_cols,
            how='outer',
            suffixes=(' (Thô)', ' (Sau GT)')
        )
        
        # Điền giá trị mặc định
        for col in df_comparison.columns:
            if 'Số phiếu' in col or 'Tổng TB' in col:
                df_comparison[col] = df_comparison[col].fillna(0).astype(int)
            elif 'Tỷ lệ' in col:
                df_comparison[col] = df_comparison[col].fillna(0)
        
        # Tính chênh lệch
        df_comparison['Chênh lệch %'] = (
            df_comparison['Tỷ lệ báo hỏng (%) (Sau GT)'].fillna(0) - 
            df_comparison['Tỷ lệ báo hỏng (%) (Thô)'].fillna(0)
        ).round(2)
        
        # Sắp xếp cột
        if 'TEN_DOI' in df_comparison.columns:
            column_order = [
                'TEN_DOI', 'NVKT', 'Tổng TB (Thô)',
                'Số phiếu báo hỏng (Thô)', 'Số phiếu báo hỏng (Sau GT)',
                'Tỷ lệ báo hỏng (%) (Thô)', 'Tỷ lệ báo hỏng (%) (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            column_order = [
                'NVKT', 'Tổng TB (Thô)',
                'Số phiếu báo hỏng (Thô)', 'Số phiếu báo hỏng (Sau GT)',
                'Tỷ lệ báo hỏng (%) (Thô)', 'Tỷ lệ báo hỏng (%) (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values('NVKT').reset_index(drop=True)
        
        # Chỉ lấy các cột tồn tại
        column_order = [c for c in column_order if c in df_comparison.columns]
        df_comparison = df_comparison[column_order]
        
        # Loại bỏ các dòng có NVKT rỗng
        df_comparison = df_comparison[df_comparison['NVKT'].notna()]
        
        # Tạo sheet tổng hợp
        bh_tho = df_before['Số phiếu báo hỏng'].sum()
        bh_sau = df_after['Số phiếu báo hỏng'].sum()
        tb_tong = df_ref_clean['Tổng TB'].sum()
        tyle_tho = round((bh_tho / tb_tong * 100), 2) if tb_tong > 0 else 0
        tyle_sau = round((bh_sau / tb_tong * 100), 2) if tb_tong > 0 else 0
        
        df_tongke = pd.DataFrame([{
            'Chỉ tiêu': 'C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng',
            'Tổng thuê bao': tb_tong,
            'Phiếu báo hỏng (Thô)': bh_tho,
            'Phiếu báo hỏng (Sau GT)': bh_sau,
            'Phiếu loại trừ': num_excluded,
            'Tỷ lệ % (Thô)': tyle_tho,
            'Tỷ lệ % (Sau GT)': tyle_sau,
            'Thay đổi %': round(tyle_sau - tyle_tho, 2)
        }])
        
        # Lấy danh sách phiếu bị loại trừ
        df_loai_tru = df_brcd[df_brcd['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        cols_to_keep = ['BAOHONG_ID', 'MA_TB', 'TEN_TB', 'TEN_KV', 'TEN_DOI', 'TEN_DICH_VU', 'NGAY_BAO_HONG']
        cols_available = [c for c in cols_to_keep if c in df_loai_tru.columns]
        df_loai_tru = df_loai_tru[cols_available]
        
        # Tính thống kê theo đơn vị (Tổ) cho C1.2 TP2
        # Chuẩn bị df_before_unit và df_after_unit để tính theo đội
        df_before_unit = df_before[merge_cols + ['Số phiếu báo hỏng', 'Tổng TB']].copy()
        df_before_unit = df_before_unit.rename(columns={
            'Số phiếu báo hỏng': 'Số phiếu đạt',
            'Tổng TB': 'Tổng phiếu'
        })
        df_after_unit = df_after[merge_cols + ['Số phiếu báo hỏng', 'Tổng TB']].copy()
        df_after_unit = df_after_unit.rename(columns={
            'Số phiếu báo hỏng': 'Số phiếu đạt',
            'Tổng TB': 'Tổng phiếu'
        })
        
        df_unit_stats = calculate_unit_stats(
            df_before_unit, df_after_unit, 
            tong_col='Tổng phiếu', 
            dat_col='Số phiếu đạt'
        )
        # Đổi tên cột cho phù hợp với C1.2 TP2
        df_unit_stats = df_unit_stats.rename(columns={
            'Phiếu đạt (Thô)': 'Phiếu báo hỏng (Thô)',
            'Phiếu đạt (Sau GT)': 'Phiếu báo hỏng (Sau GT)',
            'Tổng phiếu (Thô)': 'Tổng TB (Thô)',
            'Tổng phiếu (Sau GT)': 'Tổng TB (Sau GT)',
            'Tỷ lệ % (Thô)': 'Tỷ lệ báo hỏng % (Thô)',
            'Tỷ lệ % (Sau GT)': 'Tỷ lệ báo hỏng % (Sau GT)',
            'Thay đổi %': 'Chênh lệch %'
        })
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "SM4-C12-ti-le-su-co-dv-brcd.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_unit_stats.to_excel(writer, sheet_name='Thong_ke_theo_don_vi', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"\n✅ Đã tạo báo cáo so sánh C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng")
        print(f"   - Tổng thuê bao: {tb_tong}")
        print(f"   - Phiếu báo hỏng: {bh_tho} -> {bh_sau} (loại trừ: {num_excluded})")
        print(f"   - Tỷ lệ: {tyle_tho}% -> {tyle_sau}% (Δ: {round(tyle_sau - tyle_tho, 2)}%)")
        
        # Bổ sung cột Điểm BSC vào file vừa tạo
        add_bsc_scores_to_c12_tp2(output_file)
        
        return {
            'chi_tieu': 'C1.2 Tỷ lệ BRCĐ báo hỏng',
            'tong_tho': bh_tho,
            'loai_tru': num_excluded,
            'tong_sau_gt': bh_sau,
            'tyle_tho': tyle_tho,
            'tyle_sau_gt': tyle_sau
        }
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng: {e}")
        import traceback
        traceback.print_exc()
        return None


def load_c14_exclusion_list(exclusion_file="du_lieu_tham_chieu/LOAI_TRU_C1.4.xlsx"):
    """
    Đọc danh sách BAOHONG_ID cần loại trừ cho C1.4 từ file Excel
    
    Returns:
        set: Tập hợp các BAOHONG_ID cần loại trừ (dạng string)
    """
    try:
        if not os.path.exists(exclusion_file):
            print(f"⚠️ Không tìm thấy file loại trừ C1.4: {exclusion_file}")
            return set()
        
        df = pd.read_excel(exclusion_file)
        
        if 'BAOHONG_ID' not in df.columns:
            print(f"⚠️ Không tìm thấy cột 'BAOHONG_ID' trong file {exclusion_file}")
            return set()
        
        # Chuẩn hóa tất cả ID
        exclusion_ids = {normalize_id(idx) for idx in df['BAOHONG_ID'].tolist() if pd.notna(idx)}
        exclusion_ids.discard("")
        
        print(f"✅ Đã đọc {len(exclusion_ids)} mã BAOHONG_ID loại trừ C1.4")
        return exclusion_ids
        
    except Exception as e:
        print(f"❌ Lỗi khi đọc file loại trừ C1.4: {e}")
        return set()


def create_c14_comparison_report(exclusion_ids, output_dir):
    """
    Tạo báo cáo so sánh C1.4 - Độ hài lòng khách hàng trước/sau giảm trừ
    
    Công thức: Tỷ lệ hài lòng = (Tổng phiếu KS thành công - Tổng phiếu KHL) / Tổng phiếu KS thành công * 100
    
    Args:
        exclusion_ids: Set các BAOHONG_ID cần loại trừ
        output_dir: Thư mục xuất kết quả
        
    Returns:
        dict: Kết quả so sánh hoặc None nếu lỗi
    """
    try:
        print("\n" + "="*80)
        print("TẠO BÁO CÁO SO SÁNH C1.4 - ĐỘ HÀI LÒNG KHÁCH HÀNG")
        print("="*80)
        
        # Đọc dữ liệu chi tiết C1.4
        data_file = os.path.join("downloads", "baocao_hanoi", "c1.4_chitiet_report.xlsx")
        
        if not os.path.exists(data_file):
            print(f"❌ Không tìm thấy file dữ liệu C1.4: {data_file}")
            return None
        
        df_raw = pd.read_excel(data_file, sheet_name="Sheet1")
        print(f"✅ Đã đọc file dữ liệu: {len(df_raw)} phiếu khảo sát")
        
        # Chuẩn hóa cột NVKT
        if 'TEN_NVKT_DB' in df_raw.columns:
            df_raw['NVKT'] = df_raw['TEN_NVKT_DB'].apply(extract_nvkt_name)
        elif 'TEN_KV' in df_raw.columns:
            df_raw['NVKT'] = df_raw['TEN_KV'].apply(extract_nvkt_name)
        else:
            print("❌ Không tìm thấy cột tên NVKT")
            return None
            
        df_raw = df_raw[df_raw['NVKT'].notna()].copy()
        df_raw['NVKT'] = df_raw['NVKT'].str.strip().str.title()
        
        # Xác định cột đội
        if 'DOIVT' in df_raw.columns:
            df_raw['TEN_DOI'] = df_raw['DOIVT']
        
        has_ten_doi = 'TEN_DOI' in df_raw.columns
        
        # Kiểm tra cột cần thiết theo logic c1_process.py
        if 'DO_HL' not in df_raw.columns or 'KHL_KT' not in df_raw.columns:
            print("❌ Không tìm thấy cột DO_HL hoặc KHL_KT")
            return None
        
        # Lọc chỉ các bản ghi có DO_HL = 'HL' hoặc 'KHL' (phiếu KS thành công)
        df_ks = df_raw[df_raw['DO_HL'].isin(['HL', 'KHL'])].copy()
        print(f"✅ Số phiếu KS thành công (DO_HL = 'HL' hoặc 'KHL'): {len(df_ks)}")
        
        # Xác định phiếu không hài lòng: KHL_KT != null
        df_ks['IS_KHL'] = df_ks['KHL_KT'].notna().astype(int)
        
        # Áp dụng giảm trừ
        df_ks['BAOHONG_ID_STR'] = df_ks['BAOHONG_ID'].apply(normalize_id)
        df_excluded = df_ks[~df_ks['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        num_excluded = len(df_ks) - len(df_excluded)
        print(f"✅ Loại trừ: {num_excluded} phiếu, còn lại {len(df_excluded)} phiếu")
        
        # ========== TÍNH TOÁN TRƯỚC GIẢM TRỪ ==========
        if has_ten_doi:
            df_before = df_ks.groupby(['TEN_DOI', 'NVKT']).agg({
                'IS_KHL': 'sum',      # Tổng phiếu KHL (KHL_KT != null)
                'BAOHONG_ID': 'size'  # Tổng phiếu KS thành công
            }).reset_index()
        else:
            df_before = df_ks.groupby('NVKT').agg({
                'IS_KHL': 'sum',
                'BAOHONG_ID': 'size'
            }).reset_index()
            
        df_before.columns = list(df_before.columns[:-2]) + ['Số phiếu KHL', 'Tổng phiếu KS']
        # Tỷ lệ HL = (Tổng - KHL) / Tổng * 100 = Số phiếu hài lòng / Tổng * 100
        df_before['Tỷ lệ HL (%)'] = ((df_before['Tổng phiếu KS'] - df_before['Số phiếu KHL']) / 
                                      df_before['Tổng phiếu KS'].replace(0, 1) * 100).round(2)
        
        # ========== TÍNH TOÁN SAU GIẢM TRỪ - LOGIC MỚI: GIỮ NGUYÊN MẪU SỐ ==========
        # CHỈ tính số phiếu KHL từ dữ liệu đã loại trừ
        if has_ten_doi:
            df_khl_after = df_excluded.groupby(['TEN_DOI', 'NVKT']).agg({
                'IS_KHL': 'sum'  # Chỉ đếm số phiếu KHL sau loại trừ
            }).reset_index()
            df_khl_after.columns = list(df_khl_after.columns[:-1]) + ['Số phiếu KHL']
        else:
            df_khl_after = df_excluded.groupby('NVKT').agg({
                'IS_KHL': 'sum'
            }).reset_index()
            df_khl_after.columns = list(df_khl_after.columns[:-1]) + ['Số phiếu KHL']

        # Merge với df_before để GIỮ NGUYÊN Tổng phiếu KS (mẫu số)
        merge_cols = ['TEN_DOI', 'NVKT'] if has_ten_doi else ['NVKT']
        df_after = pd.merge(
            df_before[merge_cols + ['Tổng phiếu KS']],  # Lấy mẫu số từ TRƯỚC GT
            df_khl_after,  # Lấy số phiếu KHL từ SAU GT
            on=merge_cols,
            how='left'
        )
        df_after['Số phiếu KHL'] = df_after['Số phiếu KHL'].fillna(0).astype(int)

        # Tính tỷ lệ với MẪU SỐ GIỮ NGUYÊN
        df_after['Tỷ lệ HL (%)'] = ((df_after['Tổng phiếu KS'] - df_after['Số phiếu KHL']) /
                                     df_after['Tổng phiếu KS'].replace(0, 1) * 100).round(2)
        
        # ========== MERGE VÀ SO SÁNH ==========
        merge_cols = ['TEN_DOI', 'NVKT'] if has_ten_doi else ['NVKT']
        
        df_comparison = pd.merge(
            df_before, df_after,
            on=merge_cols,
            how='outer',
            suffixes=(' (Thô)', ' (Sau GT)')
        )
        
        # Điền giá trị mặc định
        for col in df_comparison.columns:
            if 'phiếu' in col.lower():
                df_comparison[col] = df_comparison[col].fillna(0).astype(int)
            elif 'Tỷ lệ' in col:
                df_comparison[col] = df_comparison[col].fillna(100.0)
        
        # Tính chênh lệch
        df_comparison['Chênh lệch %'] = (
            df_comparison['Tỷ lệ HL (%) (Sau GT)'].fillna(100) - 
            df_comparison['Tỷ lệ HL (%) (Thô)'].fillna(100)
        ).round(2)
        
        # Sắp xếp cột
        if has_ten_doi:
            column_order = [
                'TEN_DOI', 'NVKT', 
                'Tổng phiếu KS (Thô)', 'Số phiếu KHL (Thô)', 'Tỷ lệ HL (%) (Thô)',
                'Tổng phiếu KS (Sau GT)', 'Số phiếu KHL (Sau GT)', 'Tỷ lệ HL (%) (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            column_order = [
                'NVKT', 
                'Tổng phiếu KS (Thô)', 'Số phiếu KHL (Thô)', 'Tỷ lệ HL (%) (Thô)',
                'Tổng phiếu KS (Sau GT)', 'Số phiếu KHL (Sau GT)', 'Tỷ lệ HL (%) (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values('NVKT').reset_index(drop=True)
        
        column_order = [c for c in column_order if c in df_comparison.columns]
        df_comparison = df_comparison[column_order]
        
        # Loại bỏ các dòng có NVKT rỗng
        df_comparison = df_comparison[df_comparison['NVKT'].notna()]
        
        # ========== TÍNH THỐNG KÊ TỔNG HỢP ==========
        tong_ks_tho = df_before['Tổng phiếu KS'].sum()
        tong_khl_tho = df_before['Số phiếu KHL'].sum()
        tyle_tho = round((tong_ks_tho - tong_khl_tho) / tong_ks_tho * 100, 2) if tong_ks_tho > 0 else 100
        
        tong_ks_sau = df_after['Tổng phiếu KS'].sum()
        tong_khl_sau = df_after['Số phiếu KHL'].sum()
        tyle_sau = round((tong_ks_sau - tong_khl_sau) / tong_ks_sau * 100, 2) if tong_ks_sau > 0 else 100
        
        df_tongke = pd.DataFrame([{
            'Chỉ tiêu': 'C1.4 - Độ hài lòng khách hàng',
            'Tổng phiếu KS (Thô)': tong_ks_tho,
            'Tổng phiếu KHL (Thô)': tong_khl_tho,
            'Tổng phiếu KS (Sau GT)': tong_ks_sau,
            'Tổng phiếu KHL (Sau GT)': tong_khl_sau,
            'Phiếu loại trừ': num_excluded,
            'Tỷ lệ HL % (Thô)': tyle_tho,
            'Tỷ lệ HL % (Sau GT)': tyle_sau,
            'Thay đổi %': round(tyle_sau - tyle_tho, 2)
        }])
        
        # Lấy danh sách phiếu bị loại trừ (dùng df_ks vì đã có cột BAOHONG_ID_STR)
        df_loai_tru = df_ks[df_ks['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        cols_to_keep = ['BAOHONG_ID', 'MA_TB', 'TEN_NVKT_DB', 'NGUOI_TL', 'DO_HL', 'KHL_KT', 'NGAY_HOI']
        cols_available = [c for c in cols_to_keep if c in df_loai_tru.columns]
        df_loai_tru = df_loai_tru[cols_available]
        
        # Tính thống kê theo đơn vị (Tổ) cho C1.4
        df_unit_stats = calculate_unit_stats(
            df_before, df_after,
            tong_col='Tổng phiếu KS',
            dat_col='Số phiếu KHL'
        )
        # Đổi tên cột cho phù hợp với C1.4
        df_unit_stats = df_unit_stats.rename(columns={
            'Phiếu đạt (Thô)': 'Phiếu KHL (Thô)',
            'Phiếu đạt (Sau GT)': 'Phiếu KHL (Sau GT)',
            'Tỷ lệ % (Thô)': 'Tỷ lệ KHL % (Thô)',
            'Tỷ lệ % (Sau GT)': 'Tỷ lệ KHL % (Sau GT)'
        })
        # Tính tỷ lệ HL (ngược lại với KHL)
        df_unit_stats['Tỷ lệ HL % (Thô)'] = (100 - df_unit_stats['Tỷ lệ KHL % (Thô)']).round(2)
        df_unit_stats['Tỷ lệ HL % (Sau GT)'] = (100 - df_unit_stats['Tỷ lệ KHL % (Sau GT)']).round(2)
        df_unit_stats['Thay đổi HL %'] = (df_unit_stats['Tỷ lệ HL % (Sau GT)'] - df_unit_stats['Tỷ lệ HL % (Thô)']).round(2)
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C14.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_unit_stats.to_excel(writer, sheet_name='Thong_ke_theo_don_vi', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"\n✅ Đã tạo báo cáo so sánh C1.4 - Độ hài lòng khách hàng")
        print(f"   - Tổng phiếu KS: {tong_ks_tho} -> {tong_ks_sau} (loại trừ: {num_excluded})")
        print(f"   - Phiếu KHL: {tong_khl_tho} -> {tong_khl_sau}")
        print(f"   - Tỷ lệ HL: {tyle_tho}% -> {tyle_sau}% (Δ: {round(tyle_sau - tyle_tho, 2)}%)")
        
        # Bổ sung cột Điểm BSC vào file vừa tạo
        add_bsc_scores_to_c14(output_file)
        
        return {
            'chi_tieu': 'C1.4 Độ hài lòng KH',
            'tong_tho': tong_ks_tho,
            'loai_tru': num_excluded,
            'tong_sau_gt': tong_ks_sau,
            'tyle_tho': tyle_tho,
            'tyle_sau_gt': tyle_sau
        }
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo C1.4: {e}")
        import traceback
        traceback.print_exc()
        return None


def load_c15_exclusion_list(exclusion_file="du_lieu_tham_chieu/LOAI_TRU_C1.5.xlsx"):
    """
    Đọc danh sách HDTB_ID cần loại trừ cho C1.5 từ file Excel
    
    Returns:
        set: Tập hợp các HDTB_ID cần loại trừ (dạng string)
    """
    try:
        if not os.path.exists(exclusion_file):
            print(f"⚠️ Không tìm thấy file loại trừ C1.5: {exclusion_file}")
            return set()
        
        df = pd.read_excel(exclusion_file)
        
        if 'HDTB_ID' not in df.columns:
            print(f"⚠️ Không tìm thấy cột 'HDTB_ID' trong file {exclusion_file}")
            return set()
        
        # Chuẩn hóa tất cả ID
        exclusion_ids = {normalize_id(idx) for idx in df['HDTB_ID'].tolist() if pd.notna(idx)}
        exclusion_ids.discard("")
        
        print(f"✅ Đã đọc {len(exclusion_ids)} mã HDTB_ID loại trừ C1.5")
        return exclusion_ids
        
    except Exception as e:
        print(f"❌ Lỗi khi đọc file loại trừ C1.5: {e}")
        return set()


def create_c15_comparison_report(exclusion_ids, output_dir):
    """
    Tạo báo cáo so sánh C1.5 - Tỷ lệ thiết lập dịch vụ BRCĐ đạt thời gian quy định trước/sau giảm trừ
    
    Công thức: Tỷ lệ đạt = Phiếu đạt / Tổng Hoàn công * 100
    
    Phiếu đạt: CHITIEU = 'PT1' (đạt thời gian quy định)
    Phiếu không đạt: CHITIEU = 'PT2' (không đạt thời gian quy định)
    
    Args:
        exclusion_ids: Set các HDTB_ID cần loại trừ
        output_dir: Thư mục xuất kết quả
        
    Returns:
        dict: Kết quả so sánh hoặc None nếu lỗi
    """
    try:
        print("\n" + "="*80)
        print("TẠO BÁO CÁO SO SÁNH C1.5 - TỶ LỆ THIẾT LẬP DỊCH VỤ BRCĐ ĐẠT THỜI GIAN QUY ĐỊNH")
        print("="*80)
        
        # Đọc dữ liệu chi tiết C1.5 từ sheet DATA
        data_file = os.path.join("downloads", "baocao_hanoi", "c1.5_chitiet_report.xlsx")
        
        if not os.path.exists(data_file):
            print(f"❌ Không tìm thấy file dữ liệu C1.5: {data_file}")
            return None
        
        df_raw = pd.read_excel(data_file, sheet_name="DATA")
        print(f"✅ Đã đọc file dữ liệu: {len(df_raw)} bản ghi tổng")
        
        # Kiểm tra cột cần thiết
        required_cols = ['HDTB_ID', 'TEN_NVKT', 'DOIVT', 'NGAY_HC', 'PT2_KR16']
        missing_cols = [col for col in required_cols if col not in df_raw.columns]
        if missing_cols:
            print(f"❌ Không tìm thấy các cột: {missing_cols}")
            return None
        
        # Lọc chỉ các phiếu đã hoàn công (NGAY_HC không null)
        df_hc = df_raw[df_raw['NGAY_HC'].notna()].copy()
        print(f"✅ Số phiếu đã hoàn công (NGAY_HC != null): {len(df_hc)}")
        
        # Chuẩn hóa cột NVKT
        df_hc['NVKT'] = df_hc['TEN_NVKT'].apply(extract_nvkt_name)
        df_hc = df_hc[df_hc['NVKT'].notna()].copy()
        df_hc['NVKT'] = df_hc['NVKT'].str.strip().str.title()
        
        # Xác định đội
        df_hc['TEN_DOI'] = df_hc['DOIVT']
        has_ten_doi = 'TEN_DOI' in df_hc.columns
        
        # Xác định phiếu đạt: PT2_KR16 != 1 (phiếu không bị trễ)
        # PT2_KR16 = 1 nghĩa là không đạt thời gian quy định (trễ)
        df_hc['IS_DAT'] = (df_hc['PT2_KR16'] != 1).astype(int)
        
        # Áp dụng giảm trừ
        df_hc['HDTB_ID_STR'] = df_hc['HDTB_ID'].apply(normalize_id)
        df_excluded = df_hc[~df_hc['HDTB_ID_STR'].isin(exclusion_ids)].copy()
        num_excluded = len(df_hc) - len(df_excluded)
        print(f"✅ Loại trừ: {num_excluded} phiếu, còn lại {len(df_excluded)} phiếu")
        
        # ========== TÍNH TOÁN TRƯỚC GIẢM TRỪ ==========
        if has_ten_doi:
            df_before = df_hc.groupby(['TEN_DOI', 'NVKT']).agg({
                'IS_DAT': 'sum',       # Tổng phiếu đạt
                'HDTB_ID': 'size'      # Tổng phiếu hoàn công
            }).reset_index()
        else:
            df_before = df_hc.groupby('NVKT').agg({
                'IS_DAT': 'sum',
                'HDTB_ID': 'size'
            }).reset_index()
            
        df_before.columns = list(df_before.columns[:-2]) + ['Phiếu đạt', 'Tổng Hoàn công']
        df_before['Phiếu không đạt'] = df_before['Tổng Hoàn công'] - df_before['Phiếu đạt']
        df_before['Tỷ lệ đạt (%)'] = (df_before['Phiếu đạt'] / 
                                      df_before['Tổng Hoàn công'].replace(0, 1) * 100).round(2)
        
        # ========== TÍNH TOÁN SAU GIẢM TRỪ - LOGIC: CHUYỂN PHIẾU KHÔNG ĐẠT THÀNH ĐẠT ==========
        # Giữ nguyên mẫu số, tăng tử số bằng số phiếu KHÔNG ĐẠT bị loại trừ
        
        # Xác định các phiếu bị loại trừ
        df_loai_tru = df_hc[df_hc['HDTB_ID_STR'].isin(exclusion_ids)].copy()
        
        # Đếm số phiếu KHÔNG ĐẠT bị loại trừ theo nhóm (sẽ chuyển thành đạt)
        if has_ten_doi:
            df_not_dat_excluded = df_loai_tru[df_loai_tru['IS_DAT'] == 0].groupby(['TEN_DOI', 'NVKT']).size().reset_index(name='Phiếu KĐ loại trừ')
        else:
            df_not_dat_excluded = df_loai_tru[df_loai_tru['IS_DAT'] == 0].groupby('NVKT').size().reset_index(name='Phiếu KĐ loại trừ')
        
        # Merge để tính số phiếu đạt SAU giảm trừ = Phiếu đạt TRƯỚC + Phiếu KĐ loại trừ
        merge_cols = ['TEN_DOI', 'NVKT'] if has_ten_doi else ['NVKT']
        df_after = pd.merge(
            df_before[merge_cols + ['Phiếu đạt', 'Tổng Hoàn công']],
            df_not_dat_excluded,
            on=merge_cols,
            how='left'
        )
        df_after['Phiếu KĐ loại trừ'] = df_after['Phiếu KĐ loại trừ'].fillna(0).astype(int)
        
        # TÍNH: Phiếu đạt SAU = Phiếu đạt TRƯỚC + Phiếu KHÔNG ĐẠT bị loại trừ (chuyển thành đạt)
        df_after['Phiếu đạt'] = df_after['Phiếu đạt'] + df_after['Phiếu KĐ loại trừ']
        # Đảm bảo không vượt quá tổng
        df_after['Phiếu đạt'] = df_after[['Phiếu đạt', 'Tổng Hoàn công']].min(axis=1)

        # Tính lại phiếu không đạt và tỷ lệ với MẪU SỐ GIỮ NGUYÊN
        df_after['Phiếu không đạt'] = df_after['Tổng Hoàn công'] - df_after['Phiếu đạt']
        df_after['Tỷ lệ đạt (%)'] = (df_after['Phiếu đạt'] /
                                     df_after['Tổng Hoàn công'].replace(0, 1) * 100).round(2)
        
        # Xóa cột tạm
        df_after = df_after.drop(columns=['Phiếu KĐ loại trừ'])
        
        # ========== MERGE VÀ SO SÁNH ==========
        merge_cols = ['TEN_DOI', 'NVKT'] if has_ten_doi else ['NVKT']
        
        df_comparison = pd.merge(
            df_before, df_after,
            on=merge_cols,
            how='outer',
            suffixes=(' (Thô)', ' (Sau GT)')
        )
        
        # Điền giá trị mặc định
        for col in df_comparison.columns:
            if 'Phiếu' in col or 'Hoàn công' in col:
                df_comparison[col] = df_comparison[col].fillna(0).astype(int)
            elif 'Tỷ lệ' in col:
                df_comparison[col] = df_comparison[col].fillna(100.0)
        
        # Tính chênh lệch
        df_comparison['Chênh lệch %'] = (
            df_comparison['Tỷ lệ đạt (%) (Sau GT)'].fillna(100) - 
            df_comparison['Tỷ lệ đạt (%) (Thô)'].fillna(100)
        ).round(2)
        
        # Sắp xếp cột
        if has_ten_doi:
            column_order = [
                'TEN_DOI', 'NVKT', 
                'Tổng Hoàn công (Thô)', 'Phiếu đạt (Thô)', 'Phiếu không đạt (Thô)', 'Tỷ lệ đạt (%) (Thô)',
                'Tổng Hoàn công (Sau GT)', 'Phiếu đạt (Sau GT)', 'Phiếu không đạt (Sau GT)', 'Tỷ lệ đạt (%) (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            column_order = [
                'NVKT', 
                'Tổng Hoàn công (Thô)', 'Phiếu đạt (Thô)', 'Phiếu không đạt (Thô)', 'Tỷ lệ đạt (%) (Thô)',
                'Tổng Hoàn công (Sau GT)', 'Phiếu đạt (Sau GT)', 'Phiếu không đạt (Sau GT)', 'Tỷ lệ đạt (%) (Sau GT)',
                'Chênh lệch %'
            ]
            df_comparison = df_comparison.sort_values('NVKT').reset_index(drop=True)
        
        column_order = [c for c in column_order if c in df_comparison.columns]
        df_comparison = df_comparison[column_order]
        
        # Loại bỏ các dòng có NVKT rỗng
        df_comparison = df_comparison[df_comparison['NVKT'].notna()]
        
        # ========== TÍNH THỐNG KÊ TỔNG HỢP ==========
        tong_hc_tho = df_before['Tổng Hoàn công'].sum()
        tong_dat_tho = df_before['Phiếu đạt'].sum()
        tyle_tho = round(tong_dat_tho / tong_hc_tho * 100, 2) if tong_hc_tho > 0 else 100
        
        tong_hc_sau = df_after['Tổng Hoàn công'].sum()
        tong_dat_sau = df_after['Phiếu đạt'].sum()
        tyle_sau = round(tong_dat_sau / tong_hc_sau * 100, 2) if tong_hc_sau > 0 else 100
        
        df_tongke = pd.DataFrame([{
            'Chỉ tiêu': 'C1.5 - Tỷ lệ thiết lập dịch vụ BRCĐ đạt TG quy định',
            'Tổng Hoàn công (Thô)': tong_hc_tho,
            'Phiếu đạt (Thô)': tong_dat_tho,
            'Tổng Hoàn công (Sau GT)': tong_hc_sau,
            'Phiếu đạt (Sau GT)': tong_dat_sau,
            'Phiếu loại trừ': num_excluded,
            'Tỷ lệ đạt % (Thô)': tyle_tho,
            'Tỷ lệ đạt % (Sau GT)': tyle_sau,
            'Thay đổi %': round(tyle_sau - tyle_tho, 2)
        }])
        
        # Lấy danh sách phiếu bị loại trừ
        df_loai_tru = df_hc[df_hc['HDTB_ID_STR'].isin(exclusion_ids)].copy()
        cols_to_keep = ['HDTB_ID', 'MA_TB', 'TEN_NVKT', 'DOIVT', 'PT2_KR16', 'NGAY_HC', 'NGAY_BT']
        cols_available = [c for c in cols_to_keep if c in df_loai_tru.columns]
        df_loai_tru = df_loai_tru[cols_available]
        
        # Tính thống kê theo đơn vị (Tổ) cho C1.5
        df_unit_stats = calculate_unit_stats(
            df_before, df_after,
            tong_col='Tổng Hoàn công',
            dat_col='Phiếu đạt'
        )
        # Đổi tên cột cho phù hợp với C1.5
        df_unit_stats = df_unit_stats.rename(columns={
            'Phiếu đạt (Thô)': 'Phiếu đạt (Thô)',
            'Phiếu đạt (Sau GT)': 'Phiếu đạt (Sau GT)',
            'Tỷ lệ % (Thô)': 'Tỷ lệ đạt % (Thô)',
            'Tỷ lệ % (Sau GT)': 'Tỷ lệ đạt % (Sau GT)'
        })
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C15.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_unit_stats.to_excel(writer, sheet_name='Thong_ke_theo_don_vi', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"\n✅ Đã tạo báo cáo so sánh C1.5 - Tỷ lệ thiết lập dịch vụ BRCĐ")
        print(f"   - Tổng Hoàn công: {tong_hc_tho} -> {tong_hc_sau} (loại trừ: {num_excluded})")
        print(f"   - Phiếu đạt: {tong_dat_tho} -> {tong_dat_sau}")
        print(f"   - Tỷ lệ đạt: {tyle_tho}% -> {tyle_sau}% (Δ: {round(tyle_sau - tyle_tho, 2)}%)")
        
        # Bổ sung cột Điểm BSC vào file vừa tạo
        add_bsc_scores_to_c15(output_file)
        
        return {
            'chi_tieu': 'C1.5 Tỷ lệ thiết lập BRCĐ',
            'tong_tho': tong_hc_tho,
            'loai_tru': num_excluded,
            'tong_sau_gt': tong_hc_sau,
            'tyle_tho': tyle_tho,
            'tyle_sau_gt': tyle_sau
        }
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo C1.5: {e}")
        import traceback
        traceback.print_exc()
        return None


def create_unit_bsc_comparison_report(exclusion_folder, original_data_folder, output_folder):
    """
    Tạo báo cáo so sánh điểm BSC cấp đơn vị (Tổ) trước/sau giảm trừ
    """
    try:
        print("\n" + "="*80)
        print("TÍNH ĐIỂM BSC CẤP ĐƠN VỊ TRƯỚC VÀ SAU GIẢM TRỪ")
        print("="*80)
        
        # 1. Lấy dữ liệu điểm NVKT trước/sau giảm trừ (cho sheet Chi_tiet_Ca_nhan)
        print(">> Đang tính điểm Trước giảm trừ...")
        df_truoc = tinh_diem_kpi_nvkt(original_data_folder, None)
        
        print(">> Đang tính điểm Sau giảm trừ...")
        df_sau = tinh_diem_kpi_nvkt_sau_giam_tru(exclusion_folder, original_data_folder, None)
        
        # 2. Tính điểm CẤP ĐƠN VỊ từ các file Thong_ke_theo_don_vi
        print(">> Tính điểm đơn vị từ số liệu gộp (Thong_ke_theo_don_vi)...")
        
        def read_unit_scores_from_comparison_files(folder):
            """Tính điểm đơn vị từ các sheet Thong_ke_theo_don_vi của file so sánh."""
            import os
            
            # Đọc điểm C1.1 TP1 (SM2)
            try:
                df_tp1 = pd.read_excel(os.path.join(folder, 'So_sanh_C11_SM2.xlsx'), sheet_name='Thong_ke_theo_don_vi')
                df_tp1 = df_tp1[df_tp1['Đơn vị'] != 'TTVT Sơn Tây'].copy()  # Loại bỏ dòng TTVT (sẽ tính riêng)
                df_tp1 = df_tp1.rename(columns={'Đơn vị': 'don_vi'})
                df_tp1['c11_tp1_tho'] = df_tp1['Điểm BSC (Thô)']
                df_tp1['c11_tp1_sau'] = df_tp1['Điểm BSC (Sau GT)']
                df_tp1 = df_tp1[['don_vi', 'c11_tp1_tho', 'c11_tp1_sau']]
            except Exception as e:
                print(f"  ⚠️ Lỗi đọc C1.1 TP1: {e}")
                df_tp1 = pd.DataFrame(columns=['don_vi', 'c11_tp1_tho', 'c11_tp1_sau'])
            
            # Đọc điểm C1.1 TP2 (SM4)
            try:
                df_tp2 = pd.read_excel(os.path.join(folder, 'So_sanh_C11_SM4.xlsx'), sheet_name='Thong_ke_theo_don_vi')
                df_tp2 = df_tp2[df_tp2['Đơn vị'] != 'TTVT Sơn Tây'].copy()
                df_tp2 = df_tp2.rename(columns={'Đơn vị': 'don_vi'})
                df_tp2['c11_tp2_tho'] = df_tp2['Điểm BSC (Thô)']
                df_tp2['c11_tp2_sau'] = df_tp2['Điểm BSC (Sau GT)']
                df_tp2 = df_tp2[['don_vi', 'c11_tp2_tho', 'c11_tp2_sau']]
            except Exception as e:
                print(f"  ⚠️ Lỗi đọc C1.1 TP2: {e}")
                df_tp2 = pd.DataFrame(columns=['don_vi', 'c11_tp2_tho', 'c11_tp2_sau'])
            
            # Đọc điểm C1.2 TP1 (SM1 - Hỏng lặp lại)
            try:
                df_c12_tp1 = pd.read_excel(os.path.join(folder, 'So_sanh_C12_SM1.xlsx'), sheet_name='Thong_ke_theo_don_vi')
                df_c12_tp1 = df_c12_tp1[df_c12_tp1['Đơn vị'] != 'TTVT Sơn Tây'].copy()
                df_c12_tp1 = df_c12_tp1.rename(columns={'Đơn vị': 'don_vi'})
                df_c12_tp1['c12_tp1_tho'] = df_c12_tp1['Điểm BSC (Thô)']
                df_c12_tp1['c12_tp1_sau'] = df_c12_tp1['Điểm BSC (Sau GT)']
                df_c12_tp1 = df_c12_tp1[['don_vi', 'c12_tp1_tho', 'c12_tp1_sau']]
            except Exception as e:
                print(f"  ⚠️ Lỗi đọc C1.2 TP1: {e}")
                df_c12_tp1 = pd.DataFrame(columns=['don_vi', 'c12_tp1_tho', 'c12_tp1_sau'])
            
            # Đọc điểm C1.2 TP2 (SM4-C12 - Tỷ lệ sự cố)
            try:
                df_c12_tp2 = pd.read_excel(os.path.join(folder, 'SM4-C12-ti-le-su-co-dv-brcd.xlsx'), sheet_name='Thong_ke_theo_don_vi')
                df_c12_tp2 = df_c12_tp2[df_c12_tp2['Đơn vị'] != 'TTVT Sơn Tây'].copy()
                df_c12_tp2 = df_c12_tp2.rename(columns={'Đơn vị': 'don_vi'})
                df_c12_tp2['c12_tp2_tho'] = df_c12_tp2['Điểm BSC (Thô)']
                df_c12_tp2['c12_tp2_sau'] = df_c12_tp2['Điểm BSC (Sau GT)']
                df_c12_tp2 = df_c12_tp2[['don_vi', 'c12_tp2_tho', 'c12_tp2_sau']]
            except Exception as e:
                print(f"  ⚠️ Lỗi đọc C1.2 TP2: {e}")
                df_c12_tp2 = pd.DataFrame(columns=['don_vi', 'c12_tp2_tho', 'c12_tp2_sau'])
            
            # Đọc điểm C1.4
            try:
                df_c14 = pd.read_excel(os.path.join(folder, 'So_sanh_C14.xlsx'), sheet_name='Thong_ke_theo_don_vi')
                df_c14 = df_c14[df_c14['Đơn vị'] != 'TTVT Sơn Tây'].copy()
                df_c14 = df_c14.rename(columns={'Đơn vị': 'don_vi'})
                df_c14['c14_tho'] = df_c14['Điểm BSC (Thô)']
                df_c14['c14_sau'] = df_c14['Điểm BSC (Sau GT)']
                df_c14 = df_c14[['don_vi', 'c14_tho', 'c14_sau']]
            except Exception as e:
                print(f"  ⚠️ Lỗi đọc C1.4: {e}")
                df_c14 = pd.DataFrame(columns=['don_vi', 'c14_tho', 'c14_sau'])
            
            # Đọc điểm C1.5
            try:
                df_c15 = pd.read_excel(os.path.join(folder, 'So_sanh_C15.xlsx'), sheet_name='Thong_ke_theo_don_vi')
                df_c15 = df_c15[df_c15['Đơn vị'] != 'TTVT Sơn Tây'].copy()
                df_c15 = df_c15.rename(columns={'Đơn vị': 'don_vi'})
                df_c15['c15_tho'] = df_c15['Điểm BSC (Thô)']
                df_c15['c15_sau'] = df_c15['Điểm BSC (Sau GT)']
                df_c15 = df_c15[['don_vi', 'c15_tho', 'c15_sau']]
            except Exception as e:
                print(f"  ⚠️ Lỗi đọc C1.5: {e}")
                df_c15 = pd.DataFrame(columns=['don_vi', 'c15_tho', 'c15_sau'])
            
            # Merge tất cả
            df_all = df_tp1.copy()
            for df_merge in [df_tp2, df_c12_tp1, df_c12_tp2, df_c14, df_c15]:
                if not df_merge.empty:
                    df_all = df_all.merge(df_merge, on='don_vi', how='outer')
            
            # Tính điểm tổng hợp C1.1 và C1.2
            df_all['Diem_C1.1 (Trước)'] = df_all['c11_tp1_tho'].fillna(5) * 0.3 + df_all['c11_tp2_tho'].fillna(5) * 0.7
            df_all['Diem_C1.1 (Sau)'] = df_all['c11_tp1_sau'].fillna(5) * 0.3 + df_all['c11_tp2_sau'].fillna(5) * 0.7
            df_all['Diem_C1.1 (CL)'] = (df_all['Diem_C1.1 (Sau)'] - df_all['Diem_C1.1 (Trước)']).round(2)
            
            # C1.2: Tính từ 2 thành phần TP1 (Hỏng lặp lại) và TP2 (Tỷ lệ sự cố)
            df_all['Diem_C1.2 (Trước)'] = df_all['c12_tp1_tho'].fillna(5) * 0.5 + df_all['c12_tp2_tho'].fillna(5) * 0.5
            df_all['Diem_C1.2 (Sau)'] = df_all['c12_tp1_sau'].fillna(5) * 0.5 + df_all['c12_tp2_sau'].fillna(5) * 0.5
            df_all['Diem_C1.2 (CL)'] = (df_all['Diem_C1.2 (Sau)'] - df_all['Diem_C1.2 (Trước)']).round(2)
            
            df_all['Diem_C1.4 (Trước)'] = df_all['c14_tho'].fillna(5)
            df_all['Diem_C1.4 (Sau)'] = df_all['c14_sau'].fillna(5)
            df_all['Diem_C1.4 (CL)'] = (df_all['Diem_C1.4 (Sau)'] - df_all['Diem_C1.4 (Trước)']).round(2)
            
            # Xử lý C1.5 - kiểm tra cột có tồn tại không (do file So_sanh_C15.xlsx có thể chưa được tạo)
            if 'c15_tho' in df_all.columns:
                df_all['Diem_C1.5 (Trước)'] = df_all['c15_tho'].fillna(5)
            else:
                df_all['Diem_C1.5 (Trước)'] = 5.0  # Mặc định điểm 5 nếu không có dữ liệu
            
            if 'c15_sau' in df_all.columns:
                df_all['Diem_C1.5 (Sau)'] = df_all['c15_sau'].fillna(5)
            else:
                df_all['Diem_C1.5 (Sau)'] = 5.0
            
            df_all['Diem_C1.5 (CL)'] = (df_all['Diem_C1.5 (Sau)'] - df_all['Diem_C1.5 (Trước)']).round(2)
            
            # Làm tròn
            for col in ['Diem_C1.1 (Trước)', 'Diem_C1.1 (Sau)', 'Diem_C1.2 (Trước)', 'Diem_C1.2 (Sau)',
                        'Diem_C1.4 (Trước)', 'Diem_C1.4 (Sau)', 'Diem_C1.5 (Trước)', 'Diem_C1.5 (Sau)']:
                df_all[col] = df_all[col].round(2)
            
            # Chọn cột kết quả
            result_cols = ['don_vi', 
                          'Diem_C1.1 (Trước)', 'Diem_C1.1 (Sau)', 'Diem_C1.1 (CL)',
                          'Diem_C1.2 (Trước)', 'Diem_C1.2 (Sau)', 'Diem_C1.2 (CL)',
                          'Diem_C1.4 (Trước)', 'Diem_C1.4 (Sau)', 'Diem_C1.4 (CL)',
                          'Diem_C1.5 (Trước)', 'Diem_C1.5 (Sau)', 'Diem_C1.5 (CL)']
            return df_all[[c for c in result_cols if c in df_all.columns]]
        
        # Tính điểm đơn vị từ file so sánh
        df_compare_unit = read_unit_scores_from_comparison_files(exclusion_folder)
        score_cols = ['Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4', 'Diem_C1.5']
        
        # Thêm dòng tổng hợp TTVT Sơn Tây (tính từ số liệu sheet Thong_ke_tong_hop)
        try:
            def calculate_ttvt_from_comparison_files(exclusion_folder):
                """
                Tính điểm TTVT từ dữ liệu Thong_ke_tong_hop của các file so sánh.
                Trả về dict chứa điểm Trước và Sau.
                """
                import os
                
                # Đọc từng file và lấy số liệu tổng hợp
                # C1.1 TP1 (SM2)
                try:
                    df_c11_tp1 = pd.read_excel(os.path.join(exclusion_folder, 'So_sanh_C11_SM2.xlsx'), sheet_name='Thong_ke_tong_hop')
                    c11_tp1_tong_tho = df_c11_tp1['Tổng phiếu (Thô)'].iloc[0]
                    c11_tp1_dat_tho = df_c11_tp1['Phiếu đạt (Thô)'].iloc[0]
                    c11_tp1_tong_sau = df_c11_tp1['Tổng phiếu (Sau GT)'].iloc[0]
                    c11_tp1_dat_sau = df_c11_tp1['Phiếu đạt (Sau GT)'].iloc[0]
                except:
                    c11_tp1_tong_tho, c11_tp1_dat_tho = 1, 1
                    c11_tp1_tong_sau, c11_tp1_dat_sau = 1, 1
                
                # C1.1 TP2 (SM4)
                try:
                    df_c11_tp2 = pd.read_excel(os.path.join(exclusion_folder, 'So_sanh_C11_SM4.xlsx'), sheet_name='Thong_ke_tong_hop')
                    c11_tp2_tong_tho = df_c11_tp2['Tổng phiếu (Thô)'].iloc[0]
                    c11_tp2_dat_tho = df_c11_tp2['Phiếu đạt (Thô)'].iloc[0]
                    c11_tp2_tong_sau = df_c11_tp2['Tổng phiếu (Sau GT)'].iloc[0]
                    c11_tp2_dat_sau = df_c11_tp2['Phiếu đạt (Sau GT)'].iloc[0]
                except:
                    c11_tp2_tong_tho, c11_tp2_dat_tho = 1, 1
                    c11_tp2_tong_sau, c11_tp2_dat_sau = 1, 1
                
                # C1.2 TP1 (SM1 - Hỏng lặp lại)
                try:
                    df_c12_tp1 = pd.read_excel(os.path.join(exclusion_folder, 'So_sanh_C12_SM1.xlsx'), sheet_name='Thong_ke_tong_hop')
                    c12_tp1_hll_tho = df_c12_tp1['Phiếu HLL (Thô)'].iloc[0]
                    c12_tp1_bh_tho = df_c12_tp1['Phiếu báo hỏng (Thô)'].iloc[0]
                    c12_tp1_hll_sau = df_c12_tp1['Phiếu HLL (Sau GT)'].iloc[0]
                    c12_tp1_bh_sau = df_c12_tp1['Phiếu báo hỏng (Sau GT)'].iloc[0]
                except:
                    c12_tp1_hll_tho, c12_tp1_bh_tho = 0, 1
                    c12_tp1_hll_sau, c12_tp1_bh_sau = 0, 1
                
                # C1.2 TP2 (SM4 - Sự cố BRCĐ)
                try:
                    df_c12_tp2 = pd.read_excel(os.path.join(exclusion_folder, 'SM4-C12-ti-le-su-co-dv-brcd.xlsx'), sheet_name='Thong_ke_tong_hop')
                    c12_tp2_tb = df_c12_tp2['Tổng thuê bao'].iloc[0]
                    c12_tp2_bh_tho = df_c12_tp2['Phiếu báo hỏng (Thô)'].iloc[0]
                    c12_tp2_bh_sau = df_c12_tp2['Phiếu báo hỏng (Sau GT)'].iloc[0]
                except:
                    c12_tp2_tb, c12_tp2_bh_tho, c12_tp2_bh_sau = 1, 0, 0
                
                # C1.4 (Độ hài lòng)
                try:
                    df_c14 = pd.read_excel(os.path.join(exclusion_folder, 'So_sanh_C14.xlsx'), sheet_name='Thong_ke_tong_hop')
                    c14_ks_tho = df_c14['Tổng phiếu KS (Thô)'].iloc[0]
                    c14_khl_tho = df_c14['Tổng phiếu KHL (Thô)'].iloc[0]
                    c14_ks_sau = df_c14['Tổng phiếu KS (Sau GT)'].iloc[0]
                    c14_khl_sau = df_c14['Tổng phiếu KHL (Sau GT)'].iloc[0]
                except:
                    c14_ks_tho, c14_khl_tho = 1, 0
                    c14_ks_sau, c14_khl_sau = 1, 0
                
                # C1.5 (Thiết lập đúng hạn)
                try:
                    df_c15 = pd.read_excel(os.path.join(exclusion_folder, 'So_sanh_C15.xlsx'), sheet_name='Thong_ke_tong_hop')
                    c15_tong_tho = df_c15['Tổng Hoàn công (Thô)'].iloc[0]
                    c15_dat_tho = df_c15['Phiếu đạt (Thô)'].iloc[0]
                    c15_tong_sau = df_c15['Tổng Hoàn công (Sau GT)'].iloc[0]
                    c15_dat_sau = df_c15['Phiếu đạt (Sau GT)'].iloc[0]
                except:
                    c15_tong_tho, c15_dat_tho = 1, 1
                    c15_tong_sau, c15_dat_sau = 1, 1
                
                # Tính tỷ lệ và điểm - TRƯỚC
                r_c11_tp1_tho = c11_tp1_dat_tho / c11_tp1_tong_tho if c11_tp1_tong_tho > 0 else 1.0
                r_c11_tp2_tho = c11_tp2_dat_tho / c11_tp2_tong_tho if c11_tp2_tong_tho > 0 else 1.0
                d_c11_tho = tinh_diem_C11_TP1(r_c11_tp1_tho) * 0.3 + tinh_diem_C11_TP2(r_c11_tp2_tho) * 0.7
                
                r_c12_tp1_tho = c12_tp1_hll_tho / c12_tp1_bh_tho if c12_tp1_bh_tho > 0 else 0.0
                r_c12_tp2_tho = c12_tp2_bh_tho / c12_tp2_tb if c12_tp2_tb > 0 else 0.0
                d_c12_tho = tinh_diem_C12_TP1(r_c12_tp1_tho) * 0.5 + tinh_diem_C12_TP2(r_c12_tp2_tho) * 0.5
                
                r_c14_tho = (c14_ks_tho - c14_khl_tho) / c14_ks_tho if c14_ks_tho > 0 else 1.0
                d_c14_tho = tinh_diem_C14(r_c14_tho)
                
                r_c15_tho = c15_dat_tho / c15_tong_tho if c15_tong_tho > 0 else 1.0
                d_c15_tho = tinh_diem_C15(r_c15_tho)
                
                # Tính tỷ lệ và điểm - SAU
                r_c11_tp1_sau = c11_tp1_dat_sau / c11_tp1_tong_sau if c11_tp1_tong_sau > 0 else 1.0
                r_c11_tp2_sau = c11_tp2_dat_sau / c11_tp2_tong_sau if c11_tp2_tong_sau > 0 else 1.0
                d_c11_sau = tinh_diem_C11_TP1(r_c11_tp1_sau) * 0.3 + tinh_diem_C11_TP2(r_c11_tp2_sau) * 0.7
                
                r_c12_tp1_sau = c12_tp1_hll_sau / c12_tp1_bh_sau if c12_tp1_bh_sau > 0 else 0.0
                r_c12_tp2_sau = c12_tp2_bh_sau / c12_tp2_tb if c12_tp2_tb > 0 else 0.0
                d_c12_sau = tinh_diem_C12_TP1(r_c12_tp1_sau) * 0.5 + tinh_diem_C12_TP2(r_c12_tp2_sau) * 0.5
                
                r_c14_sau = (c14_ks_sau - c14_khl_sau) / c14_ks_sau if c14_ks_sau > 0 else 1.0
                d_c14_sau = tinh_diem_C14(r_c14_sau)
                
                r_c15_sau = c15_dat_sau / c15_tong_sau if c15_tong_sau > 0 else 1.0
                d_c15_sau = tinh_diem_C15(r_c15_sau)
                
                return {
                    'don_vi': 'TTVT Sơn Tây',
                    'Diem_C1.1 (Trước)': round(d_c11_tho, 2),
                    'Diem_C1.1 (Sau)': round(d_c11_sau, 2),
                    'Diem_C1.1 (CL)': round(d_c11_sau - d_c11_tho, 2),
                    'Diem_C1.2 (Trước)': round(d_c12_tho, 2),
                    'Diem_C1.2 (Sau)': round(d_c12_sau, 2),
                    'Diem_C1.2 (CL)': round(d_c12_sau - d_c12_tho, 2),
                    'Diem_C1.4 (Trước)': round(d_c14_tho, 2),
                    'Diem_C1.4 (Sau)': round(d_c14_sau, 2),
                    'Diem_C1.4 (CL)': round(d_c14_sau - d_c14_tho, 2),
                    'Diem_C1.5 (Trước)': round(d_c15_tho, 2),
                    'Diem_C1.5 (Sau)': round(d_c15_sau, 2),
                    'Diem_C1.5 (CL)': round(d_c15_sau - d_c15_tho, 2),
                }
            
            # Tính điểm tổng hợp TTVT từ các file so sánh
            ttvt_row = calculate_ttvt_from_comparison_files(exclusion_folder)
            
            # Tạo DF và nối vào
            ttvt_df = pd.DataFrame([ttvt_row])
            # Đảm bảo cột khớp
            ttvt_df = ttvt_df[df_compare_unit.columns]
            
            df_compare_unit = pd.concat([df_compare_unit, ttvt_df], ignore_index=True)
            print("✅ Đã tính điểm tổng hợp TTVT Sơn Tây từ các file so sánh.")
            
        except Exception as ex:
            print(f"⚠️ Lỗi khi tính điểm tổng hợp TTVT: {ex}")
        
        # 4b. Merge và so sánh CẤP CÁ NHÂN
        # Lấy các cột điểm
        cols_ind_truoc = ['don_vi', 'nvkt'] + [c for c in score_cols if c in df_truoc.columns]
        cols_ind_sau = ['don_vi', 'nvkt'] + [c for c in score_cols if c in df_sau.columns]
        
        df_ind_truoc = df_truoc[cols_ind_truoc].copy()
        df_ind_sau = df_sau[cols_ind_sau].copy()
        
        df_compare_ind = pd.merge(
            df_ind_truoc,
            df_ind_sau,
            on=['don_vi', 'nvkt'],
            how='outer',
            suffixes=(' (Trước)', ' (Sau)')
        )
        
        # Sắp xếp và tính chênh lệch CÁ NHÂN
        final_cols_ind = ['don_vi', 'nvkt']
        for col in score_cols:
            col_truoc = f"{col} (Trước)"
            col_sau = f"{col} (Sau)"
            col_diff = f"{col} (CL)"
            
            if col_truoc in df_compare_ind:
                df_compare_ind[col_truoc] = df_compare_ind[col_truoc].fillna(0)
            if col_sau in df_compare_ind:
                df_compare_ind[col_sau] = df_compare_ind[col_sau].fillna(0)
                
            if col_truoc in df_compare_ind and col_sau in df_compare_ind:
                df_compare_ind[col_diff] = (df_compare_ind[col_sau] - df_compare_ind[col_truoc]).round(2)
                final_cols_ind.extend([col_truoc, col_sau, col_diff])
            elif col_truoc in df_compare_ind:
                final_cols_ind.append(col_truoc)
            elif col_sau in df_compare_ind:
                final_cols_ind.append(col_sau)
        
        df_compare_ind = df_compare_ind[final_cols_ind].sort_values(['don_vi', 'nvkt'])
        
        # 5. Xuất file 2 sheet
        output_file = os.path.join(output_folder, "Tong_hop_Diem_BSC_Don_Vi.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_compare_unit.to_excel(writer, sheet_name='Tong_hop_Don_vi', index=False)
            df_compare_ind.to_excel(writer, sheet_name='Chi_tiet_Ca_nhan', index=False)
            
        print(f"\n✅ Đã xuất báo cáo BSC đơn vị (kèm chi tiết cá nhân): {output_file}")
        
        # In preview
        print("\nPreview kết quả Đơn vị (3 dòng đầu):")
        print(df_compare_unit.head(3).to_string())
        print("\nPreview kết quả Cá nhân (3 dòng đầu):")
        print(df_compare_ind.head(3).to_string())
        
        return True
        
    except Exception as e:
        print(f"❌ Lỗi khi tính điểm BSC đơn vị: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_exclusion_reports():
    """
    Hàm chính: Tạo tất cả báo cáo so sánh giảm trừ
    
    Quy trình:
    1. Đọc danh sách BAOHONG_ID cần loại trừ
    2. Tạo thư mục xuất kết quả
    3. Tạo báo cáo so sánh cho từng chỉ tiêu
    4. Tạo báo cáo tổng hợp
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TẠO BÁO CÁO SO SÁNH GIẢM TRỪ")
        print("="*80)
        
        # 1. Đọc danh sách loại trừ
        exclusion_ids = load_exclusion_list()
        
        if not exclusion_ids:
            print("⚠️ Không có phiếu nào cần loại trừ. Bỏ qua việc tạo báo cáo giảm trừ.")
            return False
        
        # 2. Tạo thư mục xuất kết quả
        output_dir = os.path.join("downloads", "kq_sau_giam_tru")
        os.makedirs(output_dir, exist_ok=True)
        print(f"✅ Đã tạo thư mục xuất kết quả: {output_dir}")
        
        # 3. Tạo các báo cáo so sánh
        results = []
        
        # C1.1 SM4
        result = create_c11_comparison_report(exclusion_ids, output_dir)
        if result:
            results.append(result)
        
        # C1.1 SM2 - Sử dụng file loại trừ riêng
        print("\n" + "-"*80)
        print("Đọc danh sách loại trừ riêng cho C1.1 SM2 (TP1)")
        print("-"*80)
        exclusion_ids_c11_sm2 = load_exclusion_list("du_lieu_tham_chieu/LOAI_TRU_C1.1_TP1.xlsx")
        
        if exclusion_ids_c11_sm2:
            print(f"✅ Sử dụng {len(exclusion_ids_c11_sm2)} phiếu loại trừ từ LOAI_TRU_C1.1_TP1.xlsx")
            result = create_c11_sm2_comparison_report(exclusion_ids_c11_sm2, output_dir)
        else:
            print(f"⚠️ Không tìm thấy file LOAI_TRU_C1.1_TP1.xlsx, sử dụng danh sách loại trừ chung")
            result = create_c11_sm2_comparison_report(exclusion_ids, output_dir)
        
        if result:
            results.append(result)
        
        # C1.2 SM1
        result = create_c12_comparison_report(exclusion_ids, output_dir)
        if result:
            results.append(result)
        
        # Tạo file SM1-C12 sau giảm trừ với cấu trúc gốc
        create_sm1_c12_excluded_file(exclusion_ids, output_dir)
        
        # C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng (SM4-C12)
        result = create_c12_ti_le_bao_hong_comparison_report(exclusion_ids, output_dir)
        if result:
            results.append(result)
        
        # C1.4 - Độ hài lòng khách hàng
        c14_exclusion_ids = load_c14_exclusion_list()
        if c14_exclusion_ids:
            result = create_c14_comparison_report(c14_exclusion_ids, output_dir)
            if result:
                results.append(result)
        
        # C1.5 - Tỷ lệ thiết lập dịch vụ BRCĐ đạt thời gian quy định
        # Luôn tạo báo cáo C1.5 (dù không có phiếu loại trừ - dữ liệu Trước/Sau sẽ giống nhau)
        c15_exclusion_ids = load_c15_exclusion_list()
        result = create_c15_comparison_report(c15_exclusion_ids, output_dir)
        if result:
            results.append(result)
        
        # 4. Tạo báo cáo tổng hợp tất cả chỉ tiêu
        if results:
            print("\n" + "="*80)
            print("TẠO BÁO CÁO TỔNG HỢP TẤT CẢ CHỈ TIÊU")
            print("="*80)
            
            df_summary = pd.DataFrame(results)
            df_summary.columns = [
                'Chỉ tiêu', 
                'Tổng phiếu (Thô)', 
                'Phiếu loại trừ', 
                'Tổng phiếu (Sau GT)',
                'Tỷ lệ % (Thô)',
                'Tỷ lệ % (Sau GT)'
            ]
            df_summary['Thay đổi %'] = (df_summary['Tỷ lệ % (Sau GT)'] - df_summary['Tỷ lệ % (Thô)']).round(2)
            
            summary_file = os.path.join(output_dir, "Tong_hop_giam_tru.xlsx")
            df_summary.to_excel(summary_file, index=False)
            
            print(f"✅ Đã tạo báo cáo tổng hợp: {summary_file}")
            print("\n" + "-"*80)
            print("TỔNG KẾT:")
            print("-"*80)
            for _, row in df_summary.iterrows():
                print(f"  {row['Chỉ tiêu']}:")
                print(f"    - Thô: {row['Tỷ lệ % (Thô)']}% -> Sau GT: {row['Tỷ lệ % (Sau GT)']}% (Δ: {row['Thay đổi %']}%)")
            print("-"*80)
        
        # 5. Tính toán KPI cá nhân sau giảm trừ cho NVKT
        print("\n" + "="*80)
        print("TÍNH TOÁN KPI CÁ NHÂN SAU GIẢM TRỪ CHO NVKT")
        print("="*80)
        
        # Thư mục chứa dữ liệu gốc (cho C1.4)
        original_data_folder = "downloads/baocao_hanoi"
        # Thư mục xuất file KPI (Excel)
        kpi_output_dir = os.path.join(output_dir, "KPI_sau_GT")
        
        # Gọi kpi_calculator để tính điểm
        df_kpi_nvkt = kpi_calculator.tinh_diem_kpi_nvkt_sau_giam_tru(
            exclusion_folder=output_dir,
            original_data_folder=original_data_folder,
            output_folder=kpi_output_dir
        )
        
        if df_kpi_nvkt is not None:
            # 6. Tạo báo cáo Word cho từng cá nhân và lưu theo thư mục Đội
            report_generator.generate_all_individual_reports_after_exclusion(
                kpi_folder=kpi_output_dir,
                output_root=output_dir,
                report_month=None # Tự động lấy tháng hiện tại
            )
            
            # 7. Tạo báo cáo so sánh điểm BSC cấp đơn vị
            create_unit_bsc_comparison_report(
                exclusion_folder=output_dir,
                original_data_folder=original_data_folder,
                output_folder=output_dir
            )
        
        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẤT CẢ CÁC BƯỚC XỬ LÝ GIẢM TRỪ")
        print(f"   Dữ liệu so sánh & Báo cáo cá nhân: {output_dir}")
        print("="*80)
        
        return True
        
    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo giảm trừ: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # Chạy độc lập để test
    process_exclusion_reports()
