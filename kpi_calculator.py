"""
Module tính điểm KPI cho NVKT theo BSC Q4/2025 VNPT Hà Nội
Các chỉ tiêu: C1.1, C1.2, C1.4
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime


# ============================================================================
# CÁC HÀM TÍNH ĐIỂM THÀNH PHẦN
# ============================================================================

def tinh_diem_C11_TP1(kq):
    """
    Tính điểm C1.1 Thành phần 1 (30%): Tỷ lệ sửa chữa phiếu chất lượng chủ động
    
    Args:
        kq: Tỷ lệ sửa chữa chủ động (dạng thập phân, vd: 0.98 = 98%)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có phiếu cần sửa = tốt
    
    if kq >= 0.99:
        return 5
    elif kq > 0.96:
        return 1 + 4 * (kq - 0.96) / 0.03
    else:
        return 1


def tinh_diem_C11_TP2(kq):
    """
    Tính điểm C1.1 Thành phần 2 (70%): Tỷ lệ sửa chữa báo hỏng đúng quy định (không tính hẹn)
    
    Args:
        kq: Tỷ lệ sửa chữa báo hỏng đúng quy định (dạng thập phân)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có phiếu báo hỏng = tốt
    
    if kq >= 0.85:
        return 5
    elif kq >= 0.82:
        return 4 + (kq - 0.82) / 0.03
    elif kq >= 0.79:
        return 3 + (kq - 0.79) / 0.03
    elif kq >= 0.76:
        return 2
    else:
        return 1


def tinh_diem_C12_TP1(kq):
    """
    Tính điểm C1.2 Thành phần 1 (50%): Tỷ lệ thuê bao báo hỏng lặp lại
    LƯU Ý: Càng thấp càng tốt
    
    Args:
        kq: Tỷ lệ báo hỏng lặp lại (dạng thập phân)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có hỏng lặp lại = tốt
    
    if kq <= 0.025:
        return 5
    elif kq < 0.04:
        return 5 - 4 * (kq - 0.025) / 0.015
    else:
        return 1


def tinh_diem_C12_TP2(kq):
    """
    Tính điểm C1.2 Thành phần 2 (50%): Tỷ lệ sự cố dịch vụ BRCĐ
    LƯU Ý: Càng thấp càng tốt
    
    Args:
        kq: Tỷ lệ sự cố dịch vụ BRCĐ (dạng thập phân)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return 5  # Không có dữ liệu = không có sự cố = tốt
    
    if kq <= 0.02:
        return 5
    elif kq < 0.03:
        return 5 - 4 * (kq - 0.02) / 0.01
    else:
        return 1


def tinh_diem_C14(kq):
    """
    Tính điểm C1.4: Độ hài lòng của khách hàng sau lắp đặt và sửa chữa
    
    Args:
        kq: Độ hài lòng khách hàng (dạng thập phân)
    
    Returns:
        Điểm từ 1-5
    """
    if pd.isna(kq) or kq is None:
        return np.nan
    
    if kq >= 0.995:
        return 5
    elif kq > 0.95:
        return 1 + 4 * (kq - 0.95) / 0.045
    else:
        return 1


# ============================================================================
# CÁC HÀM ĐỌC DỮ LIỆU TỪ FILE EXCEL
# ============================================================================

def chuan_hoa_ty_le(df, col_ty_le):
    """
    Chuẩn hóa tỷ lệ về dạng thập phân (0-1)
    - Nếu giá trị > 1 thì chia cho 100
    """
    df = df.copy()
    if df[col_ty_le].max() > 1:
        df[col_ty_le] = df[col_ty_le] / 100
    return df


def chuan_hoa_ten(df, col_ten):
    """
    Chuẩn hóa tên NVKT về dạng Title Case
    - Xử lý trường hợp cùng 1 người được nhập với chữ hoa/thường khác nhau
    - Ví dụ: "Bùi văn Cường" -> "Bùi Văn Cường"
    """
    df = df.copy()
    df[col_ten] = df[col_ten].str.strip().str.title()
    return df


def doc_C11_TP1(data_folder):
    """
    Đọc dữ liệu C1.1 Thành phần 1: Tỷ lệ sửa chữa chủ động
    File: SM2-C11.xlsx, Sheet: TH_SM2
    """
    file_path = Path(data_folder) / "SM2-C11.xlsx"
    df = pd.read_excel(file_path, sheet_name="TH_SM2")
    
    # Lấy các cột cần thiết
    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu', 'Phiếu đạt', 'Tỉ lệ đạt (%)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp1_tong_phieu', 'c11_tp1_phieu_dat', 'c11_tp1_ty_le']
    
    # Chuẩn hóa tên NVKT
    df = chuan_hoa_ten(df, 'nvkt')
    
    # Gộp các dòng trùng lặp (cùng đơn vị + tên NVKT)
    df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
        'c11_tp1_tong_phieu': 'sum',
        'c11_tp1_phieu_dat': 'sum',
        'c11_tp1_ty_le': 'first'  # Sẽ tính lại
    })
    # Tính lại tỷ lệ sau khi gộp
    df['c11_tp1_ty_le'] = df['c11_tp1_phieu_dat'] / df['c11_tp1_tong_phieu']
    
    return df


def doc_C11_TP2(data_folder):
    """
    Đọc dữ liệu C1.1 Thành phần 2: Tỷ lệ sửa chữa báo hỏng đúng quy định (không tính hẹn)
    File: SM4-C11.xlsx, Sheet: chi_tiet
    """
    file_path = Path(data_folder) / "SM4-C11.xlsx"
    df = pd.read_excel(file_path, sheet_name="chi_tiet")
    
    # Tên cột thực tế khá dài
    col_ty_le = 'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn'
    
    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu', 'Số phiếu đạt', col_ty_le]].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp2_tong_phieu', 'c11_tp2_phieu_dat', 'c11_tp2_ty_le']
    
    # Chuẩn hóa tên NVKT
    df = chuan_hoa_ten(df, 'nvkt')
    
    # Chuẩn hóa tỷ lệ về dạng thập phân
    df = chuan_hoa_ty_le(df, 'c11_tp2_ty_le')
    
    return df


def doc_C12_TP1(data_folder):
    """
    Đọc dữ liệu C1.2 Thành phần 1: Tỷ lệ thuê bao báo hỏng lặp lại
    File: SM1-C12.xlsx, Sheet: TH_SM1C12_HLL_Thang
    """
    file_path = Path(data_folder) / "SM1-C12.xlsx"
    df = pd.read_excel(file_path, sheet_name="TH_SM1C12_HLL_Thang")
    
    df = df[['TEN_DOI', 'NVKT', 'Số phiếu HLL', 'Số phiếu báo hỏng', 'Tỉ lệ HLL tháng (2.5%)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c12_tp1_phieu_hll', 'c12_tp1_phieu_bh', 'c12_tp1_ty_le']
    
    # Chuẩn hóa tên NVKT
    df = chuan_hoa_ten(df, 'nvkt')
    
    # Chuẩn hóa tỷ lệ về dạng thập phân
    df = chuan_hoa_ty_le(df, 'c12_tp1_ty_le')
    
    return df


def doc_C12_TP2(data_folder):
    """
    Đọc dữ liệu C1.2 Thành phần 2: Tỷ lệ sự cố dịch vụ BRCĐ
    File: SM4-C12-ti-le-su-co-dv-brcd.xlsx, Sheet: TH_C12_TiLeBaoHong
    
    LƯU Ý: Dữ liệu gốc có tỷ lệ dạng 0.37 = 0.37% (không phải 37%)
           Cần chia 100 để chuyển về dạng thập phân: 0.37 / 100 = 0.0037
    """
    file_path = Path(data_folder) / "SM4-C12-ti-le-su-co-dv-brcd.xlsx"
    df = pd.read_excel(file_path, sheet_name="TH_C12_TiLeBaoHong")
    
    df = df[['TEN_DOI', 'NVKT', 'Số phiếu báo hỏng', 'Tổng TB', 'Tỷ lệ báo hỏng (%)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c12_tp2_phieu_bh', 'c12_tp2_tong_tb', 'c12_tp2_ty_le']
    
    # Chuẩn hóa tên NVKT
    df = chuan_hoa_ten(df, 'nvkt')
    
    # Dữ liệu gốc có tỷ lệ dạng 0.37 = 0.37%, cần chia 100 để về dạng thập phân
    # VD: 0.37 / 100 = 0.0037 (tức 0.37%)
    df['c12_tp2_ty_le'] = df['c12_tp2_ty_le'] / 100
    
    return df


def doc_C14(data_folder):
    """
    Đọc dữ liệu C1.4: Độ hài lòng khách hàng
    File: c1.4_chitiet_report.xlsx, Sheet: TH_HL_NVKT
    
    LƯU Ý: Nếu Tổng phiếu KS thành công = 0 (mẫu số = 0), 
           tỷ lệ hài lòng sẽ được mặc định = 100% (1.0)
    """
    file_path = Path(data_folder) / "c1.4_chitiet_report.xlsx"
    df = pd.read_excel(file_path, sheet_name="TH_HL_NVKT")
    
    df = df[['DOIVT', 'NVKT', 'Tổng phiếu KS thành công', 'Tổng phiếu KHL', 'Tỉ lệ HL NVKT (%)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c14_phieu_ks', 'c14_phieu_khl', 'c14_ty_le']
    
    # Chuẩn hóa tên NVKT
    df = chuan_hoa_ten(df, 'nvkt')
    
    # Chuẩn hóa tỷ lệ về dạng thập phân
    df = chuan_hoa_ty_le(df, 'c14_ty_le')
    
    # Xử lý trường hợp đặc biệt: mẫu số = 0 -> tỷ lệ = 100%
    # Nếu chưa có phiếu KS thành công nào, mặc định NVKT được coi là hài lòng 100%
    mask_no_survey = (df['c14_phieu_ks'] == 0) | (df['c14_phieu_ks'].isna())
    df.loc[mask_no_survey, 'c14_ty_le'] = 1.0  # 100% dạng thập phân
    
    return df


# ============================================================================
# CÁC HÀM ĐỌC DỮ LIỆU SAU GIẢM TRỪ
# ============================================================================

def doc_C11_TP1_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.1 Thành phần 1 SAU GIẢM TRỪ: Tỷ lệ sửa chữa chủ động
    File: So_sanh_C11_SM2.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C11_SM2.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")
    
    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu (Sau GT)', 'Số phiếu đạt (Sau GT)', 'Tỷ lệ % (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp1_tong_phieu', 'c11_tp1_phieu_dat', 'c11_tp1_ty_le']
    
    # Chuẩn hóa tỷ lệ về dạng thập phân
    df = chuan_hoa_ty_le(df, 'c11_tp1_ty_le')
    
    return df


def doc_C11_TP2_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.1 Thành phần 2 SAU GIẢM TRỪ: Tỷ lệ sửa chữa báo hỏng đúng quy định
    File: So_sanh_C11_SM4.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C11_SM4.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")
    
    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu (Sau GT)', 'Số phiếu đạt (Sau GT)', 'Tỷ lệ % (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp2_tong_phieu', 'c11_tp2_phieu_dat', 'c11_tp2_ty_le']
    
    # Chuẩn hóa tỷ lệ về dạng thập phân
    df = chuan_hoa_ty_le(df, 'c11_tp2_ty_le')
    
    return df


def doc_C12_TP1_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.2 Thành phần 1 SAU GIẢM TRỪ: Tỷ lệ thuê bao báo hỏng lặp lại
    File: So_sanh_C12_SM1.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C12_SM1.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")
    
    df = df[['TEN_DOI', 'NVKT', 'Số phiếu HLL (Sau GT)', 'Số phiếu báo hỏng (Sau GT)', 'Tỷ lệ HLL % (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c12_tp1_phieu_hll', 'c12_tp1_phieu_bh', 'c12_tp1_ty_le']
    
    # Chuẩn hóa tỷ lệ về dạng thập phân
    df = chuan_hoa_ty_le(df, 'c12_tp1_ty_le')
    
    return df


def doc_C12_TP2_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.2 Thành phần 2 SAU GIẢM TRỪ: Tỷ lệ sự cố dịch vụ BRCĐ
    File: SM4-C12-ti-le-su-co-dv-brcd.xlsx, Sheet: So_sanh_chi_tiet
    
    LƯU Ý: File sau giảm trừ có tỷ lệ dạng 3.16 = 3.16% (khác với file gốc)
           chuan_hoa_ty_le sẽ chia 100 vì max > 1
    """
    file_path = Path(exclusion_folder) / "SM4-C12-ti-le-su-co-dv-brcd.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")
    
    df = df[['TEN_DOI', 'NVKT', 'Tổng TB (Thô)', 'Số phiếu báo hỏng (Sau GT)', 'Tỷ lệ báo hỏng (%) (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c12_tp2_tong_tb', 'c12_tp2_phieu_bh', 'c12_tp2_ty_le']
    
    # Chuẩn hóa tỷ lệ về dạng thập phân (3.16 / 100 = 0.0316)
    df = chuan_hoa_ty_le(df, 'c12_tp2_ty_le')
    
    return df


# ============================================================================
# HÀM TÍNH ĐIỂM KPI TỔNG HỢP
# ============================================================================

def tinh_diem_kpi_nvkt(data_folder, output_folder=None):
    """
    Tính điểm KPI cho tất cả NVKT dựa trên dữ liệu từ các file Excel
    
    Args:
        data_folder: Thư mục chứa các file dữ liệu Excel
        output_folder: Thư mục xuất kết quả (nếu không chỉ định sẽ không xuất file)
    
    Returns:
        DataFrame chứa điểm KPI của từng NVKT
    """
    print(f"Đang đọc dữ liệu từ: {data_folder}")
    
    # 1. Đọc dữ liệu từ các file
    df_c11_tp1 = doc_C11_TP1(data_folder)
    print(f"  - C1.1 TP1: {len(df_c11_tp1)} NVKT")
    
    df_c11_tp2 = doc_C11_TP2(data_folder)
    print(f"  - C1.1 TP2: {len(df_c11_tp2)} NVKT")
    
    df_c12_tp1 = doc_C12_TP1(data_folder)
    print(f"  - C1.2 TP1: {len(df_c12_tp1)} NVKT")
    
    df_c12_tp2 = doc_C12_TP2(data_folder)
    print(f"  - C1.2 TP2: {len(df_c12_tp2)} NVKT")
    
    df_c14 = doc_C14(data_folder)
    print(f"  - C1.4: {len(df_c14)} NVKT")
    
    # 2. Merge tất cả dữ liệu theo nvkt VÀ don_vi
    # (NVKT có thể xuất hiện ở nhiều đơn vị do chuyển đơn vị)
    merge_keys = ['don_vi', 'nvkt']
    
    # Bắt đầu với C1.1 TP1
    df_all = df_c11_tp1.copy()
    
    # Merge C1.1 TP2
    df_all = df_all.merge(
        df_c11_tp2, 
        on=merge_keys, 
        how='outer'
    )
    
    # Merge C1.2 TP1
    df_all = df_all.merge(
        df_c12_tp1, 
        on=merge_keys, 
        how='outer'
    )
    
    # Merge C1.2 TP2
    df_all = df_all.merge(
        df_c12_tp2, 
        on=merge_keys, 
        how='outer'
    )
    
    # Merge C1.4
    df_all = df_all.merge(
        df_c14, 
        on=merge_keys, 
        how='outer'
    )
    
    print(f"\nTổng số NVKT sau merge: {len(df_all)}")
    
    # 3. Tính điểm từng thành phần
    print("Đang tính điểm các thành phần...")
    
    df_all['diem_c11_tp1'] = df_all['c11_tp1_ty_le'].apply(tinh_diem_C11_TP1)
    df_all['diem_c11_tp2'] = df_all['c11_tp2_ty_le'].apply(tinh_diem_C11_TP2)
    df_all['diem_c12_tp1'] = df_all['c12_tp1_ty_le'].apply(tinh_diem_C12_TP1)
    df_all['diem_c12_tp2'] = df_all['c12_tp2_ty_le'].apply(tinh_diem_C12_TP2)
    df_all['diem_c14'] = df_all['c14_ty_le'].apply(tinh_diem_C14)
    
    # 4. Tính điểm tổng hợp
    df_all['Diem_C1.1'] = df_all['diem_c11_tp1'] * 0.30 + df_all['diem_c11_tp2'] * 0.70
    df_all['Diem_C1.2'] = df_all['diem_c12_tp1'] * 0.50 + df_all['diem_c12_tp2'] * 0.50
    df_all['Diem_C1.4'] = df_all['diem_c14']
    
    # 5. Làm tròn điểm (2 chữ số thập phân)
    diem_cols = ['diem_c11_tp1', 'diem_c11_tp2', 'diem_c12_tp1', 'diem_c12_tp2', 'diem_c14',
                 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']
    for col in diem_cols:
        df_all[col] = df_all[col].round(2)
    
    # Làm tròn tỷ lệ về % (2 chữ số thập phân)
    ty_le_cols = ['c11_tp1_ty_le', 'c11_tp2_ty_le', 'c12_tp1_ty_le', 'c12_tp2_ty_le', 'c14_ty_le']
    for col in ty_le_cols:
        df_all[col] = (df_all[col] * 100).round(2)  # Chuyển về %
    
    # 6. Sắp xếp các cột theo thứ tự logic
    col_order = [
        'don_vi', 'nvkt',
        # C1.1 TP1
        'c11_tp1_tong_phieu', 'c11_tp1_phieu_dat', 'c11_tp1_ty_le', 'diem_c11_tp1',
        # C1.1 TP2
        'c11_tp2_tong_phieu', 'c11_tp2_phieu_dat', 'c11_tp2_ty_le', 'diem_c11_tp2',
        # C1.1 Tổng
        'Diem_C1.1',
        # C1.2 TP1
        'c12_tp1_phieu_hll', 'c12_tp1_phieu_bh', 'c12_tp1_ty_le', 'diem_c12_tp1',
        # C1.2 TP2
        'c12_tp2_phieu_bh', 'c12_tp2_tong_tb', 'c12_tp2_ty_le', 'diem_c12_tp2',
        # C1.2 Tổng
        'Diem_C1.2',
        # C1.4
        'c14_phieu_ks', 'c14_phieu_khl', 'c14_ty_le', 'diem_c14', 'Diem_C1.4'
    ]
    
    # Chỉ lấy các cột tồn tại
    existing_cols = [col for col in col_order if col in df_all.columns]
    df_all = df_all[existing_cols]
    
    # Sắp xếp theo đơn vị và NVKT
    df_all = df_all.sort_values(['don_vi', 'nvkt']).reset_index(drop=True)
    
    # 7. Xuất file nếu có output_folder
    if output_folder:
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        
        # File đầy đủ chi tiết (ghi đè mỗi lần chạy)
        full_file = output_folder / "KPI_NVKT_ChiTiet.xlsx"
        df_all.to_excel(full_file, index=False)
        print(f"\nĐã xuất file chi tiết: {full_file}")
        
        # File tóm tắt (ghi đè mỗi lần chạy)
        summary_cols = ['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']
        df_summary = df_all[summary_cols].copy()
        summary_file = output_folder / "KPI_NVKT_TomTat.xlsx"
        df_summary.to_excel(summary_file, index=False)
        print(f"Đã xuất file tóm tắt: {summary_file}")
    
    return df_all


def tao_bao_cao_kpi(data_folder, output_folder, label=""):
    """
    Wrapper function để tạo báo cáo KPI hoàn chỉnh
    
    Args:
        data_folder: Thư mục chứa dữ liệu đầu vào
        output_folder: Thư mục xuất kết quả
        label: Nhãn cho báo cáo (ví dụ: "TRƯỚC GIẢM TRỪ", "SAU GIẢM TRỪ")
    """
    title = f"TÍNH ĐIỂM KPI NVKT - BSC Q4/2025 VNPT Hà Nội"
    if label:
        title += f" ({label})"
    
    print("="*60)
    print(title)
    print("="*60)
    print(f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    df_result = tinh_diem_kpi_nvkt(data_folder, output_folder)
    
    print()
    print("="*60)
    print("THỐNG KÊ KẾT QUẢ")
    print("="*60)
    
    # Thống kê điểm
    for col in ['Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']:
        valid_data = df_result[col].dropna()
        if len(valid_data) > 0:
            print(f"\n{col}:")
            print(f"  - Số NVKT: {len(valid_data)}")
            print(f"  - Điểm TB: {valid_data.mean():.2f}")
            print(f"  - Điểm min: {valid_data.min():.2f}")
            print(f"  - Điểm max: {valid_data.max():.2f}")
            print(f"  - Số đạt 5 điểm: {(valid_data >= 5).sum()}")
            print(f"  - Số dưới 3 điểm: {(valid_data < 3).sum()}")
    
    return df_result


def tinh_diem_kpi_nvkt_sau_giam_tru(exclusion_folder, original_data_folder, output_folder=None):
    """
    Tính điểm KPI cho tất cả NVKT dựa trên dữ liệu SAU GIẢM TRỪ
    
    Args:
        exclusion_folder: Thư mục chứa các file kết quả sau giảm trừ
        original_data_folder: Thư mục chứa file dữ liệu gốc (cho C1.4 không có giảm trừ)
        output_folder: Thư mục xuất kết quả
    
    Returns:
        DataFrame chứa điểm KPI của từng NVKT
    """
    print(f"Đang đọc dữ liệu SAU GIẢM TRỪ từ: {exclusion_folder}")
    
    # 1. Đọc dữ liệu từ các file sau giảm trừ
    df_c11_tp1 = doc_C11_TP1_sau_giam_tru(exclusion_folder)
    print(f"  - C1.1 TP1 (sau GT): {len(df_c11_tp1)} NVKT")
    
    df_c11_tp2 = doc_C11_TP2_sau_giam_tru(exclusion_folder)
    print(f"  - C1.1 TP2 (sau GT): {len(df_c11_tp2)} NVKT")
    
    df_c12_tp1 = doc_C12_TP1_sau_giam_tru(exclusion_folder)
    print(f"  - C1.2 TP1 (sau GT): {len(df_c12_tp1)} NVKT")
    
    df_c12_tp2 = doc_C12_TP2_sau_giam_tru(exclusion_folder)
    print(f"  - C1.2 TP2 (sau GT): {len(df_c12_tp2)} NVKT")
    
    # C1.4 không có giảm trừ, dùng dữ liệu gốc
    df_c14 = doc_C14(original_data_folder)
    print(f"  - C1.4 (gốc): {len(df_c14)} NVKT")
    
    # 2. Merge tất cả dữ liệu theo nvkt VÀ don_vi
    merge_keys = ['don_vi', 'nvkt']
    
    df_all = df_c11_tp1.copy()
    
    df_all = df_all.merge(df_c11_tp2, on=merge_keys, how='outer')
    df_all = df_all.merge(df_c12_tp1, on=merge_keys, how='outer')
    df_all = df_all.merge(df_c12_tp2, on=merge_keys, how='outer')
    df_all = df_all.merge(df_c14, on=merge_keys, how='outer')
    
    print(f"\nTổng số NVKT sau merge: {len(df_all)}")
    
    # 3. Tính điểm từng thành phần
    print("Đang tính điểm các thành phần...")
    
    df_all['diem_c11_tp1'] = df_all['c11_tp1_ty_le'].apply(tinh_diem_C11_TP1)
    df_all['diem_c11_tp2'] = df_all['c11_tp2_ty_le'].apply(tinh_diem_C11_TP2)
    df_all['diem_c12_tp1'] = df_all['c12_tp1_ty_le'].apply(tinh_diem_C12_TP1)
    df_all['diem_c12_tp2'] = df_all['c12_tp2_ty_le'].apply(tinh_diem_C12_TP2)
    df_all['diem_c14'] = df_all['c14_ty_le'].apply(tinh_diem_C14)
    
    # 4. Tính điểm tổng hợp
    df_all['Diem_C1.1'] = df_all['diem_c11_tp1'] * 0.30 + df_all['diem_c11_tp2'] * 0.70
    df_all['Diem_C1.2'] = df_all['diem_c12_tp1'] * 0.50 + df_all['diem_c12_tp2'] * 0.50
    df_all['Diem_C1.4'] = df_all['diem_c14']
    
    # 5. Làm tròn điểm
    diem_cols = ['diem_c11_tp1', 'diem_c11_tp2', 'diem_c12_tp1', 'diem_c12_tp2', 'diem_c14',
                 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']
    for col in diem_cols:
        df_all[col] = df_all[col].round(2)
    
    # Làm tròn tỷ lệ về %
    ty_le_cols = ['c11_tp1_ty_le', 'c11_tp2_ty_le', 'c12_tp1_ty_le', 'c12_tp2_ty_le', 'c14_ty_le']
    for col in ty_le_cols:
        if col in df_all.columns:
            df_all[col] = (df_all[col] * 100).round(2)
    
    # 6. Sắp xếp các cột
    col_order = [
        'don_vi', 'nvkt',
        'c11_tp1_tong_phieu', 'c11_tp1_phieu_dat', 'c11_tp1_ty_le', 'diem_c11_tp1',
        'c11_tp2_tong_phieu', 'c11_tp2_phieu_dat', 'c11_tp2_ty_le', 'diem_c11_tp2',
        'Diem_C1.1',
        'c12_tp1_phieu_hll', 'c12_tp1_phieu_bh', 'c12_tp1_ty_le', 'diem_c12_tp1',
        'c12_tp2_phieu_bh', 'c12_tp2_tong_tb', 'c12_tp2_ty_le', 'diem_c12_tp2',
        'Diem_C1.2',
        'c14_phieu_ks', 'c14_phieu_khl', 'c14_ty_le', 'diem_c14', 'Diem_C1.4'
    ]
    
    existing_cols = [col for col in col_order if col in df_all.columns]
    df_all = df_all[existing_cols]
    df_all = df_all.sort_values(['don_vi', 'nvkt']).reset_index(drop=True)
    
    # 7. Xuất file
    if output_folder:
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        full_file = output_folder / f"KPI_NVKT_SauGiamTru_ChiTiet_{timestamp}.xlsx"
        df_all.to_excel(full_file, index=False)
        print(f"\nĐã xuất file chi tiết: {full_file}")
        
        summary_cols = ['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']
        df_summary = df_all[summary_cols].copy()
        summary_file = output_folder / f"KPI_NVKT_SauGiamTru_TomTat_{timestamp}.xlsx"
        df_summary.to_excel(summary_file, index=False)
        print(f"Đã xuất file tóm tắt: {summary_file}")
    
    return df_all


def tao_bao_cao_kpi_sau_giam_tru(exclusion_folder, original_data_folder, output_folder):
    """
    Wrapper function để tạo báo cáo KPI SAU GIẢM TRỪ
    """
    print("="*60)
    print("TÍNH ĐIỂM KPI NVKT - BSC Q4/2025 VNPT Hà Nội (SAU GIẢM TRỪ)")
    print("="*60)
    print(f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    df_result = tinh_diem_kpi_nvkt_sau_giam_tru(exclusion_folder, original_data_folder, output_folder)
    
    print()
    print("="*60)
    print("THỐNG KÊ KẾT QUẢ (SAU GIẢM TRỪ)")
    print("="*60)
    
    for col in ['Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']:
        valid_data = df_result[col].dropna()
        if len(valid_data) > 0:
            print(f"\n{col}:")
            print(f"  - Số NVKT: {len(valid_data)}")
            print(f"  - Điểm TB: {valid_data.mean():.2f}")
            print(f"  - Điểm min: {valid_data.min():.2f}")
            print(f"  - Điểm max: {valid_data.max():.2f}")
            print(f"  - Số đạt 5 điểm: {(valid_data >= 5).sum()}")
            print(f"  - Số dưới 3 điểm: {(valid_data < 3).sum()}")
    
    return df_result


def tao_bao_cao_so_sanh_kpi(data_folder, exclusion_folder, output_folder):
    """
    Tạo báo cáo so sánh KPI trước/sau giảm trừ
    """
    print("="*60)
    print("SO SÁNH KPI TRƯỚC/SAU GIẢM TRỪ")
    print("="*60)
    
    # Tính KPI trước giảm trừ
    df_truoc = tinh_diem_kpi_nvkt(data_folder, None)
    df_truoc = df_truoc[['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']].copy()
    df_truoc.columns = ['don_vi', 'nvkt', 'C1.1_Truoc', 'C1.2_Truoc', 'C1.4_Truoc']
    
    # Tính KPI sau giảm trừ
    df_sau = tinh_diem_kpi_nvkt_sau_giam_tru(exclusion_folder, data_folder, None)
    df_sau = df_sau[['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']].copy()
    df_sau.columns = ['don_vi', 'nvkt', 'C1.1_Sau', 'C1.2_Sau', 'C1.4_Sau']
    
    # Merge và tính chênh lệch
    df_compare = df_truoc.merge(df_sau, on=['don_vi', 'nvkt'], how='outer')
    
    df_compare['C1.1_CL'] = df_compare['C1.1_Sau'] - df_compare['C1.1_Truoc']
    df_compare['C1.2_CL'] = df_compare['C1.2_Sau'] - df_compare['C1.2_Truoc']
    df_compare['C1.4_CL'] = df_compare['C1.4_Sau'] - df_compare['C1.4_Truoc']
    
    # Sắp xếp cột
    df_compare = df_compare[[
        'don_vi', 'nvkt',
        'C1.1_Truoc', 'C1.1_Sau', 'C1.1_CL',
        'C1.2_Truoc', 'C1.2_Sau', 'C1.2_CL',
        'C1.4_Truoc', 'C1.4_Sau', 'C1.4_CL'
    ]]
    
    df_compare = df_compare.sort_values(['don_vi', 'nvkt']).reset_index(drop=True)
    
    # Xuất file
    if output_folder:
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        compare_file = output_folder / f"KPI_NVKT_SoSanh_{timestamp}.xlsx"
        df_compare.to_excel(compare_file, index=False)
        print(f"\nĐã xuất file so sánh: {compare_file}")
    
    # Thống kê
    print("\nTHỐNG KÊ CHÊNH LỆCH:")
    for col in ['C1.1_CL', 'C1.2_CL', 'C1.4_CL']:
        valid_data = df_compare[col].dropna()
        if len(valid_data) > 0:
            chi_tieu = col.replace('_CL', '')
            print(f"\n{chi_tieu}:")
            print(f"  - Chênh lệch TB: {valid_data.mean():.2f}")
            print(f"  - Số tăng điểm: {(valid_data > 0).sum()}")
            print(f"  - Số giảm điểm: {(valid_data < 0).sum()}")
            print(f"  - Số không đổi: {(valid_data == 0).sum()}")
    
    return df_compare


# ============================================================================
# MAIN - Chạy trực tiếp từ command line
# ============================================================================

if __name__ == "__main__":
    # Đường dẫn mặc định
    DATA_FOLDER = "downloads/baocao_hanoi"
    EXCLUSION_FOLDER = "downloads/kq_sau_giam_tru"
    OUTPUT_FOLDER = "downloads/KPI"
    
    # 1. Tính KPI trước giảm trừ
    print("\n" + "="*60)
    print("PHẦN 1: KPI TRƯỚC GIẢM TRỪ")
    print("="*60)
    df_truoc = tao_bao_cao_kpi(DATA_FOLDER, OUTPUT_FOLDER, "TRƯỚC GIẢM TRỪ")
    
    # 2. Tính KPI sau giảm trừ
    print("\n\n" + "="*60)
    print("PHẦN 2: KPI SAU GIẢM TRỪ")
    print("="*60)
    df_sau = tao_bao_cao_kpi_sau_giam_tru(EXCLUSION_FOLDER, DATA_FOLDER, OUTPUT_FOLDER)
    
    # 3. Tạo báo cáo so sánh
    print("\n\n" + "="*60)
    print("PHẦN 3: SO SÁNH TRƯỚC/SAU GIẢM TRỪ")
    print("="*60)
    df_compare = tao_bao_cao_so_sanh_kpi(DATA_FOLDER, EXCLUSION_FOLDER, OUTPUT_FOLDER)
    
    print("\n" + "="*60)
    print("HOÀN THÀNH!")
    print("="*60)

