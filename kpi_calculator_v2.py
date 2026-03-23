"""
Module tính điểm KPI cho NVKT theo BSC Q4/2025 VNPT Hà Nội (v2)
Các chỉ tiêu: C1.1, C1.2, C1.4, C1.5

Refactored: Import scoring từ kpi_scoring.py (single source of truth)
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

from kpi_scoring import (
    tinh_diem_C11_TP1, tinh_diem_C11_TP2,
    tinh_diem_C12_TP1, tinh_diem_C12_TP2,
    tinh_diem_C14, tinh_diem_C15,
    chuan_hoa_ty_le_df, chuan_hoa_ten,
)


# ============================================================================
# CÁC HÀM ĐỌC DỮ LIỆU TỪ FILE EXCEL
# ============================================================================

def doc_C11_TP1(data_folder):
    """
    Đọc dữ liệu C1.1 Thành phần 1: Tỷ lệ sửa chữa chủ động
    File: SM2-C11.xlsx, Sheet: CT_C1.1_TP1
    """
    file_path = Path(data_folder) / "SM2-C11.xlsx"
    df = pd.read_excel(file_path, sheet_name="CT_C1.1_TP1")

    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu', 'Phiếu đạt', 'Tỉ lệ đạt (%)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp1_tong_phieu', 'c11_tp1_phieu_dat', 'c11_tp1_ty_le']

    df = chuan_hoa_ten(df, 'nvkt')

    df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
        'c11_tp1_tong_phieu': 'sum',
        'c11_tp1_phieu_dat': 'sum',
        'c11_tp1_ty_le': 'first'
    })
    df['c11_tp1_ty_le'] = df['c11_tp1_phieu_dat'] / df['c11_tp1_tong_phieu']

    return df


def doc_C11_TP2(data_folder):
    """
    Đọc dữ liệu C1.1 Thành phần 2: Tỷ lệ sửa chữa báo hỏng đúng quy định (không tính hẹn)
    File: SM4-C11.xlsx, Sheet: chi_tiet
    """
    file_path = Path(data_folder) / "SM4-C11.xlsx"
    df = pd.read_excel(file_path, sheet_name="chi_tiet")

    col_ty_le = 'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn'

    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu', 'Số phiếu đạt', col_ty_le]].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp2_tong_phieu', 'c11_tp2_phieu_dat', 'c11_tp2_ty_le']

    df = chuan_hoa_ten(df, 'nvkt')
    df = chuan_hoa_ty_le_df(df, 'c11_tp2_ty_le')

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

    df = chuan_hoa_ten(df, 'nvkt')
    df = chuan_hoa_ty_le_df(df, 'c12_tp1_ty_le')

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

    df = chuan_hoa_ten(df, 'nvkt')

    # Dữ liệu gốc có tỷ lệ dạng 0.37 = 0.37%, cần chia 100 để về dạng thập phân
    df['c12_tp2_ty_le'] = df['c12_tp2_ty_le'] / 100

    return df


def doc_C14(data_folder):
    """
    Đọc dữ liệu C1.4: Độ hài lòng khách hàng
    File: c1.4_chitiet_report.xlsx, Sheet: TH_HL_NVKT
    """
    file_path = Path(data_folder) / "c1.4_chitiet_report.xlsx"
    df = pd.read_excel(file_path, sheet_name="TH_HL_NVKT")

    df = df[['DOIVT', 'NVKT', 'Tổng phiếu KS thành công', 'Tổng phiếu KHL', 'Tỉ lệ HL NVKT (%)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c14_phieu_ks', 'c14_phieu_khl', 'c14_ty_le']

    df = chuan_hoa_ten(df, 'nvkt')
    df = chuan_hoa_ty_le_df(df, 'c14_ty_le')

    # Mẫu số = 0 -> tỷ lệ = 100%
    mask_no_survey = (df['c14_phieu_ks'] == 0) | (df['c14_phieu_ks'].isna())
    df.loc[mask_no_survey, 'c14_ty_le'] = 1.0

    return df


def doc_C15(data_folder):
    """
    Đọc dữ liệu C1.5: Tỉ lệ thiết lập dịch vụ đạt thời gian quy định
    File: c1.5_chitiet_report.xlsx, Sheet: KQ_C15_chitiet
    """
    file_path = Path(data_folder) / "c1.5_chitiet_report.xlsx"
    df = pd.read_excel(file_path, sheet_name="KQ_C15_chitiet")

    df = df[['DOIVT', 'NVKT', 'Phiếu đạt', 'Phiếu không đạt', 'Tổng Hoàn công', 'Tỉ lệ đạt (%)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c15_phieu_dat', 'c15_phieu_khong_dat', 'c15_tong_phieu', 'c15_ty_le']

    df = chuan_hoa_ten(df, 'nvkt')
    df = chuan_hoa_ty_le_df(df, 'c15_ty_le')

    return df


# ============================================================================
# CÁC HÀM ĐỌC DỮ LIỆU SAU GIẢM TRỪ
# ============================================================================

def doc_C11_TP1_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.1 TP1 SAU GIẢM TRỪ
    File: So_sanh_C11_SM2.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C11_SM2.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")

    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu (Sau GT)', 'Số phiếu đạt (Sau GT)', 'Tỷ lệ % (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp1_tong_phieu', 'c11_tp1_phieu_dat', 'c11_tp1_ty_le']

    df = chuan_hoa_ty_le_df(df, 'c11_tp1_ty_le')
    df = chuan_hoa_ten(df, 'nvkt')

    df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
        'c11_tp1_tong_phieu': 'sum',
        'c11_tp1_phieu_dat': 'sum',
        'c11_tp1_ty_le': 'first'
    })
    df['c11_tp1_ty_le'] = df['c11_tp1_phieu_dat'] / df['c11_tp1_tong_phieu'].replace(0, 1)

    return df


def doc_C11_TP2_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.1 TP2 SAU GIẢM TRỪ
    File: So_sanh_C11_SM4.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C11_SM4.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")

    df = df[['TEN_DOI', 'NVKT', 'Tổng phiếu (Sau GT)', 'Số phiếu đạt (Sau GT)', 'Tỷ lệ % (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c11_tp2_tong_phieu', 'c11_tp2_phieu_dat', 'c11_tp2_ty_le']

    df = chuan_hoa_ty_le_df(df, 'c11_tp2_ty_le')
    df = chuan_hoa_ten(df, 'nvkt')

    df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
        'c11_tp2_tong_phieu': 'sum',
        'c11_tp2_phieu_dat': 'sum',
        'c11_tp2_ty_le': 'first'
    })
    df['c11_tp2_ty_le'] = df['c11_tp2_phieu_dat'] / df['c11_tp2_tong_phieu'].replace(0, 1)

    return df


def doc_C12_TP1_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.2 TP1 SAU GIẢM TRỪ
    File: So_sanh_C12_SM1.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C12_SM1.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")

    df = df[['TEN_DOI', 'NVKT', 'Số phiếu HLL (Sau GT)', 'Số phiếu báo hỏng (Sau GT)', 'Tỷ lệ HLL % (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c12_tp1_phieu_hll', 'c12_tp1_phieu_bh', 'c12_tp1_ty_le']

    df = chuan_hoa_ty_le_df(df, 'c12_tp1_ty_le')
    df = chuan_hoa_ten(df, 'nvkt')

    df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
        'c12_tp1_phieu_hll': 'sum',
        'c12_tp1_phieu_bh': 'sum',
        'c12_tp1_ty_le': 'first'
    })
    df['c12_tp1_ty_le'] = df['c12_tp1_phieu_hll'] / df['c12_tp1_phieu_bh'].replace(0, 1)

    return df


def doc_C12_TP2_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.2 TP2 SAU GIẢM TRỪ
    File: SM4-C12-ti-le-su-co-dv-brcd.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "SM4-C12-ti-le-su-co-dv-brcd.xlsx"
    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")

    df = df[['TEN_DOI', 'NVKT', 'Tổng TB (Thô)', 'Số phiếu báo hỏng (Sau GT)', 'Tỷ lệ báo hỏng (%) (Sau GT)']].copy()
    df.columns = ['don_vi', 'nvkt', 'c12_tp2_tong_tb', 'c12_tp2_phieu_bh', 'c12_tp2_ty_le']

    df = chuan_hoa_ty_le_df(df, 'c12_tp2_ty_le')
    df = chuan_hoa_ten(df, 'nvkt')

    df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
        'c12_tp2_tong_tb': 'sum',
        'c12_tp2_phieu_bh': 'sum',
        'c12_tp2_ty_le': 'first'
    })
    df['c12_tp2_ty_le'] = df['c12_tp2_phieu_bh'] / df['c12_tp2_tong_tb'].replace(0, 1)

    return df


def doc_C14_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.4 SAU GIẢM TRỪ
    File: So_sanh_C14.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C14.xlsx"

    if not file_path.exists():
        print(f"  [WARN] Không tìm thấy file C1.4 sau giảm trừ: {file_path}")
        return None

    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")

    cols_needed = ['TEN_DOI', 'NVKT', 'Tổng phiếu KS (Sau GT)', 'Số phiếu KHL (Sau GT)', 'Tỷ lệ HL (%) (Sau GT)']
    cols_available = [c for c in cols_needed if c in df.columns]

    if 'TEN_DOI' in df.columns:
        df = df[cols_available].copy()
        df.columns = ['don_vi', 'nvkt', 'c14_phieu_ks', 'c14_phieu_khl', 'c14_ty_le']
    else:
        cols_needed = ['NVKT', 'Tổng phiếu KS (Sau GT)', 'Số phiếu KHL (Sau GT)', 'Tỷ lệ HL (%) (Sau GT)']
        df = df[cols_needed].copy()
        df.columns = ['nvkt', 'c14_phieu_ks', 'c14_phieu_khl', 'c14_ty_le']
        df['don_vi'] = None

    df = chuan_hoa_ten(df, 'nvkt')
    df = chuan_hoa_ty_le_df(df, 'c14_ty_le')

    mask_no_survey = (df['c14_phieu_ks'] == 0) | (df['c14_phieu_ks'].isna())
    df.loc[mask_no_survey, 'c14_ty_le'] = 1.0

    if 'don_vi' in df.columns and df['don_vi'].notna().any():
        df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
            'c14_phieu_ks': 'sum',
            'c14_phieu_khl': 'sum',
            'c14_ty_le': 'first'
        })
        mask = df['c14_phieu_ks'] > 0
        df.loc[mask, 'c14_ty_le'] = (df.loc[mask, 'c14_phieu_ks'] - df.loc[mask, 'c14_phieu_khl']) / df.loc[mask, 'c14_phieu_ks']
        df.loc[~mask, 'c14_ty_le'] = 1.0

    return df


def doc_C15_sau_giam_tru(exclusion_folder):
    """
    Đọc dữ liệu C1.5 SAU GIẢM TRỪ
    File: So_sanh_C15.xlsx, Sheet: So_sanh_chi_tiet
    """
    file_path = Path(exclusion_folder) / "So_sanh_C15.xlsx"

    if not file_path.exists():
        print(f"  [WARN] Không tìm thấy file C1.5 sau giảm trừ: {file_path}")
        return None

    df = pd.read_excel(file_path, sheet_name="So_sanh_chi_tiet")

    cols_needed = ['TEN_DOI', 'NVKT', 'Tổng Hoàn công (Sau GT)', 'Phiếu đạt (Sau GT)', 'Tỷ lệ đạt (%) (Sau GT)']
    cols_available = [c for c in cols_needed if c in df.columns]

    if 'TEN_DOI' in df.columns:
        df = df[cols_available].copy()
        df.columns = ['don_vi', 'nvkt', 'c15_tong_phieu', 'c15_phieu_dat', 'c15_ty_le']
    else:
        cols_needed = ['NVKT', 'Tổng Hoàn công (Sau GT)', 'Phiếu đạt (Sau GT)', 'Tỷ lệ đạt (%) (Sau GT)']
        df = df[[c for c in cols_needed if c in df.columns]].copy()
        df.columns = ['nvkt', 'c15_tong_phieu', 'c15_phieu_dat', 'c15_ty_le']
        df['don_vi'] = None

    df['c15_phieu_khong_dat'] = df['c15_tong_phieu'] - df['c15_phieu_dat']
    df = chuan_hoa_ten(df, 'nvkt')
    df = chuan_hoa_ty_le_df(df, 'c15_ty_le')

    if 'don_vi' in df.columns and df['don_vi'].notna().any():
        df = df.groupby(['don_vi', 'nvkt'], as_index=False).agg({
            'c15_tong_phieu': 'sum',
            'c15_phieu_dat': 'sum',
            'c15_phieu_khong_dat': 'sum',
            'c15_ty_le': 'first'
        })
        mask = df['c15_tong_phieu'] > 0
        df.loc[mask, 'c15_ty_le'] = df.loc[mask, 'c15_phieu_dat'] / df.loc[mask, 'c15_tong_phieu']
        df.loc[~mask, 'c15_ty_le'] = 1.0

    return df


# ============================================================================
# HÀM TÍNH ĐIỂM KPI TỔNG HỢP
# ============================================================================

def _merge_all_data(df_c11_tp1, df_c11_tp2, df_c12_tp1, df_c12_tp2, df_c14, df_c15):
    """Merge tất cả DataFrame theo don_vi + nvkt"""
    merge_keys = ['don_vi', 'nvkt']
    df_all = df_c11_tp1.copy()
    df_all = df_all.merge(df_c11_tp2, on=merge_keys, how='outer')
    df_all = df_all.merge(df_c12_tp1, on=merge_keys, how='outer')
    df_all = df_all.merge(df_c12_tp2, on=merge_keys, how='outer')
    df_all = df_all.merge(df_c14, on=merge_keys, how='outer')
    df_all = df_all.merge(df_c15, on=merge_keys, how='outer')
    return df_all


def _tinh_diem_va_lam_tron(df_all):
    """Tính điểm từng thành phần và tổng hợp, làm tròn"""
    # Tính điểm từng thành phần
    df_all['diem_c11_tp1'] = df_all['c11_tp1_ty_le'].apply(tinh_diem_C11_TP1)
    df_all['diem_c11_tp2'] = df_all['c11_tp2_ty_le'].apply(tinh_diem_C11_TP2)
    df_all['diem_c12_tp1'] = df_all['c12_tp1_ty_le'].apply(tinh_diem_C12_TP1)
    df_all['diem_c12_tp2'] = df_all['c12_tp2_ty_le'].apply(tinh_diem_C12_TP2)
    df_all['diem_c14'] = df_all['c14_ty_le'].apply(tinh_diem_C14)
    df_all['diem_c15'] = df_all['c15_ty_le'].apply(tinh_diem_C15)

    # Tính điểm tổng hợp
    df_all['Diem_C1.1'] = df_all['diem_c11_tp1'] * 0.30 + df_all['diem_c11_tp2'] * 0.70
    df_all['Diem_C1.2'] = df_all['diem_c12_tp1'] * 0.50 + df_all['diem_c12_tp2'] * 0.50
    df_all['Diem_C1.4'] = df_all['diem_c14']
    df_all['Diem_C1.5'] = df_all['diem_c15']

    # Làm tròn điểm
    diem_cols = ['diem_c11_tp1', 'diem_c11_tp2', 'diem_c12_tp1', 'diem_c12_tp2', 'diem_c14', 'diem_c15',
                 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4', 'Diem_C1.5']
    for col in diem_cols:
        if col in df_all.columns:
            df_all[col] = df_all[col].round(2)

    # Làm tròn tỷ lệ về %
    ty_le_cols = ['c11_tp1_ty_le', 'c11_tp2_ty_le', 'c12_tp1_ty_le', 'c12_tp2_ty_le', 'c14_ty_le', 'c15_ty_le']
    for col in ty_le_cols:
        if col in df_all.columns:
            df_all[col] = (df_all[col] * 100).round(2)

    return df_all


def _sap_xep_cot(df_all):
    """Sắp xếp cột theo thứ tự logic"""
    col_order = [
        'don_vi', 'nvkt',
        'c11_tp1_tong_phieu', 'c11_tp1_phieu_dat', 'c11_tp1_ty_le', 'diem_c11_tp1',
        'c11_tp2_tong_phieu', 'c11_tp2_phieu_dat', 'c11_tp2_ty_le', 'diem_c11_tp2',
        'Diem_C1.1',
        'c12_tp1_phieu_hll', 'c12_tp1_phieu_bh', 'c12_tp1_ty_le', 'diem_c12_tp1',
        'c12_tp2_phieu_bh', 'c12_tp2_tong_tb', 'c12_tp2_ty_le', 'diem_c12_tp2',
        'Diem_C1.2',
        'c14_phieu_ks', 'c14_phieu_khl', 'c14_ty_le', 'diem_c14', 'Diem_C1.4',
        'c15_phieu_dat', 'c15_phieu_khong_dat', 'c15_tong_phieu', 'c15_ty_le', 'diem_c15', 'Diem_C1.5'
    ]
    existing_cols = [col for col in col_order if col in df_all.columns]
    df_all = df_all[existing_cols]
    df_all = df_all.sort_values(['don_vi', 'nvkt']).reset_index(drop=True)
    return df_all


def tinh_diem_kpi_nvkt(data_folder, output_folder=None):
    """
    Tính điểm KPI cho tất cả NVKT dựa trên dữ liệu từ các file Excel

    Args:
        data_folder: Thư mục chứa các file dữ liệu Excel
        output_folder: Thư mục xuất kết quả (None = không xuất file)

    Returns:
        DataFrame chứa điểm KPI của từng NVKT
    """
    print(f"Đang đọc dữ liệu từ: {data_folder}")

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

    df_c15 = doc_C15(data_folder)
    print(f"  - C1.5: {len(df_c15)} NVKT")

    # Merge, tính điểm, sắp xếp
    df_all = _merge_all_data(df_c11_tp1, df_c11_tp2, df_c12_tp1, df_c12_tp2, df_c14, df_c15)
    print(f"\nTổng số NVKT sau merge: {len(df_all)}")

    print("Đang tính điểm các thành phần...")
    df_all = _tinh_diem_va_lam_tron(df_all)
    df_all = _sap_xep_cot(df_all)

    # Xuất file
    if output_folder:
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        # File chi tiết
        full_file = output_folder / "KPI_NVKT_ChiTiet.xlsx"
        c15_cols = ['don_vi', 'nvkt', 'c15_phieu_dat', 'c15_phieu_khong_dat', 'c15_tong_phieu', 'c15_ty_le', 'diem_c15', 'Diem_C1.5']
        df_c15_detail = df_all[[col for col in c15_cols if col in df_all.columns]].copy()
        df_c15_detail = df_c15_detail.dropna(subset=['c15_tong_phieu'])

        with pd.ExcelWriter(full_file, engine='openpyxl') as writer:
            df_all.to_excel(writer, sheet_name='KPI_ChiTiet', index=False)
            df_c15_detail.to_excel(writer, sheet_name='C1.5_ChiTiet', index=False)
        print(f"\nĐã xuất file chi tiết: {full_file}")

        # File tóm tắt
        summary_cols = ['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4', 'Diem_C1.5']
        df_summary = df_all[[col for col in summary_cols if col in df_all.columns]].copy()

        c15_summary_cols = ['don_vi', 'nvkt', 'c15_ty_le', 'Diem_C1.5']
        df_c15_summary = df_all[[col for col in c15_summary_cols if col in df_all.columns]].copy()
        df_c15_summary = df_c15_summary.dropna(subset=['Diem_C1.5'])

        summary_file = output_folder / "KPI_NVKT_TomTat.xlsx"
        with pd.ExcelWriter(summary_file, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='KPI_TomTat', index=False)
            df_c15_summary.to_excel(writer, sheet_name='C1.5_TomTat', index=False)
        print(f"Đã xuất file tóm tắt: {summary_file}")

    return df_all


def tao_bao_cao_kpi(data_folder, output_folder, label=""):
    """Wrapper: tạo báo cáo KPI hoàn chỉnh"""
    title = f"TÍNH ĐIỂM KPI NVKT - BSC Q4/2025 VNPT Hà Nội"
    if label:
        title += f" ({label})"

    print("=" * 60)
    print(title)
    print("=" * 60)
    print(f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()

    df_result = tinh_diem_kpi_nvkt(data_folder, output_folder)

    print()
    print("=" * 60)
    print("THỐNG KÊ KẾT QUẢ")
    print("=" * 60)

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
        original_data_folder: Thư mục chứa file dữ liệu gốc (cho C1.4, C1.5 không có giảm trừ)
        output_folder: Thư mục xuất kết quả
    """
    print(f"Đang đọc dữ liệu SAU GIẢM TRỪ từ: {exclusion_folder}")

    df_c11_tp1 = doc_C11_TP1_sau_giam_tru(exclusion_folder)
    print(f"  - C1.1 TP1 (sau GT): {len(df_c11_tp1)} NVKT")

    df_c11_tp2 = doc_C11_TP2_sau_giam_tru(exclusion_folder)
    print(f"  - C1.1 TP2 (sau GT): {len(df_c11_tp2)} NVKT")

    df_c12_tp1 = doc_C12_TP1_sau_giam_tru(exclusion_folder)
    print(f"  - C1.2 TP1 (sau GT): {len(df_c12_tp1)} NVKT")

    df_c12_tp2 = doc_C12_TP2_sau_giam_tru(exclusion_folder)
    print(f"  - C1.2 TP2 (sau GT): {len(df_c12_tp2)} NVKT")

    # C1.4 - fallback về dữ liệu gốc nếu không có sau giảm trừ
    df_c14 = doc_C14_sau_giam_tru(exclusion_folder)
    if df_c14 is not None and len(df_c14) > 0:
        print(f"  - C1.4 (sau GT): {len(df_c14)} NVKT")
    else:
        df_c14 = doc_C14(original_data_folder)
        print(f"  - C1.4 (gốc): {len(df_c14)} NVKT")

    # C1.5 - fallback về dữ liệu gốc nếu không có sau giảm trừ
    df_c15 = doc_C15_sau_giam_tru(exclusion_folder)
    if df_c15 is not None and len(df_c15) > 0:
        print(f"  - C1.5 (sau GT): {len(df_c15)} NVKT")
    else:
        df_c15 = doc_C15(original_data_folder)
        print(f"  - C1.5 (gốc): {len(df_c15)} NVKT")

    # Merge, tính điểm, sắp xếp
    df_all = _merge_all_data(df_c11_tp1, df_c11_tp2, df_c12_tp1, df_c12_tp2, df_c14, df_c15)
    print(f"\nTổng số NVKT sau merge: {len(df_all)}")

    print("Đang tính điểm các thành phần...")
    df_all = _tinh_diem_va_lam_tron(df_all)
    df_all = _sap_xep_cot(df_all)

    # Xuất file
    if output_folder:
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        full_file = output_folder / "KPI_NVKT_SauGiamTru_ChiTiet.xlsx"
        df_all.to_excel(full_file, index=False)
        print(f"\nĐã xuất file chi tiết: {full_file}")

        summary_cols = ['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4', 'Diem_C1.5']
        df_summary = df_all[[col for col in summary_cols if col in df_all.columns]].copy()
        summary_file = output_folder / "KPI_NVKT_SauGiamTru_TomTat.xlsx"
        df_summary.to_excel(summary_file, index=False)
        print(f"Đã xuất file tóm tắt: {summary_file}")

    return df_all


def tao_bao_cao_kpi_sau_giam_tru(exclusion_folder, original_data_folder, output_folder):
    """Wrapper: tạo báo cáo KPI SAU GIẢM TRỪ"""
    print("=" * 60)
    print("TÍNH ĐIỂM KPI NVKT - BSC Q4/2025 VNPT Hà Nội (SAU GIẢM TRỪ)")
    print("=" * 60)
    print(f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()

    df_result = tinh_diem_kpi_nvkt_sau_giam_tru(exclusion_folder, original_data_folder, output_folder)

    print()
    print("=" * 60)
    print("THỐNG KÊ KẾT QUẢ (SAU GIẢM TRỪ)")
    print("=" * 60)

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
    """Tạo báo cáo so sánh KPI trước/sau giảm trừ"""
    print("=" * 60)
    print("SO SÁNH KPI TRƯỚC/SAU GIẢM TRỪ")
    print("=" * 60)

    df_truoc = tinh_diem_kpi_nvkt(data_folder, None)
    df_truoc = df_truoc[['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']].copy()
    df_truoc.columns = ['don_vi', 'nvkt', 'C1.1_Truoc', 'C1.2_Truoc', 'C1.4_Truoc']

    df_sau = tinh_diem_kpi_nvkt_sau_giam_tru(exclusion_folder, data_folder, None)
    df_sau = df_sau[['don_vi', 'nvkt', 'Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4']].copy()
    df_sau.columns = ['don_vi', 'nvkt', 'C1.1_Sau', 'C1.2_Sau', 'C1.4_Sau']

    df_compare = df_truoc.merge(df_sau, on=['don_vi', 'nvkt'], how='outer')

    df_compare['C1.1_CL'] = df_compare['C1.1_Sau'] - df_compare['C1.1_Truoc']
    df_compare['C1.2_CL'] = df_compare['C1.2_Sau'] - df_compare['C1.2_Truoc']
    df_compare['C1.4_CL'] = df_compare['C1.4_Sau'] - df_compare['C1.4_Truoc']

    df_compare = df_compare[[
        'don_vi', 'nvkt',
        'C1.1_Truoc', 'C1.1_Sau', 'C1.1_CL',
        'C1.2_Truoc', 'C1.2_Sau', 'C1.2_CL',
        'C1.4_Truoc', 'C1.4_Sau', 'C1.4_CL'
    ]]
    df_compare = df_compare.sort_values(['don_vi', 'nvkt']).reset_index(drop=True)

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
    DATA_FOLDER = "downloads/baocao_hanoi"
    EXCLUSION_FOLDER = "downloads/kq_sau_giam_tru_hni"
    OUTPUT_FOLDER = "downloads/KPI"

    # 1. Tính KPI trước giảm trừ
    print("\n" + "=" * 60)
    print("PHẦN 1: KPI TRƯỚC GIẢM TRỪ")
    print("=" * 60)
    df_truoc = tao_bao_cao_kpi(DATA_FOLDER, OUTPUT_FOLDER, "TRƯỚC GIẢM TRỪ")

    # 2. Tính KPI sau giảm trừ
    print("\n\n" + "=" * 60)
    print("PHẦN 2: KPI SAU GIẢM TRỪ")
    print("=" * 60)
    df_sau = tao_bao_cao_kpi_sau_giam_tru(EXCLUSION_FOLDER, DATA_FOLDER, OUTPUT_FOLDER)

    # 3. Tạo báo cáo so sánh
    print("\n\n" + "=" * 60)
    print("PHẦN 3: SO SÁNH TRƯỚC/SAU GIẢM TRỪ")
    print("=" * 60)
    df_compare = tao_bao_cao_so_sanh_kpi(DATA_FOLDER, EXCLUSION_FOLDER, OUTPUT_FOLDER)

    print("\n" + "=" * 60)
    print("HOÀN THÀNH!")
    print("=" * 60)
