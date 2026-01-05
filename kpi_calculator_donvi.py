"""
Module tính điểm KPI cho ĐƠN VỊ theo BSC Q4/2025 VNPT Hà Nội
Với khả năng giảm trừ dựa trên danh sách BAOHONG_ID

Các chỉ tiêu: C1.1, C1.2
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import math


# ============================================================================
# CÁC HÀM TÍNH ĐIỂM (copy từ kpi_calculator.py)
# ============================================================================

def tinh_diem_C11_TP1(kq):
    """C1.1 TP1 (30%): Tỷ lệ sửa chữa phiếu chất lượng chủ động"""
    if pd.isna(kq) or kq is None:
        return np.nan
    if kq >= 0.99:
        return 5
    elif kq > 0.96:
        return 1 + 4 * (kq - 0.96) / 0.03
    else:
        return 1


def tinh_diem_C11_TP2(kq):
    """C1.1 TP2 (70%): Tỷ lệ sửa chữa báo hỏng đúng quy định (không tính hẹn)"""
    if pd.isna(kq) or kq is None:
        return np.nan
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
    """C1.2 TP1 (50%): Tỷ lệ thuê bao báo hỏng lặp lại - càng thấp càng tốt"""
    if pd.isna(kq) or kq is None:
        return np.nan
    if kq <= 0.025:
        return 5
    elif kq < 0.04:
        return 5 - 4 * (kq - 0.025) / 0.015
    else:
        return 1


def tinh_diem_C12_TP2(kq):
    """C1.2 TP2 (50%): Tỷ lệ sự cố dịch vụ BRCĐ - càng thấp càng tốt"""
    if pd.isna(kq) or kq is None:
        return np.nan
    if kq <= 0.02:
        return 5
    elif kq < 0.03:
        return 5 - 4 * (kq - 0.02) / 0.01
    else:
        return 1


# ============================================================================
# HÀM ĐỌC DỮ LIỆU
# ============================================================================

def load_exclusion_list(exclusion_file="du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx"):
    """Đọc danh sách BAOHONG_ID cần loại trừ"""
    if not Path(exclusion_file).exists():
        print(f"⚠️ Không tìm thấy file loại trừ: {exclusion_file}")
        return set()
    
    df = pd.read_excel(exclusion_file)
    if 'BAOHONG_ID' not in df.columns:
        print(f"⚠️ Không tìm thấy cột 'BAOHONG_ID'")
        return set()
    
    exclusion_ids = set(df['BAOHONG_ID'].astype(str).tolist())
    print(f"✅ Đã đọc {len(exclusion_ids)} mã BAOHONG_ID cần loại trừ")
    return exclusion_ids


def load_c11_report(data_folder):
    """Đọc báo cáo C1.1 tổng hợp theo đơn vị"""
    file_path = Path(data_folder) / "c1.1 report.xlsx"
    df = pd.read_excel(file_path, sheet_name="TH_C1.1")
    
    # Rename columns for clarity
    # SM3 = số phiếu đạt (tử số)
    # SM4 = tổng phiếu (mẫu số)
    # Tỷ lệ = SM3/SM4 * 100
    df = df.rename(columns={
        'Đơn vị': 'don_vi',
        'SM1': 'c11_sm1',  # Số phiếu sửa chữa chủ động (tử số TP1)
        'SM2': 'c11_sm2',  # Tổng phiếu chủ động (mẫu số TP1)
        'Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu': 'c11_tp1_ty_le',
        'SM3': 'c11_sm3',  # Số phiếu đạt báo hỏng (tử số TP2)
        'SM4': 'c11_sm4',  # Tổng phiếu báo hỏng (mẫu số TP2)
        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': 'c11_tp2_ty_le',
        'Chỉ tiêu BSC': 'c11_bsc_goc'
    })
    
    # Loại bỏ dòng Tổng
    df = df[df['don_vi'] != 'Tổng'].copy()
    
    return df


def load_c12_report(data_folder):
    """Đọc báo cáo C1.2 tổng hợp theo đơn vị"""
    file_path = Path(data_folder) / "c1.2 report.xlsx"
    df = pd.read_excel(file_path, sheet_name="TH_C1.2")
    
    df = df.rename(columns={
        'Đơn vị': 'don_vi',
        'SM1': 'c12_sm1',  # Số phiếu HLL
        'SM2': 'c12_sm2',  # Số phiếu báo hỏng (mẫu số HLL)
        'Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại': 'c12_tp1_ty_le',
        'SM3': 'c12_sm3',  # Số phiếu báo hỏng (tử số tỷ lệ sự cố)
        'SM4': 'c12_sm4',  # Tổng thuê bao (mẫu số tỷ lệ sự cố)
        'Tỷ lệ sự cố dịch vụ BRCĐ': 'c12_tp2_ty_le',
        'Chỉ tiêu BSC': 'c12_bsc_goc'
    })
    
    # Loại bỏ dòng Tổng
    df = df[df['don_vi'] != 'Tổng'].copy()
    
    return df


def get_exclusion_stats_c11(data_folder, exclusion_ids):
    """
    Tính số phiếu loại trừ theo đơn vị cho C1.1 (từ SM4-C11)
    
    SM3 = số phiếu đạt (tử số) → cần trừ số phiếu đạt trong loại trừ
    SM4 = tổng phiếu (mẫu số) → cần trừ tổng phiếu loại trừ
    
    Returns: DataFrame với c11_loai_tru_sm3 (phiếu đạt loại trừ) và c11_loai_tru_sm4 (tổng loại trừ)
    """
    file_path = Path(data_folder) / "SM4-C11.xlsx"
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    
    df['BAOHONG_ID_STR'] = df['BAOHONG_ID'].astype(str)
    excluded = df[df['BAOHONG_ID_STR'].isin(exclusion_ids)]
    
    if len(excluded) == 0:
        return pd.DataFrame(columns=['don_vi', 'c11_loai_tru_sm3', 'c11_loai_tru_sm4'])
    
    # Group by TEN_DOI
    stats = excluded.groupby('TEN_DOI').agg({
        'BAOHONG_ID': 'count',  # Tổng phiếu loại trừ → trừ SM4
        'DAT_TT_KO_HEN': lambda x: (x == 1).sum()  # Phiếu đạt loại trừ → trừ SM3
    }).reset_index()
    
    # SM3 = phiếu đạt, SM4 = tổng
    stats.columns = ['don_vi', 'c11_loai_tru_sm4', 'c11_loai_tru_sm3']
    
    return stats


def get_exclusion_stats_c12_sm1(data_folder, exclusion_ids):
    """
    Tính số phiếu loại trừ theo đơn vị cho C1.2 SM1 (phiếu HLL)
    """
    file_path = Path(data_folder) / "SM1-C12.xlsx"
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    
    df['BAOHONG_ID_STR'] = df['BAOHONG_ID'].astype(str)
    excluded = df[df['BAOHONG_ID_STR'].isin(exclusion_ids)]
    
    if len(excluded) == 0:
        return pd.DataFrame(columns=['don_vi', 'c12_loai_tru_sm1'])
    
    # Số phiếu HLL = số bản ghi / 2, làm tròn lên
    stats = excluded.groupby('TEN_DOI').agg({
        'BAOHONG_ID': lambda x: math.ceil(len(x) / 2)
    }).reset_index()
    
    stats.columns = ['don_vi', 'c12_loai_tru_sm1']
    
    return stats


def get_exclusion_stats_c12_sm2(data_folder, exclusion_ids):
    """
    Tính số phiếu loại trừ theo đơn vị cho C1.2 SM2/SM3 (từ SM2-C12)
    SM2 = số phiếu báo hỏng (mẫu số HLL)
    SM3 = số phiếu báo hỏng (tử số tỷ lệ sự cố)
    """
    file_path = Path(data_folder) / "SM2-C12.xlsx"
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    
    df['BAOHONG_ID_STR'] = df['BAOHONG_ID'].astype(str)
    excluded = df[df['BAOHONG_ID_STR'].isin(exclusion_ids)]
    
    if len(excluded) == 0:
        return pd.DataFrame(columns=['don_vi', 'c12_loai_tru_sm2', 'c12_loai_tru_sm3'])
    
    stats = excluded.groupby('TEN_DOI').agg({
        'BAOHONG_ID': 'count'
    }).reset_index()
    
    stats.columns = ['don_vi', 'c12_loai_tru_sm2']
    stats['c12_loai_tru_sm3'] = stats['c12_loai_tru_sm2']  # SM2 và SM3 đều trừ số phiếu BH
    
    return stats


# ============================================================================
# HÀM TÍNH KPI CHO ĐƠN VỊ
# ============================================================================

def tinh_kpi_donvi_sau_giam_tru(data_folder, exclusion_file, output_folder):
    """
    Tính KPI cho đơn vị với giảm trừ
    
    Args:
        data_folder: Thư mục chứa file báo cáo
        exclusion_file: File danh sách loại trừ
        output_folder: Thư mục xuất kết quả
    """
    print("="*70)
    print("TÍNH KPI ĐƠN VỊ SAU GIẢM TRỪ - BSC Q4/2025")
    print("="*70)
    print(f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # 1. Load exclusion list
    exclusion_ids = load_exclusion_list(exclusion_file)
    
    # 2. Load original reports
    print("\n📊 Đọc báo cáo gốc...")
    df_c11 = load_c11_report(data_folder)
    df_c12 = load_c12_report(data_folder)
    print(f"  - C1.1: {len(df_c11)} đơn vị")
    print(f"  - C1.2: {len(df_c12)} đơn vị")
    
    # 3. Get exclusion stats
    print("\n📉 Đếm phiếu loại trừ theo đơn vị...")
    stats_c11 = get_exclusion_stats_c11(data_folder, exclusion_ids)
    stats_c12_sm1 = get_exclusion_stats_c12_sm1(data_folder, exclusion_ids)
    stats_c12_sm2 = get_exclusion_stats_c12_sm2(data_folder, exclusion_ids)
    
    print(f"  - C1.1 (SM4-C11): {stats_c11['c11_loai_tru_sm3'].sum() if len(stats_c11) > 0 else 0} phiếu")
    print(f"  - C1.2 SM1 (SM1-C12): {stats_c12_sm1['c12_loai_tru_sm1'].sum() if len(stats_c12_sm1) > 0 else 0} phiếu HLL")
    print(f"  - C1.2 SM2/SM3 (SM2-C12): {stats_c12_sm2['c12_loai_tru_sm2'].sum() if len(stats_c12_sm2) > 0 else 0} phiếu BH")
    
    # 4. Merge exclusion stats với C1.1
    print("\n🔧 Tính toán sau giảm trừ...")
    
    df_result = df_c11.merge(stats_c11, on='don_vi', how='left')
    df_result['c11_loai_tru_sm3'] = df_result['c11_loai_tru_sm3'].fillna(0).astype(int)
    df_result['c11_loai_tru_sm4'] = df_result['c11_loai_tru_sm4'].fillna(0).astype(int)
    
    # Tính C1.1 sau giảm trừ
    # Tỷ lệ = SM3 (đạt) / SM4 (tổng) * 100
    df_result['c11_sm3_sau'] = df_result['c11_sm3'] - df_result['c11_loai_tru_sm3']
    df_result['c11_sm4_sau'] = df_result['c11_sm4'] - df_result['c11_loai_tru_sm4']
    df_result['c11_tp2_ty_le_sau'] = df_result.apply(
        lambda r: round(r['c11_sm3_sau'] / r['c11_sm4_sau'] * 100, 2) if r['c11_sm4_sau'] > 0 else 0,
        axis=1
    )
    
    # Merge C1.2 với exclusion stats
    df_result = df_result.merge(df_c12[['don_vi', 'c12_sm1', 'c12_sm2', 'c12_tp1_ty_le', 
                                         'c12_sm3', 'c12_sm4', 'c12_tp2_ty_le', 'c12_bsc_goc']], 
                                on='don_vi', how='left')
    
    df_result = df_result.merge(stats_c12_sm1, on='don_vi', how='left')
    df_result = df_result.merge(stats_c12_sm2, on='don_vi', how='left')
    
    df_result['c12_loai_tru_sm1'] = df_result['c12_loai_tru_sm1'].fillna(0).astype(int)
    df_result['c12_loai_tru_sm2'] = df_result['c12_loai_tru_sm2'].fillna(0).astype(int)
    df_result['c12_loai_tru_sm3'] = df_result['c12_loai_tru_sm3'].fillna(0).astype(int)
    
    # Tính C1.2 sau giảm trừ
    df_result['c12_sm1_sau'] = df_result['c12_sm1'] - df_result['c12_loai_tru_sm1']
    df_result['c12_sm2_sau'] = df_result['c12_sm2'] - df_result['c12_loai_tru_sm2']
    df_result['c12_sm3_sau'] = df_result['c12_sm3'] - df_result['c12_loai_tru_sm3']
    # SM4 (tổng thuê bao) giữ nguyên
    df_result['c12_sm4_sau'] = df_result['c12_sm4']
    
    df_result['c12_tp1_ty_le_sau'] = df_result.apply(
        lambda r: round(r['c12_sm1_sau'] / r['c12_sm2_sau'] * 100, 2) if r['c12_sm2_sau'] > 0 else 0,
        axis=1
    )
    df_result['c12_tp2_ty_le_sau'] = df_result.apply(
        lambda r: round(r['c12_sm3_sau'] / r['c12_sm4_sau'] * 100, 2) if r['c12_sm4_sau'] > 0 else 0,
        axis=1
    )
    
    # 5. Tính điểm KPI
    print("\n📈 Tính điểm KPI...")
    
    # C1.1 TRƯỚC giảm trừ
    df_result['diem_c11_tp1_truoc'] = (df_result['c11_tp1_ty_le'] / 100).apply(tinh_diem_C11_TP1).round(2)
    df_result['diem_c11_tp2_truoc'] = (df_result['c11_tp2_ty_le'] / 100).apply(tinh_diem_C11_TP2).round(2)
    df_result['diem_c11_truoc'] = (df_result['diem_c11_tp1_truoc'] * 0.30 + df_result['diem_c11_tp2_truoc'] * 0.70).round(2)
    
    # C1.1 SAU giảm trừ (TP1 giữ nguyên, chỉ TP2 thay đổi)
    df_result['diem_c11_tp1_sau'] = df_result['diem_c11_tp1_truoc']  # Không thay đổi
    df_result['diem_c11_tp2_sau'] = (df_result['c11_tp2_ty_le_sau'] / 100).apply(tinh_diem_C11_TP2).round(2)
    df_result['diem_c11_sau'] = (df_result['diem_c11_tp1_sau'] * 0.30 + df_result['diem_c11_tp2_sau'] * 0.70).round(2)
    
    # C1.2 TRƯỚC giảm trừ
    df_result['diem_c12_tp1_truoc'] = (df_result['c12_tp1_ty_le'] / 100).apply(tinh_diem_C12_TP1).round(2)
    df_result['diem_c12_tp2_truoc'] = (df_result['c12_tp2_ty_le'] / 100).apply(tinh_diem_C12_TP2).round(2)
    df_result['diem_c12_truoc'] = (df_result['diem_c12_tp1_truoc'] * 0.50 + df_result['diem_c12_tp2_truoc'] * 0.50).round(2)
    
    # C1.2 SAU giảm trừ
    df_result['diem_c12_tp1_sau'] = (df_result['c12_tp1_ty_le_sau'] / 100).apply(tinh_diem_C12_TP1).round(2)
    df_result['diem_c12_tp2_sau'] = (df_result['c12_tp2_ty_le_sau'] / 100).apply(tinh_diem_C12_TP2).round(2)
    df_result['diem_c12_sau'] = (df_result['diem_c12_tp1_sau'] * 0.50 + df_result['diem_c12_tp2_sau'] * 0.50).round(2)
    
    # 6. Tính chênh lệch
    df_result['diem_c11_chenh_lech'] = (df_result['diem_c11_sau'] - df_result['diem_c11_truoc']).round(2)
    df_result['diem_c12_chenh_lech'] = (df_result['diem_c12_sau'] - df_result['diem_c12_truoc']).round(2)
    
    # 7. Tạo báo cáo tóm tắt
    df_summary = df_result[[
        'don_vi',
        # C1.1
        'c11_sm3', 'c11_sm4', 'c11_tp2_ty_le', 'c11_loai_tru_sm3', 'c11_loai_tru_sm4',
        'c11_sm3_sau', 'c11_sm4_sau', 'c11_tp2_ty_le_sau',
        'diem_c11_truoc', 'diem_c11_sau', 'diem_c11_chenh_lech',
        # C1.2
        'c12_sm1', 'c12_sm2', 'c12_tp1_ty_le', 'c12_sm3', 'c12_sm4', 'c12_tp2_ty_le',
        'c12_loai_tru_sm1', 'c12_loai_tru_sm2', 'c12_loai_tru_sm3',
        'c12_sm1_sau', 'c12_sm2_sau', 'c12_tp1_ty_le_sau', 'c12_sm3_sau', 'c12_tp2_ty_le_sau',
        'diem_c12_truoc', 'diem_c12_sau', 'diem_c12_chenh_lech'
    ]].copy()
    
    # Rename columns for clarity
    df_summary.columns = [
        'Đơn vị',
        # C1.1
        'C1.1_SM3_Thô', 'C1.1_SM4_Thô', 'C1.1_TyLe_Thô(%)', 'C1.1_LoạiTrừ_SM3', 'C1.1_LoạiTrừ_SM4',
        'C1.1_SM3_Sau', 'C1.1_SM4_Sau', 'C1.1_TyLe_Sau(%)',
        'C1.1_Điểm_Trước', 'C1.1_Điểm_Sau', 'C1.1_ChênhLệch',
        # C1.2
        'C1.2_SM1_Thô', 'C1.2_SM2_Thô', 'C1.2_TP1_Thô(%)', 'C1.2_SM3_Thô', 'C1.2_SM4_Thô', 'C1.2_TP2_Thô(%)',
        'C1.2_LoạiTrừ_SM1', 'C1.2_LoạiTrừ_SM2', 'C1.2_LoạiTrừ_SM3',
        'C1.2_SM1_Sau', 'C1.2_SM2_Sau', 'C1.2_TP1_Sau(%)', 'C1.2_SM3_Sau', 'C1.2_TP2_Sau(%)',
        'C1.2_Điểm_Trước', 'C1.2_Điểm_Sau', 'C1.2_ChênhLệch'
    ]
    
    # 8. Xuất file
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_folder / f"KPI_DonVi_SauGiamTru_{timestamp}.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
        
        # Sheet tóm tắt điểm
        df_scores = df_result[[
            'don_vi',
            'diem_c11_truoc', 'diem_c11_sau', 'diem_c11_chenh_lech',
            'diem_c12_truoc', 'diem_c12_sau', 'diem_c12_chenh_lech'
        ]].copy()
        df_scores.columns = [
            'Đơn vị',
            'C1.1_Trước', 'C1.1_Sau', 'C1.1_CL',
            'C1.2_Trước', 'C1.2_Sau', 'C1.2_CL'
        ]
        df_scores.to_excel(writer, sheet_name='Diem_KPI', index=False)
    
    print(f"\n✅ Đã xuất file: {output_file}")
    
    # 9. In kết quả
    print("\n" + "="*70)
    print("KẾT QUẢ SO SÁNH ĐIỂM KPI")
    print("="*70)
    
    print("\n📊 C1.1 - Chất lượng sửa chữa thuê bao BRCĐ:")
    print("-"*50)
    for _, row in df_result.iterrows():
        arrow = "↑" if row['diem_c11_chenh_lech'] > 0 else ("↓" if row['diem_c11_chenh_lech'] < 0 else "=")
        print(f"  {row['don_vi'][:30]:<30} : {row['diem_c11_truoc']:.2f} → {row['diem_c11_sau']:.2f} ({arrow}{abs(row['diem_c11_chenh_lech']):.2f})")
    
    print("\n📊 C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng:")
    print("-"*50)
    for _, row in df_result.iterrows():
        arrow = "↑" if row['diem_c12_chenh_lech'] > 0 else ("↓" if row['diem_c12_chenh_lech'] < 0 else "=")
        print(f"  {row['don_vi'][:30]:<30} : {row['diem_c12_truoc']:.2f} → {row['diem_c12_sau']:.2f} ({arrow}{abs(row['diem_c12_chenh_lech']):.2f})")
    
    return df_result


# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    DATA_FOLDER = "downloads/baocao_hanoi"
    EXCLUSION_FILE = "du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx"
    OUTPUT_FOLDER = "downloads/KPI"
    
    df = tinh_kpi_donvi_sau_giam_tru(DATA_FOLDER, EXCLUSION_FILE, OUTPUT_FOLDER)
    
    print("\n" + "="*70)
    print("HOÀN THÀNH!")
    print("="*70)
