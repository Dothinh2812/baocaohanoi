"""
Module tính điểm KPI BSC - Single source of truth
Tất cả hàm tính điểm BSC cho các chỉ tiêu C1.1, C1.2, C1.4, C1.5
được tập trung tại đây. Các module khác import từ file này.

Quy ước:
- Input luôn ở dạng thập phân (0-1), vd: 0.98 = 98%
- Dùng chuan_hoa_ty_le() để chuẩn hóa nếu input có thể là % (>1)
"""

import pandas as pd
import numpy as np


def chuan_hoa_ty_le(value):
    """
    Chuẩn hóa tỷ lệ về dạng thập phân (0-1).
    Nếu value > 1 thì chia cho 100 (giả định input là %).

    Args:
        value: Giá trị tỷ lệ (có thể là 98.5 hoặc 0.985)

    Returns:
        Giá trị dạng thập phân (0-1)
    """
    if pd.isna(value) or value is None:
        return value
    return value / 100 if value > 1 else value


def chuan_hoa_ty_le_df(df, col_ty_le):
    """
    Chuẩn hóa cột tỷ lệ trong DataFrame về dạng thập phân (0-1).
    Nếu giá trị max > 1 thì chia toàn cột cho 100.

    Args:
        df: DataFrame chứa cột cần chuẩn hóa
        col_ty_le: Tên cột tỷ lệ

    Returns:
        DataFrame đã chuẩn hóa (bản copy)
    """
    df = df.copy()
    if df[col_ty_le].max() > 1:
        df[col_ty_le] = df[col_ty_le] / 100
    return df


def chuan_hoa_ten(df, col_ten):
    """
    Chuẩn hóa tên NVKT về dạng Title Case.
    Xử lý trường hợp cùng 1 người nhập hoa/thường khác nhau.
    Ví dụ: "Bùi văn Cường" -> "Bùi Văn Cường"

    Args:
        df: DataFrame chứa cột tên
        col_ten: Tên cột cần chuẩn hóa

    Returns:
        DataFrame đã chuẩn hóa (bản copy)
    """
    df = df.copy()
    df[col_ten] = df[col_ten].str.strip().str.title()
    return df


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
        Điểm từ 1-5 hoặc NaN nếu không có dữ liệu
    """
    if pd.isna(kq) or kq is None:
        return np.nan

    if kq >= 0.995:
        return 5
    elif kq > 0.95:
        return 1 + 4 * (kq - 0.95) / 0.045
    else:
        return 1


def tinh_diem_C15(kq):
    """
    Tính điểm C1.5: Tỉ lệ thiết lập dịch vụ đạt thời gian quy định

    Args:
        kq: Tỉ lệ thiết lập dịch vụ đạt (dạng thập phân, vd: 0.995 = 99.5%)

    Returns:
        Điểm từ 1-5 hoặc NaN nếu không có dữ liệu

    Công thức:
        - KQ >= 99.5% = 5
        - 89.5% < KQ < 99.5% = 1 + 4*(KQ - 89.5%)/10%
        - KQ <= 89.5% = 1
    """
    if pd.isna(kq) or kq is None:
        return np.nan

    if kq >= 0.995:
        return 5
    elif kq > 0.895:
        return 1 + 4 * (kq - 0.895) / 0.10
    else:
        return 1


# ============================================================================
# SELF-TEST
# ============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print("KPI Scoring Module - Self Test")
    print("=" * 60)

    # Test chuan_hoa_ty_le
    assert chuan_hoa_ty_le(98.5) == 0.985, "chuan_hoa_ty_le(98.5) failed"
    assert chuan_hoa_ty_le(0.985) == 0.985, "chuan_hoa_ty_le(0.985) failed"
    assert chuan_hoa_ty_le(100) == 1.0, "chuan_hoa_ty_le(100) failed"
    assert pd.isna(chuan_hoa_ty_le(None)), "chuan_hoa_ty_le(None) failed"
    print("[PASS] chuan_hoa_ty_le")

    # Test C1.1 TP1
    assert tinh_diem_C11_TP1(0.99) == 5, "C11_TP1(0.99) failed"
    assert tinh_diem_C11_TP1(1.0) == 5, "C11_TP1(1.0) failed"
    assert tinh_diem_C11_TP1(0.96) == 1, "C11_TP1(0.96) failed"
    assert tinh_diem_C11_TP1(0.90) == 1, "C11_TP1(0.90) failed"
    assert tinh_diem_C11_TP1(None) == 5, "C11_TP1(None) failed"
    score = tinh_diem_C11_TP1(0.975)
    assert 1 < score < 5, f"C11_TP1(0.975) = {score}, expected between 1-5"
    print("[PASS] tinh_diem_C11_TP1")

    # Test C1.1 TP2
    assert tinh_diem_C11_TP2(0.85) == 5, "C11_TP2(0.85) failed"
    assert tinh_diem_C11_TP2(0.76) == 2, "C11_TP2(0.76) failed"
    assert tinh_diem_C11_TP2(0.70) == 1, "C11_TP2(0.70) failed"
    assert tinh_diem_C11_TP2(None) == 5, "C11_TP2(None) failed"
    print("[PASS] tinh_diem_C11_TP2")

    # Test C1.2 TP1 (lower is better)
    assert tinh_diem_C12_TP1(0.02) == 5, "C12_TP1(0.02) failed"
    assert tinh_diem_C12_TP1(0.025) == 5, "C12_TP1(0.025) failed"
    assert tinh_diem_C12_TP1(0.04) == 1, "C12_TP1(0.04) failed"
    assert tinh_diem_C12_TP1(0.05) == 1, "C12_TP1(0.05) failed"
    assert tinh_diem_C12_TP1(None) == 5, "C12_TP1(None) failed"
    print("[PASS] tinh_diem_C12_TP1")

    # Test C1.2 TP2 (lower is better)
    assert tinh_diem_C12_TP2(0.01) == 5, "C12_TP2(0.01) failed"
    assert tinh_diem_C12_TP2(0.02) == 5, "C12_TP2(0.02) failed"
    assert tinh_diem_C12_TP2(0.03) == 1, "C12_TP2(0.03) failed"
    assert tinh_diem_C12_TP2(None) == 5, "C12_TP2(None) failed"
    print("[PASS] tinh_diem_C12_TP2")

    # Test C1.4
    assert tinh_diem_C14(0.995) == 5, "C14(0.995) failed"
    assert tinh_diem_C14(0.95) == 1, "C14(0.95) failed"
    assert pd.isna(tinh_diem_C14(None)), "C14(None) failed"
    score = tinh_diem_C14(0.975)
    assert 1 < score < 5, f"C14(0.975) = {score}, expected between 1-5"
    print("[PASS] tinh_diem_C14")

    # Test C1.5
    assert tinh_diem_C15(0.995) == 5, "C15(0.995) failed"
    assert tinh_diem_C15(0.895) == 1, "C15(0.895) failed"
    assert pd.isna(tinh_diem_C15(None)), "C15(None) failed"
    score = tinh_diem_C15(0.95)
    assert 1 < score < 5, f"C15(0.95) = {score}, expected between 1-5"
    print("[PASS] tinh_diem_C15")

    print("\n" + "=" * 60)
    print("ALL TESTS PASSED")
    print("=" * 60)
