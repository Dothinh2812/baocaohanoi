#!/usr/bin/env python3
"""
Script kiểm tra kết quả sau khi sửa logic giảm trừ
"""

import pandas as pd
import os

def verify_c11_sm4():
    """Kiểm tra C1.1 SM4"""
    print("\n" + "="*80)
    print("📊 KIỂM TRA C1.1 SM4 - SỬA CHỮA BÁO HỎNG")
    print("="*80)

    file_path = "downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx"

    # Kiểm tra sheet Thong_ke_tong_hop
    df_tong_hop = pd.read_excel(file_path, sheet_name='Thong_ke_tong_hop')
    print("\n📋 SHEET: Thong_ke_tong_hop")
    print(df_tong_hop.to_string(index=False))

    tong_truoc = df_tong_hop.iloc[0]['Tổng phiếu (Thô)']
    tong_sau = df_tong_hop.iloc[0]['Tổng phiếu (Sau GT)']

    print("\n🔍 KIỂM TRA LOGIC:")
    print(f"   Tổng phiếu TRƯỚC: {tong_truoc}")
    print(f"   Tổng phiếu SAU:   {tong_sau}")

    if tong_truoc == tong_sau:
        print("   ✅ MẪU SỐ ĐÃ ĐƯỢC GIỮ NGUYÊN - ĐÚNG!")
    else:
        print(f"   ❌ MẪU SỐ BỊ THAY ĐỔI - SAI! (Thay đổi: {tong_sau - tong_truoc})")

    # Kiểm tra sheet Thong_ke_theo_don_vi
    df_don_vi = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
    print("\n📋 SHEET: Thong_ke_theo_don_vi")
    print(df_don_vi[['don_vi', 'tong_phieu_truoc', 'tong_phieu_sau',
                      'phieu_dat_truoc', 'phieu_dat_sau',
                      'ty_le_truoc', 'ty_le_sau']].to_string(index=False))

    # Kiểm tra từng đơn vị
    print("\n🔍 KIỂM TRA TỪNG ĐƠN VỊ:")
    all_correct = True
    for _, row in df_don_vi.iterrows():
        if row['tong_phieu_truoc'] != row['tong_phieu_sau']:
            print(f"   ❌ {row['don_vi']}: Mẫu số bị thay đổi ({row['tong_phieu_truoc']} -> {row['tong_phieu_sau']})")
            all_correct = False

    if all_correct:
        print("   ✅ TẤT CẢ ĐƠN VỊ ĐỀU GIỮ NGUYÊN MẪU SỐ - ĐÚNG!")


def verify_c14():
    """Kiểm tra C1.4"""
    print("\n" + "="*80)
    print("📊 KIỂM TRA C1.4 - ĐỘ HÀI LÒNG KHÁCH HÀNG")
    print("="*80)

    file_path = "downloads/kq_sau_giam_tru/So_sanh_C14.xlsx"

    # Kiểm tra sheet Thong_ke_tong_hop
    df_tong_hop = pd.read_excel(file_path, sheet_name='Thong_ke_tong_hop')
    print("\n📋 SHEET: Thong_ke_tong_hop")
    print(df_tong_hop.to_string(index=False))

    tong_truoc = df_tong_hop.iloc[0]['Tổng KS (Thô)']
    tong_sau = df_tong_hop.iloc[0]['Tổng KS (Sau GT)']

    print("\n🔍 KIỂM TRA LOGIC:")
    print(f"   Tổng phiếu KS TRƯỚC: {tong_truoc}")
    print(f"   Tổng phiếu KS SAU:   {tong_sau}")

    if tong_truoc == tong_sau:
        print("   ✅ MẪU SỐ ĐÃ ĐƯỢC GIỮ NGUYÊN - ĐÚNG!")
    else:
        print(f"   ❌ MẪU SỐ BỊ THAY ĐỔI - SAI! (Thay đổi: {tong_sau - tong_truoc})")


def verify_c15():
    """Kiểm tra C1.5"""
    print("\n" + "="*80)
    print("📊 KIỂM TRA C1.5 - THIẾT LẬP DỊCH VỤ BRCĐ")
    print("="*80)

    file_path = "downloads/kq_sau_giam_tru/So_sanh_C15.xlsx"

    # Kiểm tra sheet Thong_ke_tong_hop
    df_tong_hop = pd.read_excel(file_path, sheet_name='Thong_ke_tong_hop')
    print("\n📋 SHEET: Thong_ke_tong_hop")
    print(df_tong_hop.to_string(index=False))

    tong_truoc = df_tong_hop.iloc[0]['Tổng Hoàn công (Thô)']
    tong_sau = df_tong_hop.iloc[0]['Tổng Hoàn công (Sau GT)']

    print("\n🔍 KIỂM TRA LOGIC:")
    print(f"   Tổng Hoàn công TRƯỚC: {tong_truoc}")
    print(f"   Tổng Hoàn công SAU:   {tong_sau}")

    if tong_truoc == tong_sau:
        print("   ✅ MẪU SỐ ĐÃ ĐƯỢC GIỮ NGUYÊN - ĐÚNG!")
    else:
        print(f"   ❌ MẪU SỐ BỊ THAY ĐỔI - SAI! (Thay đổi: {tong_sau - tong_truoc})")


if __name__ == "__main__":
    print("="*80)
    print("🔬 KIỂM TRA KẾT QUẢ SAU KHI SỬA LOGIC GIẢM TRỪ")
    print("="*80)

    verify_c11_sm4()
    verify_c14()
    verify_c15()

    print("\n" + "="*80)
    print("✅ HOÀN THÀNH KIỂM TRA")
    print("="*80)
