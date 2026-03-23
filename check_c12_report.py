#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script kiểm tra file So_sanh_C12_SM1.xlsx để xem dữ liệu đang như thế nào
"""

import pandas as pd
import os

# Đọc file so sánh
file_path = "downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx"

if not os.path.exists(file_path):
    print(f"❌ Không tìm thấy file: {file_path}")
    exit(1)

# Đọc các sheet
print("="*80)
print(f"KIỂM TRA FILE: {file_path}")
print("="*80)

# Sheet So_sanh_chi_tiet
try:
    df_detail = pd.read_excel(file_path, sheet_name='So_sanh_chi_tiet')
    print("\n📊 Sheet: So_sanh_chi_tiet")
    print(f"Số dòng: {len(df_detail)}")
    print("\nCác cột:")
    for col in df_detail.columns:
        print(f"  - {col}")
    
    # Lọc Quảng Oai
    if 'TEN_DOI' in df_detail.columns:
        quang_oai = df_detail[df_detail['TEN_DOI'].str.contains('Quảng Oai', na=False)]
        print(f"\n🔍 Dữ liệu QUẢNG OAI:")
        print(quang_oai.to_string())
    else:
        print("\n⚠️ Không có cột TEN_DOI")
except Exception as e:
    print(f"❌ Lỗi đọc sheet So_sanh_chi_tiet: {e}")

# Sheet Thong_ke_theo_don_vi
try:
    df_unit = pd.read_excel(file_path, sheet_name='Thong_ke_theo_don_vi')
    print("\n\n📊 Sheet: Thong_ke_theo_don_vi")
    print(f"Số dòng: {len(df_unit)}")
    print("\nDữ liệu:")
    print(df_unit.to_string())
except Exception as e:
    print(f"❌ Lỗi đọc sheet Thong_ke_theo_don_vi: {e}")

# Sheet TH_SM1C12_HLL_Thang trong file gốc
print("\n\n" + "="*80)
print("KIỂM TRA FILE GỐC: downloads/baocao_hanoi/SM1-C12.xlsx")
print("="*80)

try:
    df_origin = pd.read_excel("downloads/baocao_hanoi/SM1-C12.xlsx", sheet_name='TH_SM1C12_HLL_Thang')
    print("\n📊 Sheet: TH_SM1C12_HLL_Thang")
    print(f"Số dòng: {len(df_origin)}")
    
    # Lọc Quảng Oai
    if 'TEN_DOI' in df_origin.columns:
        quang_oai = df_origin[df_origin['TEN_DOI'].str.contains('Quảng Oai', na=False)]
        print(f"\n🔍 Dữ liệu QUẢNG OAI:")
        print(quang_oai.to_string())
except Exception as e:
    print(f"❌ Lỗi đọc file gốc: {e}")
