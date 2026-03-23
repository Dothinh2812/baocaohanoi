#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script kiểm tra dữ liệu C1.2 để tìm lỗi sau giảm trừ
"""

import pandas as pd
import math
import os
import re

def normalize_id(id_val):
    """Chuẩn hóa BAOHONG_ID"""
    if pd.isna(id_val):
        return ""
    s = str(id_val).strip()
    s = re.sub(r'_x000[dD]_', '', s).strip()
    s = re.sub(r'[\r\n\t]', '', s)
    if s.endswith('.0'):
        s = s[:-2]
    return s

def extract_nvkt_name(ten_kv):
    """Trích xuất tên NVKT từ cột TEN_KV"""
    if pd.isna(ten_kv):
        return None
    ten_kv = str(ten_kv).strip()
    if '-' in ten_kv:
        parts = ten_kv.split('-')
        nvkt_name = parts[-1].strip()
    else:
        nvkt_name = ten_kv
    if '(' in nvkt_name:
        nvkt_name = nvkt_name.split('(')[0].strip()
    return nvkt_name

# Đọc danh sách loại trừ
exclusion_file = "du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx"
if os.path.exists(exclusion_file):
    df_excl = pd.read_excel(exclusion_file)
    exclusion_ids = {normalize_id(idx) for idx in df_excl['BAOHONG_ID'].tolist() if pd.notna(idx)}
    exclusion_ids.discard("")
    print(f"Đã đọc {len(exclusion_ids)} mã BAOHONG_ID từ danh sách loại trừ\n")
else:
    exclusion_ids = set()
    print("Không tìm thấy file loại trừ\n")

# Đọc SM1-C12
input_file_sm1 = "downloads/baocao_hanoi/SM1-C12.xlsx"
if not os.path.exists(input_file_sm1):
    print(f"❌ Không tìm thấy file: {input_file_sm1}")
    exit(1)

df_sm1_raw = pd.read_excel(input_file_sm1, sheet_name='Sheet1')
print(f"✅ Đã đọc SM1-C12: {len(df_sm1_raw)} dòng")

# Chuẩn hóa
df_sm1_raw['NVKT'] = df_sm1_raw['TEN_KV'].apply(extract_nvkt_name)
df_sm1_raw = df_sm1_raw[df_sm1_raw['NVKT'].notna()].copy()
df_sm1_raw['BAOHONG_ID_STR'] = df_sm1_raw['BAOHONG_ID'].apply(normalize_id)

# Lọc theo đội Quảng Oai
quang_oai_before = df_sm1_raw[df_sm1_raw['TEN_DOI'].str.contains('Quảng Oai', na=False)]
print(f"\n{'='*80}")
print(f"QUẢNG OAI - TRƯỚC GIẢM TRỪ")
print(f"{'='*80}")
print(f"Tổng số dòng: {len(quang_oai_before)}")
print("\nNhóm theo NVKT:")
for nvkt, group in quang_oai_before.groupby('NVKT'):
    count = len(group)
    hll = math.ceil(count / 2)
    print(f"  {nvkt}: {count} phiếu → HLL = ceil({count}/2) = {hll}")
    
    # Hiển thị các BAOHONG_ID
    ids = group['BAOHONG_ID_STR'].tolist()
    print(f"    BAOHONG_IDs: {', '.join(ids[:10])}{'...' if len(ids) > 10 else ''}")
    
    # Kiểm tra xem có ID nào trong danh sách loại trừ không
    excluded_ids = [id for id in ids if id in exclusion_ids]
    if excluded_ids:
        print(f"    ⚠️ Có {len(excluded_ids)} ID trong danh sách loại trừ: {', '.join(excluded_ids)}")

# Sau giảm trừ
df_sm1_excluded = df_sm1_raw[~df_sm1_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
quang_oai_after = df_sm1_excluded[df_sm1_excluded['TEN_DOI'].str.contains('Quảng Oai', na=False)]

print(f"\n{'='*80}")
print(f"QUẢNG OAI - SAU GIẢM TRỪ")
print(f"{'='*80}")
print(f"Tổng số dòng: {len(quang_oai_after)}")
print(f"Đã loại trừ: {len(quang_oai_before) - len(quang_oai_after)} phiếu")
print("\nNhóm theo NVKT:")
for nvkt, group in quang_oai_after.groupby('NVKT'):
    count = len(group)
    hll = math.ceil(count / 2)
    print(f"  {nvkt}: {count} phiếu → HLL = ceil({count}/2) = {hll}")

# So sánh
print(f"\n{'='*80}")
print(f"PHÂN TÍCH VẤN ĐỀ")
print(f"{'='*80}")
print("\nCông thức hiện tại: HLL = ceil(số_phiếu / 2)")
print("Vấn đề: Khi loại 1 phiếu từ số chẵn, kết quả làm tròn lên có thể làm HLL tăng!")
print("\nVí dụ:")
print("  - Trước GT: 6 phiếu → ceil(6/2) = 3 HLL")
print("  - Loại 1 phiếu: còn 5 phiếu → ceil(5/2) = 3 HLL")  
print("  - Loại 2 phiếu: còn 4 phiếu → ceil(4/2) = 2 HLL")
print("\nNhưng nếu:")
print("  - Trước GT: 6 phiếu → ceil(6/2) = 3 HLL")
print("  - Thêm 1 phiếu: có 7 phiếu → ceil(7/2) = 4 HLL ❌")
