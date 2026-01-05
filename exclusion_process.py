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
        
        exclusion_ids = set(df['BAOHONG_ID'].astype(str).tolist())
        print(f"✅ Đã đọc {len(exclusion_ids)} mã BAOHONG_ID cần loại trừ")
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
        df_raw['BAOHONG_ID_STR'] = df_raw['BAOHONG_ID'].astype(str)
        df_excluded = df_raw[~df_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        
        num_excluded = len(df_raw) - len(df_excluded)
        print(f"✅ Đã loại trừ {num_excluded} phiếu, còn lại {len(df_excluded)} phiếu")
        
        # Tính thống kê TRƯỚC giảm trừ
        print("\n✓ Đang tính thống kê TRƯỚC giảm trừ...")
        df_stats_before = calculate_statistics(df_raw, has_ten_doi)
        
        # Tính thống kê SAU giảm trừ
        print("✓ Đang tính thống kê SAU giảm trừ...")
        df_stats_after = calculate_statistics(df_excluded, has_ten_doi)
        
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
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C11_SM4.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"✅ Đã tạo báo cáo so sánh C1.1 (SM4-C11)")
        print(f"   - Tổng phiếu thô: {tong_before}")
        print(f"   - Phiếu loại trừ: {num_excluded}")
        print(f"   - Tổng phiếu sau GT: {tong_after}")
        print(f"   - Tỷ lệ thô: {tyle_before}% -> Sau GT: {tyle_after}%")
        
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
        
        # Xác định phiếu đạt: TG <= 72 giờ
        df_raw['PHIEU_DAT'] = (df_raw['TG'] <= 72).astype(int)
        
        # Lọc dữ liệu sau giảm trừ
        df_raw['BAOHONG_ID_STR'] = df_raw['BAOHONG_ID'].astype(str)
        df_excluded = df_raw[~df_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        
        num_excluded = len(df_raw) - len(df_excluded)
        print(f"✅ Đã loại trừ {num_excluded} phiếu, còn lại {len(df_excluded)} phiếu")
        
        # Tính thống kê TRƯỚC giảm trừ
        print("\n✓ Đang tính thống kê TRƯỚC giảm trừ...")
        df_stats_before = calculate_statistics(df_raw, has_ten_doi, dat_column='PHIEU_DAT', dat_value=1)
        
        # Tính thống kê SAU giảm trừ
        print("✓ Đang tính thống kê SAU giảm trừ...")
        df_stats_after = calculate_statistics(df_excluded, has_ten_doi, dat_column='PHIEU_DAT', dat_value=1)
        
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
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C11_SM2.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"✅ Đã tạo báo cáo so sánh C1.1 SM2")
        print(f"   - Tổng phiếu thô: {tong_before}")
        print(f"   - Phiếu loại trừ: {num_excluded}")
        print(f"   - Tổng phiếu sau GT: {tong_after}")
        print(f"   - Tỷ lệ thô: {tyle_before}% -> Sau GT: {tyle_after}%")
        
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
        
        # ========== ĐỌC VÀ XỬ LÝ SM1-C12 ==========
        print("\n" + "-"*40)
        print("XỬ LÝ SM1-C12 (Phiếu hỏng lặp lại)")
        print("-"*40)
        
        df_sm1_raw = pd.read_excel(input_file_sm1, sheet_name='Sheet1')
        print(f"✅ Đã đọc SM1-C12 thô: {len(df_sm1_raw)} dòng")
        
        # Chuẩn hóa cột NVKT
        df_sm1_raw['NVKT'] = df_sm1_raw['TEN_KV'].apply(extract_nvkt_name)
        df_sm1_raw = df_sm1_raw[df_sm1_raw['NVKT'].notna()].copy()
        
        # Áp dụng giảm trừ SM1
        df_sm1_raw['BAOHONG_ID_STR'] = df_sm1_raw['BAOHONG_ID'].astype(str)
        df_sm1_excluded = df_sm1_raw[~df_sm1_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        num_excluded_sm1 = len(df_sm1_raw) - len(df_sm1_excluded)
        print(f"✅ Loại trừ SM1: {num_excluded_sm1} phiếu, còn lại {len(df_sm1_excluded)} phiếu")
        
        has_ten_doi = 'TEN_DOI' in df_sm1_raw.columns
        
        # Tính số phiếu HLL theo NVKT (TRƯỚC giảm trừ)
        def calculate_hll_by_nvkt(df, has_ten_doi):
            report_data = []
            if has_ten_doi:
                for (ten_doi, nvkt), group_df in df.groupby(['TEN_DOI', 'NVKT']):
                    if pd.isna(nvkt):
                        continue
                    so_phieu_hll = math.ceil(len(group_df) / 2)
                    report_data.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Số phiếu HLL': so_phieu_hll
                    })
            else:
                for nvkt, group_df in df.groupby('NVKT'):
                    if pd.isna(nvkt):
                        continue
                    so_phieu_hll = math.ceil(len(group_df) / 2)
                    report_data.append({
                        'NVKT': nvkt,
                        'Số phiếu HLL': so_phieu_hll
                    })
            return pd.DataFrame(report_data)
        
        df_hll_before = calculate_hll_by_nvkt(df_sm1_raw, has_ten_doi)
        df_hll_after = calculate_hll_by_nvkt(df_sm1_excluded, has_ten_doi)
        
        print(f"✅ HLL trước GT: {df_hll_before['Số phiếu HLL'].sum()} phiếu")
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
        
        # Áp dụng giảm trừ SM2
        df_sm2_raw['BAOHONG_ID_STR'] = df_sm2_raw['BAOHONG_ID'].astype(str)
        df_sm2_excluded = df_sm2_raw[~df_sm2_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        num_excluded_sm2 = len(df_sm2_raw) - len(df_sm2_excluded)
        print(f"✅ Loại trừ SM2: {num_excluded_sm2} phiếu, còn lại {len(df_sm2_excluded)} phiếu")
        
        has_ten_doi_sm2 = 'TEN_DOI' in df_sm2_raw.columns
        
        # Tính số phiếu báo hỏng theo NVKT
        def calculate_bh_by_nvkt(df, has_ten_doi):
            report_data = []
            if has_ten_doi:
                for (ten_doi, nvkt), group_df in df.groupby(['TEN_DOI', 'NVKT']):
                    if pd.isna(nvkt):
                        continue
                    report_data.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': len(group_df)
                    })
            else:
                for nvkt, group_df in df.groupby('NVKT'):
                    if pd.isna(nvkt):
                        continue
                    report_data.append({
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': len(group_df)
                    })
            return pd.DataFrame(report_data)
        
        df_bh_before = calculate_bh_by_nvkt(df_sm2_raw, has_ten_doi_sm2)
        df_bh_after = calculate_bh_by_nvkt(df_sm2_excluded, has_ten_doi_sm2)
        
        print(f"✅ Báo hỏng trước GT: {df_bh_before['Số phiếu báo hỏng'].sum()} phiếu")
        print(f"✅ Báo hỏng sau GT: {df_bh_after['Số phiếu báo hỏng'].sum()} phiếu")
        
        # ========== KẾT HỢP VÀ TÍNH TỶ LỆ ==========
        print("\n" + "-"*40)
        print("TẠO BÁO CÁO SO SÁNH")
        print("-"*40)
        
        # Merge HLL với báo hỏng (TRƯỚC giảm trừ)
        if has_ten_doi and has_ten_doi_sm2:
            merge_cols = ['TEN_DOI', 'NVKT']
        else:
            merge_cols = ['NVKT']
        
        df_before = pd.merge(df_hll_before, df_bh_before, on=merge_cols, how='outer')
        df_before['Số phiếu HLL'] = df_before['Số phiếu HLL'].fillna(0).astype(int)
        df_before['Số phiếu báo hỏng'] = df_before['Số phiếu báo hỏng'].fillna(0).astype(int)
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
            df_before, df_after,
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
            'Loại trừ SM2': num_excluded_sm2,
            'Tỷ lệ HLL % (Thô)': tyle_tho,
            'Tỷ lệ HLL % (Sau GT)': tyle_sau,
            'Thay đổi %': round(tyle_sau - tyle_tho, 2)
        }])
        
        # Lấy danh sách phiếu bị loại trừ từ cả SM1 và SM2
        df_loai_tru_sm1 = df_sm1_raw[df_sm1_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        df_loai_tru_sm2 = df_sm2_raw[df_sm2_raw['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
        
        cols_to_keep = ['BAOHONG_ID', 'MA_TB', 'TEN_TB', 'TEN_KV', 'TEN_DOI', 'NGAY_BAO_HONG']
        cols_sm1 = [c for c in cols_to_keep if c in df_loai_tru_sm1.columns]
        cols_sm2 = [c for c in cols_to_keep if c in df_loai_tru_sm2.columns]
        
        df_loai_tru_sm1 = df_loai_tru_sm1[cols_sm1]
        df_loai_tru_sm1['Nguồn'] = 'SM1-C12'
        df_loai_tru_sm2 = df_loai_tru_sm2[cols_sm2]
        df_loai_tru_sm2['Nguồn'] = 'SM2-C12'
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "So_sanh_C12_SM1.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru_sm1.to_excel(writer, sheet_name='DS_loai_tru_SM1', index=False)
            df_loai_tru_sm2.to_excel(writer, sheet_name='DS_loai_tru_SM2', index=False)
        
        print(f"\n✅ Đã tạo báo cáo so sánh C1.2")
        print(f"   - Phiếu HLL: {hll_tho} -> {hll_sau} (loại trừ SM1: {num_excluded_sm1})")
        print(f"   - Phiếu báo hỏng: {bh_tho} -> {bh_sau} (loại trừ SM2: {num_excluded_sm2})")
        print(f"   - Tỷ lệ HLL: {tyle_tho}% -> {tyle_sau}% (Δ: {round(tyle_sau - tyle_tho, 2)}%)")
        
        return {
            'chi_tieu': 'C1.2',
            'tong_tho': bh_tho,
            'loai_tru': num_excluded_sm1 + num_excluded_sm2,
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
        df_sm1['BAOHONG_ID_STR'] = df_sm1['BAOHONG_ID'].astype(str)
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
                # Số phiếu HLL = (số bản ghi) / 2, làm tròn lên
                so_phieu_hll = math.ceil(len(group_df) / 2)
                report_data_hll.append({
                    'TEN_DOI': ten_doi,
                    'NVKT': nvkt,
                    'Số phiếu HLL': so_phieu_hll
                })
        else:
            for nvkt, group_df in df_sm1_excluded.groupby('NVKT'):
                if pd.isna(nvkt):
                    continue
                so_phieu_hll = math.ceil(len(group_df) / 2)
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
            
            # ÁP DỤNG GIẢM TRỪ CHO SM2-C12
            df_sm2['BAOHONG_ID_STR'] = df_sm2['BAOHONG_ID'].astype(str)
            df_sm2_excluded = df_sm2[~df_sm2['BAOHONG_ID_STR'].isin(exclusion_ids)].copy()
            num_excluded_sm2 = len(df_sm2) - len(df_sm2_excluded)
            print(f"✅ Đã loại trừ {num_excluded_sm2} phiếu từ SM2-C12, còn lại {len(df_sm2_excluded)} phiếu")
            
            # Chuẩn hóa cột NVKT cho SM2 (sau giảm trừ)
            df_sm2_excluded['NVKT'] = df_sm2_excluded['TEN_KV'].apply(extract_nvkt_name)
            
            has_ten_doi_sm2 = 'TEN_DOI' in df_sm2_excluded.columns
            
            # Tính tổng phiếu báo hỏng theo NVKT (sau giảm trừ)
            report_data_bh = []
            
            if has_ten_doi_sm2:
                for (ten_doi, nvkt), group_df in df_sm2_excluded.groupby(['TEN_DOI', 'NVKT']):
                    if pd.isna(nvkt):
                        continue
                    report_data_bh.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': len(group_df)
                    })
            else:
                for nvkt, group_df in df_sm2_excluded.groupby('NVKT'):
                    if pd.isna(nvkt):
                        continue
                    report_data_bh.append({
                        'NVKT': nvkt,
                        'Số phiếu báo hỏng': len(group_df)
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
        
        has_ten_doi = 'TEN_DOI' in df_brcd.columns
        
        # Áp dụng giảm trừ
        df_brcd['BAOHONG_ID_STR'] = df_brcd['BAOHONG_ID'].astype(str)
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
        df_ref_clean = df_ref_clean.dropna(subset=['NVKT', 'Tổng TB'])
        
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
        
        # Ghi vào file Excel
        output_file = os.path.join(output_dir, "SM4-C12-ti-le-su-co-dv-brcd.xlsx")
        print(f"\n✓ Đang ghi kết quả vào: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='So_sanh_chi_tiet', index=False)
            df_tongke.to_excel(writer, sheet_name='Thong_ke_tong_hop', index=False)
            df_loai_tru.to_excel(writer, sheet_name='DS_phieu_loai_tru', index=False)
        
        print(f"\n✅ Đã tạo báo cáo so sánh C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng")
        print(f"   - Tổng thuê bao: {tb_tong}")
        print(f"   - Phiếu báo hỏng: {bh_tho} -> {bh_sau} (loại trừ: {num_excluded})")
        print(f"   - Tỷ lệ: {tyle_tho}% -> {tyle_sau}% (Δ: {round(tyle_sau - tyle_tho, 2)}%)")
        
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
        
        # C1.1 SM2
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
        
        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẠO BÁO CÁO SO SÁNH GIẢM TRỪ")
        print(f"   Kết quả được lưu tại: {output_dir}")
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
