# -*- coding: utf-8 -*-
"""
Module tạo báo cáo xu hướng suy hao theo tuần/tháng/ngày
"""

import pandas as pd
import sqlite3
import os
from datetime import datetime, timedelta


def generate_daily_comparison_report(today_date=None, output_file=None):
    """
    Tạo báo cáo so sánh suy hao cao ngày hôm nay với ngày hôm qua
    
    Args:
        today_date: Ngày cần so sánh (format: 'YYYY-MM-DD' hoặc datetime). 
                   Mặc định lấy ngày mới nhất trong database.
        output_file: Đường dẫn file Excel output (tùy chọn)
    
    Returns:
        Đường dẫn file Excel đã tạo
    """
    print(f"\n{'='*80}")
    print(f"TẠO BÁO CÁO SO SÁNH SHC NGÀY")
    print(f"{'='*80}\n")
    
    db_path = "suy_hao_history.db"
    if not os.path.exists(db_path):
        print(f"❌ Không tìm thấy database: {db_path}")
        return None
    
    conn = sqlite3.connect(db_path)
    
    # Xác định ngày hôm nay và ngày hôm qua
    if today_date is None:
        # Lấy ngày mới nhất trong database
        df_latest = pd.read_sql_query("SELECT MAX(ngay_bao_cao) as latest FROM suy_hao_snapshots", conn)
        if df_latest.empty or df_latest['latest'][0] is None:
            print("❌ Không có dữ liệu trong database")
            conn.close()
            return None
        today_date = datetime.strptime(df_latest['latest'][0], '%Y-%m-%d')
    elif isinstance(today_date, str):
        today_date = datetime.strptime(today_date, '%Y-%m-%d')
    
    yesterday_date = today_date - timedelta(days=1)
    
    print(f"✓ Ngày hôm nay: {today_date.strftime('%d/%m/%Y')}")
    print(f"✓ Ngày hôm qua: {yesterday_date.strftime('%d/%m/%Y')}")
    
    # Lấy dữ liệu summary ngày hôm nay
    df_today = pd.read_sql_query(f"""
        SELECT 
            doi_one,
            nvkt_db_normalized,
            tong_so_hien_tai as so_luong_hom_nay,
            so_tang_moi as so_phat_sinh,
            so_giam_het as so_da_giam,
            so_van_con,
            so_tb_quan_ly,
            ty_le_shc as ty_le_hom_nay
        FROM suy_hao_daily_summary
        WHERE ngay_bao_cao = '{today_date.strftime('%Y-%m-%d')}'
    """, conn)
    
    # Lấy dữ liệu summary ngày hôm qua
    df_yesterday = pd.read_sql_query(f"""
        SELECT 
            doi_one,
            nvkt_db_normalized,
            tong_so_hien_tai as so_luong_hom_qua,
            ty_le_shc as ty_le_hom_qua
        FROM suy_hao_daily_summary
        WHERE ngay_bao_cao = '{yesterday_date.strftime('%Y-%m-%d')}'
    """, conn)
    
    print(f"\n✓ Dữ liệu hôm nay: {len(df_today)} NVKT")
    print(f"✓ Dữ liệu hôm qua: {len(df_yesterday)} NVKT")
    
    if len(df_today) == 0:
        print(f"❌ Không có dữ liệu ngày {today_date.strftime('%d/%m/%Y')}")
        conn.close()
        return None
    
    # Merge dữ liệu
    df_compare = df_today.merge(
        df_yesterday[['doi_one', 'nvkt_db_normalized', 'so_luong_hom_qua', 'ty_le_hom_qua']],
        on=['doi_one', 'nvkt_db_normalized'],
        how='outer'
    ).fillna(0)
    
    # Tính toán chênh lệch
    df_compare['chenh_lech'] = df_compare['so_luong_hom_nay'] - df_compare['so_luong_hom_qua']
    df_compare['chenh_lech_ty_le'] = (df_compare['ty_le_hom_nay'] - df_compare['ty_le_hom_qua']).round(2)
    
    # Tạo tên cột với ngày cụ thể
    yesterday_str = yesterday_date.strftime('%d/%m')
    today_str = today_date.strftime('%d/%m')
    
    col_sl_qua = f'SL {yesterday_str}'
    col_sl_nay = f'SL {today_str}'
    col_tl_qua = f'TL% {yesterday_str}'
    col_tl_nay = f'TL% {today_str}'
    
    # Đổi tên cột
    df_compare = df_compare.rename(columns={
        'doi_one': 'Đơn vị',
        'nvkt_db_normalized': 'NVKT',
        'so_luong_hom_qua': col_sl_qua,
        'so_luong_hom_nay': col_sl_nay,
        'so_phat_sinh': 'Phát sinh',
        'so_da_giam': 'Đã giảm',
        'so_van_con': 'Vẫn còn',
        'chenh_lech': '+/- SL',
        'so_tb_quan_ly': 'TB Quản lý',
        'ty_le_hom_qua': col_tl_qua,
        'ty_le_hom_nay': col_tl_nay,
        'chenh_lech_ty_le': '+/- TL%'
    })
    
    # Sắp xếp cột
    columns_order = ['Đơn vị', 'NVKT', col_sl_qua, col_sl_nay, '+/- SL', 
                     'Phát sinh', 'Đã giảm', 'Vẫn còn', 'TB Quản lý', 
                     col_tl_qua, col_tl_nay, '+/- TL%']
    df_compare = df_compare[[c for c in columns_order if c in df_compare.columns]]
    df_compare = df_compare.sort_values(by=['Đơn vị', 'NVKT'])
    
    # Tổng hợp theo đơn vị
    df_by_unit = df_compare.groupby('Đơn vị').agg({
        col_sl_qua: 'sum',
        col_sl_nay: 'sum',
        'Phát sinh': 'sum',
        'Đã giảm': 'sum',
        'Vẫn còn': 'sum',
        'TB Quản lý': 'sum'
    }).reset_index()
    
    df_by_unit['+/- SL'] = df_by_unit[col_sl_nay] - df_by_unit[col_sl_qua]
    df_by_unit[col_tl_qua] = (df_by_unit[col_sl_qua] / df_by_unit['TB Quản lý'] * 100).round(2)
    df_by_unit[col_tl_nay] = (df_by_unit[col_sl_nay] / df_by_unit['TB Quản lý'] * 100).round(2)
    df_by_unit['+/- TL%'] = (df_by_unit[col_tl_nay] - df_by_unit[col_tl_qua]).round(2)
    
    # Xử lý inf/nan
    df_by_unit = df_by_unit.replace([float('inf'), -float('inf')], 0).fillna(0)
    
    # Thêm dòng tổng
    total_row = pd.DataFrame({
        'Đơn vị': ['TỔNG CỘNG'],
        col_sl_qua: [df_by_unit[col_sl_qua].sum()],
        col_sl_nay: [df_by_unit[col_sl_nay].sum()],
        '+/- SL': [df_by_unit['+/- SL'].sum()],
        'Phát sinh': [df_by_unit['Phát sinh'].sum()],
        'Đã giảm': [df_by_unit['Đã giảm'].sum()],
        'Vẫn còn': [df_by_unit['Vẫn còn'].sum()],
        'TB Quản lý': [df_by_unit['TB Quản lý'].sum()],
        col_tl_qua: [round(df_by_unit[col_sl_qua].sum() / df_by_unit['TB Quản lý'].sum() * 100, 2) if df_by_unit['TB Quản lý'].sum() > 0 else 0],
        col_tl_nay: [round(df_by_unit[col_sl_nay].sum() / df_by_unit['TB Quản lý'].sum() * 100, 2) if df_by_unit['TB Quản lý'].sum() > 0 else 0],
        '+/- TL%': [0]
    })
    total_row['+/- TL%'] = total_row[col_tl_nay] - total_row[col_tl_qua]
    df_by_unit = pd.concat([df_by_unit, total_row], ignore_index=True)
    
    # Sắp xếp cột cho df_by_unit
    unit_columns = ['Đơn vị', col_sl_qua, col_sl_nay, '+/- SL', 
                    'Phát sinh', 'Đã giảm', 'Vẫn còn', 'TB Quản lý',
                    col_tl_qua, col_tl_nay, '+/- TL%']
    df_by_unit = df_by_unit[[c for c in unit_columns if c in df_by_unit.columns]]
    
    conn.close()
    
    # In thống kê
    print(f"\n{'─'*60}")
    print(f"📊 THỐNG KÊ TỔNG HỢP:")
    print(f"   SL {yesterday_str}: {int(total_row[col_sl_qua].iloc[0])}")
    print(f"   SL {today_str}: {int(total_row[col_sl_nay].iloc[0])}")
    print(f"   Phát sinh:  +{int(total_row['Phát sinh'].iloc[0])}")
    print(f"   Đã giảm:    -{int(total_row['Đã giảm'].iloc[0])}")
    print(f"   Chênh lệch: {int(total_row['+/- SL'].iloc[0]):+d}")
    print(f"{'─'*60}")
    
    # Tạo file Excel
    if output_file is None:
        output_file = "downloads/baocao_hanoi/So_sanh_SHC_theo_ngay_T-1.xlsx"
    
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    print(f"\n✓ Đang ghi file Excel: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet tổng hợp theo đơn vị
        df_by_unit.to_excel(writer, sheet_name='Theo_don_vi', index=False)
        
        # Sheet chi tiết theo NVKT
        df_compare.to_excel(writer, sheet_name='Chi_tiet_NVKT', index=False)
    
    print(f"✅ Đã tạo báo cáo so sánh: {output_file}")
    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH BÁO CÁO SO SÁNH SHC NGÀY")
    print(f"{'='*80}\n")
    
    return output_file


def generate_daily_comparison_report_k2(today_date=None, output_file=None):
    """
    Tạo báo cáo so sánh suy hao cao K2 ngày hôm nay với ngày hôm qua
    Sử dụng database suy_hao_history_k2.db
    
    Args:
        today_date: Ngày cần so sánh (format: 'YYYY-MM-DD' hoặc datetime). 
                   Mặc định lấy ngày mới nhất trong database.
        output_file: Đường dẫn file Excel output (tùy chọn)
    
    Returns:
        Đường dẫn file Excel đã tạo
    """
    print(f"\n{'='*80}")
    print(f"TẠO BÁO CÁO SO SÁNH SHC K2 NGÀY")
    print(f"{'='*80}\n")
    
    db_path = "suy_hao_history_k2.db"
    if not os.path.exists(db_path):
        print(f"❌ Không tìm thấy database: {db_path}")
        return None
    
    conn = sqlite3.connect(db_path)
    
    # Xác định ngày hôm nay và ngày hôm qua
    if today_date is None:
        # Lấy ngày mới nhất trong database
        df_latest = pd.read_sql_query("SELECT MAX(ngay_bao_cao) as latest FROM suy_hao_snapshots", conn)
        if df_latest.empty or df_latest['latest'][0] is None:
            print("❌ Không có dữ liệu trong database K2")
            conn.close()
            return None
        today_date = datetime.strptime(df_latest['latest'][0], '%Y-%m-%d')
    elif isinstance(today_date, str):
        today_date = datetime.strptime(today_date, '%Y-%m-%d')
    
    yesterday_date = today_date - timedelta(days=1)
    
    print(f"✓ Ngày hôm nay: {today_date.strftime('%d/%m/%Y')}")
    print(f"✓ Ngày hôm qua: {yesterday_date.strftime('%d/%m/%Y')}")
    
    # Lấy dữ liệu summary ngày hôm nay
    df_today = pd.read_sql_query(f"""
        SELECT 
            doi_one,
            nvkt_db_normalized,
            tong_so_hien_tai as so_luong_hom_nay,
            so_tang_moi as so_phat_sinh,
            so_giam_het as so_da_giam,
            so_van_con,
            so_tb_quan_ly,
            ty_le_shc as ty_le_hom_nay
        FROM suy_hao_daily_summary
        WHERE ngay_bao_cao = '{today_date.strftime('%Y-%m-%d')}'
    """, conn)
    
    # Lấy dữ liệu summary ngày hôm qua
    df_yesterday = pd.read_sql_query(f"""
        SELECT 
            doi_one,
            nvkt_db_normalized,
            tong_so_hien_tai as so_luong_hom_qua,
            ty_le_shc as ty_le_hom_qua
        FROM suy_hao_daily_summary
        WHERE ngay_bao_cao = '{yesterday_date.strftime('%Y-%m-%d')}'
    """, conn)
    
    print(f"\n✓ Dữ liệu hôm nay: {len(df_today)} NVKT")
    print(f"✓ Dữ liệu hôm qua: {len(df_yesterday)} NVKT")
    
    if len(df_today) == 0:
        print(f"❌ Không có dữ liệu ngày {today_date.strftime('%d/%m/%Y')}")
        conn.close()
        return None
    
    # Merge dữ liệu
    df_compare = df_today.merge(
        df_yesterday[['doi_one', 'nvkt_db_normalized', 'so_luong_hom_qua', 'ty_le_hom_qua']],
        on=['doi_one', 'nvkt_db_normalized'],
        how='outer'
    ).fillna(0)
    
    # Tính toán chênh lệch
    df_compare['chenh_lech'] = df_compare['so_luong_hom_nay'] - df_compare['so_luong_hom_qua']
    df_compare['chenh_lech_ty_le'] = (df_compare['ty_le_hom_nay'] - df_compare['ty_le_hom_qua']).round(2)
    
    # Tạo tên cột với ngày cụ thể
    yesterday_str = yesterday_date.strftime('%d/%m')
    today_str = today_date.strftime('%d/%m')
    
    col_sl_qua = f'SL {yesterday_str}'
    col_sl_nay = f'SL {today_str}'
    col_tl_qua = f'TL% {yesterday_str}'
    col_tl_nay = f'TL% {today_str}'
    
    # Đổi tên cột
    df_compare = df_compare.rename(columns={
        'doi_one': 'Đơn vị',
        'nvkt_db_normalized': 'NVKT',
        'so_luong_hom_qua': col_sl_qua,
        'so_luong_hom_nay': col_sl_nay,
        'so_phat_sinh': 'Phát sinh',
        'so_da_giam': 'Đã giảm',
        'so_van_con': 'Vẫn còn',
        'chenh_lech': '+/- SL',
        'so_tb_quan_ly': 'TB Quản lý',
        'ty_le_hom_qua': col_tl_qua,
        'ty_le_hom_nay': col_tl_nay,
        'chenh_lech_ty_le': '+/- TL%'
    })
    
    # Sắp xếp cột
    columns_order = ['Đơn vị', 'NVKT', col_sl_qua, col_sl_nay, '+/- SL', 
                     'Phát sinh', 'Đã giảm', 'Vẫn còn', 'TB Quản lý', 
                     col_tl_qua, col_tl_nay, '+/- TL%']
    df_compare = df_compare[[c for c in columns_order if c in df_compare.columns]]
    df_compare = df_compare.sort_values(by=['Đơn vị', 'NVKT'])
    
    # Tổng hợp theo đơn vị
    df_by_unit = df_compare.groupby('Đơn vị').agg({
        col_sl_qua: 'sum',
        col_sl_nay: 'sum',
        'Phát sinh': 'sum',
        'Đã giảm': 'sum',
        'Vẫn còn': 'sum',
        'TB Quản lý': 'sum'
    }).reset_index()
    
    df_by_unit['+/- SL'] = df_by_unit[col_sl_nay] - df_by_unit[col_sl_qua]
    df_by_unit[col_tl_qua] = (df_by_unit[col_sl_qua] / df_by_unit['TB Quản lý'] * 100).round(2)
    df_by_unit[col_tl_nay] = (df_by_unit[col_sl_nay] / df_by_unit['TB Quản lý'] * 100).round(2)
    df_by_unit['+/- TL%'] = (df_by_unit[col_tl_nay] - df_by_unit[col_tl_qua]).round(2)
    
    # Xử lý inf/nan
    df_by_unit = df_by_unit.replace([float('inf'), -float('inf')], 0).fillna(0)
    
    # Thêm dòng tổng
    total_row = pd.DataFrame({
        'Đơn vị': ['TỔNG CỘNG'],
        col_sl_qua: [df_by_unit[col_sl_qua].sum()],
        col_sl_nay: [df_by_unit[col_sl_nay].sum()],
        '+/- SL': [df_by_unit['+/- SL'].sum()],
        'Phát sinh': [df_by_unit['Phát sinh'].sum()],
        'Đã giảm': [df_by_unit['Đã giảm'].sum()],
        'Vẫn còn': [df_by_unit['Vẫn còn'].sum()],
        'TB Quản lý': [df_by_unit['TB Quản lý'].sum()],
        col_tl_qua: [round(df_by_unit[col_sl_qua].sum() / df_by_unit['TB Quản lý'].sum() * 100, 2) if df_by_unit['TB Quản lý'].sum() > 0 else 0],
        col_tl_nay: [round(df_by_unit[col_sl_nay].sum() / df_by_unit['TB Quản lý'].sum() * 100, 2) if df_by_unit['TB Quản lý'].sum() > 0 else 0],
        '+/- TL%': [0]
    })
    total_row['+/- TL%'] = total_row[col_tl_nay] - total_row[col_tl_qua]
    df_by_unit = pd.concat([df_by_unit, total_row], ignore_index=True)
    
    # Sắp xếp cột cho df_by_unit
    unit_columns = ['Đơn vị', col_sl_qua, col_sl_nay, '+/- SL', 
                    'Phát sinh', 'Đã giảm', 'Vẫn còn', 'TB Quản lý',
                    col_tl_qua, col_tl_nay, '+/- TL%']
    df_by_unit = df_by_unit[[c for c in unit_columns if c in df_by_unit.columns]]
    
    conn.close()
    
    # In thống kê
    print(f"\n{'─'*60}")
    print(f"📊 THỐNG KÊ TỔNG HỢP K2:")
    print(f"   SL {yesterday_str}: {int(total_row[col_sl_qua].iloc[0])}")
    print(f"   SL {today_str}: {int(total_row[col_sl_nay].iloc[0])}")
    print(f"   Phát sinh:  +{int(total_row['Phát sinh'].iloc[0])}")
    print(f"   Đã giảm:    -{int(total_row['Đã giảm'].iloc[0])}")
    print(f"   Chênh lệch: {int(total_row['+/- SL'].iloc[0]):+d}")
    print(f"{'─'*60}")
    
    # Tạo file Excel
    if output_file is None:
        output_file = "downloads/baocao_hanoi/So_sanh_SHC_k2_theo_ngay_T-1.xlsx"
    
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    print(f"\n✓ Đang ghi file Excel: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet tổng hợp theo đơn vị
        df_by_unit.to_excel(writer, sheet_name='Theo_don_vi', index=False)
        
        # Sheet chi tiết theo NVKT
        df_compare.to_excel(writer, sheet_name='Chi_tiet_NVKT', index=False)
    
    print(f"✅ Đã tạo báo cáo so sánh K2: {output_file}")
    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH BÁO CÁO SO SÁNH SHC K2 NGÀY")
    print(f"{'='*80}\n")
    
    return output_file

def generate_weekly_report(year, week_number, output_file=None):
    """
    Tạo báo cáo so sánh theo tuần

    Args:
        year: Năm (VD: 2025)
        week_number: Số tuần trong năm (1-52)
        output_file: Đường dẫn file Excel output (tùy chọn)

    Returns:
        Đường dẫn file Excel đã tạo
    """
    print(f"\n{'='*80}")
    print(f"TẠO BÁO CÁO TUẦN {week_number}/{year}")
    print(f"{'='*80}\n")

    db_path = "suy_hao_history.db"
    if not os.path.exists(db_path):
        print(f"❌ Không tìm thấy database: {db_path}")
        return None

    conn = sqlite3.connect(db_path)

    # Tính ngày bắt đầu và kết thúc tuần
    week_start = datetime.strptime(f'{year}-W{week_number:02d}-1', '%Y-W%W-%w')
    week_end = week_start + timedelta(days=6)

    # Tuần trước
    prev_week_start = week_start - timedelta(days=7)
    prev_week_end = week_end - timedelta(days=7)

    print(f"✓ Tuần hiện tại: {week_start.strftime('%d/%m/%Y')} - {week_end.strftime('%d/%m/%Y')}")
    print(f"✓ Tuần trước: {prev_week_start.strftime('%d/%m/%Y')} - {prev_week_end.strftime('%d/%m/%Y')}")

    # Lấy danh sách thuê bao tuần hiện tại
    df_current = pd.read_sql_query(f"""
        SELECT DISTINCT account_cts, doi_one, nvkt_db_normalized, sa, ten_tb_one
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao BETWEEN '{week_start.strftime('%Y-%m-%d')}'
          AND '{week_end.strftime('%Y-%m-%d')}'
    """, conn)

    # Lấy danh sách thuê bao tuần trước
    df_previous = pd.read_sql_query(f"""
        SELECT DISTINCT account_cts, doi_one, nvkt_db_normalized, sa, ten_tb_one
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao BETWEEN '{prev_week_start.strftime('%Y-%m-%d')}'
          AND '{prev_week_end.strftime('%Y-%m-%d')}'
    """, conn)

    print(f"\n✓ Tuần hiện tại: {len(df_current)} thuê bao")
    print(f"✓ Tuần trước: {len(df_previous)} thuê bao")

    # Phân loại biến động
    accounts_current = set(df_current['account_cts'].tolist())
    accounts_previous = set(df_previous['account_cts'].tolist())

    tang_moi = accounts_current - accounts_previous
    giam_het = accounts_previous - accounts_current
    van_con = accounts_current & accounts_previous

    print(f"\n✓ Phân tích biến động:")
    print(f"  🆕 TĂNG MỚI: {len(tang_moi)} thuê bao")
    print(f"  ⬇️  GIẢM/HẾT: {len(giam_het)} thuê bao")
    print(f"  ↔️  VẪN CÒN: {len(van_con)} thuê bao")

    # Tạo DataFrames
    df_tang = df_current[df_current['account_cts'].isin(tang_moi)].copy()
    df_giam = df_previous[df_previous['account_cts'].isin(giam_het)].copy()
    df_van = df_current[df_current['account_cts'].isin(van_con)].copy()

    # Tổng hợp theo NVKT_DB
    summary_current = df_current.groupby(['doi_one', 'nvkt_db_normalized']).size().reset_index(name='so_luong_tuan_nay')
    summary_previous = df_previous.groupby(['doi_one', 'nvkt_db_normalized']).size().reset_index(name='so_luong_tuan_truoc')

    df_summary = summary_current.merge(
        summary_previous,
        on=['doi_one', 'nvkt_db_normalized'],
        how='outer'
    ).fillna(0)

    df_summary['tang_giam'] = df_summary['so_luong_tuan_nay'] - df_summary['so_luong_tuan_truoc']
    df_summary['ty_le_thay_doi'] = (df_summary['tang_giam'] / df_summary['so_luong_tuan_truoc'] * 100).round(1)
    df_summary['ty_le_thay_doi'] = df_summary['ty_le_thay_doi'].replace([float('inf'), -float('inf')], 0)

    # Lấy thông tin tỉ lệ SHC từ suy_hao_daily_summary (ngày cuối tuần)
    df_ratio = pd.read_sql_query(f"""
        SELECT doi_one, nvkt_db_normalized, so_tb_quan_ly, ty_le_shc
        FROM suy_hao_daily_summary
        WHERE ngay_bao_cao = (SELECT MAX(ngay_bao_cao) FROM suy_hao_daily_summary 
                              WHERE ngay_bao_cao <= '{week_end.strftime('%Y-%m-%d')}')
    """, conn)
    
    if len(df_ratio) > 0:
        df_summary = df_summary.merge(
            df_ratio,
            on=['doi_one', 'nvkt_db_normalized'],
            how='left'
        )
        df_summary['so_tb_quan_ly'] = df_summary['so_tb_quan_ly'].fillna(0).astype(int)
        df_summary['ty_le_shc'] = df_summary['ty_le_shc'].fillna(0)
        df_summary.columns = ['Đơn vị', 'NVKT_DB', 'Tuần này', 'Tuần trước', 'Tăng/Giảm', '% Thay đổi', 'Số TB quản lý', 'Tỉ lệ SHC (%)']
    else:
        df_summary.columns = ['Đơn vị', 'NVKT_DB', 'Tuần này', 'Tuần trước', 'Tăng/Giảm', '% Thay đổi']
    
    df_summary = df_summary.sort_values(by=['Đơn vị', 'NVKT_DB'])

    conn.close()

    # Tạo file Excel
    if output_file is None:
        output_file = f"downloads/baocao_hanoi/Bao_cao_tuan_{week_number}_{year}.xlsx"

    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    print(f"\n✓ Đang ghi file Excel: {output_file}")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet tổng hợp
        df_summary.to_excel(writer, sheet_name='Tong_hop', index=False)

        # Sheet chi tiết TĂNG MỚI
        if len(df_tang) > 0:
            df_tang[['account_cts', 'ten_tb_one', 'doi_one', 'nvkt_db_normalized', 'sa']].to_excel(
                writer, sheet_name='Tang_moi', index=False
            )

        # Sheet chi tiết GIẢM/HẾT
        if len(df_giam) > 0:
            df_giam[['account_cts', 'ten_tb_one', 'doi_one', 'nvkt_db_normalized', 'sa']].to_excel(
                writer, sheet_name='Giam_het', index=False
            )

        # Sheet VẪN CÒN
        if len(df_van) > 0:
            df_van[['account_cts', 'ten_tb_one', 'doi_one', 'nvkt_db_normalized', 'sa']].to_excel(
                writer, sheet_name='Van_con', index=False
            )

    print(f"✅ Đã tạo báo cáo tuần: {output_file}")
    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH BÁO CÁO TUẦN {week_number}/{year}")
    print(f"{'='*80}\n")

    return output_file


def generate_monthly_report(year, month, output_file=None):
    """
    Tạo báo cáo so sánh theo tháng

    Args:
        year: Năm (VD: 2025)
        month: Tháng (1-12)
        output_file: Đường dẫn file Excel output (tùy chọn)

    Returns:
        Đường dẫn file Excel đã tạo
    """
    print(f"\n{'='*80}")
    print(f"TẠO BÁO CÁO THÁNG {month}/{year}")
    print(f"{'='*80}\n")

    db_path = "suy_hao_history.db"
    if not os.path.exists(db_path):
        print(f"❌ Không tìm thấy database: {db_path}")
        return None

    conn = sqlite3.connect(db_path)

    # Tháng hiện tại
    month_start = datetime(year, month, 1)
    if month == 12:
        month_end = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        month_end = datetime(year, month + 1, 1) - timedelta(days=1)

    # Tháng trước
    if month == 1:
        prev_month_start = datetime(year - 1, 12, 1)
        prev_month_end = datetime(year, 1, 1) - timedelta(days=1)
    else:
        prev_month_start = datetime(year, month - 1, 1)
        prev_month_end = datetime(year, month, 1) - timedelta(days=1)

    print(f"✓ Tháng hiện tại: {month_start.strftime('%d/%m/%Y')} - {month_end.strftime('%d/%m/%Y')}")
    print(f"✓ Tháng trước: {prev_month_start.strftime('%d/%m/%Y')} - {prev_month_end.strftime('%d/%m/%Y')}")

    # Lấy danh sách thuê bao tháng hiện tại
    df_current = pd.read_sql_query(f"""
        SELECT DISTINCT account_cts, doi_one, nvkt_db_normalized, sa, ten_tb_one
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao BETWEEN '{month_start.strftime('%Y-%m-%d')}'
          AND '{month_end.strftime('%Y-%m-%d')}'
    """, conn)

    # Lấy danh sách thuê bao tháng trước
    df_previous = pd.read_sql_query(f"""
        SELECT DISTINCT account_cts, doi_one, nvkt_db_normalized, sa, ten_tb_one
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao BETWEEN '{prev_month_start.strftime('%Y-%m-%d')}'
          AND '{prev_month_end.strftime('%Y-%m-%d')}'
    """, conn)

    print(f"\n✓ Tháng hiện tại: {len(df_current)} thuê bao")
    print(f"✓ Tháng trước: {len(df_previous)} thuê bao")

    # Phân loại biến động
    accounts_current = set(df_current['account_cts'].tolist())
    accounts_previous = set(df_previous['account_cts'].tolist())

    tang_moi = accounts_current - accounts_previous
    giam_het = accounts_previous - accounts_current
    van_con = accounts_current & accounts_previous

    print(f"\n✓ Phân tích biến động:")
    print(f"  🆕 TĂNG MỚI: {len(tang_moi)} thuê bao")
    print(f"  ⬇️  GIẢM/HẾT: {len(giam_het)} thuê bao")
    print(f"  ↔️  VẪN CÒN: {len(van_con)} thuê bao")

    # Tạo DataFrames
    df_tang = df_current[df_current['account_cts'].isin(tang_moi)].copy()
    df_giam = df_previous[df_previous['account_cts'].isin(giam_het)].copy()
    df_van = df_current[df_current['account_cts'].isin(van_con)].copy()

    # Tổng hợp theo NVKT_DB
    summary_current = df_current.groupby(['doi_one', 'nvkt_db_normalized']).size().reset_index(name='so_luong_thang_nay')
    summary_previous = df_previous.groupby(['doi_one', 'nvkt_db_normalized']).size().reset_index(name='so_luong_thang_truoc')

    df_summary = summary_current.merge(
        summary_previous,
        on=['doi_one', 'nvkt_db_normalized'],
        how='outer'
    ).fillna(0)

    df_summary['tang_giam'] = df_summary['so_luong_thang_nay'] - df_summary['so_luong_thang_truoc']
    df_summary['ty_le_thay_doi'] = (df_summary['tang_giam'] / df_summary['so_luong_thang_truoc'] * 100).round(1)
    df_summary['ty_le_thay_doi'] = df_summary['ty_le_thay_doi'].replace([float('inf'), -float('inf')], 0)

    # Lấy thông tin tỉ lệ SHC từ suy_hao_daily_summary (ngày cuối tháng)
    df_ratio = pd.read_sql_query(f"""
        SELECT doi_one, nvkt_db_normalized, so_tb_quan_ly, ty_le_shc
        FROM suy_hao_daily_summary
        WHERE ngay_bao_cao = (SELECT MAX(ngay_bao_cao) FROM suy_hao_daily_summary 
                              WHERE ngay_bao_cao <= '{month_end.strftime('%Y-%m-%d')}')
    """, conn)
    
    if len(df_ratio) > 0:
        df_summary = df_summary.merge(
            df_ratio,
            on=['doi_one', 'nvkt_db_normalized'],
            how='left'
        )
        df_summary['so_tb_quan_ly'] = df_summary['so_tb_quan_ly'].fillna(0).astype(int)
        df_summary['ty_le_shc'] = df_summary['ty_le_shc'].fillna(0)
        df_summary.columns = ['Đơn vị', 'NVKT_DB', 'Tháng này', 'Tháng trước', 'Tăng/Giảm', '% Thay đổi', 'Số TB quản lý', 'Tỉ lệ SHC (%)']
    else:
        df_summary.columns = ['Đơn vị', 'NVKT_DB', 'Tháng này', 'Tháng trước', 'Tăng/Giảm', '% Thay đổi']
    
    df_summary = df_summary.sort_values(by=['Đơn vị', 'NVKT_DB'])

    # Thống kê theo ngày trong tháng
    print(f"\n✓ Tạo xu hướng theo ngày...")
    df_daily = pd.read_sql_query(f"""
        SELECT
            ngay_bao_cao,
            COUNT(DISTINCT account_cts) as so_luong
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao BETWEEN '{month_start.strftime('%Y-%m-%d')}'
          AND '{month_end.strftime('%Y-%m-%d')}'
        GROUP BY ngay_bao_cao
        ORDER BY ngay_bao_cao
    """, conn)

    df_daily['ngay_bao_cao'] = pd.to_datetime(df_daily['ngay_bao_cao']).dt.strftime('%d/%m')
    df_daily.columns = ['Ngày', 'Số lượng TB suy hao']

    conn.close()

    # Tạo file Excel
    if output_file is None:
        output_file = f"downloads/baocao_hanoi/Bao_cao_thang_{month:02d}_{year}.xlsx"

    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    print(f"\n✓ Đang ghi file Excel: {output_file}")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet tổng hợp
        df_summary.to_excel(writer, sheet_name='Tong_hop', index=False)

        # Sheet xu hướng theo ngày
        df_daily.to_excel(writer, sheet_name='Xu_huong_theo_ngay', index=False)

        # Sheet chi tiết TĂNG MỚI
        if len(df_tang) > 0:
            df_tang[['account_cts', 'ten_tb_one', 'doi_one', 'nvkt_db_normalized', 'sa']].to_excel(
                writer, sheet_name='Tang_moi', index=False
            )

        # Sheet chi tiết GIẢM/HẾT
        if len(df_giam) > 0:
            df_giam[['account_cts', 'ten_tb_one', 'doi_one', 'nvkt_db_normalized', 'sa']].to_excel(
                writer, sheet_name='Giam_het', index=False
            )

        # Sheet VẪN CÒN
        if len(df_van) > 0:
            df_van[['account_cts', 'ten_tb_one', 'doi_one', 'nvkt_db_normalized', 'sa']].to_excel(
                writer, sheet_name='Van_con', index=False
            )

    print(f"✅ Đã tạo báo cáo tháng: {output_file}")
    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH BÁO CÁO THÁNG {month}/{year}")
    print(f"{'='*80}\n")

    return output_file


def generate_trend_report(start_date, end_date, output_file=None):
    """
    Tạo báo cáo xu hướng tùy chỉnh theo khoảng thời gian

    Args:
        start_date: Ngày bắt đầu (format: 'YYYY-MM-DD' hoặc datetime)
        end_date: Ngày kết thúc (format: 'YYYY-MM-DD' hoặc datetime)
        output_file: Đường dẫn file Excel output (tùy chọn)

    Returns:
        Đường dẫn file Excel đã tạo
    """
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
    if isinstance(end_date, str):
        end_date = datetime.strptime(end_date, '%Y-%m-%d')

    print(f"\n{'='*80}")
    print(f"TẠO BÁO CÁO XU HƯỚNG: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}")
    print(f"{'='*80}\n")

    db_path = "suy_hao_history.db"
    if not os.path.exists(db_path):
        print(f"❌ Không tìm thấy database: {db_path}")
        return None

    conn = sqlite3.connect(db_path)

    # Xu hướng theo ngày
    df_daily = pd.read_sql_query(f"""
        SELECT
            ngay_bao_cao,
            doi_one,
            COUNT(DISTINCT account_cts) as so_luong
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao BETWEEN '{start_date.strftime('%Y-%m-%d')}'
          AND '{end_date.strftime('%Y-%m-%d')}'
        GROUP BY ngay_bao_cao, doi_one
        ORDER BY ngay_bao_cao, doi_one
    """, conn)

    df_daily['ngay_bao_cao'] = pd.to_datetime(df_daily['ngay_bao_cao']).dt.strftime('%d/%m/%Y')
    df_daily.columns = ['Ngày', 'Đơn vị', 'Số lượng']

    # Pivot để hiển thị theo đơn vị
    df_pivot = df_daily.pivot(index='Ngày', columns='Đơn vị', values='Số lượng').fillna(0)
    df_pivot = df_pivot.reset_index()

    # Biến động theo NVKT_DB
    df_nvkt_trend = pd.read_sql_query(f"""
        SELECT
            ngay_bao_cao,
            nvkt_db_normalized,
            COUNT(DISTINCT account_cts) as so_luong
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao BETWEEN '{start_date.strftime('%Y-%m-%d')}'
          AND '{end_date.strftime('%Y-%m-%d')}'
        GROUP BY ngay_bao_cao, nvkt_db_normalized
        ORDER BY ngay_bao_cao, nvkt_db_normalized
    """, conn)

    conn.close()

    # Pivot để hiển thị theo NVKT_DB (giống như df_pivot)
    df_nvkt_trend['ngay_bao_cao'] = pd.to_datetime(df_nvkt_trend['ngay_bao_cao']).dt.strftime('%d/%m/%Y')
    df_nvkt_pivot = df_nvkt_trend.pivot(index='ngay_bao_cao', columns='nvkt_db_normalized', values='so_luong').fillna(0)
    df_nvkt_pivot = df_nvkt_pivot.reset_index()

    # Tạo file Excel
    if output_file is None:
        output_file = f"downloads/baocao_hanoi/Bao_cao_xu_huong_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"

    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    print(f"\n✓ Đang ghi file Excel: {output_file}")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet xu hướng theo đơn vị
        df_pivot.to_excel(writer, sheet_name='Xu_huong_theo_don_vi', index=False)

        # Sheet xu hướng theo NVKT_DB (pivot table)
        df_nvkt_pivot.to_excel(writer, sheet_name='Xu_huong_theo_NVKT', index=False)

    print(f"✅ Đã tạo báo cáo xu hướng: {output_file}")
    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH BÁO CÁO XU HƯỚNG")
    print(f"{'='*80}\n")

    return output_file


if __name__ == "__main__":
    # Test các hàm
    print("Test module suy_hao_reports.py")

    # Báo cáo so sánh SHC ngày (T-1)
    generate_daily_comparison_report()

    # Lấy ngày dữ liệu mới nhất từ database
    db_path = "suy_hao_history.db"
    if os.path.exists(db_path):
        conn = sqlite3.connect(db_path)
        latest_date_query = """
            SELECT MAX(ngay_bao_cao) as latest_date
            FROM suy_hao_snapshots
        """
        df_latest = pd.read_sql_query(latest_date_query, conn)
        conn.close()

        if not df_latest.empty and df_latest['latest_date'][0]:
            latest_date = datetime.strptime(df_latest['latest_date'][0], '%Y-%m-%d')
            month_start = datetime(latest_date.year, latest_date.month, 1)
        else:
            # Nếu database trống, sử dụng tháng hiện tại
            today = datetime.now()
            month_start = datetime(today.year, today.month, 1)
            latest_date = today
    else:
        # Nếu không có database, sử dụng ngày hôm nay
        today = datetime.now()
        month_start = datetime(today.year, today.month, 1)
        latest_date = today

    # Test báo cáo tuần (tuần hiện tại)
    current_week = latest_date.isocalendar()[1]
    current_year = latest_date.year
    generate_weekly_report(current_year, current_week)

    # Test báo cáo tháng (tháng hiện tại)
    generate_monthly_report(latest_date.year, latest_date.month)

    # Test báo cáo xu hướng (từ đầu tháng đến ngày dữ liệu mới nhất)
    generate_trend_report(month_start.strftime('%Y-%m-%d'), latest_date.strftime('%Y-%m-%d'))
