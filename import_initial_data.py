# -*- coding: utf-8 -*-
"""
Script import dữ liệu khởi đầu từ 2 file I1.5 report
- File ngày 02/11: I1.5 report-3.11.xlsx
- File ngày 03/11: I1.5 report.xlsx
"""

import pandas as pd
import sqlite3
import os
from datetime import datetime
import re

def normalize_nvkt(x):
    """Chuẩn hóa tên NVKT_DB - giữ phần sau dấu '-'"""
    if not isinstance(x, str):
        return x
    if '-' in x:
        x = x.split('-')[1].strip()
    x = re.sub(r'\([^)]*\)', '', x).strip()
    return x


def import_i15_report_to_db(excel_file, report_date, db_path="suy_hao_history.db"):
    """
    Import một file I1.5 report vào database

    Args:
        excel_file: Đường dẫn file Excel
        report_date: Ngày báo cáo (format: 'YYYY-MM-DD')
        db_path: Đường dẫn database
    """
    print(f"\n{'='*80}")
    print(f"IMPORT DỮ LIỆU: {os.path.basename(excel_file)}")
    print(f"Ngày báo cáo: {report_date}")
    print(f"{'='*80}\n")

    # Đọc file Excel
    print(f"✓ Đang đọc file {excel_file}...")
    df = pd.read_excel(excel_file)
    print(f"  Tổng số dòng: {len(df)}")

    # Chuẩn hóa NVKT_DB
    if 'NVKT_DB' in df.columns:
        df['NVKT_DB_NORMALIZED'] = df['NVKT_DB'].apply(normalize_nvkt)
    else:
        df['NVKT_DB_NORMALIZED'] = None

    # Kết nối database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Import vào bảng snapshots
    print(f"\n✓ Import vào bảng suy_hao_snapshots...")

    inserted_count = 0
    duplicate_count = 0

    for idx, row in df.iterrows():
        try:
            cursor.execute("""
                INSERT INTO suy_hao_snapshots (
                    ngay_bao_cao, account_cts, ten_tb_one, dt_onediachi_one,
                    doi_one, nvkt_db, nvkt_db_normalized, sa,
                    olt_cts, port_cts, thietbi, ketcuoi, trangthai_tb
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_date,
                row.get('ACCOUNT_CTS'),
                row.get('TEN_TB_ONE'),
                row.get('DT_ONEDIACHI_ONE'),
                row.get('DOI_ONE'),
                row.get('NVKT_DB'),
                row.get('NVKT_DB_NORMALIZED'),
                row.get('SA'),
                row.get('OLT_CTS'),
                row.get('PORT_CTS'),
                row.get('THIETBI'),
                row.get('KETCUOI'),
                row.get('TRANGTHAI_TB')
            ))
            inserted_count += 1
        except sqlite3.IntegrityError:
            duplicate_count += 1

    conn.commit()
    print(f"  ✅ Đã insert: {inserted_count} bản ghi")
    if duplicate_count > 0:
        print(f"  ⚠️  Bỏ qua (trùng): {duplicate_count} bản ghi")

    # Update bảng tracking
    print(f"\n✓ Cập nhật bảng suy_hao_tracking...")

    tracking_updates = 0
    tracking_inserts = 0

    for idx, row in df.iterrows():
        account = row.get('ACCOUNT_CTS')

        # Kiểm tra thuê bao đã tồn tại chưa
        cursor.execute("""
            SELECT account_cts, ngay_xuat_hien_dau_tien, so_ngay_lien_tuc
            FROM suy_hao_tracking
            WHERE account_cts = ?
        """, (account,))

        existing = cursor.fetchone()

        if existing:
            # Cập nhật: tăng số ngày liên tục
            _, first_date, days_count = existing
            new_days_count = days_count + 1

            cursor.execute("""
                UPDATE suy_hao_tracking
                SET ngay_thay_cuoi_cung = ?,
                    so_ngay_lien_tuc = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE account_cts = ?
            """, (report_date, new_days_count, account))
            tracking_updates += 1
        else:
            # Insert mới
            cursor.execute("""
                INSERT INTO suy_hao_tracking (
                    account_cts, ngay_xuat_hien_dau_tien, ngay_thay_cuoi_cung,
                    so_ngay_lien_tuc, doi_one, nvkt_db, sa, trang_thai
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                account,
                report_date,
                report_date,
                1,
                row.get('DOI_ONE'),
                row.get('NVKT_DB_NORMALIZED'),
                row.get('SA'),
                'DANG_SUY_HAO'
            ))
            tracking_inserts += 1

    conn.commit()
    print(f"  ✅ Đã insert mới: {tracking_inserts} thuê bao")
    print(f"  ✅ Đã cập nhật: {tracking_updates} thuê bao")

    conn.close()

    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH IMPORT DỮ LIỆU NGÀY {report_date}")
    print(f"{'='*80}\n")


def calculate_daily_changes(date1, date2, db_path="suy_hao_history.db"):
    """
    Tính toán biến động giữa 2 ngày

    Args:
        date1: Ngày cũ (YYYY-MM-DD)
        date2: Ngày mới (YYYY-MM-DD)
        db_path: Đường dẫn database
    """
    print(f"\n{'='*80}")
    print(f"TÍNH TOÁN BIẾN ĐỘNG: {date1} → {date2}")
    print(f"{'='*80}\n")

    conn = sqlite3.connect(db_path)

    # Đọc dữ liệu 2 ngày
    df1 = pd.read_sql_query(f"""
        SELECT account_cts, doi_one, nvkt_db_normalized, sa,
               ten_tb_one, dt_onediachi_one, olt_cts, port_cts, thietbi, ketcuoi
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao = '{date1}'
    """, conn)

    df2 = pd.read_sql_query(f"""
        SELECT account_cts, doi_one, nvkt_db_normalized, sa,
               ten_tb_one, dt_onediachi_one, olt_cts, port_cts, thietbi, ketcuoi
        FROM suy_hao_snapshots
        WHERE ngay_bao_cao = '{date2}'
    """, conn)

    print(f"✓ Ngày {date1}: {len(df1)} thuê bao")
    print(f"✓ Ngày {date2}: {len(df2)} thuê bao")

    accounts1 = set(df1['account_cts'].tolist())
    accounts2 = set(df2['account_cts'].tolist())

    # Phân loại
    tang_moi = accounts2 - accounts1
    giam_het = accounts1 - accounts2
    van_con = accounts1 & accounts2

    print(f"\n✓ Phân tích biến động:")
    print(f"  🆕 TĂNG MỚI: {len(tang_moi)} thuê bao")
    print(f"  ⬇️  GIẢM/HẾT: {len(giam_het)} thuê bao")
    print(f"  ↔️  VẪN CÒN: {len(van_con)} thuê bao")

    cursor = conn.cursor()

    # Xóa dữ liệu cũ nếu có
    cursor.execute("""
        DELETE FROM suy_hao_daily_changes
        WHERE ngay_bao_cao = ?
    """, (date2,))

    # Insert TĂNG MỚI
    print(f"\n✓ Lưu dữ liệu TĂNG MỚI...")
    for account in tang_moi:
        row = df2[df2['account_cts'] == account].iloc[0]
        cursor.execute("""
            INSERT OR REPLACE INTO suy_hao_daily_changes (
                ngay_bao_cao, account_cts, loai_bien_dong,
                doi_one, nvkt_db, nvkt_db_normalized, sa,
                so_ngay_lien_tuc, ten_tb_one, dt_onediachi_one,
                olt_cts, port_cts, thietbi, ketcuoi
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            date2, account, 'TANG_MOI',
            row['doi_one'], row.get('nvkt_db'), row['nvkt_db_normalized'], row['sa'],
            1, row['ten_tb_one'], row['dt_onediachi_one'],
            row['olt_cts'], row['port_cts'], row['thietbi'], row['ketcuoi']
        ))

    # Insert GIẢM/HẾT
    print(f"✓ Lưu dữ liệu GIẢM/HẾT...")
    for account in giam_het:
        row = df1[df1['account_cts'] == account].iloc[0]

        # Lấy số ngày liên tục từ tracking
        cursor.execute("""
            SELECT so_ngay_lien_tuc FROM suy_hao_tracking
            WHERE account_cts = ?
        """, (account,))
        result = cursor.fetchone()
        so_ngay = result[0] if result else 1

        cursor.execute("""
            INSERT OR REPLACE INTO suy_hao_daily_changes (
                ngay_bao_cao, account_cts, loai_bien_dong,
                doi_one, nvkt_db, nvkt_db_normalized, sa,
                so_ngay_lien_tuc, ten_tb_one, dt_onediachi_one,
                olt_cts, port_cts, thietbi, ketcuoi
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            date2, account, 'GIAM_HET',
            row['doi_one'], row.get('nvkt_db'), row['nvkt_db_normalized'], row['sa'],
            so_ngay, row['ten_tb_one'], row['dt_onediachi_one'],
            row['olt_cts'], row['port_cts'], row['thietbi'], row['ketcuoi']
        ))

        # Cập nhật trạng thái tracking
        cursor.execute("""
            UPDATE suy_hao_tracking
            SET trang_thai = 'DA_HET_SUY_HAO'
            WHERE account_cts = ?
        """, (account,))

    # Insert VẪN CÒN
    print(f"✓ Lưu dữ liệu VẪN CÒN...")
    for account in van_con:
        row = df2[df2['account_cts'] == account].iloc[0]

        # Lấy số ngày liên tục từ tracking
        cursor.execute("""
            SELECT so_ngay_lien_tuc FROM suy_hao_tracking
            WHERE account_cts = ?
        """, (account,))
        result = cursor.fetchone()
        so_ngay = result[0] if result else 1

        cursor.execute("""
            INSERT OR REPLACE INTO suy_hao_daily_changes (
                ngay_bao_cao, account_cts, loai_bien_dong,
                doi_one, nvkt_db, nvkt_db_normalized, sa,
                so_ngay_lien_tuc, ten_tb_one, dt_onediachi_one,
                olt_cts, port_cts, thietbi, ketcuoi
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            date2, account, 'VAN_CON',
            row['doi_one'], row.get('nvkt_db'), row['nvkt_db_normalized'], row['sa'],
            so_ngay, row['ten_tb_one'], row['dt_onediachi_one'],
            row['olt_cts'], row['port_cts'], row['thietbi'], row['ketcuoi']
        ))

    conn.commit()

    # Tạo summary
    print(f"\n✓ Tạo daily summary...")
    cursor.execute("""
        DELETE FROM suy_hao_daily_summary
        WHERE ngay_bao_cao = ?
    """, (date2,))

    cursor.execute(f"""
        INSERT INTO suy_hao_daily_summary (
            ngay_bao_cao, doi_one, nvkt_db_normalized,
            tong_so_hien_tai, so_tang_moi, so_giam_het, so_van_con
        )
        SELECT
            '{date2}' as ngay_bao_cao,
            doi_one,
            nvkt_db_normalized,
            SUM(CASE WHEN loai_bien_dong IN ('TANG_MOI', 'VAN_CON') THEN 1 ELSE 0 END) as tong_so_hien_tai,
            SUM(CASE WHEN loai_bien_dong = 'TANG_MOI' THEN 1 ELSE 0 END) as so_tang_moi,
            SUM(CASE WHEN loai_bien_dong = 'GIAM_HET' THEN 1 ELSE 0 END) as so_giam_het,
            SUM(CASE WHEN loai_bien_dong = 'VAN_CON' THEN 1 ELSE 0 END) as so_van_con
        FROM suy_hao_daily_changes
        WHERE ngay_bao_cao = '{date2}'
        GROUP BY doi_one, nvkt_db_normalized
    """)

    conn.commit()
    conn.close()

    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH TÍNH TOÁN BIẾN ĐỘNG")
    print(f"{'='*80}\n")


def main():
    """Main function để import dữ liệu khởi đầu"""

    base_dir = os.path.dirname(__file__)
    download_dir = os.path.join(base_dir, "downloads", "baocao_hanoi")

    # File paths
    file_3nov = os.path.join(download_dir, "I1.5 report-3.11.xlsx")
    file_4nov = os.path.join(download_dir, "I1.5 report.xlsx")

    db_path = os.path.join(base_dir, "suy_hao_history.db")

    # Import ngày 02/11 (file tên 3.11 nhưng dữ liệu là 02/11)
    if os.path.exists(file_3nov):
        import_i15_report_to_db(file_3nov, "2025-11-02", db_path)
    else:
        print(f"❌ Không tìm thấy file: {file_3nov}")
        return

    # Import ngày 03/11
    if os.path.exists(file_4nov):
        import_i15_report_to_db(file_4nov, "2025-11-03", db_path)
    else:
        print(f"❌ Không tìm thấy file: {file_4nov}")
        return

    # Tính toán biến động
    calculate_daily_changes("2025-11-02", "2025-11-03", db_path)

    print("\n" + "="*80)
    print("🎉 HOÀN THÀNH IMPORT TOÀN BỘ DỮ LIỆU KHỞI ĐẦU")
    print("="*80)


if __name__ == "__main__":
    main()
