# -*- coding: utf-8 -*-
"""
Script khởi tạo database theo dõi lịch sử suy hao
Chạy một lần để tạo cấu trúc database
"""

import sqlite3
import os

def init_suy_hao_database():
    """
    Tạo database suy_hao_history.db với đầy đủ schema
    """
    db_path = os.path.join(os.path.dirname(__file__), "suy_hao_history.db")

    print(f"\n{'='*80}")
    print(f"KHỞI TẠO DATABASE: {db_path}")
    print(f"{'='*80}\n")

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Bảng 1: Snapshot hàng ngày - lưu toàn bộ dữ liệu từng ngày
    print("✓ Tạo bảng suy_hao_snapshots...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            account_cts TEXT NOT NULL,
            ten_tb_one TEXT,
            dt_onediachi_one TEXT,
            doi_one TEXT,
            nvkt_db TEXT,
            nvkt_db_normalized TEXT,
            sa TEXT,
            olt_cts TEXT,
            port_cts TEXT,
            thietbi TEXT,
            ketcuoi TEXT,
            trangthai_tb TEXT,
            olt_rx REAL,
            onu_rx REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, account_cts)
        )
    """)

    # Bảng 2: Tracking trạng thái từng thuê bao
    print("✓ Tạo bảng suy_hao_tracking...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_tracking (
            account_cts TEXT PRIMARY KEY,
            ngay_xuat_hien_dau_tien DATE NOT NULL,
            ngay_thay_cuoi_cung DATE NOT NULL,
            so_ngay_lien_tuc INTEGER DEFAULT 1,
            doi_one TEXT,
            nvkt_db TEXT,
            sa TEXT,
            trang_thai TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Bảng 3: Biến động hàng ngày
    print("✓ Tạo bảng suy_hao_daily_changes...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_daily_changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            account_cts TEXT NOT NULL,
            loai_bien_dong TEXT NOT NULL,
            doi_one TEXT,
            nvkt_db TEXT,
            nvkt_db_normalized TEXT,
            sa TEXT,
            so_ngay_lien_tuc INTEGER,
            ten_tb_one TEXT,
            dt_onediachi_one TEXT,
            olt_cts TEXT,
            port_cts TEXT,
            thietbi TEXT,
            ketcuoi TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, account_cts, loai_bien_dong)
        )
    """)

    # Bảng 4: Tổng hợp biến động theo ngày (cho báo cáo nhanh)
    print("✓ Tạo bảng suy_hao_daily_summary...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_daily_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            doi_one TEXT,
            nvkt_db_normalized TEXT,
            tong_so_hien_tai INTEGER DEFAULT 0,
            so_tang_moi INTEGER DEFAULT 0,
            so_giam_het INTEGER DEFAULT 0,
            so_van_con INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, doi_one, nvkt_db_normalized)
        )
    """)

    # Bảng 5: Biến động theo tuần
    print("✓ Tạo bảng suy_hao_weekly_changes...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_weekly_changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nam INTEGER NOT NULL,
            tuan INTEGER NOT NULL,
            ngay_bat_dau DATE NOT NULL,
            ngay_ket_thuc DATE NOT NULL,
            account_cts TEXT NOT NULL,
            loai_bien_dong TEXT NOT NULL,
            doi_one TEXT,
            nvkt_db_normalized TEXT,
            sa TEXT,
            ten_tb_one TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(nam, tuan, account_cts, loai_bien_dong)
        )
    """)

    # Bảng 6: Tổng hợp theo tuần
    print("✓ Tạo bảng suy_hao_weekly_summary...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_weekly_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nam INTEGER NOT NULL,
            tuan INTEGER NOT NULL,
            ngay_bat_dau DATE NOT NULL,
            ngay_ket_thuc DATE NOT NULL,
            doi_one TEXT,
            nvkt_db_normalized TEXT,
            so_tang_moi INTEGER DEFAULT 0,
            so_giam_het INTEGER DEFAULT 0,
            so_van_con INTEGER DEFAULT 0,
            tong_so_cuoi_tuan INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(nam, tuan, doi_one, nvkt_db_normalized)
        )
    """)

    # Bảng 7: Biến động theo tháng
    print("✓ Tạo bảng suy_hao_monthly_changes...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_monthly_changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nam INTEGER NOT NULL,
            thang INTEGER NOT NULL,
            account_cts TEXT NOT NULL,
            loai_bien_dong TEXT NOT NULL,
            doi_one TEXT,
            nvkt_db_normalized TEXT,
            sa TEXT,
            ten_tb_one TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(nam, thang, account_cts, loai_bien_dong)
        )
    """)

    # Bảng 8: Tổng hợp theo tháng
    print("✓ Tạo bảng suy_hao_monthly_summary...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS suy_hao_monthly_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nam INTEGER NOT NULL,
            thang INTEGER NOT NULL,
            doi_one TEXT,
            nvkt_db_normalized TEXT,
            so_tang_moi INTEGER DEFAULT 0,
            so_giam_het INTEGER DEFAULT 0,
            so_van_con INTEGER DEFAULT 0,
            tong_so_cuoi_thang INTEGER DEFAULT 0,
            so_tb_trung_binh_ngay REAL,
            so_tb_cao_nhat_ngay INTEGER,
            so_tb_thap_nhat_ngay INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(nam, thang, doi_one, nvkt_db_normalized)
        )
    """)

    # Tạo các index để tăng tốc query
    print("\n✓ Tạo indexes...")

    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_snapshots_date ON suy_hao_snapshots(ngay_bao_cao)",
        "CREATE INDEX IF NOT EXISTS idx_snapshots_account ON suy_hao_snapshots(account_cts)",
        "CREATE INDEX IF NOT EXISTS idx_snapshots_nvkt ON suy_hao_snapshots(nvkt_db_normalized)",
        "CREATE INDEX IF NOT EXISTS idx_tracking_dates ON suy_hao_tracking(ngay_xuat_hien_dau_tien, ngay_thay_cuoi_cung)",
        "CREATE INDEX IF NOT EXISTS idx_daily_changes_date ON suy_hao_daily_changes(ngay_bao_cao)",
        "CREATE INDEX IF NOT EXISTS idx_daily_changes_type ON suy_hao_daily_changes(loai_bien_dong)",
        "CREATE INDEX IF NOT EXISTS idx_daily_summary_date ON suy_hao_daily_summary(ngay_bao_cao)",
        "CREATE INDEX IF NOT EXISTS idx_weekly_changes_week ON suy_hao_weekly_changes(nam, tuan)",
        "CREATE INDEX IF NOT EXISTS idx_weekly_summary_week ON suy_hao_weekly_summary(nam, tuan)",
        "CREATE INDEX IF NOT EXISTS idx_monthly_changes_month ON suy_hao_monthly_changes(nam, thang)",
        "CREATE INDEX IF NOT EXISTS idx_monthly_summary_month ON suy_hao_monthly_summary(nam, thang)",
    ]

    for idx_sql in indexes:
        cursor.execute(idx_sql)

    conn.commit()

    # Hiển thị thông tin database
    print("\n✓ Kiểm tra các bảng đã tạo:")
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = cursor.fetchall()
    for i, (table_name,) in enumerate(tables, 1):
        cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
        count = cursor.fetchone()[0]
        print(f"  {i}. {table_name}: {count} bản ghi")

    conn.close()

    print(f"\n{'='*80}")
    print(f"✅ HOÀN THÀNH KHỞI TẠO DATABASE")
    print(f"{'='*80}\n")

    return db_path


if __name__ == "__main__":
    db_path = init_suy_hao_database()
    print(f"Database đã được tạo tại: {db_path}")
