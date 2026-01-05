# -*- coding: utf-8 -*-
"""
init_reports_history_db.py
Khởi tạo database reports_history.db để tracking lịch sử tất cả báo cáo

Database này bổ sung cho suy_hao_history.db (chỉ tracking I1.5)
"""

import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), 'reports_history.db')


def get_connection():
    """Tạo kết nối đến database"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_database():
    """Khởi tạo database với tất cả tables và indexes"""
    conn = get_connection()
    cursor = conn.cursor()

    print("=" * 60)
    print("KHỞI TẠO DATABASE: reports_history.db")
    print("=" * 60)

    # ============================================
    # SECTION 1: METADATA AND CONFIGURATION
    # ============================================
    print("\n[1/7] Tạo bảng metadata...")

    # Report types registry
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS report_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_code TEXT UNIQUE NOT NULL,
            report_name TEXT NOT NULL,
            report_category TEXT NOT NULL,
            has_nvkt_detail INTEGER DEFAULT 0,
            has_don_vi_detail INTEGER DEFAULT 1,
            snapshot_frequency TEXT DEFAULT 'DAILY',
            retention_days INTEGER DEFAULT 365,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Insert report type definitions
    report_types = [
        ('C1.1', 'Tỷ lệ sửa chữa phiếu chất lượng', 'QUALITY', 1, 1),
        ('C1.2', 'Tỷ lệ thuê bao báo hỏng lặp lại', 'QUALITY', 1, 1),
        ('C1.3', 'Tỷ lệ báo hỏng đúng hạn', 'QUALITY', 0, 1),
        ('C1.4', 'Báo cáo CSKH tổng hợp', 'QUALITY', 1, 1),
        ('C1.5', 'Tỷ lệ thiết lập dịch vụ BRCĐ', 'QUALITY', 1, 1),
        ('KR6', 'KR6 - Gia hạn TTTC tháng T', 'KR', 1, 1),
        ('KR7', 'KR7 - Gia hạn TTTC tháng T+1', 'KR', 1, 1),
        ('THUC_TANG_PTTB', 'Thực tăng PTTB (Fiber)', 'GROWTH', 1, 1),
        ('THUC_TANG_MYTV', 'Thực tăng MyTV', 'GROWTH', 1, 1),
        ('HOAN_CONG_PTTB', 'Hoàn công PTTB', 'GROWTH', 1, 1),
        ('HOAN_CONG_MYTV', 'Hoàn công MyTV', 'GROWTH', 1, 1),
        ('NGUNG_PSC_PTTB', 'Ngưng PSC PTTB', 'GROWTH', 1, 1),
        ('NGUNG_PSC_MYTV', 'Ngưng PSC MyTV', 'GROWTH', 1, 1),
        ('VAT_TU_THU_HOI', 'Vật tư thu hồi', 'EQUIPMENT', 1, 1),
    ]

    for rt in report_types:
        cursor.execute('''
            INSERT OR IGNORE INTO report_types
            (report_code, report_name, report_category, has_nvkt_detail, has_don_vi_detail)
            VALUES (?, ?, ?, ?, ?)
        ''', rt)

    # Import log
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS import_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_code TEXT NOT NULL,
            ngay_bao_cao DATE NOT NULL,
            file_name TEXT,
            file_path TEXT,
            record_count INTEGER DEFAULT 0,
            status TEXT DEFAULT 'SUCCESS',
            error_message TEXT,
            import_duration_ms INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_import_log_date ON import_log(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_import_log_report ON import_log(report_code)')

    print("   - report_types: OK")
    print("   - import_log: OK")

    # ============================================
    # SECTION 2: C1.x QUALITY REPORTS
    # ============================================
    print("\n[2/7] Tạo bảng C1.x...")

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            thang INTEGER NOT NULL,
            nam INTEGER NOT NULL,
            loai_bao_cao TEXT NOT NULL,
            don_vi TEXT NOT NULL,

            sm1 INTEGER,
            sm2 INTEGER,
            sm3 INTEGER,
            sm4 INTEGER,

            ty_le_chinh REAL,
            ty_le_phu REAL,
            diem_bsc REAL,

            raw_data TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_bao_cao, loai_bao_cao, don_vi)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_snapshots_date ON c1_snapshots(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_snapshots_type ON c1_snapshots(loai_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_snapshots_month ON c1_snapshots(nam, thang)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_snapshots_don_vi ON c1_snapshots(don_vi)')

    # C1 chi tiết theo NVKT
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_snapshots_nvkt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            thang INTEGER NOT NULL,
            nam INTEGER NOT NULL,
            loai_bao_cao TEXT NOT NULL,
            don_vi TEXT,
            nvkt TEXT NOT NULL,

            tong_phieu INTEGER DEFAULT 0,
            phieu_dat INTEGER DEFAULT 0,
            phieu_khong_dat INTEGER DEFAULT 0,
            ty_le_dat REAL,

            raw_data TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_bao_cao, loai_bao_cao, nvkt)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_nvkt_date ON c1_snapshots_nvkt(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_nvkt_type ON c1_snapshots_nvkt(loai_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_nvkt_nvkt ON c1_snapshots_nvkt(nvkt)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_c1_nvkt_don_vi ON c1_snapshots_nvkt(don_vi)')

    print("   - c1_snapshots: OK")
    print("   - c1_snapshots_nvkt: OK")

    # ============================================
    # SECTION 3: KR6/KR7 REPORTS
    # ============================================
    print("\n[3/7] Tạo bảng KR6/KR7...")

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS kr_snapshots_tonghop (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            thang INTEGER NOT NULL,
            nam INTEGER NOT NULL,
            loai_bao_cao TEXT NOT NULL,
            don_vi TEXT NOT NULL,

            so_giao INTEGER DEFAULT 0,
            so_hoan_thanh INTEGER DEFAULT 0,
            ty_le_hoan_thanh REAL,
            ke_hoach_giao INTEGER,
            diem_okr REAL,
            ton_chua_nghiem_thu INTEGER,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_bao_cao, loai_bao_cao, don_vi)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kr_tonghop_date ON kr_snapshots_tonghop(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kr_tonghop_type ON kr_snapshots_tonghop(loai_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kr_tonghop_month ON kr_snapshots_tonghop(nam, thang)')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS kr_snapshots_nvkt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            thang INTEGER NOT NULL,
            nam INTEGER NOT NULL,
            loai_bao_cao TEXT NOT NULL,

            nvkt TEXT NOT NULL,
            don_vi TEXT,

            so_giao INTEGER DEFAULT 0,
            so_hoan_thanh INTEGER DEFAULT 0,
            ty_le_hoan_thanh REAL,
            ke_hoach_giao INTEGER,
            diem_okr REAL,
            ton_chua_nghiem_thu INTEGER,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_bao_cao, loai_bao_cao, nvkt)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kr_nvkt_date ON kr_snapshots_nvkt(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kr_nvkt_type ON kr_snapshots_nvkt(loai_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kr_nvkt_nvkt ON kr_snapshots_nvkt(nvkt)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kr_nvkt_don_vi ON kr_snapshots_nvkt(don_vi)')

    print("   - kr_snapshots_tonghop: OK")
    print("   - kr_snapshots_nvkt: OK")

    # ============================================
    # SECTION 4: GROWTH REPORTS (Thực tăng, Hoàn công, Ngưng PSC)
    # ============================================
    print("\n[4/7] Tạo bảng Thực tăng/Hoàn công/Ngưng PSC...")

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS growth_snapshots_donvi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            loai_dich_vu TEXT NOT NULL,
            don_vi TEXT NOT NULL,

            so_hoan_cong INTEGER DEFAULT 0,
            so_ngung_psc INTEGER DEFAULT 0,
            thuc_tang INTEGER DEFAULT 0,
            ty_le_ngung_psc REAL,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_bao_cao, loai_dich_vu, don_vi)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_growth_donvi_date ON growth_snapshots_donvi(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_growth_donvi_type ON growth_snapshots_donvi(loai_dich_vu)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_growth_donvi_don_vi ON growth_snapshots_donvi(don_vi)')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS growth_snapshots_nvkt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            loai_dich_vu TEXT NOT NULL,
            nvkt TEXT NOT NULL,
            don_vi TEXT,

            so_hoan_cong INTEGER DEFAULT 0,
            so_ngung_psc INTEGER DEFAULT 0,
            thuc_tang INTEGER DEFAULT 0,
            ty_le_ngung_psc REAL,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_bao_cao, loai_dich_vu, nvkt)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_growth_nvkt_date ON growth_snapshots_nvkt(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_growth_nvkt_type ON growth_snapshots_nvkt(loai_dich_vu)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_growth_nvkt_nvkt ON growth_snapshots_nvkt(nvkt)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_growth_nvkt_don_vi ON growth_snapshots_nvkt(don_vi)')

    # Chi tiết hoàn công
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS hoan_cong_detail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            loai_dich_vu TEXT NOT NULL,

            stt INTEGER,
            ma_tb TEXT,
            ngay_nghiem_thu TEXT,
            ngay_yeu_cau TEXT,
            doi TEXT,
            nhom_dia_ban TEXT,
            ten_ttvt TEXT,
            trang_thai_phieu TEXT,
            hdtb_id TEXT,
            nhan_vien_kt TEXT,
            ma_gd TEXT,
            nvkt TEXT,
            don_vi TEXT,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_hoan_cong_detail_date ON hoan_cong_detail(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_hoan_cong_detail_nvkt ON hoan_cong_detail(nvkt)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_hoan_cong_detail_ma_tb ON hoan_cong_detail(ma_tb)')

    # Chi tiết ngưng PSC
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ngung_psc_detail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            loai_dich_vu TEXT NOT NULL,

            stt INTEGER,
            ma_tb TEXT,
            so_may TEXT,
            ten_tb TEXT,
            loai_dich_vu_tb TEXT,
            dia_chi_ld TEXT,
            ngay_tam_dung TEXT,
            ngay_khoi_phuc TEXT,
            ngay_huy TEXT,
            nhom_dia_ban TEXT,
            ten_to TEXT,
            ten_ttvt TEXT,
            ten_kh TEXT,
            dien_thoai_lh TEXT,
            trang_thai_tb TEXT,
            ly_do_huy_tam_dung TEXT,
            nvkt TEXT,
            don_vi TEXT,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_ngung_psc_detail_date ON ngung_psc_detail(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_ngung_psc_detail_nvkt ON ngung_psc_detail(nvkt)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_ngung_psc_detail_ma_tb ON ngung_psc_detail(ma_tb)')

    print("   - growth_snapshots_donvi: OK")
    print("   - growth_snapshots_nvkt: OK")
    print("   - hoan_cong_detail: OK")
    print("   - ngung_psc_detail: OK")

    # ============================================
    # SECTION 5: VAT TU THU HOI
    # ============================================
    print("\n[5/7] Tạo bảng Vật tư thu hồi...")

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS vat_tu_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            don_vi TEXT,
            nvkt TEXT,
            diem_chia TEXT,

            so_luong_chua_thu_hoi INTEGER DEFAULT 0,
            so_luong_da_thu_hoi INTEGER DEFAULT 0,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_bao_cao, don_vi, nvkt, diem_chia)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_vat_tu_snapshots_date ON vat_tu_snapshots(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_vat_tu_snapshots_nvkt ON vat_tu_snapshots(nvkt)')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS vat_tu_detail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,

            nvkt_diaban_giao TEXT,
            ma_tb TEXT,
            ten_tb TEXT,
            ten_tbi TEXT,
            ngay_giao TEXT,
            ten_loaihd TEXT,
            ten_kieuld TEXT,
            so_dt TEXT,
            ngay_sd_tb TEXT,
            diem_chia TEXT,
            trang_thai_thu_hoi TEXT,
            loai_vt TEXT,
            loai_phieu TEXT,

            nvkt TEXT,
            don_vi TEXT,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_vat_tu_detail_date ON vat_tu_detail(ngay_bao_cao)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_vat_tu_detail_nvkt ON vat_tu_detail(nvkt)')

    print("   - vat_tu_snapshots: OK")
    print("   - vat_tu_detail: OK")

    # ============================================
    # SECTION 6: UNIFIED SUMMARY TABLES
    # ============================================
    print("\n[6/7] Tạo bảng Summary...")

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL UNIQUE,

            pttb_hoan_cong INTEGER DEFAULT 0,
            pttb_ngung_psc INTEGER DEFAULT 0,
            pttb_thuc_tang INTEGER DEFAULT 0,

            mytv_hoan_cong INTEGER DEFAULT 0,
            mytv_ngung_psc INTEGER DEFAULT 0,
            mytv_thuc_tang INTEGER DEFAULT 0,

            vat_tu_chua_thu_hoi INTEGER DEFAULT 0,
            suy_hao_cao_count INTEGER DEFAULT 0,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_daily_summary_date ON daily_summary(ngay_bao_cao)')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS weekly_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nam INTEGER NOT NULL,
            tuan INTEGER NOT NULL,
            ngay_bat_dau DATE NOT NULL,
            ngay_ket_thuc DATE NOT NULL,

            pttb_hoan_cong INTEGER DEFAULT 0,
            pttb_ngung_psc INTEGER DEFAULT 0,
            pttb_thuc_tang INTEGER DEFAULT 0,

            mytv_hoan_cong INTEGER DEFAULT 0,
            mytv_ngung_psc INTEGER DEFAULT 0,
            mytv_thuc_tang INTEGER DEFAULT 0,

            pttb_thuc_tang_wow_change INTEGER DEFAULT 0,
            pttb_thuc_tang_wow_pct REAL,
            mytv_thuc_tang_wow_change INTEGER DEFAULT 0,
            mytv_thuc_tang_wow_pct REAL,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(nam, tuan)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_weekly_summary_week ON weekly_summary(nam, tuan)')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS monthly_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nam INTEGER NOT NULL,
            thang INTEGER NOT NULL,

            pttb_hoan_cong INTEGER DEFAULT 0,
            pttb_ngung_psc INTEGER DEFAULT 0,
            pttb_thuc_tang INTEGER DEFAULT 0,

            mytv_hoan_cong INTEGER DEFAULT 0,
            mytv_ngung_psc INTEGER DEFAULT 0,
            mytv_thuc_tang INTEGER DEFAULT 0,

            kr6_ty_le_trung_binh REAL,
            kr7_ty_le_trung_binh REAL,

            pttb_thuc_tang_mom_change INTEGER DEFAULT 0,
            pttb_thuc_tang_mom_pct REAL,
            mytv_thuc_tang_mom_change INTEGER DEFAULT 0,
            mytv_thuc_tang_mom_pct REAL,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(nam, thang)
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_monthly_summary_month ON monthly_summary(nam, thang)')

    print("   - daily_summary: OK")
    print("   - weekly_summary: OK")
    print("   - monthly_summary: OK")

    # ============================================
    # SECTION 7: VIEWS FOR DASHBOARD
    # ============================================
    print("\n[7/7] Tạo Views cho Dashboard...")

    # View: Daily growth trend (last 30 days)
    cursor.execute('DROP VIEW IF EXISTS v_growth_trend_30d')
    cursor.execute('''
        CREATE VIEW v_growth_trend_30d AS
        SELECT
            ngay_bao_cao,
            SUM(CASE WHEN loai_dich_vu = 'PTTB' THEN thuc_tang ELSE 0 END) as pttb_thuc_tang,
            SUM(CASE WHEN loai_dich_vu = 'MYTV' THEN thuc_tang ELSE 0 END) as mytv_thuc_tang,
            SUM(CASE WHEN loai_dich_vu = 'PTTB' THEN so_hoan_cong ELSE 0 END) as pttb_hoan_cong,
            SUM(CASE WHEN loai_dich_vu = 'PTTB' THEN so_ngung_psc ELSE 0 END) as pttb_ngung_psc,
            SUM(CASE WHEN loai_dich_vu = 'MYTV' THEN so_hoan_cong ELSE 0 END) as mytv_hoan_cong,
            SUM(CASE WHEN loai_dich_vu = 'MYTV' THEN so_ngung_psc ELSE 0 END) as mytv_ngung_psc
        FROM growth_snapshots_donvi
        WHERE ngay_bao_cao >= date('now', '-30 days')
        GROUP BY ngay_bao_cao
        ORDER BY ngay_bao_cao DESC
    ''')

    # View: Don vi performance comparison
    cursor.execute('DROP VIEW IF EXISTS v_don_vi_performance')
    cursor.execute('''
        CREATE VIEW v_don_vi_performance AS
        SELECT
            don_vi,
            ngay_bao_cao,
            SUM(CASE WHEN loai_dich_vu = 'PTTB' THEN thuc_tang ELSE 0 END) as pttb_thuc_tang,
            SUM(CASE WHEN loai_dich_vu = 'PTTB' THEN so_hoan_cong ELSE 0 END) as pttb_hoan_cong,
            SUM(CASE WHEN loai_dich_vu = 'PTTB' THEN so_ngung_psc ELSE 0 END) as pttb_ngung_psc,
            SUM(CASE WHEN loai_dich_vu = 'MYTV' THEN thuc_tang ELSE 0 END) as mytv_thuc_tang,
            SUM(CASE WHEN loai_dich_vu = 'MYTV' THEN so_hoan_cong ELSE 0 END) as mytv_hoan_cong,
            SUM(CASE WHEN loai_dich_vu = 'MYTV' THEN so_ngung_psc ELSE 0 END) as mytv_ngung_psc
        FROM growth_snapshots_donvi
        GROUP BY don_vi, ngay_bao_cao
        ORDER BY ngay_bao_cao DESC, pttb_thuc_tang DESC
    ''')

    # View: NVKT ranking by thuc tang (current month)
    cursor.execute('DROP VIEW IF EXISTS v_nvkt_ranking_month')
    cursor.execute('''
        CREATE VIEW v_nvkt_ranking_month AS
        SELECT
            nvkt,
            don_vi,
            loai_dich_vu,
            SUM(thuc_tang) as tong_thuc_tang,
            SUM(so_hoan_cong) as tong_hoan_cong,
            SUM(so_ngung_psc) as tong_ngung_psc,
            ROUND(AVG(ty_le_ngung_psc), 2) as ty_le_ngung_psc_tb
        FROM growth_snapshots_nvkt
        WHERE strftime('%Y-%m', ngay_bao_cao) = strftime('%Y-%m', 'now')
        GROUP BY nvkt, don_vi, loai_dich_vu
        ORDER BY tong_thuc_tang DESC
    ''')

    # View: KR completion rates trend
    cursor.execute('DROP VIEW IF EXISTS v_kr_trend')
    cursor.execute('''
        CREATE VIEW v_kr_trend AS
        SELECT
            ngay_bao_cao,
            loai_bao_cao,
            AVG(ty_le_hoan_thanh) as ty_le_trung_binh,
            MIN(ty_le_hoan_thanh) as ty_le_thap_nhat,
            MAX(ty_le_hoan_thanh) as ty_le_cao_nhat,
            COUNT(DISTINCT don_vi) as so_don_vi
        FROM kr_snapshots_tonghop
        GROUP BY ngay_bao_cao, loai_bao_cao
        ORDER BY ngay_bao_cao DESC
    ''')

    # View: Day-over-day comparison
    cursor.execute('DROP VIEW IF EXISTS v_dod_comparison')
    cursor.execute('''
        CREATE VIEW v_dod_comparison AS
        WITH today_data AS (
            SELECT * FROM daily_summary
            WHERE ngay_bao_cao = (SELECT MAX(ngay_bao_cao) FROM daily_summary)
        ),
        yesterday_data AS (
            SELECT * FROM daily_summary
            WHERE ngay_bao_cao = (
                SELECT MAX(ngay_bao_cao) FROM daily_summary
                WHERE ngay_bao_cao < (SELECT MAX(ngay_bao_cao) FROM daily_summary)
            )
        )
        SELECT
            t.ngay_bao_cao as hom_nay,
            y.ngay_bao_cao as hom_qua,
            t.pttb_thuc_tang as pttb_thuc_tang_hom_nay,
            y.pttb_thuc_tang as pttb_thuc_tang_hom_qua,
            t.pttb_thuc_tang - COALESCE(y.pttb_thuc_tang, 0) as pttb_thuc_tang_change,
            t.mytv_thuc_tang as mytv_thuc_tang_hom_nay,
            y.mytv_thuc_tang as mytv_thuc_tang_hom_qua,
            t.mytv_thuc_tang - COALESCE(y.mytv_thuc_tang, 0) as mytv_thuc_tang_change
        FROM today_data t
        LEFT JOIN yesterday_data y ON 1=1
    ''')

    print("   - v_growth_trend_30d: OK")
    print("   - v_don_vi_performance: OK")
    print("   - v_nvkt_ranking_month: OK")
    print("   - v_kr_trend: OK")
    print("   - v_dod_comparison: OK")

    conn.commit()
    conn.close()

    print("\n" + "=" * 60)
    print(f"HOÀN THÀNH! Database đã được tạo tại: {DB_PATH}")
    print("=" * 60)

    # Hiển thị thống kê
    show_database_stats()


def show_database_stats():
    """Hiển thị thống kê database"""
    conn = get_connection()
    cursor = conn.cursor()

    print("\nThống kê database:")
    print("-" * 40)

    # Đếm số tables
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = cursor.fetchall()
    print(f"Số tables: {len(tables)}")
    for t in tables:
        cursor.execute(f"SELECT COUNT(*) FROM {t['name']}")
        count = cursor.fetchone()[0]
        print(f"   - {t['name']}: {count} records")

    # Đếm số views
    cursor.execute("SELECT name FROM sqlite_master WHERE type='view' ORDER BY name")
    views = cursor.fetchall()
    print(f"\nSố views: {len(views)}")
    for v in views:
        print(f"   - {v['name']}")

    # Đếm số indexes
    cursor.execute("SELECT name FROM sqlite_master WHERE type='index' AND name NOT LIKE 'sqlite_%' ORDER BY name")
    indexes = cursor.fetchall()
    print(f"\nSố indexes: {len(indexes)}")

    conn.close()


def check_date_exists(report_code: str, ngay_bao_cao: str) -> bool:
    """Kiểm tra ngày đã được import chưa"""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT COUNT(*) FROM import_log
        WHERE report_code = ? AND ngay_bao_cao = ? AND status = 'SUCCESS'
    ''', (report_code, ngay_bao_cao))

    count = cursor.fetchone()[0]
    conn.close()

    return count > 0


def log_import(report_code: str, ngay_bao_cao: str, file_name: str = None,
               file_path: str = None, record_count: int = 0,
               status: str = 'SUCCESS', error_message: str = None,
               duration_ms: int = None):
    """Ghi log import"""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute('''
        INSERT INTO import_log
        (report_code, ngay_bao_cao, file_name, file_path, record_count, status, error_message, import_duration_ms)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (report_code, ngay_bao_cao, file_name, file_path, record_count, status, error_message, duration_ms))

    conn.commit()
    conn.close()


if __name__ == '__main__':
    init_database()
