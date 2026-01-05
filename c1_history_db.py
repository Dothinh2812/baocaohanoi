# -*- coding: utf-8 -*-
"""
Database schema cho lưu trữ lịch sử chỉ tiêu C1.x
Lưu dữ liệu theo từng ngày với 3 cấp độ:
1. Toàn trung tâm (TTVT Sơn Tây)
2. Theo tổ (4 tổ)
3. Theo NVKT (nhân viên kỹ thuật)
"""

import sqlite3
import os
from datetime import datetime


DB_PATH = os.path.join(os.path.dirname(__file__), 'c1_history.db')


def get_connection():
    """Lấy kết nối database"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_database():
    """Khởi tạo database với schema"""
    conn = get_connection()
    cursor = conn.cursor()

    # =============================================
    # 1. BẢNG TỔNG HỢP TOÀN TRUNG TÂM (theo ngày)
    # =============================================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_tong_hop (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_cap_nhat DATE NOT NULL,

            -- C1.1: Chất lượng sửa chữa
            c11_sm1 INTEGER DEFAULT 0,
            c11_sm2 INTEGER DEFAULT 0,
            c11_ty_le_sua_chua_clcd REAL DEFAULT 0,
            c11_sm3 INTEGER DEFAULT 0,
            c11_sm4 INTEGER DEFAULT 0,
            c11_ty_le_brcd REAL DEFAULT 0,
            c11_bsc REAL DEFAULT 0,

            -- C1.2: Báo hỏng lặp lại / Sự cố
            c12_sm1 INTEGER DEFAULT 0,
            c12_sm2 INTEGER DEFAULT 0,
            c12_ty_le_bao_hong_lap_lai REAL DEFAULT 0,
            c12_sm3 INTEGER DEFAULT 0,
            c12_sm4 INTEGER DEFAULT 0,
            c12_ty_le_su_co REAL DEFAULT 0,
            c12_bsc REAL DEFAULT 0,

            -- C1.3: Kênh thuê riêng (TSL)
            c13_sm1 INTEGER DEFAULT 0,
            c13_sm2 INTEGER DEFAULT 0,
            c13_ty_le_sua_chua_tsl REAL DEFAULT 0,
            c13_sm3 INTEGER DEFAULT 0,
            c13_sm4 INTEGER DEFAULT 0,
            c13_ty_le_bao_hong_tsl REAL DEFAULT 0,
            c13_sm5 INTEGER DEFAULT 0,
            c13_sm6 INTEGER DEFAULT 0,
            c13_ty_le_su_co_tsl REAL DEFAULT 0,
            c13_bsc REAL DEFAULT 0,

            -- C1.4: Hài lòng khách hàng
            c14_tong_phieu INTEGER DEFAULT 0,
            c14_sl_da_ks INTEGER DEFAULT 0,
            c14_sl_ks_thanh_cong INTEGER DEFAULT 0,
            c14_sl_kh_hai_long INTEGER DEFAULT 0,
            c14_khong_hl_phuc_vu INTEGER DEFAULT 0,
            c14_ty_le_hl_phuc_vu REAL DEFAULT 0,
            c14_khong_hl_dich_vu INTEGER DEFAULT 0,
            c14_ty_le_hl_dich_vu REAL DEFAULT 0,
            c14_tong_phieu_hai_long INTEGER DEFAULT 0,
            c14_ty_le_kh_hai_long REAL DEFAULT 0,
            c14_bsc REAL DEFAULT 0,

            -- C1.5: Hoàn công đúng hạn
            c15_sm1 INTEGER DEFAULT 0,
            c15_sm2 INTEGER DEFAULT 0,
            c15_kq_thuc_hien REAL DEFAULT 0,
            c15_bsc REAL DEFAULT 0,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_cap_nhat)
        )
    ''')

    # =============================================
    # 2. BẢNG DỮ LIỆU THEO TỔ (4 tổ + Tổng)
    # =============================================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_theo_to (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_cap_nhat DATE NOT NULL,
            ten_to TEXT NOT NULL,

            -- C1.1: Chất lượng sửa chữa
            c11_sm1 INTEGER DEFAULT 0,
            c11_sm2 INTEGER DEFAULT 0,
            c11_ty_le_sua_chua_clcd REAL DEFAULT 0,
            c11_sm3 INTEGER DEFAULT 0,
            c11_sm4 INTEGER DEFAULT 0,
            c11_ty_le_brcd REAL DEFAULT 0,
            c11_bsc REAL DEFAULT 0,

            -- C1.2: Báo hỏng lặp lại / Sự cố
            c12_sm1 INTEGER DEFAULT 0,
            c12_sm2 INTEGER DEFAULT 0,
            c12_ty_le_bao_hong_lap_lai REAL DEFAULT 0,
            c12_sm3 INTEGER DEFAULT 0,
            c12_sm4 INTEGER DEFAULT 0,
            c12_ty_le_su_co REAL DEFAULT 0,
            c12_bsc REAL DEFAULT 0,

            -- C1.3: Kênh thuê riêng (TSL)
            c13_sm1 INTEGER DEFAULT 0,
            c13_sm2 INTEGER DEFAULT 0,
            c13_ty_le_sua_chua_tsl REAL DEFAULT 0,
            c13_sm3 INTEGER DEFAULT 0,
            c13_sm4 INTEGER DEFAULT 0,
            c13_ty_le_bao_hong_tsl REAL DEFAULT 0,
            c13_sm5 INTEGER DEFAULT 0,
            c13_sm6 INTEGER DEFAULT 0,
            c13_ty_le_su_co_tsl REAL DEFAULT 0,
            c13_bsc REAL DEFAULT 0,

            -- C1.4: Hài lòng khách hàng
            c14_tong_phieu INTEGER DEFAULT 0,
            c14_sl_da_ks INTEGER DEFAULT 0,
            c14_sl_ks_thanh_cong INTEGER DEFAULT 0,
            c14_sl_kh_hai_long INTEGER DEFAULT 0,
            c14_khong_hl_phuc_vu INTEGER DEFAULT 0,
            c14_ty_le_hl_phuc_vu REAL DEFAULT 0,
            c14_khong_hl_dich_vu INTEGER DEFAULT 0,
            c14_ty_le_hl_dich_vu REAL DEFAULT 0,
            c14_tong_phieu_hai_long INTEGER DEFAULT 0,
            c14_ty_le_kh_hai_long REAL DEFAULT 0,
            c14_bsc REAL DEFAULT 0,

            -- C1.5: Hoàn công đúng hạn
            c15_sm1 INTEGER DEFAULT 0,
            c15_sm2 INTEGER DEFAULT 0,
            c15_kq_thuc_hien REAL DEFAULT 0,
            c15_bsc REAL DEFAULT 0,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_cap_nhat, ten_to)
        )
    ''')

    # =============================================
    # 3. BẢNG DỮ LIỆU THEO NVKT (nhân viên kỹ thuật)
    # =============================================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_theo_nvkt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_cap_nhat DATE NOT NULL,
            ten_to TEXT NOT NULL,
            ten_nvkt TEXT NOT NULL,

            -- C1.4: Hài lòng khách hàng (từ TH_HL_NVKT)
            c14_tong_phieu_ks_thanh_cong INTEGER DEFAULT 0,
            c14_tong_phieu_hai_long INTEGER DEFAULT 0,
            c14_ty_le_hai_long REAL DEFAULT 0,

            -- C1.5: Hoàn công đúng hạn (từ KQ_C15_chitiet)
            c15_phieu_dat INTEGER DEFAULT 0,
            c15_tong_hoan_cong INTEGER DEFAULT 0,
            c15_ty_le_dat REAL DEFAULT 0,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(ngay_cap_nhat, ten_to, ten_nvkt)
        )
    ''')

    # =============================================
    # 4. BẢNG CHI TIẾT C1.4 (dữ liệu phiếu hài lòng)
    # =============================================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_4_chi_tiet (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_cap_nhat DATE NOT NULL,

            ma_tb TEXT,
            baohong_id TEXT,
            hdtb_id TEXT,
            nguoi_tl TEXT,
            dia_chi_ld TEXT,
            dien_thoai_ks TEXT,
            dien_thoai_lh TEXT,
            ghi_chu TEXT,
            nguoi_cn TEXT,
            do_hl TEXT,
            ma_tl TEXT,
            hl TEXT,
            ktc TEXT,
            ktm TEXT,
            khl_kt TEXT,
            khl_kd TEXT,
            nd_ktc_ktm TEXT,
            ten_dv_hni TEXT,
            ngay_hoi DATE,
            ngay_hc DATE,
            ten_kv TEXT,
            doi_vt TEXT,
            ttvt TEXT,
            nguoi_khoa TEXT,
            ten_nvkt_db TEXT,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # =============================================
    # 5. BẢNG CHI TIẾT C1.5 (dữ liệu hoàn công)
    # =============================================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_5_chi_tiet (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_cap_nhat DATE NOT NULL,

            ma_tb TEXT,
            hdtb_id TEXT,
            ma_gd TEXT,
            ten_dvvt_hni TEXT,
            so_ngay_hoan_thanh INTEGER,
            so_gio_hoan_thanh REAL,
            toanha_id TEXT,
            ten_kieu_ld TEXT,
            ngay_giao_phieu DATE,
            ngay_hc DATE,
            dat_chi_tieu TEXT,
            ten_kv TEXT,
            nguoi_khoa TEXT,
            nvkt_dia_ban TEXT,
            diem_chia TEXT,
            doi_vt TEXT,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # =============================================
    # 6. BẢNG LOG IMPORT
    # =============================================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS c1_import_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_import TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            loai_du_lieu TEXT NOT NULL,
            so_ban_ghi INTEGER DEFAULT 0,
            trang_thai TEXT DEFAULT 'success',
            ghi_chu TEXT
        )
    ''')

    # =============================================
    # TẠO INDEX
    # =============================================
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_tong_hop_ngay
        ON c1_tong_hop(ngay_cap_nhat)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_theo_to_ngay
        ON c1_theo_to(ngay_cap_nhat)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_theo_to_ten
        ON c1_theo_to(ten_to)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_theo_nvkt_ngay
        ON c1_theo_nvkt(ngay_cap_nhat)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_theo_nvkt_to
        ON c1_theo_nvkt(ten_to)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_theo_nvkt_ten
        ON c1_theo_nvkt(ten_nvkt)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_4_chi_tiet_ngay
        ON c1_4_chi_tiet(ngay_cap_nhat)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_c1_5_chi_tiet_ngay
        ON c1_5_chi_tiet(ngay_cap_nhat)
    ''')

    # =============================================
    # TẠO VIEWS CHO DASHBOARD
    # =============================================

    # View: Xu hướng BSC 30 ngày gần nhất
    cursor.execute('''
        CREATE VIEW IF NOT EXISTS v_c1_bsc_trend_30d AS
        SELECT
            ngay_cap_nhat,
            c11_bsc,
            c12_bsc,
            c13_bsc,
            c14_bsc,
            c15_bsc,
            (c11_bsc + c12_bsc + c13_bsc + c14_bsc + c15_bsc) / 5 as bsc_trung_binh
        FROM c1_tong_hop
        WHERE ngay_cap_nhat >= date('now', '-30 days')
        ORDER BY ngay_cap_nhat
    ''')

    # View: So sánh BSC theo tổ
    cursor.execute('''
        CREATE VIEW IF NOT EXISTS v_c1_bsc_theo_to AS
        SELECT
            ngay_cap_nhat,
            ten_to,
            c11_bsc,
            c12_bsc,
            c13_bsc,
            c14_bsc,
            c15_bsc,
            (c11_bsc + c12_bsc + c13_bsc + c14_bsc + c15_bsc) / 5 as bsc_trung_binh
        FROM c1_theo_to
        WHERE ten_to != 'Tổng'
        ORDER BY ngay_cap_nhat, ten_to
    ''')

    # View: Xếp hạng NVKT theo C1.4 hài lòng
    cursor.execute('''
        CREATE VIEW IF NOT EXISTS v_c1_nvkt_ranking_c14 AS
        SELECT
            ngay_cap_nhat,
            ten_to,
            ten_nvkt,
            c14_tong_phieu_ks_thanh_cong,
            c14_tong_phieu_hai_long,
            c14_ty_le_hai_long,
            RANK() OVER (PARTITION BY ngay_cap_nhat ORDER BY c14_ty_le_hai_long DESC) as rank_c14
        FROM c1_theo_nvkt
        WHERE c14_tong_phieu_ks_thanh_cong > 0
        ORDER BY ngay_cap_nhat DESC, c14_ty_le_hai_long DESC
    ''')

    # View: Xếp hạng NVKT theo C1.5 hoàn công
    cursor.execute('''
        CREATE VIEW IF NOT EXISTS v_c1_nvkt_ranking_c15 AS
        SELECT
            ngay_cap_nhat,
            ten_to,
            ten_nvkt,
            c15_phieu_dat,
            c15_tong_hoan_cong,
            c15_ty_le_dat,
            RANK() OVER (PARTITION BY ngay_cap_nhat ORDER BY c15_ty_le_dat DESC) as rank_c15
        FROM c1_theo_nvkt
        WHERE c15_tong_hoan_cong > 0
        ORDER BY ngay_cap_nhat DESC, c15_ty_le_dat DESC
    ''')

    conn.commit()
    conn.close()

    print(f"✅ Đã tạo database C1: {DB_PATH}")
    print("   Các bảng:")
    print("   - c1_tong_hop: Tổng hợp toàn trung tâm theo ngày")
    print("   - c1_theo_to: Dữ liệu theo 4 tổ")
    print("   - c1_theo_nvkt: Dữ liệu theo NVKT")
    print("   - c1_4_chi_tiet: Chi tiết phiếu hài lòng")
    print("   - c1_5_chi_tiet: Chi tiết hoàn công")
    print("   - c1_import_log: Log import")
    print("   Views:")
    print("   - v_c1_bsc_trend_30d: Xu hướng BSC 30 ngày")
    print("   - v_c1_bsc_theo_to: So sánh BSC theo tổ")
    print("   - v_c1_nvkt_ranking_c14: Xếp hạng NVKT theo hài lòng")
    print("   - v_c1_nvkt_ranking_c15: Xếp hạng NVKT theo hoàn công")


def show_schema():
    """Hiển thị schema database"""
    conn = get_connection()
    cursor = conn.cursor()

    # Lấy danh sách bảng
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = cursor.fetchall()

    print("\n=== SCHEMA DATABASE C1 ===")
    for table in tables:
        table_name = table['name']
        print(f"\n📋 Bảng: {table_name}")

        cursor.execute(f"PRAGMA table_info({table_name})")
        columns = cursor.fetchall()

        for col in columns:
            print(f"   - {col['name']}: {col['type']}")

    # Lấy danh sách views
    cursor.execute("SELECT name FROM sqlite_master WHERE type='view' ORDER BY name")
    views = cursor.fetchall()

    print("\n📊 Views:")
    for view in views:
        print(f"   - {view['name']}")

    conn.close()


if __name__ == "__main__":
    init_database()
    show_schema()
