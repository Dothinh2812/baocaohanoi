#!/usr/bin/env python3
"""
Script import dữ liệu báo cáo từ Excel vào SQLite database.
Hỗ trợ lưu trữ dữ liệu hàng ngày để tạo báo cáo lịch sử.

Usage:
    python import_baocao.py                      # Import với ngày hiện tại
    python import_baocao.py --date 2024-12-30    # Import với ngày chỉ định
    python import_baocao.py --init               # Chỉ khởi tạo database (không import)
"""

import argparse
import sqlite3
from datetime import date, datetime
from pathlib import Path
import pandas as pd


# Cấu hình
DB_PATH = Path(__file__).parent / "baocao_hanoi.db"
REPORT_DIR = Path(__file__).parent / "downloads" / "baocao_hanoi"

REPORT_FILES = {
    "c11": ("c1.1 report.xlsx", "TH_C1.1"),
    "c12": ("c1.2 report.xlsx", "TH_C1.2"),
    "c13": ("c1.3 report.xlsx", "TH_C1.3"),
    "c14": ("c1.4 report.xlsx", "TH_C1.4"),
    "c14_nvkt": ("c1.4_chitiet_report.xlsx", "TH_HL_NVKT"),
    "sm1c12": ("SM1-C12.xlsx", "TH_SM1C12_HLL_Thang"),
    "sm4c11_chitiet": ("SM4-C11.xlsx", "chi_tiet"),
    "sm4c11_18h": ("SM4-C11.xlsx", "chi_tieu_ko_hen_18h"),
}


def init_database(conn: sqlite3.Connection):
    """Khởi tạo schema database."""
    cursor = conn.cursor()
    
    # Bảng đơn vị
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS don_vi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ten_don_vi TEXT UNIQUE NOT NULL
        )
    """)
    
    # Bảng nhân viên kỹ thuật
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS nhan_vien_kt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            don_vi_id INTEGER REFERENCES don_vi(id),
            ten_nvkt TEXT NOT NULL,
            UNIQUE(don_vi_id, ten_nvkt)
        )
    """)
    
    # Báo cáo C1.1
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_c11 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            don_vi_id INTEGER REFERENCES don_vi(id),
            sm1_cl_chu_dong INTEGER,
            sm2_cl_chu_dong INTEGER,
            ty_le_cl_chu_dong REAL,
            sm3_brcd INTEGER,
            sm4_brcd INTEGER,
            ty_le_brcd REAL,
            chi_tieu_bsc REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, don_vi_id)
        )
    """)
    
    # Báo cáo C1.2
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_c12 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            don_vi_id INTEGER REFERENCES don_vi(id),
            sm1_lap_lai INTEGER,
            sm2_lap_lai INTEGER,
            ty_le_lap_lai REAL,
            sm3_su_co INTEGER,
            sm4_su_co INTEGER,
            ty_le_su_co REAL,
            chi_tieu_bsc REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, don_vi_id)
        )
    """)
    
    # Báo cáo C1.3
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_c13 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            don_vi_id INTEGER REFERENCES don_vi(id),
            sm1_sua_chua INTEGER,
            sm2_sua_chua INTEGER,
            ty_le_sua_chua REAL,
            sm3_lap_lai INTEGER,
            sm4_lap_lai INTEGER,
            ty_le_lap_lai REAL,
            sm5_su_co INTEGER,
            sm6_su_co INTEGER,
            ty_le_su_co REAL,
            chi_tieu_bsc REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, don_vi_id)
        )
    """)
    
    # Báo cáo C1.4
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_c14 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            don_vi_id INTEGER REFERENCES don_vi(id),
            tong_phieu INTEGER,
            sl_da_ks INTEGER,
            sl_ks_thanh_cong INTEGER,
            sl_kh_hai_long INTEGER,
            khong_hl_kt_phuc_vu INTEGER,
            ty_le_hl_kt_phuc_vu REAL,
            khong_hl_kt_dich_vu INTEGER,
            ty_le_hl_kt_dich_vu REAL,
            tong_phieu_hai_long_kt INTEGER,
            ty_le_kh_hai_long REAL,
            diem_bsc REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, don_vi_id)
        )
    """)
    
    # Báo cáo C1.4 Chi tiết NVKT
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_c14_nvkt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            nvkt_id INTEGER REFERENCES nhan_vien_kt(id),
            tong_phieu_ks_thanh_cong INTEGER,
            tong_phieu_khl INTEGER,
            ty_le_hai_long REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, nvkt_id)
        )
    """)
    
    # Báo cáo SM1-C12: Hỏng lại tháng theo NVKT
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_sm1c12_hll (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            nvkt_id INTEGER REFERENCES nhan_vien_kt(id),
            so_phieu_hll INTEGER,
            so_phieu_bao_hong INTEGER,
            ty_le_hll REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, nvkt_id)
        )
    """)
    
    # Báo cáo SM4-C11: Chi tiết BRCD không hẹn theo NVKT
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_sm4c11_chitiet (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            nvkt_id INTEGER REFERENCES nhan_vien_kt(id),
            tong_phieu INTEGER,
            so_phieu_dat INTEGER,
            ty_le_dat REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, nvkt_id)
        )
    """)
    
    # Báo cáo SM4-C11: Chỉ tiêu không hẹn 18h theo NVKT
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bao_cao_sm4c11_18h (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_bao_cao DATE NOT NULL,
            nvkt_id INTEGER REFERENCES nhan_vien_kt(id),
            tong_phieu INTEGER,
            so_phieu_dat INTEGER,
            ty_le_dat REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay_bao_cao, nvkt_id)
        )
    """)
    
    # Tạo indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_c11_ngay ON bao_cao_c11(ngay_bao_cao)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_c12_ngay ON bao_cao_c12(ngay_bao_cao)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_c13_ngay ON bao_cao_c13(ngay_bao_cao)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_c14_ngay ON bao_cao_c14(ngay_bao_cao)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_c14_nvkt_ngay ON bao_cao_c14_nvkt(ngay_bao_cao)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sm1c12_hll_ngay ON bao_cao_sm1c12_hll(ngay_bao_cao)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sm4c11_chitiet_ngay ON bao_cao_sm4c11_chitiet(ngay_bao_cao)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sm4c11_18h_ngay ON bao_cao_sm4c11_18h(ngay_bao_cao)")
    
    conn.commit()
    print("✓ Database schema initialized")


def get_or_create_don_vi(conn: sqlite3.Connection, ten_don_vi: str) -> int:
    """Lấy hoặc tạo đơn vị, trả về ID."""
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM don_vi WHERE ten_don_vi = ?", (ten_don_vi,))
    row = cursor.fetchone()
    if row:
        return row[0]
    cursor.execute("INSERT INTO don_vi (ten_don_vi) VALUES (?)", (ten_don_vi,))
    conn.commit()
    return cursor.lastrowid


def get_or_create_nvkt(conn: sqlite3.Connection, don_vi_id: int, ten_nvkt: str) -> int:
    """Lấy hoặc tạo nhân viên KT, trả về ID."""
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id FROM nhan_vien_kt WHERE don_vi_id = ? AND ten_nvkt = ?",
        (don_vi_id, ten_nvkt)
    )
    row = cursor.fetchone()
    if row:
        return row[0]
    cursor.execute(
        "INSERT INTO nhan_vien_kt (don_vi_id, ten_nvkt) VALUES (?, ?)",
        (don_vi_id, ten_nvkt)
    )
    conn.commit()
    return cursor.lastrowid


def import_c11(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo C1.1."""
    file_path = REPORT_DIR / REPORT_FILES["c11"][0]
    sheet_name = REPORT_FILES["c11"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row.iloc[0]
        if pd.isna(ten_don_vi) or ten_don_vi == "Tổng":
            continue
            
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_c11 
            (ngay_bao_cao, don_vi_id, sm1_cl_chu_dong, sm2_cl_chu_dong, ty_le_cl_chu_dong,
             sm3_brcd, sm4_brcd, ty_le_brcd, chi_tieu_bsc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ngay, don_vi_id,
            int(row.iloc[1]) if pd.notna(row.iloc[1]) else None,
            int(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
            float(row.iloc[3]) if pd.notna(row.iloc[3]) else None,
            int(row.iloc[4]) if pd.notna(row.iloc[4]) else None,
            int(row.iloc[5]) if pd.notna(row.iloc[5]) else None,
            float(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
            float(row.iloc[7]) if pd.notna(row.iloc[7]) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ C1.1: {count} records imported")


def import_c12(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo C1.2."""
    file_path = REPORT_DIR / REPORT_FILES["c12"][0]
    sheet_name = REPORT_FILES["c12"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row.iloc[0]
        if pd.isna(ten_don_vi) or ten_don_vi == "Tổng":
            continue
            
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_c12 
            (ngay_bao_cao, don_vi_id, sm1_lap_lai, sm2_lap_lai, ty_le_lap_lai,
             sm3_su_co, sm4_su_co, ty_le_su_co, chi_tieu_bsc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ngay, don_vi_id,
            int(row.iloc[1]) if pd.notna(row.iloc[1]) else None,
            int(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
            float(row.iloc[3]) if pd.notna(row.iloc[3]) else None,
            int(row.iloc[4]) if pd.notna(row.iloc[4]) else None,
            int(row.iloc[5]) if pd.notna(row.iloc[5]) else None,
            float(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
            float(row.iloc[7]) if pd.notna(row.iloc[7]) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ C1.2: {count} records imported")


def import_c13(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo C1.3."""
    file_path = REPORT_DIR / REPORT_FILES["c13"][0]
    sheet_name = REPORT_FILES["c13"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row.iloc[0]
        if pd.isna(ten_don_vi) or ten_don_vi == "Tổng":
            continue
            
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_c13 
            (ngay_bao_cao, don_vi_id, sm1_sua_chua, sm2_sua_chua, ty_le_sua_chua,
             sm3_lap_lai, sm4_lap_lai, ty_le_lap_lai, sm5_su_co, sm6_su_co, ty_le_su_co, chi_tieu_bsc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ngay, don_vi_id,
            int(row.iloc[1]) if pd.notna(row.iloc[1]) else None,
            int(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
            float(row.iloc[3]) if pd.notna(row.iloc[3]) else None,
            int(row.iloc[4]) if pd.notna(row.iloc[4]) else None,
            int(row.iloc[5]) if pd.notna(row.iloc[5]) else None,
            float(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
            int(row.iloc[7]) if pd.notna(row.iloc[7]) else None,
            int(row.iloc[8]) if pd.notna(row.iloc[8]) else None,
            float(row.iloc[9]) if pd.notna(row.iloc[9]) else None,
            float(row.iloc[10]) if pd.notna(row.iloc[10]) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ C1.3: {count} records imported")


def import_c14(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo C1.4."""
    file_path = REPORT_DIR / REPORT_FILES["c14"][0]
    sheet_name = REPORT_FILES["c14"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row.iloc[0]
        if pd.isna(ten_don_vi) or ten_don_vi == "Tổng":
            continue
            
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_c14 
            (ngay_bao_cao, don_vi_id, tong_phieu, sl_da_ks, sl_ks_thanh_cong, sl_kh_hai_long,
             khong_hl_kt_phuc_vu, ty_le_hl_kt_phuc_vu, khong_hl_kt_dich_vu, ty_le_hl_kt_dich_vu,
             tong_phieu_hai_long_kt, ty_le_kh_hai_long, diem_bsc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ngay, don_vi_id,
            int(row.iloc[1]) if pd.notna(row.iloc[1]) else None,
            int(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
            int(row.iloc[3]) if pd.notna(row.iloc[3]) else None,
            int(row.iloc[4]) if pd.notna(row.iloc[4]) else None,
            int(row.iloc[5]) if pd.notna(row.iloc[5]) else None,
            float(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
            int(row.iloc[7]) if pd.notna(row.iloc[7]) else None,
            float(row.iloc[8]) if pd.notna(row.iloc[8]) else None,
            int(row.iloc[9]) if pd.notna(row.iloc[9]) else None,
            float(row.iloc[10]) if pd.notna(row.iloc[10]) else None,
            float(row.iloc[11]) if pd.notna(row.iloc[11]) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ C1.4: {count} records imported")


def import_c14_nvkt(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo C1.4 chi tiết NVKT."""
    file_path = REPORT_DIR / REPORT_FILES["c14_nvkt"][0]
    sheet_name = REPORT_FILES["c14_nvkt"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row.iloc[0]  # DOIVT
        ten_nvkt = row.iloc[1]    # NVKT
        
        if pd.isna(ten_don_vi) or pd.isna(ten_nvkt):
            continue
            
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        nvkt_id = get_or_create_nvkt(conn, don_vi_id, ten_nvkt)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_c14_nvkt 
            (ngay_bao_cao, nvkt_id, tong_phieu_ks_thanh_cong, tong_phieu_khl, ty_le_hai_long)
            VALUES (?, ?, ?, ?, ?)
        """, (
            ngay, nvkt_id,
            int(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
            int(row.iloc[3]) if pd.notna(row.iloc[3]) else None,
            float(row.iloc[4]) if pd.notna(row.iloc[4]) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ C1.4 NVKT: {count} records imported")


def import_sm1c12_hll(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo SM1-C12: Tỷ lệ hỏng lại tháng theo NVKT."""
    file_path = REPORT_DIR / REPORT_FILES["sm1c12"][0]
    sheet_name = REPORT_FILES["sm1c12"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row['TEN_DOI']
        ten_nvkt = row['NVKT']
        
        if pd.isna(ten_don_vi) or pd.isna(ten_nvkt):
            continue
        
        # Chuẩn hóa tên đơn vị thành tên ngắn
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        nvkt_id = get_or_create_nvkt(conn, don_vi_id, ten_nvkt)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_sm1c12_hll 
            (ngay_bao_cao, nvkt_id, so_phieu_hll, so_phieu_bao_hong, ty_le_hll)
            VALUES (?, ?, ?, ?, ?)
        """, (
            ngay, nvkt_id,
            int(row['Số phiếu HLL']) if pd.notna(row['Số phiếu HLL']) else None,
            int(row['Số phiếu báo hỏng']) if pd.notna(row['Số phiếu báo hỏng']) else None,
            float(row['Tỉ lệ HLL tháng (2.5%)']) if pd.notna(row['Tỉ lệ HLL tháng (2.5%)']) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ SM1-C12 HLL: {count} records imported")


def import_sm4c11_chitiet(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo SM4-C11: Chi tiết BRCD không hẹn theo NVKT."""
    file_path = REPORT_DIR / REPORT_FILES["sm4c11_chitiet"][0]
    sheet_name = REPORT_FILES["sm4c11_chitiet"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    # Lấy tên cột tỷ lệ (có thể dài)
    ty_le_col = [c for c in df.columns if 'Tỷ lệ' in c][0]
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row['TEN_DOI']
        ten_nvkt = row['NVKT']
        
        if pd.isna(ten_don_vi) or pd.isna(ten_nvkt):
            continue
        
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        nvkt_id = get_or_create_nvkt(conn, don_vi_id, ten_nvkt)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_sm4c11_chitiet 
            (ngay_bao_cao, nvkt_id, tong_phieu, so_phieu_dat, ty_le_dat)
            VALUES (?, ?, ?, ?, ?)
        """, (
            ngay, nvkt_id,
            int(row['Tổng phiếu']) if pd.notna(row['Tổng phiếu']) else None,
            int(row['Số phiếu đạt']) if pd.notna(row['Số phiếu đạt']) else None,
            float(row[ty_le_col]) if pd.notna(row[ty_le_col]) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ SM4-C11 Chi tiết: {count} records imported")


def import_sm4c11_18h(conn: sqlite3.Connection, ngay: str):
    """Import báo cáo SM4-C11: Chỉ tiêu không hẹn 18h theo NVKT."""
    file_path = REPORT_DIR / REPORT_FILES["sm4c11_18h"][0]
    sheet_name = REPORT_FILES["sm4c11_18h"][1]
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cursor = conn.cursor()
    
    # Lấy tên cột tỷ lệ (có thể dài)
    ty_le_col = [c for c in df.columns if 'Tỷ lệ' in c][0]
    
    count = 0
    for _, row in df.iterrows():
        ten_don_vi = row['TEN_DOI']
        ten_nvkt = row['NVKT']
        
        if pd.isna(ten_don_vi) or pd.isna(ten_nvkt):
            continue
        
        don_vi_id = get_or_create_don_vi(conn, ten_don_vi)
        nvkt_id = get_or_create_nvkt(conn, don_vi_id, ten_nvkt)
        
        cursor.execute("""
            INSERT OR REPLACE INTO bao_cao_sm4c11_18h 
            (ngay_bao_cao, nvkt_id, tong_phieu, so_phieu_dat, ty_le_dat)
            VALUES (?, ?, ?, ?, ?)
        """, (
            ngay, nvkt_id,
            int(row['Tổng phiếu']) if pd.notna(row['Tổng phiếu']) else None,
            int(row['Số phiếu đạt']) if pd.notna(row['Số phiếu đạt']) else None,
            float(row[ty_le_col]) if pd.notna(row[ty_le_col]) else None,
        ))
        count += 1
    
    conn.commit()
    print(f"  ✓ SM4-C11 18h: {count} records imported")


def main():
    parser = argparse.ArgumentParser(description="Import báo cáo từ Excel vào SQLite")
    parser.add_argument(
        "--date", "-d",
        type=str,
        default=date.today().isoformat(),
        help="Ngày báo cáo (YYYY-MM-DD), mặc định là ngày hiện tại"
    )
    parser.add_argument(
        "--init",
        action="store_true",
        help="Chỉ khởi tạo database, không import dữ liệu"
    )
    args = parser.parse_args()
    
    # Validate date format
    try:
        datetime.strptime(args.date, "%Y-%m-%d")
    except ValueError:
        print(f"❌ Ngày không hợp lệ: {args.date}. Định dạng: YYYY-MM-DD")
        return 1
    
    print(f"Database: {DB_PATH}")
    print(f"Report directory: {REPORT_DIR}")
    
    conn = sqlite3.connect(DB_PATH)
    
    try:
        init_database(conn)
        
        if args.init:
            print("✓ Database initialized (no data imported)")
            return 0
        
        print(f"\nImporting data for date: {args.date}")
        
        import_c11(conn, args.date)
        import_c12(conn, args.date)
        import_c13(conn, args.date)
        import_c14(conn, args.date)
        import_c14_nvkt(conn, args.date)
        import_sm1c12_hll(conn, args.date)
        import_sm4c11_chitiet(conn, args.date)
        import_sm4c11_18h(conn, args.date)
        
        print(f"\n✓ All reports imported successfully for {args.date}")
        
    finally:
        conn.close()
    
    return 0


if __name__ == "__main__":
    exit(main())
