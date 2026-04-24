-- Schema SQLite cho import du lieu tong hop theo tung sheet cua tung loai bao cao
-- Nguyen tac:
-- 1. Moi loai bao cao co nhieu bang du lieu, moi sheet tong hop la 1 bang rieng.
-- 2. Chi import cac sheet sau xu ly, khong import raw data.
-- 3. Moi bao cao chi co 1 snapshot hop le cho moi (ma_bao_cao, ngay_du_lieu).
-- 4. Chay lai cung ngay se ghi de snapshot cu trong transaction.
-- 5. Luon ghi lai file nguon processed/daily da dung de import.

PRAGMA foreign_keys = ON;

BEGIN;

CREATE TABLE IF NOT EXISTS cau_hinh_import_tong_hop (
    khoa TEXT PRIMARY KEY,
    gia_tri TEXT NOT NULL,
    mo_ta TEXT,
    thoi_gian_cap_nhat TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

INSERT INTO cau_hinh_import_tong_hop (khoa, gia_tri, mo_ta, thoi_gian_cap_nhat)
VALUES (
    'schema_version',
    'summary_per_report_sheet_v3_csv_whitelist_file_sheet',
    'Chi import sheet co trong list of table.csv; bang du lieu dat ten theo folder_name + file_name_rut_gon + sheet_name',
    CURRENT_TIMESTAMP
)
ON CONFLICT(khoa) DO UPDATE SET
    gia_tri = excluded.gia_tri,
    mo_ta = excluded.mo_ta,
    thoi_gian_cap_nhat = excluded.thoi_gian_cap_nhat;

CREATE TABLE IF NOT EXISTS danh_muc_bao_cao_tong_hop (
    ma_bao_cao TEXT PRIMARY KEY,
    ten_bao_cao TEXT NOT NULL,
    nhom_bao_cao TEXT NOT NULL,
    duong_dan_processed_mac_dinh TEXT,
    che_do_import TEXT NOT NULL DEFAULT 'summary_per_report_sheet',
    dang_su_dung INTEGER NOT NULL DEFAULT 1,
    mo_ta TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    thoi_gian_cap_nhat TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS danh_muc_bang_du_lieu_bao_cao (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_bao_cao TEXT NOT NULL,
    ten_sheet_goc TEXT NOT NULL,
    ma_sheet TEXT NOT NULL,
    ten_bang_du_lieu TEXT NOT NULL,
    che_do_luu_tru TEXT NOT NULL DEFAULT 'processed_summary_sheet_only',
    tong_so_cot INTEGER NOT NULL DEFAULT 0,
    danh_sach_cot_json TEXT,
    mo_ta TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    thoi_gian_cap_nhat TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ma_bao_cao) REFERENCES danh_muc_bao_cao_tong_hop(ma_bao_cao),
    UNIQUE (ma_bao_cao, ten_sheet_goc),
    UNIQUE (ma_bao_cao, ma_sheet),
    UNIQUE (ma_bao_cao, ten_bang_du_lieu)
);

CREATE INDEX IF NOT EXISTS idx_danh_muc_bang_du_lieu_bao_cao_ma_bao_cao
    ON danh_muc_bang_du_lieu_bao_cao(ma_bao_cao);
CREATE INDEX IF NOT EXISTS idx_danh_muc_bang_du_lieu_bao_cao_ten_bang
    ON danh_muc_bang_du_lieu_bao_cao(ten_bang_du_lieu);

CREATE TABLE IF NOT EXISTS bao_cao_tong_hop_ngay (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_bao_cao TEXT NOT NULL,
    ngay_du_lieu TEXT NOT NULL,
    tu_ngay TEXT,
    den_ngay TEXT,
    thang_bao_cao INTEGER,
    nam_bao_cao INTEGER,
    ten_tep_nguon TEXT,
    duong_dan_tep_nguon TEXT,
    ma_hash_tep TEXT,
    so_sheet_tong_hop INTEGER NOT NULL DEFAULT 0,
    so_bang_du_lieu INTEGER NOT NULL DEFAULT 0,
    so_dong_tong_hop INTEGER NOT NULL DEFAULT 0,
    so_chi_tieu_tong_hop INTEGER NOT NULL DEFAULT 0,
    trang_thai_nap TEXT NOT NULL DEFAULT 'cho_xu_ly'
        CHECK (trang_thai_nap IN ('cho_xu_ly', 'thanh_cong', 'that_bai', 'khong_co_sheet_tong_hop')),
    ghi_chu TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    thoi_gian_cap_nhat TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ma_bao_cao) REFERENCES danh_muc_bao_cao_tong_hop(ma_bao_cao),
    UNIQUE (ma_bao_cao, ngay_du_lieu)
);

CREATE INDEX IF NOT EXISTS idx_bao_cao_tong_hop_ngay_ma_bao_cao
    ON bao_cao_tong_hop_ngay(ma_bao_cao);
CREATE INDEX IF NOT EXISTS idx_bao_cao_tong_hop_ngay_ngay_du_lieu
    ON bao_cao_tong_hop_ngay(ngay_du_lieu);

CREATE TABLE IF NOT EXISTS tep_nguon_bao_cao_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_tong_hop_ngay_id INTEGER NOT NULL,
    loai_tep TEXT NOT NULL
        CHECK (loai_tep IN ('processed', 'daily')),
    ten_tep TEXT NOT NULL,
    duong_dan_tep TEXT NOT NULL,
    la_tep_nap_chinh INTEGER NOT NULL DEFAULT 0,
    kich_thuoc_byte INTEGER,
    ma_hash_tep TEXT,
    thoi_gian_sua_file TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (bao_cao_tong_hop_ngay_id) REFERENCES bao_cao_tong_hop_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_tong_hop_ngay_id, loai_tep, duong_dan_tep)
);

CREATE INDEX IF NOT EXISTS idx_tep_nguon_bao_cao_tong_hop_snapshot
    ON tep_nguon_bao_cao_tong_hop(bao_cao_tong_hop_ngay_id);

CREATE TABLE IF NOT EXISTS sheet_bao_cao_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_tong_hop_ngay_id INTEGER NOT NULL,
    ten_sheet TEXT NOT NULL,
    ma_sheet TEXT NOT NULL,
    ten_bang_du_lieu TEXT NOT NULL,
    thu_tu_sheet INTEGER NOT NULL DEFAULT 1,
    tong_so_cot INTEGER NOT NULL DEFAULT 0,
    so_dong_tong_hop INTEGER NOT NULL DEFAULT 0,
    so_chi_tieu_tong_hop INTEGER NOT NULL DEFAULT 0,
    danh_sach_cot_json TEXT,
    cot_chieu_json TEXT,
    cot_chi_tieu_json TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (bao_cao_tong_hop_ngay_id) REFERENCES bao_cao_tong_hop_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_tong_hop_ngay_id, ten_sheet)
);

CREATE INDEX IF NOT EXISTS idx_sheet_bao_cao_tong_hop_snapshot
    ON sheet_bao_cao_tong_hop(bao_cao_tong_hop_ngay_id);
CREATE INDEX IF NOT EXISTS idx_sheet_bao_cao_tong_hop_bang
    ON sheet_bao_cao_tong_hop(ten_bang_du_lieu);

CREATE TABLE IF NOT EXISTS nhat_ky_nap_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_bao_cao TEXT NOT NULL,
    ngay_du_lieu TEXT NOT NULL,
    so_bang_du_lieu INTEGER NOT NULL DEFAULT 0,
    danh_sach_bang_du_lieu TEXT,
    che_do_nap TEXT NOT NULL DEFAULT 'summary_per_report_sheet_overwrite_same_day',
    bat_dau_luc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    ket_thuc_luc TEXT,
    trang_thai TEXT NOT NULL DEFAULT 'dang_chay'
        CHECK (trang_thai IN ('dang_chay', 'thanh_cong', 'that_bai', 'bo_qua', 'khong_co_sheet_tong_hop')),
    so_sheet_tong_hop INTEGER NOT NULL DEFAULT 0,
    so_dong_tong_hop INTEGER NOT NULL DEFAULT 0,
    so_chi_tieu_tong_hop INTEGER NOT NULL DEFAULT 0,
    ma_hash_tep TEXT,
    thong_diep TEXT
);

CREATE INDEX IF NOT EXISTS idx_nhat_ky_nap_tong_hop_ngay
    ON nhat_ky_nap_tong_hop(ngay_du_lieu, ma_bao_cao);

COMMIT;
