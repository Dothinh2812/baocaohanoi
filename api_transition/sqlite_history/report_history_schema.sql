-- Schema SQLite cho api_transition/report_history.db
-- Nguyen tac:
-- 1. Moi bao cao chi co 1 ban ghi hop le cho moi (ma_bao_cao, ngay_du_lieu)
-- 2. Cac lan chay lai trong cung ngay se ghi de bang cach thay the du lieu con cua bao_cao_ngay
-- 3. Ten bang va ten cot dung tieng Viet khong dau, de hieu theo nghiep vu

PRAGMA foreign_keys = ON;

BEGIN;

CREATE TABLE IF NOT EXISTS danh_muc_bao_cao (
    ma_bao_cao TEXT PRIMARY KEY,
    ten_bao_cao TEXT NOT NULL,
    nhom_bao_cao TEXT NOT NULL,
    loai_dich_vu TEXT,
    hanh_dong_chinh TEXT,
    duong_dan_processed_mac_dinh TEXT,
    dang_su_dung INTEGER NOT NULL DEFAULT 1,
    mo_ta TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    thoi_gian_cap_nhat TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS bao_cao_ngay (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_bao_cao TEXT NOT NULL,
    ngay_du_lieu TEXT NOT NULL,
    tu_ngay TEXT,
    den_ngay TEXT,
    thang_bao_cao INTEGER,
    nam_bao_cao INTEGER,
    duong_dan_tep TEXT,
    ma_hash_tep TEXT,
    so_dong_goc INTEGER NOT NULL DEFAULT 0,
    so_dong_tong_hop INTEGER NOT NULL DEFAULT 0,
    so_dong_chi_tiet INTEGER NOT NULL DEFAULT 0,
    trang_thai_nap TEXT NOT NULL DEFAULT 'cho_xu_ly'
        CHECK (trang_thai_nap IN ('cho_xu_ly', 'thanh_cong', 'that_bai')),
    ghi_chu TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    thoi_gian_cap_nhat TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ma_bao_cao) REFERENCES danh_muc_bao_cao(ma_bao_cao),
    UNIQUE (ma_bao_cao, ngay_du_lieu)
);

CREATE INDEX IF NOT EXISTS idx_bao_cao_ngay_ngay_du_lieu
    ON bao_cao_ngay(ngay_du_lieu);
CREATE INDEX IF NOT EXISTS idx_bao_cao_ngay_ma_bao_cao
    ON bao_cao_ngay(ma_bao_cao);

CREATE TABLE IF NOT EXISTS sheet_bao_cao (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ten_sheet TEXT NOT NULL,
    loai_sheet TEXT NOT NULL
        CHECK (loai_sheet IN ('source', 'summary', 'detail', 'note', 'other')),
    thu_tu_sheet INTEGER NOT NULL DEFAULT 1,
    so_dong INTEGER NOT NULL DEFAULT 0,
    danh_sach_cot_json TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ten_sheet)
);

CREATE INDEX IF NOT EXISTS idx_sheet_bao_cao_bao_cao_ngay_id
    ON sheet_bao_cao(bao_cao_ngay_id);

CREATE TABLE IF NOT EXISTS dong_bao_cao_goc (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    sheet_bao_cao_id INTEGER,
    ten_sheet TEXT NOT NULL,
    so_dong INTEGER NOT NULL,
    khoa_ban_ghi TEXT,
    ttvt TEXT,
    doi_vien_thong TEXT,
    don_vi TEXT,
    ma_nv TEXT,
    ten_nv TEXT,
    du_lieu_json TEXT NOT NULL,
    ma_hash_dong TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    FOREIGN KEY (sheet_bao_cao_id) REFERENCES sheet_bao_cao(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ten_sheet, so_dong)
);

CREATE INDEX IF NOT EXISTS idx_dong_bao_cao_goc_bao_cao_ngay_id
    ON dong_bao_cao_goc(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_dong_bao_cao_goc_khoa_ban_ghi
    ON dong_bao_cao_goc(khoa_ban_ghi);

CREATE TABLE IF NOT EXISTS tep_luu_tru_bao_cao (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    loai_tep TEXT NOT NULL
        CHECK (loai_tep IN ('processed', 'daily', 'manifest', 'raw')),
    duong_dan_tep TEXT NOT NULL,
    kich_thuoc_byte INTEGER,
    ma_hash_tep TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, loai_tep, duong_dan_tep)
);

CREATE INDEX IF NOT EXISTS idx_tep_luu_tru_bao_cao_bao_cao_ngay_id
    ON tep_luu_tru_bao_cao(bao_cao_ngay_id);

CREATE TABLE IF NOT EXISTS nhat_ky_nap_bao_cao (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_bao_cao TEXT NOT NULL,
    ngay_du_lieu TEXT NOT NULL,
    che_do_ghi_de TEXT NOT NULL DEFAULT 'overwrite_cung_ngay',
    bat_dau_luc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    ket_thuc_luc TEXT,
    trang_thai TEXT NOT NULL DEFAULT 'dang_chay'
        CHECK (trang_thai IN ('dang_chay', 'thanh_cong', 'that_bai', 'bo_qua')),
    so_dong_goc INTEGER NOT NULL DEFAULT 0,
    so_dong_tong_hop INTEGER NOT NULL DEFAULT 0,
    so_dong_chi_tiet INTEGER NOT NULL DEFAULT 0,
    ma_hash_tep TEXT,
    thong_diep TEXT
);

CREATE INDEX IF NOT EXISTS idx_nhat_ky_nap_bao_cao_ngay
    ON nhat_ky_nap_bao_cao(ngay_du_lieu, ma_bao_cao);

CREATE TABLE IF NOT EXISTS danh_muc_don_vi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ttvt TEXT,
    doi_vien_thong TEXT,
    don_vi TEXT,
    ma_don_vi TEXT,
    khoa_chuan_hoa TEXT UNIQUE,
    du_lieu_bo_sung_json TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS danh_muc_nhan_vien (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_nv TEXT,
    ten_nv TEXT NOT NULL,
    ten_nv_chuan_hoa TEXT,
    ttvt TEXT,
    doi_vien_thong TEXT,
    don_vi TEXT,
    nguon_du_lieu TEXT,
    du_lieu_bo_sung_json TEXT,
    thoi_gian_tao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (ma_nv, ten_nv, doi_vien_thong)
);

CREATE INDEX IF NOT EXISTS idx_danh_muc_nhan_vien_ten_nv
    ON danh_muc_nhan_vien(ten_nv);

CREATE TABLE IF NOT EXISTS c11_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL,
    sm1 INTEGER,
    sm2 INTEGER,
    ty_le_sua_chua_chat_luong_chu_dong REAL,
    sm3 INTEGER,
    sm4 INTEGER,
    ty_le_bao_hong_brcd_dung_quy_dinh REAL,
    sm5 INTEGER,
    sm6 INTEGER,
    ty_le_sua_chua_trong_ngay_tai_ccco REAL,
    chi_tieu_bsc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi)
);

CREATE TABLE IF NOT EXISTS c11_chi_tiet_nvkt (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    doi_vien_thong TEXT NOT NULL,
    nvkt TEXT NOT NULL,
    tong_phieu INTEGER,
    so_phieu_dat INTEGER,
    ty_le_dat REAL,
    moc_gio TEXT NOT NULL DEFAULT 'tong',
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, doi_vien_thong, nvkt, moc_gio)
);

CREATE TABLE IF NOT EXISTS c12_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL,
    sm1 INTEGER,
    sm2 INTEGER,
    ty_le_hong_lap_lai REAL,
    sm3 INTEGER,
    sm4 INTEGER,
    ty_le_su_co_brcd REAL,
    chi_tieu_bsc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi)
);

CREATE TABLE IF NOT EXISTS c12_hong_lap_lai_nvkt (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    doi_vien_thong TEXT NOT NULL,
    nvkt TEXT NOT NULL,
    so_phieu_hll INTEGER,
    so_phieu_bao_hong INTEGER,
    ty_le_hll REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, doi_vien_thong, nvkt)
);

CREATE TABLE IF NOT EXISTS c13_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL,
    sm1 INTEGER,
    sm2 INTEGER,
    ty_le_sua_chua_dung_han REAL,
    sm3 INTEGER,
    sm4 INTEGER,
    ty_le_hong_lap_lai_kenh_tsl REAL,
    sm5 INTEGER,
    sm6 INTEGER,
    ty_le_su_co_kenh_tsl REAL,
    chi_tieu_bsc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi)
);

CREATE TABLE IF NOT EXISTS c14_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL,
    tong_phieu INTEGER,
    so_luong_da_khao_sat INTEGER,
    so_luong_khao_sat_thanh_cong INTEGER,
    so_luong_khach_hang_hai_long INTEGER,
    khong_hai_long_ky_thuat_phuc_vu INTEGER,
    ty_le_hai_long_ky_thuat_phuc_vu REAL,
    khong_hai_long_ky_thuat_dich_vu INTEGER,
    ty_le_hai_long_ky_thuat_dich_vu REAL,
    tong_phieu_hai_long_ky_thuat INTEGER,
    ty_le_khach_hang_hai_long REAL,
    diem_bsc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi)
);

CREATE TABLE IF NOT EXISTS c14_hai_long_nvkt (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    doi_vien_thong TEXT NOT NULL,
    nvkt TEXT NOT NULL,
    tong_phieu_khao_sat_thanh_cong INTEGER,
    tong_phieu_khong_hai_long INTEGER,
    ty_le_hai_long_nvkt REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, doi_vien_thong, nvkt)
);

CREATE TABLE IF NOT EXISTS ghtt_don_vi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL,
    ttvt TEXT NOT NULL DEFAULT '',
    hoan_thanh_t INTEGER,
    giao_nvkt_t INTEGER,
    ty_le_t REAL,
    hoan_thanh_t_cong_1 INTEGER,
    giao_nvkt_t_cong_1 INTEGER,
    ty_le_t_cong_1 REAL,
    so_luong_ghtt_lon_hon_6_thang INTEGER,
    hoan_thanh_lon_hon_6_thang_t_cong_1 INTEGER,
    ty_le_lon_hon_6_thang_t_cong_1 REAL,
    ty_le_tong REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi, ttvt)
);

CREATE TABLE IF NOT EXISTS ghtt_nvkt (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    nvkt TEXT NOT NULL,
    don_vi TEXT NOT NULL DEFAULT '',
    ttvt TEXT NOT NULL DEFAULT '',
    hoan_thanh_t INTEGER,
    giao_nvkt_t INTEGER,
    ty_le_t REAL,
    hoan_thanh_t_cong_1 INTEGER,
    giao_nvkt_t_cong_1 INTEGER,
    ty_le_t_cong_1 REAL,
    so_luong_ghtt_lon_hon_6_thang INTEGER,
    hoan_thanh_lon_hon_6_thang_t_cong_1 INTEGER,
    ty_le_lon_hon_6_thang_t_cong_1 REAL,
    ty_le_tong REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, nvkt, don_vi, ttvt)
);

CREATE TABLE IF NOT EXISTS kpi_nvkt_c11 (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL DEFAULT '',
    nvkt TEXT NOT NULL,
    sm1 INTEGER,
    sm2 INTEGER,
    ty_le_sua_chua_chat_luong_chu_dong REAL,
    sm3 INTEGER,
    sm4 INTEGER,
    ty_le_bao_hong_brcd_dung_quy_dinh REAL,
    chi_tieu_bsc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi, nvkt)
);

CREATE TABLE IF NOT EXISTS kpi_nvkt_c12 (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL DEFAULT '',
    nvkt TEXT NOT NULL,
    sm1 INTEGER,
    sm2 INTEGER,
    ty_le_hong_lap_lai REAL,
    sm3 INTEGER,
    sm4 INTEGER,
    ty_le_su_co_brcd REAL,
    chi_tieu_bsc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi, nvkt)
);

CREATE TABLE IF NOT EXISTS kpi_nvkt_c13 (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL DEFAULT '',
    nvkt TEXT NOT NULL,
    sm1 INTEGER,
    sm2 INTEGER,
    ty_le_sua_chua_dung_han REAL,
    sm3 INTEGER,
    sm4 INTEGER,
    ty_le_hong_lap_lai_kenh_tsl REAL,
    sm5 INTEGER,
    sm6 INTEGER,
    ty_le_su_co_kenh_tsl REAL,
    chi_tieu_bsc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi, nvkt)
);

CREATE TABLE IF NOT EXISTS ket_qua_tiep_thi_nv (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT,
    ma_nv TEXT NOT NULL DEFAULT '',
    ten_nv TEXT NOT NULL,
    dich_vu_brcd INTEGER,
    dich_vu_mytv INTEGER,
    tong INTEGER,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_nv, ten_nv)
);

CREATE TABLE IF NOT EXISTS ket_qua_tiep_thi_don_vi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    don_vi TEXT NOT NULL,
    dich_vu_brcd INTEGER,
    dich_vu_mytv INTEGER,
    tong INTEGER,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, don_vi)
);

CREATE TABLE IF NOT EXISTS hoan_cong_fiber (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    ma_thue_bao TEXT,
    so_may TEXT,
    hdtb_id TEXT,
    ma_khach_hang TEXT,
    ma_giao_dich TEXT,
    loai_hinh_thue_bao TEXT,
    ten_dich_vu TEXT,
    ten_goi TEXT,
    ten_thue_bao TEXT,
    dia_chi_thue_bao TEXT,
    dia_chi_lap_dat TEXT,
    ngay_yeu_cau TEXT,
    ngay_nghiem_thu TEXT,
    ttvt TEXT,
    doi_vien_thong TEXT,
    nvkt TEXT,
    trang_thai_hop_dong TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS ngung_psc_fiber (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    ma_thue_bao TEXT,
    ten_thue_bao TEXT,
    ngay_lap_hop_dong TEXT,
    ngay_thuc_hien TEXT,
    ten_kieu_lenh TEXT,
    ly_do_huy TEXT,
    trang_thai_thue_bao TEXT,
    ten_loai_hop_dong TEXT,
    trang_thai_hop_dong TEXT,
    ma_giao_dich TEXT,
    ttvt TEXT,
    doi_vien_thong TEXT,
    nvkt TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS khoi_phuc_fiber (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    ma_thue_bao TEXT,
    ten_thue_bao TEXT,
    ngay_lap_hop_dong TEXT,
    ngay_thuc_hien TEXT,
    ten_kieu_lenh TEXT,
    ly_do_huy TEXT,
    trang_thai_thue_bao TEXT,
    ten_loai_hop_dong TEXT,
    trang_thai_hop_dong TEXT,
    ma_giao_dich TEXT,
    ttvt TEXT,
    doi_vien_thong TEXT,
    nvkt TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS hoan_cong_mytv (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    ma_thue_bao TEXT,
    hdtb_id TEXT,
    ngay_ins TEXT,
    ngay_yeu_cau TEXT,
    nhom_dia_ban TEXT,
    doi_vien_thong TEXT,
    ten_ttvt TEXT,
    nhan_vien_ky_thuat TEXT,
    ma_giao_dich TEXT,
    trang_thai_hop_dong TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS ngung_psc_mytv (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    ma_thue_bao TEXT,
    ngay_tam_dung TEXT,
    ngay_huy TEXT,
    ma_khu_vuc TEXT,
    ten_khu_vuc TEXT,
    ten_doi TEXT,
    ten_ttvt TEXT,
    trang_thai_thue_bao TEXT,
    loai_hinh_thue_bao TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS thuc_tang_fiber (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    cap_tong_hop TEXT NOT NULL DEFAULT 'nvkt',
    ttvt TEXT NOT NULL DEFAULT '',
    doi_vien_thong TEXT NOT NULL DEFAULT '',
    nvkt TEXT NOT NULL DEFAULT '',
    hoan_cong INTEGER,
    ngung_phat_sinh_cuoc INTEGER,
    thuc_tang INTEGER,
    ty_le_ngung_psc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, cap_tong_hop, ttvt, doi_vien_thong, nvkt)
);

CREATE TABLE IF NOT EXISTS thuc_tang_mytv (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    cap_tong_hop TEXT NOT NULL DEFAULT 'nvkt',
    ttvt TEXT NOT NULL DEFAULT '',
    doi_vien_thong TEXT NOT NULL DEFAULT '',
    nvkt TEXT NOT NULL DEFAULT '',
    hoan_cong INTEGER,
    ngung_phat_sinh_cuoc INTEGER,
    thuc_tang INTEGER,
    ty_le_ngung_psc REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, cap_tong_hop, ttvt, doi_vien_thong, nvkt)
);

CREATE TABLE IF NOT EXISTS xac_minh_chi_tiet (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    ma_thue_bao TEXT,
    ten_loai_hinh_thue_bao TEXT,
    ten_thue_bao TEXT,
    ten_kieu_lenh TEXT,
    ngay_lap_hop_dong TEXT,
    ngay_hoan_thanh TEXT,
    loai_phieu TEXT,
    ten_khu_vuc TEXT,
    doi_vien_thong TEXT,
    ttvt TEXT,
    nvkt TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS xac_minh_tong_hop_nvkt (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ttvt TEXT NOT NULL DEFAULT '',
    doi_vien_thong TEXT NOT NULL DEFAULT '',
    nvkt TEXT NOT NULL,
    so_phieu_xac_minh INTEGER,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ttvt, doi_vien_thong, nvkt)
);

CREATE TABLE IF NOT EXISTS xac_minh_tong_hop_loai_phieu (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    loai_phieu TEXT NOT NULL,
    ten_kieu_lenh TEXT NOT NULL DEFAULT '',
    so_phieu_xac_minh INTEGER,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, loai_phieu, ten_kieu_lenh)
);

CREATE TABLE IF NOT EXISTS cau_hinh_tu_dong_chi_tiet (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    serial_number TEXT,
    ma_thue_bao TEXT,
    loai_hop_dong TEXT,
    loai_cau_hinh TEXT,
    trang_thai TEXT,
    trang_thai_chuan_hoa TEXT,
    ma_loi TEXT,
    thoi_gian_cap_nhat TEXT,
    ttvt TEXT,
    doi_vien_thong TEXT,
    ma_nhan_vien TEXT,
    nvkt TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS cau_hinh_tu_dong_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ttvt TEXT NOT NULL DEFAULT '',
    don_vi TEXT NOT NULL,
    loai_dong TEXT NOT NULL,
    tong_hop_dong INTEGER,
    khong_thuc_hien_cau_hinh_tu_dong INTEGER,
    da_day_cau_hinh_tu_dong INTEGER,
    khong_day_do_loi_he_thong INTEGER,
    khong_day_do_tbi_da_co_cau_hinh INTEGER,
    cau_hinh_thanh_cong INTEGER,
    ty_le_day_tu_dong REAL,
    ty_le_tbi_da_co_cau_hinh REAL,
    ty_le_cau_hinh_thanh_cong REAL,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ttvt, don_vi, loai_dong)
);

CREATE TABLE IF NOT EXISTS tong_hop_loi_cau_hinh_tu_dong (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_loi TEXT NOT NULL,
    so_luong INTEGER,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_loi)
);

CREATE TABLE IF NOT EXISTS vat_tu_thu_hoi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    nvkt_dia_ban_giao TEXT,
    trang_thai_thu_hoi TEXT,
    loai_vat_tu TEXT,
    loai_phieu TEXT,
    ma_men TEXT,
    ma_thue_bao TEXT,
    ten_thue_bao TEXT,
    dia_chi_khach_hang TEXT,
    nhan_vien_khoa TEXT,
    nhan_vien_thu TEXT,
    nhan_vien_nhap_kho TEXT,
    ngay_khoa TEXT,
    ngay_hoan_cong TEXT,
    ngay_hoan_ung TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS chi_tiet_vat_tu_thu_hoi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    diem_chia TEXT,
    nvkt_dia_ban_giao TEXT,
    ma_thue_bao TEXT,
    ten_thue_bao TEXT,
    ten_thiet_bi TEXT,
    ngay_giao TEXT,
    ten_loai_hop_dong TEXT,
    ten_kieu_lenh TEXT,
    so_dien_thoai TEXT,
    ngay_su_dung_thiet_bi TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE TABLE IF NOT EXISTS quyet_toan_vat_tu (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bao_cao_ngay_id INTEGER NOT NULL,
    ma_hash_dong TEXT,
    ma_tkyt TEXT,
    ma_spdv TEXT,
    loai TEXT,
    ten_vat_tu TEXT,
    don_vi_tinh TEXT,
    so_luong REAL,
    don_gia REAL,
    thanh_tien REAL,
    ma_vat_tu TEXT,
    du_lieu_bo_sung_json TEXT,
    FOREIGN KEY (bao_cao_ngay_id) REFERENCES bao_cao_ngay(id) ON DELETE CASCADE,
    UNIQUE (bao_cao_ngay_id, ma_hash_dong)
);

CREATE INDEX IF NOT EXISTS idx_c11_tong_hop_bao_cao_ngay_id ON c11_tong_hop(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_c11_chi_tiet_nvkt_bao_cao_ngay_id ON c11_chi_tiet_nvkt(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_c12_tong_hop_bao_cao_ngay_id ON c12_tong_hop(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_c12_hong_lap_lai_nvkt_bao_cao_ngay_id ON c12_hong_lap_lai_nvkt(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_c13_tong_hop_bao_cao_ngay_id ON c13_tong_hop(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_c14_tong_hop_bao_cao_ngay_id ON c14_tong_hop(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_c14_hai_long_nvkt_bao_cao_ngay_id ON c14_hai_long_nvkt(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_ghtt_don_vi_bao_cao_ngay_id ON ghtt_don_vi(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_ghtt_nvkt_bao_cao_ngay_id ON ghtt_nvkt(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_kpi_nvkt_c11_bao_cao_ngay_id ON kpi_nvkt_c11(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_kpi_nvkt_c12_bao_cao_ngay_id ON kpi_nvkt_c12(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_kpi_nvkt_c13_bao_cao_ngay_id ON kpi_nvkt_c13(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_ket_qua_tiep_thi_nv_bao_cao_ngay_id ON ket_qua_tiep_thi_nv(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_ket_qua_tiep_thi_don_vi_bao_cao_ngay_id ON ket_qua_tiep_thi_don_vi(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_hoan_cong_fiber_bao_cao_ngay_id ON hoan_cong_fiber(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_ngung_psc_fiber_bao_cao_ngay_id ON ngung_psc_fiber(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_khoi_phuc_fiber_bao_cao_ngay_id ON khoi_phuc_fiber(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_hoan_cong_mytv_bao_cao_ngay_id ON hoan_cong_mytv(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_ngung_psc_mytv_bao_cao_ngay_id ON ngung_psc_mytv(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_thuc_tang_fiber_bao_cao_ngay_id ON thuc_tang_fiber(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_thuc_tang_mytv_bao_cao_ngay_id ON thuc_tang_mytv(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_xac_minh_chi_tiet_bao_cao_ngay_id ON xac_minh_chi_tiet(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_xac_minh_tong_hop_nvkt_bao_cao_ngay_id ON xac_minh_tong_hop_nvkt(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_xac_minh_tong_hop_loai_phieu_bao_cao_ngay_id ON xac_minh_tong_hop_loai_phieu(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_cau_hinh_tu_dong_chi_tiet_bao_cao_ngay_id ON cau_hinh_tu_dong_chi_tiet(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_cau_hinh_tu_dong_tong_hop_bao_cao_ngay_id ON cau_hinh_tu_dong_tong_hop(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_tong_hop_loi_cau_hinh_tu_dong_bao_cao_ngay_id ON tong_hop_loi_cau_hinh_tu_dong(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_vat_tu_thu_hoi_bao_cao_ngay_id ON vat_tu_thu_hoi(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_chi_tiet_vat_tu_thu_hoi_bao_cao_ngay_id ON chi_tiet_vat_tu_thu_hoi(bao_cao_ngay_id);
CREATE INDEX IF NOT EXISTS idx_quyet_toan_vat_tu_bao_cao_ngay_id ON quyet_toan_vat_tu(bao_cao_ngay_id);

COMMIT;
