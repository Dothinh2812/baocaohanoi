-- Database Schema cho hệ thống báo cáo Hà Nội
-- Tạo: 2025-11-20

-- ============================================
-- Bảng danh mục
-- ============================================

-- Bảng Nhân viên
CREATE TABLE IF NOT EXISTS nhan_vien (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_nv TEXT UNIQUE,
    ten_nv TEXT NOT NULL,
    don_vi TEXT,
    nhom_dia_ban TEXT,
    ttvt TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Bảng Đơn vị
CREATE TABLE IF NOT EXISTS don_vi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_don_vi TEXT UNIQUE,
    ten_don_vi TEXT NOT NULL,
    ttvt TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- ============================================
-- Bảng dữ liệu chi tiết hàng ngày
-- ============================================

-- Bảng Hoàn công (FIBER)
CREATE TABLE IF NOT EXISTS hoan_cong (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    loai_dv TEXT DEFAULT 'FIBER', -- FIBER hoặc MYTV
    stt INTEGER,
    ma_tb TEXT,
    ngay_nghiem_thu DATETIME,
    ngay_yeu_cau DATETIME,
    doi TEXT,
    nhom_dia_ban TEXT,
    ten_ttvt TEXT,
    trang_thai_phieu TEXT,
    hdtb_id INTEGER,
    nhan_vien_kt TEXT,
    ma_gd TEXT,
    nvkt TEXT,
    don_vi TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    -- Index để truy vấn nhanh
    UNIQUE(ma_tb, hdtb_id, ngay_bao_cao)
);

CREATE INDEX idx_hoan_cong_ngay ON hoan_cong(ngay_bao_cao);
CREATE INDEX idx_hoan_cong_nvkt ON hoan_cong(nvkt);
CREATE INDEX idx_hoan_cong_don_vi ON hoan_cong(don_vi);

-- Bảng Ngừng PSC
CREATE TABLE IF NOT EXISTS ngung_psc (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    loai_dv TEXT DEFAULT 'FIBER', -- FIBER hoặc MYTV
    stt INTEGER,
    ma_tb TEXT,
    so_may TEXT,
    ten_tb TEXT,
    loai_dich_vu TEXT,
    dia_chi_ld TEXT,
    ngay_tam_dung DATETIME,
    ngay_khoi_phuc DATETIME,
    ngay_huy DATETIME,
    nhom_dia_ban TEXT,
    ten_to TEXT,
    ten_ttvt TEXT,
    ten_kh TEXT,
    dien_thoai_lh TEXT,
    trang_thai_tb TEXT,
    ly_do_huy_tam_dung TEXT,
    ttvt_xac_minh_huy TEXT,
    doi_tuong TEXT,
    nvkt TEXT,
    don_vi TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ma_tb, ngay_tam_dung, ngay_bao_cao)
);

CREATE INDEX idx_ngung_psc_ngay ON ngung_psc(ngay_bao_cao);
CREATE INDEX idx_ngung_psc_nvkt ON ngung_psc(nvkt);
CREATE INDEX idx_ngung_psc_don_vi ON ngung_psc(don_vi);

-- Bảng Thực tăng tổng hợp
CREATE TABLE IF NOT EXISTS thuc_tang (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    loai_dv TEXT DEFAULT 'FIBER', -- FIBER hoặc MYTV
    cap_do TEXT NOT NULL, -- 'to' hoặc 'nvkt'
    don_vi TEXT,
    nvkt TEXT,
    hoan_cong INTEGER DEFAULT 0,
    ngung_psc INTEGER DEFAULT 0,
    thuc_tang INTEGER DEFAULT 0,
    ty_le_ngung_psc REAL DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ngay_bao_cao, loai_dv, cap_do, don_vi, nvkt)
);

CREATE INDEX idx_thuc_tang_ngay ON thuc_tang(ngay_bao_cao);
CREATE INDEX idx_thuc_tang_nvkt ON thuc_tang(nvkt);

-- ============================================
-- Bảng Suy hao cao (I1.5)
-- ============================================

CREATE TABLE IF NOT EXISTS suy_hao_cao (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    ttvt_cts TEXT,
    ttvt_one TEXT,
    doi_one TEXT,
    ten_kv TEXT,
    olt_cts TEXT,
    port_cts TEXT,
    account_cts TEXT,
    ten_tb_one TEXT,
    dt_one TEXT,
    diachi_one TEXT,
    ngay_suyhao DATE,
    trangthai_tb TEXT,
    ma_module_quang_olt TEXT,
    chi_so_olt_rx REAL,
    chi_so_onu_rx REAL,
    ma_tu_hop_s2 TEXT,
    gia_tri_sh_s2_one REAL,
    s2_vi_tri_one_port TEXT,
    chenh_lech REAL,
    do_dai_day_thue_bao REAL,
    down REAL,
    up REAL,
    do_dai REAL,
    chi_so_kem_8362 TEXT,
    chi_so_kem_vthn TEXT,
    mo_ta TEXT,
    so_sanh_phieu_ton TEXT,
    ma_tu_hop TEXT,
    dia_chi_lap_dat TEXT,
    loai_kh TEXT,
    nvkt_db TEXT,
    thietbi TEXT,
    sa TEXT,
    ketcuoi TEXT,
    nvkt_db_normalized TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(account_cts, ngay_bao_cao)
);

CREATE INDEX idx_suy_hao_cao_ngay ON suy_hao_cao(ngay_bao_cao);
CREATE INDEX idx_suy_hao_cao_nvkt ON suy_hao_cao(nvkt_db_normalized);
CREATE INDEX idx_suy_hao_cao_don_vi ON suy_hao_cao(doi_one);

-- ============================================
-- Bảng Báo cáo tuần/tháng
-- ============================================

CREATE TABLE IF NOT EXISTS bao_cao_tuan_thang (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ky_bao_cao TEXT NOT NULL, -- 'tuan' hoặc 'thang'
    so_ky INTEGER NOT NULL, -- Số tuần/tháng
    nam INTEGER NOT NULL,
    ngay_bat_dau DATE,
    ngay_ket_thuc DATE,
    don_vi TEXT,
    nvkt_db TEXT,
    so_tb_ky_nay INTEGER DEFAULT 0,
    so_tb_ky_truoc INTEGER DEFAULT 0,
    tang_giam INTEGER DEFAULT 0,
    phan_tram_thay_doi REAL DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ky_bao_cao, so_ky, nam, don_vi, nvkt_db)
);

CREATE INDEX idx_bao_cao_tuan_thang_ky ON bao_cao_tuan_thang(ky_bao_cao, so_ky, nam);

-- Bảng Chi tiết TB trong báo cáo tuần/tháng
CREATE TABLE IF NOT EXISTS bao_cao_tuan_thang_chitiet (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    id_bao_cao INTEGER,
    account_cts TEXT,
    ten_tb_one TEXT,
    doi_one TEXT,
    nvkt_db_normalized TEXT,
    sa TEXT,
    trang_thai TEXT, -- 'tang_moi', 'giam_het', 'van_con'
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    FOREIGN KEY (id_bao_cao) REFERENCES bao_cao_tuan_thang(id) ON DELETE CASCADE
);

CREATE INDEX idx_bao_cao_chitiet_id ON bao_cao_tuan_thang_chitiet(id_bao_cao);

-- ============================================
-- Bảng Xu hướng theo ngày
-- ============================================

CREATE TABLE IF NOT EXISTS xu_huong_theo_ngay (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay DATE NOT NULL,
    cap_do TEXT NOT NULL, -- 'tong', 'don_vi', 'nvkt'
    don_vi TEXT,
    nvkt TEXT,
    so_luong_tb_suy_hao INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ngay, cap_do, don_vi, nvkt)
);

CREATE INDEX idx_xu_huong_ngay ON xu_huong_theo_ngay(ngay);
CREATE INDEX idx_xu_huong_don_vi ON xu_huong_theo_ngay(don_vi);

-- ============================================
-- Bảng Báo cáo C1.x (Chất lượng dịch vụ)
-- ============================================

CREATE TABLE IF NOT EXISTS bao_cao_c1 (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    thang INTEGER NOT NULL,
    nam INTEGER NOT NULL,
    loai_bao_cao TEXT NOT NULL, -- 'C1.1', 'C1.2', 'C1.3', 'C1.4', 'C1.5'
    don_vi TEXT,
    phan_loai TEXT, -- 'tong', 'ccco', 'khong_ccco'
    sm1 INTEGER,
    sm2 INTEGER,
    sm3 INTEGER,
    sm4 INTEGER,
    ty_le_1 REAL, -- Tỷ lệ chính của từng loại báo cáo
    ty_le_2 REAL, -- Tỷ lệ phụ (nếu có)
    diem_bsc REAL,
    mo_ta_ty_le_1 TEXT, -- Mô tả ý nghĩa của tỷ lệ 1
    mo_ta_ty_le_2 TEXT, -- Mô tả ý nghĩa của tỷ lệ 2
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ngay_bao_cao, loai_bao_cao, don_vi, phan_loai)
);

CREATE INDEX idx_bao_cao_c1_ngay ON bao_cao_c1(ngay_bao_cao);
CREATE INDEX idx_bao_cao_c1_loai ON bao_cao_c1(loai_bao_cao);

-- Bảng Chi tiết C1.4 và C1.5
CREATE TABLE IF NOT EXISTS bao_cao_c1_chitiet (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    id_bao_cao INTEGER,
    ma_tb TEXT,
    ten_tb TEXT,
    dia_chi TEXT,
    nvkt TEXT,
    ly_do TEXT,
    thoi_gian TEXT,
    ghi_chu TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    FOREIGN KEY (id_bao_cao) REFERENCES bao_cao_c1(id) ON DELETE CASCADE
);

CREATE INDEX idx_bao_cao_c1_chitiet_id ON bao_cao_c1_chitiet(id_bao_cao);

-- ============================================
-- Bảng Báo cáo KR6/KR7
-- ============================================

CREATE TABLE IF NOT EXISTS bao_cao_kr (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    loai_bao_cao TEXT NOT NULL, -- 'KR6', 'KR7'
    cap_do TEXT NOT NULL, -- 'nvkt', 'tong_hop'
    don_vi TEXT,
    nvkt TEXT,
    chi_tieu_1 INTEGER,
    chi_tieu_2 INTEGER,
    chi_tieu_3 INTEGER,
    ty_le REAL,
    diem REAL,
    ghi_chu TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ngay_bao_cao, loai_bao_cao, cap_do, don_vi, nvkt)
);

CREATE INDEX idx_bao_cao_kr_ngay ON bao_cao_kr(ngay_bao_cao);
CREATE INDEX idx_bao_cao_kr_loai ON bao_cao_kr(loai_bao_cao);

-- ============================================
-- Bảng Báo cáo thu hồi vật tư
-- ============================================

CREATE TABLE IF NOT EXISTS thu_hoi_vat_tu (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    stt INTEGER,
    ma_tb TEXT,
    ten_tb TEXT,
    dia_chi TEXT,
    loai_thiet_bi TEXT,
    so_luong INTEGER,
    don_vi_tinh TEXT,
    nvkt TEXT,
    don_vi TEXT,
    ngay_thu_hoi DATE,
    trang_thai TEXT,
    ghi_chu TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_thu_hoi_vat_tu_ngay ON thu_hoi_vat_tu(ngay_bao_cao);
CREATE INDEX idx_thu_hoi_vat_tu_nvkt ON thu_hoi_vat_tu(nvkt);

-- ============================================
-- Bảng SM (Chỉ tiêu KPI)
-- ============================================

CREATE TABLE IF NOT EXISTS bao_cao_sm (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    thang INTEGER NOT NULL,
    nam INTEGER NOT NULL,
    loai_sm TEXT NOT NULL, -- 'SM1', 'SM2', 'SM3', 'SM4'
    don_vi TEXT,
    nvkt TEXT,
    gia_tri INTEGER,
    muc_tieu INTEGER,
    ty_le_hoan_thanh REAL,
    xep_loai TEXT,
    ghi_chu TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ngay_bao_cao, loai_sm, don_vi, nvkt)
);

CREATE INDEX idx_bao_cao_sm_ngay ON bao_cao_sm(ngay_bao_cao);
CREATE INDEX idx_bao_cao_sm_loai ON bao_cao_sm(loai_sm);

-- ============================================
-- Bảng Suy hao theo SA (thiết bị)
-- ============================================

CREATE TABLE IF NOT EXISTS suy_hao_theo_sa (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    sa TEXT NOT NULL,
    so_luong INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ngay_bao_cao, sa)
);

CREATE INDEX idx_suy_hao_theo_sa_ngay ON suy_hao_theo_sa(ngay_bao_cao);

-- ============================================
-- Bảng Biến động suy hao cao
-- ============================================

CREATE TABLE IF NOT EXISTS bien_dong_suy_hao (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay_bao_cao DATE NOT NULL,
    don_vi TEXT,
    nvkt_db TEXT,
    tong_so_hien_tai INTEGER DEFAULT 0,
    tang_moi INTEGER DEFAULT 0,
    giam_het INTEGER DEFAULT 0,
    van_con INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,

    UNIQUE(ngay_bao_cao, don_vi, nvkt_db)
);

CREATE INDEX idx_bien_dong_suy_hao_ngay ON bien_dong_suy_hao(ngay_bao_cao);

-- ============================================
-- Bảng log import
-- ============================================

CREATE TABLE IF NOT EXISTS import_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name TEXT NOT NULL,
    file_path TEXT,
    loai_bao_cao TEXT,
    ngay_bao_cao DATE,
    so_ban_ghi INTEGER DEFAULT 0,
    trang_thai TEXT, -- 'thanh_cong', 'loi', 'canh_bao'
    thong_bao TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_import_log_file ON import_log(file_name);
CREATE INDEX idx_import_log_ngay ON import_log(created_at);

-- ============================================
-- View để truy vấn nhanh
-- ============================================

-- View tổng hợp theo ngày
CREATE VIEW IF NOT EXISTS v_tong_hop_ngay AS
SELECT
    h.ngay_bao_cao,
    h.loai_dv,
    h.don_vi,
    COUNT(h.id) as so_hoan_cong,
    COUNT(n.id) as so_ngung_psc,
    COUNT(h.id) - COUNT(n.id) as thuc_tang
FROM hoan_cong h
LEFT JOIN ngung_psc n ON h.ngay_bao_cao = n.ngay_bao_cao
    AND h.don_vi = n.don_vi
    AND h.loai_dv = n.loai_dv
GROUP BY h.ngay_bao_cao, h.loai_dv, h.don_vi;

-- View tổng hợp theo NVKT
CREATE VIEW IF NOT EXISTS v_tong_hop_nvkt AS
SELECT
    h.ngay_bao_cao,
    h.loai_dv,
    h.don_vi,
    h.nvkt,
    COUNT(h.id) as so_hoan_cong,
    COUNT(n.id) as so_ngung_psc,
    COUNT(h.id) - COUNT(n.id) as thuc_tang
FROM hoan_cong h
LEFT JOIN ngung_psc n ON h.ngay_bao_cao = n.ngay_bao_cao
    AND h.nvkt = n.nvkt
    AND h.loai_dv = n.loai_dv
GROUP BY h.ngay_bao_cao, h.loai_dv, h.don_vi, h.nvkt;

-- View xu hướng suy hao cao
CREATE VIEW IF NOT EXISTS v_xu_huong_suy_hao AS
SELECT
    ngay_bao_cao,
    doi_one as don_vi,
    COUNT(*) as so_luong_tb_suy_hao
FROM suy_hao_cao
GROUP BY ngay_bao_cao, doi_one
ORDER BY ngay_bao_cao DESC, so_luong_tb_suy_hao DESC;
