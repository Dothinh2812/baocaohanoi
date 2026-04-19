-- View layer cho api_transition/report_history.db
-- Gom 3 lop:
-- 1. view quan tri / helper
-- 2. view lich su va moi nhat cho tung nhom nghiep vu
-- 3. view tong hop cho dashboard

BEGIN;

DROP VIEW IF EXISTS v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao;
CREATE VIEW v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao AS
SELECT b.*
FROM bao_cao_ngay b
JOIN (
    SELECT ma_bao_cao, MAX(ngay_du_lieu) AS ngay_du_lieu
    FROM bao_cao_ngay
    WHERE trang_thai_nap = 'thanh_cong'
    GROUP BY ma_bao_cao
) mx
  ON b.ma_bao_cao = mx.ma_bao_cao
 AND b.ngay_du_lieu = mx.ngay_du_lieu;

DROP VIEW IF EXISTS v_bao_cao_ngay_moi_nhat_toan_bo;
CREATE VIEW v_bao_cao_ngay_moi_nhat_toan_bo AS
SELECT *
FROM bao_cao_ngay
WHERE ngay_du_lieu = (
    SELECT MAX(ngay_du_lieu) FROM bao_cao_ngay WHERE trang_thai_nap = 'thanh_cong'
);

DROP VIEW IF EXISTS v_nhat_ky_nap_gan_nhat;
CREATE VIEW v_nhat_ky_nap_gan_nhat AS
SELECT n.*
FROM nhat_ky_nap_bao_cao n
JOIN (
    SELECT ma_bao_cao, MAX(id) AS id_moi_nhat
    FROM nhat_ky_nap_bao_cao
    GROUP BY ma_bao_cao
) mx
  ON n.ma_bao_cao = mx.ma_bao_cao
 AND n.id = mx.id_moi_nhat;

DROP VIEW IF EXISTS v_tien_do_nap_bao_cao;
CREATE VIEW v_tien_do_nap_bao_cao AS
SELECT
    d.ma_bao_cao,
    d.ten_bao_cao,
    d.nhom_bao_cao,
    d.duong_dan_processed_mac_dinh,
    b.ngay_du_lieu,
    b.trang_thai_nap,
    b.so_dong_goc,
    b.so_dong_tong_hop,
    b.so_dong_chi_tiet,
    b.ghi_chu,
    n.bat_dau_luc,
    n.ket_thuc_luc,
    n.trang_thai AS trang_thai_nhat_ky,
    n.thong_diep
FROM danh_muc_bao_cao d
LEFT JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao b
  ON d.ma_bao_cao = b.ma_bao_cao
LEFT JOIN v_nhat_ky_nap_gan_nhat n
  ON d.ma_bao_cao = n.ma_bao_cao;

DROP VIEW IF EXISTS v_sheet_bao_cao_moi_nhat;
CREATE VIEW v_sheet_bao_cao_moi_nhat AS
SELECT
    b.ma_bao_cao,
    b.ngay_du_lieu,
    s.*
FROM sheet_bao_cao s
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao b
  ON s.bao_cao_ngay_id = b.id;

DROP VIEW IF EXISTS v_dong_bao_cao_goc_moi_nhat;
CREATE VIEW v_dong_bao_cao_goc_moi_nhat AS
SELECT
    b.ma_bao_cao,
    b.ngay_du_lieu,
    r.*
FROM dong_bao_cao_goc r
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao b
  ON r.bao_cao_ngay_id = b.id;

DROP VIEW IF EXISTS v_tep_luu_tru_moi_nhat;
CREATE VIEW v_tep_luu_tru_moi_nhat AS
SELECT
    b.ma_bao_cao,
    b.ngay_du_lieu,
    t.*
FROM tep_luu_tru_bao_cao t
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao b
  ON t.bao_cao_ngay_id = b.id;

DROP VIEW IF EXISTS v_c11_tong_hop_lich_su;
CREATE VIEW v_c11_tong_hop_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM c11_tong_hop t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_c11_tong_hop_moi_nhat;
CREATE VIEW v_c11_tong_hop_moi_nhat AS
SELECT v.*
FROM v_c11_tong_hop_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'chi_tieu_c_c1_1_report';

DROP VIEW IF EXISTS v_c11_nvkt_lich_su;
CREATE VIEW v_c11_nvkt_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM c11_chi_tiet_nvkt t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_c11_nvkt_moi_nhat;
CREATE VIEW v_c11_nvkt_moi_nhat AS
SELECT v.*
FROM v_c11_nvkt_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'chi_tieu_c_c1_1_chitiet_report';

DROP VIEW IF EXISTS v_c12_tong_hop_lich_su;
CREATE VIEW v_c12_tong_hop_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM c12_tong_hop t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_c12_tong_hop_moi_nhat;
CREATE VIEW v_c12_tong_hop_moi_nhat AS
SELECT v.*
FROM v_c12_tong_hop_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'chi_tieu_c_c1_2_report';

DROP VIEW IF EXISTS v_c12_nvkt_lich_su;
CREATE VIEW v_c12_nvkt_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM c12_hong_lap_lai_nvkt t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_c12_nvkt_moi_nhat;
CREATE VIEW v_c12_nvkt_moi_nhat AS
SELECT v.*
FROM v_c12_nvkt_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'chi_tieu_c_c1_2_chitiet_sm1_report';

DROP VIEW IF EXISTS v_c13_tong_hop_lich_su;
CREATE VIEW v_c13_tong_hop_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM c13_tong_hop t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_c13_tong_hop_moi_nhat;
CREATE VIEW v_c13_tong_hop_moi_nhat AS
SELECT v.*
FROM v_c13_tong_hop_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'chi_tieu_c_c1_3_report';

DROP VIEW IF EXISTS v_c14_tong_hop_lich_su;
CREATE VIEW v_c14_tong_hop_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM c14_tong_hop t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_c14_tong_hop_moi_nhat;
CREATE VIEW v_c14_tong_hop_moi_nhat AS
SELECT v.*
FROM v_c14_tong_hop_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'chi_tieu_c_c1_4_report';

DROP VIEW IF EXISTS v_c14_nvkt_lich_su;
CREATE VIEW v_c14_nvkt_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM c14_hai_long_nvkt t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_c14_nvkt_moi_nhat;
CREATE VIEW v_c14_nvkt_moi_nhat AS
SELECT v.*
FROM v_c14_nvkt_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'chi_tieu_c_c1_4_chitiet_report';

DROP VIEW IF EXISTS v_ghtt_don_vi_lich_su;
CREATE VIEW v_ghtt_don_vi_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM ghtt_don_vi t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_ghtt_don_vi_moi_nhat;
CREATE VIEW v_ghtt_don_vi_moi_nhat AS
SELECT v.*
FROM v_ghtt_don_vi_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m
  ON v.ma_bao_cao = m.ma_bao_cao
 AND v.ngay_du_lieu = m.ngay_du_lieu;

DROP VIEW IF EXISTS v_ghtt_nvkt_lich_su;
CREATE VIEW v_ghtt_nvkt_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM ghtt_nvkt t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_ghtt_nvkt_moi_nhat;
CREATE VIEW v_ghtt_nvkt_moi_nhat AS
SELECT v.*
FROM v_ghtt_nvkt_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'ghtt_ghtt_nvktdb_report';

DROP VIEW IF EXISTS v_kpi_nvkt_tong_hop_lich_su;
CREATE VIEW v_kpi_nvkt_tong_hop_lich_su AS
SELECT
    'c11' AS nhom_chi_tieu,
    b.ma_bao_cao,
    b.ngay_du_lieu,
    t.don_vi,
    t.nvkt,
    t.sm1,
    t.sm2,
    t.sm3,
    t.sm4,
    NULL AS sm5,
    NULL AS sm6,
    t.ty_le_sua_chua_chat_luong_chu_dong AS chi_so_1,
    'ty_le_sua_chua_chat_luong_chu_dong' AS ten_chi_so_1,
    t.ty_le_bao_hong_brcd_dung_quy_dinh AS chi_so_2,
    'ty_le_bao_hong_brcd_dung_quy_dinh' AS ten_chi_so_2,
    NULL AS chi_so_3,
    NULL AS ten_chi_so_3,
    t.chi_tieu_bsc,
    t.du_lieu_bo_sung_json
FROM kpi_nvkt_c11 t JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id
UNION ALL
SELECT
    'c12',
    b.ma_bao_cao,
    b.ngay_du_lieu,
    t.don_vi,
    t.nvkt,
    t.sm1,
    t.sm2,
    t.sm3,
    t.sm4,
    NULL,
    NULL,
    t.ty_le_hong_lap_lai,
    'ty_le_hong_lap_lai',
    t.ty_le_su_co_brcd,
    'ty_le_su_co_brcd',
    NULL,
    NULL,
    t.chi_tieu_bsc,
    t.du_lieu_bo_sung_json
FROM kpi_nvkt_c12 t JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id
UNION ALL
SELECT
    'c13',
    b.ma_bao_cao,
    b.ngay_du_lieu,
    t.don_vi,
    t.nvkt,
    t.sm1,
    t.sm2,
    t.sm3,
    t.sm4,
    t.sm5,
    t.sm6,
    t.ty_le_sua_chua_dung_han,
    'ty_le_sua_chua_dung_han',
    t.ty_le_hong_lap_lai_kenh_tsl,
    'ty_le_hong_lap_lai_kenh_tsl',
    t.ty_le_su_co_kenh_tsl,
    'ty_le_su_co_kenh_tsl',
    t.chi_tieu_bsc,
    t.du_lieu_bo_sung_json
FROM kpi_nvkt_c13 t JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_kpi_nvkt_tong_hop_moi_nhat;
CREATE VIEW v_kpi_nvkt_tong_hop_moi_nhat AS
SELECT v.*
FROM v_kpi_nvkt_tong_hop_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.ma_bao_cao = m.ma_bao_cao AND v.ngay_du_lieu = m.ngay_du_lieu;

DROP VIEW IF EXISTS v_ket_qua_tiep_thi_nv_lich_su;
CREATE VIEW v_ket_qua_tiep_thi_nv_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM ket_qua_tiep_thi_nv t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_ket_qua_tiep_thi_nv_moi_nhat;
CREATE VIEW v_ket_qua_tiep_thi_nv_moi_nhat AS
SELECT v.*
FROM v_ket_qua_tiep_thi_nv_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'kq_tiep_thi_kq_tiep_thi_report';

DROP VIEW IF EXISTS v_ket_qua_tiep_thi_don_vi_lich_su;
CREATE VIEW v_ket_qua_tiep_thi_don_vi_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM ket_qua_tiep_thi_don_vi t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_ket_qua_tiep_thi_don_vi_moi_nhat;
CREATE VIEW v_ket_qua_tiep_thi_don_vi_moi_nhat AS
SELECT v.*
FROM v_ket_qua_tiep_thi_don_vi_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'kq_tiep_thi_kq_tiep_thi_report';

DROP VIEW IF EXISTS v_hoan_cong_fiber_lich_su;
CREATE VIEW v_hoan_cong_fiber_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM hoan_cong_fiber t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_hoan_cong_fiber_moi_nhat;
CREATE VIEW v_hoan_cong_fiber_moi_nhat AS
SELECT v.*
FROM v_hoan_cong_fiber_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'phieu_hoan_cong_dich_vu_phieu_hoan_cong_dich_vu_chi_tiet';

DROP VIEW IF EXISTS v_ngung_psc_fiber_lich_su;
CREATE VIEW v_ngung_psc_fiber_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM ngung_psc_fiber t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_ngung_psc_fiber_moi_nhat;
CREATE VIEW v_ngung_psc_fiber_moi_nhat AS
SELECT v.*
FROM v_ngung_psc_fiber_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet';

DROP VIEW IF EXISTS v_khoi_phuc_fiber_lich_su;
CREATE VIEW v_khoi_phuc_fiber_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM khoi_phuc_fiber t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_khoi_phuc_fiber_moi_nhat;
CREATE VIEW v_khoi_phuc_fiber_moi_nhat AS
SELECT v.*
FROM v_khoi_phuc_fiber_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc';

DROP VIEW IF EXISTS v_hoan_cong_mytv_lich_su;
CREATE VIEW v_hoan_cong_mytv_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM hoan_cong_mytv t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_hoan_cong_mytv_moi_nhat;
CREATE VIEW v_hoan_cong_mytv_moi_nhat AS
SELECT v.*
FROM v_hoan_cong_mytv_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'mytv_dich_vu_mytv_hoan_cong';

DROP VIEW IF EXISTS v_ngung_psc_mytv_lich_su;
CREATE VIEW v_ngung_psc_mytv_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM ngung_psc_mytv t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_ngung_psc_mytv_moi_nhat;
CREATE VIEW v_ngung_psc_mytv_moi_nhat AS
SELECT v.*
FROM v_ngung_psc_mytv_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'mytv_dich_vu_mytv_ngung_psc';

DROP VIEW IF EXISTS v_thuc_tang_fiber_lich_su;
CREATE VIEW v_thuc_tang_fiber_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM thuc_tang_fiber t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_thuc_tang_fiber_moi_nhat;
CREATE VIEW v_thuc_tang_fiber_moi_nhat AS
SELECT v.*
FROM v_thuc_tang_fiber_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'thuc_tang_ngung_psc_fiber_thuc_tang';

DROP VIEW IF EXISTS v_thuc_tang_mytv_lich_su;
CREATE VIEW v_thuc_tang_mytv_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM thuc_tang_mytv t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_thuc_tang_mytv_moi_nhat;
CREATE VIEW v_thuc_tang_mytv_moi_nhat AS
SELECT v.*
FROM v_thuc_tang_mytv_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'mytv_dich_vu_mytv_thuc_tang';

DROP VIEW IF EXISTS v_xac_minh_chi_tiet_lich_su;
CREATE VIEW v_xac_minh_chi_tiet_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM xac_minh_chi_tiet t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_xac_minh_chi_tiet_moi_nhat;
CREATE VIEW v_xac_minh_chi_tiet_moi_nhat AS
SELECT v.*
FROM v_xac_minh_chi_tiet_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'ty_le_xac_minh_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet';

DROP VIEW IF EXISTS v_xac_minh_nvkt_lich_su;
CREATE VIEW v_xac_minh_nvkt_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM xac_minh_tong_hop_nvkt t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_xac_minh_nvkt_moi_nhat;
CREATE VIEW v_xac_minh_nvkt_moi_nhat AS
SELECT v.*
FROM v_xac_minh_nvkt_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'ty_le_xac_minh_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet';

DROP VIEW IF EXISTS v_xac_minh_loai_phieu_lich_su;
CREATE VIEW v_xac_minh_loai_phieu_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM xac_minh_tong_hop_loai_phieu t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_xac_minh_loai_phieu_moi_nhat;
CREATE VIEW v_xac_minh_loai_phieu_moi_nhat AS
SELECT v.*
FROM v_xac_minh_loai_phieu_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'ty_le_xac_minh_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet';

DROP VIEW IF EXISTS v_cau_hinh_tu_dong_chi_tiet_lich_su;
CREATE VIEW v_cau_hinh_tu_dong_chi_tiet_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM cau_hinh_tu_dong_chi_tiet t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_cau_hinh_tu_dong_chi_tiet_moi_nhat;
CREATE VIEW v_cau_hinh_tu_dong_chi_tiet_moi_nhat AS
SELECT *
FROM v_cau_hinh_tu_dong_chi_tiet_lich_su
WHERE ngay_du_lieu = (SELECT MAX(ngay_du_lieu) FROM bao_cao_ngay WHERE trang_thai_nap = 'thanh_cong');

DROP VIEW IF EXISTS v_cau_hinh_tu_dong_tong_hop_lich_su;
CREATE VIEW v_cau_hinh_tu_dong_tong_hop_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM cau_hinh_tu_dong_tong_hop t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_cau_hinh_tu_dong_tong_hop_moi_nhat;
CREATE VIEW v_cau_hinh_tu_dong_tong_hop_moi_nhat AS
SELECT *
FROM v_cau_hinh_tu_dong_tong_hop_lich_su
WHERE ngay_du_lieu = (SELECT MAX(ngay_du_lieu) FROM bao_cao_ngay WHERE trang_thai_nap = 'thanh_cong');

DROP VIEW IF EXISTS v_tong_hop_loi_cau_hinh_tu_dong_lich_su;
CREATE VIEW v_tong_hop_loi_cau_hinh_tu_dong_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM tong_hop_loi_cau_hinh_tu_dong t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_tong_hop_loi_cau_hinh_tu_dong_moi_nhat;
CREATE VIEW v_tong_hop_loi_cau_hinh_tu_dong_moi_nhat AS
SELECT *
FROM v_tong_hop_loi_cau_hinh_tu_dong_lich_su
WHERE ngay_du_lieu = (SELECT MAX(ngay_du_lieu) FROM bao_cao_ngay WHERE trang_thai_nap = 'thanh_cong');

DROP VIEW IF EXISTS v_vat_tu_thu_hoi_lich_su;
CREATE VIEW v_vat_tu_thu_hoi_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM vat_tu_thu_hoi t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_vat_tu_thu_hoi_moi_nhat;
CREATE VIEW v_vat_tu_thu_hoi_moi_nhat AS
SELECT v.*
FROM v_vat_tu_thu_hoi_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'vat_tu_thu_hoi_bc_thu_hoi_vat_tu';

DROP VIEW IF EXISTS v_chi_tiet_vat_tu_thu_hoi_lich_su;
CREATE VIEW v_chi_tiet_vat_tu_thu_hoi_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM chi_tiet_vat_tu_thu_hoi t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_chi_tiet_vat_tu_thu_hoi_moi_nhat;
CREATE VIEW v_chi_tiet_vat_tu_thu_hoi_moi_nhat AS
SELECT v.*
FROM v_chi_tiet_vat_tu_thu_hoi_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'vat_tu_thu_hoi_bc_thu_hoi_vat_tu';

DROP VIEW IF EXISTS v_quyet_toan_vat_tu_lich_su;
CREATE VIEW v_quyet_toan_vat_tu_lich_su AS
SELECT b.ma_bao_cao, b.ngay_du_lieu, b.thang_bao_cao, b.nam_bao_cao, t.*
FROM quyet_toan_vat_tu t
JOIN bao_cao_ngay b ON b.id = t.bao_cao_ngay_id;

DROP VIEW IF EXISTS v_quyet_toan_vat_tu_moi_nhat;
CREATE VIEW v_quyet_toan_vat_tu_moi_nhat AS
SELECT v.*
FROM v_quyet_toan_vat_tu_lich_su v
JOIN v_bao_cao_ngay_moi_nhat_theo_ma_bao_cao m ON v.bao_cao_ngay_id = m.id
WHERE m.ma_bao_cao = 'vat_tu_thu_hoi_quyet_toan_vat_tu';

DROP VIEW IF EXISTS v_dashboard_chat_luong_don_vi_moi_nhat;
CREATE VIEW v_dashboard_chat_luong_don_vi_moi_nhat AS
SELECT
    'c11' AS nhom_chi_tieu,
    ngay_du_lieu,
    don_vi,
    chi_tieu_bsc,
    'ty_le_sua_chua_chat_luong_chu_dong' AS ten_chi_so_1,
    ty_le_sua_chua_chat_luong_chu_dong AS chi_so_1,
    'ty_le_bao_hong_brcd_dung_quy_dinh' AS ten_chi_so_2,
    ty_le_bao_hong_brcd_dung_quy_dinh AS chi_so_2,
    'ty_le_sua_chua_trong_ngay_tai_ccco' AS ten_chi_so_3,
    ty_le_sua_chua_trong_ngay_tai_ccco AS chi_so_3
FROM v_c11_tong_hop_moi_nhat
UNION ALL
SELECT
    'c12',
    ngay_du_lieu,
    don_vi,
    chi_tieu_bsc,
    'ty_le_hong_lap_lai',
    ty_le_hong_lap_lai,
    'ty_le_su_co_brcd',
    ty_le_su_co_brcd,
    NULL,
    NULL
FROM v_c12_tong_hop_moi_nhat
UNION ALL
SELECT
    'c13',
    ngay_du_lieu,
    don_vi,
    chi_tieu_bsc,
    'ty_le_sua_chua_dung_han',
    ty_le_sua_chua_dung_han,
    'ty_le_hong_lap_lai_kenh_tsl',
    ty_le_hong_lap_lai_kenh_tsl,
    'ty_le_su_co_kenh_tsl',
    ty_le_su_co_kenh_tsl
FROM v_c13_tong_hop_moi_nhat
UNION ALL
SELECT
    'c14',
    ngay_du_lieu,
    don_vi,
    diem_bsc,
    'ty_le_hai_long_ky_thuat_phuc_vu',
    ty_le_hai_long_ky_thuat_phuc_vu,
    'ty_le_hai_long_ky_thuat_dich_vu',
    ty_le_hai_long_ky_thuat_dich_vu,
    'ty_le_khach_hang_hai_long',
    ty_le_khach_hang_hai_long
FROM v_c14_tong_hop_moi_nhat;

DROP VIEW IF EXISTS v_dashboard_kpi_nvkt_moi_nhat;
CREATE VIEW v_dashboard_kpi_nvkt_moi_nhat AS
SELECT * FROM v_kpi_nvkt_tong_hop_moi_nhat;

DROP VIEW IF EXISTS v_dashboard_dich_vu_theo_to_moi_nhat;
CREATE VIEW v_dashboard_dich_vu_theo_to_moi_nhat AS
SELECT ngay_du_lieu, 'Fiber' AS loai_dich_vu, 'hoan_cong' AS hanh_dong, ttvt, doi_vien_thong, nvkt, COUNT(*) AS so_luong
FROM v_hoan_cong_fiber_moi_nhat
GROUP BY ngay_du_lieu, ttvt, doi_vien_thong, nvkt
UNION ALL
SELECT ngay_du_lieu, 'Fiber', 'ngung_psc', ttvt, doi_vien_thong, nvkt, COUNT(*)
FROM v_ngung_psc_fiber_moi_nhat
GROUP BY ngay_du_lieu, ttvt, doi_vien_thong, nvkt
UNION ALL
SELECT ngay_du_lieu, 'Fiber', 'khoi_phuc', ttvt, doi_vien_thong, nvkt, COUNT(*)
FROM v_khoi_phuc_fiber_moi_nhat
GROUP BY ngay_du_lieu, ttvt, doi_vien_thong, nvkt
UNION ALL
SELECT ngay_du_lieu, 'MyTV', 'hoan_cong', ten_ttvt AS ttvt, doi_vien_thong, nhan_vien_ky_thuat AS nvkt, COUNT(*)
FROM v_hoan_cong_mytv_moi_nhat
GROUP BY ngay_du_lieu, ten_ttvt, doi_vien_thong, nhan_vien_ky_thuat
UNION ALL
SELECT ngay_du_lieu, 'MyTV', 'ngung_psc', ten_ttvt AS ttvt, ten_doi AS doi_vien_thong, '' AS nvkt, COUNT(*)
FROM v_ngung_psc_mytv_moi_nhat
GROUP BY ngay_du_lieu, ten_ttvt, ten_doi;

DROP VIEW IF EXISTS v_dashboard_thuc_tang_moi_nhat;
CREATE VIEW v_dashboard_thuc_tang_moi_nhat AS
SELECT ngay_du_lieu, 'Fiber' AS loai_dich_vu, cap_tong_hop, ttvt, doi_vien_thong, nvkt, hoan_cong, ngung_phat_sinh_cuoc, thuc_tang, ty_le_ngung_psc
FROM v_thuc_tang_fiber_moi_nhat
UNION ALL
SELECT ngay_du_lieu, 'MyTV', cap_tong_hop, ttvt, doi_vien_thong, nvkt, hoan_cong, ngung_phat_sinh_cuoc, thuc_tang, ty_le_ngung_psc
FROM v_thuc_tang_mytv_moi_nhat;

DROP VIEW IF EXISTS v_dashboard_xac_minh_moi_nhat;
CREATE VIEW v_dashboard_xac_minh_moi_nhat AS
SELECT ngay_du_lieu, ttvt, doi_vien_thong, nvkt, so_phieu_xac_minh
FROM v_xac_minh_nvkt_moi_nhat;

DROP VIEW IF EXISTS v_dashboard_cau_hinh_tu_dong_moi_nhat;
CREATE VIEW v_dashboard_cau_hinh_tu_dong_moi_nhat AS
SELECT
    ngay_du_lieu,
    ttvt,
    don_vi,
    loai_dong,
    tong_hop_dong,
    khong_thuc_hien_cau_hinh_tu_dong,
    da_day_cau_hinh_tu_dong,
    khong_day_do_loi_he_thong,
    khong_day_do_tbi_da_co_cau_hinh,
    cau_hinh_thanh_cong,
    ty_le_day_tu_dong,
    ty_le_tbi_da_co_cau_hinh,
    ty_le_cau_hinh_thanh_cong
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat;

DROP VIEW IF EXISTS v_dashboard_vat_tu_thu_hoi_moi_nhat;
CREATE VIEW v_dashboard_vat_tu_thu_hoi_moi_nhat AS
SELECT
    ngay_du_lieu,
    nvkt_dia_ban_giao,
    loai_vat_tu,
    trang_thai_thu_hoi,
    COUNT(*) AS so_luong
FROM v_vat_tu_thu_hoi_moi_nhat
GROUP BY ngay_du_lieu, nvkt_dia_ban_giao, loai_vat_tu, trang_thai_thu_hoi;

DROP VIEW IF EXISTS v_dashboard_quyet_toan_vat_tu_moi_nhat;
CREATE VIEW v_dashboard_quyet_toan_vat_tu_moi_nhat AS
SELECT
    ngay_du_lieu,
    loai,
    SUM(so_luong) AS tong_so_luong,
    SUM(thanh_tien) AS tong_thanh_tien
FROM v_quyet_toan_vat_tu_moi_nhat
GROUP BY ngay_du_lieu, loai;

DROP VIEW IF EXISTS v_dashboard_ttvt_son_tay_chi_so_don_vi_moi_nhat;
CREATE VIEW v_dashboard_ttvt_son_tay_chi_so_don_vi_moi_nhat AS
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây' AS ttvt,
    'chat_luong' AS nhom_du_lieu,
    nhom_chi_tieu,
    CASE WHEN don_vi = 'Tổng' THEN 'tong' ELSE 'don_vi' END AS cap_du_lieu,
    don_vi,
    NULL AS loai_dich_vu,
    NULL AS hanh_dong,
    ten_chi_so_1 AS ten_chi_so,
    chi_so_1 AS gia_tri_so,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat' AS nguon_view
FROM v_dashboard_chat_luong_don_vi_moi_nhat
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'chat_luong',
    nhom_chi_tieu,
    CASE WHEN don_vi = 'Tổng' THEN 'tong' ELSE 'don_vi' END,
    don_vi,
    NULL,
    NULL,
    ten_chi_so_2,
    chi_so_2,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat'
FROM v_dashboard_chat_luong_don_vi_moi_nhat
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'chat_luong',
    nhom_chi_tieu,
    CASE WHEN don_vi = 'Tổng' THEN 'tong' ELSE 'don_vi' END,
    don_vi,
    NULL,
    NULL,
    ten_chi_so_3,
    chi_so_3,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat'
FROM v_dashboard_chat_luong_don_vi_moi_nhat
WHERE ten_chi_so_3 IS NOT NULL
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'chat_luong',
    nhom_chi_tieu,
    CASE WHEN don_vi = 'Tổng' THEN 'tong' ELSE 'don_vi' END,
    don_vi,
    NULL,
    NULL,
    'chi_tieu_bsc',
    chi_tieu_bsc,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat'
FROM v_dashboard_chat_luong_don_vi_moi_nhat
UNION ALL
SELECT
    ngay_du_lieu,
    ttvt,
    'ghtt',
    'ghtt',
    CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
    don_vi,
    NULL,
    NULL,
    'hoan_thanh_t',
    hoan_thanh_t,
    ty_le_tong,
    'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'giao_nvkt_t', giao_nvkt_t, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'ty_le_t', ty_le_t, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'hoan_thanh_t_cong_1', hoan_thanh_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'giao_nvkt_t_cong_1', giao_nvkt_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'ty_le_t_cong_1', ty_le_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'so_luong_ghtt_lon_hon_6_thang', so_luong_ghtt_lon_hon_6_thang, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'hoan_thanh_lon_hon_6_thang_t_cong_1', hoan_thanh_lon_hon_6_thang_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'ty_le_lon_hon_6_thang_t_cong_1', ty_le_lon_hon_6_thang_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt',
       CASE WHEN don_vi = ttvt OR don_vi = 'Tổng' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'ty_le_tong', ty_le_tong, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'ket_qua_tiep_thi',
    'ket_qua_tiep_thi',
    'don_vi',
    don_vi,
    NULL,
    NULL,
    'dich_vu_brcd',
    dich_vu_brcd,
    tong,
    'v_ket_qua_tiep_thi_don_vi_moi_nhat'
FROM v_ket_qua_tiep_thi_don_vi_moi_nhat
UNION ALL
SELECT ngay_du_lieu, 'TTVT Sơn Tây', 'ket_qua_tiep_thi', 'ket_qua_tiep_thi',
       'don_vi', don_vi, NULL, NULL, 'dich_vu_mytv', dich_vu_mytv, tong, 'v_ket_qua_tiep_thi_don_vi_moi_nhat'
FROM v_ket_qua_tiep_thi_don_vi_moi_nhat
UNION ALL
SELECT ngay_du_lieu, 'TTVT Sơn Tây', 'ket_qua_tiep_thi', 'ket_qua_tiep_thi',
       'don_vi', don_vi, NULL, NULL, 'tong', tong, tong, 'v_ket_qua_tiep_thi_don_vi_moi_nhat'
FROM v_ket_qua_tiep_thi_don_vi_moi_nhat
UNION ALL
SELECT
    ngay_du_lieu,
    ttvt,
    'dich_vu',
    hanh_dong,
    'don_vi',
    doi_vien_thong,
    loai_dich_vu,
    hanh_dong,
    loai_dich_vu || '_' || hanh_dong || '_so_luong',
    so_luong,
    so_luong,
    'v_dashboard_dich_vu_theo_to_moi_nhat'
FROM v_dashboard_dich_vu_theo_to_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT
    ngay_du_lieu,
    ttvt,
    'thuc_tang',
    loai_dich_vu,
    CASE WHEN cap_tong_hop = 'to' THEN 'don_vi' ELSE cap_tong_hop END,
    doi_vien_thong,
    loai_dich_vu,
    NULL,
    'hoan_cong',
    hoan_cong,
    thuc_tang,
    'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'thuc_tang', loai_dich_vu,
       CASE WHEN cap_tong_hop = 'to' THEN 'don_vi' ELSE cap_tong_hop END,
       doi_vien_thong, loai_dich_vu, NULL, 'ngung_phat_sinh_cuoc', ngung_phat_sinh_cuoc, thuc_tang, 'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'thuc_tang', loai_dich_vu,
       CASE WHEN cap_tong_hop = 'to' THEN 'don_vi' ELSE cap_tong_hop END,
       doi_vien_thong, loai_dich_vu, NULL, 'thuc_tang', thuc_tang, thuc_tang, 'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'thuc_tang', loai_dich_vu,
       CASE WHEN cap_tong_hop = 'to' THEN 'don_vi' ELSE cap_tong_hop END,
       doi_vien_thong, loai_dich_vu, NULL, 'ty_le_ngung_psc', ty_le_ngung_psc, thuc_tang, 'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'xac_minh',
    'xac_minh',
    'don_vi',
    doi_vien_thong,
    NULL,
    NULL,
    'so_phieu_xac_minh',
    SUM(so_phieu_xac_minh),
    SUM(so_phieu_xac_minh),
    'v_dashboard_xac_minh_moi_nhat'
FROM v_dashboard_xac_minh_moi_nhat
WHERE ttvt LIKE '%Sơn Tây%'
GROUP BY ngay_du_lieu, doi_vien_thong
UNION ALL
SELECT
    ngay_du_lieu,
    ttvt,
    'cau_hinh_tu_dong',
    CASE
        WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
        WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
        ELSE loai_dong
    END,
    CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
    don_vi,
    NULL,
    NULL,
    'tong_hop_dong',
    tong_hop_dong,
    cau_hinh_thanh_cong,
    'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'khong_thuc_hien_cau_hinh_tu_dong', khong_thuc_hien_cau_hinh_tu_dong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'da_day_cau_hinh_tu_dong', da_day_cau_hinh_tu_dong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'khong_day_do_loi_he_thong', khong_day_do_loi_he_thong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'khong_day_do_tbi_da_co_cau_hinh', khong_day_do_tbi_da_co_cau_hinh, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'cau_hinh_thanh_cong', cau_hinh_thanh_cong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'ty_le_day_tu_dong', ty_le_day_tu_dong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'ty_le_tbi_da_co_cau_hinh', ty_le_tbi_da_co_cau_hinh, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       CASE WHEN loai_dong = 'TTVT' THEN 'ttvt' ELSE 'don_vi' END,
       don_vi, NULL, NULL, 'ty_le_cau_hinh_thanh_cong', ty_le_cau_hinh_thanh_cong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây';

DROP VIEW IF EXISTS v_dashboard_ttvt_son_tay_tong_hop_moi_nhat;
CREATE VIEW v_dashboard_ttvt_son_tay_tong_hop_moi_nhat AS
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây' AS ttvt,
    'chat_luong' AS nhom_du_lieu,
    nhom_chi_tieu,
    'TTVT Sơn Tây' AS don_vi,
    ten_chi_so_1 AS ten_chi_so,
    chi_so_1 AS gia_tri_so,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat' AS nguon_view
FROM v_dashboard_chat_luong_don_vi_moi_nhat
WHERE don_vi = 'Tổng'
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'chat_luong',
    nhom_chi_tieu,
    'TTVT Sơn Tây',
    ten_chi_so_2,
    chi_so_2,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat'
FROM v_dashboard_chat_luong_don_vi_moi_nhat
WHERE don_vi = 'Tổng'
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'chat_luong',
    nhom_chi_tieu,
    'TTVT Sơn Tây',
    ten_chi_so_3,
    chi_so_3,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat'
FROM v_dashboard_chat_luong_don_vi_moi_nhat
WHERE don_vi = 'Tổng' AND ten_chi_so_3 IS NOT NULL
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'chat_luong',
    nhom_chi_tieu,
    'TTVT Sơn Tây',
    'chi_tieu_bsc',
    chi_tieu_bsc,
    chi_tieu_bsc,
    'v_dashboard_chat_luong_don_vi_moi_nhat'
FROM v_dashboard_chat_luong_don_vi_moi_nhat
WHERE don_vi = 'Tổng'
UNION ALL
SELECT
    ngay_du_lieu,
    ttvt,
    'ghtt',
    'ghtt',
    'TTVT Sơn Tây',
    'hoan_thanh_t',
    hoan_thanh_t,
    ty_le_tong,
    'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'giao_nvkt_t', giao_nvkt_t, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'ty_le_t', ty_le_t, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'hoan_thanh_t_cong_1', hoan_thanh_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'giao_nvkt_t_cong_1', giao_nvkt_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'ty_le_t_cong_1', ty_le_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'so_luong_ghtt_lon_hon_6_thang', so_luong_ghtt_lon_hon_6_thang, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'hoan_thanh_lon_hon_6_thang_t_cong_1', hoan_thanh_lon_hon_6_thang_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'ty_le_lon_hon_6_thang_t_cong_1', ty_le_lon_hon_6_thang_t_cong_1, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'ghtt', 'ghtt', 'TTVT Sơn Tây', 'ty_le_tong', ty_le_tong, ty_le_tong, 'v_ghtt_don_vi_moi_nhat'
FROM v_ghtt_don_vi_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND don_vi = 'TTVT Sơn Tây'
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'ket_qua_tiep_thi',
    'ket_qua_tiep_thi',
    'TTVT Sơn Tây',
    'dich_vu_brcd',
    SUM(dich_vu_brcd),
    SUM(tong),
    'v_ket_qua_tiep_thi_don_vi_moi_nhat'
FROM v_ket_qua_tiep_thi_don_vi_moi_nhat
GROUP BY ngay_du_lieu
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'ket_qua_tiep_thi',
    'ket_qua_tiep_thi',
    'TTVT Sơn Tây',
    'dich_vu_mytv',
    SUM(dich_vu_mytv),
    SUM(tong),
    'v_ket_qua_tiep_thi_don_vi_moi_nhat'
FROM v_ket_qua_tiep_thi_don_vi_moi_nhat
GROUP BY ngay_du_lieu
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'ket_qua_tiep_thi',
    'ket_qua_tiep_thi',
    'TTVT Sơn Tây',
    'tong',
    SUM(tong),
    SUM(tong),
    'v_ket_qua_tiep_thi_don_vi_moi_nhat'
FROM v_ket_qua_tiep_thi_don_vi_moi_nhat
GROUP BY ngay_du_lieu
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'dich_vu',
    loai_dich_vu,
    'TTVT Sơn Tây',
    loai_dich_vu || '_' || hanh_dong || '_so_luong',
    SUM(so_luong),
    SUM(so_luong),
    'v_dashboard_dich_vu_theo_to_moi_nhat'
FROM v_dashboard_dich_vu_theo_to_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây'
GROUP BY ngay_du_lieu, loai_dich_vu, hanh_dong
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'thuc_tang',
    loai_dich_vu,
    'TTVT Sơn Tây',
    'hoan_cong',
    SUM(hoan_cong),
    SUM(thuc_tang),
    'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
GROUP BY ngay_du_lieu, loai_dich_vu
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'thuc_tang',
    loai_dich_vu,
    'TTVT Sơn Tây',
    'ngung_phat_sinh_cuoc',
    SUM(ngung_phat_sinh_cuoc),
    SUM(thuc_tang),
    'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
GROUP BY ngay_du_lieu, loai_dich_vu
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'thuc_tang',
    loai_dich_vu,
    'TTVT Sơn Tây',
    'thuc_tang',
    SUM(thuc_tang),
    SUM(thuc_tang),
    'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
GROUP BY ngay_du_lieu, loai_dich_vu
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'thuc_tang',
    loai_dich_vu,
    'TTVT Sơn Tây',
    'ty_le_ngung_psc',
    CASE
        WHEN SUM(COALESCE(hoan_cong, 0)) = 0 THEN NULL
        ELSE ROUND(SUM(COALESCE(ngung_phat_sinh_cuoc, 0)) * 100.0 / SUM(COALESCE(hoan_cong, 0)), 2)
    END,
    SUM(thuc_tang),
    'v_dashboard_thuc_tang_moi_nhat'
FROM v_dashboard_thuc_tang_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND cap_tong_hop = 'to'
GROUP BY ngay_du_lieu, loai_dich_vu
UNION ALL
SELECT
    ngay_du_lieu,
    'TTVT Sơn Tây',
    'xac_minh',
    'xac_minh',
    'TTVT Sơn Tây',
    'so_phieu_xac_minh',
    SUM(so_phieu_xac_minh),
    SUM(so_phieu_xac_minh),
    'v_dashboard_xac_minh_moi_nhat'
FROM v_dashboard_xac_minh_moi_nhat
WHERE ttvt LIKE '%Sơn Tây%'
GROUP BY ngay_du_lieu
UNION ALL
SELECT
    ngay_du_lieu,
    ttvt,
    'cau_hinh_tu_dong',
    CASE
        WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
        WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
        ELSE loai_dong
    END,
    'TTVT Sơn Tây',
    'tong_hop_dong',
    tong_hop_dong,
    cau_hinh_thanh_cong,
    'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'khong_thuc_hien_cau_hinh_tu_dong', khong_thuc_hien_cau_hinh_tu_dong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'da_day_cau_hinh_tu_dong', da_day_cau_hinh_tu_dong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'khong_day_do_loi_he_thong', khong_day_do_loi_he_thong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'khong_day_do_tbi_da_co_cau_hinh', khong_day_do_tbi_da_co_cau_hinh, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'cau_hinh_thanh_cong', cau_hinh_thanh_cong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'ty_le_day_tu_dong', ty_le_day_tu_dong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'ty_le_tbi_da_co_cau_hinh', ty_le_tbi_da_co_cau_hinh, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT'
UNION ALL
SELECT ngay_du_lieu, ttvt, 'cau_hinh_tu_dong',
       CASE
           WHEN ma_bao_cao LIKE '%_ptm' THEN 'ptm'
           WHEN ma_bao_cao LIKE '%_thay_the' THEN 'thay_the'
           ELSE loai_dong
       END,
       'TTVT Sơn Tây', 'ty_le_cau_hinh_thanh_cong', ty_le_cau_hinh_thanh_cong, cau_hinh_thanh_cong, 'v_cau_hinh_tu_dong_tong_hop_moi_nhat'
FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
WHERE ttvt = 'TTVT Sơn Tây' AND loai_dong = 'TTVT';

DROP VIEW IF EXISTS v_dashboard_chi_so_nvkt_moi_nhat;
CREATE VIEW v_dashboard_chi_so_nvkt_moi_nhat AS
SELECT *
FROM (
    SELECT
        ngay_du_lieu,
        NULL AS ttvt,
        don_vi,
        nvkt,
        'kpi_nvkt' AS nhom_du_lieu,
        nhom_chi_tieu,
        'sm1' AS ten_chi_so,
        CAST(sm1 AS REAL) AS gia_tri_so,
        chi_tieu_bsc,
        NULL AS loai_dich_vu,
        NULL AS hanh_dong,
        'v_dashboard_kpi_nvkt_moi_nhat' AS nguon_view
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, 'sm2', CAST(sm2 AS REAL), chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, 'sm3', CAST(sm3 AS REAL), chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, 'sm4', CAST(sm4 AS REAL), chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, 'sm5', CAST(sm5 AS REAL), chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, 'sm6', CAST(sm6 AS REAL), chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, ten_chi_so_1, chi_so_1, chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, ten_chi_so_2, chi_so_2, chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, ten_chi_so_3, chi_so_3, chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'kpi_nvkt', nhom_chi_tieu, 'chi_tieu_bsc', chi_tieu_bsc, chi_tieu_bsc, NULL, NULL, 'v_dashboard_kpi_nvkt_moi_nhat'
    FROM v_dashboard_kpi_nvkt_moi_nhat
    UNION ALL
    SELECT
        ngay_du_lieu,
        NULL AS ttvt,
        doi_vien_thong AS don_vi,
        nvkt,
        'hai_long_nvkt' AS nhom_du_lieu,
        'c14' AS nhom_chi_tieu,
        'tong_phieu_khao_sat_thanh_cong' AS ten_chi_so,
        CAST(tong_phieu_khao_sat_thanh_cong AS REAL) AS gia_tri_so,
        ty_le_hai_long_nvkt AS chi_tieu_bsc,
        NULL AS loai_dich_vu,
        NULL AS hanh_dong,
        'v_c14_nvkt_moi_nhat' AS nguon_view
    FROM v_c14_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, doi_vien_thong, nvkt, 'hai_long_nvkt', 'c14', 'tong_phieu_khong_hai_long', CAST(tong_phieu_khong_hai_long AS REAL), ty_le_hai_long_nvkt, NULL, NULL, 'v_c14_nvkt_moi_nhat'
    FROM v_c14_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, doi_vien_thong, nvkt, 'hai_long_nvkt', 'c14', 'ty_le_hai_long_nvkt', ty_le_hai_long_nvkt, ty_le_hai_long_nvkt, NULL, NULL, 'v_c14_nvkt_moi_nhat'
    FROM v_c14_nvkt_moi_nhat
    UNION ALL
    SELECT
        ngay_du_lieu,
        NULL AS ttvt,
        don_vi,
        nvkt,
        'ghtt' AS nhom_du_lieu,
        'ghtt' AS nhom_chi_tieu,
        'hoan_thanh_t' AS ten_chi_so,
        CAST(hoan_thanh_t AS REAL) AS gia_tri_so,
        ty_le_tong AS chi_tieu_bsc,
        NULL AS loai_dich_vu,
        NULL AS hanh_dong,
        'v_ghtt_nvkt_moi_nhat' AS nguon_view
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'giao_nvkt_t', CAST(giao_nvkt_t AS REAL), ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'ty_le_t', ty_le_t, ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'hoan_thanh_t_cong_1', CAST(hoan_thanh_t_cong_1 AS REAL), ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'giao_nvkt_t_cong_1', CAST(giao_nvkt_t_cong_1 AS REAL), ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'ty_le_t_cong_1', ty_le_t_cong_1, ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'so_luong_ghtt_lon_hon_6_thang', CAST(so_luong_ghtt_lon_hon_6_thang AS REAL), ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'hoan_thanh_lon_hon_6_thang_t_cong_1', CAST(hoan_thanh_lon_hon_6_thang_t_cong_1 AS REAL), ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'ty_le_lon_hon_6_thang_t_cong_1', ty_le_lon_hon_6_thang_t_cong_1, ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, nvkt, 'ghtt', 'ghtt', 'ty_le_tong', ty_le_tong, ty_le_tong, NULL, NULL, 'v_ghtt_nvkt_moi_nhat'
    FROM v_ghtt_nvkt_moi_nhat
    UNION ALL
    SELECT
        ngay_du_lieu,
        NULL AS ttvt,
        don_vi,
        ten_nv AS nvkt,
        'ket_qua_tiep_thi' AS nhom_du_lieu,
        'ket_qua_tiep_thi' AS nhom_chi_tieu,
        'dich_vu_brcd' AS ten_chi_so,
        CAST(dich_vu_brcd AS REAL) AS gia_tri_so,
        tong AS chi_tieu_bsc,
        NULL AS loai_dich_vu,
        NULL AS hanh_dong,
        'v_ket_qua_tiep_thi_nv_moi_nhat' AS nguon_view
    FROM v_ket_qua_tiep_thi_nv_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, ten_nv, 'ket_qua_tiep_thi', 'ket_qua_tiep_thi', 'dich_vu_mytv', CAST(dich_vu_mytv AS REAL), tong, NULL, NULL, 'v_ket_qua_tiep_thi_nv_moi_nhat'
    FROM v_ket_qua_tiep_thi_nv_moi_nhat
    UNION ALL
    SELECT ngay_du_lieu, NULL, don_vi, ten_nv, 'ket_qua_tiep_thi', 'ket_qua_tiep_thi', 'tong', CAST(tong AS REAL), tong, NULL, NULL, 'v_ket_qua_tiep_thi_nv_moi_nhat'
    FROM v_ket_qua_tiep_thi_nv_moi_nhat
    UNION ALL
    SELECT
        ngay_du_lieu,
        ttvt,
        doi_vien_thong AS don_vi,
        nvkt,
        'dich_vu' AS nhom_du_lieu,
        loai_dich_vu AS nhom_chi_tieu,
        loai_dich_vu || '_' || hanh_dong || '_so_luong' AS ten_chi_so,
        CAST(so_luong AS REAL) AS gia_tri_so,
        so_luong AS chi_tieu_bsc,
        loai_dich_vu,
        hanh_dong,
        'v_dashboard_dich_vu_theo_to_moi_nhat' AS nguon_view
    FROM v_dashboard_dich_vu_theo_to_moi_nhat
    WHERE TRIM(COALESCE(nvkt, '')) <> ''
    UNION ALL
    SELECT
        ngay_du_lieu,
        ttvt,
        doi_vien_thong AS don_vi,
        nvkt,
        'thuc_tang' AS nhom_du_lieu,
        loai_dich_vu AS nhom_chi_tieu,
        'hoan_cong' AS ten_chi_so,
        hoan_cong AS gia_tri_so,
        thuc_tang AS chi_tieu_bsc,
        loai_dich_vu,
        NULL AS hanh_dong,
        'v_dashboard_thuc_tang_moi_nhat' AS nguon_view
    FROM v_dashboard_thuc_tang_moi_nhat
    WHERE cap_tong_hop = 'nvkt' AND TRIM(COALESCE(nvkt, '')) <> ''
    UNION ALL
    SELECT ngay_du_lieu, ttvt, doi_vien_thong, nvkt, 'thuc_tang', loai_dich_vu, 'ngung_phat_sinh_cuoc', ngung_phat_sinh_cuoc, thuc_tang, loai_dich_vu, NULL, 'v_dashboard_thuc_tang_moi_nhat'
    FROM v_dashboard_thuc_tang_moi_nhat
    WHERE cap_tong_hop = 'nvkt' AND TRIM(COALESCE(nvkt, '')) <> ''
    UNION ALL
    SELECT ngay_du_lieu, ttvt, doi_vien_thong, nvkt, 'thuc_tang', loai_dich_vu, 'thuc_tang', thuc_tang, thuc_tang, loai_dich_vu, NULL, 'v_dashboard_thuc_tang_moi_nhat'
    FROM v_dashboard_thuc_tang_moi_nhat
    WHERE cap_tong_hop = 'nvkt' AND TRIM(COALESCE(nvkt, '')) <> ''
    UNION ALL
    SELECT ngay_du_lieu, ttvt, doi_vien_thong, nvkt, 'thuc_tang', loai_dich_vu, 'ty_le_ngung_psc', ty_le_ngung_psc, thuc_tang, loai_dich_vu, NULL, 'v_dashboard_thuc_tang_moi_nhat'
    FROM v_dashboard_thuc_tang_moi_nhat
    WHERE cap_tong_hop = 'nvkt' AND TRIM(COALESCE(nvkt, '')) <> ''
    UNION ALL
    SELECT
        ngay_du_lieu,
        ttvt,
        doi_vien_thong AS don_vi,
        nvkt,
        'xac_minh' AS nhom_du_lieu,
        'xac_minh' AS nhom_chi_tieu,
        'so_phieu_xac_minh' AS ten_chi_so,
        CAST(so_phieu_xac_minh AS REAL) AS gia_tri_so,
        so_phieu_xac_minh AS chi_tieu_bsc,
        NULL AS loai_dich_vu,
        NULL AS hanh_dong,
        'v_dashboard_xac_minh_moi_nhat' AS nguon_view
    FROM v_dashboard_xac_minh_moi_nhat
    UNION ALL
    SELECT
        ngay_du_lieu,
        ttvt,
        doi_vien_thong AS don_vi,
        nvkt,
        'cau_hinh_tu_dong' AS nhom_du_lieu,
        CASE
            WHEN TRIM(COALESCE(loai_cau_hinh, '')) = '' THEN 'khong_xac_dinh'
            ELSE loai_cau_hinh
        END AS nhom_chi_tieu,
        CASE
            WHEN TRIM(COALESCE(trang_thai_chuan_hoa, '')) = '' THEN 'khong_co_trang_thai'
            ELSE trang_thai_chuan_hoa
        END AS ten_chi_so,
        CAST(COUNT(*) AS REAL) AS gia_tri_so,
        CAST(COUNT(*) AS REAL) AS chi_tieu_bsc,
        NULL AS loai_dich_vu,
        NULL AS hanh_dong,
        'v_cau_hinh_tu_dong_chi_tiet_moi_nhat' AS nguon_view
    FROM v_cau_hinh_tu_dong_chi_tiet_moi_nhat
    WHERE TRIM(COALESCE(nvkt, '')) <> ''
    GROUP BY ngay_du_lieu, ttvt, doi_vien_thong, nvkt, loai_cau_hinh, trang_thai_chuan_hoa
    UNION ALL
    SELECT
        ngay_du_lieu,
        NULL AS ttvt,
        NULL AS don_vi,
        nvkt_dia_ban_giao AS nvkt,
        'vat_tu_thu_hoi' AS nhom_du_lieu,
        CASE
            WHEN TRIM(COALESCE(loai_vat_tu, '')) = '' THEN 'khong_xac_dinh'
            ELSE loai_vat_tu
        END AS nhom_chi_tieu,
        CASE
            WHEN TRIM(COALESCE(trang_thai_thu_hoi, '')) = '' THEN 'so_luong'
            ELSE trang_thai_thu_hoi
        END AS ten_chi_so,
        CAST(so_luong AS REAL) AS gia_tri_so,
        so_luong AS chi_tieu_bsc,
        NULL AS loai_dich_vu,
        NULL AS hanh_dong,
        'v_dashboard_vat_tu_thu_hoi_moi_nhat' AS nguon_view
    FROM v_dashboard_vat_tu_thu_hoi_moi_nhat
    WHERE TRIM(COALESCE(nvkt_dia_ban_giao, '')) <> ''
)
WHERE TRIM(COALESCE(nvkt, '')) <> ''
  AND gia_tri_so IS NOT NULL;

COMMIT;
