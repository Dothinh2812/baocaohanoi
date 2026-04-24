-- View layer cho schema summary per-report-per-sheet

BEGIN;

CREATE VIEW IF NOT EXISTS v_bao_cao_tong_hop_moi_nhat_theo_ma_bao_cao AS
SELECT b.*
FROM bao_cao_tong_hop_ngay b
JOIN (
    SELECT ma_bao_cao, MAX(ngay_du_lieu) AS ngay_du_lieu
    FROM bao_cao_tong_hop_ngay
    WHERE trang_thai_nap IN ('thanh_cong', 'khong_co_sheet_tong_hop')
    GROUP BY ma_bao_cao
) mx
  ON b.ma_bao_cao = mx.ma_bao_cao
 AND b.ngay_du_lieu = mx.ngay_du_lieu;

CREATE VIEW IF NOT EXISTS v_bao_cao_tong_hop_moi_nhat_toan_bo AS
SELECT *
FROM bao_cao_tong_hop_ngay
WHERE ngay_du_lieu = (
    SELECT MAX(ngay_du_lieu)
    FROM bao_cao_tong_hop_ngay
    WHERE trang_thai_nap IN ('thanh_cong', 'khong_co_sheet_tong_hop')
);

CREATE VIEW IF NOT EXISTS v_nhat_ky_nap_tong_hop_gan_nhat AS
SELECT n.*
FROM nhat_ky_nap_tong_hop n
JOIN (
    SELECT ma_bao_cao, MAX(id) AS id_moi_nhat
    FROM nhat_ky_nap_tong_hop
    GROUP BY ma_bao_cao
) mx
  ON n.ma_bao_cao = mx.ma_bao_cao
 AND n.id = mx.id_moi_nhat;

CREATE VIEW IF NOT EXISTS v_tong_hop_bang_du_lieu_theo_bao_cao AS
SELECT
    ma_bao_cao,
    COUNT(*) AS so_bang_du_lieu_dang_ky,
    GROUP_CONCAT(ten_bang_du_lieu, ', ') AS danh_sach_bang_du_lieu
FROM danh_muc_bang_du_lieu_bao_cao
GROUP BY ma_bao_cao;

CREATE VIEW IF NOT EXISTS v_tien_do_nap_tong_hop AS
SELECT
    d.ma_bao_cao,
    d.ten_bao_cao,
    d.nhom_bao_cao,
    d.duong_dan_processed_mac_dinh,
    g.so_bang_du_lieu_dang_ky,
    g.danh_sach_bang_du_lieu,
    b.ngay_du_lieu,
    b.trang_thai_nap,
    b.so_sheet_tong_hop,
    b.so_bang_du_lieu,
    b.so_dong_tong_hop,
    b.so_chi_tieu_tong_hop,
    b.ten_tep_nguon,
    b.duong_dan_tep_nguon,
    b.ghi_chu,
    n.bat_dau_luc,
    n.ket_thuc_luc,
    n.trang_thai AS trang_thai_nhat_ky,
    n.thong_diep
FROM danh_muc_bao_cao_tong_hop d
LEFT JOIN v_tong_hop_bang_du_lieu_theo_bao_cao g
  ON d.ma_bao_cao = g.ma_bao_cao
LEFT JOIN v_bao_cao_tong_hop_moi_nhat_theo_ma_bao_cao b
  ON d.ma_bao_cao = b.ma_bao_cao
LEFT JOIN v_nhat_ky_nap_tong_hop_gan_nhat n
  ON d.ma_bao_cao = n.ma_bao_cao;

CREATE VIEW IF NOT EXISTS v_tep_nguon_bao_cao_tong_hop_moi_nhat AS
SELECT
    b.ma_bao_cao,
    b.ngay_du_lieu,
    t.*
FROM tep_nguon_bao_cao_tong_hop t
JOIN v_bao_cao_tong_hop_moi_nhat_theo_ma_bao_cao b
  ON t.bao_cao_tong_hop_ngay_id = b.id;

CREATE VIEW IF NOT EXISTS v_sheet_bao_cao_tong_hop_moi_nhat AS
SELECT
    b.ma_bao_cao,
    b.ngay_du_lieu,
    s.*
FROM sheet_bao_cao_tong_hop s
JOIN v_bao_cao_tong_hop_moi_nhat_theo_ma_bao_cao b
  ON s.bao_cao_tong_hop_ngay_id = b.id;

CREATE VIEW IF NOT EXISTS v_danh_muc_bang_du_lieu_bao_cao AS
SELECT
    d.ma_bao_cao,
    d.ten_bao_cao,
    d.nhom_bao_cao,
    g.ten_sheet_goc,
    g.ma_sheet,
    g.ten_bang_du_lieu,
    g.che_do_luu_tru,
    g.tong_so_cot,
    g.danh_sach_cot_json,
    g.mo_ta,
    g.thoi_gian_tao,
    g.thoi_gian_cap_nhat
FROM danh_muc_bao_cao_tong_hop d
JOIN danh_muc_bang_du_lieu_bao_cao g
  ON d.ma_bao_cao = g.ma_bao_cao;

CREATE VIEW IF NOT EXISTS v_cau_hinh_tu_dong_chi_tiet_th_theo_to AS
SELECT
    "Trung tâm Viễn thông",
    "Đội Viễn thông",
    "Tổng hợp đồng",
    "Lắp mới",
    "Thay thế",
    "Cấu hình WAN",
    "Cấu hình WiFi",
    "Thành công",
    "Thất bại",
    "Chưa có trạng thái",
    "Tỷ lệ thành công (%)",
    "Tỷ lệ thất bại (%)"
FROM "cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet_th_theo_to";

CREATE VIEW IF NOT EXISTS v_cau_hinh_tu_dong_chi_tiet_th_theo_nvkt AS
SELECT
    "Trung tâm Viễn thông",
    "Đội Viễn thông",
    "NVKT",
    "Tổng hợp đồng",
    "Lắp mới",
    "Thay thế",
    "Cấu hình WAN",
    "Cấu hình WiFi",
    "Thành công",
    "Thất bại",
    "Chưa có trạng thái",
    "Tỷ lệ thành công (%)",
    "Tỷ lệ thất bại (%)"
FROM "cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet_th_theo_nvkt";

CREATE VIEW IF NOT EXISTS v_ghtt_hni_kq_hni AS
SELECT
    "Đơn vị",
    "Hoàn thành T",
    "Giao NVKT T",
    "Tỷ lệ T",
    "Hoàn thành T+1",
    "Giao NVKT T+1",
    "Tỷ lệ T+1",
    "SL GHTT >=6T",
    "Hoàn thành >=6T T+1",
    "Tỷ lệ >=6T T+1",
    "Tỷ lệ Tổng"
FROM "ghtt_ghtt_hni_report_kq_hni";

CREATE VIEW IF NOT EXISTS v_ghtt_sontay_kq_sontay AS
SELECT
    "Đơn vị",
    "Hoàn thành T",
    "Giao NVKT T",
    "Tỷ lệ T",
    "Hoàn thành T+1",
    "Giao NVKT T+1",
    "Tỷ lệ T+1",
    "SL GHTT >=6T",
    "Hoàn thành >=6T T+1",
    "Tỷ lệ >=6T T+1",
    "Tỷ lệ Tổng"
FROM "ghtt_ghtt_sontay_report_kq_sontay";

CREATE VIEW IF NOT EXISTS v_ghtt_nvktdb_kq_nvktdb AS
SELECT
    "NVKT",
    "Đơn vị",
    "TTVT",
    "Hoàn thành T",
    "Giao NVKT T",
    "Tỷ lệ T",
    "Hoàn thành T+1",
    "Giao NVKT T+1",
    "Tỷ lệ T+1",
    "SL GHTT >=6T",
    "Hoàn thành >=6T T+1",
    "Tỷ lệ >=6T T+1",
    "Tỷ lệ Tổng"
FROM "ghtt_ghtt_nvktdb_report_kq_nvktdb";

CREATE VIEW IF NOT EXISTS v_kq_tiep_thi_kq_th AS
SELECT
    "STT",
    "Đơn vị",
    "Dịch vụ BRCĐ",
    "Dịch vụ MyTV",
    "Tổng"
FROM "kq_tiep_thi_kq_tiep_thi_report_kq_th";

CREATE VIEW IF NOT EXISTS v_kq_tiep_thi_kq_tiep_thi AS
SELECT
    "STT",
    "Đơn vị",
    "Mã NV",
    "Tên NV",
    "Dịch vụ BRCĐ",
    "Dịch vụ MyTV",
    "Tổng"
FROM "kq_tiep_thi_kq_tiep_thi_report_kq_tiep_thi";

CREATE VIEW IF NOT EXISTS v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt AS
SELECT
    "TTVT",
    "DOIVT",
    "NVKT",
    "Tạm dừng Fiber",
    "Tạm dừng MyTV",
    "Khôi phục Fiber",
    "Khôi phục MyTV",
    "Chưa khôi phục Fiber",
    "Chưa khôi phục MyTV"
FROM "tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt";

CREATE VIEW IF NOT EXISTS v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to AS
SELECT
    "TTVT",
    "DOIVT",
    "Tạm dừng Fiber",
    "Tạm dừng MyTV",
    "Khôi phục Fiber",
    "Khôi phục MyTV",
    "Chưa khôi phục Fiber",
    "Chưa khôi phục MyTV"
FROM "tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to";

CREATE VIEW IF NOT EXISTS v_ngung_psc_fiber_thang_t_1_cap_ttvt AS
SELECT
    "Đơn vị/Nhân viên KT",
    "Phát triển mới | Yêu cầu(*) | Lắp mới",
    "Phát triển mới | Yêu cầu(*) | Phục hủy",
    "Phát triển mới | Yêu cầu(*) | Lũy kế tháng",
    "Phát triển mới | Yêu cầu(*) | Lũy kế năm",
    "Phát triển mới | Hoàn công | Hoàn công(*) (1.5)",
    "Phát triển mới | Hoàn công | Lũy kế tháng(1.6)",
    "Phát triển mới | Hoàn công | Lũy kế năm(1.7)",
    "Phát triển mới | Tồn(trong tháng)",
    "Phát triển mới | Tồn(lũy kế)",
    "Số liệu hủy | Số thuê bao(*)(2.1)",
    "Số liệu hủy | Lũy kế tháng(2.2)",
    "Số liệu hủy | Lũy kế năm(2.3)",
    "Số liệu hủy | Số TB hủy trong tòa nhà(*)(2.4)",
    "Số liệu Phát sinh cước | TB PSC tháng chọn(hoặc tháng T-1)(3.1)",
    "Số liệu Phát sinh cước | TB PSC tháng lũy kế năm(3.2)",
    "Số liệu Phát sinh cước | Hủy trong tháng PSC chọn(3.3)",
    "Số liệu Phát sinh cước | Hủy lũy kế PSC(3.4)",
    "Số liệu Phát sinh cước | TB PSC tháng trước(hoặc tháng T-2)(3.5)",
    "Số liệu Phát sinh cước | PSC thực tăng LK tháng T-1(3.6)",
    "Số liệu tạm dừng | Số thuê bao dừng trong ngày(*)",
    "Số liệu tạm dừng | Tổng số thuê bao có trạng thái tạm dừng tính đến ngày (*)",
    "Số liệu tạm dừng | Tạm dừng lũy kế tháng",
    "Chỉ tiêu thuê bao ngưng PSC | TB PSC thực tăng tháng chọn(4.1)(Cột 3.1 - 3.5)",
    "Chỉ tiêu thuê bao ngưng PSC | TB ngưng PSC tháng(4.2)(cột 1.6 - 4.1)",
    "Chỉ tiêu thuê bao ngưng PSC | SL TB PSC năm trước(4.3)",
    "Chỉ tiêu thuê bao ngưng PSC | SLTB thực tăng lũy kế năm(4.4)",
    "Chỉ tiêu thuê bao ngưng PSC | TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)",
    "Chỉ tiêu thuê bao ngưng PSC | PTM lũy kế năm đến tháng PSC(4.7)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % TB ngưng PSC/Thuê bao PTM(4.8) (cột 4.2/1.6)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % TB ngưng PSC/Thuê bao PSC(4.9)(cột 4.2/3.1)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % ngưng PSC/PTM Lũy kế Năm(4.9.1)(cột 4.6/4.7)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % ngưng PSC Lũy kế năm(4.9.2)(cột 4.6/3.2)",
    "Chỉ tiêu tạm tính tháng T | Ngưng PSC tạm tính tháng T(5.1)",
    "Chỉ tiêu tạm tính tháng T | Thuê bao PSC tạm tính đến ngày thứ N của tháng T",
    "Chỉ tiêu tạm tính tháng T | Số tb tạm dừng các tháng T-2, T-3 đã khôi phục tính đến ngày N của tháng T",
    "Chỉ tiêu tạm tính tháng T | Tỷ lệ Ngưng PSC/PTM tạm tính tháng T(5.2)",
    "Chỉ tiêu tạm tính tháng T | Tỷ lệ Ngưng PSC/PTM tạm tính lũy kế tháng T(5.3)"
FROM "tam_dung_khoi_phuc_dich_vu_ngung_psc_fiber_thang_t_1_cap_ttvt_th_ngung_psc_thang_t_1";

CREATE VIEW IF NOT EXISTS v_ngung_psc_mytv_thang_t_1_cap_ttvt AS
SELECT
    "Đơn vị/Nhân viên KT",
    "Phát triển mới | Yêu cầu(*) | Lắp mới",
    "Phát triển mới | Yêu cầu(*) | Phục hủy",
    "Phát triển mới | Yêu cầu(*) | Lũy kế tháng",
    "Phát triển mới | Yêu cầu(*) | Lũy kế năm",
    "Phát triển mới | Hoàn công | Hoàn công(*) (1.5)",
    "Phát triển mới | Hoàn công | Lũy kế tháng(1.6)",
    "Phát triển mới | Hoàn công | Lũy kế năm(1.7)",
    "Phát triển mới | Tồn(trong tháng)",
    "Phát triển mới | Tồn(lũy kế)",
    "Số liệu hủy | Số thuê bao(*)(2.1)",
    "Số liệu hủy | Lũy kế tháng(2.2)",
    "Số liệu hủy | Lũy kế năm(2.3)",
    "Số liệu hủy | Số TB hủy trong tòa nhà(*)(2.4)",
    "Số liệu Phát sinh cước | TB PSC tháng chọn(hoặc tháng T-1)(3.1)",
    "Số liệu Phát sinh cước | TB PSC tháng lũy kế năm(3.2)",
    "Số liệu Phát sinh cước | Hủy trong tháng PSC chọn(3.3)",
    "Số liệu Phát sinh cước | Hủy lũy kế PSC(3.4)",
    "Số liệu Phát sinh cước | TB PSC tháng trước(hoặc tháng T-2)(3.5)",
    "Số liệu Phát sinh cước | PSC thực tăng LK tháng T-1(3.6)",
    "Số liệu tạm dừng | Số thuê bao dừng trong ngày(*)",
    "Số liệu tạm dừng | Tổng số thuê bao có trạng thái tạm dừng tính đến ngày (*)",
    "Số liệu tạm dừng | Tạm dừng lũy kế tháng",
    "Chỉ tiêu thuê bao ngưng PSC | TB PSC thực tăng tháng chọn(4.1)(Cột 3.1 - 3.5)",
    "Chỉ tiêu thuê bao ngưng PSC | TB ngưng PSC tháng(4.2)(cột 1.6 - 4.1)",
    "Chỉ tiêu thuê bao ngưng PSC | SL TB PSC năm trước(4.3)",
    "Chỉ tiêu thuê bao ngưng PSC | SLTB thực tăng lũy kế năm(4.4)",
    "Chỉ tiêu thuê bao ngưng PSC | TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)",
    "Chỉ tiêu thuê bao ngưng PSC | PTM lũy kế năm đến tháng PSC(4.7)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % TB ngưng PSC/Thuê bao PTM(4.8) (cột 4.2/1.6)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % TB ngưng PSC/Thuê bao PSC(4.9)(cột 4.2/3.1)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % ngưng PSC/PTM Lũy kế Năm(4.9.1)(cột 4.6/4.7)",
    "Chỉ tiêu thuê bao ngưng PSC | Tỷ lệ % ngưng PSC Lũy kế năm(4.9.2)(cột 4.6/3.2)",
    "Chỉ tiêu tạm tính tháng T | Ngưng PSC tạm tính tháng T(5.1)",
    "Chỉ tiêu tạm tính tháng T | Thuê bao PSC tạm tính đến ngày thứ N của tháng T",
    "Chỉ tiêu tạm tính tháng T | Số tb tạm dừng các tháng T-2, T-3 đã khôi phục tính đến ngày N của tháng T",
    "Chỉ tiêu tạm tính tháng T | Tỷ lệ Ngưng PSC/PTM tạm tính tháng T(5.2)",
    "Chỉ tiêu tạm tính tháng T | Tỷ lệ Ngưng PSC/PTM tạm tính lũy kế tháng T(5.3)"
FROM "tam_dung_khoi_phuc_dich_vu_ngung_psc_mytv_thang_t_1_cap_ttvt_th_ngung_psc_thang_t_1";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_1_chitiet_report_chi_tiet AS
SELECT
    "TEN_DOI",
    "NVKT",
    "Tổng phiếu",
    "Số phiếu đạt",
    "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn"
FROM "chi_tieu_c_c1_1_chitiet_report_chi_tiet";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h AS
SELECT
    "TEN_DOI",
    "NVKT",
    "Tổng phiếu",
    "Số phiếu đạt",
    "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn"
FROM "chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_1_report_th_c1_1 AS
SELECT
    "Đơn vị",
    "SM1",
    "SM2",
    "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",
    "SM3",
    "SM4",
    "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn",
    "SM5",
    "SM6",
    "Tỷ lệ phiếu sửa chữa trong ngày tại CCCO",
    "Chỉ tiêu BSC"
FROM "chi_tieu_c_c1_1_report_th_c1_1";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang AS
SELECT
    "TEN_DOI",
    "NVKT",
    "Số phiếu HLL",
    "Số phiếu báo hỏng",
    "Tỉ lệ HLL tháng (2.5%)"
FROM "chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_2_report_th_c1_2 AS
SELECT
    "Đơn vị",
    "SM1",
    "SM2",
    "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",
    "SM3",
    "SM4",
    "Tỷ lệ sự cố dịch vụ BRCĐ",
    "Chỉ tiêu BSC"
FROM "chi_tieu_c_c1_2_report_th_c1_2";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_3_report_th_c1_3 AS
SELECT
    "Đơn vị",
    "SM1",
    "SM2",
    "Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định",
    "SM3",
    "SM4",
    "Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại",
    "SM5",
    "SM6",
    "Tỷ lệ sự cố dịch vụ kênh TSL",
    "Chỉ tiêu BSC"
FROM "chi_tieu_c_c1_3_report_th_c1_3";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt AS
SELECT
    "DOIVT",
    "NVKT",
    "Tổng phiếu KS thành công",
    "Tổng phiếu KHL",
    "Tỉ lệ HL NVKT (%)"
FROM "chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_4_report_th_c1_4 AS
SELECT
    "Đơn vị",
    "Tổng phiếu",
    "SL đã KS",
    "SL KS thành công",
    "SL KH hài lòng",
    "Không HL KT phục vụ",
    "Tỷ lệ HL KT phục vụ",
    "Không HL KT dịch vụ",
    "Tỷ lệ HL KT dịch vụ",
    "Tổng phiếu hài lòng KT",
    "Tỷ lệ KH hài lòng",
    "Điểm BSC"
FROM "chi_tieu_c_c1_4_report_th_c1_4";

CREATE VIEW IF NOT EXISTS v_chi_tieu_c_c1_5_report_th_c1_5 AS
SELECT
    "Đơn vị",
    "Tổng - SM1",
    "Tổng - SM2",
    "Tổng - KQ thực hiện chỉ tiêu",
    "Tổng - Điểm BSC",
    "CCCO - SM1",
    "CCCO - SM2",
    "CCCO - Tỷ lệ",
    "CCCO - Điểm BSC",
    "Không CCCO - SM1",
    "Không CCCO - SM2",
    "Không CCCO - Tỷ lệ",
    "Không CCCO - Điểm BSC",
    "CCCO xã hội hóa - SM1",
    "CCCO xã hội hóa - SM2"
FROM "chi_tieu_c_c1_5_report_th_c1_5";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_k2_report_bien_dong_tong_hop AS
SELECT
    "TT",
    "Đơn vị",
    "NVKT_DB",
    "Tổng số hiện tại",
    "Tăng mới",
    "Giảm/Hết",
    "Vẫn còn",
    "Số TB quản lý",
    "Tỉ lệ SHC (%)"
FROM "chi_tieu_i_i1_5_k2_report_bien_dong_tong_hop";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_k2_report_shc_theo_sa AS
SELECT
    "TT",
    "SA",
    "Số lượng"
FROM "chi_tieu_i_i1_5_k2_report_shc_theo_sa";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_k2_report_th_shc_i15 AS
SELECT
    "TT",
    "Đơn vị",
    "NVKT_DB",
    "Số TB Suy hao cao K2",
    "Số TB quản lý",
    "Tỉ lệ SHC (%)"
FROM "chi_tieu_i_i1_5_k2_report_th_shc_i15";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_k2_report_th_shc_theo_to AS
SELECT
    "TT",
    "Đơn vị",
    "Số TB Suy hao cao K2",
    "Số TB quản lý",
    "Tỉ lệ SHC (%)"
FROM "chi_tieu_i_i1_5_k2_report_th_shc_theo_to";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_report_bien_dong_tong_hop AS
SELECT
    "TT",
    "Đơn vị",
    "NVKT_DB",
    "Tổng số hiện tại",
    "Tăng mới",
    "Giảm/Hết",
    "Vẫn còn",
    "Số TB quản lý",
    "Tỉ lệ SHC (%)"
FROM "chi_tieu_i_i1_5_report_bien_dong_tong_hop";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_report_shc_theo_sa AS
SELECT
    "TT",
    "SA",
    "Số lượng"
FROM "chi_tieu_i_i1_5_report_shc_theo_sa";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_report_th_shc_i15 AS
SELECT
    "TT",
    "Đơn vị",
    "NVKT_DB",
    "Số TB Suy hao cao K1",
    "Số TB quản lý",
    "Tỉ lệ SHC (%)"
FROM "chi_tieu_i_i1_5_report_th_shc_i15";

CREATE VIEW IF NOT EXISTS v_chi_tieu_i_i1_5_report_th_shc_theo_to AS
SELECT
    "TT",
    "Đơn vị",
    "Số TB Suy hao cao K1",
    "Số TB quản lý",
    "Tỉ lệ SHC (%)"
FROM "chi_tieu_i_i1_5_report_th_shc_theo_to";

CREATE VIEW IF NOT EXISTS v_nvkt_tong_hop_da_nguon AS
WITH
c11 AS (
    SELECT
        TRIM("NVKT") AS nhan_vien,
        "TEN_DOI" AS to_doi,
        "Tổng phiếu" AS c11_tong_phieu,
        "Số phiếu đạt" AS c11_so_phieu_dat,
        "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn" AS c11_ty_le
    FROM "chi_tieu_c_c1_1_chitiet_report_chi_tiet"
    WHERE TRIM(COALESCE("NVKT", '')) <> ''
      AND TRIM("NVKT") <> 'Tổng'
),
c12_sm1 AS (
    SELECT
        TRIM("NVKT") AS nhan_vien,
        "TEN_DOI" AS to_doi,
        "Số phiếu HLL" AS c12_sm1_so_phieu_hll,
        "Số phiếu báo hỏng" AS c12_sm1_so_phieu_bao_hong,
        "Tỉ lệ HLL tháng (2.5%)" AS c12_sm1_ty_le_hll
    FROM "chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang"
    WHERE TRIM(COALESCE("NVKT", '')) <> ''
      AND TRIM("NVKT") <> 'Tổng'
),
c14 AS (
    SELECT
        TRIM("NVKT") AS nhan_vien,
        "DOIVT" AS to_doi,
        "Tổng phiếu KS thành công" AS c14_tong_phieu_ks_thanh_cong,
        "Tổng phiếu KHL" AS c14_tong_phieu_khl,
        "Tỉ lệ HL NVKT (%)" AS c14_ty_le_hl
    FROM "chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt"
    WHERE TRIM(COALESCE("NVKT", '')) <> ''
      AND TRIM("NVKT") <> 'Tổng'
),
i15_k1 AS (
    SELECT
        TRIM("NVKT_DB") AS nhan_vien,
        "Đơn vị" AS to_doi,
        "Số TB Suy hao cao K1" AS i15_k1_so_tb_shc,
        "Số TB quản lý" AS i15_k1_so_tb_quan_ly,
        "Tỉ lệ SHC (%)" AS i15_k1_ty_le_shc
    FROM "chi_tieu_i_i1_5_report_th_shc_i15"
    WHERE TRIM(COALESCE("NVKT_DB", '')) <> ''
      AND TRIM("NVKT_DB") <> 'Tổng'
),
i15_k2 AS (
    SELECT
        TRIM("NVKT_DB") AS nhan_vien,
        "Đơn vị" AS to_doi,
        "Số TB Suy hao cao K2" AS i15_k2_so_tb_shc,
        "Số TB quản lý" AS i15_k2_so_tb_quan_ly,
        "Tỉ lệ SHC (%)" AS i15_k2_ty_le_shc
    FROM "chi_tieu_i_i1_5_k2_report_th_shc_i15"
    WHERE TRIM(COALESCE("NVKT_DB", '')) <> ''
      AND TRIM("NVKT_DB") <> 'Tổng'
),
ghtt AS (
    SELECT
        TRIM("NVKT") AS nhan_vien,
        "Đơn vị" AS to_doi,
        "TTVT" AS ttvt,
        "Hoàn thành T" AS ghtt_hoan_thanh_t,
        "Giao NVKT T" AS ghtt_giao_nvkt_t,
        "Tỷ lệ T" AS ghtt_ty_le_t,
        "Hoàn thành T+1" AS ghtt_hoan_thanh_t1,
        "Giao NVKT T+1" AS ghtt_giao_nvkt_t1,
        "Tỷ lệ T+1" AS ghtt_ty_le_t1,
        "SL GHTT >=6T" AS ghtt_sl_6t,
        "Hoàn thành >=6T T+1" AS ghtt_hoan_thanh_6t_t1,
        "Tỷ lệ >=6T T+1" AS ghtt_ty_le_6t_t1,
        "Tỷ lệ Tổng" AS ghtt_ty_le_tong
    FROM "ghtt_ghtt_nvktdb_report_kq_nvktdb"
    WHERE TRIM(COALESCE("NVKT", '')) <> ''
      AND TRIM("NVKT") <> 'Tổng'
),
kpi_c11 AS (
    SELECT
        TRIM("NVKT") AS nhan_vien,
        "đơn vị" AS to_doi,
        "SM1" AS kpi_c11_sm1,
        "SM2" AS kpi_c11_sm2,
        "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu" AS kpi_c11_ty_le_dat_yeu_cau,
        "SM3" AS kpi_c11_sm3,
        "SM4" AS kpi_c11_sm4,
        "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCĐ đúng quy định không tính hẹn" AS kpi_c11_ty_le_dung_hen,
        "Chỉ tiêu BSC" AS kpi_c11_chi_tieu_bsc
    FROM "kpi_nvkt_c11_nvktdb_report_c11_kpi_nvkt"
    WHERE TRIM(COALESCE("NVKT", '')) <> ''
      AND TRIM("NVKT") <> 'Tổng'
),
kpi_c12 AS (
    SELECT
        TRIM("NVKT") AS nhan_vien,
        "đơn vị" AS to_doi,
        "SM1" AS kpi_c12_sm1,
        "SM2" AS kpi_c12_sm2,
        "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại" AS kpi_c12_ty_le_lap_lai,
        "SM3" AS kpi_c12_sm3,
        "SM4" AS kpi_c12_sm4,
        "Tỷ lệ sự cố dịch vụ BRCĐ" AS kpi_c12_ty_le_su_co,
        "Chỉ tiêu BSC" AS kpi_c12_chi_tieu_bsc
    FROM "kpi_nvkt_c12_nvktdb_report_c12_kpi_nvkt"
    WHERE TRIM(COALESCE("NVKT", '')) <> ''
      AND TRIM("NVKT") <> 'Tổng'
),
kq_tiep_thi AS (
    SELECT
        TRIM("Tên NV") AS nhan_vien,
        "Đơn vị" AS to_doi,
        "Mã NV" AS ma_nv,
        "Dịch vụ BRCĐ" AS kqtt_brcd,
        "Dịch vụ MyTV" AS kqtt_mytv,
        "Tổng" AS kqtt_tong
    FROM "kq_tiep_thi_kq_tiep_thi_report_kq_tiep_thi"
    WHERE TRIM(COALESCE("Tên NV", '')) <> ''
      AND TRIM("Tên NV") <> 'Tổng'
),
keys AS (
    SELECT nhan_vien FROM c11
    UNION
    SELECT nhan_vien FROM c12_sm1
    UNION
    SELECT nhan_vien FROM c14
    UNION
    SELECT nhan_vien FROM i15_k1
    UNION
    SELECT nhan_vien FROM i15_k2
    UNION
    SELECT nhan_vien FROM ghtt
    UNION
    SELECT nhan_vien FROM kpi_c11
    UNION
    SELECT nhan_vien FROM kpi_c12
    UNION
    SELECT nhan_vien FROM kq_tiep_thi
)
SELECT
    keys.nhan_vien AS nvkt_hoac_ten_nv,
    COALESCE(
        c11.to_doi,
        c12_sm1.to_doi,
        c14.to_doi,
        i15_k1.to_doi,
        i15_k2.to_doi,
        ghtt.to_doi,
        kpi_c11.to_doi,
        kpi_c12.to_doi,
        kq_tiep_thi.to_doi
    ) AS to_doi_hoac_don_vi,
    ghtt.ttvt,
    kq_tiep_thi.ma_nv,
    c11.c11_tong_phieu,
    c11.c11_so_phieu_dat,
    c11.c11_ty_le,
    c12_sm1.c12_sm1_so_phieu_hll,
    c12_sm1.c12_sm1_so_phieu_bao_hong,
    c12_sm1.c12_sm1_ty_le_hll,
    c14.c14_tong_phieu_ks_thanh_cong,
    c14.c14_tong_phieu_khl,
    c14.c14_ty_le_hl,
    i15_k1.i15_k1_so_tb_shc,
    i15_k1.i15_k1_so_tb_quan_ly,
    i15_k1.i15_k1_ty_le_shc,
    i15_k2.i15_k2_so_tb_shc,
    i15_k2.i15_k2_so_tb_quan_ly,
    i15_k2.i15_k2_ty_le_shc,
    ghtt.ghtt_hoan_thanh_t,
    ghtt.ghtt_giao_nvkt_t,
    ghtt.ghtt_ty_le_t,
    ghtt.ghtt_hoan_thanh_t1,
    ghtt.ghtt_giao_nvkt_t1,
    ghtt.ghtt_ty_le_t1,
    ghtt.ghtt_sl_6t,
    ghtt.ghtt_hoan_thanh_6t_t1,
    ghtt.ghtt_ty_le_6t_t1,
    ghtt.ghtt_ty_le_tong,
    kpi_c11.kpi_c11_sm1,
    kpi_c11.kpi_c11_sm2,
    kpi_c11.kpi_c11_ty_le_dat_yeu_cau,
    kpi_c11.kpi_c11_sm3,
    kpi_c11.kpi_c11_sm4,
    kpi_c11.kpi_c11_ty_le_dung_hen,
    kpi_c11.kpi_c11_chi_tieu_bsc,
    kpi_c12.kpi_c12_sm1,
    kpi_c12.kpi_c12_sm2,
    kpi_c12.kpi_c12_ty_le_lap_lai,
    kpi_c12.kpi_c12_sm3,
    kpi_c12.kpi_c12_sm4,
    kpi_c12.kpi_c12_ty_le_su_co,
    kpi_c12.kpi_c12_chi_tieu_bsc,
    kq_tiep_thi.kqtt_brcd,
    kq_tiep_thi.kqtt_mytv,
    kq_tiep_thi.kqtt_tong
FROM keys
LEFT JOIN c11 ON c11.nhan_vien = keys.nhan_vien
LEFT JOIN c12_sm1 ON c12_sm1.nhan_vien = keys.nhan_vien
LEFT JOIN c14 ON c14.nhan_vien = keys.nhan_vien
LEFT JOIN i15_k1 ON i15_k1.nhan_vien = keys.nhan_vien
LEFT JOIN i15_k2 ON i15_k2.nhan_vien = keys.nhan_vien
LEFT JOIN ghtt ON ghtt.nhan_vien = keys.nhan_vien
LEFT JOIN kpi_c11 ON kpi_c11.nhan_vien = keys.nhan_vien
LEFT JOIN kpi_c12 ON kpi_c12.nhan_vien = keys.nhan_vien
LEFT JOIN kq_tiep_thi ON kq_tiep_thi.nhan_vien = keys.nhan_vien
ORDER BY keys.nhan_vien;

CREATE VIEW IF NOT EXISTS v_don_vi_tong_hop_da_nguon AS
WITH
c11 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "SM1" AS c11_sm1,
        "SM2" AS c11_sm2,
        "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu" AS c11_ty_le_dat_yeu_cau,
        "SM3" AS c11_sm3,
        "SM4" AS c11_sm4,
        "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn" AS c11_ty_le_dung_hen,
        "SM5" AS c11_sm5,
        "SM6" AS c11_sm6,
        "Tỷ lệ phiếu sửa chữa trong ngày tại CCCO" AS c11_ty_le_ccco,
        "Chỉ tiêu BSC" AS c11_chi_tieu_bsc
    FROM "chi_tieu_c_c1_1_report_th_c1_1"
),
c12 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "SM1" AS c12_sm1,
        "SM2" AS c12_sm2,
        "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại" AS c12_ty_le_lap_lai,
        "SM3" AS c12_sm3,
        "SM4" AS c12_sm4,
        "Tỷ lệ sự cố dịch vụ BRCĐ" AS c12_ty_le_su_co,
        "Chỉ tiêu BSC" AS c12_chi_tieu_bsc
    FROM "chi_tieu_c_c1_2_report_th_c1_2"
),
c13 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "SM1" AS c13_sm1,
        "SM2" AS c13_sm2,
        "Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định" AS c13_ty_le_dung_hen,
        "SM3" AS c13_sm3,
        "SM4" AS c13_sm4,
        "Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại" AS c13_ty_le_lap_lai,
        "SM5" AS c13_sm5,
        "SM6" AS c13_sm6,
        "Tỷ lệ sự cố dịch vụ kênh TSL" AS c13_ty_le_su_co,
        "Chỉ tiêu BSC" AS c13_chi_tieu_bsc
    FROM "chi_tieu_c_c1_3_report_th_c1_3"
),
c14 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Tổng phiếu" AS c14_tong_phieu,
        "SL đã KS" AS c14_sl_da_ks,
        "SL KS thành công" AS c14_sl_ks_thanh_cong,
        "SL KH hài lòng" AS c14_sl_kh_hai_long,
        "Không HL KT phục vụ" AS c14_khong_hl_kt_phuc_vu,
        "Tỷ lệ HL KT phục vụ" AS c14_ty_le_hl_kt_phuc_vu,
        "Không HL KT dịch vụ" AS c14_khong_hl_kt_dich_vu,
        "Tỷ lệ HL KT dịch vụ" AS c14_ty_le_hl_kt_dich_vu,
        "Tổng phiếu hài lòng KT" AS c14_tong_phieu_hai_long_kt,
        "Tỷ lệ KH hài lòng" AS c14_ty_le_kh_hai_long,
        "Điểm BSC" AS c14_diem_bsc
    FROM "chi_tieu_c_c1_4_report_th_c1_4"
),
c15 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Tổng - SM1" AS c15_tong_sm1,
        "Tổng - SM2" AS c15_tong_sm2,
        "Tổng - KQ thực hiện chỉ tiêu" AS c15_tong_kq_thuc_hien,
        "Tổng - Điểm BSC" AS c15_tong_diem_bsc,
        "CCCO - SM1" AS c15_ccco_sm1,
        "CCCO - SM2" AS c15_ccco_sm2,
        "CCCO - Tỷ lệ" AS c15_ccco_ty_le,
        "CCCO - Điểm BSC" AS c15_ccco_diem_bsc,
        "Không CCCO - SM1" AS c15_khong_ccco_sm1,
        "Không CCCO - SM2" AS c15_khong_ccco_sm2,
        "Không CCCO - Tỷ lệ" AS c15_khong_ccco_ty_le,
        "Không CCCO - Điểm BSC" AS c15_khong_ccco_diem_bsc,
        "CCCO xã hội hóa - SM1" AS c15_ccco_xhh_sm1,
        "CCCO xã hội hóa - SM2" AS c15_ccco_xhh_sm2
    FROM "chi_tieu_c_c1_5_report_th_c1_5"
),
ghtt AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Hoàn thành T" AS ghtt_hoan_thanh_t,
        "Giao NVKT T" AS ghtt_giao_nvkt_t,
        "Tỷ lệ T" AS ghtt_ty_le_t,
        "Hoàn thành T+1" AS ghtt_hoan_thanh_t1,
        "Giao NVKT T+1" AS ghtt_giao_nvkt_t1,
        "Tỷ lệ T+1" AS ghtt_ty_le_t1,
        "SL GHTT >=6T" AS ghtt_sl_6t,
        "Hoàn thành >=6T T+1" AS ghtt_hoan_thanh_6t_t1,
        "Tỷ lệ >=6T T+1" AS ghtt_ty_le_6t_t1,
        "Tỷ lệ Tổng" AS ghtt_ty_le_tong
    FROM "ghtt_ghtt_sontay_report_kq_sontay"
),
cau_hinh AS (
    SELECT
        "Đội Viễn thông" AS don_vi,
        "Trung tâm Viễn thông" AS cau_hinh_ttvt,
        "Tổng hợp đồng" AS cau_hinh_tong_hop_dong,
        "Lắp mới" AS cau_hinh_lap_moi,
        "Thay thế" AS cau_hinh_thay_the,
        "Cấu hình WAN" AS cau_hinh_wan,
        "Cấu hình WiFi" AS cau_hinh_wifi,
        "Thành công" AS cau_hinh_thanh_cong,
        "Thất bại" AS cau_hinh_that_bai,
        "Chưa có trạng thái" AS cau_hinh_chua_co_trang_thai,
        "Tỷ lệ thành công (%)" AS cau_hinh_ty_le_thanh_cong,
        "Tỷ lệ thất bại (%)" AS cau_hinh_ty_le_that_bai
    FROM "cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet_th_theo_to"
),
kq_tiep_thi AS (
    SELECT
        CASE TRIM("Đơn vị")
            WHEN 'Tổ Phúc Thọ' THEN 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
            WHEN 'Tổ Quảng oai' THEN 'Tổ Kỹ thuật Địa bàn Quảng Oai'
            WHEN 'Tổ Sơn Tây' THEN 'Tổ Kỹ thuật Địa bàn Sơn Tây'
            WHEN 'Tổ Suối Hai' THEN 'Tổ Kỹ thuật Địa bàn Suối hai'
            ELSE TRIM("Đơn vị")
        END AS don_vi,
        "Dịch vụ BRCĐ" AS kqtt_brcd,
        "Dịch vụ MyTV" AS kqtt_mytv,
        "Tổng" AS kqtt_tong
    FROM "kq_tiep_thi_kq_tiep_thi_report_kq_th"
),
keys AS (
    SELECT don_vi FROM c11
    UNION
    SELECT don_vi FROM c12
    UNION
    SELECT don_vi FROM c13
    UNION
    SELECT don_vi FROM c14
    UNION
    SELECT don_vi FROM c15
    UNION
    SELECT don_vi FROM ghtt
    UNION
    SELECT don_vi FROM cau_hinh
    UNION
    SELECT don_vi FROM kq_tiep_thi
)
SELECT
    keys.don_vi AS don_vi,
    cau_hinh.cau_hinh_ttvt,
    c11.c11_sm1,
    c11.c11_sm2,
    c11.c11_ty_le_dat_yeu_cau,
    c11.c11_sm3,
    c11.c11_sm4,
    c11.c11_ty_le_dung_hen,
    c11.c11_sm5,
    c11.c11_sm6,
    c11.c11_ty_le_ccco,
    c11.c11_chi_tieu_bsc,
    c12.c12_sm1,
    c12.c12_sm2,
    c12.c12_ty_le_lap_lai,
    c12.c12_sm3,
    c12.c12_sm4,
    c12.c12_ty_le_su_co,
    c12.c12_chi_tieu_bsc,
    c13.c13_sm1,
    c13.c13_sm2,
    c13.c13_ty_le_dung_hen,
    c13.c13_sm3,
    c13.c13_sm4,
    c13.c13_ty_le_lap_lai,
    c13.c13_sm5,
    c13.c13_sm6,
    c13.c13_ty_le_su_co,
    c13.c13_chi_tieu_bsc,
    c14.c14_tong_phieu,
    c14.c14_sl_da_ks,
    c14.c14_sl_ks_thanh_cong,
    c14.c14_sl_kh_hai_long,
    c14.c14_khong_hl_kt_phuc_vu,
    c14.c14_ty_le_hl_kt_phuc_vu,
    c14.c14_khong_hl_kt_dich_vu,
    c14.c14_ty_le_hl_kt_dich_vu,
    c14.c14_tong_phieu_hai_long_kt,
    c14.c14_ty_le_kh_hai_long,
    c14.c14_diem_bsc,
    c15.c15_tong_sm1,
    c15.c15_tong_sm2,
    c15.c15_tong_kq_thuc_hien,
    c15.c15_tong_diem_bsc,
    c15.c15_ccco_sm1,
    c15.c15_ccco_sm2,
    c15.c15_ccco_ty_le,
    c15.c15_ccco_diem_bsc,
    c15.c15_khong_ccco_sm1,
    c15.c15_khong_ccco_sm2,
    c15.c15_khong_ccco_ty_le,
    c15.c15_khong_ccco_diem_bsc,
    c15.c15_ccco_xhh_sm1,
    c15.c15_ccco_xhh_sm2,
    ghtt.ghtt_hoan_thanh_t,
    ghtt.ghtt_giao_nvkt_t,
    ghtt.ghtt_ty_le_t,
    ghtt.ghtt_hoan_thanh_t1,
    ghtt.ghtt_giao_nvkt_t1,
    ghtt.ghtt_ty_le_t1,
    ghtt.ghtt_sl_6t,
    ghtt.ghtt_hoan_thanh_6t_t1,
    ghtt.ghtt_ty_le_6t_t1,
    ghtt.ghtt_ty_le_tong,
    cau_hinh.cau_hinh_tong_hop_dong,
    cau_hinh.cau_hinh_lap_moi,
    cau_hinh.cau_hinh_thay_the,
    cau_hinh.cau_hinh_wan,
    cau_hinh.cau_hinh_wifi,
    cau_hinh.cau_hinh_thanh_cong,
    cau_hinh.cau_hinh_that_bai,
    cau_hinh.cau_hinh_chua_co_trang_thai,
    cau_hinh.cau_hinh_ty_le_thanh_cong,
    cau_hinh.cau_hinh_ty_le_that_bai,
    kq_tiep_thi.kqtt_brcd,
    kq_tiep_thi.kqtt_mytv,
    kq_tiep_thi.kqtt_tong
FROM keys
LEFT JOIN c11 ON c11.don_vi = keys.don_vi
LEFT JOIN c12 ON c12.don_vi = keys.don_vi
LEFT JOIN c13 ON c13.don_vi = keys.don_vi
LEFT JOIN c14 ON c14.don_vi = keys.don_vi
LEFT JOIN c15 ON c15.don_vi = keys.don_vi
LEFT JOIN ghtt ON ghtt.don_vi = keys.don_vi
LEFT JOIN cau_hinh ON cau_hinh.don_vi = keys.don_vi
LEFT JOIN kq_tiep_thi ON kq_tiep_thi.don_vi = keys.don_vi
WHERE keys.don_vi IN (
    'Tổ Kỹ thuật Địa bàn Phúc Thọ',
    'Tổ Kỹ thuật Địa bàn Quảng Oai',
    'Tổ Kỹ thuật Địa bàn Sơn Tây',
    'Tổ Kỹ thuật Địa bàn Suối hai'
)
ORDER BY keys.don_vi;

DROP VIEW IF EXISTS "chi tieu BSC-KPI cac to";

CREATE VIEW IF NOT EXISTS "chi tieu BSC-KPI cac to" AS
WITH
c11 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu" AS c11_ty_le_sua_chua_chu_dong,
        "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn" AS c11_ty_le_brcd_dung_quy_dinh,
        "Tỷ lệ phiếu sửa chữa trong ngày tại CCCO" AS c11_ty_le_ccco,
        "Chỉ tiêu BSC" AS c11_chi_tieu_bsc
    FROM "chi_tieu_c_c1_1_report_th_c1_1"
),
c12 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại" AS c12_ty_le_lap_lai,
        "Tỷ lệ sự cố dịch vụ BRCĐ" AS c12_ty_le_su_co,
        "Chỉ tiêu BSC" AS c12_chi_tieu_bsc
    FROM "chi_tieu_c_c1_2_report_th_c1_2"
),
c13 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định" AS c13_ty_le_tsl_dung_thoi_gian
    FROM "chi_tieu_c_c1_3_report_th_c1_3"
),
c14 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Tỷ lệ KH hài lòng" AS c14_ty_le_kh_hai_long,
        "Điểm BSC" AS c14_diem_bsc
    FROM "chi_tieu_c_c1_4_report_th_c1_4"
),
c15 AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Tổng - KQ thực hiện chỉ tiêu" AS c15_tong_kq_thuc_hien_chi_tieu,
        "Tổng - Điểm BSC" AS c15_tong_diem_bsc
    FROM "chi_tieu_c_c1_5_report_th_c1_5"
),
ghtt AS (
    SELECT
        "Đơn vị" AS don_vi,
        "Hoàn thành T" AS ghtt_hoan_thanh_t,
        "Giao NVKT T" AS ghtt_giao_nvkt_t,
        "Tỷ lệ T" AS ghtt_ty_le_t,
        "Hoàn thành T+1" AS ghtt_hoan_thanh_t1,
        "Giao NVKT T+1" AS ghtt_giao_nvkt_t1,
        "Tỷ lệ T+1" AS ghtt_ty_le_t1,
        "SL GHTT >=6T" AS ghtt_sl_ghtt_6t,
        "Hoàn thành >=6T T+1" AS ghtt_hoan_thanh_6t_t1,
        "Tỷ lệ >=6T T+1" AS ghtt_ty_le_6t_t1,
        "Tỷ lệ Tổng" AS ghtt_ty_le_tong
    FROM "ghtt_ghtt_sontay_report_kq_sontay"
),
kq_tiep_thi AS (
    SELECT
        CASE TRIM("Đơn vị")
            WHEN 'Tổ Phúc Thọ' THEN 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
            WHEN 'Tổ Quảng oai' THEN 'Tổ Kỹ thuật Địa bàn Quảng Oai'
            WHEN 'Tổ Sơn Tây' THEN 'Tổ Kỹ thuật Địa bàn Sơn Tây'
            WHEN 'Tổ Suối hai' THEN 'Tổ Kỹ thuật Địa bàn Suối hai'
            WHEN 'Tổ Suối Hai' THEN 'Tổ Kỹ thuật Địa bàn Suối hai'
            ELSE TRIM("Đơn vị")
        END AS don_vi,
        "Dịch vụ BRCĐ" AS kqtt_dich_vu_brcd,
        "Dịch vụ MyTV" AS kqtt_dich_vu_mytv,
        "Tổng" AS kqtt_tong
    FROM "kq_tiep_thi_kq_tiep_thi_report_kq_th"
),
keys AS (
    SELECT don_vi FROM c11
    UNION
    SELECT don_vi FROM c12
    UNION
    SELECT don_vi FROM c13
    UNION
    SELECT don_vi FROM c14
    UNION
    SELECT don_vi FROM c15
    UNION
    SELECT don_vi FROM ghtt
    UNION
    SELECT don_vi FROM kq_tiep_thi
)
SELECT
    keys.don_vi AS "Đơn vị",
    c11.c11_ty_le_sua_chua_chu_dong AS "C1.1 - Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",
    c11.c11_ty_le_brcd_dung_quy_dinh AS "C1.1 - Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn",
    c11.c11_ty_le_ccco AS "C1.1 - Tỷ lệ phiếu sửa chữa trong ngày tại CCCO",
    c11.c11_chi_tieu_bsc AS "C1.1 - Chỉ tiêu BSC",
    c12.c12_ty_le_lap_lai AS "C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",
    c12.c12_ty_le_su_co AS "C1.2 - Tỷ lệ sự cố dịch vụ BRCĐ",
    c12.c12_chi_tieu_bsc AS "C1.2 - Chỉ tiêu BSC",
    c13.c13_ty_le_tsl_dung_thoi_gian AS "C1.3 - Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định",
    c14.c14_ty_le_kh_hai_long AS "C1.4 - Tỷ lệ KH hài lòng",
    c14.c14_diem_bsc AS "C1.4 - Điểm BSC",
    c15.c15_tong_kq_thuc_hien_chi_tieu AS "C1.5 - Tổng - KQ thực hiện chỉ tiêu",
    c15.c15_tong_diem_bsc AS "C1.5 - Tổng - Điểm BSC",
    ghtt.ghtt_hoan_thanh_t AS "GHTT - Hoàn thành T",
    ghtt.ghtt_giao_nvkt_t AS "GHTT - Giao NVKT T",
    ghtt.ghtt_ty_le_t AS "GHTT - Tỷ lệ T",
    ghtt.ghtt_hoan_thanh_t1 AS "GHTT - Hoàn thành T+1",
    ghtt.ghtt_giao_nvkt_t1 AS "GHTT - Giao NVKT T+1",
    ghtt.ghtt_ty_le_t1 AS "GHTT - Tỷ lệ T+1",
    ghtt.ghtt_sl_ghtt_6t AS "GHTT - SL GHTT >=6T",
    ghtt.ghtt_hoan_thanh_6t_t1 AS "GHTT - Hoàn thành >=6T T+1",
    ghtt.ghtt_ty_le_6t_t1 AS "GHTT - Tỷ lệ >=6T T+1",
    ghtt.ghtt_ty_le_tong AS "GHTT - Tỷ lệ Tổng",
    kq_tiep_thi.kqtt_dich_vu_brcd AS "KQTT - Dịch vụ BRCĐ",
    kq_tiep_thi.kqtt_dich_vu_mytv AS "KQTT - Dịch vụ MyTV",
    kq_tiep_thi.kqtt_tong AS "KQTT - Tổng"
FROM keys
LEFT JOIN c11 ON c11.don_vi = keys.don_vi
LEFT JOIN c12 ON c12.don_vi = keys.don_vi
LEFT JOIN c13 ON c13.don_vi = keys.don_vi
LEFT JOIN c14 ON c14.don_vi = keys.don_vi
LEFT JOIN c15 ON c15.don_vi = keys.don_vi
LEFT JOIN ghtt ON ghtt.don_vi = keys.don_vi
LEFT JOIN kq_tiep_thi ON kq_tiep_thi.don_vi = keys.don_vi
WHERE keys.don_vi IN (
    'Tổ Kỹ thuật Địa bàn Phúc Thọ',
    'Tổ Kỹ thuật Địa bàn Quảng Oai',
    'Tổ Kỹ thuật Địa bàn Sơn Tây',
    'Tổ Kỹ thuật Địa bàn Suối hai'
)
ORDER BY keys.don_vi;

DROP VIEW IF EXISTS v_chi_tieu_bsc_kpi_cac_to;

CREATE VIEW IF NOT EXISTS v_chi_tieu_bsc_kpi_cac_to AS
SELECT *
FROM "chi tieu BSC-KPI cac to";

COMMIT;
