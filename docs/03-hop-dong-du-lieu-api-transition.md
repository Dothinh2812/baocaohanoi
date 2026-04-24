# Hợp đồng dữ liệu api_transition

## Nguồn được chọn cho dashv4

DB nên dùng cho `dashv4` bản Sơn Tây:

`/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db`

Lý do:
- đây là runtime DB theo đơn vị, đúng hướng thiết kế mới
- đã có các `view` tiêu thụ cho dashboard
- tránh dùng DB root chung cho nhiều đơn vị nếu mục tiêu trước mắt là một app Sơn Tây

## Cách dữ liệu được tạo ra

Luồng hiện tại của `api_transition`:
1. tải báo cáo về `downloads`
2. chuẩn hóa thành file trong `Processed`
3. lưu snapshot theo ngày trong `ProcessedDaily/<ngày>`
4. import vào `report_history.db`
5. dựng `view` cho consumer/dashboard

Nguyên tắc cho `dashv4`:
- đọc `view` trước
- chỉ đọc bảng thô khi chưa có `view` phù hợp và phải ghi rõ lý do

Lưu ý cập nhật mới:
- với các route đã hỗ trợ hiển thị theo ngày, không đọc trực tiếp `view ... moi_nhat`
- phải truy nguyên snapshot theo `bao_cao_tong_hop_ngay.ngay_du_lieu`
- nguyên tắc chi tiết xem thêm tại:
  `docs/09-nguyen-tac-loc-ngay-report-history.md`

## Các bảng nền quan trọng

Nhóm snapshot và metadata:
- `danh_muc_bao_cao`
- `bao_cao_ngay`
- `sheet_bao_cao`
- `dong_bao_cao_goc`
- `tep_luu_tru_bao_cao`
- `nhat_ky_nap_bao_cao`
- `danh_muc_don_vi`
- `danh_muc_nhan_vien`

## Các bảng nghiệp vụ đã có

- `c11_tong_hop`, `c11_chi_tiet_nvkt`
- `c12_tong_hop`, `c12_hong_lap_lai_nvkt`
- `c13_tong_hop`
- `c14_tong_hop`, `c14_hai_long_nvkt`
- `ghtt_don_vi`, `ghtt_nvkt`
- `kpi_nvkt_c11`, `kpi_nvkt_c12`, `kpi_nvkt_c13`
- `ket_qua_tiep_thi_don_vi`, `ket_qua_tiep_thi_nv`
- `hoan_cong_fiber`, `hoan_cong_mytv`
- `khoi_phuc_fiber`
- `ngung_psc_fiber`, `ngung_psc_mytv`
- `thuc_tang_fiber`, `thuc_tang_mytv`
- `xac_minh_chi_tiet`, `xac_minh_tong_hop_loai_phieu`, `xac_minh_tong_hop_nvkt`
- `cau_hinh_tu_dong_chi_tiet`, `cau_hinh_tu_dong_tong_hop`, `tong_hop_loi_cau_hinh_tu_dong`
- `vat_tu_thu_hoi`, `chi_tiet_vat_tu_thu_hoi`, `quyet_toan_vat_tu`

## Các view dashboard-ready đã xác nhận

- `v_dashboard_chat_luong_don_vi_moi_nhat`
- `v_dashboard_kpi_nvkt_moi_nhat`
- `v_dashboard_dich_vu_theo_to_moi_nhat`
- `v_dashboard_thuc_tang_moi_nhat`
- `v_dashboard_xac_minh_moi_nhat`
- `v_dashboard_cau_hinh_tu_dong_moi_nhat`
- `v_dashboard_vat_tu_thu_hoi_moi_nhat`
- `v_dashboard_quyet_toan_vat_tu_moi_nhat`
- `v_dashboard_chi_so_nvkt_moi_nhat`
- `v_dashboard_ttvt_son_tay_chi_so_don_vi_moi_nhat`
- `v_dashboard_ttvt_son_tay_tong_hop_moi_nhat`
- `v_tien_do_nap_bao_cao`

## Các view chi tiết dùng cho drilldown

- `v_c11_tong_hop_moi_nhat`, `v_c11_nvkt_moi_nhat`
- `v_c12_tong_hop_moi_nhat`, `v_c12_nvkt_moi_nhat`
- `v_c13_tong_hop_moi_nhat`
- `v_c14_tong_hop_moi_nhat`, `v_c14_nvkt_moi_nhat`
- `v_ghtt_don_vi_moi_nhat`, `v_ghtt_nvkt_moi_nhat`
- `v_ket_qua_tiep_thi_don_vi_moi_nhat`, `v_ket_qua_tiep_thi_nv_moi_nhat`
- `v_hoan_cong_fiber_moi_nhat`, `v_hoan_cong_mytv_moi_nhat`
- `v_ngung_psc_fiber_moi_nhat`, `v_ngung_psc_mytv_moi_nhat`
- `v_khoi_phuc_fiber_moi_nhat`
- `v_thuc_tang_fiber_moi_nhat`, `v_thuc_tang_mytv_moi_nhat`
- `v_xac_minh_nvkt_moi_nhat`, `v_xac_minh_chi_tiet_moi_nhat`, `v_xac_minh_loai_phieu_moi_nhat`
- `v_cau_hinh_tu_dong_tong_hop_moi_nhat`, `v_cau_hinh_tu_dong_chi_tiet_moi_nhat`
- `v_vat_tu_thu_hoi_moi_nhat`, `v_quyet_toan_vat_tu_moi_nhat`

## Tình trạng dữ liệu kiểm tra tại ngày 2026-04-20

Snapshot runtime Sơn Tây đang có `28` báo cáo đã import thành công.

Một số view chính:

| View | Số dòng kiểm tra |
| --- | --- |
| `v_dashboard_chat_luong_don_vi_moi_nhat` | 20 |
| `v_dashboard_kpi_nvkt_moi_nhat` | 105 |
| `v_dashboard_dich_vu_theo_to_moi_nhat` | 45 |
| `v_dashboard_thuc_tang_moi_nhat` | 39 |
| `v_dashboard_xac_minh_moi_nhat` | 37 |
| `v_dashboard_cau_hinh_tu_dong_moi_nhat` | 300 |
| `v_dashboard_vat_tu_thu_hoi_moi_nhat` | 3 |
| `v_dashboard_quyet_toan_vat_tu_moi_nhat` | 16 |
| `v_dashboard_chi_so_nvkt_moi_nhat` | 2701 |
| `v_dashboard_ttvt_son_tay_tong_hop_moi_nhat` | 50 |

## Cảnh báo rất quan trọng

DB runtime Sơn Tây không đồng nghĩa mọi view đều đã tự lọc còn một `ttvt`.

Số `ttvt` distinct kiểm tra thực tế:

| View | Distinct `ttvt` |
| --- | --- |
| `v_dashboard_xac_minh_moi_nhat` | 1 |
| `v_dashboard_thuc_tang_moi_nhat` | 1 |
| `v_dashboard_dich_vu_theo_to_moi_nhat` | 1 |
| `v_dashboard_chi_so_nvkt_moi_nhat` | 19 |
| `v_dashboard_cau_hinh_tu_dong_moi_nhat` | 18 |

Kết luận:
- `dashv4` phải có tầng lọc đơn vị rõ ràng.
- Không được giả định mọi `view` trong runtime DB đã là Sơn Tây-only.

## Các gói dữ liệu đã chắc chắn có ở runtime Sơn Tây

Theo `v_tien_do_nap_bao_cao` tại ngày `2026-04-20`, các nhóm báo cáo đã nạp gồm:
- C1.1, C1.2, C1.3, C1.4 và các sheet chi tiết tương ứng
- GHTT Sơn Tây và GHTT NVKTDB
- KPI NVKT C11/C12/C13
- Kết quả tiếp thị
- MyTV hoàn công
- Phiếu hoàn công dịch vụ chi tiết
- Bộ dữ liệu tạm dừng/khôi phục dịch vụ
- Fiber thực tăng
- Xác minh và chi tiết xác minh
- Cấu hình tự động
- Vật tư thu hồi và quyết toán vật tư

## Khoảng trống dữ liệu hiện thời

Chưa thấy contract tương ứng trong SQLite runtime Sơn Tây cho:
- BRCD
- PTTB
- I1.5
- I1.5 K2
- SHC processing
- quang chủ động
- sự cố SA
- tồn kho vật tư tổng và tra cứu kho kiểu cũ
- thống kê ticket kiểu `dashv3`

Ngoài ra còn có các điểm chưa hoàn chỉnh:
- `C15` chưa có hợp đồng rõ như màn `dashv3`
- `v_dashboard_thuc_tang_moi_nhat` hiện mới thấy `Fiber`
- `MyTV` ở khối dịch vụ mới chỉ chắc cho `hoan_cong` trong `v_dashboard_dich_vu_theo_to_moi_nhat`
- `tam_dung_khoi_phuc` đã có dữ liệu gốc/snapshot nhưng chưa thấy `view` consumer ổn định tương ứng một-một với màn cũ
