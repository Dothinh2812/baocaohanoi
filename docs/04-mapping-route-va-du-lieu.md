# Mapping route và dữ liệu

## Mục tiêu

Tài liệu này là bảng source-of-truth để map:
- route/page
- API dữ liệu đi kèm
- nguồn dữ liệu hiện tại trong `dashv4`
- trạng thái hỗ trợ hiển thị theo ngày

Từ mốc có tài liệu `docs/09-nguyen-tac-loc-ngay-report-history.md`, mọi route đọc `report_history.db` khi triển khai tới đâu phải cập nhật bảng này tới đó.

## Quy ước cột `supports_date`

Giá trị chuẩn:
- `yes`: route đã hỗ trợ `?date=YYYY-MM-DD` theo nguyên tắc chuẩn
- `planned`: route dùng `report_history.db`, có thể làm theo ngày nhưng chưa triển khai
- `no`: route hiện không hỗ trợ theo ngày
- `n/a`: route không dùng `report_history.db` hoặc không thuộc phạm vi lọc ngày hiện tại

## Quy ước cột `status`

Giá trị nên dùng:
- `compatible`
- `compatible-with-adapter`
- `partially-compatible`
- `disabled-pending-data`
- `legacy-non-report-history`

## Nguyên tắc cập nhật

Mỗi khi triển khai xong một route theo ngày, bắt buộc cập nhật tối thiểu:
- cột `supports_date`
- cột `nguồn dashv4 hiện tại`
- cột `ghi chú`

Nếu route đổi logic nguồn dữ liệu:
- cập nhật file này cùng lúc với code
- cập nhật thêm `docs/08-trang-thai-thuc-thi.md`

## Bảng mapping chính

| Route/Page | API chính | Nguồn `dashv3` | Nguồn `dashv4` hiện tại | supports_date | status | Ghi chú |
| --- | --- | --- | --- | --- | --- | --- |
| `/tong-hop-bsc-kpi` | `/api/tong-hop-bsc-kpi` | payload build tay từ dữ liệu trung gian | `v_dashboard_ttvt_son_tay_tong_hop_moi_nhat` | `planned` | `compatible` | Đang đọc snapshot mới nhất theo đơn vị cấu hình. Có thể chuyển sang by-date nếu map lại từ bảng raw nền. |
| `/cau-hinh-tu-dong` | `/api/cau-hinh-tu-dong/son-tay` | Excel/process cục bộ | 2 bảng raw import `cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet_th_theo_to`, `cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet_th_theo_nvkt` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã hỗ trợ `?date=YYYY-MM-DD`, summary cards và 2 bảng cùng đọc một `selected_date` chung. |
| `/tiepthi` | `/api/tiepthi-data` | Excel kết quả tiếp thị | 2 bảng raw import `kq_tiep_thi_kq_tiep_thi_report_kq_th`, `kq_tiep_thi_kq_tiep_thi_report_kq_tiep_thi` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã hỗ trợ `?date=YYYY-MM-DD`, đồng bộ cả bảng tổng hợp theo tổ và tab chi tiết theo nhân viên. |
| `/giahan` | `/api/giahan-ghtt-hni`, `/api/giahan-ghtt-sty`, `/api/giahan-ghtt-nvktdb` | Excel GHTT HNI/STY/NVKTDB | 3 bảng raw import `ghtt_ghtt_hni_report_kq_hni`, `ghtt_ghtt_sontay_report_kq_sontay`, `ghtt_ghtt_nvktdb_report_kq_nvktdb` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã hỗ trợ `?date=YYYY-MM-DD`, cả 3 API dùng chung một `selected_date` cho HNI, STY và NVKTĐB. |
| `/thuhoi` | `/api/thu-hoi-data` | file vật tư thu hồi | `v_thu_hoi_ui_tong_hop_moi_nhat`, `v_thu_hoi_ui_chi_tiet_moi_nhat` | `planned` | `compatible` | Nếu chuyển by-date phải join metadata từ bảng raw phía dưới. |
| `/chatluong` | `/api/c1-chat-luong-data` | Excel C1.* | 9 bảng raw import C1.* truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay`, render theo thứ tự 9 view `v_chi_tieu_c_*` | `yes` | `compatible-with-adapter` | Đã hỗ trợ `?date=YYYY-MM-DD`, trả `selected_date`, `latest_available_date`, `available_dates`, và không fallback ngầm khi ngày không có dữ liệu. |
| `/i15` | `/api/i15-data`, `/api/shc-variation-data`, `/api/shc-nvkt-detail/*` | DB/Excel SHC | 4 bảng raw import `chi_tieu_i_i1_5_report_bien_dong_tong_hop`, `chi_tieu_i_i1_5_report_shc_theo_sa`, `chi_tieu_i_i1_5_report_th_shc_i15`, `chi_tieu_i_i1_5_report_th_shc_theo_to` và các bảng cá nhân `chi_tieu_i_i1_5_report_{ten_nvkt}` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã chuyển `/i15` sang đọc trực tiếp dữ liệu K1 theo ngày; khối “chi tiết suy hao cao theo cá nhân” cũng đã đọc từ DB và export Excel động theo `selected_date`. |
| `/i15k2` | `/api/i15k2-data`, `/api/shc-variation-k2-data`, `/api/shc-nvkt-detail/*` | DB/Excel SHC | 4 bảng raw import `chi_tieu_i_i1_5_k2_report_bien_dong_tong_hop`, `chi_tieu_i_i1_5_k2_report_shc_theo_sa`, `chi_tieu_i_i1_5_k2_report_th_shc_i15`, `chi_tieu_i_i1_5_k2_report_th_shc_theo_to` và các bảng cá nhân `chi_tieu_i_i1_5_k2_report_{ten_nvkt}` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã chuyển `/i15k2` sang đọc trực tiếp dữ liệu K2 theo ngày; khối “chi tiết suy hao cao theo cá nhân” dùng chung API DB-backed với `/i15`. |
| `/thuc-tang-ngung-psc` | `/api/ngungpsc-data`, `/api/ngungpsc-mytv-data` | Excel nhiều sheet | 2 bảng raw import `tam_dung_khoi_phuc_dich_vu_ngung_psc_fiber_*`, `tam_dung_khoi_phuc_dich_vu_ngung_psc_mytv_*` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã hỗ trợ `?date=YYYY-MM-DD`, 2 API cùng resolve 1 `selected_date` chung cho cả Fiber và MyTV. |
| `/xac-minh-tam-dung` | `/api/xac-minh-tam-dung-data` | Excel xác minh | `v_xac_minh_ui_chi_tiet_moi_nhat` | `planned` | `compatible` | Đang nhóm theo đội VT trên snapshot mới nhất. |
| `/tam-dung-khoi-phuc` | `/api/tam-dung-khoi-phuc-data` | Excel processed | 2 bảng raw import `tam_dung_khoi_phuc_dich_vu_*_tong_hop_theo_to`, `tam_dung_khoi_phuc_dich_vu_*_tong_hop_theo_nvkt` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã hỗ trợ `?date=YYYY-MM-DD`, trả metadata ngày chuẩn và không fallback ngầm khi ngày không có dữ liệu. |
| `/kpi` | `/api/kpi-data` | Excel KPI | 3 bảng raw import `kpi_nvkt_c11_nvktdb_report_c11_kpi_nvkt`, `kpi_nvkt_c12_nvktdb_report_c12_kpi_nvkt`, `kpi_nvkt_c13_nvktdb_report_c13_kpi_nvkt` truy qua `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay` | `yes` | `compatible-with-adapter` | Đã thay adapter cũ phụ thuộc view lỗi bằng payload dựng trực tiếp từ 3 bảng KPI theo ngày; hỗ trợ `?date=YYYY-MM-DD`. |
| `/kpi-nvkt-bchn` | `/api/kpi-nvkt-bchn-data`, `/api/nvkt-chi-tiet/<slug>` | Excel KPI BCHN | `v_nvkt_tong_hop_da_nguon` | `planned` | `compatible-with-adapter` | Có thể hỗ trợ theo ngày nếu view/tables nguồn cho phép truy nguyên snapshot. |
| `/tong-hop-cap-to` | `/api/tong-hop-cap-to-data` | tổng hợp nội bộ | `v_don_vi_tong_hop_da_nguon` | `planned` | `compatible` | Route mới của `dashv4`, hiện chưa có by-date. |
| `/bsc-kpi-cac-to` | `/api/bsc-kpi-cac-to-data` | tổng hợp nội bộ | `v_chi_tieu_bsc_kpi_cac_to` | `planned` | `compatible` | Route mới của `dashv4`, hiện chưa có by-date. |
| `/thuctang` | `n/a` | ảnh/PNG + file info | chưa có nguồn DB 1:1 | `n/a` | `disabled-pending-data` | Chưa phải route report_history chuẩn. |
| `/brcd` | `/api/excel-data`, `/api/excel-data-main`, `/api/excel-data-pending`, `/api/brcd-kiemsoat/detail`, `/api/brcd-kiemsoat/thongke`, `/api/brcd-kiemsoat/luu`, `/download/brcd-kiemsoat-report` | Excel từ repo khác | Excel runtime `1bss`: `…/kq_dhsc/bc_BRCD.xlsx`, `…/chiaTheoDoi/chiTietBrcd5Doi.xlsx` (sheet đầy đủ `ToKT_<doi>`, khóa `baohong_id`); lớp "kiểm soát tổ trưởng" ghi vào SQLite per-instance `INSTANCE_RUNTIME_DIR/brcd_kiemsoat.db` (`DASHV4_BRCD_KIEMSOAT_DB_PATH`); snapshot `brcd_phieu` hằng giờ qua `scripts/sync_brcd_phieu.py` + on-load sync | `n/a` | `legacy-non-report-history` | Không thuộc phạm vi lọc ngày của `report_history.db`. Tồn live đọc Excel read-only; annotation kiểm soát lưu riêng (write-aside, keyed `baohong_id`) để không bị `1bss` refresh ghi đè; thống kê = LEFT JOIN tồn hiện tại với annotation, lọc theo quá giờ/trạng thái/đội, xuất Excel. Snapshot `brcd_phieu` lưu lịch sử phiếu đã rời tồn; endpoint `/api/brcd-kiemsoat/thongke` trả thêm `lich_su` với 2 metric 'rời tồn đã/chưa KS'. |
| `/pttb` | `/api/pttb-data-*`, `/api/pttb-kiemsoat/detail`, `/api/pttb-kiemsoat/thongke`, `/api/pttb-kiemsoat/luu`, `/download/pttb-kiemsoat-report` | Excel từ repo khác | Excel runtime `1bss`: `/home/vtst/1bss/runtime/default/downloads/ton_pttb/baoCaoPTTB.xlsx` (sheet `ToKT_<to>`, khóa `ma_thue_bao`); lớp "kiểm soát tổ trưởng" ghi vào SQLite per-instance `INSTANCE_RUNTIME_DIR/brcd_kiemsoat.db` (`DASHV4_BRCD_KIEMSOAT_DB_PATH`); snapshot `pttb_phieu` hằng giờ qua `scripts/sync_pttb_phieu.py` + on-load sync | `n/a` | `legacy-non-report-history` | Không thuộc phạm vi lọc ngày hiện tại. Tồn live đọc Excel read-only; annotation kiểm soát lưu riêng (write-aside, keyed `ma_thue_bao`); thống kê = LEFT JOIN tồn hiện tại với annotation, lọc theo quá giờ/trạng thái/tổ, xuất Excel. Snapshot `pttb_phieu` lưu lịch sử phiếu đã rời tồn; endpoint `/api/pttb-kiemsoat/thongke` trả thêm `lich_su` với 2 metric 'rời tồn đã/chưa KS'. |
| `/shc-processing` | `/api/shc-processing-report`, `/api/shc-nvkt-detail/*` | báo cáo SHC xử lý riêng | service/nguồn riêng ngoài contract chuẩn `report_history.db` | `n/a` | `legacy-non-report-history` | Chưa đưa vào chuẩn lọc ngày chung của tài liệu 09. |
| `/shc-cts` | `/api/shc-cts-kiemsoat/detail`, `/api/shc-cts-kiemsoat/thongke`, `/api/shc-cts-kiemsoat/luu`, `/download/shc-cts-kiemsoat-report` | báo cáo SHC CTS xử lý | Excel intraday `Bao_cao_tien_trinh_YYYYMMDD.xlsx` + SQLite per-instance `INSTANCE_RUNTIME_DIR/shc_cts.db` (`DASHV4_SHC_CTS_HISTORY_DB_PATH`), 2 bảng `shc_cts_tien_do` (snapshot) và `shc_cts_kiemsoat` (annotation tổ trưởng), grain = `(ngay_xu_ly, nvkt_db)` | `n/a` | `legacy-non-report-history` | Không thuộc `report_history.db`. Snapshot sync idempotent qua cron + lazy; annotation lưu write-aside keyed `(ngay_xu_ly, nvkt_db)`; today đọc Excel live, quá khứ đọc từ DB. |
| `/ton-kho-vat-tu` | `/api/ton-kho-vat-tu`, `/api/ton-kho-vat-tu-tot-thuong-dung` | `baocao-vattu` | Excel cũ | `n/a` | `legacy-non-report-history` | Không dùng `report_history.db`. |
| `/Tong_hop_tien` | page render trực tiếp | `baocao-vattu` | Excel cũ | `n/a` | `legacy-non-report-history` | Không dùng `report_history.db`. |
| `/tra-cuu-nhanh-vat-tu` | `/api/tra-cuu-nhanh-vat-tu` | `baocao-vattu` | Excel cũ, giữ nguyên logic `dashv3` | `n/a` | `legacy-non-report-history` | Không dùng `report_history.db`; template và JS đang giống hệt `dashv3`. |
| `/quangchudong` | `/api/quangchudong/*` | DB/service riêng | snapshot JSON từ `/home/vtst/do_chu_dong_api/runtime/current_off_snapshot.json` | `n/a` | `legacy-non-report-history` | Không thuộc `report_history.db`; dashboard đọc từ snapshot do `do_chu_dong_api` sinh ra. |
| `/su_co_sa` | `/api/su-co-sa/data` | DB riêng | SQLite `/home/vtst/1bss/runtime/default/sqlite/sa_outage.db` bảng `sa_outage_incidents` | `n/a` | `legacy-non-report-history` | Không thuộc `report_history.db`; hiển thị sự cố SA đang tồn và đã kết thúc trong ngày. |
| `/dao-tao-sat-hach` | `/api/training/*`, gồm `GET /api/training/questions`, `GET /api/training/questions/<version>`, `POST /api/training/questions/{validate,import,approve,reject,publish}`, `POST /api/training/exams/<id>/cancel`, `/download/training/*` | `n/a` | SQLite ghi được per-instance `DASHV4_TRAINING_DB_PATH` (mặc định `runtime_app/<unit>/training.db`), migration 1-11 | `n/a` | `compatible` | Module độc lập, không đọc/ghi `report_history.db` và không áp dụng date-contract. UI-2 hoàn tất panel ngân hàng câu hỏi cho editor/exam-manager: list phân trang/filter, detail DTO quản trị, validate/import draft, duyệt/từ chối/phát hành. Classification domain suy ra từ evidence -> knowledge block, không thêm migration. Route/API kiểm tra RBAC module theo scope server-side; learner chỉ dùng attempt DTO không có đáp án, giải thích, evidence hoặc metadata chấm. Chưa xác nhận UI-3+, production OpenAI hoặc thi lại. |

## Route ưu tiên triển khai `supports_date = yes`

Thứ tự khuyến nghị:
1. `/chatluong`
2. `/tam-dung-khoi-phuc`
3. `/thuc-tang-ngung-psc`
4. `/giahan`
5. `/cau-hinh-tu-dong`

Lý do:
- các route này đã đọc `report_history.db`
- phạm vi dữ liệu tương đối rõ
- đang có nhu cầu hiển thị snapshot theo ngày rõ nhất

## Route chưa nên triển khai lọc ngày ngay

### Nhóm `n/a`

Không làm theo tài liệu 09 cho đến khi chuyển nguồn:
- `/brcd`
- `/pttb`
- `/ton-kho-vat-tu`
- `/Tong_hop_tien`
- `/tra-cuu-nhanh-vat-tu`
- `/quangchudong`
- `/su_co_sa`
- `/dao-tao-sat-hach`

### Nhóm `planned` nhưng cần rà thêm

- `/kpi`
- `/kpi-nvkt-bchn`
- `/tong-hop-bsc-kpi`

Lý do:
- đang dùng view tổng hợp business hoặc adapter nhiều lớp
- cần xác định lại route -> report_code -> table_name trước khi thêm `date`

## Route đã hỗ trợ `date` theo chuẩn mới

### `/chatluong`

Đã hoàn tất các yêu cầu tối thiểu:
- nhận `?date=YYYY-MM-DD`
- backend resolve ngày từ `bao_cao_tong_hop_ngay.ngay_du_lieu`
- query dữ liệu qua bảng raw + `__sheet_id`
- trả `selected_date`, `latest_available_date`, `available_dates`
- frontend có bộ lọc ngày chung cho toàn page
- nếu ngày không có dữ liệu thì giữ đúng ngày được chọn, không fallback ngầm

## Checklist cập nhật khi một route được bật lọc ngày

Khi đổi `supports_date` từ `planned/no` sang `yes`, bắt buộc cập nhật:
- file này
- `docs/08-trang-thai-thuc-thi.md`
- `docs/09-nguyen-tac-loc-ngay-report-history.md` nếu có thay đổi nguyên tắc

Đồng thời route đó phải đạt:
- có nhận `?date=YYYY-MM-DD`
- có `selected_date` trong payload
- không còn phụ thuộc view `..._moi_nhat` làm nguồn chính
- đã smoke test với:
  - không truyền `date`
  - truyền `date` hợp lệ
  - truyền `date` không có dữ liệu

## Ghi chú thực tế tại thời điểm cập nhật

- Nhiều route hiện đã chuyển sang DB mới nhưng vẫn đang đọc snapshot “mới nhất”.
- `supports_date = planned` không có nghĩa là route đã gần xong; chỉ có nghĩa là route đang dùng `report_history.db` và nằm trong phạm vi chuẩn hóa tiếp theo.
- Sau mỗi lần triển khai, không được quên đổi cột `supports_date` và ghi chú tình trạng thực tế.
