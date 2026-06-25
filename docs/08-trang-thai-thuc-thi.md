# Trạng thái thực thi

## Mốc cập nhật

- Ngày cập nhật: `2026-06-25`
- Trạng thái: đã chuyển phần lớn route có dữ liệu phù hợp sang DB mới; các route chưa đủ contract đã bắt đầu bị ngắt khỏi nguồn Excel/file cũ; `/brcd` và `/pttb` đã có lớp kiểm soát tổ trưởng (write-aside) kèm snapshot lịch sử và timestamp báo cáo

## Những gì đã làm

### 1. Fork codebase

Đã copy khung `dashv3` sang `dashv4` để giữ:
- Flask app bootstrap
- blueprint
- template
- static JS/CSS
- auth cũ

Đã loại trừ các artifact runtime chính khi copy:
- `.git`
- `logs`
- `flask_session`
- `cache`
- `archive`
- `__pycache__`

## 2. Đổi nền cấu hình sang dashv4

Đã cập nhật:
- `config.py`
- `dashboard.py`
- `gunicorn_config.py`
- `start_dashboard.sh`
- thêm `start_dashv4.sh`

Các thay đổi nền:
- dùng biến môi trường `DASHV4_*`
- DB mặc định trỏ về:
  `/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db`
- trang chủ đổi sang `/tong-hop-bsc-kpi`
- background services mặc định tắt trong `dashv4`

## 3. Thêm data layer SQLite read-only

Đã tạo:
- [repositories/sqlite_runtime.py](/home/vtst/dashv4/repositories/sqlite_runtime.py)
- [repositories/dashboard_views.py](/home/vtst/dashv4/repositories/dashboard_views.py)

Nguyên tắc đã áp dụng:
- kết nối SQLite read-only
- query qua `view`
- không đọc Excel cho các route đã chuyển

## 4. Route đã chuyển sang DB mới

### Đã chạy từ `report_history.db`

- `/api/tong-hop-bsc-kpi`
- `/api/cau-hinh-tu-dong/son-tay`
- `/api/tiepthi-data`
- `/api/thu-hoi-data`
- `/api/giahan-ghtt-sty`
- `/api/giahan-ghtt-nvktdb`
- `/api/c1-chat-luong-data`
- `/api/c11-chi-tiet-data`
- `/api/c11-chi-tiet-15h-data`
- `/api/c11-chi-tiet-16h-data`
- `/api/c11-chi-tiet-17h-data`
- `/api/c11-chi-tiet-18h-data`
- `/api/c11-chi-tiet-quality-data`
- `/api/c12-repeat-failure-data`
- `/api/c14-chi-tiet-data`
- `/api/kpi-data`
- `/api/kpi-nvkt-bchn-data`
- `/api/ngungpsc-data`
- `/api/ngungpsc-mytv-data`
- `/api/ngungpsc-fiber-to-data`
- `/api/ngungpsc-mytv-to-data`
- `/api/ngungpsc-fiber-nvkt-data`
- `/api/ngungpsc-mytv-nvkt-data`
- `/api/xac-minh-tam-dung-data`

### Đã hỗ trợ lọc ngày chuẩn từ `report_history.db`

- `/api/c1-chat-luong-data`

Chi tiết triển khai hiện tại:
- route nhận `?date=YYYY-MM-DD`
- ngày chuẩn lấy từ `bao_cao_tong_hop_ngay.ngay_du_lieu`
- query qua bảng raw import + join `sheet_bao_cao_tong_hop` + `bao_cao_tong_hop_ngay`
- payload trả:
  - `selected_date`
  - `latest_available_date`
  - `available_dates`
  - `date_has_data`

Page HTML đã bật filter ngày tương ứng:
- `/chatluong`

### Ghi chú

- một số route dùng `consumer view` trực tiếp, ví dụ `ttvt`, `cau_hinh_tu_dong`, `giahan`, `kpi`
- một số route dùng adapter để giữ shape UI cũ, ví dụ `chatluong`, `kpi-nvkt-bchn`
- riêng `ngungpsc` đang ở hai mức:
  - Fiber dùng `v_dashboard_thuc_tang_moi_nhat`
  - MyTV đang fallback tổng hợp từ `v_dashboard_dich_vu_theo_to_moi_nhat` vì runtime hiện chưa có hàng dữ liệu tương ứng trong `v_thuc_tang_mytv_moi_nhat`

## 6. Kiểm soát tổ trưởng tại `/brcd`

> Tài liệu vận hành đầy đủ (schema, cron, backup, troubleshooting):
> [11-kiemsoat-to-truong-van-hanh.md](/home/vtst/dashv4/docs/11-kiemsoat-to-truong-van-hanh.md)

Đã thêm lớp "kiểm soát tổ trưởng" cho phiếu tồn BRCD (mã `baohong_id`):

- tổ trưởng nhập "nội dung kiểm soát" (1 ô tự do) trực tiếp trên bảng chi tiết phiếu tồn
- dữ liệu lưu trong SQLite ghi được per-instance `INSTANCE_RUNTIME_DIR/brcd_kiemsoat.db` (env `DASHV4_BRCD_KIEMSOAT_DB_PATH`), bảng `brcd_kiemsoat`
- tồn live vẫn đọc Excel read-only từ `1bss`; annotation lưu riêng (write-aside) để không bị `1bss` refresh ghi đè; annotation giữ lại làm lịch sử khi phiếu rời tồn
- endpoint mới:
  - `POST /api/brcd-kiemsoat/luu` (upsert theo `baohong_id`; xóa khi nội dung rỗng; ghi `nguoi_nhap` từ session)
  - `GET /api/brcd-kiemsoat/detail` (sheet đầy đủ `ToKT_<doi>` + JOIN annotation)
  - `GET /api/brcd-kiemsoat/thongke` (lọc theo quá giờ / trạng thái / đội; trả summary + by_doi + by_nvkt + chi_tiet + lich_su)
  - `GET /download/brcd-kiemsoat-report` (xuất Excel theo bộ lọc hiện hành)
- không thuộc `report_history.db` nên không áp dụng date-contract của tài liệu 09

### 6.1. Snapshot lịch sử phiếu (`brcd_phieu`)

- **Mục đích:** Tra cứu phiếu đã rời tồn; đánh giá kiểm soát tổ trưởng trên phiếu đã xử lý xong.
- **Cơ chế:** Upsert Vũ trụ tổng (Excel) vào bảng `brcd_phieu` mỗi lần load `/brcd` + mỗi giờ qua cron `scripts/sync_brcd_phieu.py`.
- **Schema:** 1 dòng mỗi `baohong_id`. `first_seen` giữ nguyên qua sync, `last_seen` update mỗi lần. Phiếu rời tồn → dòng giữ lại với `last_seen` cũ.
- **Metric mới trong `/api/brcd-kiemsoat/thongke`:** section `lich_su` với `roi_da_ks` (rời tồn + đã KS) và `roi_chua_ks` (rời tồn + chưa KS). Filter `khoang` ∈ tuan_nay/thang_nay/nam_nay/tat_ca.
- **Cron setup (per-instance):**
  ```
  0 * * * * DASHV4_UNIT_CODE=<unit> /home/vtst/dashv4/venv/bin/python3 \
      /home/vtst/dashv4/scripts/sync_brcd_phieu.py \
      >> /home/vtst/dashv4/logs/brcd_phieu_sync.log 2>&1
  ```

## 7. Kiểm soát tổ trưởng tại `/pttb`

> Tài liệu vận hành đầy đủ (schema, cron, backup, troubleshooting):
> [11-kiemsoat-to-truong-van-hanh.md](/home/vtst/dashv4/docs/11-kiemsoat-to-truong-van-hanh.md)

Đã thêm lớp "kiểm soát tổ trưởng" cho phiếu tồn PTTB (mã `ma_thue_bao`):

- tổ trưởng nhập "nội dung kiểm soát" (1 ô tự do) trực tiếp trên bảng chi tiết phiếu tồn PTTB
- dữ liệu lưu trong SQLite per-instance `INSTANCE_RUNTIME_DIR/brcd_kiemsoat.db` (cùng DB với BRCD kiêm soát), bảng `pttb_kiemsoat`
- tồn live vẫn đọc Excel read-only từ `1bss` (`baoCaoPTTB.xlsx`); annotation lưu riêng (write-aside) để không bị `1bss` refresh ghi đè
- endpoint mới:
  - `POST /api/pttb-kiemsoat/luu` (upsert theo `ma_thue_bao`; xóa khi nội dung rỗng; ghi `nguoi_nhap` từ session)
  - `GET /api/pttb-kiemsoat/detail` (sheet `ToKT_<to>` + JOIN annotation)
  - `GET /api/pttb-kiemsoat/thongke` (lọc theo quá giờ / trạng thái / tổ; trả summary + by_doi + by_nvkt + chi_tiet + lich_su)
  - `GET /download/pttb-kiemsoat-report` (xuất Excel theo bộ lọc hiện hành)
- không thuộc `report_history.db` nên không áp dụng date-contract của tài liệu 09

### 7.1. Snapshot lịch sử phiếu PTTB (`pttb_phieu`)

- **Mục đích:** Tra cứu phiếu PTTB đã rời tồn; đánh giá kiểm soát tổ trưởng trên phiếu đã xử lý xong.
- **Cơ chế:** Upsert Vũ trụ tổng (Excel) vào bảng `pttb_phieu` mỗi lần load `/pttb` + mỗi giờ qua cron `scripts/sync_pttb_phieu.py`.
- **Schema:** 1 dòng mỗi `ma_thue_bao`. `first_seen` giữ nguyên qua sync, `last_seen` update mỗi lần. Phiếu rời tồn → dòng giữ lại với `last_seen` cũ.
- **Metric mới trong `/api/pttb-kiemsoat/thongke`:** section `lich_su` với `roi_da_ks` (rời tồn + đã KS) và `roi_chua_ks` (rời tồn + chưa KS). Filter `khoang` ∈ tuan_nay/thang_nay/nam_nay/tat_ca.
- **Cron setup (per-instance):**
  ```
  0 * * * * DASHV4_UNIT_CODE=<unit> /home/vtst/dashv4/venv/bin/python3 \
      /home/vtst/dashv4/scripts/sync_pttb_phieu.py \
      >> /home/vtst/dashv4/logs/pttb_phieu_sync.log 2>&1
  ```

### 7.2. Timestamp thời gian file báo cáo

Phía trên bảng chi tiết tồn (cả `/brcd` và `/pttb`) hiển thị dòng timestamp
màu đỏ nổi bật:

> 🕐 **Dữ liệu báo cáo cập nhật:** `<file mtime>`

- Lấy từ `file_info.modified` của endpoint `detail` (= thời gian `1bss` tải
  file Excel về, phản ánh độ mới số liệu).
- Render trong `renderKiemSoatTable` (BRCD) và `renderPttbKiemSoatTable` (PTTB)
  qua `static/js/pages/brcd.js` và `static/js/pages/pttb.js`.
- CSS inline: `color:#d32f2f; font-weight:600`.

## 5. Route chưa hỗ trợ đã bị chặn ở mức page/API

Đã thêm page chung:
- [templates/pages/pending_feature.html](/home/vtst/dashv4/templates/pages/pending_feature.html)

Hiện các page sau trả trạng thái pending thay vì chạy nguồn cũ:
- `/thuctang`
- `/shc-processing`
- `/ton-kho-vat-tu`
- `/Tong_hop_tien`
- `/tam-dung-khoi-phuc`

Hai trang I1.5 đã được mở lại và đọc trực tiếp từ `report_history.db`:
- `/i15`
- `/i15k2`

Các API/download legacy tương ứng cũng đã được cấu hình trả `501` với payload:
- `error = legacy_endpoint_disabled`
- `title`
- `reason`
- `required_display_contract`

Mục tiêu của payload này là biến endpoint bị ngắt thành backlog dữ liệu rõ ràng cho `api_transition`, thay vì để Excel cũ tiếp tục là nguồn ngầm.

## Smoke test đã chạy

### Cú pháp Python

Đã chạy `py_compile` cho các file chính đã sửa, kết quả thành công.

### API smoke test

Đã xác nhận bằng Flask `test_client`:
- `/api/tong-hop-bsc-kpi` trả `200`
- `/api/cau-hinh-tu-dong/son-tay` trả `200`
- `/api/tiepthi-data` trả `200`
- `/api/kpi-data` trả `200`
- `/api/i15-data` trả `200`
- `/api/shc-variation-data` trả `200`
- `/api/shc-nvkt-detail/options` trả `200`
- `/api/shc-nvkt-detail/files` trả `200`
- `/api/shc-nvkt-detail/preview` trả `200`
- `/download/shc-nvkt-detail/...` trả `200`
- `/download/excel-i15` và `/download/excel-i15k2` trả `200`
  Hai nút tải tổng của `/i15` và `/i15k2` đã đổi sang export động từ payload DB theo `selected_date`, không còn trỏ về file Excel cũ ngoài thư mục download.
- `/api/i15k2-data` trả `200`
- `/api/shc-variation-k2-data` trả `200`
- `/api/thu-hoi-data` trả `200`
- `/api/giahan-ghtt-hni` trả `200`
- `/api/giahan-ghtt-sty` trả `200`
- `/api/giahan-ghtt-nvktdb` trả `200`
- `/api/tiepthi-data` trả `200`
- `/api/kpi-data` trả `200`
- `/api/kpi-nvkt-bchn-data` trả `200`
- `/api/c1-chat-luong-data` trả `200`
- `/api/c11-chi-tiet-data` trả `200`
- `/api/c12-repeat-failure-data` trả `200`
- `/api/ngungpsc-data` trả `200`
- `/api/ngungpsc-mytv-data` trả `200`
- `/api/ngungpsc-fiber-to-data` trả `200`
- `/api/ngungpsc-mytv-to-data` trả `200`
- `/api/ngungpsc-fiber-nvkt-data` trả `200`
- `/api/ngungpsc-mytv-nvkt-data` trả `200`
- `/api/xac-minh-tam-dung-data` trả `200`
- `/api/excel-data` trả `501`
- `/api/pttb-data-summary` trả `501`
- `/api/tam-dung-khoi-phuc-data` trả `200`
- `/api/cau-hinh-tu-dong/son-tay` trả `200`
- `/brcd` khi có session trả `200` (đã mở lại với lớp kiểm soát tổ trưởng)

## Những gì chưa làm

- chưa refactor adapter logic về `services/serializers`
- chưa có consumer view chính thức cho phần MyTV của `thuc-tang-ngung-psc`
- vẫn còn các nhóm ngoài phạm vi DB mới: `shc`, `statistics`
- đã có `/chatluong`, `/tam-dung-khoi-phuc`, `/thuc-tang-ngung-psc`, `/giahan`, `/cau-hinh-tu-dong`, `/tiepthi`, `/kpi`, `/i15` và `/i15k2` được chuyển sang chuẩn lọc ngày; các route còn lại vẫn chủ yếu đang đọc snapshot mới nhất

## Bước tiếp theo đề xuất

Thứ tự tiếp tục hợp lý:
1. chuyển phần adapter rải trong blueprint thành `services/serializers`
2. bổ sung contract DB cho `/tam-dung-khoi-phuc`
3. bổ sung `view` tổng hợp chính thức cho MyTV thực tăng/ngưng PSC
4. tiếp tục khóa nốt endpoint legacy nếu còn sót
5. thêm smoke test tự động cho các route đã migrated và các route `501`

## Tài liệu nguyên tắc mới

Đã bổ sung tài liệu chuẩn cho hướng triển khai lọc theo ngày từ `report_history.db`:
- `docs/09-nguyen-tac-loc-ngay-report-history.md`

Quy tắc bắt buộc từ mốc này:
- route nào được triển khai hỗ trợ `date` phải cập nhật doc cùng lúc
- tối thiểu phải cập nhật:
  - `docs/08-trang-thai-thuc-thi.md`
  - `docs/04-mapping-route-va-du-lieu.md`
  - và tài liệu nguyên tắc nếu có thay đổi chuẩn chung
