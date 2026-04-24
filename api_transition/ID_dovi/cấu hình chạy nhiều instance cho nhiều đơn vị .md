# Cấu hình chạy nhiều instance cho nhiều đơn vị

Tài liệu thiết kế implementation chi tiết đi kèm:

- [thiết kế implementation chi tiết RuntimeContext và file config.md](</home/vtst/baocaohanoi/api_transition/ID_dovi/thiết kế implementation chi tiết RuntimeContext và file config.md>)

## Mục tiêu

Chạy nhiều instance từ cùng codebase `api_transition`, mỗi instance dùng một bộ cấu hình đầu vào riêng để tải cùng loại báo cáo cho các đơn vị khác nhau.

Yêu cầu chính:

- Không sửa code mỗi lần đổi đơn vị.
- Mỗi đơn vị có một file config riêng.
- Một loại báo cáo có thể cần một họ ID đơn vị khác với báo cáo khác.
- Tránh ghi đè lẫn nhau giữa các lần chạy của nhiều đơn vị.
- Có thể mở rộng dần từ mức download-only lên full pipeline.

## Nhận định hiện trạng

Codebase hiện tại chưa sẵn sàng cho multi-instance chỉ bằng cách đổi một `unit_id`.

Các điểm vướng chính:

- `batch_download.py` đang hard-code một bộ tham số toàn cục duy nhất.
- Các báo cáo đang dùng nhiều họ ID khác nhau:
  - `ptrungtamid` kiểu `14xxx` cho C1.x, KPI NVKT.
  - `vdonvi`, `vdv`, `pdonvi_id`, `vdonvi_id`, `vttvt`, `vdvvt` kiểu `28xxxx` cho nhiều báo cáo khác.
  - nhóm `ngưng PSC T-1` dùng `vdonvi_id` nhưng hiện đang nhận ID kiểu `14xxx`.
  - OneBSS dùng thêm các key riêng như `TT_ID`, `DOI_ID`, `vdonvi_id`, `vphanvung_id`, `vnhanvien_id`.
- `units.py` hiện mới mô hình hóa một phần nhỏ các loại ID.
- `downloads/`, `Processed/`, `ProcessedDaily/`, `report_history.db` đang dùng chung root mặc định.
- SQLite hiện ghi đè theo `ma_bao_cao + ngay_du_lieu`, nên nếu nhiều đơn vị cùng import vào một DB thì dữ liệu có thể đè nhau.
- Một số processor vẫn còn mang tính đặc thù Sơn Tây hoặc path cố định.

## Trạng thái triển khai hiện tại

### Đã xong

- Đã có bộ config draft cho `18` đơn vị trong `configs/units/`
- Đã có mapping `report -> ID family`
- Đã có validation cho bộ config draft
- Đã triển khai [runtime_config.py](/home/vtst/baocaohanoi/api_transition/runtime_config.py)
- Đã load thử thành công `18/18` file config
- Đã xác nhận mỗi config có thể tự tạo `instance_root` với đầy đủ:
  - `downloads`
  - `Processed`
  - `ProcessedDaily`
  - `sqlite_history`
- Đã nối [batch_download.py](/home/vtst/baocaohanoi/api_transition/batch_download.py) với config đơn vị
- Đã thêm `--config`
- Đã cho batch:
  - resolve `unit_id` theo `ID family`
  - resolve `output_dir` theo `instance_root`
  - skip report theo `enabled: false`
- Đã verify thật full batch cho `son_tay`:
  - `28` report thành công
  - `0` report thất bại
  - `5` report skip theo config
  - dữ liệu raw ghi đúng về `runtime/son_tay/downloads/...`
- Đã cho [processors/runner.py](/home/vtst/baocaohanoi/api_transition/processors/runner.py) chạy theo config đơn vị
- Đã cho processor:
  - đọc raw từ `runtime/<unit>/downloads/...`
  - ghi processed vào `runtime/<unit>/Processed/...`
  - skip processor theo `reports.<report_key>.enabled`
- Đã verify thật full processor runner cho `son_tay`:
  - `25` processor thành công
  - `0` processor thất bại
  - `5` processor skip theo config
  - dữ liệu processed ghi đúng về `runtime/son_tay/Processed/...`
- Đã nối [full_pipeline.py](/home/vtst/baocaohanoi/api_transition/full_pipeline.py) với config đơn vị
- Đã cho full pipeline:
  - archive sang `runtime/<unit>/ProcessedDaily/<snapshot_date>/...`
  - import vào `runtime/<unit>/sqlite_history/report_history.db`
- Đã verify thật Phase 4 cho `son_tay` bằng cách reuse raw đã có:
  - `25` processor thành công
  - `28` file archive trong `ProcessedDaily`
  - `28` workbook import thành công vào SQLite
  - `0` import thất bại
  - DB tạo đúng tại `runtime/son_tay/sqlite_history/report_history.db`
- Đã verify thật full end-to-end có network cho `son_tay`:
  - chạy `python3 -u -m api_transition.full_pipeline --config api_transition/configs/units/son_tay.yaml --reset-db`
  - `28` report download thành công
  - `0` report download thất bại
  - `25` processor thành công
  - `0` processor thất bại
  - `28` workbook archive thành công
  - `28` workbook import thành công vào SQLite
  - `0` import thất bại
- Đã đối soát DB `son_tay` với toàn bộ file processed:
  - `28` workbook processed khớp `28` dòng trong `bao_cao_ngay`
  - `99` sheet khớp `99` dòng trong `sheet_bao_cao`
  - `31,495` dòng raw trong workbook khớp `31,495` dòng trong `dong_bao_cao_goc`
  - kiểm tra `ma_hash_dong` cho toàn bộ `28` report cho kết quả `0` mismatch
- Đã query view dashboard thực tế:
  - `v_dashboard_chat_luong_don_vi_moi_nhat`
  - view hiện có `20` dòng dữ liệu cho `c11`, `c12`, `c13`, `c14`
  - dữ liệu đọc ra đúng ngày `2026-04-20` và đúng các đơn vị/tổ của Sơn Tây

### Chưa xong

- Chưa xử lý hardening/report đặc biệt của Phase 5

## Nguyên tắc thiết kế

- Chỉ có một codebase dùng chung cho tất cả instance.
- Không copy source code theo từng đơn vị.
- Mỗi instance chỉ khác nhau ở:
  - file config
  - thư mục dữ liệu runtime
  - service Linux tương ứng
- Tách cấu hình đơn vị ra file ngoài, không hard-code trong Python.
- Mỗi instance phải có một `instance_root` riêng theo tên đơn vị.
- Bên trong `instance_root` phải chứa đầy đủ toàn bộ dữ liệu runtime của đơn vị đó:
  - `downloads`
  - `Processed`
  - `ProcessedDaily`
  - `sqlite_history`
- File SQLite của đơn vị phải nằm trong `sqlite_history` của chính instance đó.
- Không để config phải khai báo trực tiếp theo key payload API nếu có thể tránh được.
- Chuẩn hóa theo "họ ID nghiệp vụ", sau đó map sang key thật của từng report.
- Hỗ trợ ngoại lệ per-report khi có báo cáo dùng semantics khác.

## Phương án triển khai

### Pha 1: Multi-instance an toàn, ít rủi ro theo `instance_root`

Mục tiêu của pha này là cho phép chạy nhiều đơn vị mà không đè file lên nhau, ưu tiên ổn định trước.

Thực hiện:

- Mỗi đơn vị có một file config riêng, ví dụ:
  - `configs/units/son_tay.yaml`
  - `configs/units/ha_dong.yaml`
- Mỗi lần chạy truyền `--config <path>`.
- Mỗi instance có một thư mục dữ liệu runtime riêng:
  - `runtime/<unit_code>/downloads`
  - `runtime/<unit_code>/Processed`
  - `runtime/<unit_code>/ProcessedDaily`
  - `runtime/<unit_code>/sqlite_history`
  - `runtime/<unit_code>/sqlite_history/report_history.db`
- Giữ nguyên tên file báo cáo hiện tại, chỉ đổi root thư mục theo instance.
- Dùng một SQLite riêng cho mỗi đơn vị, đặt trong `sqlite_history` của đơn vị đó.

Lợi ích:

- Không cần sửa toàn bộ processor cùng lúc.
- Không đụng sâu vào schema SQLite hiện tại.
- Có thể chạy tuần tự hoặc song song giữa các đơn vị mà không đè dữ liệu.
- Dữ liệu runtime của từng đơn vị được gom vào một chỗ, dễ backup, dọn dẹp và vận hành.

### Cấu trúc thư mục runtime đề xuất

```text
runtime/
  son_tay/
    downloads/
    Processed/
    ProcessedDaily/
    sqlite_history/
      report_history.db
      import_logs/
      exports/
  ha_dong/
    downloads/
    Processed/
    ProcessedDaily/
    sqlite_history/
      report_history.db
      import_logs/
      exports/
```

Lưu ý:

- Tất cả instance phải dùng chung cùng một source code `api_transition`.
- Không tạo bản sao code trong `runtime/<unit_code>/`.
- Thư mục `sqlite_history/` trong `runtime/<unit_code>/` chỉ chứa dữ liệu runtime.
- Source code và schema SQL của module `api_transition/sqlite_history/` vẫn giữ nguyên trong repo.
- Không copy code Python của `sqlite_history` vào từng instance.

### Pha 2: Chuẩn hóa cấu hình tham số theo report family

Mục tiêu của pha này là bỏ cách truyền `unit_id` chung chung, thay bằng profile ID có ý nghĩa theo nghiệp vụ.

Ý tưởng:

- Mỗi `ReportTask` được gán một `id_family`.
- Config đơn vị chỉ khai báo các ID ngữ nghĩa.
- Code resolver sẽ map `id_family` sang key payload thực tế của từng report.

Ví dụ `id_family`:

- `center_id_14`
- `unit_id_28`
- `service_scope_ttvt_id`
- `service_scope_team_id`
- `onebss_tt_id`
- `onebss_default_team_id`
- `onebss_region_id`

Lợi ích:

- Tránh nhồi toàn bộ key kỹ thuật như `ptrungtamid`, `vdonvi`, `pdonvi_id`, `vttvt` vào file config.
- Dễ đọc, dễ maintain, dễ kiểm tra chéo.
- Report nào đặc biệt vẫn có thể override riêng.

## Cấu trúc config đề xuất

```yaml
unit_code: son_tay
unit_name: TTVT Son Tay

runtime:
  instance_root: "runtime/son_tay"

period:
  report_month: 4
  report_year: 2026
  month_id: "98944548"
  month_label: ""
  vattu_start_date: "01/04/2025"

ids:
  center_id_14: "14324"
  unit_id_28: "284656"
  service_scope_ttvt_id: "14316"
  service_scope_team_id: "14324"
  onebss_tt_id: "14324"
  onebss_default_team_id: "0"
  onebss_region_id: "21"
  onebss_region_text: "Thành phố Hà Nội"
  onebss_unit_text: "Trung tâm Viễn thông Sơn Tây"

reports:
  ghtt_hni:
    enabled: false
  onebss_bc_chi_tiet_ket_qua_cskh_uc3:
    employee_id: "0"
    customer_batch_code: "UC3_CSKH_042026"
```

## Mapping cấu hình theo nhóm báo cáo

### Nhóm Chỉ tiêu C và KPI NVKT

- Dùng họ ID `center_id_14`.
- Map vào các key như:
  - `ptrungtamid`

Áp dụng cho:

- `download_report_c11_api`
- `download_report_c12_api`
- `download_report_c13_api`
- `download_kpi_nvkt_c11_api`
- `download_kpi_nvkt_c12_api`
- `download_kpi_nvkt_c13_api`

### Nhóm chi tiết C, I1.5, GHTT, xác minh, dịch vụ, vật tư

- Dùng họ ID `unit_id_28`.
- Map vào các key như:
  - `vdonvi`
  - `vdv`
  - `pdonvi_id`
  - `vdonvi_id`
  - `vttvt`
  - `vdvvt`

### Nhóm ngưng PSC T-1

- Không dùng cùng logic với `unit_id_28`.
- Cần tách rõ:
  - `service_scope_ttvt_id`
  - `service_scope_team_id`

Vì nhóm này hiện dùng `vdonvi_id` nhưng semantics đang khác với nhiều báo cáo khác.

### Nhóm OneBSS

- Cần profile riêng, tối thiểu:
  - `onebss_tt_id`
  - `onebss_default_team_id`
  - `onebss_region_id`
  - `onebss_region_text`
  - `onebss_unit_text`
- Một số report còn cần thêm:
  - `employee_id`
  - `customer_batch_code`

## Các thay đổi code cần làm

### 1. `batch_download.py`

- Bỏ phần cấu hình cứng ở đầu file.
- Thêm cơ chế đọc file config.
- Thêm resolver tham số theo từng `ReportTask`.
- Cho phép truyền `downloads_root` theo `instance_root` của đơn vị.

### 2. `downloaders.py`

- Giữ các hàm downloader như hiện tại để tránh vỡ API nội bộ.
- Nhưng phải cho phép caller truyền `output_dir` theo root runtime của instance.
- Dần bổ sung lớp wrapper nhận `unit profile` thay vì `unit_id` rời rạc.

### 3. `onebss_downloaders.py`

- Giữ downloader hiện có.
- Thêm lớp config adapter cho các tham số OneBSS đặc thù.

### 4. `processors/common.py`

- Bỏ phụ thuộc cứng vào:
  - `api_transition/downloads`
  - `api_transition/Processed`
- Chuyển sang root được truyền từ `instance_root` hoặc context runtime.

### 5. `processors/runner.py`

- Cho phép chạy processor với:
  - `download_root`
  - `processed_root`
- Không nên ngầm phụ thuộc vào layout mặc định của repo.

### 6. `full_pipeline.py`

- Bổ sung đầy đủ runtime roots theo instance:
  - `instance_root`
  - `download_root`
  - `processed_root`
  - `archive_root`
  - `sqlite_root`
  - `db_path`
- Khi chạy `run_batch_download()` và `run_all_processors()` phải truyền cùng một runtime context.

### 7. `units.py`

- Không nên tiếp tục là nguồn cấu hình chính.
- Có thể giữ như helper lookup tạm thời.
- Về lâu dài nên thay bằng thư mục `configs/units/`.

## Mô hình production với Linux service

### Mục tiêu vận hành

Mỗi đơn vị sẽ có một Linux service riêng.

Khi `start` service:

1. service load file config đã chỉ định
2. build `RuntimeContext`
3. chạy full pipeline:
   - download raw
   - process workbook
   - archive `ProcessedDaily`
   - import SQLite
4. ghi kết quả vào DB của chính instance
5. thoát với exit code phù hợp

### Nguyên tắc production

- `1 service = 1 config = 1 instance_root`
- mọi service dùng chung một codebase
- chỉ khác file config và thư mục runtime
- service phải chạy headless, non-interactive
- log phải đi ra `stdout/stderr` để Linux journal thu được
- service phải trả về exit code chuẩn để systemd nhận biết thành công/thất bại

### Mô hình triển khai khuyến nghị

Nếu bài toán là chạy theo lịch và mỗi lần chạy hoàn tất rồi thoát, nên dùng:

- `systemd service` kiểu chạy một lượt
- kết hợp `systemd timer`

Ví dụ:

- `api-transition@son_tay.service`
- `api-transition@ha_dong.service`

Mỗi instance service sẽ gọi cùng một entrypoint code, chỉ khác file config:

```ini
ExecStart=/usr/bin/python3 -m api_transition.full_pipeline --config /opt/api_transition/configs/units/%i.yaml
```

### Các yêu cầu cần bổ sung vào implementation

- `full_pipeline.py` phải hỗ trợ `--config`
- pipeline phải chạy được hoàn toàn trong môi trường service headless
- cần có lock theo instance để tránh 2 tiến trình cùng chạy trên một `instance_root`
- cần có log rõ ràng theo từng bước để quan sát qua `journalctl`
- cần có exit code non-zero nếu fail để service monitoring hoạt động đúng

### Kết luận production

Thiết kế hiện tại phù hợp với mô hình Linux service nếu giữ đúng nguyên tắc:

- một codebase chung
- config riêng cho từng đơn vị
- data runtime riêng cho từng đơn vị
- không nhân bản source code theo instance

## Rủi ro cần lưu ý

### Rủi ro 1: Đè file giữa nhiều instance

Nếu chỉ đổi `unit_id` mà không đổi root thư mục, các file raw và processed sẽ đè nhau.

### Rủi ro 2: Đè snapshot trong SQLite

Nếu nhiều đơn vị dùng chung một `report_history.db`, dữ liệu có thể bị ghi đè vì khóa hiện tại chưa chứa thông tin đơn vị.

Vì vậy trong thiết kế này, mỗi đơn vị phải có DB riêng trong:

- `runtime/<unit_code>/sqlite_history/report_history.db`

### Rủi ro 3: Processor chưa đủ generic

Một số processor còn assumptions theo Sơn Tây hoặc theo path mặc định. Vì vậy không nên hứa multi-unit end-to-end cho toàn bộ báo cáo ngay ở pha đầu.

### Rủi ro 4: Semantics ID không đồng nhất

Cùng tên tham số `unit_id` ở Python nhưng thực tế không phải lúc nào cũng là cùng loại ID. Nếu không chuẩn hóa bằng `id_family` thì sẽ rất dễ cấu hình sai.

## Khuyến nghị triển khai thực tế

Thứ tự nên làm:

1. Làm `config loader` và chuẩn hóa `instance_root` theo từng đơn vị.
2. Từ `instance_root`, suy ra:
   - `downloads_root`
   - `processed_root`
   - `archive_root`
   - `sqlite_root`
   - `db_path`
3. Sửa `batch_download.py` để chạy được download nhiều đơn vị bằng config.
4. Sửa `full_pipeline.py` và processor roots để toàn bộ pipeline dùng cùng runtime context của instance.
5. Tạm dùng SQLite riêng cho từng đơn vị.
6. Bổ sung chế độ chạy production qua Linux service:
   - `--config`
   - logging
   - exit code
   - instance lock
7. Sau khi ổn định mới refactor sâu processor và xem xét DB hợp nhất nhiều đơn vị.

## Kết luận

Phương án phù hợp nhất hiện tại là:

- triển khai multi-instance theo từng file config đơn vị
- dùng chung một codebase cho tất cả instance
- tách toàn bộ dữ liệu runtime theo `instance_root/<unit_code>/`
- dùng schema config theo "họ ID" thay vì một `unit_id` duy nhất
- rollout theo 2 pha để giảm rủi ro

Pha 1 đủ để chạy thực tế an toàn.
Pha 2 mới là bước chuẩn hóa dài hạn cho toàn bộ codebase.
