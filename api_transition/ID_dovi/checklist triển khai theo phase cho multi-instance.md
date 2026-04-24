# Checklist triển khai theo phase cho multi-instance

Tài liệu này là checklist triển khai ngắn để đi từ:

- plan chiến lược
- thiết kế `RuntimeContext`
- bộ draft config 18 đơn vị

sang implementation thật trong codebase `api_transition`.

Tài liệu gốc liên quan:

- [cấu hình chạy nhiều instance cho nhiều đơn vị .md](</home/vtst/baocaohanoi/api_transition/ID_dovi/cấu hình chạy nhiều instance cho nhiều đơn vị .md>)
- [thiết kế implementation chi tiết RuntimeContext và file config.md](</home/vtst/baocaohanoi/api_transition/ID_dovi/thiết kế implementation chi tiết RuntimeContext và file config.md>)

## Nguyên tắc triển khai

- Không refactor toàn bộ codebase trong một lượt.
- Ưu tiên làm cho chạy được theo `instance_root` trước.
- Tách rõ:
  - phase nào chỉ chạm vào config và runtime path
  - phase nào bắt đầu chạm downloader
  - phase nào mới nối full pipeline
- Mỗi phase phải có tiêu chí hoàn thành rõ ràng.

## Trạng thái hiện tại

### Đã có

- Hai file plan tổng và design chi tiết.
- Bộ config draft cho `18` đơn vị trong `configs/units/`.
- Mapping `report -> ID family`.
- Validation draft config.
- Đã triển khai `api_transition/runtime_config.py`.
- Đã export `RuntimeContext` và `load_runtime_context` từ `api_transition.__init__`.
- Đã load thử thành công `18/18` file config draft.
- Đã xác nhận `instance_root` được tạo đúng theo cấu trúc runtime của từng đơn vị.
- Đã nối `batch_download.py` với `RuntimeContext`.
- Đã thêm CLI `--config`.
- Đã resolve `unit_id` theo `id_family` trong `ReportTask`.
- Đã cho batch ghi `output_dir` theo `runtime/<unit>/downloads/<group>/...`
- Đã cho batch skip report theo `reports.<report_key>.enabled`.
- Đã nối `processors/runner.py` với `RuntimeContext`.
- Đã cho processor đọc raw từ `runtime/<unit>/downloads/...`
- Đã cho processor ghi processed vào `runtime/<unit>/Processed/...`
- Đã verify thật processor runner cho `son_tay`:
  - `25` processor thành công
  - `0` processor thất bại
  - `5` processor skip theo config
- Đã nối `full_pipeline.py` với `RuntimeContext`
- Đã cho full pipeline archive vào `runtime/<unit>/ProcessedDaily/...`
- Đã cho full pipeline import vào `runtime/<unit>/sqlite_history/report_history.db`
- Đã verify thật Phase 4 cho `son_tay`:
  - `28` workbook archive thành công
  - `28` workbook import thành công vào SQLite

### Chưa có

- Hardening và report đặc biệt của Phase 5.
- Đối soát sâu dữ liệu nghiệp vụ sau import cho nhiều đơn vị ngoài `son_tay`.

## Phase 1: RuntimeContext và config loader

### Mục tiêu

Đưa cấu hình đơn vị từ YAML vào code, build được đầy đủ runtime roots từ `instance_root`, nhưng chưa cần thay đổi luồng download/process toàn diện.

### File dự kiến sửa / tạo

- tạo mới `api_transition/runtime_config.py`
- có thể tạo thêm `api_transition/runtime_paths.py` nếu muốn tách nhỏ
- cập nhật nhẹ `api_transition/__init__.py` nếu cần export helper

### Việc cần làm

- Tạo dataclass:
  - `RuntimePaths`
  - `UnitProfile`
  - `PeriodConfig`
  - `DownloadConfig`
  - `RuntimeContext`
- Viết hàm:
  - `load_runtime_context(config_path)`
  - `validate_runtime_config(raw)`
  - `ensure_runtime_dirs(paths)`
- Từ `instance_root` sinh ra:
  - `downloads_root`
  - `processed_root`
  - `archive_root`
  - `sqlite_root`
  - `sqlite_db_path`
  - `sqlite_import_logs_root`
  - `sqlite_exports_root`
- Hỗ trợ đọc các file hiện có trong `configs/units/*.yaml`
- Validation tối thiểu:
  - đủ `unit.code`
  - đủ `runtime.instance_root`
  - đủ `period.report_month`, `period.report_year`
  - đủ `ids.center_id_14`, `ids.unit_id_28` với draft hiện tại

### Tiêu chí hoàn thành

- Load được `configs/units/son_tay.yaml`
- Sinh được toàn bộ path tuyệt đối
- Tự tạo thư mục runtime nếu `create_dirs: true`
- Có thể in/inspect `RuntimeContext` bằng một script test nhỏ hoặc CLI tạm

### Trạng thái triển khai

- `DONE` cho vòng đầu.
- Đã có file: [runtime_config.py](/home/vtst/baocaohanoi/api_transition/runtime_config.py)
- Đã kiểm tra:
  - load config qua package root
  - compile syntax
  - tạo thư mục runtime thực tế cho `18` đơn vị
- Chưa có CLI riêng cho loader, nhưng API dùng được:
  - `from api_transition import load_runtime_context`

### Chưa làm trong phase này

- Chưa nối vào `batch_download.py`
- Chưa đổi downloader
- Chưa đổi processor

## Phase 2: Batch download chạy theo config

### Mục tiêu

Cho `batch_download.py` chạy theo `--config`, resolve `unit_id` đúng theo `ID family`, và ghi raw vào `runtime/<unit>/downloads/...`.

### File dự kiến sửa

- `api_transition/batch_download.py`
- `api_transition/downloaders.py`
- `api_transition/units.py` nếu cần giữ làm compatibility helper

### Việc cần làm

- Thêm CLI:
  - `--config <path>`
- Khi có `--config`:
  - load `RuntimeContext`
  - dùng `period` từ config làm mặc định
  - dùng `download` config làm mặc định
- Mở rộng `ReportTask` để biết:
  - `report_key`
  - `id_family`
  - `group`
- Viết resolver:
  - `center_id_14` -> truyền vào các report C/KPI
  - `unit_id_28` -> truyền vào các report detail/I/GHTT/xác minh/dịch vụ/vật tư
- Với report `enabled: false` thì batch tự skip
- Với `group_output_dir()`:
  - không dùng root mặc định nếu đang chạy theo context
  - build output từ `context.paths.downloads_root / <group>`

### Tiêu chí hoàn thành

- Chạy được batch cho 1 đơn vị bằng:
  - `python -m api_transition.batch_download --config configs/units/son_tay.yaml --only "C1.1"`
- Chạy tiếp được cho đơn vị khác mà không đè file:
  - `ha_dong`, `ba_dinh`, ...
- Raw file được ghi đúng vào:
  - `runtime/<unit>/downloads/<group>/...`

### Trạng thái triển khai

- `DONE` cho vòng đầu của batch download
- Đã triển khai trong [batch_download.py](/home/vtst/baocaohanoi/api_transition/batch_download.py):
  - `--config`
  - load `RuntimeContext`
  - `report_key`
  - `id_family`
  - resolve `unit_id`
  - resolve `output_dir`
  - skip theo `enabled: false`
- Đã verify không cần mạng:
  - `--list --config ...` hiển thị đúng `enabled/disabled`
  - kwargs build đúng cho các report mẫu
  - report bị disable được skip sớm mà không cố login
- Đã verify download thật qua network:
  - chạy full batch cho `son_tay`
  - `28` report thành công
  - `0` report thất bại
  - `5` report bị skip theo config
  - raw file được ghi đúng dưới `runtime/son_tay/downloads/...`

### Chưa làm trong phase này

- Chưa nối processor theo runtime root
- Chưa chạy full pipeline end-to-end
- Chưa bật các report đang `enabled: false`

## Phase 3: Processors chạy theo instance root

### Mục tiêu

Cho processor đọc raw từ `runtime/<unit>/downloads` và ghi processed vào `runtime/<unit>/Processed`.

### File dự kiến sửa

- `api_transition/processors/common.py`
- `api_transition/processors/runner.py`
- các processor đang hard-code path mặc định, ưu tiên:
  - `service_flow_processors.py`
  - rồi đến các module còn lại nếu cần

### Việc cần làm

- Bỏ phụ thuộc cứng vào:
  - `api_transition/downloads`
  - `api_transition/Processed`
- Truyền `download_root` và `processed_root` từ `RuntimeContext`
- Với các processor đang nhận `input_path`:
  - giữ tương thích ngược
  - nhưng bổ sung khả năng resolve từ runtime root
- Với các processor hard-code path nội bộ:
  - thay bằng helper từ `context` hoặc root truyền vào

### Tiêu chí hoàn thành

- Chạy được:
  - download một số report
  - process ra `runtime/<unit>/Processed`
- Không còn tạo file processed trong `api_transition/Processed` khi đang chạy theo config

### Trạng thái triển khai

- `DONE` cho vòng đầu của processor runtime
- Đã triển khai trong:
  - [processors/common.py](/home/vtst/baocaohanoi/api_transition/processors/common.py)
  - [processors/runner.py](/home/vtst/baocaohanoi/api_transition/processors/runner.py)
  - [processors/service_flow_processors.py](/home/vtst/baocaohanoi/api_transition/processors/service_flow_processors.py)
- Đã bổ sung:
  - runtime roots có thể cấu hình cho processor common
  - `--config` cho processor runner
  - runtime path injection cho các processor đọc một hoặc nhiều raw file
  - skip processor theo `reports.<report_key>.enabled`
- Đã verify thật:
  - `python3 -m api_transition.processors.runner --config api_transition/configs/units/son_tay.yaml`
  - kết quả `25 success`, `0 failed`, `5 skipped`
  - output được ghi đúng vào `runtime/son_tay/Processed/...`

### Chưa làm trong phase này

- Chưa hoàn tất full pipeline
- Chưa generic hóa mọi processor đặc thù Sơn Tây ngoài phạm vi wiring path/runtime

## Phase 4: Full pipeline theo instance

### Mục tiêu

Nối `batch_download`, `processors`, archive, SQLite import vào cùng một `RuntimeContext`.

### File dự kiến sửa

- `api_transition/full_pipeline.py`
- `api_transition/sqlite_history/import_processed_to_sqlite.py`
- `api_transition/sqlite_history/init_report_history_db.py`
- `api_transition/sqlite_history/apply_report_history_views.py`

### Việc cần làm

- Thêm `--config` vào `full_pipeline.py`
- Khi có config:
  - dùng `context.paths.processed_root`
  - dùng `context.paths.archive_root`
  - dùng `context.paths.sqlite_db_path`
- Đảm bảo DB của mỗi instance nằm ở:
  - `runtime/<unit>/sqlite_history/report_history.db`
- Archive processed về:
  - `runtime/<unit>/ProcessedDaily/...`

### Tiêu chí hoàn thành

- Chạy được full pipeline cho 1 đơn vị
- Sinh đúng đủ:
  - `downloads`
  - `Processed`
  - `ProcessedDaily`
  - `sqlite_history/report_history.db`
- Không chạm vào DB mặc định ở source tree khi chạy theo config

### Chưa làm trong phase này

- Chưa hợp nhất nhiều đơn vị vào chung một SQLite DB

### Trạng thái triển khai

- `DONE` cho vòng đầu của full pipeline runtime
- Đã triển khai trong:
  - [full_pipeline.py](/home/vtst/baocaohanoi/api_transition/full_pipeline.py)
- Đã bổ sung:
  - `--config` cho full pipeline
  - kế thừa `period`, `download`, `processed_root`, `archive_root`, `db_path` từ `RuntimeContext`
  - truyền chung một `RuntimeContext` vào download và processor stage
  - archive sang `runtime/<unit>/ProcessedDaily/<snapshot_date>/...`
  - import vào `runtime/<unit>/sqlite_history/report_history.db`
- Đã tận dụng sẵn khả năng override path của:
  - [sqlite_history/import_processed_to_sqlite.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/import_processed_to_sqlite.py)
  - [sqlite_history/init_report_history_db.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/init_report_history_db.py)
  - [sqlite_history/apply_report_history_views.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/apply_report_history_views.py)
- Đã verify thật cho `son_tay` bằng cách skip download stage và reuse raw có sẵn:
  - `0` download thành công
  - `25` processor thành công
  - `28` workbook được archive
  - `28` workbook được import vào SQLite
  - `0` import thất bại
- Kết quả kiểm tra sau chạy:
  - `runtime/son_tay/ProcessedDaily` có `28` file
  - `runtime/son_tay/sqlite_history/report_history.db` được tạo thành công
  - DB có:
    - `28` dòng `danh_muc_bao_cao`
    - `28` dòng `bao_cao_ngay`
    - `99` dòng `sheet_bao_cao`
    - `30004` dòng `dong_bao_cao_goc`
- Đã verify thật end-to-end có network trong cùng một lệnh:
  - `python3 -u -m api_transition.full_pipeline --config api_transition/configs/units/son_tay.yaml --reset-db`
  - kết quả:
    - `28` download thành công
    - `0` download thất bại
    - `25` processor thành công
    - `0` processor thất bại
    - `28` workbook archive thành công
    - `28` workbook import thành công
    - `0` import thất bại
- Đã đối soát DB `son_tay` với file processed:
  - `28` workbook khớp `28` dòng `bao_cao_ngay`
  - `99` sheet khớp `99` dòng `sheet_bao_cao`
  - `31,495` dòng raw khớp `31,495` dòng `dong_bao_cao_goc`
  - `0` mismatch ở mức `ma_hash_dong`
- Đã kiểm tra view:
  - `v_dashboard_chat_luong_don_vi_moi_nhat`
  - có `20` dòng dữ liệu thực tế, đọc đúng cho `c11/c12/c13/c14`
- Đã tạo admin utility:
  - [sqlite_history/sync_all_instance_dbs.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/sync_all_instance_dbs.py)
  - đã verify:
    - `--mode status --unit son_tay`
    - `--mode apply-views --unit son_tay`

## Phase 5: Report đặc biệt và hardening

### Mục tiêu

Xử lý các report đang để `enabled: false` và làm cứng hệ thống để chạy production nhiều instance.

### Nhóm việc

- Xác minh mapping cho:
  - `ghtt_hni`
  - `ngung_psc_*`
  - OneBSS reports
- Thêm logging rõ hơn theo `unit.code`
- Thêm lock file theo instance:
  - `.pipeline.lock`
- Thêm validate config chặt hơn
- Thêm smoke test cho:
  - load config
  - build context
  - resolve report params

### Tiêu chí hoàn thành

- Bật được từng report đặc biệt khi đã có mapping xác minh
- Có thể chạy nhiều service theo nhiều đơn vị mà không đè dữ liệu nhau

## Thứ tự triển khai khuyến nghị

1. Phase 1
2. Phase 2
3. Test thực tế 2 đơn vị với vài report
4. Phase 3
5. Phase 4
6. Chỉ sau đó mới làm Phase 5

## Definition of Done cho vòng 1

Vòng 1 được coi là xong khi:

- Có thể chạy bằng `--config configs/units/<unit>.yaml`
- Dữ liệu của mỗi đơn vị nằm trọn dưới `runtime/<unit>/`
- Batch download chạy đúng ID theo từng report family
- Processor và full pipeline dùng đúng runtime roots của instance
- SQLite tách riêng theo từng đơn vị
- Không cần sửa code để đổi giữa `Sơn Tây`, `Hà Đông`, `Ba Đình`, ...
