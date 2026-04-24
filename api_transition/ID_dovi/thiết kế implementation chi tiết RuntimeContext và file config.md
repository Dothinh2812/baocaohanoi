# Thiết kế implementation chi tiết RuntimeContext và file config

## 1. Mục tiêu của bản thiết kế

Tài liệu này chốt thiết kế implementation cho bài toán chạy nhiều instance theo nhiều đơn vị từ cùng codebase `api_transition`.

Mục tiêu kỹ thuật:

- Mỗi lần chạy dùng đúng một file config đơn vị.
- Tất cả instance dùng chung một source code duy nhất.
- Mỗi đơn vị có một `instance_root` riêng.
- Toàn bộ dữ liệu runtime của đơn vị nằm dưới `instance_root`.
- Các downloader, processor, archive và SQLite cùng dùng một `RuntimeContext` thống nhất.
- Không còn phụ thuộc mặc định vào `api_transition/downloads`, `Processed`, `ProcessedDaily`, `report_history.db`.
- Không bắt người vận hành phải hiểu các key payload API mức thấp.

Không nằm trong phạm vi của bản thiết kế này:

- Hợp nhất nhiều đơn vị vào chung một SQLite DB.
- Refactor toàn bộ processor đặc thù Sơn Tây thành generic ngay trong vòng đầu.
- Thay đổi schema SQLite hiện tại.

## 1.1. Trạng thái implementation hiện tại

### Đã triển khai

- Đã tạo [runtime_config.py](/home/vtst/baocaohanoi/api_transition/runtime_config.py)
- Đã có các dataclass:
  - `RuntimePaths`
  - `UnitProfile`
  - `PeriodConfig`
  - `DownloadConfig`
  - `RuntimeContext`
- Đã có các hàm:
  - `build_runtime_paths(...)`
  - `ensure_runtime_dirs(...)`
  - `validate_runtime_config(...)`
  - `load_runtime_context(...)`
- Đã export:
  - `RuntimeContext`
  - `load_runtime_context`
  từ [__init__.py](/home/vtst/baocaohanoi/api_transition/__init__.py)
- Đã load thử thành công `18/18` file config trong `configs/units/`
- Đã xác nhận loader tự tạo cây:
  - `downloads`
  - `Processed`
  - `ProcessedDaily`
  - `sqlite_history`
  dưới `instance_root` của từng đơn vị
- Đã nối [batch_download.py](/home/vtst/baocaohanoi/api_transition/batch_download.py) với `RuntimeContext`
- Đã thêm CLI:
  - `--config <path>`
- Đã bổ sung vào `ReportTask`:
  - `report_key`
  - `id_family`
- Đã cho batch:
  - resolve `unit_id` theo `id_family`
  - resolve `output_dir` theo `runtime/<unit>/downloads/<group>`
  - skip report theo `reports.<report_key>.enabled`
- Đã verify không cần mạng:
  - list report theo config
  - build kwargs đúng theo config
  - skip sớm report đang disable
- Đã verify thật qua network cho:
  - `python3 -m api_transition.batch_download --config api_transition/configs/units/son_tay.yaml`
- Kết quả verify thật:
  - `28` report thành công
  - `0` report thất bại
  - `5` report skip theo config
  - raw outputs được ghi đúng dưới `runtime/son_tay/downloads/...`
- Đã xác nhận retry hoạt động trên report chậm:
  - `C1.2 Chi tiết SM2`
  - `Vật tư thu hồi`
- Đã nối [processors/runner.py](/home/vtst/baocaohanoi/api_transition/processors/runner.py) với `RuntimeContext`
- Đã thêm CLI:
  - `python3 -m api_transition.processors.runner --config <path>`
- Đã cập nhật [processors/common.py](/home/vtst/baocaohanoi/api_transition/processors/common.py) để cấu hình được `downloads_root` và `processed_root` theo instance
- Đã vá [service_flow_processors.py](/home/vtst/baocaohanoi/api_transition/processors/service_flow_processors.py) ở chỗ sinh processed output theo runtime root
- Đã cho processor:
  - đọc raw từ `runtime/<unit>/downloads/...`
  - ghi processed vào `runtime/<unit>/Processed/...`
  - skip processor theo `reports.<report_key>.enabled`
- Đã verify thật:
  - `python3 -m api_transition.processors.runner --config api_transition/configs/units/son_tay.yaml`
- Kết quả verify thật cho processor:
  - `25` processor thành công
  - `0` processor thất bại
  - `5` processor skip theo config
  - outputs được ghi đúng dưới `runtime/son_tay/Processed/...`
- Đã nối [full_pipeline.py](/home/vtst/baocaohanoi/api_transition/full_pipeline.py) với `RuntimeContext`
- Đã thêm CLI:
  - `python3 -m api_transition.full_pipeline --config <path>`
- Đã cho full pipeline:
  - kế thừa `period` và `download` defaults từ config đơn vị
  - dùng `runtime/<unit>/Processed` làm `processed_root`
  - dùng `runtime/<unit>/ProcessedDaily` làm `archive_root`
  - dùng `runtime/<unit>/sqlite_history/report_history.db` làm `db_path`
  - truyền chung `RuntimeContext` vào download và processor stage
- Đã verify thật phần archive + SQLite import cho `son_tay` bằng full pipeline:
  - skip toàn bộ download stage để reuse raw đã có
  - `25` processor thành công
  - `28` workbook được archive
  - `28` workbook được import vào SQLite
  - `0` import thất bại
- Kết quả kiểm tra sau verify:
  - `runtime/son_tay/ProcessedDaily` có `28` file
  - DB tạo tại `runtime/son_tay/sqlite_history/report_history.db`
  - DB có `28` dòng `bao_cao_ngay`
- Đã verify thật end-to-end có network trong cùng một lệnh:
  - `python3 -u -m api_transition.full_pipeline --config api_transition/configs/units/son_tay.yaml --reset-db`
- Kết quả verify end-to-end:
  - `28` download thành công
  - `0` download thất bại
  - `25` processor thành công
  - `0` processor thất bại
  - `28` workbook được archive
  - `28` workbook được import vào SQLite
  - `0` import thất bại
- Đã đối soát DB với processed của `son_tay`:
  - `28` workbook processed khớp `28` dòng `bao_cao_ngay`
  - `99` sheet khớp `99` dòng `sheet_bao_cao`
  - `31,495` dòng raw khớp `31,495` dòng `dong_bao_cao_goc`
  - `bao_cao_ngay.so_dong_goc`, `so_dong_tong_hop`, `so_dong_chi_tiet` đều khớp với workbook
  - kiểm tra `ma_hash_dong` giữa workbook parsed và `dong_bao_cao_goc` cho `28` report cho kết quả `0` mismatch
- Đã kiểm tra một view dashboard thực tế:
  - `v_dashboard_chat_luong_don_vi_moi_nhat`
  - view có `20` dòng, là `UNION ALL` của `v_c11_tong_hop_moi_nhat`, `v_c12_tong_hop_moi_nhat`, `v_c13_tong_hop_moi_nhat`, `v_c14_tong_hop_moi_nhat`
  - dữ liệu thực tế đọc đúng cho các tổ của Sơn Tây và ngày `2026-04-20`
- Đã tạo admin utility đồng bộ DB nhiều instance:
  - [sqlite_history/sync_all_instance_dbs.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/sync_all_instance_dbs.py)
  - script này đứng ngoài full pipeline, dùng cho:
    - `status`
    - `apply-views`
    - `init-if-missing`
    - `reset-and-init`
  - đã verify thực tế:
    - `--mode status --unit son_tay`
    - `--mode apply-views --unit son_tay`

### Chưa triển khai

- Chưa thay thế các root mặc định đang hard-code ở luồng hiện tại
- Chưa làm hardening/report đặc biệt của Phase 5

## 2. Kiến trúc tổng thể

### 2.1. Luồng runtime mới

Luồng chạy sau refactor:

1. Người dùng chọn `--config configs/units/<unit_code>.yaml`
2. Hệ thống load file config
3. Hệ thống build `RuntimeContext`
4. `RuntimeContext` tạo và chuẩn hóa toàn bộ runtime roots
5. Downloader ghi raw vào `runtime/<unit_code>/downloads/...`
6. Processor đọc raw từ `downloads` của instance và ghi vào `Processed` của instance
7. Pipeline archive sang `ProcessedDaily` của instance
8. Import SQLite vào `runtime/<unit_code>/sqlite_history/report_history.db`

### 2.2. Cấu trúc thư mục runtime

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

Nguyên tắc:

- Chỉ có một codebase `api_transition` dùng chung cho mọi instance.
- Không copy source code vào `runtime/<unit_code>/`.
- `runtime/<unit_code>/` là root runtime duy nhất của một đơn vị.
- Không ghi dữ liệu runtime vào thư mục source code mặc định.
- Source code `api_transition/sqlite_history/` chỉ là code và SQL template, không phải thư mục data runtime.

### 2.3. Mô hình triển khai production

Production dùng mô hình:

- một codebase chung
- nhiều file config đơn vị
- nhiều Linux service theo từng instance

Mỗi service:

- load đúng một file config
- build `RuntimeContext`
- chạy full pipeline cho đúng một đơn vị
- ghi dữ liệu vào đúng `instance_root` của đơn vị đó

Contract bắt buộc:

- khác service nhưng cùng source code
- không deploy nhiều bản copy code chỉ để đổi đơn vị
- mọi thay đổi code ở một nơi phải có hiệu lực cho tất cả instance sau khi restart hoặc redeploy service

## 3. Mô hình dữ liệu cấu hình

### 3.1. File config cấp đơn vị

Mỗi đơn vị có một file YAML riêng.

Ví dụ:

```yaml
version: 1

unit:
  code: son_tay
  name: TTVT Son Tay

runtime:
  instance_root: runtime/son_tay
  create_dirs: true

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

download:
  headed: false
  retry_timeouts: [180, 300, 500]
  retry_delay_seconds: 3
  max_retries: 3

reports:
  defaults:
    enabled: true

  ghtt_hni:
    enabled: false

  xac_minh_tam_dung:
    service_ids: "8,9"

  phieu_hoan_cong_dich_vu_chi_tiet:
    service_ids: "1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16"
    customer_type: "0"
    contract_type: "0"
    ticket_type: "0"

  tam_dung_khoi_phuc_dich_vu_chi_tiet:
    service_ids: "8,9"
    date_type: "1"
    report_type: "0"

  tam_dung_khoi_phuc_dich_vu_tong_hop:
    service_ids: "1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16"
    report_type: "0"

  onebss_bc_chi_tiet_ket_qua_cskh_uc3:
    enabled: false
    customer_batch_code: "UC3_CSKH_042026"
    employee_id: "0"
    employee_text: "Tất cả"
```

### 3.2. Ý nghĩa các block

`version`

- Version của schema config.
- Dùng cho migration về sau.

`unit`

- `code`: định danh ngắn, dùng để tạo thư mục runtime.
- `name`: tên hiển thị.

`runtime`

- `instance_root`: root runtime của đơn vị.
- `create_dirs`: nếu `true` thì loader tự tạo thư mục còn thiếu.

`period`

- Kỳ báo cáo mặc định cho instance.
- Có thể bị ghi đè tạm thời qua CLI nếu cần.

`ids`

- Toàn bộ ID ngữ nghĩa của đơn vị.
- Không dùng key payload API trực tiếp ở đây.

`download`

- Tùy chọn chung cho downloader.

`reports`

- Cấu hình riêng cho từng report.
- Dùng để bật/tắt report hoặc override các tham số đặc thù.

## 4. `RuntimeContext`

### 4.1. Mục tiêu của `RuntimeContext`

`RuntimeContext` là object trung tâm đại diện cho một lần chạy của một đơn vị.

Nó phải giải quyết 4 việc:

- giữ config đã parse
- suy ra các đường dẫn runtime tuyệt đối
- cung cấp helper resolve output path
- làm đầu vào thống nhất cho downloader, processor, pipeline, importer

### 4.2. Định nghĩa dataclass đề xuất

```python
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Mapping, Optional


@dataclass(frozen=True)
class RuntimePaths:
    instance_root: Path
    downloads_root: Path
    processed_root: Path
    archive_root: Path
    sqlite_root: Path
    sqlite_db_path: Path
    sqlite_import_logs_root: Path
    sqlite_exports_root: Path


@dataclass(frozen=True)
class UnitProfile:
    code: str
    name: str
    ids: Mapping[str, str]


@dataclass(frozen=True)
class PeriodConfig:
    report_month: int
    report_year: int
    month_id: str = ""
    month_label: str = ""
    vattu_start_date: str = ""


@dataclass(frozen=True)
class DownloadConfig:
    headed: bool = False
    max_retries: int = 3
    retry_timeouts: tuple[int, ...] = (180, 300, 500)
    retry_delay_seconds: int = 3


@dataclass(frozen=True)
class RuntimeContext:
    config_path: Path
    paths: RuntimePaths
    unit: UnitProfile
    period: PeriodConfig
    download: DownloadConfig
    report_configs: Mapping[str, Mapping[str, Any]] = field(default_factory=dict)
    raw_config: Mapping[str, Any] = field(default_factory=dict)
```

### 4.3. Hành vi bắt buộc của `RuntimeContext`

`RuntimeContext` phải có các method hoặc helper tương đương:

```python
def get_report_config(self, report_key: str) -> Mapping[str, Any]: ...
def is_report_enabled(self, report_key: str) -> bool: ...
def download_group_dir(self, group_name: str) -> Path: ...
def processed_group_dir(self, group_name: str) -> Path: ...
def archive_group_dir(self, group_name: str) -> Path: ...
def sqlite_log_path(self, filename: str) -> Path: ...
def lock_file_path(self) -> Path: ...
```

Trạng thái hiện tại:

- Đã triển khai đủ các helper trên trong `RuntimeContext`
- `instance_root` tương đối hiện được resolve theo thư mục [api_transition](/home/vtst/baocaohanoi/api_transition/)
- Khi `runtime.create_dirs: true`, loader sẽ tự tạo cây runtime tại thời điểm load config

### 4.4. Quy tắc build path

Từ `instance_root`, loader phải tự sinh:

- `downloads_root = instance_root / "downloads"`
- `processed_root = instance_root / "Processed"`
- `archive_root = instance_root / "ProcessedDaily"`
- `sqlite_root = instance_root / "sqlite_history"`
- `sqlite_db_path = sqlite_root / "report_history.db"`
- `sqlite_import_logs_root = sqlite_root / "import_logs"`
- `sqlite_exports_root = sqlite_root / "exports"`
- `lock_file_path = instance_root / ".pipeline.lock"`

Mọi path phải được resolve tuyệt đối tại thời điểm build context.

## 5. Loader và validation

### 5.1. Module đề xuất

Nên tạo module mới:

- `api_transition/runtime_config.py`

Module này phụ trách:

- load YAML
- validate schema
- build `RuntimeContext`
- tạo thư mục runtime nếu cần

### 5.2. API đề xuất

```python
def load_runtime_context(config_path: str | Path) -> RuntimeContext: ...
def validate_runtime_config(raw: Mapping[str, Any]) -> None: ...
def ensure_runtime_dirs(paths: RuntimePaths) -> None: ...
```

### 5.3. Validation bắt buộc

Phải validate ít nhất:

- `version` tồn tại và bằng `1`
- `unit.code` không rỗng
- `runtime.instance_root` không rỗng
- `period.report_month` trong `1..12`
- `period.report_year` hợp lệ
- `download.retry_timeouts` là list số nguyên dương
- `ids.center_id_14` là string nếu report C/KPI được enable
- `ids.unit_id_28` là string nếu report nhóm 28xxxx được enable

Validation theo quan hệ:

- nếu `reports.ghtt_hni.enabled = true` thì phải có ID tương ứng cho report đó
- nếu `reports.onebss_bc_chi_tiet_ket_qua_cskh_uc3.enabled = true` thì phải có:
  - `onebss_region_id`
  - `onebss_region_text`
  - `onebss_unit_text`

Validation production:

- `runtime.instance_root` không được trỏ vào thư mục source code của repo
- `runtime.instance_root` phải resolve được thành một path runtime độc lập

### 5.4. Nguyên tắc xử lý thiếu config

- Thiếu field bắt buộc: fail fast ngay khi load config.
- Thiếu field tùy chọn: dùng default hợp lý.
- Không fallback âm thầm sang hard-code trong code downloader.

## 6. Resolver tham số theo report family

### 6.1. Vì sao cần resolver

Hiện tại cùng tên Python `unit_id` nhưng thực tế mỗi report dùng một loại ID khác nhau. Nếu không có resolver thì cấu hình sẽ sai rất dễ.

### 6.2. `ReportTask` cần bổ sung metadata

Nên mở rộng `ReportTask` trong `batch_download.py` như sau:

```python
@dataclass
class ReportTask:
    name: str
    key: str
    func: Callable
    params_type: str
    group: str
    id_family: str = ""
    extra_kwargs: Dict[str, Any] = field(default_factory=dict)
    use_shared_session: bool = True
```

`key` là định danh ổn định để lookup cấu hình report.

### 6.3. Danh sách `id_family` đề xuất

- `center_id_14`
- `unit_id_28`
- `service_scope_ttvt_id`
- `service_scope_team_id`
- `onebss_tt_id`
- `none`

### 6.4. Mapping theo report

#### Nhóm `center_id_14`

Áp dụng cho:

- `c11`
- `c12`
- `c13`
- `kpi_nvkt_c11`
- `kpi_nvkt_c12`
- `kpi_nvkt_c13`

Resolver:

```python
kwargs["unit_id"] = context.unit.ids["center_id_14"]
```

#### Nhóm `unit_id_28`

Áp dụng cho:

- `c14`
- `c14_chi_tiet`
- `c11_chi_tiet`
- `c12_chi_tiet_sm1`
- `c12_chi_tiet_sm2`
- `i15`
- `i15_k2`
- `ghtt_sontay`
- `ghtt_nvktdb`
- `xac_minh_tam_dung`
- `phieu_hoan_cong_dich_vu_chi_tiet`
- `tam_dung_khoi_phuc_dich_vu_chi_tiet`
- `tam_dung_khoi_phuc_dich_vu_tong_hop`
- `ty_le_xac_minh_ttvtkv`
- `ty_le_xac_minh_chi_tiet`
- `kq_tiep_thi`
- `vattu_thu_hoi`

Resolver:

```python
kwargs["unit_id"] = context.unit.ids["unit_id_28"]
```

#### Nhóm `service_scope_ttvt_id`

Áp dụng cho:

- `ngung_psc_mytv_thang_t_1_cap_ttvt`
- `ngung_psc_fiber_thang_t_1_cap_ttvt`

#### Nhóm `service_scope_team_id`

Áp dụng cho:

- `ngung_psc_fiber_thang_t_1_cap_to`
- `ngung_psc_mytv_thang_t_1_cap_to`

#### `ghtt_hni`

Report này cần xem là ngoại lệ cấu hình.

Khuyến nghị:

- mặc định `enabled: false`
- nếu bật thì cho phép khai báo riêng:

```yaml
reports:
  ghtt_hni:
    enabled: true
    unit_id: "284412"
```

### 6.5. Per-report override

Resolver cuối cùng phải merge theo thứ tự ưu tiên:

1. kwargs hệ thống tính từ `params_type`
2. `unit_id` từ `id_family`
3. `extra_kwargs` của task
4. override từ `reports.<report_key>`
5. override tức thời từ CLI nếu có

## 7. Tích hợp vào `batch_download.py`

### 7.1. Thay đổi signature

Đề xuất thêm:

```python
def run_batch_download(
    *,
    runtime_context: RuntimeContext,
    skip_reports: Optional[List[str]] = None,
    only_reports: Optional[List[str]] = None,
    session=None,
) -> Dict[str, Any]:
    ...
```

Giữ backward compatibility tạm thời bằng wrapper:

```python
def run_batch_download_legacy(...):
    ...
```

### 7.2. `_build_kwargs()` mới

`_build_kwargs()` phải nhận:

- `task`
- `context`
- `computed_dates`

Nó phải:

- tính ngày theo `params_type`
- resolve `unit_id`
- resolve `output_dir` theo group và instance
- merge report overrides

Ví dụ:

```python
def _build_kwargs(task: ReportTask, context: RuntimeContext, params: dict) -> dict:
    kwargs = {}
    kwargs["headed"] = context.download.headed
    kwargs["output_dir"] = str(context.paths.downloads_root / task.group)
    ...
```

### 7.3. Session dùng chung

Session login cho `baocao.hanoi` vẫn nên dùng chung trong một instance.

Không dùng chung session giữa các instance chạy song song.

## 8. Yêu cầu production với Linux service

### 8.1. Mô hình service

Khuyến nghị dùng:

- `systemd template service`
- hoặc `systemd service + timer`

Ví dụ:

- `api-transition@son_tay.service`
- `api-transition@ha_dong.service`

Mỗi service cùng gọi một codebase:

```ini
ExecStart=/usr/bin/python3 -m api_transition.full_pipeline --config /opt/api_transition/configs/units/%i.yaml
WorkingDirectory=/opt/api_transition
```

### 8.2. Hành vi khi service start

Khi service start:

1. parse `--config`
2. load config
3. build `RuntimeContext`
4. acquire instance lock
5. chạy full pipeline
6. release lock
7. thoát với exit code đúng

### 8.3. Locking theo instance

Phải có khóa theo `instance_root`.

Mục tiêu:

- tránh cùng một service bị trigger hai lần chồng nhau
- tránh chạy tay và chạy timer đè lên nhau

Thiết kế tối thiểu:

- file lock tại `runtime/<unit_code>/.pipeline.lock`
- lock scope theo instance, không lock toàn cục toàn hệ thống

### 8.4. Exit code

Quy ước nên có:

- `0`: pipeline thành công
- non-zero: có lỗi

Có thể chi tiết hơn:

- `1`: lỗi config hoặc validation
- `2`: không acquire được lock hoặc instance đang chạy
- `3`: lỗi download, process hoặc import

### 8.5. Logging

Toàn bộ log pipeline phải ra `stdout/stderr`.

Yêu cầu:

- log bắt đầu và kết thúc pipeline
- log config path, unit code, instance root
- log từng bước download, process, import
- log exception đủ rõ để tra qua `journalctl`

Nếu cần log file riêng, có thể đặt dưới:

- `runtime/<unit_code>/sqlite_history/import_logs/`

### 8.6. Headless

Service production phải chạy được không cần terminal.

Điều này có nghĩa:

- mặc định `headed=false`
- không phụ thuộc UI
- OTP/token phải tương thích môi trường service

## 9. Tích hợp vào downloader

### 8.1. `downloaders.py`

Hiện tại các downloader đã cho phép truyền `output_dir`. Đây là điểm tốt và nên giữ.

Cần làm:

- bỏ mọi assumption rằng `output_dir=""` sẽ là nơi dùng thật trong batch multi-instance
- toàn bộ batch phải luôn truyền `output_dir` rõ ràng

### 8.2. `onebss_downloaders.py`

Tương tự:

- luôn truyền `output_dir` từ `RuntimeContext`
- thêm adapter lấy tham số từ `context.report_configs`

## 10. Tích hợp vào processor

### 9.1. Mục tiêu

Processor phải đọc và ghi theo root runtime của instance, không theo root cố định của package.

### 9.2. `processors/common.py`

Hiện tại đang có:

- `DOWNLOADS_DIR = API_TRANSITION_DIR / "downloads"`
- `PROCESSED_DIR = API_TRANSITION_DIR / "Processed"`

Thiết kế mới:

```python
@dataclass(frozen=True)
class ProcessorPaths:
    downloads_root: Path
    processed_root: Path
```

Helper mới:

```python
def build_processed_path(input_path: Path, downloads_root: Path, processed_root: Path) -> Path: ...
def copy_raw_to_processed(input_path: Path, downloads_root: Path, processed_root: Path, ...) -> Path: ...
```

### 9.3. `processors/runner.py`

`run_all_processors()` nên nhận:

```python
def run_all_processors(
    *,
    runtime_context: RuntimeContext,
    overwrite_processed: bool = False,
    ...
) -> Dict[str, List[ProcessorRunResult]]:
    ...
```

### 9.4. Mức refactor tối thiểu

Vòng đầu chưa cần refactor hết mọi processor thành generic.

Chỉ cần:

- cho phép truyền input/output roots động
- giữ nguyên tên file nghiệp vụ
- giữ nguyên logic transform hiện có

## 11. Tích hợp vào `full_pipeline.py`

### 10.1. Signature mới

```python
def run_full_pipeline(
    *,
    runtime_context: RuntimeContext,
    download_only: Optional[Sequence[str]] = None,
    download_skip: Optional[Sequence[str]] = None,
    ...
) -> Dict[str, Any]:
    ...
```

### 10.2. Nguồn path

Tất cả path trong pipeline phải lấy từ `runtime_context.paths`:

- `processed_root`
- `archive_root`
- `sqlite_db_path`

Không được dùng default root của package khi đã có context.

### 10.3. Import SQLite

Importer vẫn dùng code hiện tại, nhưng nhận:

- `db_path = runtime_context.paths.sqlite_db_path`
- `processed_root = runtime_context.paths.processed_root`
- `archive_root = runtime_context.paths.archive_root`

### 11.4. Yêu cầu CLI production

`full_pipeline.py` phải hỗ trợ trực tiếp:

```bash
python3 -m api_transition.full_pipeline --config configs/units/son_tay.yaml
```

Khi chạy với `--config`, toàn bộ root runtime phải lấy từ `RuntimeContext`, không dùng default hard-code.

## 12. Cấu trúc file config chi tiết hơn

### 11.1. Schema đề xuất

```yaml
version: 1

unit:
  code: string
  name: string

runtime:
  instance_root: string
  create_dirs: bool

period:
  report_month: int
  report_year: int
  month_id: string
  month_label: string
  vattu_start_date: string

ids:
  center_id_14: string
  unit_id_28: string
  service_scope_ttvt_id: string
  service_scope_team_id: string
  onebss_tt_id: string
  onebss_default_team_id: string
  onebss_region_id: string
  onebss_region_text: string
  onebss_unit_text: string

download:
  headed: bool
  max_retries: int
  retry_timeouts: list[int]
  retry_delay_seconds: int

reports:
  defaults:
    enabled: bool
  <report_key>:
    enabled: bool
    ... per-report fields ...
```

### 11.2. Quy tắc `reports`

`reports.defaults`

- chứa default chung cho mọi report

`reports.<report_key>`

- chứa override riêng cho report đó

Ví dụ:

```yaml
reports:
  defaults:
    enabled: true

  ghtt_hni:
    enabled: false

  vattu_thu_hoi:
    vat_tu_ids: "1,2,3,4,8,6,5"
```

## 13. CLI đề xuất

### 12.1. Batch download

```bash
python3 -m api_transition.batch_download --config configs/units/son_tay.yaml
python3 -m api_transition.batch_download --config configs/units/ha_dong.yaml --only "C1.1"
```

### 12.2. Full pipeline

```bash
python3 -m api_transition.full_pipeline --config configs/units/son_tay.yaml
python3 -m api_transition.full_pipeline --config configs/units/ha_dong.yaml --strict
```

### 12.3. Validation config

Nên có thêm command:

```bash
python3 -m api_transition.runtime_config --check configs/units/son_tay.yaml
```

Đầu ra nên in:

- unit code
- instance root
- resolved runtime paths
- danh sách report enabled
- cảnh báo thiếu ID cho report enabled

## 14. Kế hoạch sửa file cụ thể

### Giai đoạn 1: Hạ tầng config và context

- tạo `api_transition/runtime_config.py`
- tạo thư mục `configs/units/`
- thêm 1 file mẫu `configs/units/example.yaml`

### Giai đoạn 2: Batch download

- sửa `batch_download.py`
- thêm `key` và `id_family` cho `ReportTask`
- thay `_build_kwargs()` bằng resolver theo `RuntimeContext`

### Giai đoạn 3: Pipeline và SQLite

- sửa `full_pipeline.py`
- chuyển toàn bộ default runtime roots sang lấy từ context

### Giai đoạn 4: Processor path injection

- sửa `processors/common.py`
- sửa `processors/runner.py`
- vá dần các processor còn bám path cứng

## 15. Rủi ro implementation

### Rủi ro 1

Một số processor đang hard-code output path riêng, đặc biệt nhóm dịch vụ/MyTV. Các chỗ này phải rà kỹ trước khi coi là multi-unit ready.

### Rủi ro 2

Một số report như `ghtt_hni` không hoàn toàn tuân theo logic ID giống các report còn lại. Phải để dạng ngoại lệ có cấu hình riêng.

### Rủi ro 3

Nếu vẫn cho phép fallback sang root mặc định trong repo, developer rất dễ vô tình chạy sai chỗ. Vì vậy nên fail fast khi có `RuntimeContext` nhưng không truyền path đúng.

### Rủi ro 4

Nếu deploy production bằng cách copy code thành nhiều thư mục khác nhau theo đơn vị, mục tiêu "sửa một nơi áp dụng cho tất cả instance" sẽ không còn đúng. Vì vậy deployment phải giữ mô hình một source code chung.

## 16. Khuyến nghị chốt

Thiết kế nên chốt theo 4 nguyên tắc:

1. Mọi lần chạy đều đi qua `RuntimeContext`.
2. Mọi dữ liệu runtime đều nằm trong `runtime/<unit_code>/`.
3. Mọi instance dùng chung một codebase, không nhân bản source code theo đơn vị.
4. File config dùng ID ngữ nghĩa, không dùng key payload API kỹ thuật làm public contract.

Với cách này, codebase sẽ có một public contract rõ ràng cho multi-instance, và vẫn giữ được phần lớn downloader hiện tại mà không phải viết lại từ đầu.
