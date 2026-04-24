# Quy trình sửa, thêm báo cáo và đồng bộ SQLite cho multi-instance

Tài liệu này chốt cách làm việc sau khi codebase `api_transition` đã chạy được theo mô hình:

- một codebase dùng chung
- nhiều file config đơn vị
- mỗi đơn vị có một `instance_root` riêng
- dữ liệu tách riêng theo:
  - `downloads`
  - `Processed`
  - `ProcessedDaily`
  - `sqlite_history/report_history.db`

Mục tiêu:

- khi sửa logic download/process/import trong codebase, thay đổi tự áp dụng cho tất cả instance
- không phải vá thủ công từng đơn vị
- vẫn giữ dữ liệu mỗi đơn vị nằm riêng trong DB riêng
- có một cách đồng bộ schema/view SQLite cho tất cả instance bằng một script nhỏ

## 0. Vị trí của `sync_all_instance_dbs.py` trong hệ thống

Script đồng bộ SQLite được đặt tại:

- [sqlite_history/sync_all_instance_dbs.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/sync_all_instance_dbs.py)

Đây là điểm rất quan trọng:

- `sync_all_instance_dbs.py` là **admin utility**
- nó **không nằm trong** `batch_download.py`
- nó **không nằm trong** `processors/runner.py`
- nó **không nằm trong** `full_pipeline.py`
- nó **không phải** một stage của pipeline hằng ngày

Nó chỉ được dùng cho vận hành và quản trị DB khi anh:

- muốn xem trạng thái DB của tất cả instance
- muốn apply schema + views mới cho các DB instance đang tồn tại mà không reset dữ liệu
- muốn apply lại views cho tất cả DB
- muốn tạo DB cho instance mới
- muốn reset và init lại toàn bộ DB do schema đổi

Luồng đúng là:

### Chạy nghiệp vụ hằng ngày

- `python3 -m api_transition.full_pipeline --config ...`

### Chạy quản trị DB toàn hệ thống

- `python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode status`
- `python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode apply-schema`
- `python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode apply-views`
- `python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode init-if-missing`
- `python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode reset-and-init`

Kết luận vận hành:

- pipeline dùng cho download/process/archive/import của **một instance**
- `sync_all_instance_dbs.py` dùng cho đồng bộ DB của **nhiều instance**
- không nhét script này vào pipeline tự động hằng ngày

## 1. Nguyên tắc cốt lõi

Nếu muốn một thay đổi áp dụng cho toàn bộ instance mà không làm hỏng mô hình multi-instance, phải giữ 4 nguyên tắc:

1. Sửa ở code dùng chung, không sửa trong thư mục `runtime/<unit>/...`
2. Mọi luồng chạy thật phải đi qua `--config` hoặc `RuntimeContext`
3. Không hard-code lại root cũ như:
   - `api_transition/downloads`
   - `api_transition/Processed`
   - `api_transition/ProcessedDaily`
   - `api_transition/report_history.db`
4. Mọi thay đổi DB phải tách thành:
   - thay đổi code importer/view/schema trong repo
   - bước đồng bộ DB cho tất cả instance

Nói ngắn gọn:

- sửa logic: sửa một nơi trong codebase
- sửa config: sửa YAML của các đơn vị liên quan
- sửa DB schema/view: sửa trong `api_transition/sqlite_history/` rồi chạy một lệnh đồng bộ tất cả DB instance
  - nếu chỉ thêm table/index/view mới mà không muốn reset dữ liệu: dùng `--mode apply-schema`
  - nếu chỉ đổi view: dùng `--mode apply-views`
  - nếu thay đổi schema phá vỡ tương thích: dùng `--mode reset-and-init`

## 2. Khi nào không cần sửa từng instance

Các thay đổi dưới đây, nếu làm đúng chỗ, sẽ tự áp dụng cho toàn bộ instance:

- sửa payload API trong downloader hiện có
- sửa rule xử lý workbook trong processor hiện có
- sửa rule import dữ liệu vào bảng SQLite hiện có
- sửa view SQL hiện có
- sửa log, retry, timeout, transform chung

Điều kiện:

- vẫn chạy qua:
  - `python3 -m api_transition.batch_download --config ...`
  - `python3 -m api_transition.processors.runner --config ...`
  - `python3 -m api_transition.full_pipeline --config ...`

## 3. Case 1: Sửa report đã có sẵn

Đây là case phổ biến nhất.

Ví dụ:

- report đổi tham số API
- report đổi tên cột
- processor phải sửa công thức
- importer phải map thêm cột hoặc bỏ cột

### 3.1. Nếu chỉ sửa downloader

Sửa ở:

- `downloaders.py`
- `onebss_downloaders.py`
- hoặc `batch_download.py` nếu cần đổi cách truyền tham số

Checklist:

- giữ nguyên `report_key`
- giữ nguyên `id_family` nếu semantics không đổi
- không hard-code `output_dir`
- nếu có output path thì phải lấy từ tham số hoặc từ batch runtime

Sau khi sửa:

1. chạy:
   - `python3 -m api_transition.batch_download --config api_transition/configs/units/son_tay.yaml --list`
2. chạy test thật 1 report hoặc full batch cho 1 đơn vị
3. nếu OK thì toàn bộ instance khác sẽ hưởng thay đổi đó khi chạy lại cùng codebase

Không cần làm:

- không sửa từng config đơn vị nếu `id_family` cũ vẫn đúng
- không sửa DB nếu chỉ thay download mà cấu trúc processed không đổi

### 3.2. Nếu chỉ sửa processor

Sửa ở:

- module tương ứng trong `processors/`
- có thể là `processors/common.py`
- có thể là `processors/runner.py` nếu đổi wiring input/output

Checklist:

- ưu tiên giữ signature nhận `input_path`, `output_path`, `overwrite_processed`
- không ghi ngược về raw file
- output phải tiếp tục đi vào `runtime/<unit>/Processed/...` qua runner

Sau khi sửa:

1. chạy:
   - `python3 -m api_transition.processors.runner --config api_transition/configs/units/son_tay.yaml`
2. kiểm tra file trong `runtime/son_tay/Processed`
3. nếu processor đổi cấu trúc workbook processed thì phải đánh giá luôn impact tới importer SQLite

### 3.3. Nếu chỉ sửa importer SQLite

Sửa ở:

- `sqlite_history/import_processed_to_sqlite.py`
- có thể thêm/sửa:
  - `SourceField`
  - hàm `import_*`
  - `populate_business_tables(...)`

Checklist:

- không phụ thuộc path mặc định nếu hàm đã có tham số `db_path`, `processed_root`, `archive_root`
- giữ nguyên contract:
  - workbook processed -> parse sheet -> ghi raw -> ghi business tables
- nếu đổi logic import mà không đổi schema:
  - chỉ cần chạy lại full pipeline hoặc import lại cho từng instance cần cập nhật dữ liệu

Lưu ý:

- sửa importer code không tự cập nhật dữ liệu cũ đã nằm trong DB
- muốn dữ liệu cũ phản ánh logic mới, phải import lại

## 4. Case 2: Thêm report mới

Đây là case cần chạm nhiều chỗ hơn.

## 4.1. Report mới chỉ cần download

Sửa ở:

- `downloaders.py` hoặc `onebss_downloaders.py`
- `batch_download.py`

Cần làm:

1. thêm hàm download
2. đăng ký `ReportTask` trong `batch_download.py`
3. khai báo:
   - `report_key`
   - `group`
   - `id_family` nếu cần
4. thêm config mặc định vào:
   - `configs/units/_template.yaml`
5. nếu report cần tham số đặc thù, thêm vào các YAML đơn vị liên quan

Không cần làm:

- không cần processor
- không cần SQLite nếu chưa import

## 4.2. Report mới cần download + process

Sửa ở:

- phần download như trên
- thêm processor mới trong `processors/...`
- đăng ký trong `processors/runner.py`

Cần làm:

1. tạo raw file name ổn định
2. tạo processor đọc đúng raw file
3. thêm `ProcessorTask`
4. nếu report dùng path đặc thù, thêm mapping runtime path ở `processors/runner.py`

Verify:

- `batch_download --config ...`
- `processors.runner --config ...`

Luu y van hanh:

- Neu downloader tai that thanh cong nhung workbook raw hien tai khong co du lieu, processor van nen sinh bo sheet canonical rong de pipeline khong vo contract.
- Truong hop do, co the verify logic processor/importer bang fixture hoac workbook lich su co du lieu that, nhung tai lieu phai ghi ro DB instance hien tai dang co `0` business rows cho snapshot do.

## 4.3. Report mới cần full flow download + process + SQLite

Sửa ở:

- downloader
- batch registry
- processor
- processor runner
- importer SQLite
- có thể cả schema/view

Checklist đầy đủ:

1. Thêm downloader
2. Thêm `ReportTask`
3. Thêm config vào `_template.yaml`
4. Cập nhật YAML đơn vị nếu report cần ID hoặc override riêng
5. Thêm processor
6. Thêm `ProcessorTask`
7. Nếu cần import business table:
   - thêm schema SQL
   - thêm hàm import trong `import_processed_to_sqlite.py`
   - nối vào `populate_business_tables(...)`
   - thêm view SQL nếu ứng dụng ngoài cần đọc
8. Chạy:
   - `full_pipeline --config ...`
9. Đối soát:
   - processed
   - DB
   - view đọc ra đúng

## 5. Case 3: Xóa bớt hoặc ngừng dùng một report

Có 3 mức.

### 5.1. Tạm thời tắt

Cách an toàn nhất:

- để nguyên code
- set trong YAML:
  - `reports.<report_key>.enabled: false`

Ưu điểm:

- không ảnh hưởng các đơn vị khác nếu chỉ muốn tắt một số đơn vị
- rollback rất dễ

### 5.2. Bỏ khỏi batch/process nhưng giữ dữ liệu cũ

Làm khi:

- report không muốn chạy nữa nhưng vẫn muốn giữ lịch sử DB

Cách làm:

- bỏ `ReportTask` khỏi `batch_download.py`
- bỏ `ProcessorTask` khỏi `processors/runner.py`
- không xóa bảng/view cũ trong SQLite

### 5.3. Xóa hoàn toàn khỏi hệ thống

Làm khi:

- chắc chắn không còn dùng report này

Cần làm:

- xóa downloader
- xóa processor
- xóa logic import
- xóa schema/view liên quan
- chạy đồng bộ SQLite cho toàn bộ instance

Đây là case rủi ro cao nhất, nên ưu tiên `disable` trước khi xóa thật.

## 6. Trọng tâm đặc biệt: cập nhật SQLite cho tất cả instance

Đây là phần quan trọng nhất của vận hành multi-instance.

## 6.1. Phân biệt 3 loại thay đổi SQLite

### A. Chỉ đổi view SQL

Ví dụ:

- thêm cột trong view dashboard
- sửa logic `JOIN`
- sửa `UNION`

Sửa ở:

- `sqlite_history/report_history_views.sql`

Tác động:

- không cần reset DB
- chỉ cần apply lại views cho tất cả DB instance

### B. Đổi importer nhưng không đổi schema

Ví dụ:

- đổi cách map dữ liệu từ processed vào bảng business
- thêm/bỏ dòng import nghiệp vụ nhưng bảng vẫn giữ nguyên

Sửa ở:

- `sqlite_history/import_processed_to_sqlite.py`

Tác động:

- code mới chỉ áp dụng cho lần import mới
- nếu muốn dữ liệu cũ phản ánh logic mới, phải import lại

### C. Đổi schema DB

Ví dụ:

- thêm bảng mới
- thêm cột mới
- thêm index
- đổi ràng buộc

Sửa ở:

- `sqlite_history/report_history_schema.sql`
- có thể cả `report_history_views.sql`

Tác động:

- phải chạy đồng bộ schema trên tất cả DB instance
- có thể cần:
  - `reset DB`
  - hoặc `migrate DB`

## 6.2. Phương án chuẩn đã triển khai

Đã tạo script riêng:

- [sqlite_history/sync_all_instance_dbs.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/sync_all_instance_dbs.py)

Mục tiêu:

- đọc toàn bộ config trong `configs/units/*.yaml`
- build `RuntimeContext` cho từng đơn vị
- lấy `context.paths.sqlite_db_path`
- chạy một trong các mode:
  - `apply-views`
  - `init-if-missing`
  - `reset-and-init`
  - `status`

### Giao diện đề xuất

```bash
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode status

python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode apply-views

python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode init-if-missing

python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode reset-and-init
```

Script hỗ trợ thêm:

- `--unit <unit_code>` để chỉ chạy cho một số instance
- `--configs-dir <path>` nếu cần đổi thư mục config
- `--schema-path <path>`
- `--views-path <path>`

### Ý nghĩa từng mode

`status`

- liệt kê:
  - đơn vị
  - path DB
  - DB có tồn tại không
  - số bảng
- số view
- thời gian sửa cuối

Ví dụ:

```bash
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode status --unit son_tay
```

`apply-views`

- dùng khi chỉ sửa `report_history_views.sql`
- apply lại views cho tất cả DB hiện có
- an toàn nhất

Ví dụ:

```bash
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode apply-views --unit son_tay
```

`init-if-missing`

- dùng khi thêm đơn vị mới
- nếu DB chưa có thì tạo mới
- nếu đã có thì bỏ qua

Ví dụ:

```bash
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode init-if-missing
```

`reset-and-init`

- dùng khi schema thay đổi mạnh hoặc muốn dựng lại sạch
- xóa DB cũ và tạo lại cho mọi instance
- chỉ dùng khi đã chấp nhận re-import lại dữ liệu

Ví dụ:

```bash
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode reset-and-init --unit son_tay
```

## 6.3. Quy trình vận hành SQLite theo từng loại thay đổi

### Trường hợp 1: chỉ sửa view

Quy trình:

1. sửa `report_history_views.sql`
2. commit code
3. chạy:
   - `sync_all_instance_dbs --mode apply-views`
4. test 1-2 view trên 1 instance

Không cần:

- reset DB
- re-import dữ liệu

### Trường hợp 2: sửa importer, không đổi schema

Quy trình:

1. sửa `import_processed_to_sqlite.py`
2. chạy test trên `son_tay`
3. nếu dữ liệu cũ cần đồng bộ lại logic mới:
   - chạy lại `full_pipeline --config ... --reset-db`
   - hoặc chạy lại import cho từng instance
4. nếu muốn đồng bộ tất cả instance:
   - viết một script chạy `full_pipeline` hoặc `import_processed_to_report_history(...)` cho từng config

Không cần:

- apply schema mới nếu bảng không đổi

### Trường hợp 3: thêm bảng mới hoặc đổi schema

Quy trình:

1. sửa `report_history_schema.sql`
2. sửa importer nếu cần
3. sửa views nếu cần
4. test trên 1 instance
5. chạy:
   - `sync_all_instance_dbs --mode reset-and-init`
6. chạy lại import/full pipeline cho các instance cần dựng lại dữ liệu

Khuyến nghị:

- với schema change mạnh, nên chấp nhận chiến lược:
  - recreate DB toàn bộ
  - re-import lại từ processed/archive

Điều này thực tế và an toàn hơn việc cố viết migration phức tạp ở giai đoạn hiện tại.

## 6.4. Định hướng nâng cấp về sau

Khi hệ thống ổn định hơn, có thể thêm:

- `migrate` mode có version schema
- bảng metadata lưu `schema_version`
- script migration theo version

Nhưng ở giai đoạn hiện tại, phương án thực dụng nhất là:

- view change -> `apply-views`
- schema change nhỏ nhưng chấp nhận rebuild -> `reset-and-init` + re-import

## 7. Quy trình chuẩn vận hành sau mỗi thay đổi

## 7.1. Nếu chỉ sửa code logic

1. sửa code trong repo
2. test trên `son_tay`
3. cập nhật tài liệu
4. chạy cho các đơn vị khác khi cần

## 7.2. Nếu thêm report mới

1. thêm downloader
2. thêm config/schema runtime nếu cần
3. thêm processor
4. thêm importer/schema/view nếu cần
5. test end-to-end trên 1 đơn vị
6. cập nhật `_template.yaml`
7. cập nhật YAML cho các đơn vị liên quan
8. cập nhật tài liệu

## 7.3. Nếu thay đổi SQLite

1. phân loại:
   - chỉ view
   - importer only
   - schema change
2. test trên `son_tay`
3. chạy script đồng bộ tất cả DB instance
4. nếu cần thì re-import lại dữ liệu
5. kiểm tra một vài view tiêu biểu

## 8. Checklist chống làm hỏng multi-instance

Trước khi commit, luôn kiểm tra:

- có chỗ nào mới hard-code `api_transition/downloads` không
- có chỗ nào mới hard-code `api_transition/Processed` không
- có chỗ nào mới hard-code `api_transition/report_history.db` không
- hàm mới có nhận path động hoặc được runner/batch inject không
- report mới có `report_key` rõ ràng không
- report mới có `id_family` đúng không
- config template có được cập nhật không
- nếu chạm SQLite:
  - có cần update schema không
  - có cần update views không
  - có cần re-import dữ liệu cũ không

## 9. Kết luận vận hành

Câu trả lời ngắn cho bài toán của anh là:

- nếu chỉ sửa logic download/process/import trong code dùng chung, thì không cần sửa từng instance
- nếu thêm report mới hoặc thêm họ ID mới, phải cập nhật registry/config ở codebase và YAML liên quan
- nếu sửa SQLite, nên có một script đồng bộ tất cả DB instance, thay vì chạm tay từng DB

Phương án chuẩn nên đi tiếp:

1. giữ toàn bộ runtime theo `RuntimeContext`
2. thêm script `sync_all_instance_dbs.py`
3. dùng:
   - `apply-views` cho thay đổi view
   - `reset-and-init` cho thay đổi schema mạnh
4. khi cần, re-import lại dữ liệu bằng full pipeline hoặc import loop cho toàn bộ instance
