# Kiến trúc đề xuất cho dashv4

## Mục tiêu kiến trúc

- độc lập với `dashv3`
- phụ thuộc tối thiểu vào repo ngoài
- dễ thay DB path theo đơn vị
- dễ thêm màn mới khi `api_transition` bổ sung dữ liệu

## Hướng công nghệ

Giữ Flask cho `dashv4` là lựa chọn thực dụng nhất:
- tận dụng hiểu biết hiện có từ `dashv3`
- dễ giữ route tương thích
- dễ tái sử dụng một phần template và JS theo từng màn
- không phát sinh chi phí chuyển framework trong lúc đang thay data layer

## Chiến lược codebase

`dashv4` không nên khởi tạo trắng.

Chiến lược đúng là:
1. copy nguyên codebase `dashv3` sang `dashv4`
2. giữ lại shell hiện có: blueprint, template, static, route
3. thay dần nguồn dữ liệu từng route sang SQLite mới
4. vô hiệu hóa rõ các route chưa có contract dữ liệu

Điểm quan trọng:
- copy codebase chỉ để giữ tốc độ và tương thích
- không có nghĩa chỉ đổi `FILE_PATH` sang một DB path là xong
- phần phải thay thật sự là data layer, service layer, và một phần adapter JSON

## Kiến trúc lớp

```text
Browser
  -> Flask route HTML
  -> page JS gọi API nội bộ
  -> API layer
  -> adapter/service layer
  -> repository/query layer
  -> SQLite views trong api_transition runtime DB
```

## Cấu trúc thư mục đích sau khi fork

```text
dashv4/
  dashboard.py
  config.py
  gunicorn_config.py
  app_helpers.py
  auth.py
  blueprints/
  services/
  templates/
  static/
  docs/
  repositories/
  serializers/
```

## Phân lớp chi tiết

### `repositories`

Chỉ chứa truy vấn SQL và mapping record:
- một repository cho mỗi miền dữ liệu
- không render JSON ở lớp này
- ưu tiên query `view`

Ví dụ:
- `quality_repository.py`
- `ghtt_repository.py`
- `auto_config_repository.py`
- `ttvt_summary_repository.py`

### `services`

Chứa logic ghép dữ liệu và chuẩn hóa:
- lọc theo đơn vị
- hợp nhất nhiều view
- chuẩn hóa tên chỉ tiêu
- tính lại một số metric phụ nếu cần
- bọc các phần route cũ của `dashv3` để giảm sửa trực tiếp trong blueprint

Trong giai đoạn hiện tại, một phần logic adapter vẫn còn nằm ngay trong blueprint. Hướng refactor tiếp theo là kéo các phần này về `services/` hoặc `serializers/` khi route đã ổn định.

### `serializers`

Chứa adapter để trả JSON giống `dashv3` khi cần:
- table payload
- summary cards
- grouped sections
- drilldown payload

Đây là lớp quan trọng nhất để giữ UI/route mà vẫn đổi nguồn dữ liệu.

Ngoài serializer cho route đã hỗ trợ, cần có một serializer chuẩn cho endpoint bị ngắt:
- `error = legacy_endpoint_disabled`
- `title`
- `reason`
- `required_display_contract`

Payload này là contract giao tiếp giữa `dashv4` và backlog dữ liệu của `api_transition`.

### `blueprints`

Nên tổ chức theo miền gần với `dashv3`:
- `operations`
- `quality`
- `growth`
- `retention`
- `inventory`

Trong pha đầu có thể giữ nguyên blueprint từ `dashv3`.

Tuy nhiên:
- route chưa có dữ liệu phải trả trạng thái rõ ràng
- route đã có dữ liệu mới nên chuyển dần sang repository/service riêng

## Nguyên tắc query

- Mỗi request mở connection read-only tới SQLite.
- Không viết trực tiếp vào DB nguồn.
- Mọi truy vấn có lọc đơn vị phải dùng tham số cấu hình, không hard-code rải rác.
- Không dùng `SELECT *` ở lớp API final.
- Chuẩn hóa timezone và ngày dữ liệu trả ra từ một chỗ.
- Ưu tiên `consumer view` như `v_dashboard_*_moi_nhat`; chỉ fallback sang view chi tiết khi DB chưa có summary view chính thức.
- Khi phải fallback aggregation, cần ghi rõ trong docs và trong backlog dữ liệu để thay thế sau này.

## Cấu hình tối thiểu cần có

- `DASHV4_DB_PATH`
- `DASHV4_UNIT_CODE`
- `DASHV4_UNIT_NAME`
- `DASHV4_HOST`
- `DASHV4_PORT`
- `DASHV4_SECRET_KEY`
- `DASHV4_TIMEZONE`

## Chính sách tái sử dụng frontend

Do mục tiêu hiện tại là đi nhanh, `dashv4` sẽ copy toàn bộ `static/js` và `templates` từ `dashv3`.

Sau khi copy, tách làm hai nhóm:
- nhóm có thể giữ gần nguyên trạng:
  - `cau_hinh_tu_dong`
  - `tong_hop_bsc_kpi`
- nhóm chỉ giữ giao diện và phải thay adapter:
  - `chatluong`
  - `giahan`
  - `tiepthi`
  - `ngungpsc`
  - `kpi`

Và thêm một nhóm thứ ba:
- nhóm phải ngắt khỏi nguồn cũ và trả `pending/501` cho đến khi DB có contract hiển thị:
  - `tam-dung-khoi-phuc`
  - `brcd`
  - `pttb`
  - `i15`
  - `i15k2`
  - `shc-processing`
  - `inventory` legacy

## Auth

Đợt đầu có thể giữ nguyên auth của `dashv3` để giảm số biến đổi.

Sau khi data layer ổn định mới quyết định có thay auth hay không.

## Chiến lược mở rộng

Khi `api_transition` có thêm `view` mới:
1. bổ sung repository tương ứng
2. bật route đang chờ
3. thêm test query tối thiểu
4. cập nhật tài liệu mapping

Khi `api_transition` chưa có đủ dữ liệu:
1. không giữ fallback Excel cũ trong `dashv4`
2. ghi rõ `required_display_contract` ở endpoint bị ngắt
3. coi contract này là backlog chính thức cho lần bổ sung dữ liệu tiếp theo
