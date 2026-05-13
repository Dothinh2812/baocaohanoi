# Hướng dẫn triển khai và vận hành

## Phạm vi tài liệu này

Đây là playbook triển khai mục tiêu cho `dashv4`.

Tại thời điểm `2026-04-20`, `dashv4` được triển khai theo chiến lược fork từ `dashv3`, sau đó thay data layer dần sang SQLite mới.

Nếu triển khai nhiều instance song song từ cùng một codebase, đọc thêm:

- [10-trien-khai-multi-instance.md](/home/vtst/dashv4/docs/10-trien-khai-multi-instance.md)

## Môi trường mục tiêu

- Linux server hoặc VM nội bộ
- Python 3
- truy cập read-only tới DB:
  `/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db`
- reverse proxy nội bộ hoặc bind trực tiếp nội mạng

## Biến môi trường cần chuẩn hóa

```bash
export DASHV4_DB_PATH=/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db
export DASHV4_UNIT_CODE=son_tay
export DASHV4_UNIT_NAME="TTVT Sơn Tây"
export DASHV4_HOST=0.0.0.0
export DASHV4_PORT=5010
export DASHV4_TIMEZONE=Asia/Ho_Chi_Minh
export DASHV4_SECRET_KEY=change-me
```

## Quy ước vận hành

- App chỉ đọc DB, không ghi DB.
- Mọi connection tới SQLite phải ở chế độ read-only nếu thực hiện được.
- Log app tách riêng khỏi log pipeline.
- Không dùng chung thư mục `logs` với `dashv3`.
- Khi fork codebase, phải loại trừ `flask_session`, `logs`, `cache`, `__pycache__`, `.git` và artifact tạm.

## Cấu trúc chạy khuyến nghị

```text
systemd
  -> gunicorn
    -> dashv4 Flask app
      -> SQLite runtime DB
```

## Checklist trước khi go-live

### Dữ liệu

- xác nhận `DASHV4_DB_PATH` tồn tại
- kiểm tra `v_tien_do_nap_bao_cao` có dữ liệu ngày hiện tại
- kiểm tra các view pha 1 trả dữ liệu
- xác nhận route nào chưa có contract phải bị disable

### Ứng dụng

- app khởi động dưới tên `dashv4`, không đụng runtime của `dashv3`
- có healthcheck
- có trang lỗi rõ ràng khi DB không đọc được
- timeout hợp lý cho truy vấn lớn
- không có route nào còn đọc Excel trực tiếp

### Vận hành

- có file service `systemd`
- có log rotate
- có tài liệu restart/reload
- có kiểm tra quyền đọc DB

## Smoke test tối thiểu sau triển khai

1. mở trang chủ
2. mở `/tong-hop-bsc-kpi`
3. mở `/cau-hinh-tu-dong`
4. gọi API JSON của từng route pha 1
5. kiểm tra route disabled trả thông báo rõ ràng
6. kiểm tra ngày dữ liệu hiển thị đúng với DB

## Rủi ro vận hành cần ghi nhớ

### DB runtime không hoàn toàn Sơn Tây-only

Một số view trong runtime DB vẫn chứa nhiều `ttvt`. Do đó:
- app phải lọc theo đơn vị ở tầng query/service
- không được tin rằng “runtime/son_tay” là đủ

### Contract dữ liệu còn đang tiến hóa

Khi `api_transition` thêm hoặc sửa view:
- cần cập nhật repository và tài liệu mapping
- không nên khóa chặt frontend vào tên cột raw

### Route giữ nhưng dữ liệu chưa đủ

Với các route treo, nên trả trạng thái tường minh:
- `501 Not Implemented` cho route backend chưa hỗ trợ
- hoặc page thông báo “chưa có dữ liệu trên nguồn DB mới”

## Đề xuất hồ sơ triển khai trong repo khi bắt đầu code

Khi `dashv4` bắt đầu có code, nên có:
- `gunicorn_config.py`
- `start_dashv4.sh`
- `systemd/dashv4.service`
- `scripts/smoke_test.sh`
- `tests/`
- lớp `repositories/`
- lớp `serializers/`

## Kết luận vận hành

`dashv4` nên được triển khai như một app mới hoàn toàn, không phải bản chạy kèm trực tiếp bên trong `dashv3`.

Điểm khác biệt quan trọng nhất của triển khai mới:
- phụ thuộc chính là DB runtime của `api_transition`
- route được kế thừa có chọn lọc
- các màn chưa có contract dữ liệu phải được quản lý như backlog, không phải “sửa sau nhưng vẫn mở”
