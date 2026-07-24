# Vận hành Đào tạo & sát hạch

## Phạm vi MVP

- Dữ liệu ghi nằm tại `DASHV4_TRAINING_DB_PATH`, mặc định `runtime_app/<unit>/training.db`.
- Không ghi vào `report_history.db` và không áp dụng lọc ngày.
- Nguồn tri thức MVP là paste text; AI chỉ tạo draft `single_choice` có evidence.
- Một assignment chỉ có một attempt; thi lại là exam/assignment mới.
- Route/API dùng RBAC module server-side theo `training_user_roles` và audience mapping; mọi write tiếp tục cần session, quyền và CSRF.
- Mốc integrity hiện tại: migration v10, fixed template bất biến, attempt/report snapshot checksum, close recovery và finalize revision.

## Migration và worker

```bash
venv/bin/python3 -m training.cli db-migrate
venv/bin/python3 -m training.cli worker --once --provider fake
```

- Chạy `db-migrate` trước khi đưa instance vào vận hành; migration hiện có đến v10. v10 sửa lại `exam_template_items` và đồng bộ tổng số câu cache của template.
- Queue AI dùng lease/heartbeat/retry để worker chết có thể được claim lại sau khi lease hết hạn. Worker là process riêng, không chạy tác vụ AI dài trong Gunicorn.

Production AI chỉ được bật khi `DASHV4_TRAINING_AI_ENABLED=true`, provider/key hợp lệ được quản lý ngoài source và worker riêng chạy cùng environment file instance. Fake provider là đường kiểm thử; chưa có xác nhận bật hoặc nghiệm thu production OpenAI.

## Backup và restore

- Dừng worker hoặc dùng SQLite backup API/`.backup`; không sao chép thô `training.db` khi WAL đang hoạt động.
- Backup kèm `DASHV4_TRAINING_FILES_DIR`.
- Trước restore, kiểm tra `training_instance_metadata.unit_code` khớp `DASHV4_UNIT_CODE`; chạy `db-migrate` sau restore.
- Thử restore định kỳ và đối chiếu attempt snapshot checksum, result và report revision checksum.

## Retention

Chưa có cleanup job MVP. Giữ document version, raw AI response, audit, attempt snapshot và report revision theo chính sách đơn vị; xác nhận thời hạn lưu và backup ngoài host trước nghiệm thu production.

## API vận hành

- Knowledge: `POST /api/training/knowledge`.
- Question draft/review/publish: `/api/training/questions/import`, `/<version>/approve`, `/<version>/publish`.
- Template/exam/assignment: `/api/training/templates`, `/api/training/exams`, `/api/training/exams/<id>/assignments`.
- Close/finalize/report/export: `/api/training/exams/<id>/close`, `/api/training/exams/<id>/finalize`, `/api/training/exams/<id>/report`, `/download/training/exams/<id>/report.xlsx`.

## Giới hạn nghiệm thu hiện tại

- Backend close/finalize và integrity snapshot đã có, nhưng chưa ghi nhận hoàn tất UI người vận hành/người học.
- Chưa xác nhận production OpenAI.
- Chưa xác nhận hoàn tất luồng thi lại; không coi các mục này là hoàn thành chỉ dựa trên API/backend.

Mọi write API cần session, quyền module và CSRF `X-CSRF-Token`.
