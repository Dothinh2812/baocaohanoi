# Vận hành Đào tạo & sát hạch

## Phạm vi MVP

- Dữ liệu ghi nằm tại `DASHV4_TRAINING_DB_PATH`, mặc định `runtime_app/<unit>/training.db`.
- Không ghi vào `report_history.db` và không áp dụng lọc ngày.
- Nguồn tri thức MVP là paste text; AI chỉ tạo draft `single_choice` có evidence.
- Một assignment chỉ có một attempt; thi lại là exam/assignment mới.

## Migration và worker

```bash
venv/bin/python3 -m training.cli db-migrate
venv/bin/python3 -m training.cli worker --once --provider fake
```

Production AI chỉ được bật khi `DASHV4_TRAINING_AI_ENABLED=true`, provider/key hợp lệ được quản lý ngoài source và worker riêng chạy cùng environment file instance. Không chạy tác vụ AI dài trong Gunicorn.

## Backup và restore

- Dừng worker hoặc dùng SQLite backup API/`.backup`; không sao chép thô `training.db` khi WAL đang hoạt động.
- Backup kèm `DASHV4_TRAINING_FILES_DIR`.
- Trước restore, kiểm tra `training_instance_metadata.unit_code` khớp `DASHV4_UNIT_CODE`; chạy `db-migrate` sau restore.
- Thử restore định kỳ và đối chiếu attempt snapshot, result và report revision.

## Retention

Chưa có cleanup job MVP. Giữ document version, raw AI response, audit, attempt snapshot và report revision theo chính sách đơn vị; xác nhận thời hạn lưu và backup ngoài host trước nghiệm thu production.

## API vận hành

- Knowledge: `POST /api/training/knowledge`.
- Question draft/review/publish: `/api/training/questions/import`, `/<version>/approve`, `/<version>/publish`.
- Template/exam/assignment: `/api/training/templates`, `/api/training/exams`, `/api/training/exams/<id>/assignments`.
- Finalize/report/export: `/api/training/exams/<id>/finalize`, `/api/training/exams/<id>/report`, `/download/training/exams/<id>/report.xlsx`.

Mọi write API cần session, quyền module và CSRF `X-CSRF-Token`.
