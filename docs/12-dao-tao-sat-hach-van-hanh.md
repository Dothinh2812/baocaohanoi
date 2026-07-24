# Vận hành Đào tạo & sát hạch

## Phạm vi MVP

- Dữ liệu ghi nằm tại `DASHV4_TRAINING_DB_PATH`, mặc định `runtime_app/<unit>/training.db`.
- Không ghi vào `report_history.db` và không áp dụng lọc ngày.
- Nguồn tri thức MVP là paste text; AI chỉ tạo draft `single_choice` có evidence.
- Một assignment chỉ có một attempt; thi lại là exam/assignment mới.
- Route/API dùng RBAC module server-side theo `training_user_roles` và audience mapping; mọi write tiếp tục cần session, quyền và CSRF.
- Mốc integrity hiện tại: migration v11, fixed template bất biến, attempt/report snapshot checksum, close recovery có summary và report chỉ được tạo khi recovery không còn lỗi.

## Migration và worker

```bash
venv/bin/python3 -m training.cli db-migrate
venv/bin/python3 -m training.cli worker --once --provider fake
```

- Chạy `db-migrate` trước khi đưa instance vào vận hành; migration hiện có đến v11. v10 sửa lại `exam_template_items` và đồng bộ tổng số câu cache của template; v11 lưu `exam_events.closed_at_ms` để retry recovery giữ nguyên chính sách phân loại tại thời điểm đóng.
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

## Chính sách engine đã triển khai

- Trạng thái kỳ thi chỉ chuyển `draft -> ready -> open -> closed`; `draft` hoặc `ready` có thể `-> cancelled`. Không mở lại `closed`/`cancelled`, không close từ trạng thái khác `open`.
- `close` commit `open -> closed` trước để chặn start/autosave mới, rồi recovery từng attempt bằng transaction riêng. Retry close trên kỳ thi đã `closed` tiếp tục recovery attempt còn `active`.
- Close trước deadline chuyển attempt active thành `administratively_submitted` với `ended_reason=exam_closed`. Close đúng deadline hoặc sau deadline chuyển thành `timed_out` với `ended_reason=timeout`; mỗi attempt chỉ có một result gốc.
- Autosave kiểm tra theo thứ tự: attempt không còn `active` trả `ATTEMPT_ALREADY_COMPLETED`; kỳ thi không `open` hoặc đã qua `end_at` trả `EXAM_NOT_OPEN`; sau đó deadline `now >= deadline_at` trả `ATTEMPT_EXPIRED`. Vì vậy request đúng deadline không ghi response mới.
- Submit sau deadline chuyển attempt active thành `timed_out` và chỉ chấm response đã lưu. Khi close đã commit nhưng recovery chưa xong, autosave vẫn bị chặn bởi `EXAM_NOT_OPEN` và không làm thay đổi response cũ.
- Finalize chỉ nhận kỳ thi `closed`, hoặc kỳ thi `open` đã đến/hết `end_at`; finalize sớm giữ nguyên trạng thái `open` và không tạo report. Với kỳ thi open đã hết giờ, finalize close trước, nên attempt active được phân loại `timed_out` ngay cả khi deadline cá nhân còn muộn hơn `end_at`.
- Recovery trả `processed_attempt_ids`, `already_completed_ids` và `failed_attempts`. Lỗi snapshot của một attempt không rollback các attempt khác; nhưng finalize chặn report khi `failed_attempts` còn phần tử và trả `blocking_attempts` cùng `recovery_summary`.
- Chuỗi audience là question version -> template -> exam -> assignment. Question đã cấu hình audience phải khớp audience template; audience exam và assignment phải khớp audience trước đó. User đã có cấu hình audience phải thuộc audience được giao; user chưa cấu hình vẫn được giao và audience được snapshot vào assignment.

## Phạm vi smoke đồng thời

- Test hiện hành là smoke local in-process: service chạy từ thread với SQLite connection độc lập; không khẳng định scheduler đa process hay tải production.
- Có 50 thao tác autosave đồng thời cho một item (giữ revision cao nhất) và 50 submit đồng thời cho một attempt (một result/audit kết thúc); thêm race close/autosave/submit để kiểm tra một result terminal.

## Giới hạn nghiệm thu hiện tại

- Backend close/finalize và integrity snapshot đã có, nhưng chưa ghi nhận hoàn tất UI người vận hành/người học.
- Chưa xác nhận production OpenAI.
- Chưa xác nhận hoàn tất luồng thi lại; không coi các mục này là hoàn thành chỉ dựa trên API/backend.

Mọi write API cần session, quyền module và CSRF `X-CSRF-Token`.
