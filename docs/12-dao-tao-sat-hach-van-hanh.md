# Vận hành Đào tạo & sát hạch

## Phạm vi MVP

- Dữ liệu ghi nằm tại `DASHV4_TRAINING_DB_PATH`, mặc định `runtime_app/<unit>/training.db`.
- Không ghi vào `report_history.db` và không áp dụng lọc ngày.
- Nguồn tri thức MVP là paste text; AI chỉ tạo draft `single_choice` có evidence.
- Một assignment chỉ có một attempt; thi lại là exam/assignment mới.
- Route/API dùng RBAC module server-side theo `training_user_roles` và audience mapping; mọi write tiếp tục cần session, quyền và CSRF.
- Mốc integrity hiện tại: migration v11, fixed template bất biến, attempt/report snapshot checksum, close recovery có summary và report chỉ được tạo khi recovery không còn lỗi.

## Cài đặt dependencies

Module đào tạo dùng thêm package ngoài runtime chính. Cài vào đúng venv production/dev:

```bash
venv/bin/python3 -m pip install -r requirements-training.txt
```

Package trong `requirements-training.txt` (đã pin version):

| Package | Phiên bản | Dùng cho | Import |
|---------|-----------|----------|--------|
| `jsonschema` | 4.26.0 | Validate JSON Schema Draft 2020-12 | Luôn (schema validation) |
| `openai` | 2.29.0 | Provider AI Structured Outputs | Lazy (chỉ khi `--provider openai`) |
| `python-docx` | 1.2.0 | Đọc file DOCX | Lazy (chỉ khi import DOCX) |

- **`jsonschema`** cần thiết cho toàn bộ module; nếu thiếu sẽ lỗi khi validate question batch.
- **`openai`** lazy import: app/CLI không-AI vẫn chạy khi thiếu SDK. Lỗi rõ ràng khi thiếu: `openai SDK chưa cài đặt; chạy: pip install openai`.
- **`python-docx`** lazy import: CLI nhận `--file *.docx` sẽ lỗi rõ ràng khi thiếu. TXT và paste không cần python-docx.
- Không thay đổi package toàn hệ thống; chỉ cài vào venv đang dùng.
- Không đưa API key mẫu thật vào docs/log/test.

## Cấu hình AI environment

OpenAI provider chỉ hoạt động khi **cả hai** điều kiện thỏa:

```bash
export OPENAI_API_KEY='sk-...'          # từ secret store, KHÔNG ghi vào DB/log
export DASHV4_TRAINING_AI_ENABLED=1     # bật tính năng AI
```

Các env var cấu hình thêm (không bắt buộc):

| Env var | Mặc định | Mô tả |
|---------|----------|-------|
| `DASHV4_TRAINING_AI_MODEL` | `gpt-4o-mini` | Model OpenAI |
| `DASHV4_TRAINING_GENERATION_TIMEOUT_SECONDS` | `120` | Timeout mỗi API call |
| `DASHV4_TRAINING_LEASE_SECONDS` | `300` | Lease worker claim job |

Fake provider (`--provider fake`) không cần key hay env AI; dùng cho kiểm thử.

## Migration và schema

```bash
venv/bin/python3 -m training.cli db-migrate
```

- Chạy `db-migrate` trước khi đưa instance vào vận hành; migration hiện có đến v11.
- v10 sửa lại `exam_template_items` và đồng bộ tổng số câu cache của template.
- v11 lưu `exam_events.closed_at_ms` để retry recovery giữ nguyên chính sách phân loại tại thời điểm đóng.

## Luồng CLI Knowledge Base + AI sinh câu hỏi

Luồng đầy đủ từ nhập tài liệu đến xuất câu hỏi sẵn sàng thi:

### Bước 1 — Import tài liệu vào kho tri thức

```bash
# Import DOCX
python3 -m training.cli knowledge-import \
  --file c1_1.docx \
  --title "C1.1 Chất lượng sửa chữa thuê bao BRCĐ" \
  --domain quality --topics brcd_repair \
  --audiences nvkt,to_truong,b2a \
  --actor thinhdx.hni

# Import TXT
python3 -m training.cli knowledge-import \
  --file c1_1.txt \
  --title "C1.1 Chất lượng sửa chữa thuê bao BRCĐ" \
  --domain quality --topics brcd_repair \
  --audiences nvkt,to_truong,b2a \
  --actor thinhdx.hni

# Paste text trực tiếp
python3 -m training.cli knowledge-import \
  --paste-text "Nội dung tài liệu..." \
  --title "C1.1 ..." \
  --domain quality --topics brcd_repair \
  --audiences nvkt,to_truong,b2a \
  --actor thinhdx.hni
```

- Validate: extension, MIME signature, zip bomb, max bytes.
- Idempotent theo `content_sha256`: trùng nội dung trả conflict có hướng dẫn; dùng `--force` để tạo version mới.
- Output: `document_id`, `document_version_id`, extraction revision, block count.

### Bước 2 — Kiểm tra tài liệu đã import

```bash
# Liệt kê tài liệu
python3 -m training.cli knowledge-list --domain quality

# Xem chi tiết version (blocks, classification)
python3 -m training.cli knowledge-show --document-version-id <id>

# Kiểm tra issues (blocking issues phải resolve trước khi generate)
python3 -m training.cli knowledge-issues --document-version-id <id>

# Tạo extraction revision mới khi tài liệu dài/chunk chưa phù hợp.
# Revision cũ vẫn bất biến để evidence lịch sử không thay đổi.
python3 -m training.cli knowledge-reprocess \
  --document-version-id <id> --actor <editor_username>
```

### Bước 3 — Tạo generation job

```bash
python3 -m training.cli generate-create \
  --document-version-ids <ver_id_1>,<ver_id_2> \
  --audiences nvkt \
  --count 15 \
  --actor thinhdx.hni
```

- `--document-version-ids`: CSV document version IDs (bắt buộc).
- `--audiences`: CSV audience codes (bắt buộc).
- `--count`: số câu muốn sinh (bắt buộc).
- `--idempotency-key`: tuỳ chọn, tự tạo nếu bỏ trống.
- **Không có** `--selection-id`, `--provider`, `--model` ở lệnh này.
- Provider (fake/openai) chỉ chọn ở bước worker (bước 4).

### Bước 4 — Chạy worker sinh câu hỏi

```bash
# Fake provider (không cần API key, dùng cho kiểm thử)
python3 -m training.cli worker --provider fake --once

# OpenAI (cần OPENAI_API_KEY + DASHV4_TRAINING_AI_ENABLED=1)
python3 -m training.cli worker --provider openai --once

# Worker liên tục (không --once)
python3 -m training.cli worker --provider fake --poll-interval 5
```

- `--provider`: `fake` (mặc định) hoặc `openai`.
- `--once`: chạy 1 job rồi thoát; bỏ qua để worker chạy liên tục.
- `--worker-id`: tuỳ chọn, mặc định `cli-<pid>`.
- Worker claim job, resolve snapshot blocks, gọi provider, import draft câu hỏi.

### Bước 5 — Kiểm tra kết quả

```bash
# Liệt kê jobs
python3 -m training.cli generate-list

# Xem chi tiết job
python3 -m training.cli generate-show --job-id <id>

# Hủy job đang chờ
python3 -m training.cli generate-cancel --job-id <id> --actor thinhdx.hni
```

### Bước 6 — Review và publish câu hỏi

```bash
# Liệt kê câu hỏi draft
python3 -m training.cli questions-list --status draft --topic brcd_repair

# Xem chi tiết câu hỏi (stem, options, correct, evidence)
python3 -m training.cli questions-show --version-id <id>

# Approve (CAS, cần quyền editor)
python3 -m training.cli questions-approve --version-id <id> --actor thinhdx.hni

# Reject (cần comment)
python3 -m training.cli questions-reject --version-id <id> --actor thinhdx.hni \
  --comment "Evidence không chính xác"

# Publish (CAS, câu phải ở trạng thái approved)
python3 -m training.cli questions-publish --version-id <id> --actor thinhdx.hni
```

- State machine: `draft → approved → published`.
- Approve/publish dùng CAS (`BEGIN IMMEDIATE` + status check) chống concurrent transition.
- Sau publish, câu hỏi sẵn sàng dùng trong UI tạo mẫu đề, kỳ thi và giao bài.

### Bước 7 — Tạo kỳ thi và thi (UI hiện có)

Sau khi publish câu hỏi, dùng UI hiện có:
1. Tạo fixed template từ câu hỏi đã publish.
2. Tạo kỳ thi gán template.
3. Giao bài cho học viên.
4. Học viên làm bài, submit.
5. Close → finalize → xem report → export Excel.

## Worker AI

- Queue AI dùng lease/heartbeat/retry để worker chết có thể được claim lại sau khi lease hết hạn.
- Worker là process riêng, không chạy tác vụ AI dài trong Gunicorn.
- OpenAI provider: lazy import SDK; thiếu SDK trả lỗi rõ ràng, không crash app.

Production AI chỉ được bật khi `DASHV4_TRAINING_AI_ENABLED=1`, provider/key hợp lệ được quản lý ngoài source và worker riêng chạy cùng environment file instance. Fake provider là đường kiểm thử; chưa có xác nhận bật hoặc nghiệm thu production OpenAI.

## Backup và restore

- Dừng worker hoặc dùng SQLite backup API/`.backup`; không sao chép thô `training.db` khi WAL đang hoạt động.
- Backup kèm `DASHV4_TRAINING_FILES_DIR`.
- Trước restore, kiểm tra `training_instance_metadata.unit_code` khớp `DASHV4_UNIT_CODE`; chạy `db-migrate` sau restore.
- Thử restore định kỳ và đối chiếu attempt snapshot checksum, result và report revision checksum.

## Retention

Chưa có cleanup job MVP. Giữ document version, raw AI response, audit, attempt snapshot và report revision theo chính sách đơn vị; xác nhận thời hạn lưu và backup ngoài host trước nghiệm thu production.

## API vận hành

- Knowledge: `POST /api/training/knowledge`.
- Question bank UI-2 đã hoàn tất và được test (editor/exam-manager): `GET /api/training/questions` (page, page_size <= 100; filter status/audience/domain/topic/q), `GET /api/training/questions/<version>` (trả `Cache-Control: no-store`), `POST /api/training/questions/validate`, `/import`, `/<version>/approve`, `/<version>/reject`, `/<version>/publish`. List/detail are management DTOs, never raw SQLite rows; detail gồm scoring metadata và creation/approval/publication history khi có. Domain derives from evidence joined to `knowledge_blocks`; no schema migration was added. Import always creates draft; publish chỉ nhận unpublished approved; review và publish đã có CAS (`BEGIN IMMEDIATE` + `WHERE ... AND status=?`)防止 concurrent transition, review không được đổi published hoặc lặp state và publish lặp trả conflict không thêm audit. Learner không có question-bank API và attempt DTO không chứa đáp án, giải thích, evidence, distractor rationales hoặc metadata chấm.
- Template management UI-3 đã hoàn tất: `GET /api/training/templates` (list, pagination), `GET /api/training/templates/<id>` (detail với items), `POST /api/training/templates` (tạo từ câu hỏi đã publish). Frontend panel trong `static/js/training-templates.js` với danh sách, form tạo và question picker.
- Exam management UI-3 đã hoàn tất: `GET /api/training/exams` (list, filter status, pagination), `GET /api/training/exams/<id>` (detail với template info và assignment_summary), `GET /api/training/exams/<id>/assignments` (danh sách assignment), `GET /api/training/users` (danh sách user có thể giao bài, không trả password), `POST /api/training/exams` (tạo kỳ thi), `POST /api/training/exams/<id>/assignments` (giao bài), `POST /api/training/exams/<id>/cancel`, `POST /api/training/exams/<id>/ready`, `POST /api/training/exams/<id>/open`. Frontend panel trong `static/js/training-exams.js` với danh sách, form tạo, quản lý assignment, lifecycle buttons (Ready/Open/Close/Cancel/Finalize) và recovery summary display.
- Learner experience UI-4 đã hoàn tất: `GET /api/training/my-assignments` (danh sách assignment của learner, phân nhóm theo trạng thái, permission `_learner_required` + CSRF). Frontend `static/js/training-my-exams.js` (assignment catalog phân nhóm chưa bắt đầu/đang làm/đã hoàn thành/hết hạn, nút bắt đầu/tiếp tục). Frontend `static/js/training-attempt.js` (đồng hồ countdown từ `deadline_at_ms`, câu nav grid, autosave debounce 400ms `client_revision` CAS, submit confirm + unanswered count, lock UI khi terminal/expired). CSS attempt workspace styles. JS behavioral test `tests/js/test_training_attempt_workspace.mjs`.
- Report dashboard UI-5 đã hoàn tất: `GET /api/training/exams/<id>/report` (report snapshot, permission exam_manager/admin), `POST /api/training/exams/<id>/finalize` (chốt kỳ thi, idempotent), `GET /download/training/exams/<id>/report.xlsx` (export Excel). Frontend `static/js/training-reports.js` hiển thị danh sách finalized exams, chi tiết báo cáo (summary cards: được giao/đã làm/chưa làm/hết hạn/đạt/không đạt/điểm TB, individual results table), finalize action button với double-click protection và error handling blocking_attempts, export Excel button. Route RBAC: learner blocked từ report/finalize/export. Snapshot bất biến, report không lộ đáp án/explanation/evidence. JS behavioral test `tests/js/test_training_report_panel.mjs`.
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
- Assignment chỉ tạo khi exam chưa finalized và ở `draft` hoặc `ready`. Username trùng trong một request hoặc đã có assignment cùng `(exam_event_id, username, audience_code)` trả `ASSIGNMENT_ALREADY_EXISTS` (409); batch không ghi một phần.

## Phạm vi smoke đồng thời

- Test hiện hành là smoke local in-process: service chạy từ thread với SQLite connection độc lập; không khẳng định scheduler đa process hay tải production.
- Có 50 thao tác autosave đồng thời cho một item (giữ revision cao nhất) và 50 submit đồng thời cho một attempt (một result/audit kết thúc); thêm race close/autosave/submit để kiểm tra một result terminal.

## Giới hạn nghiệm thu hiện tại

- Backend close/finalize/report/export đã có. UI-3 (template/exam panels), UI-4 (learner experience) và UI-5 (report dashboard) đã hoàn tất.
- Chưa xác nhận production OpenAI.
- Chưa xác nhận hoàn tất luồng thi lại; không coi các mục này là hoàn thành chỉ dựa trên API/backend.

Mọi write API cần session, quyền module và CSRF `X-CSRF-Token`.
