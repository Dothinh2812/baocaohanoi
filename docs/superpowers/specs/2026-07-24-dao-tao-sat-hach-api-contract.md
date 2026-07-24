# API contract — Module Đào tạo & sát hạch (MVP)

- Ngày lập: `2026-07-24`
- Spec nguồn: `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-design.md` (1.2)
- Data dictionary: `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-data-dictionary.md`
- Tất cả endpoint dưới `blueprints/training_routes.py` (blueprint `training`).
- **Auth**: mọi route yêu cầu đăng nhập (auth chung dashboard). Không thêm vào `PUBLIC_ENDPOINTS`.
- **CSRF**: mọi `POST/PUT/PATCH/DELETE` qua `@csrf_protect` (header `X-CSRF-Token` hoặc form `csrf_token`).
- **Permission**: kiểm tra role module ở server (`training/permissions.py`).
- **Content-Type**: JSON API dùng `application/json`.
- **Timestamp**: UTC RFC 3339 (kết thúc `Z`) trong payload; DB lưu epoch ms.
- **Cache-Control**: response làm bài + response chứa đề = `no-store`.
- Endpoint names Flask dùng cho route policy: tiền tố `training.`.

## Error envelope ổn định

```json
{ "error": { "code": "ATTEMPT_ALREADY_ACTIVE", "message": "...", "details": {} } }
```

Mã lỗi tối thiểu (stable):

| code | HTTP | ngữ cảnh |
| --- | --- | --- |
| `ASSIGNMENT_NOT_FOUND` | 404 | assignment không tồn tại |
| `ATTEMPT_ALREADY_ACTIVE` | 409 | đã có attempt active |
| `ATTEMPT_ALREADY_COMPLETED` | 409 | attempt đã kết thúc |
| `ATTEMPT_EXPIRED` | 410 | quá deadline |
| `EXAM_NOT_OPEN` | 409 | kỳ thi chưa mở/đã đóng |
| `AUDIENCE_MISMATCH` | 400 | audience không khớp trong chuỗi question/template/exam/assignment hoặc user đã cấu hình không thuộc audience được giao |
| `ATTEMPT_SNAPSHOT_INVALID` | 409 | snapshot attempt bị sửa/không còn checksum hợp lệ; finalize bị chặn khi recovery còn lỗi này |
| `QUESTION_SUPPLY_INSUFFICIENT` | 409 | blueprint thiếu câu (sau MVP) |
| `VERSION_CONFLICT` | 409 | PATCH sai expected_version |
| `DOCUMENT_HAS_BLOCKING_ISSUES` | 409 | sinh câu khi có issue open mức cao |
| `PERMISSION_SCOPE_DENIED` | 403 | learner truy cập dữ liệu người khác |
| `CSRF_INVALID` | 400 | sai token (envelope chung dashboard) |

## Pagination contract chung

`?page=1&page_size=25&sort=-created_at&status=published&domain_code=quality&q=CCCO`
- `page_size` mặc định 25, tối đa 100.
- Sort chỉ allowlist field.
- Luôn trả `{ "items": [...], "page", "page_size", "total" }`.

---

## 1. Luồng làm bài (exam engine) — MVP trọng tâm

### 1.1 POST /api/training/assignments/<id>/attempts  — start attempt
- Permission: `learner` + ownership (assignment.username == session username).
- Idempotent: nếu đã có attempt `active`/`created`, trả attempt đó (không tạo mới).
- Transaction: `BEGIN IMMEDIATE` kiểm tra exam.status==open, start_at<=now<end_at, assignment.status, attempt hiện hữu → snapshot → commit.
- Response 200 (hoặc 409 nếu đã active):
```json
{
  "attempt_id": "att-...",
  "status": "active",
  "started_at": "2026-07-24T01:00:00Z",
  "deadline_at": "2026-07-24T01:25:00Z",
  "server_now": "2026-07-24T01:00:01Z",
  "attempt_version": 1
}
```
- Lỗi: `ASSIGNMENT_NOT_FOUND` (404), `EXAM_NOT_OPEN` (409), `PERMISSION_SCOPE_DENIED` (403), `ATTEMPT_ALREADY_ACTIVE` (409, kèm attempt hiện hành).

### 1.2 GET /api/training/attempts/<id> — load attempt (learner DTO)
- Permission: ownership. Header `Cache-Control: no-store`.
- **Learner DTO không có**: `correct_option_ids`, `explanation`, `distractor_rationales`, `evidence`, `scoring_policy`, `max_score` (chỉ có `points` hiển thị nếu cần). Serialize qua allowlist, không qua deny-list.
```json
{
  "attempt_id": "att-...",
  "status": "active",
  "deadline_at": "2026-07-24T01:25:00Z",
  "server_now": "2026-07-24T01:05:00Z",
  "items": [
    {
      "item_id": "atti-...",
      "sequence": 1,
      "type": "single_choice",
      "stem": "...",
      "stimulus": null,
      "options": [ {"id": "A", "text": "..."} ],
      "points": 1.0,
      "response": { "selected_option_ids": ["B"], "client_revision": 7 }
    }
  ]
}
```

### 1.3 PUT /api/training/attempts/<id>/responses/<item_id> — autosave
- Permission: ownership. Precedence: attempt không `active` trả `ATTEMPT_ALREADY_COMPLETED`; exam không `open` hoặc `server_now >= end_at` trả `EXAM_NOT_OPEN`; sau đó `server_now >= deadline` trả `ATTEMPT_EXPIRED`.
- Body: `{ "selected_option_ids": ["B"], "client_revision": 7 }`.
- Compare-and-set: chỉ ghi khi `client_revision > stored_revision`. Request cũ đến muộn → `accepted=false`, trả stored response, không lỗi.
- Response:
```json
{ "accepted": true, "stored_response": { "selected_option_ids": ["B"], "client_revision": 7 } }
```
- Lỗi: `ATTEMPT_EXPIRED` (410, không nhận đáp án mới, gồm đúng deadline), `EXAM_NOT_OPEN` (409), `ATTEMPT_ALREADY_COMPLETED` (409).

### 1.4 POST /api/training/attempts/<id>/submit — submit
- Permission: ownership. Idempotent: retry trả cùng result gốc.
- Nếu đã kết thúc có result thì retry trả result gốc, trừ `administratively_submitted` trả `ATTEMPT_ALREADY_COMPLETED`. Nếu quá deadline → chuyển `timed_out` + chấm response đã lưu (không nhận đáp án mới).
- Compare-and-set `status=active→submitted`. Chấm từ snapshot attempt_items/options.
- Response (single_choice):
```json
{
  "attempt_id": "att-...",
  "status": "submitted",
  "submitted_at": "2026-07-24T01:20:00Z",
  "result": { "score": 17, "maximum_score": 20, "percent": 85, "passed": true },
  "answers_released": false
}
```
- `answers_released` = `exam.reveal_answers_after_finalize AND exam.finalized`.

### 1.5 POST /api/training/exams/<id>/open — mở kỳ thi
- Permission: `exam_manager`/`admin`. Idempotent: đã open → 200.
- Compare-and-set `draft/ready→open`. Chỉ cho phép khi `start_at<=now<end_at`.

### 1.6 POST /api/training/exams/<id>/close — đóng kỳ thi
- Permission: `exam_manager`/`admin`. Idempotent.
- Compare-and-set `open→closed`, lưu `closed_at_ms`, rồi chặn attempt/response mới trước recovery. Active attempt được xử lý từng transaction: close trước deadline -> `administratively_submitted`/`exam_closed`; close tại hoặc sau deadline -> `timed_out`/`timeout`. Response có thêm `recovery_summary` gồm `processed_attempt_ids`, `already_completed_ids`, `failed_attempts`. Retry khi đã closed tiếp tục recovery. **Không** tạo report snapshot.

### 1.7 POST /api/training/exams/<id>/finalize — chốt kỳ thi
- Permission: `exam_manager`/`admin`. Idempotent: đã có report snapshot → trả revision hiện hành.
- Chỉ nhận exam `closed`, hoặc `open` khi `server_now >= end_at`; finalize sớm giữ exam open và trả `EXAM_NOT_OPEN`. Với open đã hết giờ, finalize close trước nên active attempt được `timed_out`. Recover attempt dở theo transaction riêng; assignment chưa bắt đầu -> `expired`; chỉ tạo report snapshot revision N khi `failed_attempts` rỗng. Nếu còn lỗi, trả lỗi attempt (ví dụ `ATTEMPT_SNAPSHOT_INVALID`) với `details.blocking_attempts` và `details.recovery_summary`.

### 1.8 GET /api/training/exams/<id>/report — xem báo cáo
- Permission: `exam_manager`/`admin`. Trả revision mới nhất (hoặc `?revision=N`).

### 1.9 GET /download/training/exams/<id>/report.xlsx — xuất Excel
- Permission: giống API report. File tạo tại yêu cầu, **không** nằm trong static/public, tên nội bộ không đoán được. Escape text bắt đầu `=`,`+`,`-`,`@`.

---

## 2. DTO theo đối tượng

### AttemptItemLearnerView (allowlist)
`item_id, sequence, type, stem, stimulus, options[{id,text}], points, response{selected_option_ids, client_revision}`.
**Không có**: correct_option_ids, explanation, distractor_rationales, evidence, scoring_policy, max_score.

### AttemptItemReviewView / ManagerView
Bổ sung: `correct_option_ids, explanation, distractor_rationales, evidence, max_score, scoring_policy, topic/competency codes`.

### ResultLearnerView
`score, maximum_score, percent, passed`. Không trả topic breakdown chi tiết nếu policy ẩn.

---

## 3. Kho tri thức (MVP subset)

- `GET /api/training/knowledge` — list (pagination, filter domain/category/status).
- `POST /api/training/knowledge` — paste text tạo draft. Body: `{document_code, title, content_text, classification{...}, audience_codes[]}`. Max `DASHV4_TRAINING_MAX_TEXT_BYTES`.
- `GET /api/training/knowledge/<id>` — chi tiết + versions.
- `POST /api/training/knowledge/<id>/versions` — tạo version mới.
- `GET /api/training/knowledge/<id>/issues` — danh sách issue.
- `POST /api/training/knowledge/issues/<id>/resolve` — resolve/exclude issue.

## 4. Câu hỏi & AI (MVP subset)

- `GET /api/training/questions` — list (pagination/filter/sort allowlist).
- `POST /api/training/questions` — tạo draft thủ công (qua cùng validator JSON).
- `PATCH /api/training/questions/<id>` — sửa draft; `expected_version` → `409 VERSION_CONFLICT`.
- `POST /api/training/questions/<id>/approve` — `needs_review→approved`.
- `POST /api/training/questions/<id>/publish` — `approved→published`; chặn duplicate normalized hash, chặn thiếu evidence span.
- `POST /api/training/generation-jobs` — enqueue AI job (idempotency_key).
- `GET /api/training/generation-jobs/<id>` — poll trạng thái.

## 5. Đề & kỳ thi (MVP subset)

- `POST /api/training/templates` — fixed template (published question_versions).
- `GET /api/training/templates` — list.
- `POST /api/training/exams` — tạo kỳ thi (template_id, audience, start/end/duration, pass_score).
- `POST /api/training/exams/<id>/assignments` — chọn users, snapshot assignment.

Audience bắt buộc nhất quán theo chuỗi question version -> template -> exam -> assignment. User đã có row `training_user_audiences` phải có audience được giao; user chưa có row được phép giao và audience được snapshot vào assignment.

## 6. Thi lại

- `POST /api/training/exams/<id>/retake-proposal` — chọn người chưa đạt → tạo draft exam/assignment mới có `retake_of_assignment_id`. Operator chọn template mới. Không tự open.
