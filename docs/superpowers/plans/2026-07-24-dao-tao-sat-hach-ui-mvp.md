# Kế hoạch triển khai UI MVP Đào tạo & Sát hạch

## 1. Mục tiêu

Xây dựng một lát cắt nghiệp vụ hoàn chỉnh trên Flask/Jinja hiện tại:

```text
Nhập JSON câu hỏi
-> duyệt và publish
-> tạo fixed template
-> tạo kỳ thi và giao người thi
-> người học làm bài/autosave/nộp bài
-> đóng và finalize
-> xem báo cáo và tải Excel
```

Sau khi lát cắt này ổn định, bổ sung giao diện kho tri thức và adapter OpenAI production. MVP không có lập lịch tự động, blueprint sinh đề theo ma trận, nhiều attempt, proctoring hoặc QTI.

Mỗi template/kỳ thi chỉ phục vụ một audience. Muốn tổ chức cho `nvkt`, `to_truong`, `b2a` thì tạo template và kỳ thi riêng.

## 2. Nguyên tắc triển khai

- Chỉ làm trong `/home/vtst/dashv4-training`, branch `feat/dao-tao-sat-hach-mvp`.
- Không checkout/merge `main`, không sửa `/home/vtst/dashv4`, không restart `dashv4@son_tay`.
- Dùng Flask/Jinja, CSS và JavaScript hiện có; tránh đưa thêm frontend framework hoặc build pipeline.
- Mọi thao tác ghi dùng API hiện có, CSRF và error envelope chuẩn.
- UI learner không được nhận đáp án đúng, explanation, evidence hoặc scoring metadata.
- AI chỉ tạo draft; không tự duyệt hoặc publish.
- Ưu tiên màn hình rõ ràng, vận hành được trên desktop và điện thoại; không mở rộng MVP bằng tính năng chưa cần thiết.

## 3. Pha UI-0 — Khóa guardrail backend còn thiếu

Trước khi UI phụ thuộc vào API kỳ thi:

1. Chỉ cho tạo assignment khi exam ở `draft` hoặc `ready`.
2. Từ chối username trùng trong cùng payload và assignment đã tồn tại bằng `ASSIGNMENT_ALREADY_EXISTS`, HTTP 409, error envelope ổn định.
3. Expose `POST /api/training/exams/<id>/cancel` với RBAC, CSRF và route tests.
4. Đồng bộ policy manual close trong spec: close trước deadline là `administratively_submitted`; close tại/sau deadline là `timed_out`.

Tiêu chí hoàn thành:

- Không thể thay đổi danh sách giao bài sau khi exam đã open/closed/cancelled/finalized.
- Không còn raw `sqlite3.IntegrityError` đi ra route.
- Test state, route và docs đều đồng bộ.

## 4. Pha UI-1 — Khung workspace theo vai trò

Nâng trang `/dao-tao-sat-hach` từ entry page thành workspace có điều hướng:

- Tổng quan.
- Kho tri thức.
- Ngân hàng câu hỏi.
- Mẫu đề.
- Kỳ thi.
- Bài thi của tôi.
- Báo cáo.

Phân quyền hiển thị:

- `learner`: Bài thi của tôi, kết quả được phép xem.
- `editor`: Kho tri thức, ngân hàng câu hỏi.
- `exam_manager`: Mẫu đề, kỳ thi, giao bài, báo cáo.
- `admin`: toàn bộ.

Tạo helper JavaScript dùng chung cho API/CSRF, loading, toast, modal xác nhận, error envelope, format thời gian Việt Nam và `Cache-Control: no-store` cho dữ liệu thi.

Tiêu chí hoàn thành:

- Không chỉ ẩn menu: backend vẫn kiểm tra role/scope.
- Người dùng không thấy thao tác ngoài quyền.
- Không phát sinh frontend framework hoặc bước build mới.

## 5. Pha UI-2 — Vertical slice quản trị nội dung

Triển khai trước luồng JSON để kiểm chứng toàn hệ thống mà chưa phụ thuộc OpenAI:

- Danh sách câu hỏi, lọc theo trạng thái, audience, domain/topic và tìm kiếm.
- Form nhập batch JSON theo schema hiện có.
- Hiển thị validation error theo câu/field.
- Xem chi tiết stem, options, đáp án, explanation, classification và evidence.
- Duyệt, từ chối và publish với modal xác nhận.
- Chỉ editor/manager phù hợp mới thấy đáp án và evidence.

Tiêu chí hoàn thành:

- Có thể nhập JSON mẫu, duyệt và publish hoàn toàn qua UI.
- Không serialize trực tiếp learner DTO từ entity quản trị.
- Validation lỗi không làm mất nội dung người dùng vừa nhập.

## 6. Pha UI-3 — Mẫu đề và tổ chức kỳ thi ✅ HOÀN THÀNH

> Commit range: `9643043..HEAD`. Chi tiết tại `docs/superpowers/plans/2026-07-24-training-ui3-templates-exams.md`.

### Mẫu đề cố định ✅

- Tạo template từ câu hỏi đã publish.
- Chọn một audience, thời lượng, điểm đạt, trộn câu và trộn đáp án.
- Lọc/chọn câu hỏi và xem trước tổng số câu/tổng điểm.
- Cảnh báo audience không phù hợp.
- Không sửa cấu trúc template đã được kỳ thi sử dụng.

### Kỳ thi ✅

- Tạo exam từ template.
- Nhập thời gian bắt đầu/kết thúc và cấu hình công bố kết quả.
- Chọn người thi, snapshot tên/tổ/đơn vị/audience.
- Hiển thị trạng thái `draft`, `ready`, `open`, `closed`, `cancelled`.
- Các nút `Ready`, `Open`, `Close`, `Cancel`, `Finalize` phải có modal mô tả tác động.
- Hiển thị số người chưa bắt đầu, đang làm, đã hoàn thành và recovery blocker.

### Hardening ✅

- `Cache-Control: no-store` trên management question detail.
- CAS cho review transitions (`BEGIN IMMEDIATE` + `WHERE ... AND review_status=?`).
- CAS cho publish (`BEGIN IMMEDIATE` + duplicate guard + `WHERE ... AND publication_status='unpublished'`).

### Read model APIs ✅

- `GET /api/training/templates` — list DTO.
- `GET /api/training/templates/<id>` — detail với items.
- `GET /api/training/exams` — list DTO.
- `GET /api/training/exams/<id>` — detail với template info và assignment_summary.
- `GET /api/training/exams/<id>/assignments` — assignment list.
- `GET /api/training/users` — assignable users (no password).

### Frontend ✅

- `static/js/training-templates.js` — template panel: list, create form, question picker.
- `static/js/training-exams.js` — exam panel: list, create form, assignment, lifecycle, status tracking.

### Tests ✅

- `tests/test_training_ui3_hardening.py` — Cache-Control + CAS concurrency tests.
- `tests/test_training_template_exam_routes.py` — route-level tests cho templates/exams/users.
- `tests/js/test_training_exams_lifecycle.mjs` — JS behavioral test cho exam lifecycle buttons.

Tiêu chí hoàn thành:

- Người vận hành thực hiện được toàn bộ vòng đời kỳ thi thủ công.
- Không tự open theo thời gian và không tự giao bài.
- Close/finalize hiển thị `recovery_summary` và không che giấu attempt lỗi.

## 7. Pha UI-4 — Trải nghiệm làm bài của learner ✅ HOÀN THÀNH

### Bài thi của tôi ✅

- Danh sách bài chưa làm, đang làm và đã hoàn thành.
- Hiển thị trạng thái, thời hạn, thời lượng và nút bắt đầu/tiếp tục.
- Backend: `GET /api/training/my-assignments` (`_learner_required` + CSRF).
- Frontend: `static/js/training-my-exams.js` (assignment catalog phân nhóm).

### Màn hình làm bài ✅

- Đồng hồ dùng deadline server trả về.
- Hiển thị danh sách câu và trạng thái đã trả lời.
- Chọn đáp án, autosave với `client_revision`, trạng thái `Đang lưu/Đã lưu/Lỗi kết nối`.
- Refresh trang phải khôi phục response đã lưu.
- Xác nhận trước khi submit.
- Khi exam đóng hoặc hết giờ, ngừng nhận input và đồng bộ trạng thái từ server.
- Không đưa correct options, explanation, evidence hoặc đáp án qua HTML/API learner.
- Frontend: `static/js/training-attempt.js`, CSS attempt workspace.
- HTML panels: `#training-my-exams`, `#training-attempt` trong `templates/pages/training/index.html`.
- JS behavioral test: `tests/js/test_training_attempt_workspace.mjs`.

Tiêu chí hoàn thành:

- Làm bài được trên desktop và màn hình điện thoại thông dụng ✅.
- Không mất câu trả lời sau refresh ✅.
- Revision cũ không ghi đè revision mới ✅.
- Submit retry không tạo result/audit trùng ✅.

## 8. Pha UI-5 — Báo cáo sau thi ✅ HOÀN THÀNH

Sau finalize, hiển thị:

- Tổng số được giao, đã thi, chưa thi, đạt, không đạt.
- Điểm trung bình và kết quả từng người.
- Phân tích theo topic và câu sai nhiều nếu payload hiện có hỗ trợ.
- Revision, checksum/thời điểm chốt phù hợp với report contract.
- Nút tải Excel từ snapshot đã chốt.

MVP chưa có điều chỉnh điểm và report revision mới qua UI.

## 9. Pha UI-6 — Kho tri thức và OpenAI production

### Kho tri thức

- Danh sách tài liệu/phiên bản.
- Form dán văn bản, chọn domain/topic/audience.
- Hiển thị block và issue cần xác nhận.
- Khởi tạo generation job từ document version đã chọn.

### OpenAI adapter

- Lazy import SDK; ứng dụng vẫn khởi động nếu chưa cấu hình OpenAI.
- Structured output theo JSON Schema hiện có.
- Web request chỉ tạo job; worker xử lý bên ngoài request.
- UI polling trạng thái `pending/running/completed/failed/cancelled`.
- Lưu provider/model/request metadata/raw response theo contract audit.
- Mock toàn bộ provider trong tests; không gọi mạng thật.
- Xem xét heartbeat/lease trong lúc provider call dài.

Tiêu chí hoàn thành:

- Từ tài liệu có thể tạo draft câu hỏi, xem evidence, duyệt và publish.
- AI không thể tự phát hành câu hỏi.

## 10. Pha UI-7 — Kiểm thử và nghiệm thu

Bổ sung tối thiểu:

- Route/UI authorization và CSRF tests.
- Test learner payload/HTML không chứa đáp án.
- Test 50 assignment và 50 attempt độc lập autosave đồng thời.
- Test 50 attempt độc lập submit trong burst ngắn: đúng một result mỗi attempt, không `SQLITE_BUSY` lọt ra API.
- Test close cạnh tranh với autosave/submit.
- Test refresh/resume và timeout UI ở mức JavaScript/helper phù hợp với stack hiện tại.
- Test finalize bị chặn khi snapshot lỗi và UI hiển thị blocker.
- Full suite `python3 -m pytest tests/`.
- Targeted `python3 -m pytest tests/test_training_*.py -q`.
- `python3 -m py_compile` cho file Python thay đổi và `git diff --check`.

## 11. Thứ tự commit đề xuất

1. `fix(training): close remaining exam API guardrails`
2. `feat(training-ui): add role-aware training workspace`
3. `feat(training-ui): add question review workflow`
4. `feat(training-ui): add template and exam management`
5. `feat(training-ui): add learner attempt experience`
6. `feat(training-ui): add finalized report dashboard`
7. `feat(training-ui): add knowledge generation workflow`
8. `feat(training-ai): add OpenAI structured generation adapter`
9. `test(training): add independent-attempt load coverage`
10. `docs(training): complete UI and operations runbook`

## 12. Điểm kiểm tra giữa các pha

- Sau UI-2: JSON question -> review -> publish chạy được qua UI.
- Sau UI-4: vertical slice đầy đủ tới learner submit chạy được.
- Sau UI-5: vertical slice đầy đủ tới finalize/report/Excel chạy được.
- Sau UI-6: thay JSON nhập tay bằng AI-generated draft nhưng giữ nguyên review/publish pipeline.

Không triển khai toàn bộ UI trong một commit lớn. Sau mỗi điểm kiểm tra phải chạy targeted tests, review learner data exposure và push commit lên đúng feature branch.
