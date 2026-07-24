# Module Đào tạo & sát hạch — MVP Implementation Plan

**Goal:** Triển khai một vertical slice hoàn chỉnh trong từng instance `dashv4`: dán văn bản → AI sinh câu hỏi `single_choice` có evidence → người vận hành duyệt → tạo template cố định → giao một lần thi → autosave/chấm điểm → finalize và xuất báo cáo Excel.

**Architecture:** Flask modular monolith, một SQLite ghi được `training.db` per-instance, UI/API/CLI dùng chung service và validator. Một assignment có tối đa một attempt record; thi lại tạo assignment/kỳ thi mới. AI chạy bằng worker SQLite đơn giản, không chạy lâu trong Gunicorn worker.

**Spec:** `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-design.md` phiên bản `1.1`.

## 1. Global constraints

- Chạy mọi lệnh từ `/home/vtst/dashv4`.
- Entry point mới phải `import runtime_limits` trước pandas/numpy hoặc thư viện kéo chúng vào.
- Test bằng `python3 -m pytest ...`, không dùng `venv/bin/pytest`.
- Không có lint/typecheck/formatter; dùng `python3 -m py_compile` và `git diff --check`.
- Không ghi vào `report_history.db`; dữ liệu module nằm trong `training.db` per-instance.
- Không sửa/xóa question version, attempt snapshot hoặc report revision đã được sử dụng.
- Không gọi API AI thật trong test.
- Tất cả route ghi dùng CSRF và permission check phía server.
- Không gọi AI, render Excel hoặc I/O file lớn trong transaction SQLite.
- Không tự nới điều kiện chọn câu và không tự publish nội dung AI.
- Không commit/push nếu người dùng chưa yêu cầu; giữ nguyên thay đổi không liên quan trong worktree.

## 2. MVP cut line

### Có trong MVP

- Paste text; danh mục; version tài liệu; extraction revision; blocks và issues.
- Worker AI sinh `single_choice` có evidence; review/publish question version.
- Import/export JSON câu hỏi nháp.
- Fixed template cho một audience mỗi kỳ thi.
- Kỳ thi do operator mở/đóng/finalize; một assignment–một attempt.
- Randomization có seed/version; autosave revision; deadline server; submit idempotent.
- Chấm `single_choice`, report revision, Excel và thi lại bằng kỳ thi mới.

### Để sau MVP

- PDF/DOCX/XLSX/OCR ingestion.
- Blueprint chọn ngẫu nhiên, supply validation và nhiều variant trong một kỳ thi.
- Multiple-choice, partial credit, pause và nhiều attempt.
- Adaptive retake, scope theo tổ, QTI, proctoring và PDF report chính thức.

## 3. Target file structure

### Tạo mới

```text
training/
  __init__.py
  cli.py
  constants.py
  db.py
  errors.py
  permissions.py
  time_policy.py
  migrations.py
  schemas/
    knowledge_manifest.schema.json
    question_batch.schema.json
    exam_template.schema.json
  providers/
    __init__.py
    base.py
    fake.py
    openai_provider.py

repositories/training_repository.py

services/
  training_catalog_service.py
  training_knowledge_service.py
  training_question_service.py
  training_generation_service.py
  training_exam_service.py
  training_attempt_service.py
  training_scoring_service.py
  training_report_service.py

blueprints/training_routes.py
templates/pages/training/*.html
static/js/pages/training/*.js
tests/test_training_*.py
deploy/dashv4-training-worker@.service
docs/12-dao-tao-sat-hach-van-hanh.md
```

### Sửa

```text
config.py
dashboard.py
blueprints/__init__.py
route_policy.py
templates/base.html
scripts/generate_instances.py
deploy/README-multi-instance.md
docs/00-doc-index.md
docs/04-mapping-route-va-du-lieu.md
docs/08-trang-thai-thuc-thi.md
```

## 4. Milestone 0 — Technical contract và dependency

### Task 0.1: Data dictionary và API contract

**Files:**

- Create: `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-data-dictionary.md`
- Create: `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-api-contract.md`

- [ ] Liệt kê từng bảng MVP với column/type/PK/FK/null/default/unique/index/immutable.
- [ ] Chốt timestamp là `INTEGER` UTC epoch milliseconds.
- [ ] Chốt enum state và delete policy; dữ liệu đã snapshot dùng `RESTRICT`/soft-retire.
- [ ] Ghi request/response/error cho start/load/autosave/submit/close/finalize/report.
- [ ] Chốt learner/reviewer/manager DTO; learner không có dữ liệu chấm.
- [ ] Ghi transaction boundary của start, autosave, submit, close và finalize.
- [ ] Map từng invariant trong spec sang constraint hoặc test.

**Gate:** Không tạo migration exam/attempt trước khi hai contract này được review.

### Task 0.2: Dependency strategy

**Files:**

- Create: `requirements-training.txt`

- [ ] Dùng tài liệu chính thức để xác minh phiên bản tương thích của `jsonschema` và OpenAI SDK tại thời điểm code.
- [ ] Pin dependency; không dựa vào Python hệ thống vì `venv/` hiện chưa có chúng.
- [ ] Provider AI dùng lazy import để app vẫn khởi động khi AI disabled.
- [ ] Cài vào `venv/` và ghi cách khôi phục dependency trong runbook.
- [ ] Test import bằng cả runtime venv và Python hệ thống.

**Verify:**

```bash
venv/bin/python3 -c "import jsonschema, openai; print('ok')"
python3 -c "import jsonschema, openai; print('ok')"
```

## 5. Milestone 1 — Nền dữ liệu, danh mục và RBAC

### Task 1.1: Config per-instance

**Files:** `config.py`, `scripts/generate_instances.py`, `tests/test_training_config.py`, `tests/test_generate_instances.py`.

- [ ] Test default path và env override cho DB/files/export.
- [ ] Thêm `TRAINING_DB_PATH`, `TRAINING_FILES_DIR`, `TRAINING_EXPORT_DIR`.
- [ ] Thêm AI enabled/provider/model, max text bytes, lease seconds và max retries.
- [ ] Expose trong `DashboardConfig`.
- [ ] Generator sinh path per-instance và AI-disabled mặc định; không ghi API key vào generated env.
- [ ] Chạy config/generator tests.

### Task 1.2: Domain constants, errors và time policy

**Files:** `training/constants.py`, `training/errors.py`, `training/time_policy.py`, `tests/test_training_domain.py`.

- [ ] Định nghĩa enum string cho review/publication/exam/assignment/attempt/result/job.
- [ ] Định nghĩa stable error codes từ spec.
- [ ] Implement `utc_now_ms()` và `compute_attempt_deadline_ms()`.
- [ ] Test deadline là min của exam end và start + duration, gồm boundary chính xác.
- [ ] Không hardcode timezone hiển thị; dùng `DASHV4_TIMEZONE`.

### Task 1.3: Migration runner và instance metadata

**Files:** `training/db.py`, `training/migrations.py`, `training/cli.py`, `tests/test_training_migrations.py`.

- [ ] Test DB mới, migrate lặp và unit-code mismatch.
- [ ] Connection riêng, WAL, busy timeout, foreign keys.
- [ ] Tạo `training_instance_metadata` và `training_schema_migrations`.
- [ ] Migration chạy trong transaction và file lock per-instance.
- [ ] Startup chỉ kiểm tra schema; CLI `db migrate` là đường vận hành chính.
- [ ] Test hai process/connection không làm hỏng schema.

### Task 1.4: Catalog và quyền module

**Files:** `training/migrations.py`, `repositories/training_repository.py`, `services/training_catalog_service.py`, `training/permissions.py`, `tests/test_training_permissions.py`.

- [ ] Tạo domain/category/indicator/topic/competency/audience/service/tag.
- [ ] Tạo `training_user_roles`, `training_user_audiences`.
- [ ] Admin dashboard có module admin theo policy, không đổi ý nghĩa role chung của `users.xlsx`.
- [ ] Implement permission service/decorator.
- [ ] MVP exam manager có scope toàn instance; learner chỉ xem dữ liệu cá nhân.
- [ ] Test nhiều role/audience và lỗi `PERMISSION_SCOPE_DENIED`.

**Milestone gate:** Migration/config/RBAC tests pass; app khởi động khi AI disabled.

## 6. Milestone 2 — Kho tri thức, câu hỏi và AI

### Task 2.1: Knowledge tables và service

**Files:** `training/migrations.py`, `repositories/training_repository.py`, `services/training_knowledge_service.py`, `tests/test_training_knowledge.py`.

- [ ] Tạo document/version/block/issue và classification/audience mappings.
- [ ] Paste-text create draft và SHA-256.
- [ ] Extraction revision; reprocess không ghi đè revision được tham chiếu.
- [ ] Deterministic block IDs trong version+revision.
- [ ] Review status tách khỏi effective dates.
- [ ] Chặn generation khi có blocking issue open.
- [ ] Test source version cũ vẫn đọc được sau version mới.

### Task 2.2: Knowledge UI/API

**Files:** `blueprints/training_routes.py`, `templates/pages/training/index.html`, `knowledge.html`, `static/js/pages/training/common.js`, `knowledge.js`, `tests/test_training_routes.py`.

- [ ] Trang tổng quan và kho tri thức.
- [ ] API CRUD draft, analyze request, issue list/resolve.
- [ ] Form paste text có max bytes và CSRF.
- [ ] Cây classification và filter cơ bản.
- [ ] Stable error envelope; sanitize nội dung AI trước render.

### Task 2.3: JSON Schema và semantic validator

**Files:** `training/schemas/*.schema.json`, `services/training_question_service.py`, `tests/test_training_question_schema.py`.

- [ ] Validate Draft 2020-12; object chính `additionalProperties: false`.
- [ ] MVP nhận `single_choice`; schema có version để mở rộng.
- [ ] Validate unique option IDs, đúng một correct ID và evidence block/revision tồn tại.
- [ ] Backend xác minh quote span, offset và content hash trước publish.
- [ ] Exact normalized hash chặn publish; Jaccard/token similarity chỉ cảnh báo.
- [ ] Calculation chỉ dùng allowlisted `rule_code`, không eval expression AI.
- [ ] Test payload đúng/sai, evidence sai, duplicate và calculation rule.

### Task 2.4: Question bank, review và versioning

**Files:** migration/repository/question service, `templates/pages/training/questions.html`, `question_review.html`, JS tương ứng, `tests/test_training_questions.py`.

- [ ] Tạo item/version/options/sources/reviews/classification mappings.
- [ ] Manual draft và JSON import cùng validator.
- [ ] Review UI hai cột câu hỏi/evidence.
- [ ] Publish version bất biến; sửa published tạo version mới.
- [ ] Source đổi gắn `source_changed`, không đổi history version cũ.
- [ ] List API có pagination/filter/sort allowlist.
- [ ] PATCH dùng expected version; conflict trả `409 VERSION_CONFLICT`.

### Task 2.5: AI job queue và fake provider

**Files:** migration, `services/training_generation_service.py`, `training/providers/base.py`, `fake.py`, `training/cli.py`, `tests/test_training_generation_jobs.py`.

- [ ] Tạo job/batch và state machine.
- [ ] Atomic claim bằng `BEGIN IMMEDIATE`, lease/heartbeat/retry/cancel.
- [ ] Worker command `python3 -m training.cli worker [--once]`.
- [ ] Fake provider trả fixture JSON, không mạng.
- [ ] Completed job chỉ tạo drafts, không publish.
- [ ] Test hai worker không claim cùng job và recover lease hết hạn.

### Task 2.6: Production AI provider adapter

**Files:** `training/providers/openai_provider.py`, `tests/test_training_ai_provider.py`.

- [ ] Trước code, dùng `openai-docs` kiểm tra API/model/Structured Outputs hiện hành.
- [ ] Provider nhận prompt/schema version, approved blocks và audience profile.
- [ ] Không gửi blocking issue hoặc tài liệu ngoài selection.
- [ ] Chuẩn hóa refusal/timeout/truncated/provider errors.
- [ ] Lưu provider/model/prompt/schema/usage metadata, không lưu API key.
- [ ] Unit test mock SDK/transport; smoke thật là thao tác có cấu hình riêng.
- [ ] AI disabled/thiếu key không làm app crash.

### Task 2.7: Generation UI và worker service

**Files:** knowledge/review templates+JS, `deploy/dashv4-training-worker@.service`, `deploy/README-multi-instance.md`, route tests.

- [ ] UI enqueue, poll, progress/error và link review.
- [ ] Refresh không tạo job trùng khi idempotency key giống nhau.
- [ ] Systemd worker dùng cùng env file instance, một process per-instance.
- [ ] Worker restart qua lease; dev hỗ trợ `worker --once`.

**Milestone gate:** Paste text → fake/real provider → draft có evidence → reviewer publish; test không cần Internet.

## 7. Milestone 3 — Fixed-template exam engine

### Task 3.1: Template, exam và assignment

**Files:** migration/repository, `services/training_exam_service.py`, `tests/test_training_templates.py`, `tests/test_training_exam_states.py`.

- [ ] Tạo fixed template/items, exam event, single variant và assignment.
- [ ] Template chỉ tham chiếu published question versions.
- [ ] Lock template/exam trước `ready`.
- [ ] Assignment unique theo exam+username+audience và snapshot tên/tổ/đơn vị.
- [ ] State draft→ready→open→closed/cancelled; không pause/open lại.
- [ ] Open/close/finalize dùng compare-and-set và audit.

### Task 3.2: Trang quản lý template/kỳ thi

**Files:** `templates/pages/training/templates.html`, `exams.html`, JS tương ứng, routes, route tests.

- [ ] Chọn question versions bằng filter và tạo fixed template.
- [ ] Tạo kỳ thi với start/end/duration/pass score/audience.
- [ ] Chọn users và snapshot assignment.
- [ ] Preview trước ready; open/close/finalize có confirmation/idempotent result.

### Task 3.3: Start attempt và immutable snapshot

**Files:** migration, `services/training_attempt_service.py`, `tests/test_training_attempt_start.py`, concurrency tests.

- [ ] Tạo attempt/items/options/result tables đúng data dictionary.
- [ ] Unique `exam_attempts(assignment_id)`.
- [ ] Start dùng `BEGIN IMMEDIATE`, kiểm tra exam/time/ownership cùng transaction.
- [ ] Snapshot scoring/classification/evidence/explanation; learner DTO không lộ chúng.
- [ ] Random seed server-side, algorithm version và checksum.
- [ ] Deadline UTC epoch ms theo time policy.
- [ ] Hai start đồng thời chỉ tạo một attempt; retry trả attempt hiện hành.

### Task 3.4: Learner page và autosave

**Files:** `templates/pages/training/take_exam.html`, JS, routes, `tests/test_training_autosave.py`, security tests.

- [ ] “Bài thi của tôi” và trang làm bài.
- [ ] Learner DTO allowlist; test không lộ correct IDs/explanation/evidence/scoring.
- [ ] Countdown từ deadline/server_now; không tin clock client để chấm.
- [ ] Autosave per-item với revision tăng đơn điệu; request cũ không ghi đè mới.
- [ ] `Cache-Control: no-store`, CSRF, ownership và phục hồi sau refresh.

### Task 3.5: Submit, timeout và scoring

**Files:** `services/training_scoring_service.py`, attempt service/routes, `tests/test_training_submit.py`, concurrency tests.

- [ ] Chấm chỉ từ snapshot.
- [ ] Compare-and-set active→submitted/timed_out/administratively_submitted.
- [ ] Tạo đúng một result trong transaction ngắn; retry trả cùng result.
- [ ] Sau deadline không nhận đáp án mới.
- [ ] Close đánh dấu exam trước, rồi kết thúc từng active attempt bằng transaction riêng.
- [ ] Không cập nhật counter tổng hợp chung khi submit.
- [ ] Test burst submit không tạo result trùng.

**Milestone gate:** Start, refresh, autosave, submit/timeout hoạt động; result độc lập question bank hiện hành.

## 8. Milestone 4 — Report, finalize và thi lại

### Task 4.1: Report aggregation/finalize revision

**Files:** migration, `services/training_report_service.py`, `tests/test_training_reports.py`.

- [ ] Aggregate từ immutable attempts/results, không dùng counter mutation.
- [ ] Report gồm assignment snapshot, result cá nhân, topic/question summary và chưa thi.
- [ ] Finalize recover attempt dở, expire assignment chưa bắt đầu, tạo revision 1.
- [ ] Snapshot có schema/algorithm version/checksum/payload.
- [ ] Finalize retry trả revision hiện hành.
- [ ] Score adjustment append-only tạo report revision mới.

### Task 4.2: Report UI và Excel

**Files:** report template/JS/routes, report/security tests.

- [ ] Cards, bảng cá nhân, nhóm/chủ đề và câu sai nhiều.
- [ ] Chọn revision để xem/xuất.
- [ ] Excel có summary/individual/topics/questions/retake.
- [ ] Escape text bắt đầu `=`, `+`, `-`, `@`.
- [ ] Download kiểm tra quyền, file không nằm static/public.
- [ ] Export từ snapshot, không join user hiện tại.

### Task 4.3: Thi lại thủ công

**Files:** exam/report services, report UI, report tests.

- [ ] Chọn người chưa đạt từ report revision.
- [ ] Tạo draft exam/assignment mới có `retake_of_assignment_id`.
- [ ] Operator chọn fixed template mới; chưa sinh adaptive blueprint.
- [ ] Không tái sử dụng assignment/attempt cũ và không tự open.
- [ ] Audit người tạo và revision nguồn.

**Milestone gate:** Report revision/Excel ổn định; thi lại không đổi lịch sử cũ.

## 9. Milestone 5 — Wiring, vận hành và nghiệm thu

### Task 5.1: Blueprint/sidebar/route policy

**Files:** `blueprints/__init__.py`, `dashboard.py`, `route_policy.py`, `templates/base.html`, route policy/tests.

- [ ] Register `training_bp`; thêm `training.page_index` vào active keys.
- [ ] Sidebar một mục “Đào tạo & sát hạch”.
- [ ] Disable per-instance qua policy; không thêm vào `PUBLIC_ENDPOINTS`.
- [ ] Sidebar ẩn khi endpoint bị disable.

### Task 5.2: Backup, restore và worker runbook

**Files:** `docs/12-dao-tao-sat-hach-van-hanh.md`, deploy README, `training/cli.py`, migration tests.

- [ ] CLI `db migrate/backup/verify`, `worker`, `questions validate/import`.
- [ ] Backup dùng SQLite backup API và kèm files dir.
- [ ] Restore kiểm tra unit code/schema/checksum.
- [ ] Hướng dẫn systemd worker, logs, lease recovery, AI-disabled mode và retention.
- [ ] Không copy DB đang ghi theo cách bỏ qua WAL.

### Task 5.3: Load/concurrency/recovery

**Files:** `tests/test_training_concurrency.py`, optional `scripts/training_load_smoke.py`.

- [ ] 50 learner autosave mỗi 10 giây, không mất/đảo revision.
- [ ] 50 submit trong 5 giây, đúng một result mỗi attempt.
- [ ] Hai worker claim/finalize không tạo duplicate.
- [ ] Restart app không mất response committed; worker chết được recover qua lease.
- [ ] Backup/restore tái hiện template/snapshot/result/report revision.
- [ ] Ghi host config và kết quả test vào runbook.

### Task 5.4: Doc sync và full verification

**Files:** docs 00/04/08/12.

- [ ] Mapping route ghi `training.db`, `supports_date = n/a`.
- [ ] Status doc phân biệt đã triển khai và sau MVP.
- [ ] Doc index liên kết spec/plan/runbook.
- [ ] Compile toàn bộ Python mới/sửa.
- [ ] Chạy targeted tests, full suite và `git diff --check`.

**Final verify:**

```bash
python3 -m py_compile training/*.py training/providers/*.py \
  repositories/training_repository.py services/training_*.py \
  blueprints/training_routes.py

python3 -m pytest tests/test_training_*.py -q
python3 -m pytest tests/
git diff --check
```

## 10. Dependency graph

```text
M0 contracts/dependencies
  -> M1 DB/config/catalog/RBAC
    -> M2 knowledge/question/AI
      -> M3 fixed-template exam engine
        -> M4 report/retake
          -> M5 wiring/ops/load/docs
```

Không chạy song song task cùng sửa migrations/state machine. UI chỉ bắt đầu sau khi API contract liên quan được khóa.

## 11. Definition of done

- Operator dán văn bản, enqueue AI, review và publish câu hỏi có evidence.
- Operator tạo fixed template/kỳ thi cho một audience, chọn người và mở thi.
- Learner chỉ thấy bài của mình, autosave có revision và không lộ đáp án.
- Deadline/close/submit nhất quán; retry không tạo duplicate.
- Chấm chỉ từ immutable attempt snapshot.
- Finalize tạo report revision và Excel an toàn.
- Thi lại tạo kỳ thi/assignment mới.
- Backup/restore và worker runbook đã được thử.
- Targeted/full tests pass; docs 00/04/08/12 đồng bộ.

## 12. User decisions theo gate

### Trước production AI

- Provider/model, chính sách gửi tài liệu nội bộ và nơi lưu API key/hạn mức chi phí.

### Trước giao thi thật

- Nguồn chính thức username–họ tên–tổ–audience.
- Mô tả trách nhiệm B2A/audience bổ sung.
- Chính sách công bố điểm/đáp án.

### Trước nghiệm thu báo cáo

- Mẫu Excel chính thức nếu có.
- Thời hạn lưu report/raw AI/audit/tài liệu nguồn và nơi backup ngoài host.

Các task nền, fake provider và engine fixture có thể làm trước; không bật AI/kỳ thi production khi gate chưa chốt.
