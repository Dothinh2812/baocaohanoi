# Đào tạo & sát hạch Engine Integrity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Khóa tính toàn vẹn đề thi fixed-template, autosave, close/finalize và RBAC MVP trước khi xây UI/API mới.

**Architecture:** `training_exam_service` chịu trách nhiệm kiểm tra cấu hình template và transition kỳ thi; `training_attempt_service` sở hữu snapshot, randomization, autosave và kết thúc attempt. Constraints SQLite chặn dữ liệu template sai ở tầng storage, trong khi service/route kiểm tra quyền và contract API trước khi ghi.

**Tech Stack:** Python 3, Flask, SQLite WAL, pytest, existing training migration runner.

---

## File Map

- `training/errors.py`: thêm các mã lỗi autosave ổn định.
- `training/migrations.py`: migration 8 tái tạo `exam_template_items` với FK question version và unique template/question.
- `services/training_exam_service.py`: validate template và close active attempts theo transaction ngắn.
- `services/training_attempt_service.py`: deterministic shuffle, snapshot checksum đầy đủ, validation autosave, helper administrative submit.
- `services/training_report_service.py`: finalize không đổi snapshot report và dùng path administrative-submit đúng semantics.
- `blueprints/training_routes.py`: decorator role riêng cho editor/manager, CSRF thống nhất và ownership attempt.
- `tests/test_training_exam_states.py`: test integrity template/close.
- `tests/test_training_attempt_lifecycle.py`: test shuffle, autosave invalid và close/concurrency.
- `tests/test_training_reports.py`: test finalize không thay report đã chốt.
- `tests/test_training_routes.py`: test route RBAC và error envelope.

### Task 1: Lock Template References

**Files:**
- Modify: `training/errors.py`
- Modify: `training/migrations.py`
- Modify: `services/training_exam_service.py`
- Modify: `tests/test_training_exam_states.py`

- [ ] **Step 1: Write the failing service and migration tests**

```python
def test_create_template_rejects_missing_duplicate_and_wrong_audience(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_id = _publish_questions(db_path)[0]
    with pytest.raises(TrainingError) as missing:
        es.create_template(db_path, unit_code="son_tay", actor="mgr", code="missing", title="T",
            target_audience_code="nvkt", question_version_ids=["missing"], duration_seconds=60, pass_score_percent=80)
    assert missing.value.code == "TEMPLATE_QUESTION_INVALID"
    with pytest.raises(TrainingError) as duplicate:
        es.create_template(db_path, unit_code="son_tay", actor="mgr", code="duplicate", title="T",
            target_audience_code="nvkt", question_version_ids=[version_id, version_id], duration_seconds=60, pass_score_percent=80)
    assert duplicate.value.code == "TEMPLATE_QUESTION_DUPLICATE"
```

Add a migration test which runs migrations through version 8 and asserts `PRAGMA foreign_key_list(exam_template_items)` contains `question_versions` and duplicate `(template_id, question_version_id)` insertion raises `sqlite3.IntegrityError`.

- [ ] **Step 2: Run the focused tests to verify failure**

Run: `python3 -m pytest tests/test_training_exam_states.py -k "missing or duplicate" -q`

Expected: FAIL because template creation accepts nonexistent IDs and duplicates.

- [ ] **Step 3: Add migration 8 and service validation**

Add these error constants:

```python
TEMPLATE_QUESTION_INVALID = "TEMPLATE_QUESTION_INVALID"
TEMPLATE_QUESTION_DUPLICATE = "TEMPLATE_QUESTION_DUPLICATE"
```

Create `migration_008(conn)` that disables foreign keys, creates `exam_template_items_new` with both foreign keys and both unique constraints, copies valid existing rows, replaces the old table, recreates `idx_titems_template_qv`, then reenables foreign keys. Append `(8, migration_008)` to `_MIGRATIONS`.

At the start of `create_template`, reject duplicates before SQL, fetch every requested version with publication status and audience mappings, then reject a missing/unpublished version with `TEMPLATE_QUESTION_INVALID`. Reject a version whose nonempty `question_audiences` has no row matching `target_audience_code`. Keep `total_questions=len(question_version_ids)` server-derived and insert only after all validation passes.

- [ ] **Step 4: Run focused tests**

Run: `python3 -m pytest tests/test_training_exam_states.py tests/test_training_migrations.py -q`

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add training/errors.py training/migrations.py services/training_exam_service.py tests/test_training_exam_states.py tests/test_training_migrations.py
```

### Task 2: Snapshot Deterministic Shuffle

**Files:**
- Modify: `services/training_attempt_service.py`
- Modify: `tests/test_training_attempt_lifecycle.py`

- [ ] **Step 1: Write failing reproducibility tests**

```python
def test_snapshot_shuffle_reproduces_seed(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    first = att.snapshot_order(db_path, started["attempt_id"])
    assert first == att.reproduce_snapshot_order(db_path, started["attempt_id"])

def test_snapshot_respects_question_and_option_shuffle_flags(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    # Set both flags false before start, then assert template sequence and display_order are unchanged.
    # Repeat with both flags true and assert at least one seeded ordering differs from source ordering.
```

- [ ] **Step 2: Run tests to verify failure**

Run: `python3 -m pytest tests/test_training_attempt_lifecycle.py -k "shuffle" -q`

Expected: FAIL because `exam_row_has_shuffle()` always returns `False` and question order is never shuffled.

- [ ] **Step 3: Implement a versioned deterministic ordering helper**

```python
def _seeded_rng(seed, label):
    digest = hashlib.sha256(f"{SHUFFLE_ALGORITHM_VERSION}:{seed}:{label}".encode()).digest()
    return random.Random(int.from_bytes(digest[:16], "big"))

def _ordered(values, *, seed, label, enabled):
    ordered = list(values)
    if enabled:
        _seeded_rng(seed, label).shuffle(ordered)
    return ordered
```

Load `shuffle_questions` and `shuffle_options` from `exam_templates` in `_build_snapshot`. Apply `_ordered` to template item rows before assigning snapshot `sequence_number`; apply it to each item option list using a label containing its immutable question-version ID. Expand `_snapshot_checksum` input to include assigned item sequence and option code/display order. Remove `exam_row_has_shuffle`.

Expose small read-only helpers used only by tests: `snapshot_order(db_path, attempt_id)` reads attempt item question IDs and option codes by snapshot display order; `reproduce_snapshot_order(db_path, attempt_id)` re-reads template input plus stored seed/version and uses the same ordering helper.

- [ ] **Step 4: Run focused tests**

Run: `python3 -m pytest tests/test_training_attempt_lifecycle.py -k "shuffle or snapshot" -q`

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add services/training_attempt_service.py tests/test_training_attempt_lifecycle.py
```

### Task 3: Reject Invalid Autosave Without Mutation

**Files:**
- Modify: `training/errors.py`
- Modify: `services/training_attempt_service.py`
- Modify: `tests/test_training_attempt_lifecycle.py`

- [ ] **Step 1: Write failing validation tests**

```python
def test_autosave_rejects_item_from_other_attempt_without_response(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, first_assignment = _make_open_exam_with_assignment(db_path, "u1")
    _, second_assignment = _make_open_exam_with_assignment(db_path, "u2")
    first = att.start_attempt(db_path, unit_code="son_tay", actor="u1", assignment_id=first_assignment)
    second = att.start_attempt(db_path, unit_code="son_tay", actor="u2", assignment_id=second_assignment)
    foreign_item = att.get_attempt_learner_view(db_path, second["attempt_id"])["items"][0]["item_id"]
    with pytest.raises(TrainingError) as exc:
        att.save_response(db_path, attempt_id=first["attempt_id"], attempt_item_id=foreign_item,
                          selected_option_ids=["A"], client_revision=1)
    assert exc.value.code == "ATTEMPT_ITEM_NOT_FOUND"
    assert att.response_count(db_path, first["attempt_id"]) == 0
```

Add equivalent tests for selected `"Z"` (`INVALID_OPTION_SELECTION`) and `selected_option_ids=["A", "B"]` on a single-choice item (`SINGLE_CHOICE_REQUIRES_ONE_OPTION`), each asserting no row was inserted.

- [ ] **Step 2: Run tests to verify failure**

Run: `python3 -m pytest tests/test_training_attempt_lifecycle.py -k "autosave_rejects" -q`

Expected: FAIL because the service currently inserts the foreign item and unvalidated option codes.

- [ ] **Step 3: Validate snapshot relation before revision comparison**

Add these stable error constants:

```python
ATTEMPT_ITEM_NOT_FOUND = "ATTEMPT_ITEM_NOT_FOUND"
INVALID_OPTION_SELECTION = "INVALID_OPTION_SELECTION"
SINGLE_CHOICE_REQUIRES_ONE_OPTION = "SINGLE_CHOICE_REQUIRES_ONE_OPTION"
```

Inside the existing `BEGIN IMMEDIATE`, query `exam_attempt_items` by both `id` and `attempt_id`. Query its allowed option codes from `exam_attempt_options`. Require `selected_option_ids` to be a JSON-list-like Python list of unique strings, require all values be in that set, and for `single_choice` require list length <= 1. Raise `TrainingError` before querying/upserting `exam_responses`. Add `response_count(db_path, attempt_id)` as a read-only test helper.

- [ ] **Step 4: Run focused tests**

Run: `python3 -m pytest tests/test_training_attempt_lifecycle.py -k "autosave" -q`

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add training/errors.py services/training_attempt_service.py tests/test_training_attempt_lifecycle.py
```

### Task 4: Administrative Close and Immutable Finalize

**Files:**
- Modify: `services/training_attempt_service.py`
- Modify: `services/training_exam_service.py`
- Modify: `services/training_report_service.py`
- Modify: `tests/test_training_attempt_lifecycle.py`
- Modify: `tests/test_training_reports.py`

- [ ] **Step 1: Write failing close/finalize tests**

```python
def test_close_scores_active_attempt_and_blocks_new_writes(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, started["attempt_id"])["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["B"], client_revision=1)
    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    assert att.get_attempt(db_path, started["attempt_id"])["status"] == "administratively_submitted"
    assert att.get_attempt(db_path, started["attempt_id"])["ended_reason"] == "exam_closed_by_manager"
    assert att.get_result(db_path, started["attempt_id"]) is not None
    with pytest.raises(TrainingError):
        att.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["A"], client_revision=2)
    with pytest.raises(TrainingError):
        att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
```

Add a report test that finalizes twice, records the stored `payload_json` and checksum after the first finalize, and asserts both byte-for-byte unchanged after the second.

- [ ] **Step 2: Run tests to verify failure**

Run: `python3 -m pytest tests/test_training_attempt_lifecycle.py -k "close" tests/test_training_reports.py -q`

Expected: FAIL because close only flips exam status and finalize delegates active attempts to learner submit semantics.

- [ ] **Step 3: Implement administrative terminal transition**

Add `administratively_submit_attempt(db_path, *, unit_code, actor, attempt_id, ended_reason)` in attempt service. It must open `BEGIN IMMEDIATE`, return the existing result if present, CAS only `active` to `administratively_submitted`, score using `score_attempt_with_conn`, create one result, set assignment `completed`, write `administratively_submit_attempt` audit including `ended_reason`, and commit.

Change `close_exam` to first CAS `open -> closed` and commit that transaction. Then select active attempts for the exam, call the helper one at a time, and return only after all have been processed. The second close must return successfully without reopening or rescoring.

Change `finalize_exam` to call `close_exam` only for an open exam. If an older/incomplete closed exam has active attempts, process them with `administratively_submit_attempt(..., ended_reason="finalize_recovery")`; do not use `submit_attempt`. Keep the existing snapshot uniqueness query before insert and never update `payload_json` or `checksum` for an existing revision.

- [ ] **Step 4: Run focused tests**

Run: `python3 -m pytest tests/test_training_attempt_lifecycle.py tests/test_training_reports.py -q`

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add services/training_attempt_service.py services/training_exam_service.py services/training_report_service.py tests/test_training_attempt_lifecycle.py tests/test_training_reports.py
```

### Task 5: Enforce MVP RBAC at Routes

**Files:**
- Modify: `blueprints/training_routes.py`
- Modify: `tests/test_training_routes.py`

- [ ] **Step 1: Write failing route RBAC tests**

```python
def test_editor_can_create_knowledge_but_cannot_publish_or_create_exam(monkeypatch, tmp_path):
    client, db_path = _client_as(monkeypatch, tmp_path, "editor")
    grant_role(db_path, "son_tay", "seed", "editor", "editor")
    assert client.post("/api/training/knowledge", headers={"X-CSRF-Token": "csrf"}, json=VALID_KNOWLEDGE).status_code == 201
    assert client.post("/api/training/questions/version/approve", headers={"X-CSRF-Token": "csrf"}).status_code == 403
    assert client.post("/api/training/exams", headers={"X-CSRF-Token": "csrf"}, json={}).status_code == 403

def test_learner_cannot_read_other_assignment_attempt(monkeypatch, tmp_path):
    client, attempt_id = _learner_client_with_foreign_attempt(monkeypatch, tmp_path)
    response = client.get(f"/api/training/attempts/{attempt_id}")
    assert response.status_code == 403
    assert response.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED"
```

Also test `exam_manager` can approve/publish/template/exam/finalize but gets 403 on knowledge create/import, and dashboard admin can access all operations.

- [ ] **Step 2: Run tests to verify failure**

Run: `python3 -m pytest tests/test_training_routes.py -q`

Expected: FAIL because `_manager_required` currently protects editor operations and grants manager both editor and exam-manager capabilities.

- [ ] **Step 3: Implement role-specific decorators and apply them**

Implement `_role_required(required_role, message)` using dashboard `role == "admin"` or `has_module_role`. Define `_editor_required` and `_manager_required` from it. Apply editor to knowledge create and question import; apply manager to approve, publish, template, exam, assignment, transition, report and Excel. Keep attempt routes learner ownership checks and return every authorization error via `_error_response(TrainingError(...))`. Leave all write methods protected with `@csrf_protect`; replace the hand-written knowledge CSRF branch with that decorator.

- [ ] **Step 4: Run route and permission tests**

Run: `python3 -m pytest tests/test_training_permissions.py tests/test_training_routes.py -q`

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/training_routes.py tests/test_training_routes.py
```

### Task 6: Regression Verification and Documentation Sync

**Files:**
- Modify: `docs/00-doc-index.md`
- Modify: `docs/04-mapping-route-va-du-lieu.md`
- Modify: `docs/08-trang-thai-thuc-thi.md`
- Modify: `docs/12-dao-tao-sat-hach-van-hanh.md`

- [ ] **Step 1: Document the engine integrity milestone**

Add the delta spec and this plan to the training section of `docs/00-doc-index.md`. Add one `/dao-tao-sat-hach` mapping row showing `training.db` per instance and `supports_date = n/a`. Add a concise status entry in docs 08 for deterministic snapshot shuffle, response validation, administrative close and RBAC. In runbook, document that close immediately blocks starts/autosaves and administratively submits active attempts.

- [ ] **Step 2: Compile modified Python files**

Run: `python3 -m py_compile training/errors.py training/migrations.py services/training_attempt_service.py services/training_exam_service.py services/training_report_service.py blueprints/training_routes.py`

Expected: no output and exit code 0.

- [ ] **Step 3: Run required tests**

Run: `python3 -m pytest tests/test_training_*.py -q`

Expected: all training tests pass.

Run: `python3 -m pytest tests/`

Expected: all repository tests pass.

- [ ] **Step 4: Verify diff and commit**

Run: `git diff --check`

Expected: no output and exit code 0.

```bash
git add docs/00-doc-index.md docs/04-mapping-route-va-du-lieu.md docs/08-trang-thai-thuc-thi.md docs/12-dao-tao-sat-hach-van-hanh.md
git push origin feat/dao-tao-sat-hach-mvp
```
