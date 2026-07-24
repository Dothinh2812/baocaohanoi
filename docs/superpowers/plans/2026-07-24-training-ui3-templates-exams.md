# UI-3: Mẫu đề và tổ chức kỳ thi — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

> **STATUS: ✅ ALL TASKS COMPLETE** (commit range `9643043..HEAD`)

**Goal:** Build the full operator workflow: published questions → fixed template → exam → assign users → Ready → Open → track → Close/Cancel/Finalize → recovery summary. Stop before learner UI and report dashboard.

**Architecture:** Flask/Jinja/JS (no framework). New read-model DTOs + list endpoints in `training_exam_service.py`; new routes in `training_routes.py`; two new JS panels (`training-templates.js`, `training-exams.js`); workspace template gains real panels replacing `#training-pending`. Hardening: `Cache-Control: no-store` on management question detail + CAS for review/publish.

**Tech Stack:** Python 3.10, Flask 3.1, SQLite 3.37+ (WAL), Jinja2, vanilla JS (IIFE on `window.TrainingUI`), CSS grid.

**Safety:** Worktree `/home/vtst/dashv4-training` only. Branch `feat/dao-tao-sat-hach-mvp`. Never touch `/home/vtst/dashv4`, never merge `main`, never restart production, never use port 5011, never commit secrets/runtime.

---

## File Structure

**Create:**
- `static/js/training-templates.js` — template panel: list, create form, question picker
- `static/js/training-exams.js` — exam panel: list, create form, assignment, lifecycle, status tracking
- `tests/test_training_ui3_hardening.py` — Cache-Control + CAS concurrency tests
- `tests/test_training_template_exam_routes.py` — route-level tests for templates/exams/users
- `tests/js/test_training_exams_lifecycle.mjs` — JS behavioral test for exam lifecycle buttons

**Modify:**
- `services/training_question_service.py` — CAS: wrap review/publish in `BEGIN IMMEDIATE` + `WHERE ... AND status=?`
- `services/training_exam_service.py` — add `list_templates`, `get_template_detail`, `list_exams`, `get_exam_detail`, `get_exam_assignments_dto`, `list_assignable_users`
- `blueprints/training_routes.py` — add GET endpoints + Cache-Control header on management detail
- `templates/pages/training/index.html` — replace `#training-pending` with real panels for Mẫu đề / Kỳ thi
- `static/css/training.css` — any new component styles needed
- `docs/04-mapping-route-va-du-lieu.md`, `docs/08-trang-thai-thuc-thi.md`, `docs/12-dao-tao-sat-hach-van-hanh.md`, `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-api-contract.md` — doc sync
- `docs/superpowers/plans/2026-07-24-dao-tao-sat-hach-ui-mvp.md` — mark UI-3 done

---

## Task 1: Cache-Control no-store on management question detail

**Files:**
- Modify: `blueprints/training_routes.py` (the `get_question_bank_detail` route, around line 147-155)
- Test: `tests/test_training_ui3_hardening.py`

- [x] **Step 1: Write failing test**

```python
# tests/test_training_ui3_hardening.py
"""Hardening tests for UI-3: Cache-Control + CAS for review/publish."""
import json
from tests.test_training_routes import _client


def test_management_question_detail_has_no_store_header(monkeypatch, tmp_path):
    """GET /api/training/questions/<version_id> must return Cache-Control: no-store."""
    for client in _client(monkeypatch, tmp_path):
        # import a draft question
        resp = client.post(
            "/api/training/questions/import",
            data=json.dumps(_BATCH),
            content_type="application/json",
            headers={"X-CSRF-Token": "csrf"},
        )
        assert resp.status_code == 201
        version_id = resp.get_json()["version_ids"][0]

        resp = client.get(f"/api/training/questions/{version_id}")
        assert resp.status_code == 200
        assert resp.headers.get("Cache-Control") == "no-store"

_BATCH = {
    "schema_version": "1.0",
    "batch": {
        "title": "Test",
        "language": "vi",
        "source_document_version_ids": [],
        "target_audience_codes": ["nvkt"],
        "requested_count": 1,
    },
    "questions": [
        {
            "local_ref": "Q1",
            "type": "single_choice",
            "stem": "1+1=?",
            "options": [
                {"id": "A", "text": "1"},
                {"id": "B", "text": "2"},
                {"id": "C", "text": "3"},
                {"id": "D", "text": "4"},
            ],
            "correct_option_ids": ["B"],
            "explanation": "2",
            "distractor_rationales": {"A": "sai", "B": "dung"},
            "classification": {
                "domain_code": "quality",
                "topic_codes": ["test_topic"],
                "audience_codes": ["nvkt"],
            },
            "difficulty": "easy",
        }
    ],
}
```

- [x] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_training_ui3_hardening.py::test_management_question_detail_has_no_store_header -q`
Expected: FAIL — `Cache-Control` header missing or wrong value.

- [x] **Step 3: Add Cache-Control header to the management detail route**

In `blueprints/training_routes.py`, in `get_question_bank_detail`, add after building the response:

```python
resp = jsonify(detail)
resp.headers["Cache-Control"] = "no-store"
return resp
```

- [x] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_training_ui3_hardening.py::test_management_question_detail_has_no_store_header -q`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add tests/test_training_ui3_hardening.py blueprints/training_routes.py
git commit -m "fix(training): add Cache-Control no-store to management question detail"
```

---

## Task 2: CAS for question review transitions (concurrent approve/reject)

**Files:**
- Modify: `services/training_question_service.py` — `add_review_action` (around line 409-463)
- Test: `tests/test_training_ui3_hardening.py`

- [x] **Step 1: Write failing concurrency test**

```python
# Append to tests/test_training_ui3_hardening.py
import threading
from tests.test_training_exam_states import _publish_questions, _setup
from services import training_question_service as qs


def test_concurrent_approve_and_reject_only_one_succeeds(monkeypatch, tmp_path):
    """Two threads call approve and reject on the same draft question.
    Only one transition must succeed; the other gets CONFLICT."""
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    # Use a fresh draft question by importing another
    # Actually we need a draft question — re-import one
    from tests.test_training_exam_states import VALID_BATCH
    result = qs.import_question_batch(
        db_path, unit_code="son_tay", actor="seed",
        batch=VALID_BATCH, status="draft",
    )
    qid = result["version_ids"][0]

    errors = []
    barrier = threading.Barrier(2)

    def call(action):
        try:
            barrier.wait(timeout=5)
            qs.add_review_action(
                db_path, unit_code="son_tay", actor="seed",
                version_id=qid, action=action,
            )
        except Exception as exc:
            errors.append((action, exc))

    t1 = threading.Thread(target=call, args=("approve",))
    t2 = threading.Thread(target=call, args=("reject",))
    t1.start(); t2.start()
    t1.join(timeout=10); t2.join(timeout=10)

    # Exactly one succeeds, one gets CONFLICT
    success_count = sum(1 for a, e in [(None, None)] if e is None)
    # Check actual DB state
    conn = __import__("sqlite3").connect(db_path)
    row = conn.execute(
        "SELECT review_status FROM question_versions WHERE id=?", (qid,)
    ).fetchone()
    conn.close()
    assert row["review_status"] in ("approved", "rejected")
    # Exactly one thread raised CONFLICT
    conflict_errors = [e for a, e in errors if isinstance(e, Exception)]
    assert len(conflict_errors) == 1, f"Expected 1 conflict, got {conflict_errors}"
```

- [x] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_training_ui3_hardening.py::test_concurrent_approve_and_reject_only_one_succeeds -q`
Expected: FAIL — both threads may succeed because `UPDATE` lacks `AND review_status=?`.

- [x] **Step 3: Add CAS guard to `add_review_action`**

In `services/training_question_service.py`, modify `add_review_action` to:
1. Open `write_connection` and `BEGIN IMMEDIATE` before validation.
2. Replace the final `UPDATE question_versions SET review_status=? ... WHERE id=?` with `WHERE id=? AND review_status=?` (old status).
3. If `cursor.rowcount == 0`, raise `TrainingError(CONFLICT, "Trạng thái đã thay đổi, vui lòng tải lại.", status=409)`.
4. Insert `question_reviews` and `write_audit` within the same transaction, then `commit`.

- [x] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_training_ui3_hardening.py::test_concurrent_approve_and_reject_only_one_succeeds -q`
Expected: PASS

- [x] **Step 5: Run existing question tests to verify no regression**

Run: `python3 -m pytest tests/test_training_questions.py tests/test_training_routes.py -q`
Expected: All PASS.

- [x] **Step 6: Commit**

```bash
git add services/training_question_service.py tests/test_training_ui3_hardening.py
git commit -m "fix(training): CAS review transitions with BEGIN IMMEDIATE"
```

---

## Task 3: CAS for publish (concurrent duplicate publish prevention)

**Files:**
- Modify: `services/training_question_service.py` — `publish_question_version` (around line 466-518)
- Test: `tests/test_training_ui3_hardening.py`

- [x] **Step 1: Write failing concurrency test**

```python
# Append to tests/test_training_ui3_hardening.py
def test_concurrent_duplicate_publish_prevents_both(monkeypatch, tmp_path):
    """Two threads publish the same approved question.
    Only one must succeed; the other gets CONFLICT."""
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    # Get a fresh approved-but-unpublished question
    from tests.test_training_exam_states import VALID_BATCH
    result = qs.import_question_batch(
        db_path, unit_code="son_tay", actor="seed",
        batch=VALID_BATCH, status="draft",
    )
    qid = result["version_ids"][0]
    qs.add_review_action(db_path, unit_code="son_tay", actor="seed",
                         version_id=qid, action="approve")

    errors = []
    barrier = threading.Barrier(2)

    def publish():
        try:
            barrier.wait(timeout=5)
            qs.publish_question_version(
                db_path, unit_code="son_tay", actor="seed", version_id=qid,
            )
        except Exception as exc:
            errors.append(exc)

    t1 = threading.Thread(target=publish)
    t2 = threading.Thread(target=publish)
    t1.start(); t2.start()
    t1.join(timeout=10); t2.join(timeout=10)

    assert len(errors) == 1, f"Expected exactly 1 conflict, got {len(errors)}"
    assert errors[0].code in ("CONFLICT", "QUESTION_ALREADY_PUBLISHED")


def test_audit_not_duplicated_for_failed_transition(monkeypatch, tmp_path):
    """A failed CAS transition must not write audit."""
    db_path = _setup(monkeypatch, tmp_path)
    from tests.test_training_exam_states import VALID_BATCH
    result = qs.import_question_batch(
        db_path, unit_code="son_tay", actor="seed",
        batch=VALID_BATCH, status="draft",
    )
    qid = result["version_ids"][0]
    qs.add_review_action(db_path, unit_code="son_tay", actor="seed",
                         version_id=qid, action="approve")
    # Publish successfully
    qs.publish_question_version(db_path, unit_code="son_tay", actor="seed",
                                version_id=qid)
    import sqlite3
    conn = sqlite3.connect(db_path); conn.row_factory = sqlite3.Row
    before = conn.execute(
        "SELECT COUNT(*) as c FROM training_audit_log WHERE entity_id=?", (qid,)
    ).fetchone()["c"]
    # Try to publish again — should fail
    try:
        qs.publish_question_version(db_path, unit_code="son_tay", actor="seed",
                                    version_id=qid)
        assert False, "Should have raised"
    except Exception:
        pass
    after = conn.execute(
        "SELECT COUNT(*) as c FROM training_audit_log WHERE entity_id=?", (qid,)
    ).fetchone()["c"]
    conn.close()
    assert after == before, f"Audit grew from {before} to {after} on failed transition"
```

- [x] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_training_ui3_hardening.py -k publish -q`
Expected: FAIL.

- [x] **Step 3: Add CAS guard to `publish_question_version`**

In `services/training_question_service.py`, modify `publish_question_version`:
1. Open `write_connection` and `BEGIN IMMEDIATE`.
2. Re-read `publication_status` within the transaction.
3. If already `published`, raise `TrainingError(CONFLICT, "Câu hỏi đã phát hành.", status=409)` (idempotent retry on published = error per the spec).
4. Duplicate check via `normalized_stem_hash` within the same transaction.
5. `UPDATE ... SET publication_status='published' WHERE id=? AND publication_status='unpublished'`.
6. If `cursor.rowcount == 0`, raise CONFLICT.
7. `write_audit` + `commit` within the same transaction.

- [x] **Step 4: Run tests to verify they pass**

Run: `python3 -m pytest tests/test_training_ui3_hardening.py -k publish -q`
Expected: PASS.

- [x] **Step 5: Run all training tests**

Run: `python3 -m pytest tests/test_training_*.py -q`
Expected: All PASS.

- [x] **Step 6: Commit**

```bash
git add services/training_question_service.py tests/test_training_ui3_hardening.py
git commit -m "fix(training): CAS publish with BEGIN IMMEDIATE and duplicate guard"
```

---

## Task 4: Template + Exam read-model service functions

**Files:**
- Modify: `services/training_exam_service.py`
- Test: `tests/test_training_template_exam_routes.py` (new)

- [x] **Step 1: Write failing service tests**

```python
# tests/test_training_template_exam_routes.py
"""Tests for template/exam read-model DTOs and list endpoints."""
import json
from tests.test_training_routes import _client, QUESTION_BATCH
from tests.test_training_exam_states import _setup, _publish_questions
from services import training_exam_service as exams


def test_list_templates_returns_dto(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = exams.create_template(
        db_path, unit_code="son_tay", actor="seed",
        code="TPL-1", title="Test template",
        target_audience_code="nvkt",
        question_version_ids=version_ids,
        duration_seconds=1800, pass_score_percent=80,
    )
    result = exams.list_templates(db_path, page=1, page_size=25)
    assert "items" in result and "total" in result
    item = result["items"][0]
    # Allowlist: no raw columns leaked
    assert "id" in item
    assert "code" in item
    assert "title" in item
    assert "target_audience_code" in item
    assert "total_questions" in item
    assert "duration_seconds" in item
    assert "pass_score_percent" in item
    assert "shuffle_questions" in item
    assert "shuffle_options" in item
    assert "locked" in item
    assert "created_by" in item
    assert "created_at_ms" in item
    # No sensitive/technical columns
    assert "exam_events" not in str(item)


def test_get_template_detail_includes_items(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = exams.create_template(
        db_path, unit_code="son_tay", actor="seed",
        code="TPL-1", title="Test template",
        target_audience_code="nvkt",
        question_version_ids=version_ids,
        duration_seconds=1800, pass_score_percent=80,
    )
    detail = exams.get_template_detail(db_path, tpl["id"])
    assert detail["id"] == tpl["id"]
    assert detail["code"] == "TPL-1"
    assert len(detail["items"]) == len(version_ids)
    item = detail["items"][0]
    assert "sequence_number" in item
    assert "question_version_id" in item
    assert "stem" in item  # snapshot stem for display
    assert "type" in item
    assert "difficulty" in item


def test_list_exams_returns_dto(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = exams.create_template(
        db_path, unit_code="son_tay", actor="seed",
        code="TPL-1", title="Test",
        target_audience_code="nvkt",
        question_version_ids=version_ids,
        duration_seconds=1800, pass_score_percent=80,
    )
    now = __import__("time").time() * 1000
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="seed",
        code="EX-1", title="Test exam", template_id=tpl["id"],
        target_audience_code="nvkt",
        start_at_ms=int(now), end_at_ms=int(now + 3600000),
        duration_seconds=1800, pass_score_percent=80,
    )
    result = exams.list_exams(db_path, page=1, page_size=25)
    assert "items" in result
    item = result["items"][0]
    for key in ("id", "code", "title", "template_id", "target_audience_code",
                "status", "start_at_ms", "end_at_ms", "duration_seconds",
                "pass_score_percent", "created_by", "created_at_ms"):
        assert key in item, f"Missing key: {key}"


def test_get_exam_detail_includes_summary(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = exams.create_template(
        db_path, unit_code="son_tay", actor="seed",
        code="TPL-1", title="Test",
        target_audience_code="nvkt",
        question_version_ids=version_ids,
        duration_seconds=1800, pass_score_percent=80,
    )
    now = __import__("time").time() * 1000
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="seed",
        code="EX-1", title="Test", template_id=tpl["id"],
        target_audience_code="nvkt",
        start_at_ms=int(now), end_at_ms=int(now + 3600000),
        duration_seconds=1800, pass_score_percent=80,
    )
    detail = exams.get_exam_detail(db_path, exam["id"])
    assert detail["id"] == exam["id"]
    assert detail["status"] == "draft"
    assert "template" in detail
    assert "assignment_summary" in detail
    assert detail["assignment_summary"]["total"] == 0


def test_list_assignable_users_excludes_passwords(monkeypatch, tmp_path):
    """User API must not return password hashes."""
    from services import training_exam_service as svc
    from unittest.mock import patch

    db_path = _setup(monkeypatch, tmp_path)
    # Stub auth.get_all_users to return fake data with password
    fake_users = [
        {"username": "u1", "name": "User One", "role": "user",
         "password": "SECRET_HASH", "is_active": 1, "is_first_login": 0},
        {"username": "u2", "name": "User Two", "role": "user",
         "password": "SECRET_HASH2", "is_active": 1, "is_first_login": 0},
    ]
    monkeypatch.setattr("services.training_exam_service._get_all_users",
                        lambda: fake_users)
    result = svc.list_assignable_users(db_path, q="")
    assert len(result) == 2
    for u in result:
        assert "username" in u
        assert "display_name" in u
        assert "password" not in u
        assert "password_hash" not in u
```

- [x] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_training_template_exam_routes.py -q`
Expected: FAIL — functions `list_templates`, `get_template_detail`, `list_exams`, `get_exam_detail`, `list_assignable_users` don't exist yet.

- [x] **Step 3: Implement the read-model functions**

In `services/training_exam_service.py`, add:

```python
def list_templates(db_path, *, page=1, page_size=25):
    """Return paginated template list DTO."""
    # read_connection, SELECT allowlist from exam_templates
    # return {"items": [...], "page", "page_size", "total"}

def get_template_detail(db_path, template_id):
    """Return template + items with question stem/type/difficulty snapshots."""
    # read_connection, JOIN exam_template_items + question_versions

def list_exams(db_path, *, page=1, page_size=25, status=None):
    """Return paginated exam list DTO."""
    # read_connection, SELECT allowlist from exam_events

def get_exam_detail(db_path, exam_id):
    """Return exam + template info + assignment_summary aggregation."""
    # read_connection, aggregate assignment counts via GROUP BY

def get_exam_assignments_dto(db_path, exam_id):
    """Return list of assignment DTOs (no sensitive data)."""
    # read_connection, SELECT allowlist from exam_assignments

def list_assignable_users(db_path, *, q=""):
    """Return users from username.xlsx stripped of passwords."""
    # read auth.get_all_users(), filter/transform
```

Key rules:
- `list_templates` DTO keys: `id, code, title, target_audience_code, total_questions, duration_seconds, pass_score_percent, shuffle_questions, shuffle_options, locked, created_by, created_at_ms`
- `get_template_detail` adds `items: [{sequence_number, question_version_id, stem, type, difficulty, section_label, points}]`
- `list_exams` DTO keys: `id, code, title, template_id, target_audience_code, status, start_at_ms, end_at_ms, duration_seconds, pass_score_percent, reveal_answers_after_finalize, created_by, created_at_ms, finalized_at_ms`
- `get_exam_detail` adds `template: {code, title}`, `assignment_summary: {total, assigned, completed, expired, cancelled, in_progress}`
- `list_assignable_users` returns `[{username, display_name}]` — strips ALL other fields including `password`, `role`, `is_active`, etc.

- [x] **Step 4: Run tests to verify they pass**

Run: `python3 -m pytest tests/test_training_template_exam_routes.py -q`
Expected: PASS.

- [x] **Step 5: Commit**

```bash
git add services/training_exam_service.py tests/test_training_template_exam_routes.py
git commit -m "feat(training-api): expose template and exam read models"
```

---

## Task 5: Route endpoints for template/exam reads + user API

**Files:**
- Modify: `blueprints/training_routes.py`
- Test: `tests/test_training_template_exam_routes.py`

- [x] **Step 1: Write failing route tests**

Append route-level tests to `tests/test_training_template_exam_routes.py`:

```python
def test_list_templates_route(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        resp = client.get("/api/training/templates")
        assert resp.status_code == 200
        data = resp.get_json()
        assert "items" in data and "total" in data

def test_get_template_detail_route_404(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        resp = client.get("/api/training/templates/nonexistent")
        assert resp.status_code == 404

def test_list_exams_route(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        resp = client.get("/api/training/exams")
        assert resp.status_code == 200

def test_get_exam_detail_route(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        db_path = str(tmp_path / "training.db")  # _client already seeds this
        version_ids = _publish_questions(db_path)
        tpl = exams.create_template(
            db_path, unit_code="son_tay", actor="admin",
            code="TPL", title="T", target_audience_code="nvkt",
            question_version_ids=version_ids,
            duration_seconds=1800, pass_score_percent=80,
        )
        import time
        now = int(time.time() * 1000)
        exam = exams.create_exam(
            db_path, unit_code="son_tay", actor="admin",
            code="EX", title="E", template_id=tpl["id"],
            target_audience_code="nvkt",
            start_at_ms=now, end_at_ms=now + 3600000,
            duration_seconds=1800, pass_score_percent=80,
        )
        resp = client.get(f"/api/training/exams/{exam['id']}")
        assert resp.status_code == 200
        assert resp.get_json()["status"] == "draft"

def test_list_assignments_route(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        resp = client.get("/api/training/exams/fake/assignments")
        assert resp.status_code == 200
        assert "items" in resp.get_json()

def test_list_users_route_no_password(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        # Need to stub auth for user list
        from unittest.mock import patch
        monkeypatch.setattr(
            "services.training_exam_service._get_all_users",
            lambda: [{"username": "u1", "name": "User 1", "password": "hash"}]
        )
        resp = client.get("/api/training/users?q=u")
        assert resp.status_code == 200
        data = resp.get_json()
        assert len(data["items"]) == 1
        assert "password" not in str(data)

def test_template_routes_require_exam_manager(monkeypatch, tmp_path):
    from tests.test_training_routes import _role_client
    for client, db_path in _role_client(monkeypatch, tmp_path, roles=["learner"]):
        resp = client.get("/api/training/templates")
        assert resp.status_code == 403
```

- [x] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_training_template_exam_routes.py -k route -q`
Expected: FAIL — routes don't exist yet.

- [x] **Step 3: Add route functions**

In `blueprints/training_routes.py`, add these GET endpoints (mirroring `list_question_bank`):

```python
@training_bp.route("/api/training/templates")
@_exam_manager_required
def list_templates():
    page = max(1, request.args.get("page", 1, type=int))
    page_size = min(100, max(1, request.args.get("page_size", 25, type=int)))
    try:
        result = exams.list_templates(config.TRAINING_DB_PATH, page=page, page_size=page_size)
        return jsonify(result)
    except TrainingError as exc:
        return _error_response(exc)

@training_bp.route("/api/training/templates/<template_id>")
@_exam_manager_required
def get_template_detail(template_id):
    try:
        result = exams.get_template_detail(config.TRAINING_DB_PATH, template_id)
        return jsonify(result)
    except TrainingError as exc:
        return _error_response(exc)

@training_bp.route("/api/training/exams")
@_exam_manager_required
def list_exams():
    page = max(1, request.args.get("page", 1, type=int))
    page_size = min(100, max(1, request.args.get("page_size", 25, type=int)))
    status = request.args.get("status") or None
    try:
        result = exams.list_exams(config.TRAINING_DB_PATH, page=page, page_size=page_size, status=status)
        return jsonify(result)
    except TrainingError as exc:
        return _error_response(exc)

@training_bp.route("/api/training/exams/<exam_id>")
@_exam_manager_required
def get_exam_detail(exam_id):
    try:
        result = exams.get_exam_detail(config.TRAINING_DB_PATH, exam_id)
        resp = jsonify(result)
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except TrainingError as exc:
        return _error_response(exc)

@training_bp.route("/api/training/exams/<exam_id>/assignments")
@_exam_manager_required
def list_exam_assignments(exam_id):
    try:
        result = exams.get_exam_assignments_dto(config.TRAINING_DB_PATH, exam_id)
        return jsonify({"items": result})
    except TrainingError as exc:
        return _error_response(exc)

@training_bp.route("/api/training/users")
@_exam_manager_required
def list_assignable_users():
    q = request.args.get("q", "")
    try:
        result = exams.list_assignable_users(config.TRAINING_DB_PATH, q=q)
        return jsonify({"items": result})
    except TrainingError as exc:
        return _error_response(exc)
```

Also add `list_assignable_users` service that calls `auth.get_all_users()` internally (via a `_get_all_users` indirection for testability):

```python
# In services/training_exam_service.py
def _get_all_users():
    """Indirection for testing; calls auth.get_all_users()."""
    from auth import get_all_users
    return get_all_users()

def list_assignable_users(db_path, *, q=""):
    users = _get_all_users()
    result = []
    for u in users:
        if u.get("is_active", 1) == 0:
            continue
        display_name = u.get("name") or u.get("username", "")
        username = u.get("username", "")
        if q and q.lower() not in username.lower() and q.lower() not in display_name.lower():
            continue
        result.append({"username": username, "display_name": display_name})
    return result
```

- [x] **Step 4: Run tests to verify they pass**

Run: `python3 -m pytest tests/test_training_template_exam_routes.py -q`
Expected: All PASS.

- [x] **Step 5: Commit**

```bash
git add blueprints/training_routes.py services/training_exam_service.py tests/test_training_template_exam_routes.py
git commit -m "feat(training-api): expose template/exam/user read endpoints"
```

---

## Task 6: Template UI panel (list + create + question picker)

**Files:**
- Create: `static/js/training-templates.js`
- Modify: `templates/pages/training/index.html`
- Modify: `static/css/training.css` (if needed)

- [x] **Step 1: Add template panel section to workspace template**

In `templates/pages/training/index.html`, replace the Mẫu đề `#training-pending` href with a real panel:

```html
<!-- Inside the workspace panel, after the question-bank section -->
<section class="training-templates-panel" id="training-templates"
         data-can-edit="{{ 'true' if 'exam_manager' in module_roles or 'admin' in module_roles else 'false' }}"
         aria-labelledby="training-templates-title" hidden>
  <h3 id="training-templates-title">Mẫu đề</h3>

  <div class="training-action-row">
    <button type="button" class="training-action" id="template-create-btn">Tạo mẫu đề mới</button>
    <button type="button" class="training-action-secondary" id="template-list-btn">Danh sách mẫu đề</button>
  </div>

  <!-- Create form -->
  <div id="template-create-form" hidden>
    <form id="template-form">
      <div class="training-question-filters">
        <label>Mã đề <input type="text" name="code" required></label>
        <label>Tiêu đề <input type="text" name="title" required></label>
        <label>Nhóm đối tượng
          <select name="target_audience_code" id="template-audience" required></select>
        </label>
        <label>Thời lượng (phút) <input type="number" name="duration_minutes" value="30" min="1" required></label>
        <label>Điểm đạt (%) <input type="number" name="pass_score_percent" value="80" min="0" max="100" step="0.1" required></label>
      </div>
      <div class="training-question-filters">
        <label><input type="checkbox" name="shuffle_questions"> Trộn câu hỏi</label>
        <label><input type="checkbox" name="shuffle_options"> Trộn đáp án</label>
      </div>
      <h4>Chọn câu hỏi đã phát hành</h4>
      <div class="training-question-filters" id="template-question-filters">
        <label>Từ khóa <input type="text" name="q" placeholder="Tìm câu hỏi..."></label>
        <label>Lĩnh vực <select name="domain"><option value="">Tất cả</option></select></label>
        <label>Chủ đề <input type="text" name="topic"></label>
        <button type="submit" class="training-action-secondary">Lọc</button>
      </div>
      <p id="template-selected-count">Đã chọn: 0 câu</p>
      <div id="template-question-list" role="region" aria-live="polite"></div>
      <nav id="template-question-pagination" aria-label="Phân trang câu hỏi"></nav>
      <div class="training-form-errors" id="template-errors" role="alert" aria-live="assertive"></div>
      <div class="training-action-row">
        <button type="button" class="training-action" id="template-submit-btn">Tạo mẫu đề</button>
        <button type="button" class="training-action-secondary" id="template-cancel-btn">Hủy</button>
      </div>
    </form>
  </div>

  <!-- List -->
  <div id="template-list-view">
    <table class="training-question-table" id="template-list-table">
      <thead><tr><th>Mã</th><th>Tiêu đề</th><th>Nhóm đối tượng</th><th>Số câu</th><th>Thời lượng</th><th>Trạng thái</th><th></th></tr></thead>
      <tbody id="template-list-body"></tbody>
    </table>
    <nav id="template-pagination" aria-label="Phân trang mẫu đề"></nav>
  </div>

  <!-- Detail -->
  <div id="template-detail" hidden tabindex="-1" aria-live="polite"></div>
</section>
```

Update nav anchors: Mẫu đề → `href="#training-templates"`, Kỳ thi → `href="#training-exams"`, Giao bài → merged into exam panel.

Add script include:
```html
{% if 'exam_manager' in module_roles or 'admin' in module_roles %}
<script src="{{ url_for('static', filename='js/training-templates.js') }}"></script>
<script src="{{ url_for('static', filename='js/training-exams.js') }}"></script>
{% endif %}
```

- [x] **Step 2: Implement `training-templates.js`**

Follow the IIFE pattern from `training-question-bank.js`:
- Guard on `#training-templates` panel presence.
- List templates via `GET /api/training/templates`.
- Create form: collect fields, fetch published questions via `GET /api/training/questions?status=published`, let user select with checkboxes, validate all selected match audience, POST to `/api/training/templates`.
- Convert `duration_minutes` → `duration_seconds` in JS before sending.
- Render question list with stem, type, difficulty, topic, audience, and a checkbox.
- Track selected question IDs in a `Set`, update count display.
- After create success: toast, switch to list view, reload list.
- All DOM via `textContent`/`createElement`.

- [x] **Step 3: Run existing workspace UI tests**

Run: `python3 -m pytest tests/test_training_workspace_ui.py -q`
Expected: May need updating if test checks for `#training-pending`. Update tests if needed.

- [x] **Step 4: Commit**

```bash
git add templates/pages/training/index.html static/js/training-templates.js static/css/training.css tests/
git commit -m "feat(training-ui): add fixed template workflow"
```

---

## Task 7: Exam UI panel (list + create + assignment + lifecycle)

**Files:**
- Create: `static/js/training-exams.js`
- Modify: `templates/pages/training/index.html`

- [x] **Step 1: Add exam panel section to workspace template**

Add to `index.html`:

```html
<section class="training-exams-panel" id="training-exams"
         aria-labelledby="training-exams-title" hidden>
  <h3 id="training-exams-title">Kỳ thi</h3>

  <div class="training-action-row">
    <button type="button" class="training-action" id="exam-create-btn">Tạo kỳ thi mới</button>
    <button type="button" class="training-action-secondary" id="exam-list-btn">Danh sách kỳ thi</button>
  </div>

  <!-- Create form -->
  <div id="exam-create-form" hidden>
    <form id="exam-form">
      <div class="training-question-filters">
        <label>Mã kỳ thi <input type="text" name="code" required></label>
        <label>Tiêu đề <input type="text" name="title" required></label>
        <label>Mẫu đề
          <select name="template_id" id="exam-template-select" required></select>
        </label>
        <label>Nhóm đối tượng <select name="target_audience_code" id="exam-audience" readonly></select></label>
      </div>
      <div class="training-question-filters">
        <label>Mô tả <input type="text" name="description"></label>
        <label>Thời lượng (phút) <input type="number" name="duration_minutes" value="30" min="1" required></label>
        <label>Điểm đạt (%) <input type="number" name="pass_score_percent" value="80" min="0" max="100" step="0.1" required></label>
      </div>
      <div class="training-question-filters">
        <label>Bắt đầu (giờ VN)
          <input type="datetime-local" name="start_at_local" required>
        </label>
        <label>Kết thúc (giờ VN)
          <input type="datetime-local" name="end_at_local" required>
        </label>
        <label><input type="checkbox" name="reveal_answers_after_finalize" checked> Công bố đáp án sau chốt</label>
      </div>
      <div class="training-form-errors" id="exam-errors" role="alert" aria-live="assertive"></div>
      <div class="training-action-row">
        <button type="button" class="training-action" id="exam-submit-btn">Tạo kỳ thi</button>
        <button type="button" class="training-action-secondary" id="exam-cancel-btn">Hủy</button>
      </div>
    </form>
  </div>

  <!-- List -->
  <div id="exam-list-view">
    <table class="training-question-table">
      <thead><tr><th>Mã</th><th>Tiêu đề</th><th>Trạng thái</th><th>Nhóm đối tượng</th><th>Bắt đầu</th><th>Kết thúc</th><th></th></tr></thead>
      <tbody id="exam-list-body"></tbody>
    </table>
    <nav id="exam-pagination" aria-label="Phân trang kỳ thi"></nav>
  </div>

  <!-- Detail -->
  <div id="exam-detail" hidden tabindex="-1" aria-live="polite"></div>
</section>
```

- [x] **Step 2: Implement `training-exams.js`**

Structure:
- **List**: GET `/api/training/exams`, render table with status badge, start/end times in VN timezone.
- **Create form**:
  - Load templates via `GET /api/training/templates` into the select.
  - On template select, auto-fill audience (read-only, derived from template).
  - Convert `datetime-local` (browser local) → epoch ms. Use `new Date(localString).getTime()`.
  - POST `/api/training/exams`.
- **Detail**:
  - GET `/api/training/exams/<id>` → render status, template info, time config, assignment_summary.
  - GET `/api/training/exams/<id>/assignments` → render assignment table with status per person.
  - **Lifecycle buttons** based on status:
    - `draft`: [Sẵn sàng (Ready)] [Hủy (Cancel)]
    - `ready`: [Mở (Open)] [Hủy (Cancel)]
    - `open`: [Đóng (Close)]
    - `closed`: [Chốt (Finalize)]
    - `cancelled`: no buttons
    - `finalized`: no buttons, show revision
  - Each button: `TrainingUI.confirm("...?")` → POST → handle success/error → reload detail + list.
  - **Close**: modal warns about active attempts; display `recovery_summary` after.
  - **Finalize**: display `recovery_summary` + `blocking_attempts` if present; on success display revision.
  - **Assignment sub-panel**: shown only when status is `draft`/`ready`.
    - Search users via `GET /api/training/users?q=...`.
    - Select multiple via checkboxes.
    - POST `/api/training/exams/<id>/assignments` with selected usernames + audience.
    - Handle 409 ASSIGNMENT_ALREADY_EXISTS gracefully.
  - Double-click prevention: disable button after click, re-enable on response.
  - All DOM via `textContent`/`createElement`.

- [x] **Step 3: Run workspace + route tests**

Run: `python3 -m pytest tests/test_training_workspace_ui.py tests/test_training_routes.py -q`
Expected: PASS.

- [x] **Step 4: Commit**

```bash
git add templates/pages/training/index.html static/js/training-exams.js tests/
git commit -m "feat(training-ui): add exam assignment workflow"
```

---

## Task 8: Exam lifecycle controls and status tracking

This task focuses on the JS that renders lifecycle buttons, handles transitions, and displays recovery/blocker information. If Task 7 already covers the full panel, this task adds the remaining lifecycle-specific tests and polish.

**Files:**
- Modify: `static/js/training-exams.js`
- Create: `tests/js/test_training_exams_lifecycle.mjs`
- Modify: `tests/test_training_question_bank_client.py` (or create `tests/test_training_exams_client.py`)

- [x] **Step 1: Write the JS behavioral test**

Create `tests/js/test_training_exams_lifecycle.mjs` following the pattern of `test_question_bank_review_actions.mjs`:
- Load `static/js/training-exams.js` via eval.
- Mock DOM + `window.TrainingUI`.
- Test that lifecycle buttons appear/disappear correctly for each status:
  - `draft` → buttons: "Sẵn sàng", "Hủy"
  - `ready` → buttons: "Mở", "Hủy"
  - `open` → buttons: "Đóng"
  - `closed` → buttons: "Chốt"
  - `cancelled` → no buttons
- Assert each action label contains a Vietnamese confirm prompt.
- Assert double-click prevention (button disabled after click).

- [x] **Step 2: Create pytest bridge**

Create `tests/test_training_exams_client.py` mirroring `test_training_question_bank_client.py`:
- Skip if `node` absent.
- Run the `.mjs` script and assert exit 0.

- [x] **Step 3: Run JS test**

Run: `python3 -m pytest tests/test_training_exams_client.py -q`
Expected: PASS (if all lifecycle buttons match).

- [x] **Step 4: Commit**

```bash
git add static/js/training-exams.js tests/js/test_training_exams_lifecycle.mjs tests/test_training_exams_client.py
git commit -m "test(training-ui): cover exam lifecycle controls"
```

---

## Task 9: Workspace navigation (tab switching)

**Files:**
- Modify: `templates/pages/training/index.html`
- Modify: `static/js/training-ui.js` (add panel switcher) or inline in template
- Modify: `tests/test_training_workspace_ui.py`

- [x] **Step 1: Add nav switching JS**

In `training-ui.js` or inline in the template `<script>`, add a simple panel switcher:

```javascript
// Panel switching: clicking a nav item shows the corresponding panel, hides others.
document.querySelectorAll('.training-workspace-nav a[data-panel]').forEach(function(link) {
    link.addEventListener('click', function(e) {
        e.preventDefault();
        var targetId = link.getAttribute('data-panel');
        document.querySelectorAll('.training-workspace-panel').forEach(function(p) {
            p.hidden = (p.id !== targetId);
        });
        document.querySelectorAll('.training-nav-item').forEach(function(n) {
            n.classList.remove('active');
        });
        link.querySelector('.training-nav-item')?.classList.add('active') 
            || link.classList.add('active');
    });
});
// Show first panel by default (overview or question-bank)
```

Update nav items to use `data-panel` attributes and `href` fallbacks.

- [x] **Step 2: Update workspace UI tests**

Update `tests/test_training_workspace_ui.py`:
- Assert `#training-templates` panel renders for exam_manager/admin.
- Assert `#training-exams` panel renders for exam_manager/admin.
- Assert learner page does NOT contain these panels.
- Assert `training-templates.js` + `training-exams.js` scripts are included for exam_manager.
- Assert nav items point to real panel IDs, not `#training-pending` (for Mẫu đề / Kỳ thi).

- [x] **Step 3: Run workspace tests**

Run: `python3 -m pytest tests/test_training_workspace_ui.py -q`
Expected: PASS.

- [x] **Step 4: Commit**

```bash
git add templates/pages/training/index.html static/js/training-ui.js tests/test_training_workspace_ui.py
git commit -m "feat(training-ui): wire workspace panel navigation"
```

---

## Task 10: Comprehensive route + integration tests

**Files:**
- Modify: `tests/test_training_template_exam_routes.py`

- [x] **Step 1: Add integration tests**

Add tests covering:
- Template CRUD: create → list → detail → locked check.
- Template validation: empty question_version_ids, unpublished questions, duplicates, audience mismatch.
- Exam full lifecycle: create draft → ready → open → close → recovery_summary displayed → finalize → revision shown.
- Assignment: create exam → assign multiple users → 409 on duplicate → blocked after open.
- User API: no password in response.
- RBAC: learner cannot access template/exam/user endpoints (403).
- CSRF: POST without header → 400.
- Finalize blocking: force a bad attempt snapshot → finalize returns blocking_attempts.

- [x] **Step 2: Run all training tests**

Run: `python3 -m pytest tests/test_training_*.py -q`
Expected: All PASS.

- [x] **Step 3: Commit**

```bash
git add tests/test_training_template_exam_routes.py
git commit -m "test(training-ui): cover template and exam workflows"
```

---

## Task 11: Manual smoke test

- [x] **Step 1: Check port 5111 availability**

Run: `lsof -i :5111 || echo "port free"`
Expected: port free (no output from lsof, or "port free").

- [x] **Step 2: Start dev server**

```bash
cd /home/vtst/dashv4-training && \
DASHV4_PORT=5111 \
DASHV4_UNIT_CODE=son_tay \
DASHV4_RUNTIME_DIR=/tmp/dashv4-training-runtime \
DASHV4_TRAINING_DB_PATH=/tmp/dashv4-training.db \
python3 dashboard.py &
```

- [x] **Step 3: Smoke test via curl or browser**

1. Migrate DB: `python3 -m training.cli db-migrate` (with same env).
2. Seed: insert a question via API, approve, publish.
3. Create template via API.
4. Create exam via API.
5. Assign users.
6. Ready → Open (if within time window).
7. Close → verify recovery_summary.
8. Finalize → verify revision.
9. Cancel a different draft exam.

- [x] **Step 4: Kill dev server**

```bash
kill %1  # or find the PID and kill it
```

- [x] **Step 5: Record smoke results**

No commit needed; document in the final report.

---

## Task 12: Doc sync + checkpoint

**Files:**
- Modify: `docs/04-mapping-route-va-du-lieu.md`
- Modify: `docs/08-trang-thai-thuc-thi.md`
- Modify: `docs/12-dao-tao-sat-hach-van-hanh.md`
- Modify: `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-api-contract.md`
- Modify: `docs/superpowers/plans/2026-07-24-dao-tao-sat-hach-ui-mvp.md`

- [x] **Step 1: Update docs**

Update all doc files to reflect new routes:
- `docs/04`: add new template/exam/user GET endpoints with `supports_date = n/a`.
- `docs/08`: mark template/exam management as implemented.
- `docs/12`: document the new APIs in the operations section.
- API contract: add the new read endpoints to §5.
- Plan: mark UI-3 checkpoint as complete.

- [x] **Step 2: Run full verification**

```bash
cd /home/vtst/dashv4-training
python3 -m pytest tests/test_training_*.py -q
python3 -m pytest tests/ -q
python3 -m py_compile services/training_exam_service.py services/training_question_service.py blueprints/training_routes.py
git diff --check
```

- [x] **Step 3: Commit and push**

```bash
git add docs/
git commit -m "docs(training): record UI-3 checkpoint"
git push origin feat/dao-tao-sat-hach-mvp
```

---

## Self-Review Checklist

**Spec coverage:**
- [x] A1: Cache-Control no-store on management question detail — Task 1
- [x] A2: CAS for review/publish — Tasks 2-3
- [x] B1: Read APIs (templates, questions for picker) — Tasks 4-5
- [x] B2: Template create form — Task 6
- [x] B3: Question picker with filter/select — Task 6
- [x] B4: Validation (empty, dup, unpublished, audience) — Tasks 4-5 (service), Task 6 (UI)
- [x] C1: Read APIs (exams, assignments, users) — Tasks 4-5
- [x] C2: Exam create form — Task 7
- [x] C3: Assign users — Task 7
- [x] C4: Lifecycle buttons per state — Task 7-8
- [x] C5: Open requires manual click — Task 7
- [x] C6: Close/Cancel/Finalize with recovery — Task 7-8
- [x] C7: Status tracking (assignment_summary) — Tasks 4-5
- [x] D: Workspace + responsive — Task 9
- [x] F: Tests (hardening, template, exam, RBAC, CSRF, client) — Tasks 1-3, 4-5, 10, 8
- [x] G: Verification — Task 12
- [x] H: Docs — Task 12
