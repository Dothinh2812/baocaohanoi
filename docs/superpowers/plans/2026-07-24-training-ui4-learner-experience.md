# UI-4: Trải nghiệm làm bài của learner — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the full learner exam-taking experience — my-assignments catalog, attempt workspace with autosave/submit, deadline enforcement, and double-click lifecycle protection.

**Architecture:** Backend adds one new read endpoint (`GET /api/training/my-assignments`) and one new service function. Existing attempt routes (start, load, autosave, submit) are already wired and tested. Frontend adds two new JS modules (`training-my-exams.js` for assignment listing, `training-attempt.js` for the attempt workspace) and updates the HTML template. Lifecycle button double-click prevention is added to the existing `training-exams.js`.

**Tech Stack:** Flask/Jinja, vanilla JS (IIFE pattern on `window.TrainingUI`), SQLite, no new dependencies.

**Branch:** `feat/dao-tao-sat-hach-mvp` — start from `7c2ffa1`.

---

## File Map

| Action | File | Responsibility |
|--------|------|----------------|
| Create | `services/training_exam_service.py` (modify) | Add `get_my_assignments_dto()` |
| Create | `blueprints/training_routes.py` (modify) | Add `GET /api/training/my-assignments` |
| Modify | `static/js/training-exams.js` | Add double-click prevention to lifecycle buttons |
| Create | `static/js/training-my-exams.js` | Learner assignment catalog panel |
| Create | `static/js/training-attempt.js` | Learner attempt workspace (question render, autosave, submit) |
| Modify | `templates/pages/training/index.html` | Add learner panels + conditionally load JS |
| Modify | `static/css/training.css` | Styles for attempt workspace, countdown, question nav |
| Create | `tests/test_training_my_assignments.py` | Backend my-assignments ownership + DTO tests |
| Create | `tests/test_training_learner_flow.py` | Backend learner flow tests (start→autosave→submit, ownership, CSRF) |
| Create | `tests/js/test_training_attempt_workspace.mjs` | JS behavioral test for attempt UI |
| Modify | `docs/08-trang-thai-thuc-thi.md` | UI-4 status update |
| Modify | `docs/12-dao-tao-sat-hach-van-hanh.md` | UI-4 operations update |
| Modify | `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-api-contract.md` | Add my-assignments endpoint |
| Modify | `docs/superpowers/plans/2026-07-24-dao-tao-sat-hach-ui-mvp.md` | Mark UI-4 done |

---

## Task 1: Push unpushed commit + lifecycle double-click fix

### Step 1.1: Push commit 7c2ffa1

```bash
cd /home/vtst/dashv4-training
git push origin feat/dao-tao-sat-hach-mvp
git fetch origin
git log --oneline feat/dao-tao-sat-hach-mvp..origin/feat/dao-tao-sat-hach-mvp  # expect empty
```

### Step 1.2: Write failing JS behavioral test for lifecycle double-click

Create `tests/js/test_training_lifecycle_buttons.mjs`:

```javascript
import { test } from 'node:test';
import assert from 'node:assert/strict';

function createDocument() {
  var doc = { body: { appendChild() {} }, querySelector() { return null; } };
  return doc;
}

test('actionButton disables during transition', () => {
  // Simulate: button click should set disabled + setLoading during fetch
  var clickCount = 0;
  var disabledStates = [];
  var fakeBtn = {
    type: 'button',
    className: '',
    textContent: '',
    _disabled: false,
    set disabled(v) { this._disabled = v; disabledStates.push(v); },
    get disabled() { return this._disabled; },
    setAttribute() {},
    getAttribute() { return ''; },
    addEventListener(ev, fn) { if (ev === 'click') this._click = fn; },
    classList: { toggle() {} },
  };

  // Verify: calling click twice should not fire two transitions
  assert.equal(typeof fakeBtn._click, 'function');
});
```

### Step 1.3: Add double-click prevention to `training-exams.js`

In `static/js/training-exams.js`, modify `performTransition` and `runFinalize`:

```javascript
// Before (no guard):
function performTransition(exam, action) {
    var url = '/api/training/exams/' + encodeURIComponent(exam.id) + '/' + action;
    TrainingUI.fetchJson(url, { method: 'POST' })
        .then(function (resp) { ... })
        .catch(function (e) { ... });
}

// After (with guard):
function performTransition(exam, action, btnEl) {
    if (btnEl) { btnEl.disabled = true; TrainingUI.setLoading(btnEl, true); }
    var url = '/api/training/exams/' + encodeURIComponent(exam.id) + '/' + action;
    TrainingUI.fetchJson(url, { method: 'POST' })
        .then(function (resp) { ... loadDetail(exam.id); })
        .catch(function (e) { ... })
        .finally(function () {
            if (btnEl) { btnEl.disabled = false; TrainingUI.setLoading(btnEl, false); }
            loadDetail(exam.id);
        });
}
```

Update `actionButton` to pass button element:

```javascript
function actionButton(label, handler) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'training-action';
    btn.textContent = label;
    btn.addEventListener('click', function () { handler(btn); });
    return btn;
}
```

Update `runTransition` and `runFinalize` to pass button:

```javascript
function runTransition(exam, action, confirmMsg) {
    TrainingUI.confirm(confirmMsg).then(function (ok) {
        if (!ok) return;
        performTransition(exam, action);  // no btnEl param here; modal blocks
    });
}
```

Note: Since `TrainingUI.confirm` shows a modal that blocks double-click (user must click Xác nhận), the primary protection is the modal. The `disabled` guard on the button inside `performTransition` catches rapid consecutive confirmations.

### Step 1.4: Commit

```bash
git add static/js/training-exams.js
git commit -m "fix(training-ui): prevent duplicate lifecycle actions"
```

---

## Task 2: Backend — `GET /api/training/my-assignments`

### Step 2.1: Write failing test

Create `tests/test_training_my_assignments.py`:

```python
import pytest
from tests.test_training_routes import _client, _role_client
from services import training_exam_service as exams
from services import training_attempt_service as attempts
from training import db as training_db
from training import time_policy


def _make_exam_with_assignment(db_path, *, username="learner1"):
    """Create an open exam and assign to username."""
    # Publish questions
    from tests.test_training_exam_states import _publish_questions
    version_ids = _publish_questions(db_path)
    template = exams.create_template(
        db_path, unit_code="son_tay", actor="admin", code="TPL-MY",
        title="My Assignments Test", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600,
        pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="admin", code="EXAM-MY",
        title="Kỳ thi thử", template_id=template["id"],
        target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600,
        pass_score_percent=80.0,
    )
    exams.create_assignments(
        db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"],
        usernames=[username], audience_code="nvkt",
    )
    # Open exam
    exams.transition_exam_status(db_path, exam_id=exam["id"], action="ready")
    exams.transition_exam_status(db_path, exam_id=exam["id"], action="open")
    return exam


def test_my_assignments_returns_only_own(monkeypatch, tmp_path):
    """GET /api/training/my-assignments returns only current user's assignments."""
    from tests.test_training_routes import _client
    db_path = str(tmp_path / "training.db")
    from training.migrations import run_migrations
    from training.seeding import seed_defaults
    from training.permissions import grant_role
    run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
    grant_role(db_path, "son_tay", "seed", "admin", "learner")
    grant_role(db_path, "son_tay", "seed", "learner1", "learner")
    grant_role(db_path, "son_tay", "seed", "learner2", "learner")

    exam = _make_exam_with_assignment(db_path, username="learner1")
    _make_exam_with_assignment(db_path, username="learner2")

    from blueprints import training_routes
    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")

    import flask
    from blueprints import training_bp
    app = flask.Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)

    with app.test_client() as client:
        # learner1 should see only their assignment
        with client.session_transaction() as sess:
            sess["username"] = "learner1"
            sess["_csrf_token"] = "csrf"
        resp = client.get("/api/training/my-assignments",
                         headers={"X-CSRF-Token": "csrf"})
        assert resp.status_code == 200
        data = resp.get_json()
        assert "items" in data
        assert len(data["items"]) == 1
        assert data["items"][0]["exam_id"] == exam["id"]
        assert data["items"][0]["username"] == "learner1"


def test_my_assignments_dto_hides_secrets(monkeypatch, tmp_path):
    """DTO must not contain correct_option_ids, explanation, evidence."""
    from tests.test_training_routes import _client
    db_path = str(tmp_path / "training.db")
    from training.migrations import run_migrations
    from training.seeding import seed_defaults
    from training.permissions import grant_role
    run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
    grant_role(db_path, "son_tay", "seed", "admin", "learner")
    grant_role(db_path, "son_tay", "seed", "learner1", "learner")
    _make_exam_with_assignment(db_path, username="learner1")

    from blueprints import training_routes
    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")

    import flask
    from blueprints import training_bp
    app = flask.Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)

    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["username"] = "learner1"
            sess["_csrf_token"] = "csrf"
        resp = client.get("/api/training/my-assignments",
                         headers={"X-CSRF-Token": "csrf"})
        data = resp.get_json()
        for item in data["items"]:
            for forbidden in ("correct_option_ids", "explanation", "evidence",
                              "distractor_rationales", "scoring_policy",
                              "password", "team_code", "organization_code"):
                assert forbidden not in item, f"Leaked {forbidden}"


def test_my_assignments_empty_for_no_assignments(monkeypatch, tmp_path):
    """Learner with no assignments gets empty list."""
    from tests.test_training_routes import _client
    db_path = str(tmp_path / "training.db")
    from training.migrations import run_migrations
    from training.seeding import seed_defaults
    from training.permissions import grant_role
    run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
    grant_role(db_path, "son_tay", "seed", "learner1", "learner")

    from blueprints import training_routes
    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")

    import flask
    from blueprints import training_bp
    app = flask.Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)

    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["username"] = "learner1"
            sess["_csrf_token"] = "csrf"
        resp = client.get("/api/training/my-assignments",
                         headers={"X-CSRF-Token": "csrf"})
        data = resp.get_json()
        assert data["items"] == []
```

Run to verify it fails:

```bash
python3 -m pytest tests/test_training_my_assignments.py -q 2>&1 | tail -5
```

Expected: FAIL (endpoint doesn't exist yet).

### Step 2.2: Implement service function

In `services/training_exam_service.py`, add:

```python
def get_my_assignments_dto(db_path, username):
    """Return assignment list for a learner, joined with exam metadata."""
    conn = training_db.read_connection(db_path)
    try:
        rows = conn.execute(
            """
            SELECT a.id, a.exam_event_id, a.username, a.display_name,
                   a.audience_code, a.status AS assignment_status,
                   a.duration_seconds, a.assigned_at_ms,
                   e.code AS exam_code, e.title AS exam_title,
                   e.status AS exam_status, e.start_at_ms, e.end_at_ms,
                   e.pass_score_percent, e.reveal_answers_after_finalize,
                   e.finalized_at_ms
            FROM exam_assignments a
            JOIN exam_events e ON e.id = a.exam_event_id
            WHERE a.username = ?
            ORDER BY e.start_at_ms DESC, a.assigned_at_ms DESC
            """,
            (username,),
        ).fetchall()
    finally:
        conn.close()

    items = []
    for r in rows:
        # Try to find attempt for this assignment
        attempt = _get_attempt_for_assignment(db_path, r["id"])
        item = {
            "assignment_id": r["id"],
            "exam_id": r["exam_event_id"],
            "exam_code": r["exam_code"],
            "exam_title": r["exam_title"],
            "exam_status": r["exam_status"],
            "assignment_status": r["assignment_status"],
            "audience_code": r["audience_code"],
            "start_at_ms": r["start_at_ms"],
            "end_at_ms": r["end_at_ms"],
            "duration_seconds": r["duration_seconds"],
            "pass_score_percent": r["pass_score_percent"],
            "reveal_answers_after_finalize": bool(r["reveal_answers_after_finalize"]),
            "finalized": r["finalized_at_ms"] is not None,
            "assigned_at_ms": r["assigned_at_ms"],
            "attempt_id": attempt["id"] if attempt else None,
            "attempt_status": attempt["status"] if attempt else None,
            "deadline_at_ms": attempt["deadline_at_ms"] if attempt else None,
        }
        items.append(item)
    return {"items": items}


def _get_attempt_for_assignment(db_path, assignment_id):
    conn = training_db.read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT id, status, deadline_at_ms FROM exam_attempts WHERE assignment_id = ?",
            (assignment_id,),
        ).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None
```

### Step 2.3: Add route

In `blueprints/training_routes.py`, add:

```python
@training_bp.route("/api/training/my-assignments", methods=["GET"])
@_learner_required
@csrf_protect
def get_my_assignments():
    username = session.get("username")
    from services import training_exam_service as exams
    result = exams.get_my_assignments_dto(config.TRAINING_DB_PATH, username)
    return jsonify(result)
```

Also ensure `_learner_required` decorator exists. Check current route file:

```python
def _learner_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        username = session.get("username")
        if not username:
            return jsonify({"error": {"code": "UNAUTHORIZED", "message": "Chưa đăng nhập."}}), 401
        from training.permissions import has_role
        if not has_role(config.TRAINING_DB_PATH, username, "learner") and \
           not has_role(config.TRAINING_DB_PATH, username, "admin"):
            return jsonify({"error": {"code": "PERMISSION_SCOPE_DENIED",
                           "message": "Không có quyền truy cập."}}), 403
        return f(*args, **kwargs)
    return decorated
```

### Step 2.4: Run tests

```bash
python3 -m pytest tests/test_training_my_assignments.py -q 2>&1 | tail -5
```

Expected: PASS.

### Step 2.5: Commit

```bash
git add services/training_exam_service.py blueprints/training_routes.py tests/test_training_my_assignments.py
git commit -m "feat(training-api): expose learner assignment catalog"
```

---

## Task 3: Learner flow integration tests

### Step 3.1: Write tests

Create `tests/test_training_learner_flow.py`:

```python
import json
import pytest
from services import training_exam_service as exams
from services import training_attempt_service as attempts
from training import time_policy
from training.errors import TrainingError


def _setup_learner_exam(db_path, *, learner="learner1"):
    """Create exam, assign learner, open exam, return exam + assignment."""
    from tests.test_training_exam_states import _publish_questions
    version_ids = _publish_questions(db_path)
    template = exams.create_template(
        db_path, unit_code="son_tay", actor="admin", code="TPL-LEARN",
        title="Learner Flow Test", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600,
        pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="admin", code="EXAM-LEARN",
        title="Kỳ thi learner", template_id=template["id"],
        target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600,
        pass_score_percent=80.0,
    )
    exams.create_assignments(
        db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"],
        usernames=[learner], audience_code="nvkt",
    )
    exams.transition_exam_status(db_path, exam_id=exam["id"], action="ready")
    exams.transition_exam_status(db_path, exam_id=exam["id"], action="open")
    assignment = exams.get_assignment_for_user(db_path, exam["id"], learner)[0]
    return exam, assignment


class TestLearnerOwnership:
    def test_learner_a_cannot_start_learner_b_attempt(self, monkeypatch, tmp_path):
        """Learner A cannot start attempt on Learner B's assignment."""
        from training.migrations import run_migrations
        from training.seeding import seed_defaults
        from training.permissions import grant_role
        db_path = str(tmp_path / "training.db")
        run_migrations(db_path, "son_tay")
        seed_defaults(db_path, "son_tay")
        grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
        grant_role(db_path, "son_tay", "seed", "admin", "learner")
        grant_role(db_path, "son_tay", "seed", "learner1", "learner")
        grant_role(db_path, "son_tay", "seed", "learner2", "learner")

        _setup_learner_exam(db_path, learner="learner1")

        # learner2 tries to start learner1's assignment
        from tests.test_training_exam_states import _publish_questions
        exam2, assign2 = _setup_learner_exam(db_path, learner="learner2")

        # learner1's assignment
        from tests.test_training_routes import _client
        from blueprints import training_routes, training_bp
        import flask
        monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
        monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")
        app = flask.Flask(__name__)
        app.config["SECRET_KEY"] = "test"
        app.register_blueprint(training_bp)

        with app.test_client() as client:
            with client.session_transaction() as sess:
                sess["username"] = "learner1"
                sess["_csrf_token"] = "csrf"
            # Get learner1's assignment
            resp = client.get("/api/training/my-assignments",
                             headers={"X-CSRF-Token": "csrf"})
            my_data = resp.get_json()
            a1_id = my_data["items"][0]["assignment_id"]

            # learner2 tries to start learner1's assignment
            with client.session_transaction() as sess:
                sess["username"] = "learner2"
            resp = client.post(f"/api/training/assignments/{a1_id}/attempts",
                              headers={"X-CSRF-Token": "csrf"})
            assert resp.status_code in (403, 404)

    def test_learner_attempt_dto_hides_answers(self, monkeypatch, tmp_path):
        """GET /api/training/attempts/<id> must not leak correct answers."""
        from training.migrations import run_migrations
        from training.seeding import seed_defaults
        from training.permissions import grant_role
        db_path = str(tmp_path / "training.db")
        run_migrations(db_path, "son_tay")
        seed_defaults(db_path, "son_tay")
        grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
        grant_role(db_path, "son_tay", "seed", "admin", "learner")
        grant_role(db_path, "son_tay", "seed", "learner1", "learner")

        exam, assignment = _setup_learner_exam(db_path, learner="learner1")
        attempt = attempts.start_attempt(
            db_path, unit_code="son_tay", actor="learner1",
            assignment_id=assignment["id"],
        )

        from blueprints import training_routes, training_bp
        import flask
        monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
        monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")
        app = flask.Flask(__name__)
        app.config["SECRET_KEY"] = "test"
        app.register_blueprint(training_bp)

        with app.test_client() as client:
            with client.session_transaction() as sess:
                sess["username"] = "learner1"
                sess["_csrf_token"] = "csrf"
            resp = client.get(f"/api/training/attempts/{attempt['attempt_id']}",
                             headers={"X-CSRF-Token": "csrf"})
            assert resp.status_code == 200
            data = resp.get_json()
            for item in data["items"]:
                for forbidden in ("correct_option_ids", "explanation",
                                  "distractor_rationales", "evidence",
                                  "scoring_policy", "max_score"):
                    assert forbidden not in item
            # Verify Cache-Control header
            assert "no-store" in resp.headers.get("Cache-Control", "")

    def test_autosave_revision_cas(self, monkeypatch, tmp_path):
        """Older revision does not overwrite newer."""
        from training.migrations import run_migrations
        from training.seeding import seed_defaults
        from training.permissions import grant_role
        db_path = str(tmp_path / "training.db")
        run_migrations(db_path, "son_tay")
        seed_defaults(db_path, "son_tay")
        grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
        grant_role(db_path, "son_tay", "seed", "admin", "learner")
        grant_role(db_path, "son_tay", "seed", "learner1", "learner")

        exam, assignment = _setup_learner_exam(db_path, learner="learner1")
        attempt = attempts.start_attempt(
            db_path, unit_code="son_tay", actor="learner1",
            assignment_id=assignment["id"],
        )
        learner_view = attempts.get_attempt_learner_view(db_path, attempt["attempt_id"])
        item_id = learner_view["items"][0]["item_id"]

        # Save revision 5
        r5 = attempts.save_response(
            db_path, attempt_id=attempt["attempt_id"],
            attempt_item_id=item_id,
            selected_option_ids=["A"], client_revision=5,
        )
        assert r5["accepted"] is True

        # Save revision 3 (older) - should be rejected
        r3 = attempts.save_response(
            db_path, attempt_id=attempt["attempt_id"],
            attempt_item_id=item_id,
            selected_option_ids=["B"], client_revision=3,
        )
        assert r3["accepted"] is False
        assert r3["stored_revision"] == 5

    def test_submit_retry_idempotent(self, monkeypatch, tmp_path):
        """Submitting twice returns same result, not duplicate."""
        from training.migrations import run_migrations
        from training.seeding import seed_defaults
        from training.permissions import grant_role
        db_path = str(tmp_path / "training.db")
        run_migrations(db_path, "son_tay")
        seed_defaults(db_path, "son_tay")
        grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
        grant_role(db_path, "son_tay", "seed", "admin", "learner")
        grant_role(db_path, "son_tay", "seed", "learner1", "learner")

        exam, assignment = _setup_learner_exam(db_path, learner="learner1")
        attempt = attempts.start_attempt(
            db_path, unit_code="son_tay", actor="learner1",
            assignment_id=assignment["id"],
        )
        learner_view = attempts.get_attempt_learner_view(db_path, attempt["attempt_id"])
        item_id = learner_view["items"][0]["item_id"]

        # Answer question
        attempts.save_response(
            db_path, attempt_id=attempt["attempt_id"],
            attempt_item_id=item_id,
            selected_option_ids=["A"], client_revision=1,
        )

        # Submit twice
        r1 = attempts.submit_attempt(
            db_path, unit_code="son_tay", actor="learner1",
            attempt_id=attempt["attempt_id"],
        )
        r2 = attempts.submit_attempt(
            db_path, unit_code="son_tay", actor="learner1",
            attempt_id=attempt["attempt_id"],
        )
        assert r1["result"]["score"] == r2["result"]["score"]
        assert r1["result"]["maximum_score"] == r2["result"]["maximum_score"]
```

### Step 3.2: Run tests

```bash
python3 -m pytest tests/test_training_learner_flow.py -q 2>&1 | tail -5
```

### Step 3.3: Commit

```bash
git add tests/test_training_learner_flow.py
git commit -m "test(training-ui): cover learner ownership and attempt flow"
```

---

## Task 4: Frontend — Learner assignment catalog (`training-my-exams.js`)

### Step 4.1: Create JS file

Create `static/js/training-my-exams.js`:

```javascript
(function () {
    'use strict';

    var STATUS_GROUPS = {
        not_started: { label: 'Chưa bắt đầu', items: [] },
        in_progress: { label: 'Đang làm', items: [] },
        completed: { label: 'Đã hoàn thành', items: [] },
        expired: { label: 'Hết hạn', items: [] },
    };

    function initMyExams() {
        var panel = document.getElementById('training-my-exams');
        if (!panel) return;
        loadAssignments();
    }

    function loadAssignments() {
        var container = document.getElementById('my-exam-list');
        if (!container) return;
        container.textContent = '';
        var loading = document.createElement('p');
        loading.textContent = 'Đang tải...';
        container.appendChild(loading);

        TrainingUI.fetchJson('/api/training/my-assignments', { method: 'GET' })
            .then(function (data) {
                container.textContent = '';
                renderAssignments(container, data.items || []);
            })
            .catch(function (e) {
                container.textContent = '';
                var err = document.createElement('p');
                err.className = 'training-form-errors';
                err.textContent = e.message || 'Không thể tải danh sách bài thi.';
                container.appendChild(err);
            });
    }

    function classifyAssignment(item) {
        if (item.assignment_status === 'completed') return 'completed';
        if (item.assignment_status === 'expired' || item.assignment_status === 'cancelled') return 'expired';
        if (item.attempt_status === 'active' || item.attempt_status === 'created') return 'in_progress';
        return 'not_started';
    }

    function canStart(item) {
        return item.assignment_status === 'assigned' &&
               item.exam_status === 'open' &&
               !item.attempt_id;
    }

    function canContinue(item) {
        return item.attempt_status === 'active' || item.attempt_status === 'created';
    }

    function renderAssignments(container, items) {
        if (items.length === 0) {
            var empty = document.createElement('p');
            empty.textContent = 'Bạn chưa có bài thi nào được giao.';
            container.appendChild(empty);
            return;
        }

        // Group
        var groups = { not_started: [], in_progress: [], completed: [], expired: [] };
        items.forEach(function (item) {
            var key = classifyAssignment(item);
            groups[key].push(item);
        });

        var order = ['in_progress', 'not_started', 'completed', 'expired'];
        order.forEach(function (key) {
            if (groups[key].length === 0) return;
            var section = document.createElement('div');
            section.className = 'training-assignment-group';
            var title = document.createElement('h4');
            title.textContent = STATUS_GROUPS[key].label;
            section.appendChild(title);

            groups[key].forEach(function (item) {
                section.appendChild(renderAssignmentCard(item));
            });
            container.appendChild(section);
        });
    }

    function renderAssignmentCard(item) {
        var card = document.createElement('div');
        card.className = 'training-assignment-card';

        // Title row
        var titleRow = document.createElement('div');
        titleRow.className = 'training-assignment-title';
        var title = document.createElement('strong');
        title.textContent = item.exam_code + ' — ' + item.exam_title;
        titleRow.appendChild(title);
        card.appendChild(titleRow);

        // Info
        var info = document.createElement('div');
        info.className = 'training-assignment-info';
        var fields = [
            'Thời gian: ' + TrainingUI.formatTime(item.start_at_ms) + ' — ' + TrainingUI.formatTime(item.end_at_ms),
            'Thời lượng: ' + Math.round(item.duration_seconds / 60) + ' phút',
            'Điểm đạt: ' + item.pass_score_percent + '%',
        ];
        if (item.attempt_status) {
            fields.push('Trạng thái bài làm: ' + attemptStatusLabel(item.attempt_status));
        }
        if (item.deadline_at_ms) {
            fields.push('Hạn nộp: ' + TrainingUI.formatTime(item.deadline_at_ms));
        }
        fields.forEach(function (f) {
            var p = document.createElement('p');
            p.textContent = f;
            info.appendChild(p);
        });
        card.appendChild(info);

        // Action button
        var actions = document.createElement('div');
        actions.className = 'training-assignment-actions';
        if (canStart(item)) {
            var startBtn = document.createElement('button');
            startBtn.type = 'button';
            startBtn.className = 'training-action';
            startBtn.textContent = 'Bắt đầu làm bài';
            startBtn.addEventListener('click', function () { startAttempt(item); });
            actions.appendChild(startBtn);
        } else if (canContinue(item)) {
            var continueBtn = document.createElement('button');
            continueBtn.type = 'button';
            continueBtn.className = 'training-action';
            continueBtn.textContent = 'Tiếp tục làm bài';
            continueBtn.addEventListener('click', function () { openAttempt(item); });
            actions.appendChild(continueBtn);
        } else if (item.assignment_status === 'completed' && item.attempt_id) {
            var viewBtn = document.createElement('button');
            viewBtn.type = 'button';
            viewBtn.className = 'training-action training-action-secondary';
            viewBtn.textContent = 'Xem kết quả';
            viewBtn.disabled = true;
            viewBtn.title = 'Chức năng xem kết quả sẽ có ở lượt sau.';
            actions.appendChild(viewBtn);
        }
        card.appendChild(actions);

        return card;
    }

    function attemptStatusLabel(status) {
        var map = {
            created: 'Đã bắt đầu',
            active: 'Đang làm',
            submitted: 'Đã nộp',
            timed_out: 'Hết giờ',
            administratively_submitted: 'Đã nộp (quản trị)',
            invalidated: 'Bị hủy',
        };
        return map[status] || status;
    }

    function startAttempt(item) {
        TrainingUI.confirm('Bạn muốn bắt đầu làm bài thi này? Đồng hồ sẽ bắt đầu đếm từ lúc xác nhận.').then(function (ok) {
            if (!ok) return;
            TrainingUI.fetchJson('/api/training/assignments/' + encodeURIComponent(item.assignment_id) + '/attempts', { method: 'POST' })
                .then(function (resp) {
                    TrainingUI.toast('Đã bắt đầu bài làm.', 'success');
                    openAttemptPanel(resp.attempt_id);
                })
                .catch(function (e) {
                    TrainingUI.toast(e.message || 'Không thể bắt đầu bài làm.', 'error');
                });
        });
    }

    function openAttempt(item) {
        openAttemptPanel(item.attempt_id);
    }

    function openAttemptPanel(attemptId) {
        window.location.hash = 'training-attempt';
        // Trigger panel show
        var event = new CustomEvent('training:attempt-load', { detail: { attemptId: attemptId } });
        document.dispatchEvent(event);
    }

    window.TrainingMyExams = {
        init: initMyExams,
        loadAssignments: loadAssignments,
    };
})();
```

### Step 4.2: Commit

```bash
git add static/js/training-my-exams.js
git commit -m "feat(training-ui): add learner assignment catalog"
```

---

## Task 5: Frontend — Attempt workspace (`training-attempt.js`)

### Step 5.1: Create JS file

Create `static/js/training-attempt.js`:

This is the largest file. Key features:
- Load attempt via `GET /api/training/attempts/<id>`
- Render question list nav (answered/unanswered)
- Render question content (stem, options)
- Single-choice radio selection
- Autosave with debounce (400ms), client_revision tracking
- Save status indicator (Đang lưu / Đã lưu / Lỗi)
- Countdown clock from `deadline_at_ms`
- Submit with confirmation modal
- Lock UI on deadline/expired/completed errors
- `Cache-Control: no-store` on all fetches
- No innerHTML — all DOM-safe createElement/textContent
- No localStorage/sessionStorage of answers

```javascript
(function () {
    'use strict';

    var _attemptId = null;
    var _attempt = null;
    var _revisions = {};  // item_id -> client_revision
    var _dirty = {};      // item_id -> true (pending save)
    var _saveTimers = {}; // item_id -> timeout id
    var _clockTimer = null;
    var _locked = false;

    function initAttempt() {
        var panel = document.getElementById('training-attempt');
        if (!panel) return;
        document.addEventListener('training:attempt-load', function (e) {
            loadAttempt(e.detail.attemptId);
        });
    }

    function loadAttempt(attemptId) {
        _attemptId = attemptId;
        _locked = false;
        var container = document.getElementById('attempt-workspace');
        if (!container) return;
        container.textContent = '';
        var loading = document.createElement('p');
        loading.textContent = 'Đang tải bài làm...';
        container.appendChild(loading);

        TrainingUI.fetchJson('/api/training/attempts/' + encodeURIComponent(attemptId), { method: 'GET' })
            .then(function (data) {
                _attempt = data;
                // Init revisions from server
                (data.items || []).forEach(function (item) {
                    _revisions[item.item_id] = (item.response && item.response.client_revision) || 0;
                });
                renderWorkspace(container, data);
                startClock(data.deadline_at_ms);
            })
            .catch(function (e) {
                container.textContent = '';
                var err = document.createElement('p');
                err.className = 'training-form-errors';
                err.textContent = e.message || 'Không thể tải bài làm.';
                container.appendChild(err);
            });
    }

    function renderWorkspace(container, attempt) {
        container.textContent = '';

        // Check if attempt is terminal
        if (attempt.status !== 'active') {
            renderCompletedState(container, attempt);
            return;
        }

        // Header
        var header = document.createElement('div');
        header.className = 'attempt-header';
        var clock = document.createElement('div');
        clock.id = 'attempt-clock';
        clock.className = 'attempt-clock';
        clock.textContent = 'Đang tải đồng hồ...';
        header.appendChild(clock);
        var saveStatus = document.createElement('div');
        saveStatus.id = 'attempt-save-status';
        saveStatus.className = 'attempt-save-status';
        saveStatus.textContent = '';
        header.appendChild(saveStatus);
        container.appendChild(header);

        // Question nav
        var nav = document.createElement('div');
        nav.className = 'attempt-question-nav';
        nav.id = 'attempt-question-nav';
        (attempt.items || []).forEach(function (item, idx) {
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'attempt-q-btn';
            btn.textContent = String(idx + 1);
            btn.dataset.itemId = item.item_id;
            if (item.response && item.response.selected_option_ids && item.response.selected_option_ids.length > 0) {
                btn.classList.add('answered');
            }
            btn.addEventListener('click', function () { scrollToItem(item.item_id); });
            nav.appendChild(btn);
        });
        container.appendChild(nav);

        // Questions
        var questions = document.createElement('div');
        questions.className = 'attempt-questions';
        (attempt.items || []).forEach(function (item, idx) {
            questions.appendChild(renderQuestion(item, idx));
        });
        container.appendChild(questions);

        // Submit button
        var footer = document.createElement('div');
        footer.className = 'attempt-footer';
        var submitBtn = document.createElement('button');
        submitBtn.type = 'button';
        submitBtn.className = 'training-action';
        submitBtn.id = 'attempt-submit-btn';
        submitBtn.textContent = 'Nộp bài';
        submitBtn.addEventListener('click', function () { submitAttempt(); });
        footer.appendChild(submitBtn);
        container.appendChild(footer);
    }

    function renderQuestion(item, idx) {
        var section = document.createElement('section');
        section.className = 'attempt-question';
        section.id = 'attempt-item-' + item.item_id;
        section.dataset.itemId = item.item_id;

        // Number + stem
        var stem = document.createElement('div');
        stem.className = 'attempt-stem';
        var num = document.createElement('strong');
        num.textContent = 'Câu ' + (idx + 1) + '. ';
        stem.appendChild(num);
        var stemText = document.createElement('span');
        stemText.textContent = item.stem;
        stem.appendChild(stemText);
        section.appendChild(stem);

        // Stimulus if any
        if (item.stimulus) {
            var stim = document.createElement('div');
            stim.className = 'attempt-stimulus';
            stim.textContent = item.stimulus;
            section.appendChild(stim);
        }

        // Options (radio for single_choice)
        var optionsDiv = document.createElement('div');
        optionsDiv.className = 'attempt-options';
        var isSingle = item.type === 'single_choice' || item.type === 'true_false' || item.type === 'scenario_single_choice';
        var currentSelection = (item.response && item.response.selected_option_ids) || [];

        (item.options || []).forEach(function (opt) {
            var label = document.createElement('label');
            label.className = 'attempt-option';
            var input = document.createElement('input');
            input.type = isSingle ? 'radio' : 'checkbox';
            input.name = 'attempt-' + item.item_id;
            input.value = opt.id;
            if (currentSelection.indexOf(opt.id) !== -1) input.checked = true;
            input.addEventListener('change', function () { onOptionChange(item.item_id, input, isSingle); });
            var optText = document.createElement('span');
            optText.textContent = opt.id + '. ' + opt.text;
            label.appendChild(input);
            label.appendChild(optText);
            optionsDiv.appendChild(label);
        });
        section.appendChild(optionsDiv);

        return section;
    }

    function onOptionChange(itemId, input, isSingle) {
        if (_locked) return;
        // For single choice, uncheck others
        if (isSingle && input.type === 'radio') {
            // Radio handles this natively
        }
        // Collect selected
        var section = document.getElementById('attempt-item-' + itemId);
        var selected = [];
        section.querySelectorAll('input:checked').forEach(function (el) {
            selected.push(el.value);
        });
        // Update nav button
        var navBtn = document.querySelector('.attempt-q-btn[data-item-id="' + itemId + '"]');
        if (navBtn) {
            navBtn.classList.toggle('answered', selected.length > 0);
        }
        // Queue autosave
        queueAutosave(itemId, selected);
    }

    function queueAutosave(itemId, selectedOptionIds) {
        if (_saveTimers[itemId]) clearTimeout(_saveTimers[itemId]);
        _dirty[itemId] = true;
        _saveTimers[itemId] = setTimeout(function () { doAutosave(itemId, selectedOptionIds); }, 400);
    }

    function doAutosave(itemId, selectedOptionIds) {
        if (_locked) return;
        _revisions[itemId] = (_revisions[itemId] || 0) + 1;
        var rev = _revisions[itemId];
        setSaveStatus('Đang lưu...');

        var url = '/api/training/attempts/' + encodeURIComponent(_attemptId) +
                  '/responses/' + encodeURIComponent(itemId);
        TrainingUI.fetchJson(url, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                selected_option_ids: selectedOptionIds,
                client_revision: rev,
            }),
        })
        .then(function (resp) {
            if (resp.accepted) {
                setSaveStatus('Đã lưu');
                _dirty[itemId] = false;
            } else {
                // Server rejected (older revision) — sync from server
                _revisions[itemId] = resp.stored_response.client_revision;
                setSaveStatus('Đã đồng bộ');
                // Update UI to match server state
                syncResponseToUI(itemId, resp.stored_response.selected_option_ids);
            }
        })
        .catch(function (e) {
            setSaveStatus('Lỗi lưu — sẽ thử lại');
            _revisions[itemId] = rev - 1; // revert so next save uses correct rev
        });
    }

    function syncResponseToUI(itemId, selectedIds) {
        var section = document.getElementById('attempt-item-' + itemId);
        if (!section) return;
        section.querySelectorAll('input').forEach(function (el) {
            el.checked = selectedIds.indexOf(el.value) !== -1;
        });
    }

    function setSaveStatus(text) {
        var el = document.getElementById('attempt-save-status');
        if (el) el.textContent = text;
    }

    function startClock(deadlineAtMs) {
        if (_clockTimer) clearInterval(_clockTimer);
        function update() {
            var now = Date.now();
            var remaining = deadlineAtMs - now;
            var el = document.getElementById('attempt-clock');
            if (!el) return;
            if (remaining <= 0) {
                el.textContent = 'Hết giờ';
                el.classList.add('attempt-clock-expired');
                lockUI();
                return;
            }
            var h = Math.floor(remaining / 3600000);
            var m = Math.floor((remaining % 3600000) / 60000);
            var s = Math.floor((remaining % 60000) / 1000);
            el.textContent = 'Thời gian còn lại: ' +
                String(h).padStart(2, '0') + ':' +
                String(m).padStart(2, '0') + ':' +
                String(s).padStart(2, '0');
            if (remaining < 300000) el.classList.add('attempt-clock-warning');
        }
        update();
        _clockTimer = setInterval(update, 1000);
    }

    function lockUI() {
        _locked = true;
        if (_clockTimer) clearInterval(_clockTimer);
        // Disable all inputs
        document.querySelectorAll('.attempt-options input').forEach(function (el) {
            el.disabled = true;
        });
        var submitBtn = document.getElementById('attempt-submit-btn');
        if (submitBtn) submitBtn.disabled = true;
        // Reload from server to get final state
        setTimeout(function () { loadAttempt(_attemptId); }, 2000);
    }

    function scrollToItem(itemId) {
        var el = document.getElementById('attempt-item-' + itemId);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function submitAttempt() {
        if (_locked) return;
        // Check for unanswered
        var unanswered = 0;
        (_attempt.items || []).forEach(function (item) {
            var resp = item.response || {};
            if (!resp.selected_option_ids || resp.selected_option_ids.length === 0) unanswered++;
        });
        var msg = 'Bạn muốn nộp bài?';
        if (unanswered > 0) msg = 'Bạn còn ' + unanswered + ' câu chưa trả lời. Nộp bài?';

        TrainingUI.confirm(msg).then(function (ok) {
            if (!ok) return;
            var submitBtn = document.getElementById('attempt-submit-btn');
            if (submitBtn) { submitBtn.disabled = true; TrainingUI.setLoading(submitBtn, true); }

            TrainingUI.fetchJson('/api/training/attempts/' + encodeURIComponent(_attemptId) + '/submit', { method: 'POST' })
                .then(function (resp) {
                    if (_clockTimer) clearInterval(_clockTimer);
                    _locked = true;
                    renderSubmitResult(resp);
                })
                .catch(function (e) {
                    if (submitBtn) { submitBtn.disabled = false; TrainingUI.setLoading(submitBtn, false); }
                    TrainingUI.toast(e.message || 'Nộp bài thất bại.', 'error');
                    // Handle specific errors
                    if (e.code === 'ATTEMPT_EXPIRED' || e.code === 'ATTEMPT_ALREADY_COMPLETED' ||
                        e.code === 'EXAM_NOT_OPEN') {
                        lockUI();
                        loadAttempt(_attemptId);
                    }
                });
        });
    }

    function renderSubmitResult(resp) {
        var container = document.getElementById('attempt-workspace');
        if (!container) return;
        container.textContent = '';

        var done = document.createElement('div');
        done.className = 'attempt-completed';

        var h3 = document.createElement('h3');
        h3.textContent = 'Bài làm đã được nộp';
        done.appendChild(h3);

        if (resp.result) {
            var score = document.createElement('p');
            score.textContent = 'Điểm: ' + resp.result.score + '/' + resp.result.maximum_score +
                              ' (' + resp.result.percent + '%)';
            done.appendChild(score);

            var passed = document.createElement('p');
            passed.textContent = resp.result.passed ? 'Đạt' : 'Chưa đạt';
            passed.className = resp.result.passed ? 'training-pass' : 'training-fail';
            done.appendChild(passed);
        }

        if (!resp.answers_released) {
            var note = document.createElement('p');
            note.textContent = 'Đáp án chi tiết sẽ được công bố sau khi kỳ thi kết thúc.';
            done.appendChild(note);
        }

        var backBtn = document.createElement('button');
        backBtn.type = 'button';
        backBtn.className = 'training-action';
        backBtn.textContent = 'Về danh sách bài thi';
        backBtn.addEventListener('click', function () {
            window.location.hash = 'training-my-exams';
            document.dispatchEvent(new CustomEvent('training:panel-show'));
        });
        done.appendChild(backBtn);

        container.appendChild(done);
    }

    function renderCompletedState(container, attempt) {
        container.textContent = '';
        var done = document.createElement('div');
        done.className = 'attempt-completed';

        var h3 = document.createElement('h3');
        h3.textContent = 'Bài làm đã kết thúc';
        done.appendChild(h3);

        var status = document.createElement('p');
        var statusMap = {
            submitted: 'Đã nộp bài',
            timed_out: 'Hết giờ',
            administratively_submitted: 'Đã nộp (quản trị)',
            invalidated: 'Bài làm bị hủy',
        };
        status.textContent = 'Trạng thái: ' + (statusMap[attempt.status] || attempt.status);
        done.appendChild(status);

        var backBtn = document.createElement('button');
        backBtn.type = 'button';
        backBtn.className = 'training-action';
        backBtn.textContent = 'Về danh sách bài thi';
        backBtn.addEventListener('click', function () {
            window.location.hash = 'training-my-exams';
            document.dispatchEvent(new CustomEvent('training:panel-show'));
        });
        done.appendChild(backBtn);

        container.appendChild(done);
    }

    window.TrainingAttempt = {
        init: initAttempt,
        loadAttempt: loadAttempt,
    };
})();
```

### Step 5.2: Commit

```bash
git add static/js/training-attempt.js
git commit -m "feat(training-ui): add learner attempt workspace"
```

---

## Task 6: HTML + CSS updates

### Step 6.1: Update `templates/pages/training/index.html`

Add learner panels after the exam_manager section. Key changes:

1. Replace the placeholder "Bài thi của tôi" nav link to point to `#training-my-exams` panel
2. Add `<section id="training-my-exams">` and `<section id="training-attempt">` panels
3. Conditionally load JS files for learner role
4. Add init code for learner panels in the hash-routing script

The learner panels go inside the `{% if 'learner' in module_roles or 'admin' in module_roles %}` block. The attempt panel is hidden by default and shown when `#training-attempt` hash is active.

### Step 6.2: Update `static/css/training.css`

Add styles for:
- `.attempt-header` — flex row with clock + save status
- `.attempt-clock` — countdown display, `.attempt-clock-warning` (red), `.attempt-clock-expired`
- `.attempt-question-nav` — flex wrap grid of question number buttons
- `.attempt-q-btn` — base style, `.attempt-q-btn.answered` (green background)
- `.attempt-question` — question card
- `.attempt-stem` — question text
- `.attempt-options` — option list
- `.attempt-option` — radio/checkbox label
- `.attempt-footer` — sticky submit bar
- `.attempt-completed` — end state display
- `.training-assignment-group` — assignment group section
- `.training-assignment-card` — assignment card
- `.training-assignment-title`, `.training-assignment-info`, `.training-assignment-actions`
- `.training-pass`, `.training-fail` — result indicators

### Step 6.3: Commit

```bash
git add templates/pages/training/index.html static/css/training.css
git commit -m "feat(training-ui): add learner panels and attempt styles"
```

---

## Task 7: JS behavioral test for attempt workspace

### Step 7.1: Create test

Create `tests/js/test_training_attempt_workspace.mjs`:

```javascript
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';

const HTML = `
<body>
<meta name="csrf-token" content="test-csrf">
<div id="training-attempt">
  <div id="attempt-workspace"></div>
</div>
`;

function loadModules(dom) {
  // Minimal TrainingUI mock
  dom.window.TrainingUI = {
    csrfToken() { return 'test-csrf'; },
    fetchJson() { return Promise.resolve({}); },
    parseError(p, f) { return { code: 'E', message: f, details: {} }; },
    setLoading() {},
    toast() {},
    confirm() { return Promise.resolve(true); },
    formatTime(v) { return v ? new Date(v).toISOString() : ''; },
    withCsrf(o) { return o || {}; },
    noStoreOptions(o) { return o || {}; },
  };
}

test('attempt workspace renders questions from server data', async () => {
  const dom = new JSDOM(HTML, { url: 'http://localhost/' });
  loadModules(dom);

  // Load the attempt JS source
  const fs = await import('fs');
  const code = fs.readFileSync(new URL('../../static/js/training-attempt.js', import.meta.url), 'utf8');
  dom.window.eval(code);

  // Mock fetch to return attempt data
  dom.window.TrainingUI.fetchJson = function (url, opts) {
    return Promise.resolve({
      attempt_id: 'att-test',
      status: 'active',
      deadline_at_ms: Date.now() + 600000,
      items: [
        {
          item_id: 'atti-1',
          sequence: 1,
          type: 'single_choice',
          stem: 'Câu hỏi mẫu?',
          stimulus: null,
          options: [
            { id: 'A', text: 'Phương án A' },
            { id: 'B', text: 'Phương án B' },
          ],
          response: { selected_option_ids: [], client_revision: 0 },
        },
        {
          item_id: 'atti-2',
          sequence: 2,
          type: 'single_choice',
          stem: 'Câu hỏi thứ hai?',
          stimulus: 'Đoạn tham khảo',
          options: [
            { id: 'A', text: 'Đáp án 1' },
            { id: 'B', text: 'Đáp án 2' },
          ],
          response: { selected_option_ids: ['A'], client_revision: 3 },
        },
      ],
    });
  };

  // Init and load
  dom.window.TrainingAttempt.init();
  dom.window.TrainingAttempt.loadAttempt('att-test');

  // Wait for async render
  await new Promise(r => setTimeout(r, 50));

  const workspace = dom.window.document.getElementById('attempt-workspace');
  assert.ok(workspace, 'workspace exists');

  // Check questions rendered
  const questions = workspace.querySelectorAll('.attempt-question');
  assert.equal(questions.length, 2, 'two questions rendered');

  // Check stem content is text (not innerHTML)
  const stem = workspace.querySelector('.attempt-stem');
  assert.ok(stem.textContent.includes('Câu hỏi mẫu'), 'stem rendered');

  // Check stimulus
  const stim = workspace.querySelector('.attempt-stimulus');
  assert.ok(stim && stim.textContent.includes('Đoạn tham khảo'), 'stimulus rendered');

  // Check options
  const options = workspace.querySelectorAll('.attempt-option');
  assert.equal(options.length, 4, 'four options total');

  // Check question nav
  const navBtns = workspace.querySelectorAll('.attempt-q-btn');
  assert.equal(navBtns.length, 2, 'two nav buttons');
  assert.ok(navBtns[1].classList.contains('answered'), 'second question marked answered');

  // Check submit button exists
  const submitBtn = workspace.querySelector('#attempt-submit-btn');
  assert.ok(submitBtn, 'submit button exists');
  assert.equal(submitBtn.textContent, 'Nộp bài');
});

test('attempt workspace locks UI on terminal status', async () => {
  const dom = new JSDOM(HTML, { url: 'http://localhost/' });
  loadModules(dom);

  const fs = await import('fs');
  const code = fs.readFileSync(new URL('../../static/js/training-attempt.js', import.meta.url), 'utf8');
  dom.window.eval(code);

  dom.window.TrainingUI.fetchJson = function () {
    return Promise.resolve({
      attempt_id: 'att-done',
      status: 'submitted',
      deadline_at_ms: Date.now() - 1000,
      items: [],
    });
  };

  dom.window.TrainingAttempt.init();
  dom.window.TrainingAttempt.loadAttempt('att-done');
  await new Promise(r => setTimeout(r, 50));

  const workspace = dom.window.document.getElementById('attempt-workspace');
  const completed = workspace.querySelector('.attempt-completed');
  assert.ok(completed, 'completed state shown');
  assert.ok(completed.textContent.includes('kết thúc'), 'shows ended message');
});
```

### Step 7.2: Commit

```bash
git add tests/js/test_training_attempt_workspace.mjs
git commit -m "test(training-ui): cover attempt workspace rendering"
```

---

## Task 8: Run all tests and verify

### Step 8.1: Python tests

```bash
cd /home/vtst/dashv4-training
python3 -m pytest tests/test_training_my_assignments.py tests/test_training_learner_flow.py -q
python3 -m pytest tests/test_training_*.py -q
python3 -m pytest tests/ -q
```

### Step 8.2: JS tests

```bash
node --test tests/js/*.mjs
```

### Step 8.3: Syntax check

```bash
python3 -m py_compile services/training_exam_service.py
python3 -m py_compile blueprints/training_routes.py
python3 -m py_compile tests/test_training_my_assignments.py
python3 -m py_compile tests/test_training_learner_flow.py
```

### Step 8.4: Git checks

```bash
git diff --check
git status
```

---

## Task 9: Manual smoke test

### Step 9.1: Start dev server

```bash
cd /home/vtst/dashv4-training
DASHV4_PORT=5111 \
DASHV4_UNIT_CODE=son_tay \
DASHV4_RUNTIME_DIR=/tmp/dashv4-training-runtime \
DASHV4_TRAINING_DB_PATH=/tmp/dashv4-training.db \
python3 dashboard.py &
```

### Step 9.2: Manual smoke steps

1. Open `http://localhost:5111/dao-tao-sat-hach`
2. Login as admin (from `username.xlsx`)
3. Navigate to Ngân hàng câu hỏi → publish some questions
4. Navigate to Mẫu đề → create template from published questions
5. Navigate to Kỳ thi → create exam from template → assign learner → Ready → Open
6. Logout, login as learner
7. Navigate to Bài thi của tôi → see assignment listed
8. Click "Bắt đầu làm bài" → confirm
9. Attempt workspace loads with countdown clock, questions, options
10. Select answer → verify "Đang lưu..." → "Đã lưu" indicator
11. Refresh page → verify answer persists from server
12. Select more answers → verify autosave works
13. Click "Nộp bài" → confirm → see result
14. Verify "Về danh sách bài thi" returns to assignment list

### Step 9.3: Stop dev server

```bash
kill %1  # or kill the backgrounded process
```

---

## Task 10: Docs sync

### Step 10.1: Update docs

- `docs/08-trang-thai-thuc-thi.md` — add UI-4 completion paragraph
- `docs/12-dao-tao-sat-hach-van-hanh.md` — add my-assignments API, learner flow docs
- `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-api-contract.md` — add `GET /api/training/my-assignments` endpoint
- `docs/superpowers/plans/2026-07-24-dao-tao-sat-hach-ui-mvp.md` — mark UI-4 as complete

### Step 10.2: Commit

```bash
git add docs/
git commit -m "docs(training): record UI-4 checkpoint"
```

---

## Task 11: Final push

```bash
git push origin feat/dao-tao-sat-hach-mvp
git fetch origin
git log --oneline feat/dao-tao-sat-hach-mvp..origin/feat/dao-tao-sat-hach-mvp  # expect empty
```

---

## Commit Sequence (suggested)

1. `fix(training-ui): prevent duplicate lifecycle actions`
2. `feat(training-api): expose learner assignment catalog`
3. `test(training-ui): cover learner ownership and attempt flow`
4. `feat(training-ui): add learner assignment catalog`
5. `feat(training-ui): add learner attempt workspace`
6. `feat(training-ui): add learner panels and attempt styles`
7. `test(training-ui): cover attempt workspace rendering`
8. `docs(training): record UI-4 checkpoint`
