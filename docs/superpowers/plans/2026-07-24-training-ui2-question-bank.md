# UI-2 Question Bank Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a secure, role-aware question-bank review workflow to the existing training workspace.

**Architecture:** `training_question_service` owns explicit list/detail management DTOs, classification derivation, and state validation. `training_routes` only applies RBAC, CSRF, request normalization, and service delegation. The Jinja workspace renders role-gated controls while a dedicated small JS module uses `TrainingUI` to call the APIs.

**Tech Stack:** Flask, Jinja, SQLite, pytest, vanilla JavaScript, existing `TrainingUI` helpers.

---

### Task 1: Management DTOs and filters

**Files:**
- Modify: `services/training_question_service.py`
- Test: `tests/test_training_questions.py`

- [ ] **Step 1: Write the failing service tests**

```python
def test_list_questions_filters_and_returns_management_list_dto(monkeypatch, tmp_path):
    db_path = _db(monkeypatch, tmp_path)
    imported = import_question_batch(db_path, unit_code="son_tay", actor="editor", batch=VALID_BATCH)
    result = list_questions(db_path, status="draft", audience="nvkt", domain="quality", topic="sample_topic", page=1, page_size=1)
    assert result["total"] == 1
    assert set(result["items"][0]) == {"id", "stem", "type", "difficulty", "audience", "topic", "status", "version"}
    assert result["items"][0]["id"] == imported["version_ids"][0]

def test_get_question_management_detail_composes_sensitive_fields(monkeypatch, tmp_path):
    db_path = _db(monkeypatch, tmp_path)
    version_id = import_question_batch(db_path, unit_code="son_tay", actor="editor", batch=VALID_BATCH)["version_ids"][0]
    detail = get_question_management_detail(db_path, version_id)
    assert detail["correct_option_ids"] == ["B"]
    assert detail["classification"]["audience_codes"] == ["nvkt"]
    assert detail["evidence"][0]["block_id"] == "DOC-B001"
```

- [ ] **Step 2: Run the tests and verify RED**

Run: `python3 -m pytest tests/test_training_questions.py -k 'management_list or management_detail' -q`

Expected: import/attribute failure because the DTO API does not exist.

- [ ] **Step 3: Implement the service DTOs and filters**

```python
def list_questions(db_path, *, status=None, audience=None, domain=None, topic=None, page=1, page_size=25, q=None):
    # Use EXISTS for every mapping so joins cannot duplicate question versions.
    ...

def get_question_management_detail(db_path, version_id):
    # Return an explicit DTO, never dict(row) from SELECT *.
    return {
        "id": version_id, "stem": row["stem"], "options": options,
        "correct_option_ids": json.loads(row["correct_option_ids_json"]),
        "classification": classification, "evidence": evidence,
        "review_history": reviews, "publication": publication,
    }
```

Map filter status as draft (`unpublished` and review `draft`/`needs_review`), approved (`unpublished` and review `approved`), and published (`published`). Derive domains with `question_sources` to `knowledge_blocks`; use `EXISTS` for audience/topic/domain. Include only the stipulated short list fields in list items.

- [ ] **Step 4: Run the tests and verify GREEN**

Run: `python3 -m pytest tests/test_training_questions.py -k 'management_list or management_detail' -q`

Expected: PASS.

### Task 2: Question-bank routes and review transition

**Files:**
- Modify: `blueprints/training_routes.py`
- Test: `tests/test_training_routes.py`

- [ ] **Step 1: Write failing route tests**

```python
def test_question_bank_reads_require_editor_or_manager(monkeypatch, tmp_path):
    for client, _ in _role_client(monkeypatch, tmp_path, username="learner", roles=("learner",)):
        responses = [client.get("/api/training/questions"), client.get("/api/training/questions/missing")]
    assert [response.status_code for response in responses] == [403, 403]

def test_question_bank_invalid_import_uses_stable_envelope(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        response = client.post("/api/training/questions/import", headers={"X-CSRF-Token": "csrf"}, json={})
    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "VALIDATION_ERROR"

def test_manager_can_reject_approve_and_publish_question(monkeypatch, tmp_path):
    ...
    assert reject.status_code == 200
    assert approve.status_code == 200
    assert publish.status_code == 200
```

- [ ] **Step 2: Run the tests and verify RED**

Run: `python3 -m pytest tests/test_training_routes.py -k 'question_bank or reject_approve' -q`

Expected: FAIL because GET/read and reject routes do not exist and import errors lack field details.

- [ ] **Step 3: Implement minimal protected routes**

```python
_question_reader_required = _module_role_required(("editor", "exam_manager"), "Không có quyền xem ngân hàng câu hỏi.")

@training_bp.route("/api/training/questions")
@_question_reader_required
def list_question_bank():
    return jsonify(questions.list_questions(config.TRAINING_DB_PATH, ...))

@training_bp.route("/api/training/questions/<version_id>/reject", methods=["POST"])
@csrf_protect
@_exam_manager_required
def reject_question(version_id):
    questions.add_review_action(..., action="reject", comment=(request.get_json(silent=True) or {}).get("comment"))
    return jsonify({"version_id": version_id, "review_status": "rejected"})
```

Extend `_module_role_required` to accept one role or a tuple. Parse page safely and cap page size at 100. Convert invalid JSON batch failures into `VALIDATION_ERROR` with `details={"errors": errors}` through the service error, while keeping existing envelope shape. Do not add a learner question endpoint.

- [ ] **Step 4: Run the tests and verify GREEN**

Run: `python3 -m pytest tests/test_training_routes.py -k 'question_bank or reject_approve' -q`

Expected: PASS.

### Task 3: Learner secrecy regression

**Files:**
- Test: `tests/test_training_routes.py`

- [ ] **Step 1: Write the failing regression test against the real route**

```python
def test_learner_attempt_route_hides_question_secrets_and_is_no_store(monkeypatch, tmp_path):
    for client, attempt_id, _, _ in _learner_client_with_attempt(monkeypatch, tmp_path):
        response = client.get(f"/api/training/attempts/{attempt_id}")
    serialized = response.get_data(as_text=True)
    assert response.headers["Cache-Control"] == "no-store"
    for key in ("correct_option_ids", "explanation", "evidence", "distractor_rationales", "scoring_policy", "max_score"):
        assert key not in serialized
```

- [ ] **Step 2: Run the test and verify its baseline**

Run: `python3 -m pytest tests/test_training_routes.py::test_learner_attempt_route_hides_question_secrets_and_is_no_store -q`

Expected: PASS once it proves the unchanged learner DTO contract. If it fails, fix only `get_attempt_learner_view` and `get_attempt` so the route returns no secret key and exactly `Cache-Control: no-store`.

- [ ] **Step 3: Run related route tests**

Run: `python3 -m pytest tests/test_training_routes.py -k 'attempt or learner' -q`

Expected: PASS.

### Task 4: Workspace panel and client behavior

**Files:**
- Modify: `templates/pages/training/index.html`
- Create: `static/js/training-question-bank.js`
- Modify: `static/css/training.css`

- [ ] **Step 1: Write failing template assertions**

```python
def test_training_workspace_exposes_question_bank_only_to_operator_roles(monkeypatch, tmp_path):
    ...
    assert 'id="training-question-bank"' in editor_page
    assert 'id="training-question-bank"' not in learner_page
```

- [ ] **Step 2: Run the test and verify RED**

Run: `python3 -m pytest tests/test_training_routes.py::test_training_workspace_exposes_question_bank_only_to_operator_roles -q`

Expected: FAIL because the panel is currently pending-only.

- [ ] **Step 3: Add role-gated accessible markup and dedicated client code**

```html
<section class="training-workspace-panel" id="training-question-bank" hidden>
  <form id="question-bank-filters">...</form>
  <div id="question-bank-errors" role="alert" aria-live="assertive"></div>
  <div id="question-bank-list" aria-live="polite"></div>
  <textarea id="question-batch-json" aria-label="JSON lô câu hỏi"></textarea>
</section>
```

```javascript
TrainingUI.fetchJson('/api/training/questions?' + params).then(renderList).catch(renderError);
// Validation/import never assigns textarea.value, preserving operator input.
```

Render management detail fields only from the protected detail endpoint. Render import validation error `details.errors` per question, disable manager-only actions unless list/detail status permits them, require `TrainingUI.confirm` before publishing, and bind Escape/cancel focus behavior through existing semantic controls. Add responsive grid/table styles without changing global styles.

- [ ] **Step 4: Run the template test and inspect client syntax**

Run: `python3 -m pytest tests/test_training_routes.py::test_training_workspace_exposes_question_bank_only_to_operator_roles -q`

Expected: PASS.

Run: `node --check static/js/training-question-bank.js`

Expected: exit code 0.

### Task 5: Documentation and final verification

**Files:**
- Modify: `docs/04-mapping-route-va-du-lieu.md`
- Modify: `docs/08-trang-thai-thuc-thi.md`
- Modify: `docs/12-dao-tao-sat-hach-van-hanh.md`

- [ ] **Step 1: Update only the UI-2 checkpoint**

Record protected management question list/detail/reject APIs, the UI-2 workspace completion, no migration, and preserved learner secrecy. Do not state that UI-3, learner UI completion, production AI, or retake flow is complete.

- [ ] **Step 2: Run full verification**

Run: `python3 -m pytest tests/`

Expected: all tests PASS.

Run: `python3 -m py_compile services/training_question_service.py blueprints/training_routes.py`

Expected: exit code 0.

Run: `git diff --check`

Expected: no output.

- [ ] **Step 3: Commit focused changes and push only the feature branch**

```bash
git add services/training_question_service.py blueprints/training_routes.py templates/pages/training/index.html static/js/training-question-bank.js static/css/training.css
git commit -m "feat(training-ui): add question review workflow"
git add tests/test_training_questions.py tests/test_training_routes.py docs/04-mapping-route-va-du-lieu.md docs/08-trang-thai-thuc-thi.md docs/12-dao-tao-sat-hach-van-hanh.md docs/superpowers/plans/2026-07-24-training-ui2-question-bank.md
git commit -m "test(training-ui): cover question bank workflow"
git push origin feat/dao-tao-sat-hach-mvp
```
