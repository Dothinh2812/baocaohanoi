import copy

import pytest

from services import training_attempt_service as attempts
from services import training_exam_service as exams
from tests.test_training_attempt_lifecycle import _make_open_exam_with_assignment
from tests.test_training_exam_states import _publish_questions, _setup
from tests.test_training_routes import QUESTION_BATCH, _client, _role_client
from training import db as training_db
from training import time_policy
from training.errors import TrainingError


def _make_template(db_path, *, code="TPL-LIST", shuffle_questions=True, shuffle_options=False):
    version_ids = _publish_questions(db_path)
    return exams.create_template(
        db_path, unit_code="son_tay", actor="alice", code=code, title="Template",
        target_audience_code="nvkt", question_version_ids=version_ids,
        duration_seconds=600, pass_score_percent=80.0,
        shuffle_questions=shuffle_questions, shuffle_options=shuffle_options,
    )


def _make_exam(db_path, *, code="EXAM-LIST"):
    template = _make_template(db_path, code=f"TPL-{code}")
    now = time_policy.utc_now_ms()
    return exams.create_exam(
        db_path, unit_code="son_tay", actor="alice", code=code, title="Kỳ thi",
        template_id=template["id"], target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
    )


_TEMPLATE_DTO_KEYS = {
    "id", "code", "title", "target_audience_code", "total_questions",
    "duration_seconds", "pass_score_percent", "shuffle_questions",
    "shuffle_options", "locked", "created_by", "created_at_ms",
}

_EXAM_DTO_KEYS = {
    "id", "code", "title", "template_id", "target_audience_code", "status",
    "start_at_ms", "end_at_ms", "duration_seconds", "pass_score_percent",
    "reveal_answers_after_finalize", "created_by", "created_at_ms",
    "finalized_at_ms",
}

_ASSIGNMENT_FORBIDDEN_KEYS = {
    "duration_seconds", "retake_of_assignment_id", "user_source_updated_at_ms",
}


# --- Service-level DTO allowlist checks ---

def test_list_templates_returns_explicit_dto_with_booleans(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _make_template(db_path, shuffle_questions=True, shuffle_options=False)
    result = exams.list_templates(db_path, page=1, page_size=25)

    assert result["page"] == 1
    assert result["page_size"] == 25
    assert result["total"] == 1
    assert len(result["items"]) == 1
    item = result["items"][0]
    assert set(item) == _TEMPLATE_DTO_KEYS
    assert item["shuffle_questions"] is True
    assert item["shuffle_options"] is False
    assert item["locked"] is False


def test_get_template_detail_includes_items_with_question_metadata(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    template = _make_template(db_path)
    detail = exams.get_template_detail(db_path, template["id"])

    assert set(detail) == _TEMPLATE_DTO_KEYS | {"items"}
    assert len(detail["items"]) == 2
    item = detail["items"][0]
    assert set(item) == {
        "sequence_number", "question_version_id", "stem", "type",
        "difficulty", "section_label", "points",
    }
    assert item["stem"]
    assert item["type"] == "single_choice"
    assert item["difficulty"] == "easy"
    assert [row["sequence_number"] for row in detail["items"]] == [1, 2]


def test_get_template_detail_raises_not_found(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    with pytest.raises(TrainingError) as exc:
        exams.get_template_detail(db_path, "missing")
    assert exc.value.code == "NOT_FOUND"
    assert exc.value.status == 404


def test_list_exams_returns_explicit_dto(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _make_exam(db_path, code="EXAM-DTO")
    result = exams.list_exams(db_path, page=1, page_size=25)

    assert result["total"] == 1
    item = result["items"][0]
    assert set(item) == _EXAM_DTO_KEYS
    assert item["status"] == "draft"
    assert item["reveal_answers_after_finalize"] is True
    assert item["finalized_at_ms"] is None


def test_list_exams_filters_by_status(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _make_exam(db_path, code="EXAM-STATUS-FILTER")
    assert exams.list_exams(db_path, status="draft")["total"] == 1
    assert exams.list_exams(db_path, status="open")["total"] == 0


def test_get_exam_detail_includes_template_and_assignment_summary(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    template = _make_template(db_path, code="TPL-DETAIL")
    now = time_policy.utc_now_ms()
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="alice", code="EXAM-DETAIL", title="Kỳ thi",
        template_id=template["id"], target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
    )
    exams.create_assignments(
        db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
        users=[{"username": "u1", "display_name": "User 1"},
               {"username": "u2", "display_name": "User 2"}],
        audience_code="nvkt",
    )
    detail = exams.get_exam_detail(db_path, exam["id"])

    assert set(detail) == _EXAM_DTO_KEYS | {"template", "assignment_summary"}
    assert detail["template"] == {
        "id": template["id"], "code": "TPL-DETAIL", "title": "Template",
    }
    assert detail["assignment_summary"] == {
        "total": 2, "assigned": 2, "completed": 0,
        "expired": 0, "cancelled": 0, "in_progress": 0,
    }


def test_get_exam_detail_raises_not_found(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    with pytest.raises(TrainingError) as exc:
        exams.get_exam_detail(db_path, "missing")
    assert exc.value.code == "NOT_FOUND"
    assert exc.value.status == 404


def test_get_exam_assignments_dto_strips_sensitive_fields(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _make_exam(db_path, code="EXAM-ASG-DTO")
    exams.create_assignments(
        db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
        users=[{"username": "u1", "display_name": "User 1",
                "team_code": "t1", "team_name": "Tổ 1"}],
        audience_code="nvkt",
    )
    items = exams.get_exam_assignments_dto(db_path, exam["id"])

    assert len(items) == 1
    item = items[0]
    assert _ASSIGNMENT_FORBIDDEN_KEYS.isdisjoint(item)
    assert item["username"] == "u1"
    assert item["display_name"] == "User 1"
    assert item["team_code"] == "t1"
    assert item["status"] == "assigned"
    assert item["attempt_status"] is None


def test_list_assignable_users_never_returns_password(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    monkeypatch.setattr(
        "services.training_exam_service._get_all_users",
        lambda: [
            {"username": "u1", "name": "User One", "password": "hash", "is_active": 1},
            {"username": "u2", "name": "User Two", "password": "hash", "is_active": 0},
        ],
    )
    items = exams.list_assignable_users(db_path)

    assert len(items) == 1
    item = items[0]
    assert set(item) == {"username", "display_name"}
    assert "password" not in item
    assert item == {"username": "u1", "display_name": "User One"}


def test_list_assignable_users_filters_inactive_and_query(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    monkeypatch.setattr(
        "services.training_exam_service._get_all_users",
        lambda: [
            {"username": "alice", "name": "Alice Smith", "password": "h", "is_active": 1},
            {"username": "bob", "name": "Bob Jones", "password": "h", "is_active": 1},
        ],
    )
    assert len(exams.list_assignable_users(db_path)) == 2
    matched = exams.list_assignable_users(db_path, q="ALICE")
    assert len(matched) == 1
    assert matched[0]["username"] == "alice"


# --- Route-level ---

def test_read_routes_work_for_admin_client(monkeypatch, tmp_path):
    from blueprints import training_routes

    for client in _client(monkeypatch, tmp_path):
        db_path = training_routes.config.TRAINING_DB_PATH
        version_ids = _publish_questions(db_path)
        template = exams.create_template(
            db_path, unit_code="son_tay", actor="alice", code="TPL-ROUTE", title="Template",
            target_audience_code="nvkt", question_version_ids=version_ids,
            duration_seconds=600, pass_score_percent=80.0,
        )
        now = time_policy.utc_now_ms()
        exam = exams.create_exam(
            db_path, unit_code="son_tay", actor="alice", code="EXAM-ROUTE", title="Kỳ thi",
            template_id=template["id"], target_audience_code="nvkt", start_at_ms=now,
            end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
        )
        responses = {
            "templates": client.get("/api/training/templates"),
            "template_detail": client.get(f"/api/training/templates/{template['id']}"),
            "exams": client.get("/api/training/exams"),
            "exam_detail": client.get(f"/api/training/exams/{exam['id']}"),
            "assignments": client.get(f"/api/training/exams/{exam['id']}/assignments"),
            "users": client.get("/api/training/users"),
        }

    assert [resp.status_code for resp in responses.values()] == [200] * 6
    assert responses["templates"].get_json()["total"] == 1
    assert responses["exams"].get_json()["total"] == 1
    assert responses["template_detail"].get_json()["id"] == template["id"]
    assert responses["exam_detail"].get_json()["template"]["id"] == template["id"]


def test_read_routes_return_404_for_missing_template_and_exam(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        responses = [
            client.get("/api/training/templates/missing"),
            client.get("/api/training/exams/missing"),
        ]

    assert [resp.status_code for resp in responses] == [404, 404]
    assert all(resp.get_json()["error"]["code"] == "NOT_FOUND" for resp in responses)


def test_read_routes_reject_learner_role(monkeypatch, tmp_path):
    for client, _ in _role_client(monkeypatch, tmp_path, username="learner", roles=("learner",)):
        responses = [
            client.get("/api/training/templates"),
            client.get("/api/training/templates/missing"),
            client.get("/api/training/exams"),
            client.get("/api/training/exams/missing"),
            client.get("/api/training/exams/missing/assignments"),
            client.get("/api/training/users"),
        ]

    assert [resp.status_code for resp in responses] == [403] * 6
    assert all(resp.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED" for resp in responses)


def test_assignments_route_returns_empty_for_non_finalized_exam(monkeypatch, tmp_path):
    from blueprints import training_routes

    for client in _client(monkeypatch, tmp_path):
        db_path = training_routes.config.TRAINING_DB_PATH
        exam = _make_exam(db_path, code="EXAM-EMPTY-ASG")
        response = client.get(f"/api/training/exams/{exam['id']}/assignments")

    assert response.status_code == 200
    assert response.get_json() == {"items": []}


def test_exam_detail_route_has_no_store_cache_header(monkeypatch, tmp_path):
    from blueprints import training_routes

    for client in _client(monkeypatch, tmp_path):
        db_path = training_routes.config.TRAINING_DB_PATH
        exam = _make_exam(db_path, code="EXAM-NOSTORE")
        response = client.get(f"/api/training/exams/{exam['id']}")

    assert response.status_code == 200
    assert response.headers["Cache-Control"] == "no-store"


def test_users_route_never_exposes_password(monkeypatch, tmp_path):
    monkeypatch.setattr(
        "services.training_exam_service._get_all_users",
        lambda: [{"username": "u1", "name": "User One", "password": "secret", "is_active": 1}],
    )
    for client in _client(monkeypatch, tmp_path):
        response = client.get("/api/training/users")

    assert response.status_code == 200
    serialized = response.get_data(as_text=True)
    assert "password" not in serialized
    assert "secret" not in serialized
    assert response.get_json()["items"] == [{"username": "u1", "display_name": "User One"}]


# --- Template create + validation via route ---

def _publish_via_route(client, batch):
    imported = client.post(
        "/api/training/questions/import", headers={"X-CSRF-Token": "csrf"}, json=batch,
    )
    assert imported.status_code == 201
    version_ids = imported.get_json()["version_ids"]
    for vid in version_ids:
        approved = client.post(
            f"/api/training/questions/{vid}/approve", headers={"X-CSRF-Token": "csrf"},
        )
        published = client.post(
            f"/api/training/questions/{vid}/publish", headers={"X-CSRF-Token": "csrf"},
        )
        assert approved.status_code == 200
        assert published.status_code == 200
    return version_ids


def _two_question_route_batch():
    batch = copy.deepcopy(QUESTION_BATCH)
    batch["batch"]["requested_count"] = 2
    batch["questions"] = [
        {**QUESTION_BATCH["questions"][0], "local_ref": "Q1", "stem": "Câu lifecycle 1?"},
        {**QUESTION_BATCH["questions"][0], "local_ref": "Q2", "stem": "Câu lifecycle 2?"},
    ]
    return batch


def _post_template(client, payload):
    return client.post(
        "/api/training/templates", headers={"X-CSRF-Token": "csrf"}, json=payload,
    )


def test_create_template_route_success(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        version_ids = _publish_via_route(client, _two_question_route_batch())
        response = _post_template(client, {
            "code": "TPL-ROUTE-OK", "title": "Template route",
            "target_audience_code": "nvkt",
            "question_version_ids": version_ids,
            "duration_seconds": 600, "pass_score_percent": 80.0,
            "shuffle_questions": True, "shuffle_options": False,
        })

    assert response.status_code == 201
    body = response.get_json()
    assert body["code"] == "TPL-ROUTE-OK"
    assert body["total_questions"] == 2
    assert body["locked"] == 0


def test_create_template_route_rejects_empty_questions(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        response = _post_template(client, {
            "code": "TPL-EMPTY-ROUTE", "title": "Empty",
            "target_audience_code": "nvkt", "question_version_ids": [],
            "duration_seconds": 600, "pass_score_percent": 80.0,
        })

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "TEMPLATE_QUESTION_INVALID"


def test_create_template_route_rejects_unpublished_question(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        imported = client.post(
            "/api/training/questions/import", headers={"X-CSRF-Token": "csrf"}, json=QUESTION_BATCH,
        )
        version_id = imported.get_json()["version_ids"][0]
        response = _post_template(client, {
            "code": "TPL-UNPUB-ROUTE", "title": "Unpublished",
            "target_audience_code": "nvkt", "question_version_ids": [version_id],
            "duration_seconds": 600, "pass_score_percent": 80.0,
        })

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "TEMPLATE_QUESTION_INVALID"


def test_create_template_route_rejects_duplicate_question(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        version_ids = _publish_via_route(client, QUESTION_BATCH)
        response = _post_template(client, {
            "code": "TPL-DUP-ROUTE", "title": "Duplicate",
            "target_audience_code": "nvkt",
            "question_version_ids": [version_ids[0], version_ids[0]],
            "duration_seconds": 600, "pass_score_percent": 80.0,
        })

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "TEMPLATE_QUESTION_DUPLICATE"


def test_create_template_route_rejects_audience_mismatch(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        # QUESTION_BATCH gắn audience_codes=["nvkt"]; tạo template cho đối tượng khác.
        version_ids = _publish_via_route(client, QUESTION_BATCH)
        response = _post_template(client, {
            "code": "TPL-AUD-ROUTE", "title": "Audience mismatch",
            "target_audience_code": "kinh_doanh",
            "question_version_ids": version_ids,
            "duration_seconds": 600, "pass_score_percent": 80.0,
        })

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "TEMPLATE_QUESTION_INVALID"


def test_template_becomes_immutable_after_exam_created(monkeypatch, tmp_path):
    # Không có route update template; kiểm chứng bất biến ở service layer mà route dựa vào.
    db_path = _setup(monkeypatch, tmp_path)
    template = _make_template(db_path, code="TPL-IMMUTABLE-ROUTE")
    now = time_policy.utc_now_ms()
    exams.create_exam(
        db_path, unit_code="son_tay", actor="alice", code="EXAM-IMMUTABLE-ROUTE", title="Kỳ thi",
        template_id=template["id"], target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
    )
    version_ids = [item["question_version_id"]
                   for item in exams.get_template_items(db_path, template["id"])]

    with pytest.raises(TrainingError) as exc:
        exams.update_template(
            db_path, unit_code="son_tay", actor="alice", template_id=template["id"],
            question_version_ids=list(reversed(version_ids)),
            shuffle_questions=True, shuffle_options=False,
        )

    assert exc.value.code == "TEMPLATE_IMMUTABLE"
    assert exc.value.status == 409


# --- Exam lifecycle + assignments via route ---

def test_exam_full_lifecycle_via_route(monkeypatch, tmp_path):
    from blueprints import training_routes

    for client in _client(monkeypatch, tmp_path):
        db_path = training_routes.config.TRAINING_DB_PATH
        template = _make_template(db_path, code="TPL-LIFECYCLE")
        now = time_policy.utc_now_ms()
        create = client.post(
            "/api/training/exams", headers={"X-CSRF-Token": "csrf"},
            json={
                "code": "EXAM-LIFECYCLE", "title": "Kỳ thi", "template_id": template["id"],
                "target_audience_code": "nvkt", "start_at_ms": now,
                "end_at_ms": now + 3_600_000, "duration_seconds": 600,
            },
        )
        exam_id = create.get_json()["id"]
        ready = client.post(f"/api/training/exams/{exam_id}/ready", headers={"X-CSRF-Token": "csrf"})
        assign = client.post(
            f"/api/training/exams/{exam_id}/assignments", headers={"X-CSRF-Token": "csrf"},
            json={"users": [{"username": "u1", "display_name": "User 1"}], "audience_code": "nvkt"},
        )
        opened = client.post(f"/api/training/exams/{exam_id}/open", headers={"X-CSRF-Token": "csrf"})
        closed = client.post(f"/api/training/exams/{exam_id}/close", headers={"X-CSRF-Token": "csrf"})
        finalized = client.post(f"/api/training/exams/{exam_id}/finalize", headers={"X-CSRF-Token": "csrf"})

    assert create.status_code == 201
    assert ready.status_code == 200
    assert assign.status_code == 201
    assert opened.status_code == 200
    assert closed.status_code == 200
    assert set(closed.get_json()["recovery_summary"]) == {
        "processed_attempt_ids", "already_completed_ids", "failed_attempts",
    }
    assert finalized.status_code == 200
    assert isinstance(finalized.get_json()["revision"], int)


def test_duplicate_assignment_returns_409(monkeypatch, tmp_path):
    from blueprints import training_routes

    for client in _client(monkeypatch, tmp_path):
        db_path = training_routes.config.TRAINING_DB_PATH
        exam = _make_exam(db_path, code="EXAM-DUP-ASSIGN-ROUTE")
        first = client.post(
            f"/api/training/exams/{exam['id']}/assignments", headers={"X-CSRF-Token": "csrf"},
            json={"users": [{"username": "u1"}], "audience_code": "nvkt"},
        )
        duplicate = client.post(
            f"/api/training/exams/{exam['id']}/assignments", headers={"X-CSRF-Token": "csrf"},
            json={"users": [{"username": "u1"}], "audience_code": "nvkt"},
        )

    assert first.status_code == 201
    assert duplicate.status_code == 409
    assert duplicate.get_json()["error"]["code"] == "ASSIGNMENT_ALREADY_EXISTS"


def test_assignment_blocked_after_open(monkeypatch, tmp_path):
    from blueprints import training_routes

    for client in _client(monkeypatch, tmp_path):
        db_path = training_routes.config.TRAINING_DB_PATH
        exam = _make_exam(db_path, code="EXAM-ASSIGN-AFTER-OPEN")
        client.post(f"/api/training/exams/{exam['id']}/ready", headers={"X-CSRF-Token": "csrf"})
        client.post(f"/api/training/exams/{exam['id']}/open", headers={"X-CSRF-Token": "csrf"})
        response = client.post(
            f"/api/training/exams/{exam['id']}/assignments", headers={"X-CSRF-Token": "csrf"},
            json={"users": [{"username": "u1"}], "audience_code": "nvkt"},
        )

    assert response.status_code == 409
    assert response.get_json()["error"]["code"] == "CONFLICT"


# --- RBAC + CSRF ---

def test_editor_cannot_access_exam_routes(monkeypatch, tmp_path):
    for client, _ in _role_client(monkeypatch, tmp_path, username="editor", roles=("editor",)):
        responses = [
            client.get("/api/training/exams"),
            client.get("/api/training/exams/missing"),
            client.post("/api/training/exams", headers={"X-CSRF-Token": "csrf"}, json={}),
            client.post("/api/training/exams/missing/cancel", headers={"X-CSRF-Token": "csrf"}),
        ]

    assert [response.status_code for response in responses] == [403] * 4
    assert all(response.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED" for response in responses)


def test_create_template_without_csrf_returns_400(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        response = client.post(  # cố tình thiếu X-CSRF-Token
            "/api/training/templates",
            json={
                "code": "TPL-NO-CSRF", "title": "No CSRF",
                "target_audience_code": "nvkt", "question_version_ids": [],
                "duration_seconds": 600, "pass_score_percent": 80.0,
            },
        )

    assert response.status_code == 400
    assert response.get_json()["error"] == "CSRF token không hợp lệ"


# --- XSS guard: dynamic data must never reach innerHTML ---

def test_exam_and_template_js_never_assign_dynamic_innerhtml():
    import re
    from pathlib import Path

    repo_root = Path(__file__).resolve().parent.parent
    sources = [
        repo_root / "static" / "js" / "training-exams.js",
        repo_root / "static" / "js" / "training-templates.js",
    ]
    assignment = re.compile(r"\.innerHTML\s*=\s*([^;\n]+)")
    literal = re.compile(r"'[^']*'")
    for source_path in sources:
        source = source_path.read_text(encoding="utf-8")
        for match in assignment.finditer(source):
            rhs = match.group(1).strip()
            assert literal.fullmatch(rhs), (
                f"{source_path.name}: .innerHTML gán giá trị động (không phải hằng chuỗi): {rhs}"
            )


# --- Finalize blocking: corrupted attempt snapshot surfaces blocking_attempts ---

def test_finalize_route_reports_blocking_attempts(monkeypatch, tmp_path):
    from blueprints import training_routes

    for client in _client(monkeypatch, tmp_path):
        db_path = training_routes.config.TRAINING_DB_PATH
        exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
        attempt = attempts.start_attempt(
            db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id,
        )
        # Làm hỏng snapshot: sửa nội dung item nhưng không cập nhật checksum.
        conn = training_db.write_connection(db_path)
        try:
            conn.execute(
                "UPDATE exam_attempt_items SET stem=stem||' (sửa)' WHERE attempt_id=?",
                (attempt["attempt_id"],),
            )
            conn.commit()
        finally:
            conn.close()
        closed = client.post(f"/api/training/exams/{exam_id}/close", headers={"X-CSRF-Token": "csrf"})
        finalized = client.post(f"/api/training/exams/{exam_id}/finalize", headers={"X-CSRF-Token": "csrf"})

    assert closed.status_code == 200
    failed = closed.get_json()["recovery_summary"]["failed_attempts"]
    assert any(item["attempt_id"] == attempt["attempt_id"] for item in failed)
    assert finalized.status_code == 409
    details = finalized.get_json()["error"]["details"]
    assert "blocking_attempts" in details
    assert any(item["attempt_id"] == attempt["attempt_id"] for item in details["blocking_attempts"])
