import pytest

from services import training_exam_service as exams
from tests.test_training_exam_states import _publish_questions, _setup
from tests.test_training_routes import _client, _role_client
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
