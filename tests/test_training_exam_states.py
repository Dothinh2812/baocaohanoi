import pytest

from training import db as training_db
from training import migrations, time_policy
from training.errors import TrainingError
from services import training_exam_service as es
from services import training_question_service as qs
from services.training_catalog_service import seed_defaults


VALID_BATCH = {
    "schema_version": "1.0",
    "batch": {
        "title": "Test",
        "language": "vi",
        "source_document_version_ids": ["docver-001"],
        "target_audience_codes": ["nvkt"],
        "requested_count": 2,
    },
    "questions": [
        {
            "local_ref": f"Q{i}",
            "type": "single_choice",
            "stem": f"Câu {i}?",
            "options": [{"id": "A", "text": "sai"}, {"id": "B", "text": "đúng"}],
            "correct_option_ids": ["B"],
            "explanation": "giải thích",
            "distractor_rationales": {"A": "sai"},
            "classification": {"domain_code": "quality", "topic_codes": ["t1"]},
            "difficulty": "easy",
            "evidence": [
                {"document_version_id": "docver-001", "block_id": "DOC-B001",
                 "extraction_revision": 1, "quoted_text": "q", "supports": "correct_answer"}
            ],
        }
        for i in range(1, 3)
    ],
}


def _setup(monkeypatch, tmp_path):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", "son_tay")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    qs.reset_duplicate_cache()
    return db_path


def _publish_questions(db_path):
    result = qs.import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                       batch=VALID_BATCH, status="draft")
    for vid in result["version_ids"]:
        qs.add_review_action(db_path, unit_code="son_tay", actor="bob",
                             version_id=vid, action="approve")
        qs.publish_question_version(db_path, unit_code="son_tay", actor="bob", version_id=vid)
    return result["version_ids"]


def test_create_template_from_published_versions(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(
        db_path, unit_code="son_tay", actor="alice",
        code="TPL-001", title="Template test",
        target_audience_code="nvkt",
        question_version_ids=version_ids,
        duration_seconds=1500, pass_score_percent=80.0,
    )
    assert tpl["total_questions"] == 2
    assert tpl["locked"] == 0


def test_create_template_rejects_unpublished(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = qs.import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                       batch=VALID_BATCH, status="draft")
    with pytest.raises(TrainingError):
        es.create_template(
            db_path, unit_code="son_tay", actor="alice",
            code="TPL-BAD", title="Bad",
            target_audience_code="nvkt",
            question_version_ids=result["version_ids"],
            duration_seconds=1500, pass_score_percent=80.0,
        )


def test_create_exam_draft(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(
        db_path, unit_code="son_tay", actor="alice",
        code="TPL-001", title="T",
        target_audience_code="nvkt",
        question_version_ids=version_ids,
        duration_seconds=1500, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = es.create_exam(
        db_path, unit_code="son_tay", actor="alice",
        code="EXAM-001", title="Kỳ thi test", template_id=tpl["id"],
        target_audience_code="nvkt",
        start_at_ms=now, end_at_ms=now + 3600_000,
        duration_seconds=1500, pass_score_percent=80.0,
    )
    assert exam["status"] == "draft"


def test_exam_state_transitions(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(
        db_path, unit_code="son_tay", actor="alice",
        code="TPL-001", title="T",
        target_audience_code="nvkt",
        question_version_ids=version_ids,
        duration_seconds=1500, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = es.create_exam(
        db_path, unit_code="son_tay", actor="alice",
        code="EXAM-001", title="Kỳ thi", template_id=tpl["id"],
        target_audience_code="nvkt",
        start_at_ms=now - 1000, end_at_ms=now + 3600_000,
        duration_seconds=1500, pass_score_percent=80.0,
    )
    es.ready_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    exam = es.get_exam(db_path, exam["id"])
    assert exam["status"] == "open"


def test_open_exam_idempotent(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="a",
        code="T", title="T", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600, pass_score_percent=80.0)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="a",
        code="E", title="E", template_id=tpl["id"], target_audience_code="nvkt",
        start_at_ms=now - 1000, end_at_ms=now + 3600_000,
        duration_seconds=600, pass_score_percent=80.0)
    es.ready_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
    exam = es.get_exam(db_path, exam["id"])
    assert exam["status"] == "open"


def test_open_exam_rejected_before_start(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="a",
        code="T", title="T", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600, pass_score_percent=80.0)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="a",
        code="E", title="E", template_id=tpl["id"], target_audience_code="nvkt",
        start_at_ms=now + 3600_000, end_at_ms=now + 7200_000,
        duration_seconds=600, pass_score_percent=80.0)
    es.ready_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
    with pytest.raises(TrainingError) as exc:
        es.open_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
    assert exc.value.code == "EXAM_NOT_OPEN"


def test_create_assignments_snapshots_users(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="a",
        code="T", title="T", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600, pass_score_percent=80.0)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="a",
        code="E", title="E", template_id=tpl["id"], target_audience_code="nvkt",
        start_at_ms=now, end_at_ms=now + 3600_000,
        duration_seconds=600, pass_score_percent=80.0)
    assignments = es.create_assignments(
        db_path, unit_code="son_tay", actor="a", exam_id=exam["id"],
        users=[
            {"username": "u1", "display_name": "User 1", "team_code": "t1", "team_name": "Tổ 1"},
            {"username": "u2", "display_name": "User 2", "team_code": None, "team_name": None},
        ],
        audience_code="nvkt",
    )
    assert len(assignments) == 2
    a1 = es.get_assignment(db_path, assignments[0])
    assert a1["display_name"] == "User 1"
    assert a1["status"] == "assigned"


def test_assignment_unique_per_exam_username_audience(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="a",
        code="T", title="T", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600, pass_score_percent=80.0)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="a",
        code="E", title="E", template_id=tpl["id"], target_audience_code="nvkt",
        start_at_ms=now, end_at_ms=now + 3600_000,
        duration_seconds=600, pass_score_percent=80.0)
    es.create_assignments(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"],
        users=[{"username": "u1", "display_name": "U1"}], audience_code="nvkt")
    with pytest.raises(Exception):
        es.create_assignments(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"],
            users=[{"username": "u1", "display_name": "U1 dup"}], audience_code="nvkt")
