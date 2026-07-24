import threading

import pytest

from training import db as training_db
from training import migrations, time_policy
from training.errors import ErrorCode, TrainingError
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


def _create_transition_exam(db_path, *, code="EXAM-TRANSITION"):
    version_ids = _publish_questions(db_path)
    template = es.create_template(
        db_path, unit_code="son_tay", actor="alice", code=f"TPL-{code}", title="Template",
        target_audience_code="nvkt", question_version_ids=version_ids,
        duration_seconds=600, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    return es.create_exam(
        db_path, unit_code="son_tay", actor="alice", code=code, title="Exam",
        template_id=template["id"], target_audience_code="nvkt", start_at_ms=now - 1_000,
        end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
    )


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
    with pytest.raises(TrainingError) as exc:
        es.create_template(
            db_path, unit_code="son_tay", actor="alice",
            code="TPL-BAD", title="Bad",
            target_audience_code="nvkt",
            question_version_ids=result["version_ids"],
            duration_seconds=1500, pass_score_percent=80.0,
        )
    assert exc.value.code == "TEMPLATE_QUESTION_INVALID"


def test_create_template_rejects_no_question_versions(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)

    with pytest.raises(TrainingError) as exc:
        es.create_template(
            db_path, unit_code="son_tay", actor="alice",
            code="TPL-EMPTY", title="Empty", target_audience_code="nvkt",
            question_version_ids=[], duration_seconds=1500, pass_score_percent=80.0,
        )

    assert exc.value.code == "TEMPLATE_QUESTION_INVALID"


def test_create_template_rejects_missing_question_version(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)

    with pytest.raises(TrainingError) as exc:
        es.create_template(
            db_path, unit_code="son_tay", actor="alice",
            code="TPL-MISSING", title="Missing", target_audience_code="nvkt",
            question_version_ids=["missing-version"],
            duration_seconds=1500, pass_score_percent=80.0,
        )

    assert exc.value.code == "TEMPLATE_QUESTION_INVALID"


def test_create_template_rejects_duplicate_question_version(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_id = _publish_questions(db_path)[0]

    with pytest.raises(TrainingError) as exc:
        es.create_template(
            db_path, unit_code="son_tay", actor="alice",
            code="TPL-DUP", title="Duplicate", target_audience_code="nvkt",
            question_version_ids=[version_id, version_id],
            duration_seconds=1500, pass_score_percent=80.0,
        )

    assert exc.value.code == "TEMPLATE_QUESTION_DUPLICATE"


def test_update_template_rejects_question_and_shuffle_changes_after_exam_created(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    template = es.create_template(
        db_path, unit_code="son_tay", actor="alice", code="TPL-IMM", title="Template",
        target_audience_code="nvkt", question_version_ids=version_ids,
        duration_seconds=600, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    es.create_exam(
        db_path, unit_code="son_tay", actor="alice", code="EXAM-IMM", title="Exam",
        template_id=template["id"], target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 60_000, duration_seconds=600, pass_score_percent=80.0,
    )

    with pytest.raises(TrainingError) as exc:
        es.update_template(
            db_path, unit_code="son_tay", actor="alice", template_id=template["id"],
            question_version_ids=list(reversed(version_ids)), shuffle_questions=True,
            shuffle_options=False,
        )
    assert exc.value.code == "TEMPLATE_IMMUTABLE"


def test_create_template_rejects_question_with_incompatible_audience(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_id = _publish_questions(db_path)[0]
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO question_audiences (question_version_id, audience_code) VALUES (?, ?)",
            (version_id, "kinh_doanh"),
        )
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc:
        es.create_template(
            db_path, unit_code="son_tay", actor="alice",
            code="TPL-AUDIENCE", title="Audience", target_audience_code="nvkt",
            question_version_ids=[version_id],
            duration_seconds=1500, pass_score_percent=80.0,
        )

    assert exc.value.code == "TEMPLATE_QUESTION_INVALID"


def test_update_template_rejects_question_with_incompatible_audience(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    template = es.create_template(
        db_path, unit_code="son_tay", actor="alice", code="TPL-UPDATE-AUDIENCE", title="Template",
        target_audience_code="nvkt", question_version_ids=version_ids,
        duration_seconds=600, pass_score_percent=80.0,
    )
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO question_audiences (question_version_id, audience_code) VALUES (?, ?)",
            (version_ids[0], "kinh_doanh"),
        )
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc:
        es.update_template(
            db_path, unit_code="son_tay", actor="alice", template_id=template["id"],
            question_version_ids=version_ids, shuffle_questions=False, shuffle_options=False,
        )

    assert exc.value.code == "TEMPLATE_QUESTION_INVALID"


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


def test_create_exam_rejects_missing_template_with_stable_error(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    now = time_policy.utc_now_ms()

    with pytest.raises(TrainingError) as exc:
        es.create_exam(
            db_path, unit_code="son_tay", actor="alice", code="EXAM-MISSING-TEMPLATE",
            title="Kỳ thi", template_id="missing-template", target_audience_code="nvkt",
            start_at_ms=now, end_at_ms=now + 3_600_000, duration_seconds=600,
            pass_score_percent=80.0,
        )

    assert exc.value.code == "NOT_FOUND"
    assert exc.value.status == 404


def test_create_exam_rejects_template_audience_mismatch(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    template = es.create_template(
        db_path, unit_code="son_tay", actor="alice", code="TPL-AUDIENCE-CHAIN", title="Template",
        target_audience_code="nvkt", question_version_ids=version_ids,
        duration_seconds=600, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()

    with pytest.raises(TrainingError) as exc:
        es.create_exam(
            db_path, unit_code="son_tay", actor="alice", code="EXAM-AUDIENCE-CHAIN",
            title="Kỳ thi", template_id=template["id"], target_audience_code="b2a",
            start_at_ms=now, end_at_ms=now + 3_600_000, duration_seconds=600,
            pass_score_percent=80.0,
        )

    assert exc.value.code == "AUDIENCE_MISMATCH"
    assert exc.value.status == 400


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


def test_open_requires_ready_and_does_not_audit_draft(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path)

    with pytest.raises(TrainingError) as exc_info:
        es.open_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])

    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409
    assert es.get_exam(db_path, exam["id"])["status"] == "draft"
    conn = training_db.read_connection(db_path)
    try:
        assert conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='open_exam' AND entity_id=?",
            (exam["id"],),
        ).fetchone()["c"] == 0
    finally:
        conn.close()


def test_close_requires_open_and_is_idempotent_after_closed(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path)

    with pytest.raises(TrainingError) as exc_info:
        es.close_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409

    es.ready_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    es.close_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    es.close_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    assert es.get_exam(db_path, exam["id"])["status"] == "closed"


@pytest.mark.parametrize("initial_status", ["draft", "ready"])
def test_cancel_allows_only_pre_open_states(monkeypatch, tmp_path, initial_status):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code=f"EXAM-CANCEL-{initial_status}")
    if initial_status == "ready":
        es.ready_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])

    es.cancel_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])

    assert es.get_exam(db_path, exam["id"])["status"] == "cancelled"


def test_cancel_open_exam_is_conflict_and_cannot_be_reversed(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-CANCEL-OPEN")
    es.ready_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])

    with pytest.raises(TrainingError) as exc_info:
        es.cancel_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409
    assert es.get_exam(db_path, exam["id"])["status"] == "open"

    es.close_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    with pytest.raises(TrainingError) as reopen_error:
        es.open_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    assert reopen_error.value.code == "CONFLICT"
    assert reopen_error.value.status == 409


@pytest.mark.parametrize(
    ("action", "status"),
    [
        ("ready", "ready"), ("ready", "open"), ("ready", "closed"), ("ready", "cancelled"),
        ("open", "closed"), ("open", "cancelled"),
        ("close", "ready"), ("close", "cancelled"),
        ("cancel", "closed"), ("cancel", "cancelled"),
    ],
)
def test_invalid_exam_transitions_are_conflicts(monkeypatch, tmp_path, action, status):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code=f"EXAM-INVALID-{action}-{status}")
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET status=? WHERE id=?", (status, exam["id"]))
        conn.commit()
    finally:
        conn.close()

    service = {
        "ready": es.ready_exam,
        "open": es.open_exam,
        "close": es.close_exam,
        "cancel": es.cancel_exam,
    }[action]
    with pytest.raises(TrainingError) as exc_info:
        service(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])

    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409
    assert es.get_exam(db_path, exam["id"])["status"] == status


def test_concurrent_ready_writes_one_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-CONCURRENT-READY")
    barrier = threading.Barrier(2)

    def ready_exam():
        barrier.wait()
        try:
            es.ready_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
        except TrainingError:
            pass

    threads = [threading.Thread(target=ready_exam) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=15)

    conn = training_db.read_connection(db_path)
    try:
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='ready_exam' AND entity_id=?", (exam["id"],)
        ).fetchone()["c"]
    finally:
        conn.close()
    assert es.get_exam(db_path, exam["id"])["status"] == "ready"
    assert audit_count == 1


def test_concurrent_close_writes_one_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-CONCURRENT-CLOSE")
    es.ready_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])
    barrier = threading.Barrier(2)

    def close_exam():
        barrier.wait()
        es.close_exam(db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"])

    threads = [threading.Thread(target=close_exam) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=15)

    conn = training_db.read_connection(db_path)
    try:
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='close_exam' AND entity_id=?", (exam["id"],)
        ).fetchone()["c"]
    finally:
        conn.close()
    assert es.get_exam(db_path, exam["id"])["status"] == "closed"
    assert audit_count == 1


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


def test_concurrent_open_writes_one_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="a", code="T", title="T",
        target_audience_code="nvkt", question_version_ids=version_ids,
        duration_seconds=600, pass_score_percent=80.0)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="a", code="E", title="E",
        template_id=tpl["id"], target_audience_code="nvkt", start_at_ms=now - 1000,
        end_at_ms=now + 3600_000, duration_seconds=600, pass_score_percent=80.0)
    es.ready_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
    barrier = threading.Barrier(2)
    errors = []

    def open_exam():
        barrier.wait()
        try:
            es.open_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
        except Exception as exc:
            errors.append(exc)

    threads = [threading.Thread(target=open_exam) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=15)
    conn = training_db.read_connection(db_path)
    try:
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='open_exam' AND entity_id=?", (exam["id"],)
        ).fetchone()["c"]
    finally:
        conn.close()
    assert not errors
    assert es.get_exam(db_path, exam["id"])["status"] == "open"
    assert audit_count == 1


def test_concurrent_cancel_writes_one_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_ids = _publish_questions(db_path)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="a", code="T", title="T",
        target_audience_code="nvkt", question_version_ids=version_ids,
        duration_seconds=600, pass_score_percent=80.0)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="a", code="E", title="E",
        template_id=tpl["id"], target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3600_000, duration_seconds=600, pass_score_percent=80.0)
    barrier = threading.Barrier(2)

    def cancel_exam():
        barrier.wait()
        try:
            es.cancel_exam(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"])
        except TrainingError:
            pass

    threads = [threading.Thread(target=cancel_exam) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=15)
    conn = training_db.read_connection(db_path)
    try:
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='cancel_exam' AND entity_id=?", (exam["id"],)
        ).fetchone()["c"]
    finally:
        conn.close()
    assert es.get_exam(db_path, exam["id"])["status"] == "cancelled"
    assert audit_count == 1


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


def test_create_assignments_rejects_exam_audience_mismatch(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-ASSIGNMENT-AUDIENCE")

    with pytest.raises(TrainingError) as exc:
        es.create_assignments(
            db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
            users=[{"username": "u1"}], audience_code="b2a",
        )

    assert exc.value.code == "AUDIENCE_MISMATCH"
    assert exc.value.status == 400


def test_create_assignments_rejects_configured_user_outside_audience(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-CONFIGURED-USER")
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO training_user_audiences (id, username, audience_code) VALUES (?, ?, ?)",
            ("aud-u1-b2a", "u1", "b2a"),
        )
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc:
        es.create_assignments(
            db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
            users=[{"username": "u1"}], audience_code="nvkt",
        )

    assert exc.value.code == "AUDIENCE_MISMATCH"
    assert exc.value.status == 400


def test_create_assignments_allows_unconfigured_user_and_snapshots_audience(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-UNCONFIGURED-USER")

    assignment_id = es.create_assignments(
        db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
        users=[{"username": "u1"}], audience_code="nvkt",
    )[0]

    assert es.get_assignment(db_path, assignment_id)["audience_code"] == "nvkt"


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
    with pytest.raises(TrainingError) as exc_info:
        es.create_assignments(db_path, unit_code="son_tay", actor="a", exam_id=exam["id"],
            users=[{"username": "u1", "display_name": "U1 dup"}], audience_code="nvkt")

    assert exc_info.value.code == "ASSIGNMENT_ALREADY_EXISTS"
    assert exc_info.value.status == 409


@pytest.mark.parametrize("status", ["open", "closed", "cancelled"])
def test_create_assignments_rejects_non_preparation_exam_states(monkeypatch, tmp_path, status):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code=f"EXAM-ASSIGN-{status}")
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET status=? WHERE id=?", (status, exam["id"]))
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc_info:
        es.create_assignments(
            db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
            users=[{"username": "u1"}], audience_code="nvkt",
        )

    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409
    assert es.get_assignments_for_exam(db_path, exam["id"]) == []


def test_create_assignments_rejects_finalized_exam(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-ASSIGN-FINALIZED")
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "UPDATE exam_events SET finalized_at_ms=? WHERE id=?",
            (time_policy.utc_now_ms(), exam["id"]),
        )
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc_info:
        es.create_assignments(
            db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
            users=[{"username": "u1"}], audience_code="nvkt",
        )

    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409


def test_create_assignments_rejects_duplicate_username_in_batch_atomically(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-ASSIGN-BATCH-DUPLICATE")

    with pytest.raises(TrainingError) as exc_info:
        es.create_assignments(
            db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
            users=[{"username": "u1"}, {"username": "u1"}], audience_code="nvkt",
        )

    assert exc_info.value.code == "ASSIGNMENT_ALREADY_EXISTS"
    assert exc_info.value.status == 409
    assert es.get_assignments_for_exam(db_path, exam["id"]) == []


def test_create_assignments_rejects_existing_assignment_atomically(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-ASSIGN-EXISTING")
    es.create_assignments(
        db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
        users=[{"username": "u1"}], audience_code="nvkt",
    )

    with pytest.raises(TrainingError) as exc_info:
        es.create_assignments(
            db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
            users=[{"username": "u2"}, {"username": "u1"}], audience_code="nvkt",
        )

    assert exc_info.value.code == "ASSIGNMENT_ALREADY_EXISTS"
    assert exc_info.value.status == 409
    assert [row["username"] for row in es.get_assignments_for_exam(db_path, exam["id"])] == ["u1"]


def test_create_assignments_rejects_missing_username_as_validation_error(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam = _create_transition_exam(db_path, code="EXAM-ASSIGN-MISSING-USERNAME")

    with pytest.raises(TrainingError) as exc_info:
        es.create_assignments(
            db_path, unit_code="son_tay", actor="alice", exam_id=exam["id"],
            users=[{"username": None}], audience_code="nvkt",
        )

    assert exc_info.value.code == ErrorCode.VALIDATION_ERROR
    assert exc_info.value.status == 400
    assert es.get_assignments_for_exam(db_path, exam["id"]) == []
