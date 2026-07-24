import sqlite3
import threading

from training import db as training_db
from training import migrations, time_policy
from training.errors import TrainingError
from services import training_attempt_service as attempts
from services import training_exam_service as exams
from services import training_question_service as questions
from services.training_catalog_service import seed_defaults


VALID_BATCH = {
    "schema_version": "1.0",
    "batch": {
        "title": "Concurrent smoke", "language": "vi",
        "source_document_version_ids": ["docver-001"],
        "target_audience_codes": ["nvkt"], "requested_count": 3,
    },
    "questions": [
        {
            "local_ref": f"Q{i}", "type": "single_choice", "stem": f"Câu {i}?",
            "options": [{"id": "A", "text": "sai"}, {"id": "B", "text": "đúng"}],
            "correct_option_ids": ["B"], "explanation": "Giải thích",
            "distractor_rationales": {"A": "sai"},
            "classification": {"domain_code": "quality", "topic_codes": ["topic"]},
            "difficulty": "easy",
            "evidence": [{
                "document_version_id": "docver-001", "block_id": "DOC-B001",
                "extraction_revision": 1, "quoted_text": "q", "supports": "correct_answer",
            }],
        }
        for i in range(1, 4)
    ],
}


def _setup_attempt(monkeypatch, tmp_path):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", "son_tay")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    questions.reset_duplicate_cache()
    imported = questions.import_question_batch(
        db_path, unit_code="son_tay", actor="editor", batch=VALID_BATCH, status="draft",
    )
    for version_id in imported["version_ids"]:
        questions.add_review_action(
            db_path, unit_code="son_tay", actor="manager", version_id=version_id, action="approve",
        )
        questions.publish_question_version(
            db_path, unit_code="son_tay", actor="manager", version_id=version_id,
        )
    template = exams.create_template(
        db_path, unit_code="son_tay", actor="manager", code="SMOKE", title="Smoke",
        target_audience_code="nvkt", question_version_ids=imported["version_ids"],
        duration_seconds=600, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="manager", code="SMOKE", title="Smoke",
        template_id=template["id"], target_audience_code="nvkt", start_at_ms=now - 1_000,
        end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
    )
    exams.ready_exam(db_path, unit_code="son_tay", actor="manager", exam_id=exam["id"])
    exams.open_exam(db_path, unit_code="son_tay", actor="manager", exam_id=exam["id"])
    assignment_id = exams.create_assignments(
        db_path, unit_code="son_tay", actor="manager", exam_id=exam["id"],
        users=[{"username": "learner", "display_name": "Learner"}], audience_code="nvkt",
    )[0]
    started = attempts.start_attempt(
        db_path, unit_code="son_tay", actor="learner", assignment_id=assignment_id,
    )
    view = attempts.get_attempt_learner_view(db_path, started["attempt_id"])
    return db_path, exam["id"], started["attempt_id"], [item["item_id"] for item in view["items"]]


def _run_concurrently(count, worker):
    barrier = threading.Barrier(count)
    outcomes = []
    errors = []
    outcomes_lock = threading.Lock()

    def run(index):
        try:
            barrier.wait(timeout=10)
            outcome = worker(index)
            with outcomes_lock:
                outcomes.append(outcome)
        except Exception as exc:
            with outcomes_lock:
                errors.append(exc)

    threads = [threading.Thread(target=run, args=(index,)) for index in range(1, count + 1)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=20)
    assert not [thread for thread in threads if thread.is_alive()]
    return outcomes, errors


def _assert_no_sqlite_lock(errors):
    assert not [
        error for error in errors
        if isinstance(error, sqlite3.OperationalError) and "locked" in str(error).lower()
    ]


def test_fifty_concurrent_autosaves_keep_highest_revision(monkeypatch, tmp_path):
    db_path, _, attempt_id, item_ids = _setup_attempt(monkeypatch, tmp_path)
    item_id = item_ids[0]
    outcomes, errors = _run_concurrently(
        50,
        lambda revision: attempts.save_response(
            db_path, attempt_id=attempt_id, attempt_item_id=item_id,
            selected_option_ids=["B"], client_revision=revision,
        ),
    )

    _assert_no_sqlite_lock(errors)
    assert not errors
    assert outcomes
    assert all(outcome["accepted"] or outcome["stored_revision"] >= 1 for outcome in outcomes)
    response = attempts.get_attempt_learner_view(db_path, attempt_id)["items"][0]["response"]
    assert response == {"selected_option_ids": ["B"], "client_revision": 50}


def test_fifty_concurrent_submits_create_one_result_and_one_terminal_audit(monkeypatch, tmp_path):
    db_path, _, attempt_id, item_ids = _setup_attempt(monkeypatch, tmp_path)
    for item_id in item_ids:
        attempts.save_response(
            db_path, attempt_id=attempt_id, attempt_item_id=item_id,
            selected_option_ids=["B"], client_revision=1,
        )
    outcomes, errors = _run_concurrently(
        50,
        lambda _: attempts.submit_attempt(
            db_path, unit_code="son_tay", actor="learner", attempt_id=attempt_id,
        ),
    )

    _assert_no_sqlite_lock(errors)
    assert not errors
    assert len(outcomes) == 50
    assert len({outcome["result"]["score"] for outcome in outcomes}) == 1
    conn = training_db.read_connection(db_path)
    try:
        assert conn.execute("SELECT COUNT(*) AS count FROM exam_results WHERE attempt_id=?", (attempt_id,)).fetchone()["count"] == 1
        assert conn.execute(
            "SELECT COUNT(*) AS count FROM training_audit_log WHERE action='submit_attempt' AND entity_id=?",
            (attempt_id,),
        ).fetchone()["count"] == 1
    finally:
        conn.close()


def test_close_racing_autosave_and_submit_keeps_one_terminal_result_and_audits(monkeypatch, tmp_path):
    db_path, exam_id, attempt_id, item_ids = _setup_attempt(monkeypatch, tmp_path)
    item_id = item_ids[0]
    barrier = threading.Barrier(3)
    errors = []
    errors_lock = threading.Lock()

    def run(operation):
        try:
            barrier.wait(timeout=10)
            operation()
        except TrainingError:
            pass
        except Exception as exc:
            with errors_lock:
                errors.append(exc)

    workers = [
        threading.Thread(target=run, args=(lambda: attempts.save_response(
            db_path, attempt_id=attempt_id, attempt_item_id=item_id,
            selected_option_ids=["B"], client_revision=1,
        ),)),
        threading.Thread(target=run, args=(lambda: attempts.submit_attempt(
            db_path, unit_code="son_tay", actor="learner", attempt_id=attempt_id,
        ),)),
        threading.Thread(target=run, args=(lambda: exams.close_exam(
            db_path, unit_code="son_tay", actor="manager", exam_id=exam_id,
        ),)),
    ]
    for worker in workers:
        worker.start()
    for worker in workers:
        worker.join(timeout=20)
    assert not [worker for worker in workers if worker.is_alive()]

    _assert_no_sqlite_lock(errors)
    assert not errors
    attempt = attempts.get_attempt(db_path, attempt_id)
    assert attempt["status"] in {"submitted", "administratively_submitted", "timed_out"}
    conn = training_db.read_connection(db_path)
    try:
        assert conn.execute("SELECT COUNT(*) AS count FROM exam_results WHERE attempt_id=?", (attempt_id,)).fetchone()["count"] == 1
        for action in ("submit_attempt", "administratively_submit_attempt"):
            assert conn.execute(
                "SELECT COUNT(*) AS count FROM training_audit_log WHERE action=? AND entity_id=?",
                (action, attempt_id),
            ).fetchone()["count"] <= 1
        assert conn.execute(
            "SELECT COUNT(*) AS count FROM training_audit_log WHERE action='close_exam' AND entity_id=?",
            (exam_id,),
        ).fetchone()["count"] == 1
    finally:
        conn.close()
