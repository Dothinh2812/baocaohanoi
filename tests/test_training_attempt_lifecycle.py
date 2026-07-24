import json
import threading

import pytest

from training import db as training_db
from training import migrations, time_policy
from training.errors import TrainingError
from services import training_exam_service as es
from services import training_attempt_service as att
from services import training_question_service as qs
from services.training_catalog_service import seed_defaults


VALID_BATCH = {
    "schema_version": "1.0",
    "batch": {
        "title": "Test", "language": "vi",
        "source_document_version_ids": ["docver-001"],
        "target_audience_codes": ["nvkt"], "requested_count": 3,
    },
    "questions": [
        {
            "local_ref": f"Q{i}", "type": "single_choice",
            "stem": f"Câu {i}?",
            "options": [{"id": "A", "text": "sai"}, {"id": "B", "text": "đúng"}],
            "correct_option_ids": ["B" if i != 2 else "A"],
            "explanation": f"Giải thích {i}",
            "distractor_rationales": {"A": "sai"},
            "classification": {"domain_code": "quality", "topic_codes": [f"topic_{i}"]},
            "difficulty": "easy",
            "evidence": [
                {"document_version_id": "docver-001", "block_id": "DOC-B001",
                 "extraction_revision": 1, "quoted_text": "q", "supports": "correct_answer"}
            ],
        }
        for i in range(1, 4)
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


def _make_open_exam_with_assignment(db_path, username="learner1"):
    result = qs.import_question_batch(db_path, unit_code="son_tay", actor="ed",
                                       batch=VALID_BATCH, status="draft")
    for vid in result["version_ids"]:
        qs.add_review_action(db_path, unit_code="son_tay", actor="mgr", version_id=vid, action="approve")
        qs.publish_question_version(db_path, unit_code="son_tay", actor="mgr", version_id=vid)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="mgr",
        code="T", title="T", target_audience_code="nvkt",
        question_version_ids=result["version_ids"], duration_seconds=600, pass_score_percent=80.0)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="mgr",
        code="E", title="E", template_id=tpl["id"], target_audience_code="nvkt",
        start_at_ms=now - 1000, end_at_ms=now + 3600_000,
        duration_seconds=600, pass_score_percent=80.0)
    es.ready_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam["id"])
    assignment_ids = es.create_assignments(db_path, unit_code="son_tay", actor="mgr",
        exam_id=exam["id"],
        users=[{"username": username, "display_name": "Learner"}],
        audience_code="nvkt")
    return exam["id"], assignment_ids[0]


# --- M3.3: start attempt + snapshot ---

def test_start_attempt_creates_active(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1",
                               assignment_id=assignment_id)
    assert result["status"] == "active"
    assert result["deadline_at_ms"] > result["started_at_ms"]
    assert result["attempt_id"]


def test_start_attempt_idempotent_returns_existing(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    r1 = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    r2 = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    assert r1["attempt_id"] == r2["attempt_id"]


def test_two_concurrent_starts_create_one_attempt(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    results = []
    barrier = threading.Barrier(2)

    def do_start():
        barrier.wait()
        try:
            r = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
            results.append(r["attempt_id"])
        except TrainingError:
            results.append(None)

    t1 = threading.Thread(target=do_start)
    t2 = threading.Thread(target=do_start)
    t1.start()
    t2.start()
    t1.join(timeout=15)
    t2.join(timeout=15)

    attempt_ids = [r for r in results if r is not None]
    assert len(set(attempt_ids)) == 1
    conn = training_db.read_connection(db_path)
    count = conn.execute(
        "SELECT COUNT(*) AS c FROM exam_attempts WHERE assignment_id=?", (assignment_id,)
    ).fetchone()["c"]
    conn.close()
    assert count == 1


def test_learner_view_has_no_scoring_data(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        assert "correct_option_ids" not in item
        assert "explanation" not in item
        assert "evidence" not in item
        assert "scoring_policy" not in item
        assert "max_score" not in item
        assert "options" in item
        assert "stem" in item
        assert "response" in item


def test_snapshot_has_correct_data(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    conn = training_db.read_connection(db_path)
    items = conn.execute(
        "SELECT * FROM exam_attempt_items WHERE attempt_id=?", (result["attempt_id"],)
    ).fetchall()
    opts = conn.execute(
        "SELECT * FROM exam_attempt_options WHERE attempt_item_id IN "
        "(SELECT id FROM exam_attempt_items WHERE attempt_id=?)",
        (result["attempt_id"],),
    ).fetchall()
    conn.close()
    assert len(items) == 3
    assert len(opts) == 6


# --- M3.4: autosave ---

def test_autosave_stores_response(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    item_id = view["items"][0]["item_id"]
    save_result = att.save_response(db_path, attempt_id=result["attempt_id"],
                                    attempt_item_id=item_id,
                                    selected_option_ids=["B"], client_revision=1)
    assert save_result["accepted"] is True
    assert save_result["stored_revision"] == 1


def test_old_revision_does_not_overwrite_new(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    item_id = view["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["B"], client_revision=5)
    old = att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item_id,
                            selected_option_ids=["A"], client_revision=2)
    assert old["accepted"] is False
    assert old["stored_revision"] == 5


# --- M3.5: submit + scoring ---

def test_submit_creates_result(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    submit = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1",
                                attempt_id=result["attempt_id"])
    assert submit["status"] == "submitted"
    assert submit["result"]["maximum_score"] == 3.0


def test_submit_is_idempotent(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    s1 = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    s2 = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    assert s1["result"]["score"] == s2["result"]["score"]


def test_scoring_from_snapshot(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    # Q1,Q2 correct=B, Q3 correct=A. Answer all B → 2 correct.
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    submit = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    assert submit["result"]["score"] == 2.0
    assert submit["result"]["percent"] == pytest.approx(66.67, abs=0.1)
    assert submit["result"]["passed"] is False


def test_submit_after_deadline_scores_saved_only(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    # simulate deadline passed
    conn = training_db.write_connection(db_path)
    conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?",
                 (time_policy.utc_now_ms() - 1000, result["attempt_id"]))
    conn.commit()
    conn.close()
    submit = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    assert submit["status"] == "timed_out"


def test_restart_does_not_lose_responses(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    item_id = view["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["B"], client_revision=3)
    # simulate restart: reload view
    view2 = att.get_attempt_learner_view(db_path, result["attempt_id"])
    assert view2["items"][0]["response"]["client_revision"] == 3


def test_burst_submit_one_result(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)

    results = []
    barrier = threading.Barrier(5)

    def do_submit():
        barrier.wait()
        r = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
        results.append(r["result"]["score"])

    threads = [threading.Thread(target=do_submit) for _ in range(5)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=15)

    conn = training_db.read_connection(db_path)
    count = conn.execute(
        "SELECT COUNT(*) AS c FROM exam_results WHERE attempt_id=?", (result["attempt_id"],)
    ).fetchone()["c"]
    conn.close()
    assert count == 1
    assert len(results) == 5
