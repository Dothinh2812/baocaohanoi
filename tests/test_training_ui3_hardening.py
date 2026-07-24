import threading

import pytest

from tests.test_training_exam_states import VALID_BATCH, _publish_questions, _setup
from tests.test_training_routes import QUESTION_BATCH, _client
from training import db as training_db
from training.errors import TrainingError
from services import training_question_service as qs


def test_question_bank_detail_has_no_store_cache_header(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        imported = client.post(
            "/api/training/questions/import", headers={"X-CSRF-Token": "csrf"},
            json=QUESTION_BATCH,
        )
        version_id = imported.get_json()["version_ids"][0]
        response = client.get(f"/api/training/questions/{version_id}")

    assert response.status_code == 200
    assert response.headers["Cache-Control"] == "no-store"


def _import_draft(db_path, *, version_index=0):
    result = qs.import_question_batch(
        db_path, unit_code="son_tay", actor="alice", batch=VALID_BATCH, status="draft",
    )
    return result["version_ids"][version_index]


def _import_approved_unpublished(db_path, *, version_index=0):
    version_id = _import_draft(db_path, version_index=version_index)
    qs.add_review_action(
        db_path, unit_code="son_tay", actor="bob",
        version_id=version_id, action="approve",
    )
    return version_id


def test_concurrent_approve_writes_one_audit_and_one_review(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_id = _import_draft(db_path)
    barrier = threading.Barrier(2)
    successes = []
    errors = []

    def approve():
        barrier.wait()
        try:
            qs.add_review_action(
                db_path, unit_code="son_tay", actor="bob",
                version_id=version_id, action="approve",
            )
            successes.append(1)
        except TrainingError as exc:
            errors.append(exc)

    threads = [threading.Thread(target=approve) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=15)

    assert len(successes) == 1
    assert len(errors) == 1
    assert errors[0].code == "CONFLICT"
    assert errors[0].status == 409

    conn = training_db.read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT review_status FROM question_versions WHERE id=?", (version_id,),
        ).fetchone()
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log "
            "WHERE action='review_approve' AND entity_id=?", (version_id,),
        ).fetchone()["c"]
        review_count = conn.execute(
            "SELECT COUNT(*) AS c FROM question_reviews "
            "WHERE question_version_id=? AND action='approve'", (version_id,),
        ).fetchone()["c"]
    finally:
        conn.close()
    assert row["review_status"] == "approved"
    assert audit_count == 1
    assert review_count == 1


def test_failed_review_transition_writes_no_audit_or_review(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_id = _import_draft(db_path)
    qs.add_review_action(
        db_path, unit_code="son_tay", actor="bob",
        version_id=version_id, action="approve",
    )

    with pytest.raises(TrainingError) as exc_info:
        qs.add_review_action(
            db_path, unit_code="son_tay", actor="bob",
            version_id=version_id, action="approve",
        )
    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409

    conn = training_db.read_connection(db_path)
    try:
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log "
            "WHERE action='review_approve' AND entity_id=?", (version_id,),
        ).fetchone()["c"]
        review_count = conn.execute(
            "SELECT COUNT(*) AS c FROM question_reviews "
            "WHERE question_version_id=? AND action='approve'", (version_id,),
        ).fetchone()["c"]
    finally:
        conn.close()
    assert audit_count == 1
    assert review_count == 1


def test_concurrent_publish_writes_one_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_id = _import_approved_unpublished(db_path)
    barrier = threading.Barrier(2)
    successes = []
    errors = []

    def publish():
        barrier.wait()
        try:
            qs.publish_question_version(
                db_path, unit_code="son_tay", actor="bob", version_id=version_id,
            )
            successes.append(1)
        except TrainingError as exc:
            errors.append(exc)

    threads = [threading.Thread(target=publish) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=15)

    assert len(successes) == 1
    assert len(errors) == 1
    assert errors[0].code == "CONFLICT"
    assert errors[0].status == 409

    conn = training_db.read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT publication_status FROM question_versions WHERE id=?", (version_id,),
        ).fetchone()
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log "
            "WHERE action='publish' AND entity_id=?", (version_id,),
        ).fetchone()["c"]
    finally:
        conn.close()
    assert row["publication_status"] == "published"
    assert audit_count == 1


def test_failed_publish_writes_no_additional_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    version_id = _import_approved_unpublished(db_path)
    qs.publish_question_version(
        db_path, unit_code="son_tay", actor="bob", version_id=version_id,
    )

    with pytest.raises(TrainingError) as exc_info:
        qs.publish_question_version(
            db_path, unit_code="son_tay", actor="bob", version_id=version_id,
        )
    assert exc_info.value.code == "CONFLICT"
    assert exc_info.value.status == 409

    conn = training_db.read_connection(db_path)
    try:
        audit_count = conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log "
            "WHERE action='publish' AND entity_id=?", (version_id,),
        ).fetchone()["c"]
    finally:
        conn.close()
    assert audit_count == 1
