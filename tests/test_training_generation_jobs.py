import threading
import time
from argparse import Namespace

import pytest

from training import db as training_db
from training import migrations, time_policy
from training.errors import ErrorCode, TrainingError
from services import training_generation_service as gs
from training.providers import fake as fake_provider
from services.training_catalog_service import seed_defaults


VALID_BATCH = {
    "schema_version": "1.0",
    "batch": {
        "title": "Fake batch",
        "language": "vi",
        "source_document_version_ids": ["docver-001"],
        "target_audience_codes": ["nvkt"],
        "requested_count": 1,
    },
    "questions": [
        {
            "local_ref": "FQ001",
            "type": "single_choice",
            "stem": "Câu fake provider?",
            "options": [
                {"id": "A", "text": "Sai"},
                {"id": "B", "text": "Đúng"},
            ],
            "correct_option_ids": ["B"],
            "explanation": "Fake đúng.",
            "distractor_rationales": {"A": "Sai"},
            "classification": {"domain_code": "quality", "topic_codes": ["fake"]},
            "difficulty": "easy",
            "evidence": [
                {
                    "document_version_id": "docver-001",
                    "block_id": "DOC-B001",
                    "extraction_revision": 1,
                    "quoted_text": "fake quote",
                    "supports": "correct_answer",
                }
            ],
        }
    ],
}


def _setup(monkeypatch, tmp_path, unit_code="son_tay"):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", unit_code)
    migrations.run_migrations(db_path, unit_code)
    seed_defaults(db_path, unit_code)
    return db_path


def test_enqueue_job_creates_pending(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"],
        requested_count=5,
        idempotency_key="key-1",
    )
    job = gs.get_job(db_path, job_id)
    assert job["status"] == "pending"
    assert job["requested_count"] == 5


def test_idempotency_key_prevents_duplicate(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    j1 = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=5,
        idempotency_key="dup-key",
    )
    j2 = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=5,
        idempotency_key="dup-key",
    )
    assert j1 == j2


def test_claim_job_atomic(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=5,
    )
    claimed = gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)
    assert claimed is not None
    assert claimed["id"] == job_id
    assert claimed["status"] == "running"


def test_two_workers_dont_claim_same_job(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=5,
    )
    results = []
    barrier = threading.Barrier(2)

    def worker(wid):
        barrier.wait()
        claimed = gs.claim_next_job(db_path, worker_id=wid, lease_seconds=300)
        results.append(claimed)

    t1 = threading.Thread(target=worker, args=("w1",))
    t2 = threading.Thread(target=worker, args=("w2",))
    t1.start()
    t2.start()
    t1.join(timeout=10)
    t2.join(timeout=10)

    claimed = [r for r in results if r is not None]
    assert len(claimed) == 1


def test_expired_lease_can_be_reclaimed(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=5,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)
    # simulate expired lease
    conn = training_db.write_connection(db_path)
    conn.execute(
        "UPDATE ai_generation_jobs SET lease_expires_at_ms=? WHERE id=?",
        (time_policy.utc_now_ms() - 1000, job_id),
    )
    conn.commit()
    conn.close()

    reclaimed = gs.claim_next_job(db_path, worker_id="w2", lease_seconds=300)
    assert reclaimed is not None
    assert reclaimed["id"] == job_id
    assert reclaimed["claimed_by"] == "w2"
    assert reclaimed["retry_count"] == 1


def test_complete_job_creates_drafts(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)
    gs.complete_job(
        db_path, job_id=job_id, worker_id="w1",
        batch=VALID_BATCH,
        provider_metadata={"provider": "fake", "model": "test"},
    )
    job = gs.get_job(db_path, job_id)
    assert job["status"] == "completed"

    from services.training_question_service import list_questions
    result = list_questions(db_path)
    assert result["total"] == 1


def test_fail_job_increments_retry(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1, max_retries=2,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)
    gs.fail_job(db_path, job_id=job_id, worker_id="w1",
                error_code="TIMEOUT", error_detail="timed out")
    job = gs.get_job(db_path, job_id)
    assert job["retry_count"] == 1
    assert job["status"] == "pending"


def test_fail_job_exhausted_sets_failed(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1, max_retries=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)
    gs.fail_job(db_path, job_id=job_id, worker_id="w1",
                error_code="TIMEOUT", error_detail="timed out")
    job = gs.get_job(db_path, job_id)
    assert job["status"] == "failed"
    assert job["error_code"] == "TIMEOUT"


def test_fake_provider_returns_valid_batch(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    batch = fake_provider.generate(
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"],
        requested_count=1,
    )
    from services.training_question_service import validate_question_batch
    errors = validate_question_batch(batch)
    assert errors == []


def test_cancel_job(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.cancel_job(db_path, unit_code="son_tay", actor="alice", job_id=job_id)
    job = gs.get_job(db_path, job_id)
    assert job["status"] == "cancelled"


def test_stale_worker_cannot_complete_after_lease_expires(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=-1)

    with pytest.raises(TrainingError) as exc_info:
        gs.complete_job(
            db_path, job_id=job_id, worker_id="w1", batch=VALID_BATCH,
            provider_metadata={"provider": "fake", "model": "test"},
        )

    assert exc_info.value.code == ErrorCode.CONFLICT
    assert gs.get_job(db_path, job_id)["status"] == "running"


def test_stale_worker_cannot_fail_after_lease_expires(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=-1)

    with pytest.raises(TrainingError) as exc_info:
        gs.fail_job(db_path, job_id=job_id, worker_id="w1",
                    error_code="TIMEOUT", error_detail="timed out")

    assert exc_info.value.code == ErrorCode.CONFLICT
    job = gs.get_job(db_path, job_id)
    assert job["status"] == "running"
    assert job["retry_count"] == 0


def test_stale_worker_cannot_heartbeat_after_lease_expires(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=-1)

    with pytest.raises(TrainingError) as exc_info:
        gs.heartbeat(db_path, job_id=job_id, worker_id="w1", lease_seconds=300)

    assert exc_info.value.code == ErrorCode.CONFLICT


def test_cancelled_running_job_cannot_complete(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)
    gs.cancel_job(db_path, unit_code="son_tay", actor="alice", job_id=job_id)

    with pytest.raises(TrainingError) as exc_info:
        gs.complete_job(
            db_path, job_id=job_id, worker_id="w1", batch=VALID_BATCH,
            provider_metadata={"provider": "fake", "model": "test"},
        )

    assert exc_info.value.code == ErrorCode.CONFLICT
    assert gs.get_job(db_path, job_id)["status"] == "cancelled"


def test_draft_import_failure_rolls_back_batch_and_completion(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    job_id = gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)

    def fail_import(*args, **kwargs):
        raise RuntimeError("draft import failed")

    monkeypatch.setattr(gs, "import_question_batch", fail_import)
    with pytest.raises(RuntimeError, match="draft import failed"):
        gs.complete_job(
            db_path, job_id=job_id, worker_id="w1", batch=VALID_BATCH,
            provider_metadata={"provider": "fake", "model": "test"},
        )

    assert gs.get_job(db_path, job_id)["status"] == "running"
    conn = training_db.read_connection(db_path)
    try:
        assert conn.execute("SELECT COUNT(*) FROM ai_generation_batches").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM question_versions").fetchone()[0] == 0
    finally:
        conn.close()


def test_job_records_unit_code_and_complete_audit_uses_it(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path, unit_code="ba_vi")
    job_id = gs.enqueue_job(
        db_path, unit_code="ba_vi", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    gs.claim_next_job(db_path, worker_id="w1", lease_seconds=300)
    gs.complete_job(
        db_path, job_id=job_id, worker_id="w1", batch=VALID_BATCH,
        provider_metadata={"provider": "fake", "model": "test"},
    )

    assert gs.get_job(db_path, job_id)["unit_code"] == "ba_vi"
    conn = training_db.read_connection(db_path)
    try:
        audit = conn.execute(
            "SELECT unit_code FROM training_audit_log "
            "WHERE action='complete_job' AND entity_id=?", (job_id,),
        ).fetchone()
    finally:
        conn.close()
    assert audit["unit_code"] == "ba_vi"


def test_concurrent_expired_reclaim_counts_one_abandoned_attempt(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    gs.enqueue_job(
        db_path, unit_code="son_tay", actor="alice",
        source_document_version_ids=["docver-001"],
        target_audience_codes=["nvkt"], requested_count=1,
    )
    claimed = gs.claim_next_job(db_path, worker_id="w1", lease_seconds=-1)
    barrier = threading.Barrier(2)
    results = []

    def reclaim(worker_id):
        barrier.wait()
        results.append(gs.claim_next_job(db_path, worker_id=worker_id, lease_seconds=300))

    threads = [threading.Thread(target=reclaim, args=(worker_id,)) for worker_id in ("w2", "w3")]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=10)

    successful = [job for job in results if job is not None]
    assert len(successful) == 1
    assert successful[0]["id"] == claimed["id"]
    assert successful[0]["retry_count"] == 1


def test_worker_without_once_keeps_polling_after_empty_queue(monkeypatch, tmp_path):
    from training import cli

    polls = []

    def claim(*args, **kwargs):
        polls.append(None)
        if len(polls) == 1:
            return None
        raise KeyboardInterrupt

    monkeypatch.setattr(gs, "claim_next_job", claim)
    monkeypatch.setattr("time.sleep", lambda interval: None)
    monkeypatch.setattr(cli, "_get_provider", lambda provider: object())
    args = Namespace(db_path=str(tmp_path / "training.db"), provider="fake", worker_id="w1",
                     once=False, poll_interval=0)

    with pytest.raises(KeyboardInterrupt):
        cli.cmd_worker(args)

    assert len(polls) == 2
