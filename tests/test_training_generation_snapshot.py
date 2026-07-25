"""Tests cho generation snapshot: blocks resolution + evidence selection guard."""
import pytest

from training import db as training_db
from training import migrations
from services.training_catalog_service import seed_defaults
from services.training_knowledge_service import (
    create_document, resolve_generation_snapshot, find_version_by_checksum,
)
from services.training_generation_service import list_jobs, get_job_detail


def _setup(monkeypatch, tmp_path, unit_code="son_tay"):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", unit_code)
    migrations.run_migrations(db_path, unit_code)
    seed_defaults(db_path, unit_code)
    return db_path


def _create_doc(db_path, content="Đoạn 1.\n\nĐoạn 2 nội dung."):
    return create_document(
        db_path, unit_code="son_tay", actor="admin",
        document_code="testdoc", title="Test",
        content_text=content,
        classification={"domain_code": "quality", "topic_codes": ["t1"]},
        audience_codes=["nvkt"],
    )


class TestResolveGenerationSnapshot:
    def test_snapshot_includes_blocks(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        snapshot = resolve_generation_snapshot(db_path, [doc["version_id"]], ["nvkt"])
        assert len(snapshot["document_versions"]) == 1
        blocks = snapshot["document_versions"][0]["blocks"]
        assert len(blocks) >= 2
        for blk in blocks:
            assert "block_id" in blk
            assert "content" in blk
            assert "content_sha256" in blk
            assert "char_start" in blk
            assert "char_end" in blk

    def test_allowed_block_ids(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        snapshot = resolve_generation_snapshot(db_path, [doc["version_id"]], ["nvkt"])
        assert doc["version_id"] in snapshot["allowed_document_version_ids"]
        for blk in snapshot["document_versions"][0]["blocks"]:
            assert blk["block_id"] in snapshot["allowed_block_ids"]

    def test_nonexistent_version_skipped(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        snapshot = resolve_generation_snapshot(db_path, ["fake-id"], ["nvkt"])
        assert len(snapshot["document_versions"]) == 0
        assert len(snapshot["allowed_block_ids"]) == 0

    def test_includes_audience_and_topics(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        snapshot = resolve_generation_snapshot(db_path, [doc["version_id"]], ["nvkt"])
        dv = snapshot["document_versions"][0]
        assert "nvkt" in dv["audience_codes"]
        assert "t1" in dv["topic_codes"]


class TestFindVersionByChecksum:
    def test_finds_existing(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path, content="Unique content here.")
        from services.training_knowledge_service import get_version
        ver = get_version(db_path, doc["version_id"])
        found = find_version_by_checksum(db_path, ver["content_sha256"])
        assert found is not None
        assert found["id"] == doc["version_id"]

    def test_returns_none_for_new(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        assert find_version_by_checksum(db_path, "nonexistent-hash") is None


class TestListJobs:
    def test_list_empty(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        result = list_jobs(db_path)
        assert result["total"] == 0
        assert result["items"] == []

    def test_list_with_jobs(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        from services.training_generation_service import enqueue_job
        enqueue_job(db_path, unit_code="son_tay", actor="alice",
                    source_document_version_ids=[doc["version_id"]],
                    target_audience_codes=["nvkt"], requested_count=5)
        result = list_jobs(db_path)
        assert result["total"] == 1
        assert result["items"][0]["created_by"] == "alice"

    def test_list_filter_by_status(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        from services.training_generation_service import enqueue_job
        enqueue_job(db_path, unit_code="son_tay", actor="alice",
                    source_document_version_ids=[doc["version_id"]],
                    target_audience_codes=["nvkt"], requested_count=5)
        result = list_jobs(db_path, status="pending")
        assert result["total"] == 1
        result = list_jobs(db_path, status="completed")
        assert result["total"] == 0

    def test_list_pagination(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        from services.training_generation_service import enqueue_job
        for i in range(5):
            enqueue_job(db_path, unit_code="son_tay", actor="alice",
                        source_document_version_ids=[doc["version_id"]],
                        target_audience_codes=["nvkt"], requested_count=1,
                        idempotency_key=f"k{i}")
        result = list_jobs(db_path, page=1, page_size=2)
        assert len(result["items"]) == 2
        assert result["total"] == 5

    def test_no_sensitive_columns(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        from services.training_generation_service import enqueue_job
        enqueue_job(db_path, unit_code="son_tay", actor="alice",
                    source_document_version_ids=[doc["version_id"]],
                    target_audience_codes=["nvkt"], requested_count=5)
        result = list_jobs(db_path)
        item = result["items"][0]
        assert "request_payload_json" not in item
        assert "source_document_version_ids_json" not in item


class TestGetJobDetail:
    def test_includes_batch_meta(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        doc = _create_doc(db_path)
        from services.training_generation_service import (
            enqueue_job, claim_next_job, complete_job,
        )
        from training.providers.fake import FakeProvider
        job_id = enqueue_job(db_path, unit_code="son_tay", actor="alice",
                             source_document_version_ids=[doc["version_id"]],
                             target_audience_codes=["nvkt"], requested_count=1)
        job = claim_next_job(db_path, worker_id="w1", lease_seconds=60)
        provider = FakeProvider()
        result = provider.generate(
            source_document_version_ids=[doc["version_id"]],
            target_audience_codes=["nvkt"],
            requested_count=1,
        )
        complete_job(db_path, job_id=job_id, worker_id="w1",
                     batch=result.batch,
                     provider_metadata={
                         "provider": "fake", "model": "fake-model",
                         "prompt_version": "1.0", "usage": {"tokens": 0},
                         "raw_response": None,
                     })
        detail = get_job_detail(db_path, job_id)
        assert detail is not None
        assert detail["batch_meta"] is not None
        assert detail["batch_meta"]["provider"] == "fake"

    def test_returns_none_for_missing(self, monkeypatch, tmp_path):
        db_path = _setup(monkeypatch, tmp_path)
        assert get_job_detail(db_path, "nonexistent") is None
