import hashlib
import json

import pytest

from training import db as training_db
from training import migrations
from services import training_knowledge_service as ks
from services.training_catalog_service import seed_defaults


def _setup(monkeypatch, tmp_path, unit_code="son_tay"):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", unit_code)
    migrations.run_migrations(db_path, unit_code)
    seed_defaults(db_path, unit_code)
    return db_path


SAMPLE_TEXT = """# C1.1 Chất lượng sửa chữa

Đối với các thuê bao ngoài CCCO báo hỏng sau 22h, thời gian sửa chữa
được tính từ 8h sáng hôm sau. Nếu hoàn thành trước 8h thì tính bằng 0.

Tiêu chí đạt: tỷ lệ sửa đúng hẹn >= 95%.
"""


def test_create_document_from_paste_text(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Chất lượng sửa chữa",
        content_text=SAMPLE_TEXT,
        classification={"domain_code": "quality"},
        audience_codes=["nvkt"],
    )
    assert doc["document_id"]
    assert doc["version_id"]
    assert doc["version_number"] == 1


def test_content_sha256_stored(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="hello world",
        classification={}, audience_codes=[],
    )
    conn = training_db.read_connection(db_path)
    row = conn.execute(
        "SELECT content_sha256 FROM knowledge_document_versions WHERE id=?",
        (doc["version_id"],),
    ).fetchone()
    conn.close()
    expected = hashlib.sha256("hello world".encode("utf-8")).hexdigest()
    assert row["content_sha256"] == expected


def test_blocks_have_stable_ids(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="Para one.\n\nPara two.\n\nPara three.",
        classification={}, audience_codes=[],
    )
    conn = training_db.read_connection(db_path)
    blocks = [dict(r) for r in conn.execute(
        "SELECT block_id, content FROM knowledge_blocks "
        "WHERE document_version_id=? ORDER BY block_id",
        (doc["version_id"],),
    ).fetchall()]
    conn.close()
    assert len(blocks) >= 2
    for b in blocks:
        assert b["block_id"]
        assert b["content"]


def test_extraction_revision_increments_on_reprocess(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="Para one.\n\nPara two.",
        classification={}, audience_codes=[],
    )
    rev1 = ks.get_extraction_revision(db_path, doc["version_id"])
    ks.reprocess_blocks(db_path, unit_code="son_tay", actor="alice",
                        version_id=doc["version_id"], content_text="Para one.\n\nPara two.\n\nPara three.")
    rev2 = ks.get_extraction_revision(db_path, doc["version_id"])
    assert rev2 == rev1 + 1


def test_create_new_version(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="v1 content",
        classification={}, audience_codes=[],
    )
    v2 = ks.create_version(
        db_path, unit_code="son_tay", actor="alice",
        document_id=doc["document_id"],
        content_text="v2 content",
        classification={}, audience_codes=[],
    )
    assert v2["version_number"] == 2


def test_old_version_readable_after_new_version(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="v1 content",
        classification={}, audience_codes=[],
    )
    old_version_id = doc["version_id"]
    ks.create_version(
        db_path, unit_code="son_tay", actor="alice",
        document_id=doc["document_id"],
        content_text="v2 content",
        classification={}, audience_codes=[],
    )
    old = ks.get_version(db_path, old_version_id)
    assert old["content_text"] == "v1 content"


def test_create_and_list_issues(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="Some ambiguous text.",
        classification={}, audience_codes=[],
    )
    issue_id = ks.create_issue(
        db_path, unit_code="son_tay", actor="alice",
        version_id=doc["version_id"], block_id=None,
        severity="high", issue_type="ambiguous",
        description="Câu có nhiều cách hiểu",
    )
    issues = ks.list_issues(db_path, doc["version_id"])
    assert len(issues) == 1
    assert issues[0]["id"] == issue_id
    assert issues[0]["status"] == "open"


def test_resolve_issue(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="text",
        classification={}, audience_codes=[],
    )
    issue_id = ks.create_issue(
        db_path, unit_code="son_tay", actor="alice",
        version_id=doc["version_id"], severity="medium",
        issue_type="missing_effective_date", description="No date",
    )
    ks.resolve_issue(db_path, unit_code="son_tay", actor="bob",
                     issue_id=issue_id, status="excluded")
    issues = ks.list_issues(db_path, doc["version_id"])
    assert issues[0]["status"] == "excluded"
    assert issues[0]["resolved_by"] == "bob"


def test_blocking_high_severity_open_issue(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="text",
        classification={}, audience_codes=[],
    )
    ks.create_issue(
        db_path, unit_code="son_tay", actor="alice",
        version_id=doc["version_id"], severity="high",
        issue_type="conflict", description="conflicting rules",
    )
    assert ks.has_blocking_issues(db_path, doc["version_id"]) is True


def test_no_blocking_issue_when_resolved(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    doc = ks.create_document(
        db_path, unit_code="son_tay", actor="alice",
        document_code="C1.1", title="Test",
        content_text="text",
        classification={}, audience_codes=[],
    )
    issue_id = ks.create_issue(
        db_path, unit_code="son_tay", actor="alice",
        version_id=doc["version_id"], severity="high",
        issue_type="conflict", description="conflict",
    )
    ks.resolve_issue(db_path, unit_code="son_tay", actor="alice",
                     issue_id=issue_id, status="confirmed")
    assert ks.has_blocking_issues(db_path, doc["version_id"]) is False
