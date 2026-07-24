import copy

import pytest

from training import db as training_db
from training import migrations
from services import training_question_service as qs
from services.training_question_service import (
    import_question_batch,
    list_questions,
    get_question_management_detail,
    get_question_version,
    publish_question_version,
    update_question_draft,
    add_review_action,
)
from services.training_catalog_service import seed_defaults
from services.training_knowledge_service import create_document
from training.errors import TrainingError


VALID_BATCH = {
    "schema_version": "1.0",
    "batch": {
        "title": "Test batch",
        "language": "vi",
        "source_document_version_ids": ["docver-001"],
        "target_audience_codes": ["nvkt"],
        "requested_count": 1,
    },
    "questions": [
        {
            "local_ref": "Q001",
            "type": "single_choice",
            "stem": "Câu hỏi test?",
            "options": [
                {"id": "A", "text": "Sai"},
                {"id": "B", "text": "Đúng"},
            ],
            "correct_option_ids": ["B"],
            "explanation": "Giải thích",
            "distractor_rationales": {"A": "Sai"},
            "classification": {
                "domain_code": "quality",
                "topic_codes": ["test_topic"],
            },
            "difficulty": "easy",
            "evidence": [
                {
                    "document_version_id": "docver-001",
                    "block_id": "DOC-B001",
                    "extraction_revision": 1,
                    "quoted_text": "quote",
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
    qs.reset_duplicate_cache()
    return db_path


def test_import_batch_creates_draft_versions(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                   batch=VALID_BATCH, status="draft")
    assert len(result["version_ids"]) == 1
    ver = get_question_version(db_path, result["version_ids"][0])
    assert ver["review_status"] == "draft"
    assert ver["publication_status"] == "unpublished"
    assert ver["stem"] == "Câu hỏi test?"


def test_import_creates_options(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                   batch=VALID_BATCH, status="draft")
    conn = training_db.read_connection(db_path)
    opts = [dict(r) for r in conn.execute(
        "SELECT option_code, option_text, display_order FROM question_options "
        "WHERE question_version_id=? ORDER BY display_order",
        (result["version_ids"][0],),
    ).fetchall()]
    conn.close()
    assert len(opts) == 2
    assert opts[0]["option_code"] == "A"
    assert opts[1]["option_code"] == "B"


def test_publish_sets_published(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                   batch=VALID_BATCH, status="draft")
    vid = result["version_ids"][0]
    add_review_action(db_path, unit_code="son_tay", actor="bob",
                      version_id=vid, action="approve")
    publish_question_version(db_path, unit_code="son_tay", actor="bob", version_id=vid)
    ver = get_question_version(db_path, vid)
    assert ver["publication_status"] == "published"
    assert ver["review_status"] == "approved"


def test_publish_blocked_when_not_approved(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                   batch=VALID_BATCH, status="draft")
    vid = result["version_ids"][0]
    with pytest.raises(TrainingError) as exc:
        publish_question_version(db_path, unit_code="son_tay", actor="bob", version_id=vid)
    assert "approved" in exc.value.message.lower() or "review" in exc.value.message.lower()


def test_publish_blocked_by_duplicate_normalized_stem(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    r1 = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                               batch=VALID_BATCH, status="draft")
    add_review_action(db_path, unit_code="son_tay", actor="bob",
                      version_id=r1["version_ids"][0], action="approve")
    publish_question_version(db_path, unit_code="son_tay", actor="bob",
                             version_id=r1["version_ids"][0])

    batch2 = dict(VALID_BATCH)
    batch2["questions"] = [dict(VALID_BATCH["questions"][0])]
    batch2["questions"][0]["local_ref"] = "Q002"
    batch2["questions"][0]["stem"] = "CÂU HỎI test?"  # normalized dup
    r2 = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                               batch=batch2, status="draft")
    add_review_action(db_path, unit_code="son_tay", actor="bob",
                      version_id=r2["version_ids"][0], action="approve")
    with pytest.raises(TrainingError) as exc:
        publish_question_version(db_path, unit_code="son_tay", actor="bob",
                                 version_id=r2["version_ids"][0])
    assert "duplicate" in exc.value.message.lower()


def test_list_questions_pagination(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    batch = dict(VALID_BATCH)
    questions = []
    for i in range(3):
        q = dict(VALID_BATCH["questions"][0])
        q["local_ref"] = f"Q{i:03d}"
        q["stem"] = f"Câu hỏi số {i}?"
        questions.append(q)
    batch["questions"] = questions
    import_question_batch(db_path, unit_code="son_tay", actor="alice",
                          batch=batch, status="draft")
    result = list_questions(db_path, page=1, page_size=2)
    assert result["total"] == 3
    assert len(result["items"]) == 2
    result2 = list_questions(db_path, page=2, page_size=2)
    assert len(result2["items"]) == 1


def test_question_bank_list_derives_classification_and_filters(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    document = create_document(
        db_path, unit_code="son_tay", actor="alice", document_code="DOC",
        title="Tài liệu", content_text="Nội dung", classification={"domain_code": "quality"},
    )
    batch = copy.deepcopy(VALID_BATCH)
    batch["batch"]["source_document_version_ids"] = [document["version_id"]]
    batch["questions"][0]["classification"]["audience_codes"] = ["nvkt"]
    batch["questions"][0]["evidence"][0]["document_version_id"] = document["version_id"]
    imported = import_question_batch(db_path, unit_code="son_tay", actor="alice", batch=batch)

    result = list_questions(
        db_path, status="draft", audience="nvkt", domain="quality", topic="test_topic",
        q="hỏi test", page=1, page_size=1,
    )

    assert result["total"] == 1
    assert result["items"] == [{
        "id": imported["version_ids"][0], "stem": "Câu hỏi test?", "type": "single_choice",
        "difficulty": "easy", "audience": ["nvkt"], "topic": ["test_topic"],
        "status": "draft", "version": 1,
    }]


def test_question_bank_detail_is_explicit_management_dto(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    imported = import_question_batch(db_path, unit_code="son_tay", actor="alice", batch=VALID_BATCH)

    detail = get_question_management_detail(db_path, imported["version_ids"][0])

    assert detail["correct_option_ids"] == ["B"]
    assert detail["distractor_rationales"] == {"A": "Sai"}
    assert detail["classification"] == {
        "domain_codes": [], "topic_codes": ["test_topic"], "audience_codes": [], "indicator_codes": [],
    }
    assert detail["evidence"] == [{
        "document_version_id": "docver-001", "block_id": "DOC-B001", "extraction_revision": 1,
        "quoted_text": "quote", "quote_start": None, "quote_end": None, "supports": "correct_answer",
    }]
    assert detail["review_history"] == []
    assert detail["publication"] == {"status": "unpublished", "approved_by": None, "approved_at_ms": None}


def test_question_bank_list_keeps_rejected_state_distinct_from_draft(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    imported = import_question_batch(db_path, unit_code="son_tay", actor="alice", batch=VALID_BATCH)
    add_review_action(
        db_path, unit_code="son_tay", actor="manager", version_id=imported["version_ids"][0],
        action="reject", comment="Chưa đạt",
    )

    result = list_questions(db_path)

    assert result["items"][0]["status"] == "rejected"
    assert list_questions(db_path, status="draft")["items"] == []


def test_update_draft_with_expected_version(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                   batch=VALID_BATCH, status="draft")
    vid = result["version_ids"][0]
    update_question_draft(db_path, unit_code="son_tay", actor="alice",
                          version_id=vid, expected_version=1,
                          updates={"stem": "Câu đã sửa?"})
    ver = get_question_version(db_path, vid)
    assert ver["stem"] == "Câu đã sửa?"


def test_update_draft_version_conflict(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                   batch=VALID_BATCH, status="draft")
    vid = result["version_ids"][0]
    with pytest.raises(TrainingError) as exc:
        update_question_draft(db_path, unit_code="son_tay", actor="alice",
                              version_id=vid, expected_version=99,
                              updates={"stem": "sai version"})
    assert exc.value.code == "VERSION_CONFLICT"


def test_published_version_immutable(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    result = import_question_batch(db_path, unit_code="son_tay", actor="alice",
                                   batch=VALID_BATCH, status="draft")
    vid = result["version_ids"][0]
    add_review_action(db_path, unit_code="son_tay", actor="bob",
                      version_id=vid, action="approve")
    publish_question_version(db_path, unit_code="son_tay", actor="bob", version_id=vid)
    with pytest.raises(TrainingError):
        update_question_draft(db_path, unit_code="son_tay", actor="alice",
                              version_id=vid, expected_version=1,
                              updates={"stem": "đổi sau publish"})
