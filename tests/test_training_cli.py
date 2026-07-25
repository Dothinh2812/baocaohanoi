"""Tests cho CLI commands: service delegation, RBAC, no direct SQL.

Mỗi test dùng temp training.db qua monkeypatch. CLI mutation kiểm tra role.
"""
import os
from argparse import Namespace

import pytest

from training import db as training_db
from training import migrations
from services.training_catalog_service import seed_defaults, grant_role


def _setup_db(monkeypatch, tmp_path, unit_code="son_tay"):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", unit_code)
    migrations.run_migrations(db_path, unit_code)
    seed_defaults(db_path, unit_code)
    grant_role(db_path, unit_code, "system", "admin", "admin")
    return db_path


def _import_doc(db_path, actor="admin", title="Test Doc", domain="quality",
                audiences="nvkt", topics=None):
    from services.training_knowledge_service import create_document
    classification = {"domain_code": domain}
    if topics:
        classification["topic_codes"] = [t.strip() for t in topics.split(",")]
    return create_document(
        db_path, unit_code="son_tay", actor=actor,
        document_code=title.lower().replace(" ", "_"),
        title=title, content_text="Đoạn 1.\n\nĐoạn 2 nội dung test.",
        classification=classification,
        audience_codes=[a.strip() for a in audiences.split(",")] if audiences else [],
    )


def _import_question_draft(db_path, doc_ver_id="docver-001", actor="admin"):
    from services.training_question_service import import_question_batch
    batch = {
        "schema_version": "1.0",
        "batch": {
            "title": "Test", "language": "vi",
            "source_document_version_ids": [doc_ver_id],
            "target_audience_codes": ["nvkt"],
            "requested_count": 1,
        },
        "questions": [{
            "local_ref": "Q1", "type": "single_choice",
            "stem": "Câu hỏi test?", "stimulus": None,
            "options": [{"id": "A", "text": "Sai"}, {"id": "B", "text": "Đúng"}],
            "correct_option_ids": ["B"],
            "explanation": "Đúng.", "distractor_rationales": {"A": "Sai"},
            "classification": {"domain_code": "quality", "topic_codes": ["t1"],
                               "audience_codes": ["nvkt"]},
            "difficulty": "easy",
            "evidence": [{"document_version_id": doc_ver_id, "block_id": "DOC-B001",
                          "extraction_revision": 1, "quoted_text": "quote",
                          "supports": "correct_answer"}],
        }],
    }
    result = import_question_batch(db_path, unit_code="son_tay", actor=actor, batch=batch)
    return result["version_ids"][0]


class TestCliParser:
    def test_help_lists_all_commands(self):
        from training.cli import build_parser
        parser = build_parser()
        subparser_action = next(
            a for a in parser._actions if hasattr(a, "choices") and a.dest == "command"
        )
        commands = set(subparser_action.choices.keys())
        for cmd in ("db-migrate", "worker", "knowledge-import", "knowledge-list",
                     "knowledge-show", "knowledge-issues", "generate-create",
                     "generate-list", "generate-show", "generate-cancel",
                     "questions-list", "questions-show", "questions-approve",
                     "questions-reject", "questions-publish"):
            assert cmd in commands, f"Missing command: {cmd}"


class TestKnowledgeImportCli:
    def test_import_txt_file(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        txt_path = str(tmp_path / "test.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("Nội dung tài liệu test.\n\nĐoạn thứ hai.")
        from training.cli import cmd_knowledge_import
        args = Namespace(
            db_path=db_path, unit_code="son_tay", file=txt_path,
            paste_file=None, paste_text=None, title="Test TXT",
            document_code=None, domain="quality", topics=None,
            audiences="nvkt", document_type="kpi_definition",
            issuer=None, actor="admin", force=False,
        )
        assert cmd_knowledge_import(args) == 0

    def test_import_paste_text(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        from training.cli import cmd_knowledge_import
        args = Namespace(
            db_path=db_path, unit_code="son_tay", file=None,
            paste_file=None, paste_text="Nội dung paste trực tiếp.",
            title="Paste Doc", document_code=None, domain="quality",
            topics=None, audiences="nvkt", document_type="kpi_definition",
            issuer=None, actor="admin", force=False,
        )
        assert cmd_knowledge_import(args) == 0

    def test_import_requires_editor_role(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        from training.cli import cmd_knowledge_import
        from training.errors import TrainingError
        args = Namespace(
            db_path=db_path, unit_code="son_tay", file=None,
            paste_file=None, paste_text="Test content.",
            title="Doc", document_code=None, domain="quality",
            topics=None, audiences="nvkt", document_type="kpi_definition",
            issuer=None, actor="nobody", force=False,
        )
        with pytest.raises(TrainingError) as exc_info:
            cmd_knowledge_import(args)
        assert exc_info.value.code == "PERMISSION_SCOPE_DENIED"

    def test_import_idempotent_checksum(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        content = "Nội dung trùng lặp."
        from training.cli import cmd_knowledge_import
        args1 = Namespace(
            db_path=db_path, unit_code="son_tay", file=None,
            paste_file=None, paste_text=content,
            title="Doc 1", document_code=None, domain="quality",
            topics=None, audiences=None, document_type="kpi_definition",
            issuer=None, actor="admin", force=False,
        )
        assert cmd_knowledge_import(args1) == 0
        args2 = Namespace(
            db_path=db_path, unit_code="son_tay", file=None,
            paste_file=None, paste_text=content,
            title="Doc 2", document_code=None, domain="quality",
            topics=None, audiences=None, document_type="kpi_definition",
            issuer=None, actor="admin", force=False,
        )
        with pytest.raises(ValueError, match="trùng checksum"):
            cmd_knowledge_import(args2)

    def test_import_unsupported_file_type(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        bad_path = str(tmp_path / "test.pdf")
        with open(bad_path, "wb") as f:
            f.write(b"%PDF-1.4 fake")
        from training.cli import cmd_knowledge_import
        from services.training_file_ingestion import FileValidationError
        args = Namespace(
            db_path=db_path, unit_code="son_tay", file=bad_path,
            paste_file=None, paste_text=None, title="Bad",
            document_code=None, domain="quality", topics=None,
            audiences=None, document_type="kpi_definition",
            issuer=None, actor="admin", force=False,
        )
        with pytest.raises(FileValidationError):
            cmd_knowledge_import(args)

    def test_import_mime_mismatch_txt_with_zip_sig(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        bad_path = str(tmp_path / "fake.txt")
        with open(bad_path, "wb") as f:
            f.write(b"PK\x03\x04 fake zip content")
        from training.cli import cmd_knowledge_import
        from services.training_file_ingestion import FileValidationError
        args = Namespace(
            db_path=db_path, unit_code="son_tay", file=bad_path,
            paste_file=None, paste_text=None, title="Fake",
            document_code=None, domain="quality", topics=None,
            audiences=None, document_type="kpi_definition",
            issuer=None, actor="admin", force=False,
        )
        with pytest.raises(FileValidationError, match="MIME|sai đuôi"):
            cmd_knowledge_import(args)


class TestKnowledgeListShowIssuesCli:
    def test_list(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        _import_doc(db_path, title="Listed Doc")
        from training.cli import cmd_knowledge_list
        args = Namespace(db_path=db_path, unit_code="son_tay", domain=None,
                         status=None, page=1, page_size=25)
        assert cmd_knowledge_list(args) == 0
        out = capsys.readouterr().out
        assert "Total: 1" in out

    def test_show(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        result = _import_doc(db_path)
        from training.cli import cmd_knowledge_show
        args = Namespace(db_path=db_path, unit_code="son_tay",
                         document_version_id=result["version_id"])
        assert cmd_knowledge_show(args) == 0
        out = capsys.readouterr().out
        assert "block_count" in out

    def test_show_not_found(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        from training.cli import cmd_knowledge_show
        args = Namespace(db_path=db_path, unit_code="son_tay",
                         document_version_id="nonexistent")
        assert cmd_knowledge_show(args) == 1

    def test_issues(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        result = _import_doc(db_path)
        from training.cli import cmd_knowledge_issues
        args = Namespace(db_path=db_path, unit_code="son_tay",
                         document_version_id=result["version_id"])
        assert cmd_knowledge_issues(args) == 0
        out = capsys.readouterr().out
        assert "has_blocking_issues" in out


class TestGenerateCli:
    def test_create_job(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from training.cli import cmd_generate_create
        args = Namespace(
            db_path=db_path, unit_code="son_tay",
            document_version_ids=doc["version_id"], audiences="nvkt",
            count=5, idempotency_key=None, actor="admin",
        )
        assert cmd_generate_create(args) == 0
        out = capsys.readouterr().out
        assert "job_id" in out

    def test_create_blocks_on_blocking_issue(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from services.training_knowledge_service import create_issue
        create_issue(db_path, unit_code="son_tay", actor="admin",
                     version_id=doc["version_id"], severity="high",
                     issue_type="test", description="blocking")
        from training.cli import cmd_generate_create
        args = Namespace(
            db_path=db_path, unit_code="son_tay",
            document_version_ids=doc["version_id"], audiences="nvkt",
            count=5, idempotency_key=None, actor="admin",
        )
        with pytest.raises(ValueError, match="blocking"):
            cmd_generate_create(args)

    def test_create_requires_editor_role(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from training.cli import cmd_generate_create
        from training.errors import TrainingError
        args = Namespace(
            db_path=db_path, unit_code="son_tay",
            document_version_ids=doc["version_id"], audiences="nvkt",
            count=5, idempotency_key=None, actor="nobody",
        )
        with pytest.raises(TrainingError):
            cmd_generate_create(args)

    def test_list_jobs(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from services.training_generation_service import enqueue_job
        enqueue_job(db_path, unit_code="son_tay", actor="admin",
                    source_document_version_ids=[doc["version_id"]],
                    target_audience_codes=["nvkt"], requested_count=3)
        from training.cli import cmd_generate_list
        args = Namespace(db_path=db_path, unit_code="son_tay", status=None,
                         actor=None, page=1, page_size=25)
        assert cmd_generate_list(args) == 0
        out = capsys.readouterr().out
        assert "Total: 1" in out

    def test_show_job(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from services.training_generation_service import enqueue_job
        job_id = enqueue_job(db_path, unit_code="son_tay", actor="admin",
                             source_document_version_ids=[doc["version_id"]],
                             target_audience_codes=["nvkt"], requested_count=3)
        from training.cli import cmd_generate_show
        args = Namespace(db_path=db_path, unit_code="son_tay", job_id=job_id)
        assert cmd_generate_show(args) == 0
        out = capsys.readouterr().out
        assert job_id in out

    def test_cancel_job(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from services.training_generation_service import enqueue_job
        job_id = enqueue_job(db_path, unit_code="son_tay", actor="admin",
                             source_document_version_ids=[doc["version_id"]],
                             target_audience_codes=["nvkt"], requested_count=3)
        from training.cli import cmd_generate_cancel
        args = Namespace(db_path=db_path, unit_code="son_tay", job_id=job_id,
                         actor="admin")
        assert cmd_generate_cancel(args) == 0

    def test_cancel_requires_exam_manager_role(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from services.training_generation_service import enqueue_job
        job_id = enqueue_job(db_path, unit_code="son_tay", actor="admin",
                             source_document_version_ids=[doc["version_id"]],
                             target_audience_codes=["nvkt"], requested_count=3)
        grant_role(db_path, "son_tay", "admin", "someone", "editor")
        from training.cli import cmd_generate_cancel
        from training.errors import TrainingError
        args = Namespace(db_path=db_path, unit_code="son_tay", job_id=job_id,
                         actor="someone")
        with pytest.raises(TrainingError):
            cmd_generate_cancel(args)


class TestQuestionsCli:
    def test_list(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        _import_question_draft(db_path, doc_ver_id=doc["version_id"])
        from training.cli import cmd_questions_list
        args = Namespace(db_path=db_path, unit_code="son_tay", status="draft",
                         topic=None, page=1, page_size=25)
        assert cmd_questions_list(args) == 0
        out = capsys.readouterr().out
        assert "Total: 1" in out

    def test_show(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        qv_id = _import_question_draft(db_path, doc_ver_id=doc["version_id"])
        from training.cli import cmd_questions_show
        args = Namespace(db_path=db_path, unit_code="son_tay", version_id=qv_id)
        assert cmd_questions_show(args) == 0

    def test_approve_via_service_cas(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        qv_id = _import_question_draft(db_path, doc_ver_id=doc["version_id"])
        from training.cli import cmd_questions_approve
        args = Namespace(db_path=db_path, unit_code="son_tay", version_id=qv_id,
                         actor="admin", comment=None)
        assert cmd_questions_approve(args) == 0

    def test_approve_requires_exam_manager_role(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        qv_id = _import_question_draft(db_path, doc_ver_id=doc["version_id"])
        grant_role(db_path, "son_tay", "admin", "ed", "editor")
        from training.cli import cmd_questions_approve
        from training.errors import TrainingError
        args = Namespace(db_path=db_path, unit_code="son_tay", version_id=qv_id,
                         actor="ed", comment=None)
        with pytest.raises(TrainingError):
            cmd_questions_approve(args)

    def test_publish_via_service_cas(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        qv_id = _import_question_draft(db_path, doc_ver_id=doc["version_id"])
        from training.cli import cmd_questions_approve, cmd_questions_publish
        ap = Namespace(db_path=db_path, unit_code="son_tay", version_id=qv_id,
                       actor="admin", comment=None)
        cmd_questions_approve(ap)
        pu = Namespace(db_path=db_path, unit_code="son_tay", version_id=qv_id,
                       actor="admin")
        assert cmd_questions_publish(pu) == 0

    def test_reject_requires_comment(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        qv_id = _import_question_draft(db_path, doc_ver_id=doc["version_id"])
        from training.cli import cmd_questions_reject
        args = Namespace(db_path=db_path, unit_code="son_tay", version_id=qv_id,
                         actor="admin", comment=None)
        with pytest.raises(ValueError, match="comment"):
            cmd_questions_reject(args)

    def test_publish_blocked_if_not_approved(self, monkeypatch, tmp_path):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        qv_id = _import_question_draft(db_path, doc_ver_id=doc["version_id"])
        from training.cli import cmd_questions_publish
        from training.errors import TrainingError
        args = Namespace(db_path=db_path, unit_code="son_tay", version_id=qv_id,
                         actor="admin")
        with pytest.raises(TrainingError) as exc_info:
            cmd_questions_publish(args)
        assert exc_info.value.code == "CONFLICT"


class TestWorkerSnapshot:
    def test_worker_fake_completes_job(self, monkeypatch, tmp_path, capsys):
        db_path = _setup_db(monkeypatch, tmp_path)
        doc = _import_doc(db_path)
        from services.training_generation_service import enqueue_job, get_job
        job_id = enqueue_job(db_path, unit_code="son_tay", actor="admin",
                             source_document_version_ids=[doc["version_id"]],
                             target_audience_codes=["nvkt"], requested_count=2)
        from training.cli import cmd_worker
        args = Namespace(db_path=db_path, unit_code="son_tay", provider="fake",
                         worker_id="test-w", once=True, poll_interval=1)
        assert cmd_worker(args) == 0
        job = get_job(db_path, job_id)
        assert job["status"] == "completed"
