#!/usr/bin/env python3
"""CLI cho module Đào tạo & sát hạch.

Dùng cùng service và schema với UI — không SQL trực tiếp.
Entry point phải import runtime_limits trước pandas/numpy.
Mọi lệnh mutating kiểm tra module role theo --actor qua permissions.require_module_role.
"""

import runtime_limits  # noqa: F401  # must be first

import argparse
import json
import os
import sys

from training import migrations


def _print_json(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


def _require_role(db_path, username, role):
    """Kiểm tra module role cho CLI mutation. Raise TrainingError nếu không có quyền."""
    from training.permissions import require_module_role
    require_module_role(db_path, username, role)


def _handle_error(exc):
    """In error thân thiện và trả exit code."""
    from training.errors import TrainingError
    if isinstance(exc, TrainingError):
        print(f"[{exc.code}] {exc.message}", file=sys.stderr)
        return 1 if exc.status >= 500 else 2
    print(f"ERROR: {exc}", file=sys.stderr)
    return 1


# ---- db-migrate ----

def cmd_db_migrate(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    applied = migrations.run_migrations(db_path, unit_code)
    version = migrations.get_schema_version(db_path)
    print(f"schema_version={version}, applied_new={len(applied) if applied else 0}")
    return 0


# ---- provider/worker ----

def _get_provider(provider_name):
    if provider_name == "fake":
        from training.providers.fake import FakeProvider
        return FakeProvider()
    if provider_name == "openai":
        from training.providers.openai_provider import OpenAIProvider
        return OpenAIProvider()
    raise ValueError(f"Unknown provider: {provider_name}")


def cmd_worker(args):
    from training.db import TRAINING_DB_PATH
    from services import training_generation_service as gs
    from services.training_knowledge_service import resolve_generation_snapshot

    db_path = args.db_path or TRAINING_DB_PATH
    provider = _get_provider(args.provider)

    worker_id = args.worker_id or f"cli-{os.getpid()}"
    lease_seconds = int(os.getenv("DASHV4_TRAINING_LEASE_SECONDS", "300"))

    import time
    while True:
        job = gs.claim_next_job(db_path, worker_id=worker_id, lease_seconds=lease_seconds)
        if job is None:
            if args.once:
                print("No pending jobs.")
                return 0
            time.sleep(args.poll_interval)
            continue

        print(f"Claimed job {job['id']} (retry {job['retry_count']})")
        try:
            source_ids = json.loads(job["source_document_version_ids_json"])
            audience_codes = json.loads(job["target_audience_codes_json"])

            snapshot = resolve_generation_snapshot(db_path, source_ids, audience_codes)

            is_fake = getattr(provider, "name", "") == "fake"
            result = provider.generate(
                source_document_version_ids=source_ids,
                target_audience_codes=audience_codes,
                requested_count=job["requested_count"],
                **({} if is_fake else {"snapshot": snapshot}),
            )
            gs.complete_job(
                db_path, job_id=job["id"], worker_id=worker_id,
                batch=result.batch,
                provider_metadata={
                    "provider": result.provider,
                    "model": result.model,
                    "prompt_version": result.prompt_version,
                    "usage": result.usage,
                    "raw_response": result.raw_response,
                },
            )
            print(f"Completed job {job['id']}")
        except Exception as exc:
            gs.fail_job(db_path, job_id=job["id"], worker_id=worker_id,
                        error_code="WORKER_ERROR", error_detail=str(exc))
            print(f"Failed job {job['id']}: {exc}", file=sys.stderr)
        if args.once:
            return 0


# ---- knowledge-import ----

def cmd_knowledge_import(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE
    from training.permissions import require_module_role
    from services.training_knowledge_service import (
        create_document, find_version_by_checksum,
    )
    from services.training_file_ingestion import (
        FileValidationError, detect_file_type, validate_txt_file,
        validate_docx_file, extract_docx_text, validate_paste_text,
    )

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    require_module_role(db_path, args.actor, "editor")

    if args.paste_text is not None:
        content_text = validate_paste_text(args.paste_text)
    elif args.paste_file:
        content_text = validate_paste_text(_read_file_text(args.paste_file))
    elif args.file:
        file_type = detect_file_type(args.file)
        if file_type == "txt":
            content_text = validate_txt_file(args.file)
        elif file_type == "docx":
            validate_docx_file(args.file)
            content_text = extract_docx_text(args.file)
        else:
            raise FileValidationError(
                f"Loại file không hỗ trợ: {file_type}", code="FILE_TYPE_UNSUPPORTED",
            )
    else:
        raise ValueError("Yêu cầu --file, --paste-file, hoặc --paste-text")

    if not content_text.strip():
        raise ValueError("Nội dung rỗng")

    import hashlib
    content_sha256 = hashlib.sha256(content_text.encode("utf-8")).hexdigest()
    existing = find_version_by_checksum(db_path, content_sha256)
    if existing and not args.force:
        raise ValueError(
            f"Nội dung trùng checksum với version đã có: "
            f"document_version_id={existing['id']}, "
            f"version_number={existing['version_number']}. "
            f"Dùng --force để tạo version mới cho cùng document."
        )

    topic_codes = [t.strip() for t in args.topics.split(",")] if args.topics else []
    classification = {"domain_code": args.domain, "topic_codes": topic_codes}
    audience_codes = [a.strip() for a in args.audiences.split(",")] if args.audiences else []

    document_code = args.document_code or args.title.lower().replace(" ", "_")[:60]

    result = create_document(
        db_path,
        unit_code=unit_code,
        actor=args.actor,
        document_code=document_code,
        title=args.title,
        content_text=content_text,
        classification=classification,
        audience_codes=audience_codes,
        document_type=args.document_type,
        issuer=args.issuer,
    )

    _print_json({
        "status": "created",
        "document_id": result["document_id"],
        "document_version_id": result["version_id"],
        "version_number": result["version_number"],
        "content_sha256": content_sha256,
    })
    return 0


def _read_file_text(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


# ---- knowledge-list ----

def cmd_knowledge_list(args):
    from training.db import TRAINING_DB_PATH
    from services.training_knowledge_service import list_documents

    db_path = args.db_path or TRAINING_DB_PATH
    result = list_documents(
        db_path, domain_code=args.domain, status=args.status,
        page=args.page, page_size=args.page_size,
    )
    print(f"Total: {result['total']}, Page: {result['page']}")
    for item in result["items"]:
        print(f"  {item['id']} | {item['document_code']} | {item['title'][:50]} | {item['review_status']}")
    return 0


# ---- knowledge-show ----

def cmd_knowledge_show(args):
    from training.db import TRAINING_DB_PATH
    from services.training_knowledge_service import get_version, list_blocks

    db_path = args.db_path or TRAINING_DB_PATH
    version = get_version(db_path, args.document_version_id)
    if not version:
        print(f"Version {args.document_version_id} không tồn tại", file=sys.stderr)
        return 1
    blocks = list_blocks(db_path, args.document_version_id)
    _print_json({
        "version": {k: version[k] for k in
                    ("id", "document_id", "version_number", "review_status",
                     "content_sha256", "created_by", "created_at_ms")},
        "block_count": len(blocks),
        "blocks": [
            {
                "block_id": b["block_id"],
                "extraction_revision": b["extraction_revision"],
                "char_start": b["char_start"],
                "char_end": b["char_end"],
                "content_preview": b["content"][:100],
                "content_sha256": b["content_sha256"],
            }
            for b in blocks
        ],
    })
    return 0


# ---- knowledge-issues ----

def cmd_knowledge_issues(args):
    from training.db import TRAINING_DB_PATH
    from services.training_knowledge_service import list_issues, has_blocking_issues

    db_path = args.db_path or TRAINING_DB_PATH
    issues = list_issues(db_path, args.document_version_id)
    blocking = has_blocking_issues(db_path, args.document_version_id)
    _print_json({
        "document_version_id": args.document_version_id,
        "has_blocking_issues": blocking,
        "issue_count": len(issues),
        "issues": [
            {
                "id": i["id"],
                "severity": i["severity"],
                "issue_type": i["issue_type"],
                "status": i["status"],
                "description": i["description"],
                "block_id": i["block_id"],
            }
            for i in issues
        ],
    })
    return 0


# ---- generate-create ----

def cmd_generate_create(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE
    from training.permissions import require_module_role
    from services.training_generation_service import enqueue_job
    from services.training_knowledge_service import has_blocking_issues, get_version

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    require_module_role(db_path, args.actor, "editor")

    source_ids = [s.strip() for s in args.document_version_ids.split(",")]
    for version_id in source_ids:
        ver = get_version(db_path, version_id)
        if not ver:
            raise ValueError(f"Document version {version_id} không tồn tại")
        if has_blocking_issues(db_path, version_id):
            raise ValueError(
                f"Document version {version_id} có blocking issue — "
                "resolve trước khi generate"
            )

    audience_codes = [a.strip() for a in args.audiences.split(",")]

    idempotency_key = args.idempotency_key
    if not idempotency_key:
        import hashlib
        raw = "|".join(sorted(source_ids)) + "|" + "|".join(sorted(audience_codes)) + f"|{args.count}"
        idempotency_key = "auto-" + hashlib.sha256(raw.encode()).hexdigest()[:16]

    job_id = enqueue_job(
        db_path,
        unit_code=unit_code,
        actor=args.actor,
        source_document_version_ids=source_ids,
        target_audience_codes=audience_codes,
        requested_count=args.count,
        idempotency_key=idempotency_key,
    )
    _print_json({"status": "created", "job_id": job_id, "idempotency_key": idempotency_key})
    return 0


# ---- generate-list ----

def cmd_generate_list(args):
    from training.db import TRAINING_DB_PATH
    from services.training_generation_service import list_jobs

    db_path = args.db_path or TRAINING_DB_PATH
    result = list_jobs(
        db_path, status=args.status, actor=args.actor,
        page=args.page, page_size=args.page_size,
    )
    print(f"Total: {result['total']}, Page: {result['page']}")
    for item in result["items"]:
        print(f"  {item['id']} | {item['status']} | {item['requested_count']}q | "
              f"by={item['created_by']} | retry={item['retry_count']}")
    return 0


# ---- generate-show ----

def cmd_generate_show(args):
    from training.db import TRAINING_DB_PATH
    from services.training_generation_service import get_job_detail

    db_path = args.db_path or TRAINING_DB_PATH
    job = get_job_detail(db_path, args.job_id)
    if not job:
        print(f"Job {args.job_id} không tồn tại", file=sys.stderr)
        return 1
    _print_json({
        "id": job["id"],
        "status": job["status"],
        "requested_count": job["requested_count"],
        "created_by": job["created_by"],
        "created_at_ms": job["created_at_ms"],
        "source_document_version_ids": json.loads(job["source_document_version_ids_json"]),
        "target_audience_codes": json.loads(job["target_audience_codes_json"]),
        "claimed_by": job.get("claimed_by"),
        "retry_count": job.get("retry_count"),
        "error_code": job.get("error_code"),
        "error_detail": job.get("error_detail"),
        "completed_at_ms": job.get("completed_at_ms"),
        "batch_meta": job.get("batch_meta"),
    })
    return 0


# ---- generate-cancel ----

def cmd_generate_cancel(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE
    from training.permissions import require_module_role
    from services.training_generation_service import cancel_job

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    require_module_role(db_path, args.actor, "exam_manager")

    cancel_job(db_path, unit_code=unit_code, actor=args.actor, job_id=args.job_id)
    print(f"Cancelled job {args.job_id}")
    return 0


# ---- questions-list ----

def cmd_questions_list(args):
    from training.db import TRAINING_DB_PATH
    from services.training_question_service import list_questions

    db_path = args.db_path or TRAINING_DB_PATH
    result = list_questions(
        db_path, status=args.status, topic=args.topic,
        page=args.page, page_size=args.page_size,
    )
    print(f"Total: {result['total']}, Page: {result['page']}")
    for item in result["items"]:
        print(f"  {item['id']} | {item['stem'][:60]} | {item.get('_question_status', '?')}")
    return 0


# ---- questions-show ----

def cmd_questions_show(args):
    from training.db import TRAINING_DB_PATH
    from services.training_question_service import get_question_management_detail

    db_path = args.db_path or TRAINING_DB_PATH
    detail = get_question_management_detail(db_path, args.version_id)
    _print_json(detail)
    return 0


# ---- questions-approve ----

def cmd_questions_approve(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE
    from training.permissions import require_module_role
    from services.training_question_service import add_review_action

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    require_module_role(db_path, args.actor, "exam_manager")

    add_review_action(
        db_path, unit_code=unit_code, actor=args.actor,
        version_id=args.version_id, action="approve", comment=args.comment,
    )
    print(f"Approved version {args.version_id}")
    return 0


# ---- questions-reject ----

def cmd_questions_reject(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE
    from training.permissions import require_module_role
    from services.training_question_service import add_review_action

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    require_module_role(db_path, args.actor, "exam_manager")

    if not args.comment:
        raise ValueError("--comment bắt buộc khi reject")
    add_review_action(
        db_path, unit_code=unit_code, actor=args.actor,
        version_id=args.version_id, action="reject", comment=args.comment,
    )
    print(f"Rejected version {args.version_id}")
    return 0


# ---- questions-publish ----

def cmd_questions_publish(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE
    from training.permissions import require_module_role
    from services.training_question_service import publish_question_version

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    require_module_role(db_path, args.actor, "exam_manager")

    publish_question_version(
        db_path, unit_code=unit_code, actor=args.actor, version_id=args.version_id,
    )
    print(f"Published version {args.version_id}")
    return 0


# ---- parser ----

def build_parser():
    parser = argparse.ArgumentParser(
        prog="training.cli", description="CLI đào tạo & sát hạch",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    _add_db_migrate(sub)
    _add_worker(sub)
    _add_knowledge_import(sub)
    _add_knowledge_list(sub)
    _add_knowledge_show(sub)
    _add_knowledge_issues(sub)
    _add_generate_create(sub)
    _add_generate_list(sub)
    _add_generate_show(sub)
    _add_generate_cancel(sub)
    _add_questions_list(sub)
    _add_questions_show(sub)
    _add_questions_approve(sub)
    _add_questions_reject(sub)
    _add_questions_publish(sub)

    return parser


def _add_common_db(p):
    p.add_argument("--db-path", default=None, help="Override training.db path")
    p.add_argument("--unit-code", default=None, help="Override unit_code")


def _add_db_migrate(sub):
    p = sub.add_parser("db-migrate", help="Chạy schema migration")
    _add_common_db(p)
    p.set_defaults(func=cmd_db_migrate)


def _add_worker(sub):
    p = sub.add_parser("worker", help="Chạy AI generation worker")
    _add_common_db(p)
    p.add_argument("--provider", default="fake", choices=["fake", "openai"])
    p.add_argument("--worker-id", default=None)
    p.add_argument("--once", action="store_true")
    p.add_argument("--poll-interval", type=int, default=5)
    p.set_defaults(func=cmd_worker)


def _add_knowledge_import(sub):
    p = sub.add_parser("knowledge-import", help="Import tài liệu vào kho tri thức")
    _add_common_db(p)
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument("--file", help="Path file .txt hoặc .docx")
    src.add_argument("--paste-file", help="Path file text đã paste nội dung")
    src.add_argument("--paste-text", help="Paste text trực tiếp")
    p.add_argument("--title", required=True)
    p.add_argument("--document-code", default=None)
    p.add_argument("--domain", required=True, help="Domain code (vd: quality)")
    p.add_argument("--topics", default=None, help="CSV topic codes")
    p.add_argument("--audiences", default=None, help="CSV audience codes")
    p.add_argument("--document-type", default="kpi_definition")
    p.add_argument("--issuer", default=None)
    p.add_argument("--actor", required=True, help="Username thực hiện")
    p.add_argument("--force", action="store_true",
                   help="Tạo version mới ngay cả khi trùng checksum")
    p.set_defaults(func=cmd_knowledge_import)


def _add_knowledge_list(sub):
    p = sub.add_parser("knowledge-list", help="Liệt kê tài liệu")
    _add_common_db(p)
    p.add_argument("--domain", default=None)
    p.add_argument("--status", default=None)
    p.add_argument("--page", type=int, default=1)
    p.add_argument("--page-size", type=int, default=25)
    p.set_defaults(func=cmd_knowledge_list)


def _add_knowledge_show(sub):
    p = sub.add_parser("knowledge-show", help="Hiển thị chi tiết version")
    _add_common_db(p)
    p.add_argument("--document-version-id", required=True)
    p.set_defaults(func=cmd_knowledge_show)


def _add_knowledge_issues(sub):
    p = sub.add_parser("knowledge-issues", help="Hiển thị issues của version")
    _add_common_db(p)
    p.add_argument("--document-version-id", required=True)
    p.set_defaults(func=cmd_knowledge_issues)


def _add_generate_create(sub):
    p = sub.add_parser("generate-create", help="Tạo AI generation job")
    _add_common_db(p)
    p.add_argument("--document-version-ids", required=True,
                   help="CSV document version IDs")
    p.add_argument("--audiences", required=True, help="CSV audience codes")
    p.add_argument("--count", type=int, required=True)
    p.add_argument("--idempotency-key", default=None)
    p.add_argument("--actor", required=True)
    p.set_defaults(func=cmd_generate_create)


def _add_generate_list(sub):
    p = sub.add_parser("generate-list", help="Liệt kê generation jobs")
    _add_common_db(p)
    p.add_argument("--status", default=None)
    p.add_argument("--actor", default=None)
    p.add_argument("--page", type=int, default=1)
    p.add_argument("--page-size", type=int, default=25)
    p.set_defaults(func=cmd_generate_list)


def _add_generate_show(sub):
    p = sub.add_parser("generate-show", help="Hiển thị chi tiết job")
    _add_common_db(p)
    p.add_argument("--job-id", required=True)
    p.set_defaults(func=cmd_generate_show)


def _add_generate_cancel(sub):
    p = sub.add_parser("generate-cancel", help="Hủy generation job")
    _add_common_db(p)
    p.add_argument("--job-id", required=True)
    p.add_argument("--actor", required=True)
    p.set_defaults(func=cmd_generate_cancel)


def _add_questions_list(sub):
    p = sub.add_parser("questions-list", help="Liệt kê câu hỏi")
    _add_common_db(p)
    p.add_argument("--status", default=None,
                   help="draft/approved/published")
    p.add_argument("--topic", default=None)
    p.add_argument("--page", type=int, default=1)
    p.add_argument("--page-size", type=int, default=25)
    p.set_defaults(func=cmd_questions_list)


def _add_questions_show(sub):
    p = sub.add_parser("questions-show", help="Hiển thị chi tiết câu hỏi")
    _add_common_db(p)
    p.add_argument("--version-id", required=True)
    p.set_defaults(func=cmd_questions_show)


def _add_questions_approve(sub):
    p = sub.add_parser("questions-approve", help="Approve câu hỏi (CAS)")
    _add_common_db(p)
    p.add_argument("--version-id", required=True)
    p.add_argument("--actor", required=True)
    p.add_argument("--comment", default=None)
    p.set_defaults(func=cmd_questions_approve)


def _add_questions_reject(sub):
    p = sub.add_parser("questions-reject", help="Reject câu hỏi (CAS)")
    _add_common_db(p)
    p.add_argument("--version-id", required=True)
    p.add_argument("--actor", required=True)
    p.add_argument("--comment", required=True)
    p.set_defaults(func=cmd_questions_reject)


def _add_questions_publish(sub):
    p = sub.add_parser("questions-publish", help="Publish câu hỏi (CAS + evidence)")
    _add_common_db(p)
    p.add_argument("--version-id", required=True)
    p.add_argument("--actor", required=True)
    p.set_defaults(func=cmd_questions_publish)


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:
        return _handle_error(exc)


if __name__ == "__main__":
    sys.exit(main())
