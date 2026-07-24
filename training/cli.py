#!/usr/bin/env python3
"""CLI cho module Đào tạo & sát hạch.

Dùng cùng service và schema với UI. Entry point phải import runtime_limits
trước pandas/numpy.
"""

import runtime_limits  # noqa: F401  # must be first

import argparse
import sys

from training import migrations


def cmd_db_migrate(args):
    from training.db import TRAINING_DB_PATH, UNIT_CODE

    db_path = args.db_path or TRAINING_DB_PATH
    unit_code = args.unit_code or UNIT_CODE
    applied = migrations.run_migrations(db_path, unit_code)
    version = migrations.get_schema_version(db_path)
    print(f"Migration done. schema_version={version}, applied_new={len(applied) if applied else 0}")
    return 0


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

    db_path = args.db_path or TRAINING_DB_PATH
    provider = _get_provider(args.provider)
    import os

    worker_id = args.worker_id or f"cli-{os.getpid()}"
    lease_seconds = int(os.getenv("DASHV4_TRAINING_LEASE_SECONDS", "300"))

    job = gs.claim_next_job(db_path, worker_id=worker_id, lease_seconds=lease_seconds)
    if job is None:
        if args.once:
            print("No pending jobs.")
            return 0
        import time

        time.sleep(args.poll_interval)
        return 0

    print(f"Claimed job {job['id']} (retry {job['retry_count']})")
    try:
        import json

        source_ids = json.loads(job["source_document_version_ids_json"])
        audience_codes = json.loads(job["target_audience_codes_json"])
        result = provider.generate(
            source_document_version_ids=source_ids,
            target_audience_codes=audience_codes,
            requested_count=job["requested_count"],
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
        print(f"Failed job {job['id']}: {exc}")
    return 0


def build_parser():
    parser = argparse.ArgumentParser(prog="training.cli", description="CLI đào tạo & sát hạch")
    sub = parser.add_subparsers(dest="command", required=True)

    p_migrate = sub.add_parser("db-migrate", help="Chạy schema migration cho training.db")
    p_migrate.add_argument("--db-path", default=None)
    p_migrate.add_argument("--unit-code", default=None)
    p_migrate.set_defaults(func=cmd_db_migrate)

    p_worker = sub.add_parser("worker", help="Chạy AI generation worker")
    p_worker.add_argument("--db-path", default=None)
    p_worker.add_argument("--provider", default="fake")
    p_worker.add_argument("--worker-id", default=None)
    p_worker.add_argument("--once", action="store_true", help="Chạy một lần rồi thoát")
    p_worker.add_argument("--poll-interval", type=int, default=5)
    p_worker.set_defaults(func=cmd_worker)

    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
