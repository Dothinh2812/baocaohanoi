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


def build_parser():
    parser = argparse.ArgumentParser(prog="training.cli", description="CLI đào tạo & sát hạch")
    sub = parser.add_subparsers(dest="command", required=True)

    p_migrate = sub.add_parser("db-migrate", help="Chạy schema migration cho training.db")
    p_migrate.add_argument("--db-path", default=None)
    p_migrate.add_argument("--unit-code", default=None)
    p_migrate.set_defaults(func=cmd_db_migrate)

    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
