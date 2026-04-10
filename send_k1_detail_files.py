#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gửi chủ động các file chi tiết suy hao cao K1 qua Zalo.

Module này tách riêng khỏi i15_cts_converter.py để không gửi ngay sau khi tạo.
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

from config_for_send_zalo.openzca_file_sender import (
    DEFAULT_DELAY_SECONDS,
    DEFAULT_UPLOAD_TIMEOUT_SECONDS,
    build_summary,
    send_nvkt_files_in_folder,
)


DEFAULT_REPORT_FILE = Path("downloads/baocao_hanoi/I1.5 report.xlsx")
DEFAULT_DETAIL_DIR = Path("downloads/baocao_hanoi/shc_NVKT_danh_sach_chi_tiet_K1")


def get_k1_day_month_from_report(report_file: Path | str = DEFAULT_REPORT_FILE) -> str:
    """
    Lấy ngày/tháng từ file I1.5 report.xlsx để dùng làm tiêu đề gửi Zalo.
    """
    report_file = Path(report_file)
    fallback_value = datetime.now().strftime("%d/%m")

    if not report_file.exists():
        return fallback_value

    try:
        df_report = pd.read_excel(report_file, nrows=1)
        if "NGAY_SUYHAO" not in df_report.columns or len(df_report) == 0:
            return fallback_value

        ngay_val = df_report["NGAY_SUYHAO"].iloc[0]
        if pd.isna(ngay_val):
            return fallback_value

        parsed_date = pd.to_datetime(ngay_val, dayfirst=True, errors="coerce")
        if pd.isna(parsed_date):
            return fallback_value

        return parsed_date.strftime("%d/%m")
    except Exception:
        return fallback_value


def build_k1_title_text(report_file: Path | str = DEFAULT_REPORT_FILE) -> str:
    title_day_month = get_k1_day_month_from_report(report_file)
    return f"K/g các anh chi tiết shc k1 ngày {title_day_month}:"


def send_k1_detail_files(
    *,
    detail_dir: Path | str = DEFAULT_DETAIL_DIR,
    report_file: Path | str = DEFAULT_REPORT_FILE,
    dry_run: bool = True,
    delay_seconds: float = DEFAULT_DELAY_SECONDS,
    timeout_seconds: int = DEFAULT_UPLOAD_TIMEOUT_SECONDS,
    debug_openzca: bool = False,
) -> dict[str, int]:
    detail_dir = Path(detail_dir)
    if not detail_dir.is_dir():
        raise FileNotFoundError(f"Không tìm thấy thư mục: {detail_dir}")

    title_text = build_k1_title_text(report_file)
    results = send_nvkt_files_in_folder(
        detail_dir,
        dry_run=dry_run,
        delay_seconds=delay_seconds,
        timeout_seconds=timeout_seconds,
        debug=debug_openzca,
        message_text=title_text,
    )
    return build_summary(results)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Gửi chủ động các file chi tiết suy hao cao K1 qua Zalo."
    )
    parser.add_argument(
        "--detail-dir",
        default=str(DEFAULT_DETAIL_DIR),
        help="Thư mục chứa các file chi tiết K1 theo từng tổ.",
    )
    parser.add_argument(
        "--report-file",
        default=str(DEFAULT_REPORT_FILE),
        help="File I1.5 report.xlsx để lấy ngày gửi lên tiêu đề.",
    )
    parser.add_argument(
        "--delay-seconds",
        type=float,
        default=DEFAULT_DELAY_SECONDS,
        help="Số giây nghỉ giữa các lần upload file.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=DEFAULT_UPLOAD_TIMEOUT_SECONDS,
        help="Timeout cho mỗi file khi gọi openzca.",
    )
    parser.add_argument(
        "--debug-openzca",
        action="store_true",
        help="Bật OPENZCA_DEBUG=1 cho từng lệnh gửi.",
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Thực hiện gửi thật. Mặc định chỉ dry run để kiểm tra.",
    )
    args = parser.parse_args()

    title_text = build_k1_title_text(args.report_file)
    print(f"Tiêu đề gửi: {title_text}")
    print(f"Thư mục chi tiết K1: {args.detail_dir}")
    print(f"Chế độ: {'GỬI THẬT' if args.send else 'DRY RUN'}")

    summary = send_k1_detail_files(
        detail_dir=args.detail_dir,
        report_file=args.report_file,
        dry_run=not args.send,
        delay_seconds=args.delay_seconds,
        timeout_seconds=args.timeout_seconds,
        debug_openzca=args.debug_openzca,
    )

    print(f"Tổng file: {summary['total']}")
    print(f"Thành công: {summary['success']}")
    print(f"Thất bại: {summary['failed']}")

    return 0 if summary["failed"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
