# -*- coding: utf-8 -*-
"""PoC export C1.1 qua API trong thư mục chuyển đổi."""

import argparse
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from api_transition.auth import capture_authorization, login
from api_transition.report_api_client import (
    build_report_page_url,
    export_report,
    find_value_by_label,
    get_info_report,
    make_common_headers,
    print_candidate_pairs,
    save_export_file,
)
from api_transition.settings import Settings


DEFAULT_REPORT_ID = "534964"
DEFAULT_MENU_ID = "535020"
DEFAULT_UNIT_ID = "14324"
DEFAULT_MONTH_LABEL = "Tháng 04/2026"
DEFAULT_MONTH_ID = ""
DEFAULT_PLOAICT = "1"
DEFAULT_PLOAI = "1"
DEFAULT_OUTPUT_DIR = str(Path(__file__).resolve().parent / "downloads" / "chi_tieu_c")


def parse_args():
    parser = argparse.ArgumentParser(description="Export C1.1 qua API")
    parser.add_argument("--report-id", default=DEFAULT_REPORT_ID)
    parser.add_argument("--menu-id", default=DEFAULT_MENU_ID)
    parser.add_argument("--unit-id", default=DEFAULT_UNIT_ID)
    parser.add_argument("--month-label", default=DEFAULT_MONTH_LABEL)
    parser.add_argument("--month-id", default=DEFAULT_MONTH_ID)
    parser.add_argument("--ploaict", default=DEFAULT_PLOAICT)
    parser.add_argument("--ploai", default=DEFAULT_PLOAI)
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--output-name", default="")
    parser.add_argument("--headed", action="store_true")
    return parser.parse_args()


def build_export_payload(args, month_id):
    return {
        "reportId": str(args.report_id),
        "lstInputParams": [
            {"name": "ptrungtamid", "dataType": "NUMBER", "value": str(args.unit_id)},
            {"name": "pthang", "dataType": "VARCHAR2", "value": str(month_id)},
            {"name": "ploaict", "dataType": "NUMBER", "value": str(args.ploaict)},
            {"name": "ploai", "dataType": "NUMBER", "value": str(args.ploai)},
        ],
        "lstOutputParams": [
            {"name": "odata", "dataType": "CURSOR", "value": ""}
        ],
    }


def main():
    args = parse_args()
    Settings.validate()

    playwright = browser = context = page = None
    try:
        playwright, browser, context, page = login(headless=not args.headed)

        report_page_url = build_report_page_url(args.report_id, args.menu_id)
        auth_state = capture_authorization(page, report_page_url)
        headers = make_common_headers(auth_state, context.cookies())

        info_payload = get_info_report(args.report_id, args.menu_id, headers)

        month_id = args.month_id.strip()
        if not month_id:
            try:
                month_id = find_value_by_label(info_payload, args.month_label)
                print(f"✅ Đã map month-label '{args.month_label}' -> pthang={month_id}")
            except RuntimeError:
                print_candidate_pairs(info_payload, args.month_label)
                raise
        else:
            print(f"✅ Sử dụng month-id truyền vào: {month_id}")

        export_payload = build_export_payload(args, month_id)
        export_response = export_report(headers, export_payload)
        output_path = save_export_file(export_response, args.output_dir, args.output_name)
        print(f"✅ Đã lưu file: {output_path}")

    finally:
        if browser is not None:
            browser.close()
        if playwright is not None:
            playwright.stop()


if __name__ == "__main__":
    main()
