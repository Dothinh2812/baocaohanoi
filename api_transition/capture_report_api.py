# -*- coding: utf-8 -*-
"""Capture report-api cho một report bất kỳ và sinh recipe JSON."""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urlparse

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from api_transition.auth import login
from api_transition.settings import Settings


DEFAULT_FILTER = "baocaobe.myhanoi.vn/report-api"


def parse_args():
    parser = argparse.ArgumentParser(
        description="Capture request/response report-api và sinh recipe cho chuyển đổi API."
    )
    parser.add_argument("--name", required=True, help="Tên recipe đầu ra, ví dụ c12_q2_2026")
    parser.add_argument("--report-url", required=True, help="URL report-info cần mở để thao tác tay")
    parser.add_argument("--output-dir", default="api_transition/captures")
    parser.add_argument("--recipe-dir", default="api_transition/recipes")
    parser.add_argument("--body-limit", type=int, default=8000)
    parser.add_argument("--headed", action="store_true")
    return parser.parse_args()


def ensure_text_preview(raw_value, limit):
    if raw_value is None:
        return ""
    text = raw_value.decode("utf-8", errors="replace") if isinstance(raw_value, bytes) else str(raw_value)
    if len(text) <= limit:
        return text
    return f"{text[:limit]}\n... [truncated {len(text) - limit} chars]"


def parse_json_safe(text):
    if not text:
        return None
    try:
        return json.loads(text)
    except (TypeError, json.JSONDecodeError):
        return None


def build_recipe(args, report_id, menu_id, export_payload, file_download_name=""):
    return {
        "name": args.name,
        "report_page_url": args.report_url,
        "report_id": report_id,
        "menu_id": menu_id,
        "export_payload": export_payload,
        "default_output_name": file_download_name or "",
        "captured_at": datetime.now().isoformat(),
        "notes": "Tự sinh từ capture_report_api.py",
    }


def main():
    args = parse_args()
    Settings.validate()

    capture_dir = Path(args.output_dir)
    recipe_dir = Path(args.recipe_dir)
    capture_dir.mkdir(parents=True, exist_ok=True)
    recipe_dir.mkdir(parents=True, exist_ok=True)

    log_path = capture_dir / f"{args.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jsonl"
    recipe_path = recipe_dir / f"{args.name}.json"

    parsed_report_url = urlparse(args.report_url)
    query = parse_qs(parsed_report_url.query)
    report_id = query.get("id", [""])[0]
    menu_id = query.get("menu_id", [""])[0]

    state = {
        "last_export_payload": None,
        "last_export_filename": "",
    }

    with log_path.open("a", encoding="utf-8") as log_file:
        def write_event(record):
            log_file.write(json.dumps(record, ensure_ascii=False) + "\n")
            log_file.flush()

        def on_request(request):
            if DEFAULT_FILTER not in request.url:
                return
            post_data = request.post_data or ""
            parsed_payload = parse_json_safe(post_data)

            record = {
                "timestamp": datetime.now().isoformat(),
                "event": "request",
                "method": request.method,
                "url": request.url,
                "headers": dict(request.headers),
                "post_data": ensure_text_preview(post_data, args.body_limit),
            }
            write_event(record)

            if request.url.endswith("/get-data-export") and parsed_payload:
                state["last_export_payload"] = parsed_payload
                print("✅ Bắt được payload get-data-export")

        def on_response(response):
            if DEFAULT_FILTER not in response.url:
                return
            try:
                body = response.body()
                content_type = response.headers.get("content-type", "")
                body_preview = ensure_text_preview(body, args.body_limit)
                parsed_body = parse_json_safe(body_preview if "application/json" in content_type else "")
            except Exception as exc:
                body_preview = f"[unable to read response body: {exc}]"
                parsed_body = None

            record = {
                "timestamp": datetime.now().isoformat(),
                "event": "response",
                "status": response.status,
                "url": response.url,
                "headers": dict(response.headers),
                "body_preview": body_preview,
            }
            write_event(record)

            if response.url.endswith("/get-data-export") and isinstance(parsed_body, dict):
                state["last_export_filename"] = parsed_body.get("FileDownloadName", "")
                print(f"✅ Bắt được response export: {state['last_export_filename']}")

        playwright = browser = context = page = None
        try:
            playwright, browser, context, page = login(headless=not args.headed)
            context.on("request", on_request)
            context.on("response", on_response)

            print(f"Đang mở report URL: {args.report_url}")
            page.goto(args.report_url, wait_until="networkidle", timeout=Settings.PAGE_LOAD_TIMEOUT)
            page.wait_for_load_state("networkidle", timeout=Settings.PAGE_LOAD_TIMEOUT)

            print("\n" + "=" * 80)
            print("CAPTURE ĐÃ BẬT")
            print("=" * 80)
            print(f"Log file  : {log_path}")
            print(f"Recipe    : {recipe_path}")
            print("Tiếp theo:")
            print("1. Thao tác tay trên trình duyệt.")
            print("2. Chạy báo cáo và bấm Xuất Excel.")
            print("3. Quay lại terminal và nhấn Enter để lưu recipe.")
            print("=" * 80 + "\n")
            input()

            if not state["last_export_payload"]:
                raise RuntimeError(
                    "Chưa bắt được get-data-export. Hãy chạy lại và thực hiện thao tác export trên giao diện."
                )

            recipe = build_recipe(
                args,
                report_id=report_id or state["last_export_payload"].get("reportId", ""),
                menu_id=menu_id,
                export_payload=state["last_export_payload"],
                file_download_name=state["last_export_filename"],
            )
            recipe_path.write_text(
                json.dumps(recipe, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"✅ Đã lưu recipe: {recipe_path}")

        finally:
            if browser is not None:
                browser.close()
            if playwright is not None:
                playwright.stop()


if __name__ == "__main__":
    main()
