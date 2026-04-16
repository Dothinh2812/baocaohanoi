# -*- coding: utf-8 -*-
"""Chạy hàm download cũ ở chế độ headless và đồng thời capture report-api."""

import argparse
import importlib
import json
import sys
from datetime import datetime
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from api_transition.settings import Settings
from login import login_baocao_hanoi


DEFAULT_FILTER = "baocaobe.myhanoi.vn/report-api"
LEGACY_FUNCTIONS = {
    "download_report_c11": ("c1_report_download", "download_report_c11"),
    "download_report_c11_chitiet": ("c1_report_download", "download_report_c11_chitiet"),
    "download_report_c11_chitiet_SM2": ("c1_report_download", "download_report_c11_chitiet_SM2"),
    "download_report_c12": ("c1_report_download", "download_report_c12"),
    "download_report_c12_chitiet_SM1": ("c1_report_download", "download_report_c12_chitiet_SM1"),
    "download_report_c12_chitiet_SM2": ("c1_report_download", "download_report_c12_chitiet_SM2"),
    "download_report_c13": ("c1_report_download", "download_report_c13"),
    "download_report_c14": ("c1_report_download", "download_report_c14"),
    "download_report_c14_chitiet": ("c1_report_download", "download_report_c14_chitiet"),
    "download_report_c15": ("c1_report_download", "download_report_c15"),
    "download_report_c15_chitiet": ("c1_report_download", "download_report_c15_chitiet"),
    "download_report_I15": ("c1_report_download", "download_report_I15"),
    "download_report_I15_k2": ("c1_report_download", "download_report_I15_k2"),
    "c11_download_report_nvkt": ("kpi_download_from_baocaohanoi", "c11_download_report_nvkt"),
    "c12_download_report_nvkt": ("kpi_download_from_baocaohanoi", "c12_download_report_nvkt"),
    "c13_download_report_nvkt": ("kpi_download_from_baocaohanoi", "c13_download_report_nvkt"),
    "download_GHTT_report_HNI": ("KR_download", "download_GHTT_report_HNI"),
    "download_GHTT_report_Son_Tay": ("KR_download", "download_GHTT_report_Son_Tay"),
    "download_GHTT_report_nvktdb": ("KR_download", "download_GHTT_report_nvktdb"),
    "download_report_vattu_thuhoi": ("vat_tu_thu_hoi_download", "download_report_vattu_thuhoi"),
    "kq_tiep_thi_download": ("kq_tiep_thi_download", "kq_tiep_thi_download"),
    "xac_minh_tam_dung_download": ("xac_minh_tam_dung_download", "xac_minh_tam_dung_download"),
}


def parse_args():
    parser = argparse.ArgumentParser(
        description="Capture report-api bằng cách chạy hàm download cũ ở chế độ headless."
    )
    parser.add_argument("--name", required=True, help="Tên recipe đầu ra.")
    parser.add_argument("--legacy-func", required=True, choices=sorted(LEGACY_FUNCTIONS))
    parser.add_argument("--report-url", default="", help="Tùy chọn: URL report-info để lưu vào recipe.")
    parser.add_argument("--report-month", default="", help="Truyền cho các hàm cũ nhận tham số report_month.")
    parser.add_argument("--start-date", default="", help="Truyền cho các hàm chi tiết nhận start_date.")
    parser.add_argument("--end-date", default="", help="Truyền cho các hàm chi tiết nhận end_date.")
    parser.add_argument("--output-dir", default="api_transition/captures")
    parser.add_argument("--recipe-dir", default="api_transition/recipes")
    parser.add_argument("--body-limit", type=int, default=8000)
    return parser.parse_args()


def parse_json_safe(text):
    if not text:
        return None
    try:
        return json.loads(text)
    except (TypeError, json.JSONDecodeError):
        return None


def ensure_text_preview(raw_value, limit):
    if raw_value is None:
        return ""
    text = raw_value.decode("utf-8", errors="replace") if isinstance(raw_value, bytes) else str(raw_value)
    if len(text) <= limit:
        return text
    return f"{text[:limit]}\n... [truncated {len(text) - limit} chars]"


def parse_report_ids(report_url, export_payload):
    report_id = ""
    menu_id = ""
    if report_url and "?" in report_url:
        query = report_url.split("?", 1)[1]
        params = dict(item.split("=", 1) for item in query.split("&") if "=" in item)
        report_id = params.get("id", "")
        menu_id = params.get("menu_id", "")
    if not report_id and export_payload:
        report_id = str(export_payload.get("reportId", ""))
    return report_id, menu_id


def build_recipe(name, report_url, export_payload, file_download_name):
    report_id, menu_id = parse_report_ids(report_url, export_payload)
    return {
        "name": name,
        "report_page_url": report_url,
        "report_id": report_id,
        "menu_id": menu_id,
        "export_payload": export_payload,
        "default_output_name": file_download_name or "",
        "captured_at": datetime.now().isoformat(),
        "notes": "Sinh tự động từ capture_with_legacy_flow.py",
    }


def load_legacy_function(func_name):
    module_name, attr_name = LEGACY_FUNCTIONS[func_name]
    module = importlib.import_module(module_name)
    return getattr(module, attr_name)


def invoke_legacy_function(func, page, args):
    if args.start_date and args.end_date:
        return func(page, args.start_date, args.end_date)
    if args.report_month:
        return func(page, args.report_month)
    return func(page)


def main():
    args = parse_args()
    Settings.validate()

    capture_dir = Path(args.output_dir)
    recipe_dir = Path(args.recipe_dir)
    capture_dir.mkdir(parents=True, exist_ok=True)
    recipe_dir.mkdir(parents=True, exist_ok=True)

    log_path = capture_dir / f"{args.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jsonl"
    recipe_path = recipe_dir / f"{args.name}.json"

    state = {
        "last_export_payload": None,
        "last_export_filename": "",
    }

    legacy_function = load_legacy_function(args.legacy_func)

    with log_path.open("a", encoding="utf-8") as log_file:
        def write_event(record):
            log_file.write(json.dumps(record, ensure_ascii=False) + "\n")
            log_file.flush()

        def on_request(request):
            if DEFAULT_FILTER not in request.url:
                return
            post_data = request.post_data or ""
            parsed_payload = parse_json_safe(post_data)
            write_event(
                {
                    "timestamp": datetime.now().isoformat(),
                    "event": "request",
                    "method": request.method,
                    "url": request.url,
                    "headers": dict(request.headers),
                    "post_data": ensure_text_preview(post_data, args.body_limit),
                }
            )
            if request.url.endswith("/get-data-export") and parsed_payload:
                state["last_export_payload"] = parsed_payload
                print("✅ Bắt được payload get-data-export")

        def on_response(response):
            if DEFAULT_FILTER not in response.url:
                return
            try:
                body = response.body()
                body_preview = ensure_text_preview(body, args.body_limit)
                parsed_body = parse_json_safe(body_preview)
            except Exception as exc:
                body_preview = f"[unable to read response body: {exc}]"
                parsed_body = None

            write_event(
                {
                    "timestamp": datetime.now().isoformat(),
                    "event": "response",
                    "status": response.status,
                    "url": response.url,
                    "headers": dict(response.headers),
                    "body_preview": body_preview,
                }
            )
            if response.url.endswith("/get-data-export") and isinstance(parsed_body, dict):
                state["last_export_filename"] = parsed_body.get("FileDownloadName", "")
                print(f"✅ Bắt được response export: {state['last_export_filename']}")

        page = browser = playwright = None
        try:
            Settings.validate()
            from config import Config
            Config.BROWSER_HEADLESS = True
            page, browser, playwright = login_baocao_hanoi()
            context = page.context
            context.on("request", on_request)
            context.on("response", on_response)

            print(f"Đang chạy hàm cũ: {args.legacy_func}")
            invoke_legacy_function(legacy_function, page, args)

            if not state["last_export_payload"]:
                raise RuntimeError(
                    "Chưa bắt được get-data-export. Hàm cũ có thể chưa đi qua luồng export JSON."
                )

            recipe = build_recipe(
                args.name,
                args.report_url,
                state["last_export_payload"],
                state["last_export_filename"],
            )
            recipe_path.write_text(
                json.dumps(recipe, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"✅ Đã lưu log: {log_path}")
            print(f"✅ Đã lưu recipe: {recipe_path}")

        finally:
            if browser is not None:
                browser.close()
            if playwright is not None:
                playwright.stop()


if __name__ == "__main__":
    main()
