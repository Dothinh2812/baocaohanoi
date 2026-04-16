# -*- coding: utf-8 -*-
"""
PoC tải báo cáo qua API thật sau khi login bằng Playwright.

Luồng:
1. Login như code hiện tại (gồm OTP)
2. Mở trang report để bắt Authorization header của report-api
3. Gọi get-info-report để tra ID kỳ báo cáo nếu cần
4. Gọi get-data-export
5. Decode FileContents base64 và lưu ra .xlsx

Ví dụ:
    python3 download_report_via_api.py --headed
    python3 download_report_via_api.py --headed --month-label "Tháng 04/2026"
    python3 download_report_via_api.py --month-id 98944548
"""

import argparse
import base64
import json
import re
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from config import Config
from login import login_baocao_hanoi


API_BASE_URL = "https://baocaobe.myhanoi.vn/report-api"
DEFAULT_REPORT_ID = "534964"
DEFAULT_MENU_ID = "535020"
DEFAULT_UNIT_ID = "14324"
DEFAULT_UNIT_LABEL = "TTVT Sơn Tây"
DEFAULT_MONTH_LABEL = "Tháng 04/2026"
DEFAULT_PLOAICT = "1"
DEFAULT_PLOAI = "1"


def parse_args():
    parser = argparse.ArgumentParser(
        description="Tải báo cáo trực tiếp qua API get-data-export."
    )
    parser.add_argument("--report-id", default=DEFAULT_REPORT_ID)
    parser.add_argument("--menu-id", default=DEFAULT_MENU_ID)
    parser.add_argument("--unit-id", default=DEFAULT_UNIT_ID)
    parser.add_argument("--unit-label", default=DEFAULT_UNIT_LABEL)
    parser.add_argument("--month-id", default="")
    parser.add_argument("--month-label", default=DEFAULT_MONTH_LABEL)
    parser.add_argument("--ploaict", default=DEFAULT_PLOAICT)
    parser.add_argument("--ploai", default=DEFAULT_PLOAI)
    parser.add_argument(
        "--output-dir",
        default="downloads/baocao_hanoi",
        help="Thư mục lưu file xlsx.",
    )
    parser.add_argument(
        "--output-name",
        default="",
        help="Tên file đầu ra. Nếu bỏ trống sẽ dùng tên từ API.",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Ép mở browser có giao diện khi login.",
    )
    return parser.parse_args()


def normalize_spaces(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().casefold()


def sanitize_filename(filename):
    safe_name = re.sub(r'[\\/:*?"<>|]+', "_", str(filename or "").strip())
    safe_name = re.sub(r"\s+", " ", safe_name).strip(" .")
    return safe_name or "report.xlsx"


def build_report_page_url(report_id, menu_id):
    return f"{Config.BAOCAO_BASE_URL}/report/report-info?id={report_id}&menu_id={menu_id}"


def build_cookie_header(cookies):
    return "; ".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in cookies
        if cookie.get("name") and cookie.get("value")
    )


def extract_auth_headers(page, report_url, timeout_seconds=30):
    state = {}

    def on_request(request):
        if "/report-api/" not in request.url:
            return
        authorization = request.headers.get("authorization")
        if not authorization:
            return
        if "authorization" not in state:
            state["authorization"] = authorization
            state["user_agent"] = request.headers.get("user-agent", "")
            state["accept"] = request.headers.get("accept", "application/json, text/plain, */*")
            state["referer"] = request.headers.get("referer", f"{Config.BAOCAO_BASE_URL}/")

    context = page.context
    context.on("request", on_request)

    print(f"Đang mở trang report để bắt Authorization: {report_url}")
    page.goto(report_url, wait_until="networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)

    started = time.time()
    while time.time() - started < timeout_seconds:
        if state.get("authorization"):
            break
        page.wait_for_timeout(500)

    if not state.get("authorization"):
        raise RuntimeError("Không bắt được Authorization header từ request /report-api/.")

    print("✅ Đã bắt được Authorization header")
    return state


def http_json_request(url, method="GET", headers=None, payload=None, timeout=120):
    request_headers = dict(headers or {})
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request_headers.setdefault("Content-Type", "application/json")

    req = Request(url=url, data=data, headers=request_headers, method=method.upper())
    try:
        with urlopen(req, timeout=timeout) as response:
            body = response.read().decode("utf-8")
    except HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"HTTP {exc.code} khi gọi {url}\nResponse:\n{error_body}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(f"Lỗi kết nối khi gọi {url}: {exc}") from exc

    try:
        return json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Response không phải JSON từ {url}:\n{body[:1000]}") from exc


def collect_text_value_pairs(node, path="root"):
    results = []
    if isinstance(node, dict):
        if "text" in node and "value" in node:
            results.append(
                {
                    "text": node.get("text"),
                    "value": node.get("value"),
                    "path": path,
                    "node": node,
                }
            )
        for key, value in node.items():
            results.extend(collect_text_value_pairs(value, f"{path}.{key}"))
    elif isinstance(node, list):
        for index, item in enumerate(node):
            results.extend(collect_text_value_pairs(item, f"{path}[{index}]"))
    return results


def find_pairs_by_keywords(info_payload, keywords=None):
    keywords = [normalize_spaces(item) for item in (keywords or []) if item]
    pairs = collect_text_value_pairs(info_payload)
    if not keywords:
        return pairs
    matches = []
    for pair in pairs:
        normalized_text = normalize_spaces(pair["text"])
        if any(keyword in normalized_text for keyword in keywords):
            matches.append(pair)
    return matches


def find_value_by_label(info_payload, label):
    normalized_target = normalize_spaces(label)
    pairs = collect_text_value_pairs(info_payload)

    exact_matches = [
        pair for pair in pairs if normalize_spaces(pair["text"]) == normalized_target
    ]
    if exact_matches:
        return exact_matches[0]["value"]

    contains_matches = [
        pair for pair in pairs if normalized_target and normalized_target in normalize_spaces(pair["text"])
    ]
    if contains_matches:
        return contains_matches[0]["value"]

    raise RuntimeError(
        f"Không tìm thấy option có text '{label}' trong get-info-report."
    )


def print_candidate_pairs(info_payload, label, limit=20):
    print("\nKhông map được month-label. Một số option text/value gần đúng để tham chiếu:")

    search_keywords = []
    normalized_label = normalize_spaces(label)
    if normalized_label:
        search_keywords.extend(
            token for token in re.split(r"[\s/()-]+", normalized_label) if len(token) >= 2
        )
    search_keywords.extend(["tháng", "quý", "2026", "2025"])

    seen = set()
    candidates = []
    for pair in find_pairs_by_keywords(info_payload, search_keywords):
        key = (str(pair["text"]), str(pair["value"]))
        if key in seen:
            continue
        seen.add(key)
        candidates.append(pair)

    if not candidates:
        candidates = collect_text_value_pairs(info_payload)[:limit]

    for index, pair in enumerate(candidates[:limit], start=1):
        print(f"{index}. text={pair['text']!r} value={pair['value']!r} path={pair['path']}")


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


def save_exported_file(export_response, output_dir, output_name=""):
    if "FileContents" not in export_response:
        raise RuntimeError(
            f"Response export không chứa FileContents:\n{json.dumps(export_response, ensure_ascii=False)[:1000]}"
        )

    file_name = output_name or export_response.get("FileDownloadName") or "report.xlsx"
    safe_name = sanitize_filename(file_name)

    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    output_path = output_root / safe_name

    file_bytes = base64.b64decode(export_response["FileContents"])
    output_path.write_bytes(file_bytes)
    return output_path


def main():
    args = parse_args()
    Config.validate()

    if args.headed:
        Config.BROWSER_HEADLESS = False

    page = browser = playwright = None
    try:
        page, browser, playwright = login_baocao_hanoi()
        context = page.context

        report_url = build_report_page_url(args.report_id, args.menu_id)
        auth_state = extract_auth_headers(page, report_url)
        cookies = context.cookies()

        common_headers = {
            "Authorization": auth_state["authorization"],
            "Accept": auth_state["accept"],
            "Referer": auth_state["referer"],
            "User-Agent": auth_state["user_agent"],
            "Cookie": build_cookie_header(cookies),
        }

        info_url = f"{API_BASE_URL}/get-info-report/{args.report_id}?{urlencode({'menu_id': args.menu_id})}"
        print(f"Đang gọi metadata report: {info_url}")
        info_response = http_json_request(info_url, method="GET", headers=common_headers)

        month_id = args.month_id.strip()
        if not month_id:
            try:
                month_id = find_value_by_label(info_response, args.month_label)
                print(f"✅ Đã map month-label '{args.month_label}' -> pthang={month_id}")
            except RuntimeError:
                print_candidate_pairs(info_response, args.month_label)
                raise
        else:
            print(f"✅ Sử dụng month-id truyền vào: {month_id}")

        payload = build_export_payload(args, month_id)
        export_url = f"{API_BASE_URL}/get-data-export"
        print(f"Đang gọi export API: {export_url}")
        print(json.dumps(payload, ensure_ascii=False, indent=2))

        export_response = http_json_request(
            export_url,
            method="POST",
            headers=common_headers,
            payload=payload,
        )

        output_path = save_exported_file(export_response, args.output_dir, args.output_name)
        print(f"✅ Đã lưu file: {output_path}")

    finally:
        if browser is not None:
            browser.close()
        if playwright is not None:
            playwright.stop()


if __name__ == "__main__":
    main()
