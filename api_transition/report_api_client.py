# -*- coding: utf-8 -*-
"""Helper gọi report-api và xử lý response export."""

import base64
import json
import re
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from api_transition.settings import Settings


def build_cookie_header(cookies):
    return "; ".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in cookies
        if cookie.get("name") and cookie.get("value")
    )


def make_common_headers(auth_state, cookies):
    return {
        "Authorization": auth_state["authorization"],
        "Accept": auth_state["accept"],
        "Referer": auth_state["referer"],
        "User-Agent": auth_state["user_agent"],
        "Cookie": build_cookie_header(cookies),
    }


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
        candidate_pairs = [
            ("text", "value"),
            ("Text", "Value"),
            ("TEXT", "VALUE"),
            ("TEXT", "ID"),
            ("text", "id"),
            ("Text", "ID"),
        ]
        for text_key, value_key in candidate_pairs:
            if text_key in node and value_key in node:
                results.append(
                    {
                        "text": node.get(text_key),
                        "value": node.get(value_key),
                        "path": path,
                    }
                )
                break
        for key, value in node.items():
            results.extend(collect_text_value_pairs(value, f"{path}.{key}"))
    elif isinstance(node, list):
        for index, item in enumerate(node):
            results.extend(collect_text_value_pairs(item, f"{path}[{index}]"))
    return results


def normalize_spaces(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().casefold()


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

    raise RuntimeError(f"Không tìm thấy option có text '{label}' trong get-info-report.")


def print_candidate_pairs(info_payload, label, limit=20):
    print("\nKhông map được month-label. Một số option text/value gần đúng:")
    pairs = collect_text_value_pairs(info_payload)
    search_tokens = [
        token
        for token in re.split(r"[\s/()-]+", normalize_spaces(label))
        if len(token) >= 2
    ]
    matches = []
    seen = set()
    for pair in pairs:
        normalized_text = normalize_spaces(pair["text"])
        if search_tokens and not any(token in normalized_text for token in search_tokens):
            continue
        key = (str(pair["text"]), str(pair["value"]))
        if key in seen:
            continue
        seen.add(key)
        matches.append(pair)
    if not matches:
        matches = pairs[:limit]
    for index, pair in enumerate(matches[:limit], start=1):
        print(f"{index}. text={pair['text']!r} value={pair['value']!r} path={pair['path']}")


def build_report_page_url(report_id, menu_id):
    return f"{Settings.BAOCAO_BASE_URL}/report/report-info?id={report_id}&menu_id={menu_id}"


def get_info_report(report_id, menu_id, headers):
    url = f"{Settings.API_BASE_URL}/get-info-report/{report_id}?{urlencode({'menu_id': menu_id})}"
    print(f"Đang gọi metadata report: {url}")
    return http_json_request(url, method="GET", headers=headers)


def export_report(headers, payload, timeout=120):
    url = f"{Settings.API_BASE_URL}/get-data-export"
    print(f"Đang gọi export API: {url}")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return http_json_request(url, method="POST", headers=headers, payload=payload, timeout=timeout)


def sanitize_filename(filename):
    safe_name = re.sub(r'[\\/:*?"<>|]+', "_", str(filename or "").strip())
    safe_name = re.sub(r"\s+", " ", safe_name).strip(" .")
    return safe_name or "report.xlsx"


def save_export_file(export_response, output_dir, output_name=""):
    if "FileContents" not in export_response:
        raise RuntimeError(
            f"Response export không chứa FileContents:\n{json.dumps(export_response, ensure_ascii=False)[:1000]}"
        )

    file_name = output_name or export_response.get("FileDownloadName") or "report.xlsx"
    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    output_path = output_root / sanitize_filename(file_name)
    output_path.write_bytes(base64.b64decode(export_response["FileContents"]))
    return output_path
