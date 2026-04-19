# -*- coding: utf-8 -*-
"""Helper gọi report BI của OneBSS và lưu file export binary."""

import json
import re
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from api_transition.onebss_auth import OneBSSSettings


API_TRANSITION_DIR = Path(__file__).resolve().parent
DOWNLOADS_DIR = API_TRANSITION_DIR / "downloads"


def sanitize_filename(filename):
    safe_name = re.sub(r'[\\/:*?"<>|]+', "_", str(filename or "").strip())
    safe_name = re.sub(r"\s+", " ", safe_name).strip(" .")
    return safe_name or "report.xlsx"


def build_onebss_headers(session_headers, include_apikey=False, extra_headers=None):
    headers = dict(session_headers or {})
    if include_apikey:
        headers["apikey"] = OneBSSSettings.DEFAULT_API_KEY
    headers.update(extra_headers or {})
    return headers


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
        raise RuntimeError(f"HTTP {exc.code} khi gọi {url}\nResponse:\n{error_body}") from exc
    except URLError as exc:
        raise RuntimeError(f"Lỗi kết nối khi gọi {url}: {exc}") from exc

    try:
        return json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Response không phải JSON từ {url}:\n{body[:1000]}") from exc


def http_binary_request(url, method="GET", headers=None, payload=None, timeout=120):
    request_headers = dict(headers or {})
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request_headers.setdefault("Content-Type", "application/json")

    req = Request(url=url, data=data, headers=request_headers, method=method.upper())
    try:
        with urlopen(req, timeout=timeout) as response:
            body = response.read()
            response_headers = dict(response.headers.items())
    except HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} khi gọi {url}\nResponse:\n{error_body}") from exc
    except URLError as exc:
        raise RuntimeError(f"Lỗi kết nối khi gọi {url}: {exc}") from exc

    return {
        "body": body,
        "headers": response_headers,
    }


def build_report_parameters_url(report_path, api_base_url=""):
    base_url = (api_base_url or OneBSSSettings.API_BASE_URL).rstrip("/")
    query = urlencode({"report": report_path})
    return f"{base_url}/web-report/report/bi/parameters?{query}"


def get_report_parameters(report_path, headers, api_base_url="", timeout=120):
    url = build_report_parameters_url(report_path, api_base_url=api_base_url)
    print(f"Đang gọi metadata report OneBSS: {url}")
    effective_headers = build_onebss_headers(headers, include_apikey=True)
    return http_json_request(url, method="GET", headers=effective_headers, timeout=timeout)


def refresh_report_parameters(report_path, parameter_items, headers, api_base_url="", timeout=120):
    base_url = (api_base_url or OneBSSSettings.API_BASE_URL).rstrip("/")
    url = f"{base_url}/web-report/report/bi/parameters"
    payload = {
        "report": report_path,
        "parameterNameValues": {
            "listOfParamNameValues": {
                "item": parameter_items,
            }
        },
    }
    print(f"Đang refresh metadata report OneBSS: {url}")
    return http_json_request(url, method="POST", headers=headers, payload=payload, timeout=timeout)


def prepare_parameter_item(parameter_definition, value):
    item = dict(parameter_definition)
    normalized_values = value if isinstance(value, list) else [value]
    normalized_values = [str(v) for v in normalized_values]

    if item.get("multiValuesAllowed"):
        item["defaultValue"] = normalized_values
        item["temp"] = normalized_values
        item["values"] = {"item": normalized_values}
    else:
        normalized_value = normalized_values[0] if normalized_values else ""
        item["defaultValue"] = normalized_value
        item["values"] = {"item": normalized_value}

    return {
        "defaultValue": item.get("defaultValue"),
        "temp": item.get("temp", []),
        "name": item["name"],
        "values": item["values"],
    }


def fill_parameter_values(parameter_items, overrides):
    override_map = {str(key): value for key, value in (overrides or {}).items()}
    filled_items = []

    for item in parameter_items:
        item_copy = json.loads(json.dumps(item, ensure_ascii=False))
        name = str(item_copy.get("name", ""))
        raw_override = override_map.get(name)

        if raw_override is not None:
            values = raw_override if isinstance(raw_override, list) else [raw_override]
            normalized_values = [str(v) for v in values]
            if item_copy.get("multiValuesAllowed"):
                item_copy["temp"] = normalized_values
                item_copy["mvalue"] = [{"value": v} for v in normalized_values]
                item_copy["value"] = normalized_values[0] if normalized_values else ""
                item_copy["defaultValue"] = normalized_values
                item_copy["values"] = {"item": normalized_values}
            else:
                selected_value = normalized_values[0] if normalized_values else ""
                item_copy["value"] = selected_value
                item_copy["defaultValue"] = selected_value
                item_copy["mvalue"] = [{"value": selected_value}] if selected_value else []
                item_copy["values"] = {"item": item_copy.get("values", {}).get("item", [])}
        else:
            default_value = item_copy.get("defaultValue")
            if item_copy.get("multiValuesAllowed"):
                if isinstance(default_value, list):
                    current_values = [str(v) for v in default_value]
                else:
                    current_values = []
                item_copy["temp"] = item_copy.get("temp", current_values)
                item_copy["mvalue"] = item_copy.get("mvalue", [{"value": v} for v in current_values])
            else:
                if default_value is None:
                    default_value = ""
                default_value = str(default_value)
                item_copy["value"] = item_copy.get("value", default_value)
                item_copy["mvalue"] = item_copy.get("mvalue", ([{"value": default_value}] if default_value != "" else []))

        if "options" not in item_copy:
            labels = item_copy.get("lovLabels", {}).get("item", [])
            values = item_copy.get("values", {}).get("item", [])
            if not isinstance(labels, list):
                labels = [labels]
            if not isinstance(values, list):
                values = [values]
            item_copy["options"] = [
                {"text": str(text), "id": str(identifier)}
                for text, identifier in zip(labels, values)
            ]

        filled_items.append(item_copy)

    return filled_items


def build_run_v3_payload(baocao_id, report_path, items, file_name="", export_type="xlsx", multiselect=1):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    effective_name = file_name or f"{report_path.rsplit('/', 1)[-1]}_{timestamp}.{export_type}"
    return {
        "baocao_id": baocao_id,
        "report": report_path,
        "type": export_type,
        "multiselect": multiselect,
        "file_name": effective_name,
        "items": items,
    }


def run_report_export(payload, headers, api_base_url="", timeout=120):
    base_url = (api_base_url or OneBSSSettings.API_BASE_URL).rstrip("/")
    url = f"{base_url}/web-report/report/bi/run_v3"
    print(f"Đang export report OneBSS: {url}")
    effective_headers = build_onebss_headers(headers, include_apikey=True)
    return http_binary_request(url, method="POST", headers=effective_headers, payload=payload, timeout=timeout)


def save_binary_export_file(export_response, output_dir="", output_name=""):
    output_root = Path(output_dir or DOWNLOADS_DIR / "onebss")
    output_root.mkdir(parents=True, exist_ok=True)
    file_name = output_name or "report.xlsx"
    output_path = output_root / sanitize_filename(file_name)
    output_path.write_bytes(export_response["body"])
    return output_path
