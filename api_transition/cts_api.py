# -*- coding: utf-8 -*-
"""Downloader API riêng cho CTS, độc lập với baocao.hanoi."""

import argparse
import json
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import Config
from cts import login_cts


CTS_BASE_URL = "https://cts.vnpt.vn"
DEFAULT_REPORT_URL = f"{CTS_BASE_URL}/Linetest/Report/GponQualityByUnitvb8362"
DEFAULT_EXPORT_URL = f"{CTS_BASE_URL}/Linetest/Report/RIMSGponQualityDetailByUnitvb8362KemAsync_N34"
DEFAULT_UNIT_LIST_URL = f"{CTS_BASE_URL}/Linetest/Report/RIMSUnitListByAccount_N34"
DOWNLOADS_DIR = Path(__file__).resolve().parent / "downloads" / "cts"


def build_cookie_header(cookies):
    return "; ".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in cookies
        if cookie.get("name") and cookie.get("value")
    )


def sanitize_filename(filename):
    safe_name = re.sub(r'[\\/:*?"<>|]+', "_", str(filename or "").strip())
    safe_name = re.sub(r"\s+", " ", safe_name).strip(" .")
    return safe_name or "cts_report.xlsx"


def parse_date_input(value, field_name):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    for fmt in (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(
        f"{field_name} không đúng định dạng. Hỗ trợ: dd/mm/yyyy, yyyy-mm-dd, mm/dd/yyyy."
    )


def format_display_date(value):
    return value.strftime("%d/%m/%Y")


def format_iso_date(value):
    return value.strftime("%Y-%m-%d")


def build_cts_headers(cookies, referer, user_agent):
    return {
        "Accept": "application/octet-stream, application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Cookie": build_cookie_header(cookies),
        "Origin": CTS_BASE_URL,
        "Referer": referer,
        "User-Agent": user_agent,
        "X-Requested-With": "XMLHttpRequest",
    }


def http_json_request(url, method="GET", headers=None, payload=None, timeout=120):
    request_headers = dict(headers or {})
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request_headers.setdefault("Content-Type", "application/json;charset=UTF-8")

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


def http_binary_request(url, method="POST", headers=None, payload=None, timeout=120):
    request_headers = dict(headers or {})
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request_headers.setdefault("Content-Type", "application/json;charset=UTF-8")

    req = Request(url=url, data=data, headers=request_headers, method=method.upper())
    try:
        with urlopen(req, timeout=timeout) as response:
            body = response.read()
            response_headers = {key: value for key, value in response.info().items()}
            response_url = response.geturl()
    except HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"HTTP {exc.code} khi gọi {url}\nResponse:\n{error_body}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(f"Lỗi kết nối khi gọi {url}: {exc}") from exc

    return body, response_headers, response_url


def create_cts_session(headed=False, report_url=DEFAULT_REPORT_URL):
    """Đăng nhập CTS, mở report URL, gom cookies và headers dùng cho API."""
    previous_headless = Config.BROWSER_HEADLESS
    page = browser = playwright = None

    try:
        Config.BROWSER_HEADLESS = not headed
        page, browser, playwright = login_cts()
    finally:
        Config.BROWSER_HEADLESS = previous_headless

    if page is None:
        raise RuntimeError("Đăng nhập CTS thất bại.")

    page.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)

    context = page.context
    cookies = context.cookies([CTS_BASE_URL])
    user_agent = page.evaluate("() => navigator.userAgent")
    headers = build_cts_headers(cookies, referer=report_url, user_agent=user_agent)

    return {
        "page": page,
        "browser": browser,
        "playwright": playwright,
        "cookies": cookies,
        "headers": headers,
        "report_url": report_url,
        "user_agent": user_agent,
    }


def close_cts_session(session):
    if session is None:
        return

    browser = session.get("browser")
    playwright = session.get("playwright")

    if browser is not None:
        try:
            browser.close()
        except Exception:
            pass
    if playwright is not None:
        try:
            playwright.stop()
        except Exception:
            pass


def get_cts_unit_list(session=None, headed=False):
    own_session = session is None
    try:
        if own_session:
            session = create_cts_session(headed=headed)
        return http_json_request(
            DEFAULT_UNIT_LIST_URL,
            method="GET",
            headers=session["headers"],
            timeout=max(30, Config.DOWNLOAD_TIMEOUT // 1000),
        )
    finally:
        if own_session:
            close_cts_session(session)


def find_cts_unit_id_by_name(unit_name, session=None, headed=False):
    normalized_target = re.sub(r"\s+", " ", str(unit_name or "")).strip().casefold()
    if not normalized_target:
        raise ValueError("unit_name không được để trống.")

    units = get_cts_unit_list(session=session, headed=headed)
    for item in units:
        normalized_name = re.sub(r"\s+", " ", str(item.get("UnitName", ""))).strip().casefold()
        if normalized_name == normalized_target:
            return str(item["UnitID"])

    for item in units:
        normalized_name = re.sub(r"\s+", " ", str(item.get("UnitName", ""))).strip().casefold()
        if normalized_target in normalized_name:
            return str(item["UnitID"])

    raise RuntimeError(f"Không tìm thấy UnitID cho đơn vị '{unit_name}'.")


def build_cts_gpon_quality_payload(
    report_date=None,
    start_date=None,
    end_date=None,
    unit_id="87756",
    exclusive_den_ngay=False,
    loss_max=30,
    loss_ok=27,
    report_type=1,
    search_type=2,
    quarterly=0,
    province_code=1,
):
    """Dựng payload export cho CTS GPON Quality detail."""
    if report_date not in (None, "") and (start_date not in (None, "") or end_date not in (None, "")):
        raise ValueError("Chỉ dùng report_date hoặc cặp start_date/end_date, không truyền cả hai.")

    if report_date in (None, "") and start_date in (None, "") and end_date in (None, ""):
        end_value = date.today() - timedelta(days=1)
        start_value = end_value
    elif report_date not in (None, ""):
        start_value = parse_date_input(report_date, "report_date")
        end_value = start_value
    else:
        start_value = parse_date_input(start_date, "start_date")
        end_value = parse_date_input(end_date or start_date, "end_date")

    if start_value is None or end_value is None:
        raise ValueError("Không xác định được khoảng ngày cho payload CTS.")
    if end_value < start_value:
        raise ValueError("end_date phải lớn hơn hoặc bằng start_date.")

    display_den_ngay = end_value + timedelta(days=1) if exclusive_den_ngay else end_value

    return {
        "searchType": int(search_type),
        "maDonVi": "",
        "tuNgay": format_display_date(start_value),
        "denNgay": format_display_date(display_den_ngay),
        "tuthang": str(start_value.month - 1),
        "denthang": str(end_value.month - 1),
        "UnitID": str(unit_id),
        "Loss_Max": int(loss_max),
        "Loss_Ok": int(loss_ok),
        "ReportType": int(report_type),
        "Quaterly": int(quarterly),
        "unitId": "",
        "Year": int(end_value.year),
        "province_code": "",
        "nam": int(end_value.year),
        "BeginDate": format_iso_date(start_value),
        "endDate": format_iso_date(end_value),
        "ProvinceCode": int(province_code),
    }


def save_binary_file(content, output_dir, output_name):
    output_root = Path(output_dir or DOWNLOADS_DIR)
    output_root.mkdir(parents=True, exist_ok=True)
    output_path = output_root / sanitize_filename(output_name)
    output_path.write_bytes(content)
    return output_path


def download_cts_gpon_quality_detail_api(
    report_date=None,
    start_date=None,
    end_date=None,
    unit_id="87756",
    unit_name="",
    headed=False,
    output_dir="",
    output_name="cts_shc_ngay.xlsx",
    exclusive_den_ngay=False,
    session=None,
):
    """Tải báo cáo CTS GPON Quality detail bằng API binary."""
    own_session = session is None
    try:
        if own_session:
            session = create_cts_session(headed=headed)

        effective_unit_id = str(unit_id or "").strip()
        if unit_name:
            effective_unit_id = find_cts_unit_id_by_name(unit_name, session=session)
        if not effective_unit_id:
            effective_unit_id = "87756"

        payload = build_cts_gpon_quality_payload(
            report_date=report_date,
            start_date=start_date,
            end_date=end_date,
            unit_id=effective_unit_id,
            exclusive_den_ngay=exclusive_den_ngay,
        )

        print("Đang gọi CTS export API:")
        print(DEFAULT_EXPORT_URL)
        print(json.dumps(payload, ensure_ascii=False, indent=2))

        content, response_headers, _ = http_binary_request(
            DEFAULT_EXPORT_URL,
            method="POST",
            headers=session["headers"],
            payload=payload,
            timeout=max(30, Config.DOWNLOAD_TIMEOUT // 1000),
        )

        content_type = str(response_headers.get("Content-Type", "")).lower()
        if "application/json" in content_type:
            try:
                decoded = json.loads(content.decode("utf-8"))
            except Exception:
                decoded = content.decode("utf-8", errors="replace")
            raise RuntimeError(f"CTS export không trả file binary:\n{decoded}")

        return save_binary_file(content, output_dir or DOWNLOADS_DIR, output_name)
    finally:
        if own_session:
            close_cts_session(session)


def parse_args():
    parser = argparse.ArgumentParser(description="Tải báo cáo CTS GPON Quality detail bằng API.")
    parser.add_argument("--date", default="", help="Ngày báo cáo. Hỗ trợ dd/mm/yyyy, yyyy-mm-dd, mm/dd/yyyy.")
    parser.add_argument("--start-date", default="", help="Ngày bắt đầu nếu muốn truyền range.")
    parser.add_argument("--end-date", default="", help="Ngày kết thúc nếu muốn truyền range.")
    parser.add_argument("--unit-id", default="", help="UnitID CTS. Mặc định là Viễn thông Hà Nội nếu không truyền.")
    parser.add_argument("--unit-name", default="", help="Tên đơn vị CTS để tự resolve UnitID.")
    parser.add_argument("--output-dir", default="", help="Thư mục lưu file. Mặc định: api_transition/downloads/cts/")
    parser.add_argument("--output-name", default="cts_shc_ngay.xlsx", help="Tên file đầu ra.")
    parser.add_argument("--headed", action="store_true", help="Mở trình duyệt có giao diện khi login CTS.")
    parser.add_argument(
        "--exclusive-den-ngay",
        action="store_true",
        help="Gửi denNgay = end_date + 1 ngày theo đúng biến thể capture.",
    )
    parser.add_argument("--list-units", action="store_true", help="In danh sách đơn vị CTS rồi thoát.")
    return parser.parse_args()


def main():
    args = parse_args()

    if args.list_units:
        units = get_cts_unit_list(headed=args.headed)
        for item in units:
            print(f"{item.get('UnitID')}\t{item.get('UnitName')}")
        return

    output_path = download_cts_gpon_quality_detail_api(
        report_date=args.date,
        start_date=args.start_date,
        end_date=args.end_date,
        unit_id=args.unit_id,
        unit_name=args.unit_name,
        headed=args.headed,
        output_dir=args.output_dir,
        output_name=args.output_name,
        exclusive_den_ngay=args.exclusive_den_ngay,
    )
    print(f"✅ Đã lưu file: {output_path}")


if __name__ == "__main__":
    main()
