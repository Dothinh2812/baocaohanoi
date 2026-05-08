# -*- coding: utf-8 -*-
"""
Script tải báo cáo I1.5 (K1 và K2) cho một ngày.

Mặc định tải ngày T-1, với T là ngày hiện tại trên máy chạy script.

Cách dùng:
    python3 download_daily.py
    python3 download_daily.py --date 23/04/2026
    python3 download_daily.py --year 2026 --month 4 --day 23
    python3 download_daily.py --year 2026 --month 4

Nếu chỉ chọn --year/--month mà không chọn --day, script dùng ngày trong tháng
của T-1. Ví dụ hôm nay 16/05/2026, --month 4 sẽ tải 15/04/2026.
"""

import argparse
import calendar
from datetime import date, datetime, timedelta
from pathlib import Path

from auth import capture_authorization, login
from downloaders import download_report_i15_api, download_report_i15_k2_api
from report_api_client import make_common_headers

# ──────────────────────────────────────────────
#  CẤU HÌNH – có thể chỉnh tại đây trước khi chạy
# ──────────────────────────────────────────────
REPORT_YEAR = None       # None = dùng năm của T-1; ví dụ: 2026
REPORT_MONTH = None      # None = dùng tháng của T-1; ví dụ: 4
REPORT_DAY = None        # None = dùng ngày của T-1; ví dụ: 23
UNIT_ID = "284656"       # ID đơn vị
HEADED = False           # True = hiện trình duyệt, False = chạy nền
API_TIMEOUT = 180
# ──────────────────────────────────────────────

THIS_DIR = Path(__file__).resolve().parent
K1_DIR = THIS_DIR / "K1_month"
K2_DIR = THIS_DIR / "K2_month"

REPORT_PAGE_URL_I15 = (
    "https://baocao.hanoi.vnpt.vn/report/report-info"
    "?id=509918&menu_id=509934"
)


def ensure_dirs():
    K1_DIR.mkdir(parents=True, exist_ok=True)
    K2_DIR.mkdir(parents=True, exist_ok=True)


def fmt_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def parse_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            "Ngày phải theo định dạng DD/MM/YYYY, ví dụ 23/04/2026."
        ) from exc


def resolve_target_date(
    today=None,
    report_year=None,
    report_month=None,
    report_day=None,
    explicit_date=None,
):
    if explicit_date is not None:
        return explicit_date

    current_day = today or date.today()
    yesterday = current_day - timedelta(days=1)

    year = report_year if report_year is not None else REPORT_YEAR
    month = report_month if report_month is not None else REPORT_MONTH
    day = report_day if report_day is not None else REPORT_DAY

    year = year if year is not None else yesterday.year
    month = month if month is not None else yesterday.month
    day = day if day is not None else yesterday.day

    _, last_day = calendar.monthrange(year, month)
    if day < 1 or day > last_day:
        raise ValueError(
            f"Ngày {day} không hợp lệ cho tháng {month:02d}/{year}. "
            "Hãy chọn REPORT_DAY hoặc truyền --day phù hợp."
        )
    return date(year, month, day)


def build_session():
    playwright, browser, context, page = login(headless=not HEADED)
    auth_state = capture_authorization(page, REPORT_PAGE_URL_I15)
    headers = make_common_headers(auth_state, context.cookies())
    return {
        "playwright": playwright,
        "browser": browser,
        "context": context,
        "page": page,
        "headers": headers,
        "api_timeout": API_TIMEOUT,
    }


def close_session(session: dict):
    try:
        session["browser"].close()
    except Exception:
        pass
    try:
        session["playwright"].stop()
    except Exception:
        pass


def download_one_day(target_day: date):
    ensure_dirs()
    date_str = fmt_date(target_day)
    output_name = f"{target_day.day}.xlsx"

    print(f"\n{'=' * 60}")
    print(f"  Tải báo cáo I1.5 K1 & K2 - ngày {date_str}")
    print(f"  K1 -> {K1_DIR}")
    print(f"  K2 -> {K2_DIR}")
    print(f"{'=' * 60}\n")

    print("Dang dang nhap...")
    session = build_session()
    print("Dang nhap thanh cong.\n")

    try:
        saved = download_report_i15_api(
            start_date=date_str,
            end_date=date_str,
            unit_id=UNIT_ID,
            headed=HEADED,
            output_dir=str(K1_DIR),
            session=session,
        )
        target_path = K1_DIR / output_name
        saved_path = Path(saved)
        if saved_path.resolve() != target_path.resolve():
            saved_path.rename(target_path)
        print(f"K1 -> {target_path.name}")

        saved = download_report_i15_k2_api(
            start_date=date_str,
            end_date=date_str,
            unit_id=UNIT_ID,
            headed=HEADED,
            output_dir=str(K2_DIR),
            session=session,
        )
        target_path = K2_DIR / output_name
        saved_path = Path(saved)
        if saved_path.resolve() != target_path.resolve():
            saved_path.rename(target_path)
        print(f"K2 -> {target_path.name}")
    finally:
        close_session(session)

    print(f"\nHoan thanh ngay {date_str}.\n")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Tải báo cáo I1.5 K1/K2 cho một ngày, mặc định là T-1."
    )
    parser.add_argument("--date", type=parse_date, help="Ngày cần tải, dạng DD/MM/YYYY.")
    parser.add_argument("--year", type=int, help="Năm báo cáo.")
    parser.add_argument("--month", type=int, choices=range(1, 13), help="Tháng báo cáo.")
    parser.add_argument("--day", type=int, help="Ngày báo cáo.")
    return parser.parse_args()


def run():
    args = parse_args()
    target_day = resolve_target_date(
        report_year=args.year,
        report_month=args.month,
        report_day=args.day,
        explicit_date=args.date,
    )
    download_one_day(target_day)


if __name__ == "__main__":
    run()
