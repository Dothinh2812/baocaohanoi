# -*- coding: utf-8 -*-
"""
Script tải báo cáo I1.5 (K1 và K2) cho từng ngày trong tháng.

Cách dùng:
    python3 download_month.py          (chạy trực tiếp)
    hoặc chỉnh REPORT_YEAR / REPORT_MONTH ở phần cấu hình bên dưới rồi chạy.

File được lưu tên là: 1.xlsx, 2.xlsx, ..., 30.xlsx (hoặc 31.xlsx / 28.xlsx / 29.xlsx)
vào 2 thư mục:
    k1k2MonthDownload/K1_month/
    k1k2MonthDownload/K2_month/
"""

import calendar
import time
from datetime import date
from pathlib import Path

from auth import capture_authorization, login
from report_api_client import make_common_headers
from downloaders import (
    download_report_i15_api,
    download_report_i15_k2_api,
)

# ──────────────────────────────────────────────
#  CẤU HÌNH – chỉnh tại đây trước khi chạy
# ──────────────────────────────────────────────
REPORT_YEAR = 2026
REPORT_MONTH = 4          # tháng cần tải (1–12)
UNIT_ID = "284656"        # ID đơn vị
HEADED = False            # True = hiện trình duyệt, False = chạy nền
DELAY_BETWEEN_DAYS = 2    # giây nghỉ giữa 2 ngày liên tiếp (tránh quá tải server)
# ──────────────────────────────────────────────

THIS_DIR = Path(__file__).resolve().parent
K1_DIR = THIS_DIR / "K1_month"
K2_DIR = THIS_DIR / "K2_month"


def ensure_dirs():
    K1_DIR.mkdir(parents=True, exist_ok=True)
    K2_DIR.mkdir(parents=True, exist_ok=True)


def fmt_date(d: date) -> str:
    """Chuyển date → chuỗi DD/MM/YYYY."""
    return d.strftime("%d/%m/%Y")


def days_in_month(year: int, month: int):
    """Trả về danh sách tất cả các ngày trong tháng."""
    _, last_day = calendar.monthrange(year, month)
    return [date(year, month, day) for day in range(1, last_day + 1)]


def build_session(recipe_report_page_url: str):
    """
    Đăng nhập và bắt Authorization header.
    Trả về dict session để tái sử dụng cho nhiều lần tải.
    """
    playwright, browser, context, page = login(headless=not HEADED)
    auth_state = capture_authorization(page, recipe_report_page_url)
    headers = make_common_headers(auth_state, context.cookies())
    return {
        "playwright": playwright,
        "browser": browser,
        "context": context,
        "page": page,
        "headers": headers,
        "api_timeout": 180,
    }


def close_session(session: dict):
    """Đóng trình duyệt sau khi tải xong."""
    try:
        session["browser"].close()
    except Exception:
        pass
    try:
        session["playwright"].stop()
    except Exception:
        pass


def run():
    ensure_dirs()

    all_days = days_in_month(REPORT_YEAR, REPORT_MONTH)
    total = len(all_days)
    print(f"\n{'='*60}")
    print(f"  Tải báo cáo I1.5 K1 & K2 – {REPORT_MONTH:02d}/{REPORT_YEAR}")
    print(f"  Tổng số ngày: {total}")
    print(f"  K1 → {K1_DIR}")
    print(f"  K2 → {K2_DIR}")
    print(f"{'='*60}\n")

    # Trang báo cáo dùng để bắt Authorization (lấy từ recipe i15)
    # Nếu recipe đổi URL, cập nhật dòng này cho phù hợp.
    REPORT_PAGE_URL_I15 = (
        "https://baocao.hanoi.vnpt.vn/report/report-info"
        "?id=509918&menu_id=509934"
    )

    # ── Đăng nhập một lần duy nhất ──
    print("🔐 Đang đăng nhập...")
    session = build_session(REPORT_PAGE_URL_I15)
    print("✅ Đăng nhập thành công.\n")

    errors_k1 = []
    errors_k2 = []

    try:
        for i, day in enumerate(all_days, start=1):
            date_str = fmt_date(day)
            day_num = day.day
            output_name = f"{day_num}.xlsx"

            print(f"[{i:2d}/{total}] 📅 Ngày {date_str}")

            # ── K1 ──
            k1_output = str(K1_DIR / output_name)
            try:
                saved = download_report_i15_api(
                    start_date=date_str,
                    end_date=date_str,
                    unit_id=UNIT_ID,
                    headed=HEADED,
                    output_dir=str(K1_DIR),
                    session=session,
                )
                # Đổi tên file về <ngay>.xlsx nếu hàm lưu tên mặc định
                saved_path = Path(saved)
                target_path = K1_DIR / output_name
                if saved_path.resolve() != target_path.resolve():
                    saved_path.rename(target_path)
                print(f"         ✅ K1 → {target_path.name}")
            except Exception as exc:
                print(f"         ❌ K1 lỗi: {exc}")
                errors_k1.append((date_str, str(exc)))

            # ── K2 ──
            try:
                saved = download_report_i15_k2_api(
                    start_date=date_str,
                    end_date=date_str,
                    unit_id=UNIT_ID,
                    headed=HEADED,
                    output_dir=str(K2_DIR),
                    session=session,
                )
                saved_path = Path(saved)
                target_path = K2_DIR / output_name
                if saved_path.resolve() != target_path.resolve():
                    saved_path.rename(target_path)
                print(f"         ✅ K2 → {target_path.name}")
            except Exception as exc:
                print(f"         ❌ K2 lỗi: {exc}")
                errors_k2.append((date_str, str(exc)))

            # Nghỉ giữa các ngày (trừ ngày cuối)
            if i < total and DELAY_BETWEEN_DAYS > 0:
                time.sleep(DELAY_BETWEEN_DAYS)

    finally:
        close_session(session)

    # ── Tổng kết ──
    print(f"\n{'='*60}")
    print("📊 TỔNG KẾT")
    print(f"  K1: {total - len(errors_k1)}/{total} thành công")
    print(f"  K2: {total - len(errors_k2)}/{total} thành công")
    if errors_k1:
        print("\n  ❌ K1 – Các ngày lỗi:")
        for d, err in errors_k1:
            print(f"     {d}: {err}")
    if errors_k2:
        print("\n  ❌ K2 – Các ngày lỗi:")
        for d, err in errors_k2:
            print(f"     {d}: {err}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    run()
