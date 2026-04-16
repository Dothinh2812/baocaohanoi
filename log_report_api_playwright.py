# -*- coding: utf-8 -*-
"""
Log toàn bộ request/response tới baocaobe.myhanoi.vn/report-api bằng Playwright.

Mục đích:
- Đăng nhập hệ thống báo cáo như luồng hiện tại
- Giữ trình duyệt mở để thao tác tay
- Ghi lại request/response chi tiết phục vụ bóc API export

Ví dụ:
    python log_report_api_playwright.py
    python log_report_api_playwright.py --url "https://baocao.hanoi.vnpt.vn/report/report-info?id=534964&menu_id=535020"
    python log_report_api_playwright.py --host baocaobe.myhanoi.vn/report-api --body-limit 12000
"""

import argparse
import json
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from config import Config
from login import read_otp_from_file


DEFAULT_FILTER = "baocaobe.myhanoi.vn/report-api"


def parse_args():
    parser = argparse.ArgumentParser(
        description="Log request/response report-api trong lúc thao tác tay trên website báo cáo."
    )
    parser.add_argument(
        "--host",
        default=DEFAULT_FILTER,
        help="Chuỗi cần match trong URL request/response.",
    )
    parser.add_argument(
        "--url",
        default="",
        help="URL tùy chọn để tự động mở sau khi đăng nhập thành công.",
    )
    parser.add_argument(
        "--body-limit",
        type=int,
        default=8000,
        help="Số ký tự tối đa lưu cho body/response preview.",
    )
    parser.add_argument(
        "--output-dir",
        default="logs",
        help="Thư mục lưu file JSONL.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Chạy không giao diện. Mặc định là mở browser để thao tác tay.",
    )
    return parser.parse_args()


def ensure_text_preview(raw_value, limit):
    if raw_value is None:
        return ""
    if isinstance(raw_value, bytes):
        text = raw_value.decode("utf-8", errors="replace")
    else:
        text = str(raw_value)
    if len(text) <= limit:
        return text
    return f"{text[:limit]}\n... [truncated {len(text) - limit} chars]"


def mask_sensitive_headers(headers):
    masked = {}
    sensitive_keys = {"authorization", "cookie", "set-cookie", "x-csrf-token"}
    for key, value in headers.items():
        if key.lower() in sensitive_keys:
            masked[key] = "***REDACTED***"
        else:
            masked[key] = value
    return masked


def build_output_path(output_dir):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    return output_root / f"report_api_capture_{timestamp}.jsonl"


def login_with_existing_flow(headless=False):
    print(f"=== Bắt đầu đăng nhập vào {Config.BAOCAO_URL} ===")

    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context(accept_downloads=Config.ACCEPT_DOWNLOADS)
    page = context.new_page()

    print("Đang truy cập trang đăng nhập...")
    page.goto(Config.BAOCAO_URL, timeout=Config.PAGE_LOAD_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)

    print(f"Đang điền username: {Config.BAOCAO_USERNAME}")
    username_field = page.locator('//*[@id="username"]')
    username_field.wait_for(state="visible", timeout=30000)
    username_field.fill(Config.BAOCAO_USERNAME)
    time.sleep(1)

    print("Đang điền password...")
    password_field = page.locator('//*[@id="password"]')
    password_field.wait_for(state="visible", timeout=30000)
    password_field.fill(Config.BAOCAO_PASSWORD)
    time.sleep(1)

    print("Đang click button Đăng nhập...")
    login_button = page.locator('//*[@id="fm1"]/section/button')
    login_button.wait_for(state="visible", timeout=30000)
    login_button.click()
    time.sleep(3)

    print("Đang đợi trường nhập OTP...")
    otp_field = page.locator('//*[@id="passOTP"]')
    otp_field.wait_for(state="visible", timeout=30000)

    otp_code = read_otp_from_file()
    if otp_code is None:
        print("❌ Không thể đọc OTP tự động.")
        print("⏸️  Vui lòng nhập OTP thủ công rồi xác nhận trên trình duyệt.")
        print("⏸️  Script sẽ chờ 20 giây để hoàn tất bước đăng nhập...")
        time.sleep(20)
    else:
        print(f"Đang điền OTP: {otp_code}")
        otp_field.fill(otp_code)
        time.sleep(1)

        print("Đang click button xác nhận OTP...")
        otp_confirm_button = page.locator('//*[@id="loginForm"]/div[1]/button')
        otp_confirm_button.wait_for(state="visible", timeout=30000)
        otp_confirm_button.click()
        time.sleep(5)

    page.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)
    print("✅ Đăng nhập thành công")
    print(f"📝 Cookies hiện có: {len(context.cookies())}")
    return playwright, browser, context, page


def main():
    args = parse_args()
    Config.validate()

    output_path = build_output_path(args.output_dir)
    request_seq = {"value": 0}
    request_ids = {}

    with output_path.open("a", encoding="utf-8") as log_file:
        def write_event(record):
            log_file.write(json.dumps(record, ensure_ascii=False) + "\n")
            log_file.flush()

        def on_request(request):
            if args.host not in request.url:
                return

            request_seq["value"] += 1
            seq = request_seq["value"]
            request_ids[id(request)] = seq

            post_data = request.post_data
            record = {
                "timestamp": datetime.now().isoformat(),
                "seq": seq,
                "event": "request",
                "method": request.method,
                "url": request.url,
                "resource_type": request.resource_type,
                "headers": mask_sensitive_headers(request.headers),
                "post_data": ensure_text_preview(post_data, args.body_limit),
            }
            write_event(record)

            print(f"\n=== REQUEST #{seq} ===")
            print(f"{request.method} {request.url}")
            if post_data:
                print(ensure_text_preview(post_data, args.body_limit))

        def on_response(response):
            request = response.request
            if args.host not in response.url:
                return

            seq = request_ids.get(id(request), 0)
            body_preview = ""
            try:
                content_type = response.headers.get("content-type", "")
                body = response.body()
                if "application/json" in content_type or "text" in content_type or not content_type:
                    body_preview = ensure_text_preview(body, args.body_limit)
                else:
                    body_preview = f"[binary body omitted] content-type={content_type} bytes={len(body)}"
            except Exception as exc:
                body_preview = f"[unable to read response body: {exc}]"

            record = {
                "timestamp": datetime.now().isoformat(),
                "seq": seq,
                "event": "response",
                "status": response.status,
                "url": response.url,
                "headers": mask_sensitive_headers(response.headers),
                "body_preview": body_preview,
            }
            write_event(record)

            print(f"--- RESPONSE #{seq} ---")
            print(f"{response.status} {response.url}")
            if body_preview:
                print(body_preview)

        playwright = browser = context = page = None
        try:
            playwright, browser, context, page = login_with_existing_flow(
                headless=args.headless
            )
            context.on("request", on_request)
            context.on("response", on_response)

            if args.url:
                print(f"Đang mở URL theo tham số --url: {args.url}")
                page.goto(args.url, timeout=Config.PAGE_LOAD_TIMEOUT)
                page.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)

            print("\n" + "=" * 80)
            print("ĐÃ BẬT LOG NETWORK")
            print("=" * 80)
            print(f"Filter URL contains : {args.host}")
            print(f"Log file           : {output_path}")
            print("Tiếp theo:")
            print("1. Thao tác tay trên trình duyệt.")
            print("2. Chạy báo cáo, Xuất Excel, 2.Tất cả dữ liệu / 3.Database.")
            print("3. Theo dõi log trên terminal hoặc mở file JSONL sau khi xong.")
            print("Nhấn Ctrl+C để dừng.")
            print("=" * 80 + "\n")

            while True:
                page.wait_for_timeout(1000)

        except KeyboardInterrupt:
            print("\n⏹️  Đã dừng theo yêu cầu người dùng.")
            print(f"📁 Log đã lưu tại: {output_path}")
        finally:
            if browser is not None:
                browser.close()
            if playwright is not None:
                playwright.stop()


if __name__ == "__main__":
    main()
