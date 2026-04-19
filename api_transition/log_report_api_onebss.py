# -*- coding: utf-8 -*-
"""
Log toàn bộ request/response tới api của onebss.vnpt.vn bằng Playwright.
Mục đích:
- Đăng nhập hệ thống OneBSS
- Giữ trình duyệt mở để thao tác tay
- Ghi lại request/response chi tiết để bắt các API nội bộ
"""

import argparse
import json
import time
import os
import re
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# Tải cấu hình từ .env
load_dotenv()

DEFAULT_FILTER = "onebss.vnpt.vn"

# Sử dụng biến môi trường (như cách app.py đang lấy gián tiếp từ config hoặc .env)
ONEBSS_USERNAME = os.getenv("ONEBSS_USERNAME", "")
ONEBSS_PASSWORD = os.getenv("ONEBSS_PASSWORD", "")
ONEBSS_URL = os.getenv("ONEBSS_URL", "https://onebss.vnpt.vn")
OTP_FILE_PATH = os.getenv("OTP_FILE_PATH", "/home/vtst/otp/otp_logs.txt")
PAGE_LOAD_TIMEOUT = int(os.getenv("PAGE_LOAD_TIMEOUT", "60000"))


def parse_args():
    parser = argparse.ArgumentParser(
        description="Log request/response trong lúc thao tác tay trên website OneBSS."
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
    return output_root / f"onebss_api_capture_{timestamp}.jsonl"


def login_with_existing_flow(headless=False):
    print(f"=== Bắt đầu đăng nhập vào {ONEBSS_URL} ===")
    if not ONEBSS_USERNAME or not ONEBSS_PASSWORD:
        print("⚠️  CẢNH BÁO: ONEBSS_USERNAME hoặc ONEBSS_PASSWORD đang trống trong biến môi trường!")

    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    print("Đang truy cập trang đăng nhập...")
    page.goto(ONEBSS_URL, timeout=PAGE_LOAD_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)

    # Các selector dựa trên hàm perform_login trong app.py
    print(f"Đang điền username: {ONEBSS_USERNAME}")
    username_field = page.locator('//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[1]/div/input')
    username_field.wait_for(state="visible", timeout=45000)
    username_field.fill(ONEBSS_USERNAME)
    time.sleep(1)

    print("Đang điền password...")
    password_field = page.locator('//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[2]/div/input')
    password_field.wait_for(state="visible", timeout=45000)
    password_field.fill(ONEBSS_PASSWORD)
    time.sleep(1)

    print("Check remember me...")
    checkbox = page.locator('//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[3]/div/input')
    checkbox.wait_for(state="visible", timeout=30000)
    checkbox.check()
    time.sleep(1)

    print("Đang click button đăng nhập...")
    login_button = page.locator('//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[4]/button')
    login_button.wait_for(state="visible", timeout=30000)
    login_button.click()
    time.sleep(3)
    
    page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)
    print("Kiểm tra và xử lý OTP...")

    # Wait for OTP input field if required
    otp_field = page.locator('//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[1]/div/input')
    try:
        # Chờ tối đa 10s để xem có bị yêu cầu OTP không, nếu không có thì bỏ qua
        otp_field.wait_for(state="visible", timeout=10000)
        
        file_path = OTP_FILE_PATH
        print(f"📂 Được yêu cầu nhập OTP. Đường dẫn file OTP: {file_path}")
        max_retries = 30
        retry_count = 0
        otp_code = None
        OTP_MAX_AGE = 30

        while retry_count < max_retries:
            if os.path.exists(file_path):
                file_time = os.path.getmtime(file_path)
                current_time = time.time()
                time_diff = current_time - file_time

                if time_diff <= OTP_MAX_AGE:
                    with open(file_path, 'r') as f:
                        content = f.read()
                        otp_match = re.search(r'\b\d{6}\b', content)
                        if otp_match:
                            otp_code = otp_match.group(0)
                            print(f"✅ Found OTP code: {otp_code}")
                            break
                else:
                    print(f"⏳ OTP file cũ ({time_diff:.1f}s > {OTP_MAX_AGE}s), chờ mã mới... [{retry_count + 1}/{max_retries}]")
            else:
                print(f"⏳ Chờ file OTP... [{retry_count + 1}/{max_retries}]")
            
            retry_count += 1
            if retry_count < max_retries:
                time.sleep(2)
        
        if otp_code is None:
            print("❌ Không tìm thấy OTP mới sau nhiều lần thử.")
            print("⏸️  Vui lòng nhập OTP thủ công rồi xác nhận trên trình duyệt (script đang chờ 30 giây).")
            time.sleep(30)
        else:
            print(f"Đang điền OTP: {otp_code}")
            otp_field.fill(otp_code)
            time.sleep(1)

            print("Click xác nhận OTP...")
            confirm_button = page.locator('//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[2]/button[2]')
            confirm_button.wait_for(state="visible", timeout=30000)
            confirm_button.click()
            time.sleep(5)
            
            # Xóa OTP
            try:
                if os.path.exists(OTP_FILE_PATH):
                    with open(OTP_FILE_PATH, 'w') as f:
                        f.write('')
                    print(f"✅ Đã xóa mã OTP trong file {OTP_FILE_PATH}")
            except Exception as e:
                print(f"⚠️ Không thể xóa mã OTP: {e}")

    except Exception:
        print("👉 Không có ô nhập OTP xuất hiện, có thể do đã đăng nhập thẳng hoặc lưu máy.")

    page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)
    print("✅ Đăng nhập hoàn tất")
    print(f"📝 Cookies hiện có: {len(context.cookies())}")
    return playwright, browser, context, page


def main():
    args = parse_args()

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
                page.goto(args.url, timeout=PAGE_LOAD_TIMEOUT)
                page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)

            print("\n" + "=" * 80)
            print("ĐÃ BẬT LOG NETWORK CHUẨN BỊ CHO THAO TÁC ONEBSS")
            print("=" * 80)
            print(f"Filter URL contains : {args.host}")
            print(f"Log file           : {output_path}")
            print("Tiếp theo:")
            print("1. Thao tác tay trên trình duyệt OneBSS đang mở.")
            print("2. Theo dõi log trên terminal hoặc kiểm tra thư mục 'logs' sau khi xong.")
            print("Nhấn Ctrl+C để dừng script và đóng trình duyệt.")
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
