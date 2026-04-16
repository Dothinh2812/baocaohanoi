# -*- coding: utf-8 -*-
"""Login và bắt auth header cho report-api."""

import os
import re
import time

from playwright.sync_api import sync_playwright

from api_transition.settings import Settings


def read_otp_from_file():
    file_path = Settings.OTP_FILE_PATH
    max_age_seconds = Settings.OTP_MAX_AGE_SECONDS

    for retry_count in range(10):
        if os.path.exists(file_path):
            file_time = os.path.getmtime(file_path)
            time_diff = time.time() - file_time
            if time_diff <= max_age_seconds:
                with open(file_path, "r", encoding="utf-8") as handle:
                    content = handle.read()
                otp_match = re.search(r"\b\d{6}\b", content)
                if otp_match:
                    otp_code = otp_match.group(0)
                    try:
                        with open(file_path, "w", encoding="utf-8") as handle:
                            handle.write("")
                    except OSError:
                        pass
                    return otp_code
        if retry_count < 9:
            time.sleep(2)
    return None


def login(headless=False):
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context(accept_downloads=Settings.ACCEPT_DOWNLOADS)
    page = context.new_page()

    print(f"=== Đăng nhập {Settings.BAOCAO_URL} ===")
    page.goto(Settings.BAOCAO_URL, timeout=Settings.PAGE_LOAD_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=Settings.PAGE_LOAD_TIMEOUT)

    username_field = page.locator('//*[@id="username"]')
    username_field.wait_for(state="visible", timeout=30000)
    username_field.fill(Settings.BAOCAO_USERNAME)

    password_field = page.locator('//*[@id="password"]')
    password_field.wait_for(state="visible", timeout=30000)
    password_field.fill(Settings.BAOCAO_PASSWORD)

    login_button = page.locator('//*[@id="fm1"]/section/button')
    login_button.wait_for(state="visible", timeout=30000)
    login_button.click()
    time.sleep(3)

    otp_field = page.locator('//*[@id="passOTP"]')
    otp_field.wait_for(state="visible", timeout=30000)

    otp_code = read_otp_from_file()
    if otp_code is None:
        print("Không đọc được OTP tự động. Vui lòng nhập thủ công.")
        time.sleep(20)
    else:
        otp_field.fill(otp_code)
        otp_confirm_button = page.locator('//*[@id="loginForm"]/div[1]/button')
        otp_confirm_button.wait_for(state="visible", timeout=30000)
        otp_confirm_button.click()
        time.sleep(5)

    page.wait_for_load_state("networkidle", timeout=Settings.PAGE_LOAD_TIMEOUT)
    print(f"✅ Đăng nhập thành công, cookies={len(context.cookies())}")
    return playwright, browser, context, page


def capture_authorization(page, report_page_url, timeout_seconds=30):
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
            state["referer"] = request.headers.get("referer", Settings.DEFAULT_REFERER)

    page.context.on("request", on_request)

    print(f"Đang mở trang report để bắt Authorization: {report_page_url}")
    page.goto(report_page_url, wait_until="networkidle", timeout=Settings.PAGE_LOAD_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=Settings.PAGE_LOAD_TIMEOUT)

    started = time.time()
    while time.time() - started < timeout_seconds:
        if state.get("authorization"):
            print("✅ Đã bắt được Authorization header")
            return state
        page.wait_for_timeout(500)

    raise RuntimeError("Không bắt được Authorization header từ request /report-api/.")
