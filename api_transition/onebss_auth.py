# -*- coding: utf-8 -*-
"""Đăng nhập OneBSS và tạo session dùng lại cho các downloader riêng."""

import json
import os
import re
import time
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


ROOT_DIR = Path(__file__).resolve().parent.parent
load_dotenv(ROOT_DIR / ".env")


class OneBSSSettings:
    USERNAME = os.getenv("BAOCAO_USERNAME", "")
    PASSWORD = os.getenv("BAOCAO_PASSWORD", "")

    BASE_URL = os.getenv("ONEBSS_BASE_URL", "https://onebss.vnpt.vn").rstrip("/")
    LOGIN_URL = os.getenv("ONEBSS_LOGIN_URL", BASE_URL)
    DEFAULT_TOKEN_PAGE_URL = os.getenv("ONEBSS_TOKEN_PAGE_URL", BASE_URL)
    REQUEST_HOST_FILTER = os.getenv("ONEBSS_REQUEST_HOST_FILTER", "api-onebss.vnpt.vn")
    API_BASE_URL = os.getenv("ONEBSS_API_BASE_URL", "https://api-onebss.vnpt.vn").rstrip("/")
    DEFAULT_SELECTED_MENU_ID = os.getenv("ONEBSS_SELECTED_MENU_ID", "13177")
    DEFAULT_SELECTED_PATH = os.getenv("ONEBSS_SELECTED_PATH", "/#/report/bi")
    DEFAULT_MAC_ADDRESS = os.getenv("ONEBSS_MAC_ADDRESS", "WEB")
    DEFAULT_TOKEN_ID = os.getenv("ONEBSS_TOKEN_ID", "97388db0-6ce9-11ea-bc55-0242ac130003")
    DEFAULT_API_KEY = os.getenv("ONEBSS_API_KEY", "x")

    OTP_FILE_PATH = os.getenv("OTP_FILE_PATH", "/home/vtst/otp/otp_logs.txt")
    OTP_MAX_AGE_SECONDS = int(os.getenv("OTP_MAX_AGE_SECONDS", "120"))
    OTP_MAX_RETRIES = int(os.getenv("ONEBSS_OTP_MAX_RETRIES", "30"))
    OTP_RETRY_INTERVAL_SECONDS = float(os.getenv("ONEBSS_OTP_RETRY_INTERVAL_SECONDS", "2"))

    PAGE_LOAD_TIMEOUT = int(os.getenv("PAGE_LOAD_TIMEOUT", "60000"))
    ACCEPT_DOWNLOADS = os.getenv("ACCEPT_DOWNLOADS", "True").lower() == "true"

    @classmethod
    def validate(cls):
        errors = []
        if not cls.USERNAME:
            errors.append("BAOCAO_USERNAME không được để trống")
        if not cls.PASSWORD:
            errors.append("BAOCAO_PASSWORD không được để trống")
        if not cls.OTP_FILE_PATH:
            errors.append("OTP_FILE_PATH không được để trống")
        if errors:
            raise ValueError("Config validation failed:\n- " + "\n- ".join(errors))
        return True


USERNAME_SELECTOR = '//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[1]/div/input'
PASSWORD_SELECTOR = '//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[2]/div/input'
REMEMBER_ME_SELECTOR = '//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[3]/div/input'
LOGIN_BUTTON_SELECTOR = '//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[4]/button'
OTP_INPUT_SELECTOR = '//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[1]/div/input'
OTP_CONFIRM_SELECTOR = '//*[@id="app"]/div/div[1]/div/div[3]/div[2]/div[2]/button[2]'

TOKEN_STORAGE_KEYS = (
    "token",
    "access_token",
    "accessToken",
    "authToken",
    "authorization",
    "Authorization",
)

CAPTURED_HEADER_KEYS = (
    "authorization",
    "token-id",
    "selectedmenuid",
    "selectedpath",
    "mac-address",
    "apikey",
)


def build_cookie_header(cookies):
    return "; ".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in cookies
        if cookie.get("name") and cookie.get("value")
    )


def read_otp_from_file(
    file_path="",
    max_age_seconds=None,
    max_retries=None,
    retry_interval_seconds=None,
    clear_after_read=True,
):
    otp_path = file_path or OneBSSSettings.OTP_FILE_PATH
    max_age = max_age_seconds or OneBSSSettings.OTP_MAX_AGE_SECONDS
    retries = max_retries or OneBSSSettings.OTP_MAX_RETRIES
    interval = (
        retry_interval_seconds
        if retry_interval_seconds is not None
        else OneBSSSettings.OTP_RETRY_INTERVAL_SECONDS
    )

    for retry_count in range(retries):
        if os.path.exists(otp_path):
            file_time = os.path.getmtime(otp_path)
            time_diff = time.time() - file_time
            if time_diff <= max_age:
                with open(otp_path, "r", encoding="utf-8") as handle:
                    content = handle.read()
                otp_match = re.search(r"\b\d{6}\b", content)
                if otp_match:
                    otp_code = otp_match.group(0)
                    if clear_after_read:
                        try:
                            with open(otp_path, "w", encoding="utf-8") as handle:
                                handle.write("")
                        except OSError:
                            pass
                    return otp_code
        if retry_count < retries - 1:
            time.sleep(interval)
    return None


def _try_fill_otp(page):
    otp_field = page.locator(OTP_INPUT_SELECTOR)
    try:
        otp_field.wait_for(state="visible", timeout=10000)
    except PlaywrightTimeoutError:
        return False

    otp_code = read_otp_from_file()
    if otp_code is None:
        print("Không đọc được OTP tự động cho OneBSS. Vui lòng nhập thủ công.")
        time.sleep(30)
        return True

    otp_field.fill(otp_code)
    confirm_button = page.locator(OTP_CONFIRM_SELECTOR)
    confirm_button.wait_for(state="visible", timeout=30000)
    confirm_button.click()
    time.sleep(5)
    return True


def _wait_for_networkidle_if_possible(page, timeout=None):
    effective_timeout = timeout or OneBSSSettings.PAGE_LOAD_TIMEOUT
    try:
        page.wait_for_load_state("networkidle", timeout=effective_timeout)
        return True
    except PlaywrightTimeoutError:
        return False


def login(headless=True):
    """Đăng nhập OneBSS bằng tài khoản dùng chung trong `.env`."""
    OneBSSSettings.validate()

    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context(accept_downloads=OneBSSSettings.ACCEPT_DOWNLOADS)
    page = context.new_page()

    print(f"=== Đăng nhập {OneBSSSettings.LOGIN_URL} ===")
    username_field = page.locator(USERNAME_SELECTOR)
    page.goto(
        OneBSSSettings.LOGIN_URL,
        wait_until="domcontentloaded",
        timeout=OneBSSSettings.PAGE_LOAD_TIMEOUT,
    )
    username_field.wait_for(state="visible", timeout=45000)
    username_field.fill(OneBSSSettings.USERNAME)

    password_field = page.locator(PASSWORD_SELECTOR)
    password_field.wait_for(state="visible", timeout=45000)
    password_field.fill(OneBSSSettings.PASSWORD)

    remember_checkbox = page.locator(REMEMBER_ME_SELECTOR)
    remember_checkbox.wait_for(state="visible", timeout=30000)
    remember_checkbox.check()

    login_button = page.locator(LOGIN_BUTTON_SELECTOR)
    login_button.wait_for(state="visible", timeout=30000)
    login_button.click()
    time.sleep(3)

    _wait_for_networkidle_if_possible(page, timeout=min(15000, OneBSSSettings.PAGE_LOAD_TIMEOUT))
    _try_fill_otp(page)
    _wait_for_networkidle_if_possible(page, timeout=min(15000, OneBSSSettings.PAGE_LOAD_TIMEOUT))

    print(f"✅ Đăng nhập OneBSS thành công, cookies={len(context.cookies())}")
    return playwright, browser, context, page


def _extract_token_from_text(raw_text):
    if not raw_text:
        return ""

    text = str(raw_text).strip()
    if re.match(r"^[Bb]earer\s+\S+$", text):
        return text.split(None, 1)[1].strip()
    if re.match(r"^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$", text):
        return text
    return ""


def _extract_token_from_json_like(raw_text):
    if not raw_text:
        return ""

    text = str(raw_text).strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        return ""

    candidates = [parsed]
    while candidates:
        current = candidates.pop(0)
        if isinstance(current, dict):
            for key, value in current.items():
                if key in TOKEN_STORAGE_KEYS:
                    token = _extract_token_from_text(value)
                    if token:
                        return token
                candidates.append(value)
        elif isinstance(current, list):
            candidates.extend(current)
    return ""


def extract_token_from_storage(page):
    """Quét local/session storage để lấy token nếu request chưa lộ Authorization."""
    storage_values = page.evaluate(
        """() => {
            const result = [];
            for (const storage of [window.localStorage, window.sessionStorage]) {
                for (let i = 0; i < storage.length; i += 1) {
                    const key = storage.key(i);
                    result.push({ key, value: storage.getItem(key) });
                }
            }
            return result;
        }"""
    )

    for item in storage_values:
        token = _extract_token_from_text(item.get("value"))
        if token:
            return {
                "token": token,
                "source": f"storage:{item.get('key')}",
            }

        token = _extract_token_from_json_like(item.get("value"))
        if token:
            return {
                "token": token,
                "source": f"storage-json:{item.get('key')}",
            }
    return {}


def capture_authorization(page, target_url="", timeout_seconds=30, host_filter=""):
    """Bắt Authorization header/token từ request XHR/fetch của OneBSS."""
    state = {}
    request_host_filter = host_filter or OneBSSSettings.REQUEST_HOST_FILTER

    def on_request(request):
        if request_host_filter and request_host_filter not in request.url:
            return

        headers = request.headers
        authorization = headers.get("authorization") or headers.get("Authorization")
        if authorization and "authorization" not in state:
            state["authorization"] = authorization
            state["token"] = authorization.split(None, 1)[1] if " " in authorization else authorization
            state["user_agent"] = headers.get("user-agent", "")
            state["accept"] = headers.get("accept", "application/json, text/plain, */*")
            state["referer"] = headers.get("referer", page.url or OneBSSSettings.BASE_URL + "/")
            state["source"] = "request-header"

        for key in CAPTURED_HEADER_KEYS:
            header_value = headers.get(key) or headers.get(key.title())
            if header_value and key not in state:
                state[key] = header_value

    page.context.on("request", on_request)

    if target_url:
        print(f"Đang mở trang OneBSS để bắt token: {target_url}")
        page.goto(
            target_url,
            wait_until="domcontentloaded",
            timeout=OneBSSSettings.PAGE_LOAD_TIMEOUT,
        )
        _wait_for_networkidle_if_possible(
            page,
            timeout=min(15000, OneBSSSettings.PAGE_LOAD_TIMEOUT),
        )

    started = time.time()
    while time.time() - started < timeout_seconds:
        if state.get("authorization"):
            print("✅ Đã bắt được Authorization header của OneBSS")
            return state

        storage_state = extract_token_from_storage(page)
        if storage_state.get("token"):
            state["token"] = storage_state["token"]
            state["authorization"] = f"Bearer {storage_state['token']}"
            state["user_agent"] = page.evaluate("() => navigator.userAgent")
            state["accept"] = "application/json, text/plain, */*"
            state["referer"] = page.url or OneBSSSettings.BASE_URL + "/"
            state["source"] = storage_state["source"]
            state.setdefault("token-id", OneBSSSettings.DEFAULT_TOKEN_ID)
            state.setdefault("selectedmenuid", OneBSSSettings.DEFAULT_SELECTED_MENU_ID)
            state.setdefault("selectedpath", OneBSSSettings.DEFAULT_SELECTED_PATH)
            state.setdefault("mac-address", OneBSSSettings.DEFAULT_MAC_ADDRESS)
            state.setdefault("apikey", OneBSSSettings.DEFAULT_API_KEY)
            print(f"✅ Đã lấy token OneBSS từ {state['source']}")
            return state

        page.wait_for_timeout(500)

    raise RuntimeError("Không bắt được Authorization/token từ OneBSS.")


def capture_token(page, target_url="", timeout_seconds=30, host_filter=""):
    """Trả về token thô để tái sử dụng ở các API downloader OneBSS."""
    auth_state = capture_authorization(
        page,
        target_url=target_url,
        timeout_seconds=timeout_seconds,
        host_filter=host_filter,
    )
    return auth_state["token"]


def make_common_headers(auth_state, cookies, extra_headers=None):
    headers = {
        "Authorization": auth_state["authorization"],
        "Accept": auth_state.get("accept", "application/json, text/plain, */*"),
        "Referer": auth_state.get("referer", OneBSSSettings.BASE_URL + "/"),
        "User-Agent": auth_state.get("user_agent", ""),
        "Cookie": build_cookie_header(cookies),
        "token-id": auth_state.get("token-id", OneBSSSettings.DEFAULT_TOKEN_ID),
        "selectedmenuid": auth_state.get("selectedmenuid", OneBSSSettings.DEFAULT_SELECTED_MENU_ID),
        "selectedpath": auth_state.get("selectedpath", OneBSSSettings.DEFAULT_SELECTED_PATH),
        "mac-address": auth_state.get("mac-address", OneBSSSettings.DEFAULT_MAC_ADDRESS),
    }
    headers.update(extra_headers or {})
    return headers


def create_session(headed=False, token_page_url="", extra_headers=None):
    """Login và bắt token, trả về session dict cho downloader OneBSS."""
    token_url = token_page_url or OneBSSSettings.DEFAULT_TOKEN_PAGE_URL
    playwright, browser, context, page = login(headless=not headed)
    auth_state = capture_authorization(page, target_url=token_url)
    headers = make_common_headers(auth_state, context.cookies(), extra_headers=extra_headers)

    return {
        "token": auth_state["token"],
        "api_base_url": OneBSSSettings.API_BASE_URL,
        "headers": headers,
        "auth_state": auth_state,
        "playwright": playwright,
        "browser": browser,
        "context": context,
        "page": page,
    }


def close_session(session):
    """Đóng browser và playwright của session OneBSS."""
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
