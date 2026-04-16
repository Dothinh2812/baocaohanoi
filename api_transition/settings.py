# -*- coding: utf-8 -*-
"""Cấu hình riêng cho thư mục api_transition."""

import os
from pathlib import Path

from dotenv import load_dotenv


ROOT_DIR = Path(__file__).resolve().parent.parent
load_dotenv(ROOT_DIR / ".env")


class Settings:
    BAOCAO_USERNAME = os.getenv("BAOCAO_USERNAME", "")
    BAOCAO_PASSWORD = os.getenv("BAOCAO_PASSWORD", "")
    BAOCAO_BASE_URL = os.getenv("BAOCAO_BASE_URL", "https://baocao.hanoi.vnpt.vn")
    BAOCAO_URL = BAOCAO_BASE_URL + "/"

    OTP_FILE_PATH = os.getenv("OTP_FILE_PATH", "")
    OTP_MAX_AGE_SECONDS = int(os.getenv("OTP_MAX_AGE_SECONDS", "120"))

    PAGE_LOAD_TIMEOUT = int(os.getenv("PAGE_LOAD_TIMEOUT", "60000"))
    NETWORK_IDLE_TIMEOUT = int(os.getenv("NETWORK_IDLE_TIMEOUT", "500000"))
    DOWNLOAD_TIMEOUT = int(os.getenv("DOWNLOAD_TIMEOUT", "120000"))

    ACCEPT_DOWNLOADS = os.getenv("ACCEPT_DOWNLOADS", "True").lower() == "true"

    API_BASE_URL = "https://baocaobe.myhanoi.vn/report-api"
    DEFAULT_REFERER = BAOCAO_BASE_URL + "/"

    @classmethod
    def validate(cls):
        errors = []
        if not cls.BAOCAO_USERNAME:
            errors.append("BAOCAO_USERNAME không được để trống")
        if not cls.BAOCAO_PASSWORD:
            errors.append("BAOCAO_PASSWORD không được để trống")
        if not cls.OTP_FILE_PATH:
            errors.append("OTP_FILE_PATH không được để trống")
        if errors:
            raise ValueError("Config validation failed:\n- " + "\n- ".join(errors))
        return True
