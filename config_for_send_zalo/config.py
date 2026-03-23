#!/usr/bin/env python3
"""
Configuration Module
Centralized configuration loading from .env file using python-dotenv
"""

import os
from pathlib import Path
from dotenv import load_dotenv
from team_config import get_location_thread_mapping, get_location_chat_mapping

# Load .env file from the project root
env_path = Path(__file__).parent / '.env'
load_dotenv(dotenv_path=env_path)

# ================================
# CLOUDINARY CONFIGURATION
# ================================
CLOUDINARY_CLOUD_NAME = os.getenv('CLOUDINARY_CLOUD_NAME', '')
CLOUDINARY_API_KEY = os.getenv('CLOUDINARY_API_KEY', '')
CLOUDINARY_API_SECRET = os.getenv('CLOUDINARY_API_SECRET', '')

# ================================
# ZALO THREAD IDS
# ================================
ZALO_THREAD_QUANGOAI = os.getenv('ZALO_THREAD_QUANGOAI', '')
ZALO_THREAD_SUOIHAI = os.getenv('ZALO_THREAD_SUOIHAI', '')
ZALO_THREAD_SONTAY = os.getenv('ZALO_THREAD_SONTAY', '')
ZALO_THREAD_PHUCTHO = os.getenv('ZALO_THREAD_PHUCTHO', '')
ZALO_THREAD_DEFAULT = os.getenv('ZALO_THREAD_DEFAULT', '')

# Location to ThreadID mapping (auto-generated from team_config)
LOCATION_THREAD_MAPPING = get_location_thread_mapping()
LOCATION_THREAD_MAPPING['default'] = ZALO_THREAD_DEFAULT  # Add default fallback

# ================================
# WEBHOOK URLS
# ================================
WEBHOOK_IMAGE_URL = os.getenv('WEBHOOK_IMAGE_URL', '')
WEBHOOK_TEXT_URL = os.getenv('WEBHOOK_TEXT_URL', '')

# ================================
# ONEBSS LOGIN CREDENTIALS
# ================================
ONEBSS_USERNAME = os.getenv('ONEBSS_USERNAME', '')
ONEBSS_PASSWORD = os.getenv('ONEBSS_PASSWORD', '')
ONEBSS_URL = os.getenv('ONEBSS_URL', 'https://onebss.vnpt.vn')
OTP_FILE_PATH = os.getenv('OTP_FILE_PATH', '/home/vtst/otp/otp_logs.txt')

# ================================
# TELEGRAM CREDENTIALS
# ================================
TELEGRAM_TOKEN = os.getenv('TELEGRAM_TOKEN', '')
TELEGRAM_CHAT_ID_MAIN = os.getenv('TELEGRAM_CHAT_ID_MAIN', '')
TELEGRAM_CHAT_ID_QUANGOAI = os.getenv('TELEGRAM_CHAT_ID_QUANGOAI', '')
TELEGRAM_CHAT_ID_SUOIHAI = os.getenv('TELEGRAM_CHAT_ID_SUOIHAI', '')
TELEGRAM_CHAT_ID_SONTAY = os.getenv('TELEGRAM_CHAT_ID_SONTAY', '')
TELEGRAM_CHAT_ID_PHUCTHO = os.getenv('TELEGRAM_CHAT_ID_PHUCTHO', '')

# Location to Chat ID mapping for Telegram (auto-generated from team_config)
LOCATION_CHAT_MAPPING = get_location_chat_mapping()
LOCATION_CHAT_MAPPING['default'] = TELEGRAM_CHAT_ID_MAIN  # Add default fallback

# ================================
# TIME RESTRICTION SETTINGS
# ================================
SEND_START_HOUR = int(os.getenv('SEND_START_HOUR', '6'))
SEND_START_MINUTE = int(os.getenv('SEND_START_MINUTE', '30'))
SEND_END_HOUR = int(os.getenv('SEND_END_HOUR', '21'))
SEND_END_MINUTE = int(os.getenv('SEND_END_MINUTE', '0'))

# ================================
# FILE PATHS
# ================================
WARNING_DB = os.getenv('WARNING_DB', 'warning_tracking.db')
EXCEL_FILE_BRCD = os.getenv('EXCEL_FILE_BRCD', 'chiaTheoDoi/chiTietBrcd5Doi.xlsx')
LOG_FILE_WARNING = os.getenv('LOG_FILE_WARNING', 'warning_sender.log')
LOG_FILE_TELEGRAM = os.getenv('LOG_FILE_TELEGRAM', 'telegram_sender.log')

# ================================
# FOLDER PATHS
# ================================
IMAGE_FOLDER = os.getenv('IMAGE_FOLDER', 'image')
CHART_FOLDER = os.getenv('CHART_FOLDER', 'chart')

# ================================
# BLACKLISTED LOCATIONS
# ================================
# Danh sách địa bàn cũ không gửi nữa
# NOTE: Blacklist cleared - now using team_config for team management
# All teams are managed through team_config.py with active/inactive flags
LOCATION_BLACKLIST = []


def get_cloudinary_config():
    """
    Returns Cloudinary configuration as a dict for direct usage.

    Returns:
        dict: Configuration dict with cloud_name, api_key, api_secret, secure
    """
    return {
        'cloud_name': CLOUDINARY_CLOUD_NAME,
        'api_key': CLOUDINARY_API_KEY,
        'api_secret': CLOUDINARY_API_SECRET,
        'secure': True
    }


def validate_config():
    """
    Validate that all required configuration values are loaded.
    Prints warnings for missing values.

    Returns:
        bool: True if all required configs are present, False otherwise
    """
    required_configs = {
        'CLOUDINARY_CLOUD_NAME': CLOUDINARY_CLOUD_NAME,
        'CLOUDINARY_API_KEY': CLOUDINARY_API_KEY,
        'CLOUDINARY_API_SECRET': CLOUDINARY_API_SECRET,
        'ZALO_THREAD_QUANGOAI': ZALO_THREAD_QUANGOAI,
        'ZALO_THREAD_SUOIHAI': ZALO_THREAD_SUOIHAI,
        'ZALO_THREAD_SONTAY': ZALO_THREAD_SONTAY,
        'WEBHOOK_IMAGE_URL': WEBHOOK_IMAGE_URL,
        'WEBHOOK_TEXT_URL': WEBHOOK_TEXT_URL,
        'ONEBSS_USERNAME': ONEBSS_USERNAME,
        'ONEBSS_PASSWORD': ONEBSS_PASSWORD,
        'TELEGRAM_TOKEN': TELEGRAM_TOKEN,
    }

    all_valid = True
    for key, value in required_configs.items():
        if not value:
            print(f"⚠️  WARNING: {key} is not set in .env file!")
            all_valid = False

    if all_valid:
        print("✅ All required configuration values are loaded successfully")

    return all_valid


if __name__ == "__main__":
    print("=" * 60)
    print("🔧 Configuration Module Test")
    print("=" * 60)

    # Validate configuration
    validate_config()

    print("\n📋 Configuration Summary:")
    print(f"  Cloudinary Cloud: {CLOUDINARY_CLOUD_NAME}")
    print(f"  Zalo Threads: quangoai={ZALO_THREAD_QUANGOAI}, suoihai={ZALO_THREAD_SUOIHAI}, sontay={ZALO_THREAD_SONTAY}")
    print(f"  Webhook Image: {WEBHOOK_IMAGE_URL}")
    print(f"  Webhook Text: {WEBHOOK_TEXT_URL}")
    print(f"  OneBSS User: {ONEBSS_USERNAME}")
    print(f"  Telegram Token: {TELEGRAM_TOKEN[:20]}...")
    print(f"  Time Window: {SEND_START_HOUR:02d}:{SEND_START_MINUTE:02d} - {SEND_END_HOUR:02d}:{SEND_END_MINUTE:02d}")
    print("=" * 60)
