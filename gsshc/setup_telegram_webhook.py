"""
Script để setup webhook với Telegram
Chạy script này một lần để đăng ký webhook URL với Telegram
"""
import os
import requests
import json
import sys

# Cấu hình
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHANNEL_ID = os.environ.get("TELEGRAM_CHANNEL_ID", "")

# Domain webhook
WEBHOOK_DOMAIN = os.environ.get("WEBHOOK_DOMAIN", "n8n2.ttvt8.online")
WEBHOOK_PORT = os.environ.get("WEBHOOK_PORT", 443)  # HTTPS mặc định là 443
WEBHOOK_PATH = "/telegram-webhook"

# Xây dựng webhook URL
if WEBHOOK_PORT == 443 or WEBHOOK_PORT == "443":
    WEBHOOK_URL = f"https://{WEBHOOK_DOMAIN}{WEBHOOK_PATH}"
else:
    WEBHOOK_URL = f"https://{WEBHOOK_DOMAIN}:{WEBHOOK_PORT}{WEBHOOK_PATH}"


def set_telegram_webhook() -> bool:
    """
    Đăng ký webhook URL với Telegram

    Returns:
        True nếu thành công
    """
    try:
        print("=" * 60)
        print("SETUP TELEGRAM WEBHOOK")
        print("=" * 60)

        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/setWebhook"

        data = {
            "url": WEBHOOK_URL,
            "allowed_updates": ["message"],  # Chỉ lắng nghe message updates
            "drop_pending_updates": True  # Bỏ các message cũ pending
        }

        print(f"\nBot Token: {TELEGRAM_BOT_TOKEN[:20]}...")
        print(f"Webhook URL: {WEBHOOK_URL}")
        print(f"Data: {json.dumps(data, indent=2)}")

        print("\nGửi yêu cầu setWebhook tới Telegram...")
        response = requests.post(url, json=data, timeout=10)

        if response.status_code != 200:
            print(f"❌ Lỗi từ Telegram API: Status {response.status_code}")
            print(f"Response: {response.text}")
            return False

        result = response.json()
        if result.get("ok"):
            print(f"✅ Webhook đã được setup thành công!")
            print(f"Response: {json.dumps(result, indent=2, ensure_ascii=False)}")
            return True
        else:
            print(f"❌ Telegram trả về lỗi:")
            print(f"Response: {json.dumps(result, indent=2, ensure_ascii=False)}")
            return False

    except requests.exceptions.Timeout:
        print("❌ Timeout khi gửi yêu cầu")
        return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Lỗi request: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ Lỗi: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return False


def get_webhook_info() -> bool:
    """
    Lấy thông tin webhook hiện tại

    Returns:
        True nếu thành công
    """
    try:
        print("\n" + "=" * 60)
        print("GET WEBHOOK INFO")
        print("=" * 60)

        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getWebhookInfo"

        print(f"\nGửi yêu cầu getWebhookInfo tới Telegram...")
        response = requests.get(url, timeout=10)

        if response.status_code != 200:
            print(f"❌ Lỗi từ Telegram API: Status {response.status_code}")
            print(f"Response: {response.text}")
            return False

        result = response.json()
        if result.get("ok"):
            webhook_info = result.get("result", {})
            print(f"✅ Thông tin webhook hiện tại:")
            print(json.dumps(webhook_info, indent=2, ensure_ascii=False))
            return True
        else:
            print(f"❌ Telegram trả về lỗi:")
            print(json.dumps(result, indent=2, ensure_ascii=False))
            return False

    except Exception as e:
        print(f"❌ Lỗi: {str(e)}")
        return False


def delete_webhook() -> bool:
    """
    Xóa webhook (nếu cần)

    Returns:
        True nếu thành công
    """
    try:
        print("\n" + "=" * 60)
        print("DELETE WEBHOOK")
        print("=" * 60)

        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/deleteWebhook"
        data = {"drop_pending_updates": True}

        print(f"\nGửi yêu cầu deleteWebhook tới Telegram...")
        response = requests.post(url, json=data, timeout=10)

        if response.status_code != 200:
            print(f"❌ Lỗi từ Telegram API: Status {response.status_code}")
            print(f"Response: {response.text}")
            return False

        result = response.json()
        if result.get("ok"):
            print(f"✅ Webhook đã được xóa!")
            return True
        else:
            print(f"❌ Telegram trả về lỗi:")
            print(json.dumps(result, indent=2, ensure_ascii=False))
            return False

    except Exception as e:
        print(f"❌ Lỗi: {str(e)}")
        return False


def setup_gcs_bucket() -> bool:
    """
    Setup Google Cloud Storage bucket (nếu có)

    Returns:
        True nếu thành công
    """
    try:
        print("\n" + "=" * 60)
        print("SETUP GOOGLE CLOUD STORAGE BUCKET")
        print("=" * 60)

        try:
            from gcs_storage import create_bucket_if_not_exists, make_bucket_public
        except ImportError:
            print("⚠️ gcs_storage module không tìm thấy, bỏ qua setup GCS")
            return False

        bucket_name = os.environ.get("GCS_BUCKET_NAME", "bts-telegram-images")

        print(f"\nBucket name: {bucket_name}")

        # Tạo bucket nếu chưa tồn tại
        print("\n1. Tạo bucket nếu chưa tồn tại...")
        create_result = create_bucket_if_not_exists(bucket_name, location="us-central1")
        print(json.dumps(create_result, indent=2, ensure_ascii=False))

        if "error" in create_result:
            print(f"⚠️ Lỗi tạo bucket (có thể đã tồn tại)")

        # Cấu hình bucket để public read
        print("\n2. Cấu hình bucket để public read...")
        public_result = make_bucket_public(bucket_name)
        print(json.dumps(public_result, indent=2, ensure_ascii=False))

        if "error" in public_result:
            print(f"⚠️ Lỗi cấu hình bucket")
            return False

        print("\n✅ GCS bucket đã được setup!")
        return True

    except Exception as e:
        print(f"❌ Lỗi setup GCS: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return False


def main():
    """
    Main function - chạy setup workflow
    """
    print("""
╔════════════════════════════════════════════════════════╗
║     TELEGRAM BOT WEBHOOK SETUP SCRIPT                 ║
╚════════════════════════════════════════════════════════╝

Các bước setup:
1. Setup GCS bucket (nếu chưa có)
2. Đăng ký webhook với Telegram
3. Kiểm tra webhook info

Lưu ý:
- Bot token: {0}
- Channel ID: {1}
- Webhook URL: {2}
- Cần HTTPS domain (hiện có: {3})

""".format(
        TELEGRAM_BOT_TOKEN[:20] + "...",
        TELEGRAM_CHANNEL_ID,
        WEBHOOK_URL,
        WEBHOOK_DOMAIN
    ))

    # Kiểm tra environment variables
    if not TELEGRAM_BOT_TOKEN:
        print("⚠️ Cảnh báo: TELEGRAM_BOT_TOKEN chưa được cấu hình")
        print("   Hãy set TELEGRAM_BOT_TOKEN environment variable")
        response = input("Tiếp tục? (y/n): ")
        if response.lower() != 'y':
            print("Hủy bỏ.")
            return False

    # Step 1: Setup GCS
    print("\nBước 1: Setup GCS bucket (nếu cần)")
    setup_gcs = input("Bạn muốn setup GCS bucket? (y/n): ")
    if setup_gcs.lower() == 'y':
        setup_gcs_bucket()

    # Step 2: Setup webhook
    print("\nBước 2: Đăng ký webhook với Telegram")
    if not set_telegram_webhook():
        print("\n❌ Setup webhook thất bại!")
        return False

    # Step 3: Get webhook info
    print("\nBước 3: Kiểm tra webhook info")
    get_webhook_info()

    print("\n" + "=" * 60)
    print("✅ SETUP HOÀN THÀNH!")
    print("=" * 60)
    print("""
Bước tiếp theo:
1. Chắc chắn server của bạn đang chạy trên port 8007
2. HTTPS domain n8n2.ttvt8.online phải forward tới server của bạn
3. Gửi ảnh vào group Telegram để test

Test endpoint:
curl -X POST https://n8n2.ttvt8.online/telegram-webhook \\
  -H "Content-Type: application/json" \\
  -d '{"message": {"photo": [{"file_id": "test"}]}}'
""")
    return True


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        command = sys.argv[1]
        if command == "delete":
            delete_webhook()
        elif command == "info":
            get_webhook_info()
        elif command == "gcs":
            setup_gcs_bucket()
        else:
            print(f"Lệnh không nhận ra: {command}")
            print("Lệnh có sẵn: delete, info, gcs")
    else:
        main()
