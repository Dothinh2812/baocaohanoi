"""
Module xử lý Telegram bot webhook
Nhận ảnh từ group Telegram, upload lên GCS, sau đó xử lý trực tiếp với Vision API và OpenAI
"""
import os
import json
from datetime import datetime
from typing import Optional, Dict, Any

from gcs_storage import upload_image_from_telegram
# Import các function xử lý từ gsmnv.py để gọi trực tiếp (không qua HTTP)
from gsmnv import (
    process_image_with_vision,
    analyze_text_with_openai,
    save_to_google_sheets
)

# Cấu hình
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHANNEL_ID = os.environ.get("TELEGRAM_CHANNEL_ID", "")

# GCS bucket name
GCS_BUCKET_NAME = os.environ.get("GCS_BUCKET_NAME", "bts-telegram-images")


def extract_telegram_metadata(update: dict) -> Optional[Dict[str, Any]]:
    """
    Trích xuất metadata từ message object của Telegram

    Args:
        update: Telegram update object từ webhook

    Returns:
        Dictionary chứa metadata hoặc None nếu không có ảnh
    """
    try:
        message = update.get("message")
        if not message:
            print("⚠️ Không có message trong update")
            return None

        # Kiểm tra có ảnh không
        if not message.get("photo"):
            print("⚠️ Message không chứa ảnh")
            return None

        # Lấy ảnh có kích thước lớn nhất
        photos = message["photo"]
        largest_photo = photos[-1]  # Ảnh cuối cùng là kích thước lớn nhất
        file_id = largest_photo["file_id"]

        # Trích xuất metadata
        sender = message.get("from", {})
        sender_name = sender.get("first_name", "")
        if sender.get("last_name"):
            sender_name += f" {sender['last_name']}"
        sender_username = sender.get("username", "")

        caption = message.get("caption", "") or message.get("text", "")
        message_id = message.get("message_id", "")
        message_date = message.get("date", 0)
        chat_id = message.get("chat", {}).get("id", "")
        chat_title = message.get("chat", {}).get("title", "Telegram Group")

        # Chuyển timestamp thành datetime
        from datetime import datetime as dt
        message_datetime = dt.fromtimestamp(message_date)

        metadata = {
            "file_id": file_id,
            "sender_name": sender_name,
            "sender_username": sender_username,
            "caption": caption,
            "message_id": str(message_id),
            "message_date": message_date,
            "message_datetime": message_datetime.isoformat(),
            "chat_id": str(chat_id),
            "chat_title": chat_title
        }

        print(f"✅ Trích xuất metadata từ Telegram:")
        print(f"   - Người gửi: {sender_name} ({sender_username})")
        print(f"   - Caption: {caption}")
        print(f"   - Message ID: {message_id}")

        return metadata

    except Exception as e:
        print(f"❌ Lỗi trích xuất metadata: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return None


def process_telegram_image(update: dict) -> Optional[Dict[str, Any]]:
    """
    Xử lý ảnh từ Telegram webhook - DIRECT PROCESSING (không qua HTTP)

    1. Trích xuất metadata từ Telegram
    2. Upload ảnh lên Google Cloud Storage (GCS)
    3. Gọi trực tiếp Vision API để OCR text từ ảnh
    4. Gọi trực tiếp OpenAI để phân tích text
    5. Lưu vào Google Sheets và gửi Telegram response

    Args:
        update: Telegram update object từ webhook

    Returns:
        Kết quả xử lý hoặc None nếu có lỗi
    """
    try:
        print("=" * 60)
        print("Xử lý ảnh từ Telegram")
        print("=" * 60)

        # Bước 1: Trích xuất metadata
        print("\nBước 1: Trích xuất metadata từ Telegram")
        metadata = extract_telegram_metadata(update)
        if not metadata:
            return {"error": "Không thể trích xuất metadata từ message"}

        # Bước 2: Upload ảnh lên GCS
        print("\nBước 2: Upload ảnh lên Google Cloud Storage")
        file_id = metadata["file_id"]
        filename = f"telegram_{metadata['message_id']}_{metadata['sender_username']}.jpg"

        gcs_result = upload_image_from_telegram(
            file_id=file_id,
            bot_token=TELEGRAM_BOT_TOKEN,
            filename=filename
        )

        if "error" in gcs_result:
            print(f"❌ Lỗi upload ảnh: {gcs_result['error']}")
            return gcs_result

        image_url = gcs_result["public_url"]
        print(f"✅ Ảnh đã upload: {image_url}")

        # Chuẩn bị metadata
        threadId = metadata["message_id"]
        name = metadata["sender_name"] or metadata["sender_username"] or "Unknown"
        title = metadata["caption"] or f"Telegram Group: {metadata['chat_title']}"

        # Bước 3: Xử lý ảnh với Google Cloud Vision API
        print("\nBước 3: Xử lý ảnh với Google Cloud Vision API")
        vision_text = process_image_with_vision(image_url)
        print(f"Vision API - Text detected: {vision_text[:200] if vision_text else 'No text'}...")

        # Bước 4: Phân tích text với OpenAI
        print("\nBước 4: Phân tích text với OpenAI")
        openai_analysis = analyze_text_with_openai(vision_text, name)
        print(f"OpenAI - Analysis complete: {openai_analysis[:200] if openai_analysis else 'No analysis'}...")

        # Bước 5: Lưu vào Google Sheets (sẽ tự động gửi Telegram response)
        print("\nBước 5: Lưu vào Google Sheets và gửi Telegram response")
        sheets_success = save_to_google_sheets(
            threadId,
            name,
            title,
            image_url,
            vision_text,
            openai_analysis
        )

        if sheets_success:
            print("✅ Xử lý hoàn tất: GCS → Vision API → OpenAI → Sheets → Telegram")
            return {
                "success": True,
                "message": "Ảnh đã được xử lý thành công",
                "image_url": image_url,
                "thread_id": threadId,
                "vision_text": vision_text[:200] + "..." if len(vision_text) > 200 else vision_text,
                "openai_analysis": openai_analysis[:200] + "..." if len(openai_analysis) > 200 else openai_analysis
            }
        else:
            error_msg = "Lỗi khi lưu vào Google Sheets"
            print(f"❌ {error_msg}")
            return {"error": error_msg}
    except Exception as e:
        error_msg = f"Lỗi xử lý ảnh: {str(e)}"
        print(f"❌ {error_msg}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return {"error": error_msg}


def validate_telegram_webhook(body: dict, bot_token: str) -> bool:
    """
    Xác thực webhook đến từ Telegram (optional)
    Hiện tại chỉ kiểm tra cơ bản

    Args:
        body: Request body từ webhook
        bot_token: Telegram bot token

    Returns:
        True nếu hợp lệ
    """
    # Telegram không gửi signature, chỉ kiểm tra cơ bản
    if "message" in body or "edited_message" in body:
        return True
    return False


# Test function
if __name__ == "__main__":
    print("=" * 60)
    print("TEST TELEGRAM BOT MODULE")
    print("=" * 60)

    # Test data - giả lập Telegram webhook
    test_update = {
        "update_id": 123456789,
        "message": {
            "message_id": 1001,
            "date": int(datetime.now().timestamp()),
            "chat": {
                "id": -4863386433,
                "title": "Test Group",
                "type": "supergroup"
            },
            "from": {
                "id": 12345,
                "is_bot": False,
                "first_name": "Test",
                "last_name": "User",
                "username": "testuser"
            },
            "photo": [
                {
                    "file_id": "AgACAgIAAxkBAAIF...",  # File ID mẫu
                    "file_unique_id": "AQADUoD...",
                    "width": 320,
                    "height": 240,
                    "file_size": 8000
                }
            ],
            "caption": "Test ảnh từ Telegram"
        }
    }

    print("\nTest: Trích xuất metadata")
    metadata = extract_telegram_metadata(test_update)
    if metadata:
        print("✅ Metadata:")
        print(json.dumps(metadata, indent=2, ensure_ascii=False))
    else:
        print("❌ Lỗi trích xuất metadata")

    print("\n" + "=" * 60)
