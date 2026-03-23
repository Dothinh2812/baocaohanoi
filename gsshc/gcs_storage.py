"""
Module quản lý Google Cloud Storage (GCS) cho việc upload ảnh từ Telegram
"""
import os
import json
from google.cloud import storage
from google.oauth2 import service_account
from datetime import datetime
import requests
from io import BytesIO

# Cấu hình GCS
GCS_BUCKET_NAME = os.environ.get("GCS_BUCKET_NAME", "bts-telegram-images")  # Sẽ thay bằng tên bucket thực tế
GCS_PROJECT_ID = os.environ.get("GCS_PROJECT_ID", "your-project-id")  # Sẽ thay bằng project ID thực tế


def get_gcs_client():
    """
    Khởi tạo Google Cloud Storage client
    Sử dụng credentials từ file vision-key.json
    """
    try:
        credentials = service_account.Credentials.from_service_account_file(
            "vision-key.json"
        )
        client = storage.Client(credentials=credentials, project=GCS_PROJECT_ID)
        return client
    except Exception as e:
        print(f"Lỗi khởi tạo GCS client: {str(e)}")
        return None


def upload_image_from_telegram(file_id: str, bot_token: str, filename: str = None) -> dict:
    """
    Tải ảnh từ Telegram và upload lên Google Cloud Storage

    Args:
        file_id: ID file của ảnh từ Telegram API
        bot_token: Token của Telegram bot
        filename: Tên file để lưu (nếu None sẽ dùng file_id)

    Returns:
        Dictionary chứa URL public và thông tin upload, hoặc error message
    """
    try:
        # Bước 1: Lấy đường dẫn file từ Telegram API
        print(f"Step 1: Lấy file path từ Telegram (file_id: {file_id})")
        get_file_url = f"https://api.telegram.org/bot{bot_token}/getFile?file_id={file_id}"

        file_info_response = requests.get(get_file_url)
        if file_info_response.status_code != 200:
            error_msg = f"Lỗi lấy file info từ Telegram: {file_info_response.text}"
            print(f"❌ {error_msg}")
            return {"error": error_msg}

        file_path = file_info_response.json()["result"]["file_path"]
        file_url = f"https://api.telegram.org/file/bot{bot_token}/{file_path}"
        print(f"✅ File path: {file_path}")

        # Bước 2: Tải ảnh từ URL
        print("Step 2: Tải ảnh từ Telegram server")
        image_response = requests.get(file_url)
        if image_response.status_code != 200:
            error_msg = f"Lỗi tải ảnh từ Telegram: Status {image_response.status_code}"
            print(f"❌ {error_msg}")
            return {"error": error_msg}

        image_data = image_response.content
        print(f"✅ Đã tải ảnh: {len(image_data)} bytes")

        # Bước 3: Upload lên GCS
        print("Step 3: Upload ảnh lên Google Cloud Storage")
        client = get_gcs_client()
        if not client:
            return {"error": "Không thể khởi tạo GCS client"}

        bucket = client.bucket(GCS_BUCKET_NAME)

        # Tạo tên file nếu không có
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"telegram_{timestamp}_{file_id}.jpg"

        # Thêm ngày vào path để organize
        today = datetime.now().strftime("%Y/%m/%d")
        blob_path = f"telegram-images/{today}/{filename}"

        blob = bucket.blob(blob_path)
        blob.upload_from_string(image_data, content_type="image/jpeg")

        print(f"✅ Đã upload lên GCS: {blob_path}")

        # Bước 4: Tạo public URL
        # Lưu ý: Bucket phải có public read permission
        public_url = f"https://storage.googleapis.com/{GCS_BUCKET_NAME}/{blob_path}"

        print(f"✅ Public URL: {public_url}")

        return {
            "success": True,
            "gcs_path": blob_path,
            "public_url": public_url,
            "filename": filename,
            "file_size": len(image_data),
            "uploaded_at": datetime.now().isoformat()
        }

    except Exception as e:
        error_msg = f"Lỗi upload ảnh: {str(e)}"
        print(f"❌ {error_msg}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return {"error": error_msg}


def make_bucket_public(bucket_name: str = None) -> dict:
    """
    Cấu hình bucket để cho phép public read
    Chỉ chạy một lần khi setup

    Args:
        bucket_name: Tên bucket (nếu None dùng GCS_BUCKET_NAME)

    Returns:
        Status của việc cấu hình
    """
    try:
        if not bucket_name:
            bucket_name = GCS_BUCKET_NAME

        client = get_gcs_client()
        if not client:
            return {"error": "Không thể khởi tạo GCS client"}

        bucket = client.bucket(bucket_name)

        # Cấp quyền AllUsers để có thể read các object
        bucket.make_public()
        print(f"✅ Bucket {bucket_name} đã được set thành public")

        return {"success": True, "message": f"Bucket {bucket_name} configured for public read"}

    except Exception as e:
        error_msg = f"Lỗi cấu hình bucket: {str(e)}"
        print(f"❌ {error_msg}")
        return {"error": error_msg}


def create_bucket_if_not_exists(bucket_name: str, location: str = "us-central1") -> dict:
    """
    Tạo bucket GCS nếu chưa tồn tại

    Args:
        bucket_name: Tên bucket
        location: Vị trí bucket (ví dụ: us-central1, asia-southeast1, ...)

    Returns:
        Status của việc tạo bucket
    """
    try:
        client = get_gcs_client()
        if not client:
            return {"error": "Không thể khởi tạo GCS client"}

        bucket = client.bucket(bucket_name)

        # Kiểm tra bucket đã tồn tại
        if bucket.exists():
            print(f"✅ Bucket {bucket_name} đã tồn tại")
            return {"success": True, "message": f"Bucket {bucket_name} already exists"}

        # Tạo bucket mới
        bucket = client.create_bucket(bucket_name, location=location)
        print(f"✅ Bucket {bucket_name} đã được tạo tại {location}")

        return {"success": True, "message": f"Bucket {bucket_name} created"}

    except Exception as e:
        error_msg = f"Lỗi tạo bucket: {str(e)}"
        print(f"❌ {error_msg}")
        return {"error": error_msg}


# Test function
if __name__ == "__main__":
    print("=" * 60)
    print("TEST GCS STORAGE MODULE")
    print("=" * 60)

    # Test 1: Kiểm tra GCS client
    print("\nTest 1: Khởi tạo GCS client")
    client = get_gcs_client()
    if client:
        print("✅ GCS client khởi tạo thành công")
    else:
        print("❌ Lỗi khởi tạo GCS client")

    print("\n" + "=" * 60)
