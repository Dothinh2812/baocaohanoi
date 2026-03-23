# CLAUDE.md

Tệp này cung cấp hướng dẫn cho Claude Code (claude.ai/code) khi làm việc với code trong repository này.

## Tổng Quan Dự Án

Đây là một hệ thống webhook API dựa trên Python xử lý ảnh chứa dữ liệu GPS và công suất từ các tủ BTS (Base Transceiver Station). Hệ thống tích hợp nhiều API cloud (Google Cloud Vision, OpenAI, Google Sheets, Google Cloud Storage) với backend FastAPI để trích xuất, phân tích và lưu trữ dữ liệu cơ sở hạ tầng viễn thông.

**Điểm vào chính:** `gsmnv.py` (FastAPI webhook server trên cổng 8007)

Hệ thống hỗ trợ hai nguồn ảnh:
- **N8N workflow**: Webhook `/webhook` nhận URL ảnh trực tiếp
- **Telegram group**: Webhook `/telegram-webhook` nhận ảnh từ group Telegram qua GCS

## Kiến Trúc & Luồng Dữ Liệu

### Luồng từ Telegram (mới)

1. **Telegram Group** → Bot nhận ảnh từ group
2. **Download ảnh** → Lấy ảnh từ Telegram server qua Bot API
3. **Upload GCS** → `gcs_storage.py` upload ảnh lên Google Cloud Storage
4. **Trích xuất metadata** → `telegram_bot.py` trích xuất tên người gửi, caption, message_id, timestamp
5. **Gọi webhook chính** → Gửi image URL từ GCS + metadata tới `/webhook` (xử lý giống n8n)
6. **Xử lý như n8n** → Tiếp tục quy trình OCR, phân tích AI, lưu Google Sheets

### Luồng từ N8N (cũ)

1. **N8N Workflow** → Gửi ảnh + metadata tới webhook `/webhook`
2. **Xử lý OCR** → Google Cloud Vision API trích xuất text từ ảnh
3. **Phân tích AI** → OpenAI GPT-4o-mini phân tích text thành JSON có cấu trúc:
   - Tên tủ (định dạng: H-ABC/xxxx hoặc O-ABC/xxxx)
   - Tọa độ GPS (vĩ độ, kinh độ)
   - Ngày và giờ (định dạng Việt chuẩn hóa thành YYYY-MM-DD HH:MM:SS)
   - Giá trị công suất (định dạng dBm, có thể cần chuyển đổi từ số 4 chữ số)
4. **Tính toán khoảng cách** → `distance_bts.py` tính toán khoảng cách Haversine đến vị trí tủ cụ thể
5. **Lưu trữ dữ liệu** → Kết quả được lưu vào Google Sheets + gửi thông báo qua webhook
6. **Đầu ra** → Phản hồi JSON trả về workflow n8n

### Các Module Chính

- **`gsmnv.py`** - Ứng dụng FastAPI chính xử lý các webhook và điều phối đường ống xử lý
- **`distance_bts.py`** - Tính toán khoảng cách Haversine giữa tọa độ GPS và vị trí tủ (sử dụng `ket_qua_gop.xlsx` làm nguồn dữ liệu)
- **`telegram_bot.py`** - Xử lý webhook từ Telegram, trích xuất metadata, gọi webhook chính
- **`gcs_storage.py`** - Upload ảnh từ Telegram lên Google Cloud Storage
- **`setup_telegram_webhook.py`** - Script setup webhook với Telegram
- **`ggsh.py`** - Tiện ích ví dụ cho tương tác Google Sheets API (không dùng nữa, chức năng được chuyển sang `gsmnv.py`)

### Nguồn Dữ Liệu

- **`ket_qua_gop.xlsx`** - Dữ liệu vị trí tủ với các cột: "Tên kết cuối", "Vĩ độ" (latitude), "Kinh độ" (longitude)
- **`map_gps_all_bts.xlsx`** - Dữ liệu trạm BTS tham khảo (hiện không sử dụng trong luồng chính)
- **Google Sheets** - Lưu trữ dữ liệu trực tiếp (ID sheet hardcoded trong code, worksheet "thang10")

## Các Lệnh Thường Dùng

### Cài Đặt Thư Viện Phụ Thuộc
```bash
pip install -r requirements.txt
```

### Chạy API Server
```bash
python gsmnv.py
```
Khởi động FastAPI trên `http://0.0.0.0:8007`

### Setup Telegram Webhook (lần đầu tiên)
```bash
# Với default config (từ biến môi trường)
python setup_telegram_webhook.py

# Hoặc với environment variables tùy chỉnh
TELEGRAM_BOT_TOKEN="your-token" \
WEBHOOK_DOMAIN="your-domain.com" \
GCS_BUCKET_NAME="your-bucket" \
python setup_telegram_webhook.py
```

Lệnh con (không chạy workflow chính):
```bash
# Xóa webhook
python setup_telegram_webhook.py delete

# Xem thông tin webhook hiện tại
python setup_telegram_webhook.py info

# Setup GCS bucket
python setup_telegram_webhook.py gcs
```

### Kiểm Tra Webhook Endpoint (N8N)
```bash
curl -X POST http://localhost:8007/webhook \
  -H "Content-Type: application/json" \
  -d '{
    "threadId": "test-123",
    "name": "Tên Người Dùng",
    "title": "Tiêu Đề Test",
    "image_url": "https://example.com/image.jpg"
  }'
```

### Kiểm Tra Telegram Webhook Endpoint
```bash
curl -X POST https://your-domain.com/telegram-webhook \
  -H "Content-Type: application/json" \
  -d '{
    "message": {
      "message_id": 123,
      "date": 1700000000,
      "chat": {"id": -4863386433, "title": "Test Group"},
      "from": {"id": 123, "first_name": "Test", "username": "testuser"},
      "photo": [{"file_id": "YOUR_FILE_ID"}],
      "caption": "Test"
    }
  }'
```

### Chạy Kiểm Tra Tính Toán Khoảng Cách
```bash
python distance_bts.py
```
Thực thi các hàm kiểm tra trong khối `if __name__ == "__main__"`

### Kiểm Tra GCS Storage Module
```bash
python gcs_storage.py
```

### Kiểm Tra Telegram Bot Module
```bash
python telegram_bot.py
```

### Kiểm Tra Sức Khỏe API
```bash
curl http://localhost:8007/
```

## Cấu Hình & Thông Tin Nhạy Cảm

### File Credentials
Ứng dụng yêu cầu những tệp này ở thư mục gốc dự án:

- **`vision-key.json`** - Thông tin đăng nhập Google Cloud Vision API (dùng chung cho GCS)
- **`ggsheet-key.json`** - Khóa tài khoản dịch vụ Google Sheets API
- **`openai-vison-key.txt`** - Khóa API OpenAI (tệp văn bản, một dòng)

### Environment Variables (để cấu hình động)

```bash
# Telegram
export TELEGRAM_BOT_TOKEN="<your-telegram-bot-token>"
export TELEGRAM_CHANNEL_ID="-4863386433"

# Google Cloud Storage
export GCS_BUCKET_NAME="bts-telegram-images"
export GCS_PROJECT_ID="your-gcp-project-id"

# Webhook chính (nếu khác default)
export MAIN_WEBHOOK_URL="http://localhost:8007/webhook"

# Domain webhook
export WEBHOOK_DOMAIN="n8n2.ttvt8.online"
export WEBHOOK_PORT="443"
```

### Hardcoded Config
Những giá trị được hardcode trong code (nên chuyển sang env):

**trong `gsmnv.py`:**
```python
WEBHOOK_URL = "https://your-webhook.example.com/webhook/dhsc-text"  # Line 26
sheet = client.open_by_key("10mFy9EzRNG2VvOOWnZe0Tl8OK2QnN3cJXTJiMaeUK9Q").worksheet("thang10")  # Line 231
```

**trong `telegram_bot.py`:**
```python
TELEGRAM_BOT_TOKEN = "<your-telegram-bot-token>"  # Line 12
TELEGRAM_CHANNEL_ID = "-4863386433"  # Line 13
MAIN_WEBHOOK_URL = "https://your-webhook.example.com/webhook/dhsc-text"  # Line 18
```

**trong `gcs_storage.py`:**
```python
GCS_BUCKET_NAME = "bts-telegram-images"  # Line 13
GCS_PROJECT_ID = "your-project-id"  # Line 14
```

## Chi Tiết Triển Khai Chính

### Luồng Xử Lý Telegram

1. **Webhook nhận** (`gsmnv.py:438-467`)
   - Endpoint `/telegram-webhook` nhận update từ Telegram

2. **Trích xuất metadata** (`telegram_bot.py:42-87`)
   - Lấy ảnh có kích thước lớn nhất từ mảng `photo`
   - Trích xuất: sender_name, sender_username, caption, message_id, chat_id, timestamp

3. **Download ảnh** (`telegram_bot.py:89-125`)
   - Gọi Telegram Bot API `getFile` để lấy file path
   - Download ảnh từ `https://api.telegram.org/file/bot{token}/{file_path}`

4. **Upload GCS** (`gcs_storage.py:39-96`)
   - Sử dụng credentials từ `vision-key.json`
   - Upload ảnh với path: `telegram-images/{YYYY}/{MM}/{DD}/{filename}.jpg`
   - Tạo public URL: `https://storage.googleapis.com/{bucket}/{path}`

5. **Gọi webhook chính** (`telegram_bot.py:126-173`)
   - Gửi POST tới `/webhook` với payload:
     ```json
     {
       "threadId": "message_id",
       "name": "sender_name",
       "title": "caption",
       "image_url": "GCS public URL",
       "source": "telegram",
       "telegram_metadata": {...}
     }
     ```

6. **Xử lý như n8n** - Tiếp tục quy trình OCR, phân tích AI, lưu sheets

### Kỹ Thuật Prompt OpenAI

Prompt hệ thống trong `gsmnv.py:79-138` rất quan trọng để trích xuất dữ liệu chính xác. Nó xử lý:

- **Các biến thể tên tủ**: "H-ABC/xxxx", "O-ABC/xxxx", hoặc "ABC/xxxx" (tự động thêm tiền tố "H-")
- **Phân tích tọa độ**: Trích xuất độ thập phân có hoặc không có ký hiệu ° và chỉ báo hướng (°N, °E)
- **Định dạng giá trị công suất**: Chuyển đổi giá trị dBm có định dạng đa dòng hoặc khác nhau (ví dụ: "-1700" → "-17.00dBm")
- **Chuẩn hóa ngày/giờ**: Chuyển đổi định dạng Việt thành tiêu chuẩn ISO (YYYY-MM-DD HH:MM:SS)

### Tính Toán Khoảng Cách

Công thức Haversine (`distance_bts.py:5-36`) tính toán khoảng cách trên mặt cầu (great-circle distance) tính bằng kilometers. Thuật toán khớp (`distance_bts.py:65-70`) sử dụng khớp chuỗi từng phần trên cột "Tên kết cuối" để tìm tủ chính xác.

### Xử Lý Lỗi

- Lỗi Vision API mặc định trả về chuỗi rỗng
- Lỗi phân tích OpenAI được ghi log với cảnh báo nhưng không dừng thực thi
- Thiếu trường bắt buộc trong tọa độ sẽ bỏ qua thông báo nhưng vẫn lưu vào sheets
- Lỗi Google Sheets trả về False nhưng được ghi log với đầy đủ traceback
- Lỗi upload GCS trả về error dict nhưng không dừng webhook Telegram
- Lỗi download từ Telegram được log và trả về error response

## Setup Telegram Bot (Bước Đầu Tiên)

### 1. Tạo Google Cloud Storage Bucket (nếu chưa có)

```bash
# Trên Google Cloud Console:
# 1. Vào Cloud Storage > Buckets
# 2. Tạo bucket mới với tên "bts-telegram-images" (hoặc tên khác)
# 3. Chọn region: us-central1 (hoặc asia-southeast1 gần nhất)
# 4. Permissions: Public Read
#    - Đi tới bucket > Permissions
#    - Thêm role "Storage Object Viewer" cho "allUsers"
#    - Hoặc chạy: gsutil iam ch allUsers:objectViewer gs://bucket-name
```

### 2. Setup Webhook Telegram (lần đầu)

```bash
# Cài đặt thư viện
pip install -r requirements.txt

# Setup webhook (theo hướng dẫn script)
TELEGRAM_BOT_TOKEN="<your-telegram-bot-token>" \
WEBHOOK_DOMAIN="n8n2.ttvt8.online" \
GCS_BUCKET_NAME="bts-telegram-images" \
GCS_PROJECT_ID="your-gcp-project-id" \
python setup_telegram_webhook.py

# Script sẽ hỏi:
# - Bạn muốn setup GCS bucket? (y/n)
# - Nó sẽ tạo bucket nếu chưa tồn tại
# - Đăng ký webhook với Telegram
# - Kiểm tra webhook info
```

### 3. Chạy API Server

```bash
python gsmnv.py
```

### 4. Kiểm Tra Setup

```bash
# Xem webhook info
python setup_telegram_webhook.py info

# Output sẽ hiển thị:
# - url: https://n8n2.ttvt8.online/telegram-webhook
# - allowed_updates: ["message"]
# - pending_update_count: 0
```

### 5. Test với Telegram

- Mở group Telegram (hoặc @botfather với bot của bạn)
- Gửi ảnh vào group
- Kiểm tra logs API để xem ảnh có được xử lý không

## Ghi Chú Phát Triển

### Các Sửa Đổi Thường Gặp

**Telegram:**
- **Thay token bot**: Cập nhật `TELEGRAM_BOT_TOKEN` trong `telegram_bot.py:12` hoặc `setup_telegram_webhook.py`
- **Thay channel ID**: Cập nhật `TELEGRAM_CHANNEL_ID` trong `telegram_bot.py:13`
- **Thay webhook URL**: Cập nhật `MAIN_WEBHOOK_URL` trong `telegram_bot.py:18`

**GCS:**
- **Thay bucket**: Cập nhật `GCS_BUCKET_NAME` trong `gcs_storage.py:13`
- **Thay project**: Cập nhật `GCS_PROJECT_ID` trong `gcs_storage.py:14`

**Webhook xử lý chính:**
- **Thay đổi webhook thông báo**: Cập nhật `WEBHOOK_URL` trong `gsmnv.py:26`
- **Thay đổi mục tiêu Google Sheets**: Cập nhật ID sheet và tên worksheet trong `gsmnv.py:231`

**AI:**
- **Điều chỉnh mô hình AI**: Thay đổi tham số `model` trong `gsmnv.py:142` (hiện tại là "gpt-4o-mini")
- **Sửa đổi định dạng dữ liệu**: Cập nhật prompt hệ thống trong `gsmnv.py:79-138` hoặc logic phân tích trong `save_to_google_sheets()`

### Các Cân Nhắc Kiểm Tra

**Telegram:**
- Kiểm tra bot có thể truy cập group Telegram không
- Kiểm tra file_id có hợp lệ (thường là chuỗi dài base64-like)
- Test download ảnh từ Telegram server

**GCS:**
- Kiểm tra credentials từ `vision-key.json` có quyền upload không
- Kiểm tra bucket có public read permission không
- Test tạo public URL của ảnh upload

**OCR/AI:**
- Kiểm tra với ảnh thực tế có chất lượng OCR khác nhau
- Xác minh khớp tên tủ hoạt động với các định dạng dòng Excel khác nhau
- Kiểm tra hàm làm sạch tọa độ (`clean_coordinate()` trong `gsmnv.py:278-281`) với các ký hiệu độ Unicode khác nhau
- Xác thực chuyển đổi dBm 4 chữ số (ví dụ: "-1700" phải thành "-17.00dBm")
- Kiểm tra xử lý múi giờ nếu cần điều chỉnh dấu thời gian

### Những Hạn Chế Đã Biết

- Khớp tên tủ dựa trên chuỗi từng phần; không có khớp mức độ mềm dẻo cho lỗi chính tả
- Kết nối Google Sheets là đồng bộ, có thể timeout khi kết nối tệ
- Không có logic thử lại cho các lệnh gọi API bị lỗi đến các dịch vụ bên ngoài
- Các giá trị cấu hình hardcoded nên được chuyển sang các biến môi trường
- Không có tính năng lưu trữ cơ sở dữ liệu; phụ thuộc hoàn toàn vào Google Sheets làm kho dữ liệu
- Upload GCS yêu cầu public read permission trên bucket (không an toàn cho production)
- Telegram Bot API file URLs chỉ có hiệu lực trong 1 giờ
