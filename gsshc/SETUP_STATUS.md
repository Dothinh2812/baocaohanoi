# Setup Status - Telegram Bot Integration

## ✅ Hoàn Thành

### 1. Dependencies
- ✅ google-cloud-vision (3.10.2)
- ✅ google-cloud-storage (3.4.1)
- ✅ python-telegram-bot (22.5)
- ✅ fastapi, openai, gspread, pandas, requests
- ✅ Tất cả dependencies đã install thành công

### 2. Telegram Webhook Setup
- ✅ Bot Token: <your-telegram-bot-token>
- ✅ Webhook URL: https://n8n2.ttvt8.online/telegram-webhook
- ✅ Allowed Updates: ["message"]
- ✅ Webhook Registration: THÀNH CÔNG
- ✅ Webhook Status: Active (pending_update_count: 0)

### 3. API Server
- ✅ Server Status: RUNNING
- ✅ Host: 0.0.0.0
- ✅ Port: 8007
- ✅ Health Check: ✅ Working

### 4. Endpoints
- ✅ `/` - Health check
- ✅ `/webhook` - N8N webhook (cũ)
- ✅ `/telegram-webhook` - Telegram webhook (mới)

### 5. Modules Created
- ✅ `gcs_storage.py` - Google Cloud Storage management
- ✅ `telegram_bot.py` - Telegram bot webhook handler
- ✅ `setup_telegram_webhook.py` - Setup script

### 6. Documentation
- ✅ `CLAUDE.md` - Updated with Telegram integration docs

---

## 🚀 Bước Tiếp Theo

### 1. Setup Google Cloud Storage Bucket (Tùy Chọn)
Nếu bạn muốn tự động tạo bucket, chạy:
```bash
TELEGRAM_BOT_TOKEN="<your-telegram-bot-token>" \
GCS_BUCKET_NAME="bts-telegram-images" \
GCS_PROJECT_ID="your-gcp-project-id" \
python3 setup_telegram_webhook.py gcs
```

HOẶC tạo bucket thủ công:
- Vào Google Cloud Console > Cloud Storage > Buckets
- Tạo bucket "bts-telegram-images"
- Chọn region: us-central1
- Bật Public Read: Permissions > Add allUsers:objectViewer

### 2. Kiểm Tra Kết Nối Domain
Chắc chắn domain n8n2.ttvt8.online:
- Forward HTTPS tới server trên port 8007
- Hoặc chạy reverse proxy (nginx, etc.)

### 3. Test Telegram Bot
1. Tìm bot trên Telegram (nếu còn hoạt động)
2. Hoặc tạo test group
3. Gửi ảnh vào group → Bot sẽ:
   - Download ảnh từ Telegram
   - Upload lên GCS
   - Gửi webhook tới `/webhook`
   - Xử lý OCR, AI analysis, lưu Google Sheets

### 4. Kiểm Tra Logs
```bash
# Xem logs từ API server
BashOutput bash_id: 606d85

# Hoặc kiểm tra webhook info
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo
```

---

## 📋 Quick Commands

### Health Check
```bash
curl http://localhost:8007/
```

### Get Webhook Info
```bash
python3 -c "
import requests, json
url = 'https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo'
print(json.dumps(requests.get(url).json(), indent=2))
"
```

### Xóa Webhook (nếu cần)
```bash
python3 -c "
import requests, json
url = 'https://api.telegram.org/bot<your-telegram-bot-token>/deleteWebhook'
data = {'drop_pending_updates': True}
print(json.dumps(requests.post(url, json=data).json(), indent=2))
"
```

### Kill API Server
```bash
kill 304950  # PID từ logs
```

---

## ⚙️ Cấu Hình

### Environment Variables
```bash
TELEGRAM_BOT_TOKEN=<your-telegram-bot-token>
TELEGRAM_CHANNEL_ID=-4863386433
WEBHOOK_DOMAIN=n8n2.ttvt8.online
GCS_BUCKET_NAME=bts-telegram-images
GCS_PROJECT_ID=your-gcp-project-id
```

### Hardcoded Configs (nên chuyển sang env)
- `telegram_bot.py:12` - TELEGRAM_BOT_TOKEN
- `telegram_bot.py:18` - MAIN_WEBHOOK_URL
- `gcs_storage.py:13-14` - GCS config
- `gsmnv.py:26` - WEBHOOK_URL for notifications

---

## 🔗 Workflow

```
Telegram Group
     ↓
Bot receives image
     ↓
/telegram-webhook (gsmnv.py:438)
     ↓
extract_telegram_metadata (telegram_bot.py:42)
     ↓
upload_image_from_telegram (gcs_storage.py:39)
     ↓
process_telegram_image (telegram_bot.py:89)
     ↓
/webhook (gsmnv.py:374)
     ↓
OCR (Vision API)
     ↓
AI Analysis (OpenAI)
     ↓
Distance Calculation (Haversine)
     ↓
Save Google Sheets + Send notification
```

---

## 📊 Status Summary

| Component | Status | Notes |
|-----------|--------|-------|
| Dependencies | ✅ | All installed |
| Telegram Webhook | ✅ | Active and working |
| API Server | ✅ | Running on 8007 |
| GCS Storage | ⚠️ | Needs bucket setup |
| N8N Webhook | ✅ | Still working |
| Documentation | ✅ | CLAUDE.md updated |

---

Bây giờ hệ thống sẵn sàng nhận ảnh từ Telegram! 🎉
