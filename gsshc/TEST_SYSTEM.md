# 🧪 Hướng Dẫn Kiểm Tra Hệ Thống

## 1️⃣ Kiểm Tra Server Hoạt Động

### Cách 1: Health Check
```bash
curl https://gsmnvs2.ttvt8.online/
```

**Kết quả mong đợi:**
```json
{"message":"Webhook API is running"}
```

### Cách 2: Từ Local Machine
```bash
curl http://localhost:8007/
```

---

## 2️⃣ Kiểm Tra Telegram Webhook

### Xem Webhook Info
```bash
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo
```

**Kết quả mong đợi:**
```json
{
  "ok": true,
  "result": {
    "url": "https://gsmnvs2.ttvt8.online/telegram-webhook",
    "has_custom_certificate": false,
    "pending_update_count": 0,
    "max_connections": 40,
    "allowed_updates": ["message"]
  }
}
```

---

## 3️⃣ Kiểm Tra Real-Time (Cách Tốt Nhất)

### Bước 1: Xem Server Logs
```bash
# Mở terminal thứ 2 trên server
tail -f /tmp/telegram_bot.log
# hoặc kiểm tra background process
# Bash ID: 6f823f
```

### Bước 2: Gửi Ảnh Test từ Telegram
1. Mở Telegram → Tìm bot
2. Gửi ảnh vào group hoặc chat riêng với bot
3. Ngay lập tức xem logs

### Bước 3: Kiểm Tra Logs
Bạn sẽ thấy:
```
============================================================
TELEGRAM WEBHOOK RECEIVED
============================================================
Update data: {...}

Step 1: Trích xuất metadata từ Telegram
✅ Trích xuất metadata từ Telegram:
   - Người gửi: Your Name
   - Caption: Your caption
   - Message ID: 123456

Step 2: Upload ảnh lên Google Cloud Storage
Step 1: Lấy file path từ Telegram (file_id: AgAC...)
✅ File path: photos/XXX
Step 2: Tải ảnh từ Telegram server
✅ Đã tải ảnh: 12345 bytes
Step 3: Upload ảnh lên Google Cloud Storage
✅ Đã upload lên GCS: telegram-images/2025/10/20/telegram_123456_username.jpg
✅ Public URL: https://storage.googleapis.com/bts-telegram-images/...

Step 3: Chuẩn bị dữ liệu gửi đến webhook chính
Step 4: Gửi dữ liệu tới webhook xử lý chính
✅ Gửi webhook thành công: Status 200

============================================================
TELEGRAM PROCESSING COMPLETED
============================================================
```

---

## 4️⃣ Kiểm Tra Google Sheets

### Bước 1: Vào Google Sheets
- Mở: https://docs.google.com/spreadsheets/
- Tìm sheet ID: `10mFy9EzRNG2VvOOWnZe0Tl8OK2QnN3cJXTJiMaeUK9Q`
- Hoặc kiểm tra trong code: gsmnv.py line 231

### Bước 2: Kiểm Tra Dữ Liệu
Hàng mới sẽ được thêm với các cột:
- name (Tên người gửi)
- threadId (Message ID)
- title (Caption)
- cabinet_name (Tên tủ BTS)
- extracted_date (Ngày)
- extracted_time (Giờ)
- lat (Vĩ độ)
- long (Kinh độ)
- distance_m (Khoảng cách)
- power_after_s2 (Công suất)
- image_url (URL ảnh GCS)
- created_date (Ngày tạo)
- created_time (Giờ tạo)

---

## 5️⃣ Kiểm Tra Google Cloud Storage

### Vào GCS Console
1. Mở: https://console.cloud.google.com/storage
2. Tìm bucket: `bts-telegram-images`
3. Kiểm tra folder: `telegram-images/{YYYY}/{MM}/{DD}/`
4. Ảnh sẽ có tên: `telegram_MESSAGE_ID_USERNAME.jpg`

---

## 6️⃣ Kiểm Tra N8N Notification

### Webhook Destination
- URL: https://your-webhook.example.com/webhook/dhsc-text
- Payload sẽ được gửi như:
```
📍 Kết quả phân tích GPS
👤 Người gửi: Tên người dùng
🗄️ Tủ: H-ABC/1234
📅 Ngày: 2025-10-20
🕐 Thời gian: 14:30:45
🌍 Tọa độ: 21.1234, 105.5678
📏 Khoảng cách so với capman: 1234 mét
⚡ Công suất sau S2: -20.13dBm
```

---

## 7️⃣ Kiểm Tra Lỗi (Nếu Có)

### Nếu Server Không Phản Hồi
```bash
# Kiểm tra port 8007 có chạy không
lsof -i :8007

# Kiểm tra Cloudflare tunnel
ps aux | grep cloudflared

# Restart server
kill 305716
python3 /home/vtst/s2/gsmnv.py &
```

### Nếu Telegram Webhook Không Nhận Được
```bash
# Kiểm tra webhook info
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo

# Kiểm tra pending updates
# Nếu có pending_update_count > 0, hãy:
curl https://api.telegram.org/bot<your-telegram-bot-token>/deleteWebhook \
  -X POST -H "Content-Type: application/json" \
  -d '{"drop_pending_updates": true}'

# Rồi set lại webhook
curl https://api.telegram.org/bot<your-telegram-bot-token>/setWebhook \
  -X POST -H "Content-Type: application/json" \
  -d '{"url": "https://gsmnvs2.ttvt8.online/telegram-webhook", "allowed_updates": ["message"]}'
```

### Nếu GCS Upload Thất Bại
```bash
# Kiểm tra credentials
cat /home/vtst/s2/vision-key.json | python3 -m json.tool

# Test GCS connection
python3 /home/vtst/s2/gcs_storage.py
```

### Nếu OpenAI Phân Tích Thất Bại
```bash
# Kiểm tra API key
cat /home/vtst/s2/openai-vison-key.txt

# Test OpenAI
python3 -c "
from openai import OpenAI
api_key = open('/home/vtst/s2/openai-vison-key.txt').read().strip()
client = OpenAI(api_key=api_key)
print('✅ OpenAI connection OK')
"
```

---

## 8️⃣ Test End-to-End (Recommended)

### Script Test Đơn Giản
```bash
#!/bin/bash

echo "🧪 Testing Telegram Bot Integration"
echo ""

# Test 1: Server health
echo "1️⃣ Testing server health..."
curl -s https://gsmnvs2.ttvt8.online/ && echo "✅ Server OK" || echo "❌ Server failed"
echo ""

# Test 2: Webhook info
echo "2️⃣ Testing Telegram webhook..."
curl -s https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo \
  | python3 -c "import sys, json; data=json.load(sys.stdin); print('✅ Webhook active' if data.get('ok') else '❌ Webhook failed')"
echo ""

# Test 3: Port check
echo "3️⃣ Testing port 8007..."
lsof -i :8007 > /dev/null && echo "✅ Port 8007 running" || echo "❌ Port 8007 not running"
echo ""

# Test 4: Credentials check
echo "4️⃣ Checking credentials..."
[ -f "/home/vtst/s2/vision-key.json" ] && echo "✅ vision-key.json found" || echo "❌ vision-key.json missing"
[ -f "/home/vtst/s2/ggsheet-key.json" ] && echo "✅ ggsheet-key.json found" || echo "❌ ggsheet-key.json missing"
[ -f "/home/vtst/s2/openai-vison-key.txt" ] && echo "✅ openai-vison-key.txt found" || echo "❌ openai-vison-key.txt missing"
echo ""

echo "✅ Test completed!"
```

---

## 🎯 Quy Trình Kiểm Tra Đầy Đủ

### Cách tốt nhất để kiểm tra:

1. **Xem logs trực tiếp:**
```bash
# Terminal 1: Watch logs
tail -f /tmp/telegram.log
```

2. **Gửi ảnh test từ Telegram:**
   - Mở Telegram
   - Gửi ảnh vào group

3. **Xem kết quả trong logs:**
   - Kiểm tra "TELEGRAM WEBHOOK RECEIVED"
   - Kiểm tra "Vision API - Text detected"
   - Kiểm tra "Google Sheets append result"

4. **Kiểm tra Google Sheets:**
   - Mở sheet
   - Kiểm tra hàng mới

5. **Kiểm tra GCS:**
   - Mở bucket
   - Xem ảnh upload

---

## 📞 Nếu Có Lỗi

Hãy chạy và share output:
```bash
# Check everything
echo "=== Server Status ==="
lsof -i :8007

echo "=== Webhook Info ==="
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo

echo "=== Credentials ==="
ls -la /home/vtst/s2/*.json /home/vtst/s2/*-key.txt

echo "=== Server Logs ==="
# Check background process logs
```

---

**Bây giờ hãy thử gửi ảnh test từ Telegram! 📸**
