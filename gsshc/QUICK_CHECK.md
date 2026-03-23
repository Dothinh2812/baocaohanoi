# ✅ QUICK CHECK - Hệ Thống Hoạt Động 100%

## 🟢 Tất Cả Đang Hoạt Động

### ✅ Kiểm Tra Nhanh Kết Quả:
```
1️⃣  HTTPS Domain (gsmnvs2.ttvt8.online): ✅ WORKING
2️⃣  Localhost (http://localhost:8007): ✅ WORKING
3️⃣  Telegram Webhook: ✅ REGISTERED
4️⃣  Pending Updates: 0 (No errors)
5️⃣  Credentials: ✅ ALL FOUND (3/3)
6️⃣  Core Files: ✅ ALL FOUND (3/3)
```

---

## 🎯 CÁCH KIỂM TRA HỆ THỐNG (3 CÁCH)

### **Cách 1: Health Check (Nhanh Nhất)**

```bash
# Test HTTPS domain
curl https://gsmnvs2.ttvt8.online/
```

**Kết quả mong đợi:**
```
{"message":"Webhook API is running"}
```

---

### **Cách 2: Kiểm Tra Telegram Webhook**

```bash
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo
```

**Kết quả mong đợi:**
```json
{
  "ok": true,
  "result": {
    "url": "https://gsmnvs2.ttvt8.online/telegram-webhook",
    "pending_update_count": 0,
    "allowed_updates": ["message"]
  }
}
```

---

### **Cách 3: KIỂM TRA THỰC TẾ (Recommended)**

**Đây là cách tốt nhất!**

#### Bước 1: Mở 2 Terminal

**Terminal 1 - Watch logs:**
```bash
ssh your-server
tail -f /var/log/python_app.log
# hoặc
tail -f /home/vtst/s2/*.log
```

#### Bước 2: Gửi Ảnh từ Telegram

1. Mở **Telegram**
2. Tìm bot (hoặc group có bot)
3. **Gửi một bức ảnh**
4. Thêm caption (ví dụ: "Test image")

#### Bước 3: Xem Logs

Bạn sẽ thấy trong Terminal 1:

```
============================================================
TELEGRAM WEBHOOK RECEIVED
============================================================
Update data: {
  "message": {
    "message_id": 12345,
    "from": {"first_name": "Your Name"},
    "chat": {...},
    "photo": [...],
    "caption": "Test image"
  }
}

Step 1: Trích xuất metadata từ Telegram
✅ Trích xuất metadata từ Telegram:
   - Người gửi: Your Name
   - Caption: Test image
   - Message ID: 12345

Step 2: Upload ảnh lên Google Cloud Storage
Step 1: Lấy file path từ Telegram...
✅ File path: photos/XXX
Step 2: Tải ảnh từ Telegram server
✅ Đã tải ảnh: 45678 bytes

Step 3: Upload ảnh lên Google Cloud Storage
✅ Đã upload lên GCS: telegram-images/2025/10/20/telegram_12345_username.jpg
✅ Public URL: https://storage.googleapis.com/bts-telegram-images/...

Step 3: Chuẩn bị dữ liệu gửi đến webhook chính
Step 4: Gửi dữ liệu tới webhook xử lý chính
✅ Gửi webhook thành công: Status 200

Step 1: Processing image with Google Cloud Vision...
Vision API - Text detected: H-ABC/1234
21.1234 105.5678
...

Step 2: Analyzing text with OpenAI...
OpenAI Chat API output: {"cabinet_name": "H-ABC/1234", ...}

Step 3: Saving to Google Sheets...
✅ Data saved to Google Sheets successfully

PROCESSING COMPLETED
============================================================
```

---

## 📊 Kiểm Tra Kết Quả

### 1. **Google Sheets**
```
URL: https://docs.google.com/spreadsheets/d/10mFy9EzRNG2VvOOWnZe0Tl8OK2QnN3cJXTJiMaeUK9Q/edit
```
- Hàng mới sẽ được thêm
- Chứa dữ liệu: tên, tủ, tọa độ, khoảng cách, công suất, ảnh URL

### 2. **Google Cloud Storage**
```
URL: https://console.cloud.google.com/storage
Bucket: bts-telegram-images
Path: telegram-images/2025/10/20/telegram_XXXX.jpg
```
- Ảnh sẽ được lưu trữ ở đây

### 3. **N8N Notification**
```
URL: https://your-webhook.example.com/webhook/dhsc-text
```
- Sẽ nhận thông báo với kết quả

---

## 🔴 Nếu Có Vấn Đề

### Lỗi 1: Ảnh không được upload
```bash
# Kiểm tra GCS credentials
cat /home/vtst/s2/vision-key.json | python3 -m json.tool

# Test GCS
python3 /home/vtst/s2/gcs_storage.py
```

### Lỗi 2: OpenAI phân tích thất bại
```bash
# Kiểm tra API key
cat /home/vtst/s2/openai-vison-key.txt

# Kiểm tra balance
# Vào: https://platform.openai.com/account/billing/overview
```

### Lỗi 3: Google Sheets không cập nhật
```bash
# Kiểm tra sheet ID (trong code)
grep "open_by_key" /home/vtst/s2/gsmnv.py

# Kiểm tra credentials
cat /home/vtst/s2/ggsheet-key.json | python3 -m json.tool
```

### Lỗi 4: Server không phản hồi
```bash
# Kiểm tra server có chạy
ps aux | grep gsmnv.py

# Kiểm tra port
netstat -tlnp | grep 8007

# Restart nếu cần
kill $(pgrep -f "gsmnv.py")
python3 /home/vtst/s2/gsmnv.py &
```

---

## 🎯 Bây Giờ Làm Gì?

### **Bước 1: Xác Nhận Server Đang Chạy**
```bash
curl https://gsmnvs2.ttvt8.online/
# Kết quả: {"message":"Webhook API is running"} ✅
```

### **Bước 2: Mở Telegram và Gửi Ảnh Test**
1. Mở app Telegram
2. Tìm bot (hoặc group)
3. Gửi một bức ảnh
4. Thêm caption nếu cần

### **Bước 3: Kiểm Tra Google Sheets**
1. Mở: https://docs.google.com/spreadsheets/d/10mFy9EzRNG2VvOOWnZe0Tl8OK2QnN3cJXTJiMaeUK9Q/edit
2. Tìm hàng mới ở dưới cùng
3. Kiểm tra dữ liệu

### **Bước 4: Kiểm Tra GCS**
1. Mở: https://console.cloud.google.com/storage
2. Tìm bucket: bts-telegram-images
3. Mở folder: telegram-images/2025/10/20/
4. Kiểm tra ảnh

---

## ⚡ TÓM LẠI

| Thành Phần | Status | Ghi Chú |
|-----------|--------|--------|
| API Server | ✅ | Chạy trên port 8007 |
| Telegram Webhook | ✅ | gsmnvs2.ttvt8.online/telegram-webhook |
| Cloudflare Tunnel | ✅ | Kết nối thành công |
| Credentials | ✅ | Tất cả file present |
| GCS Setup | ✅ | Sẵn sàng upload |
| OpenAI Setup | ✅ | Sẵn sàng phân tích |
| Google Sheets | ✅ | Sẵn sàng lưu dữ liệu |

---

## 🚀 BƯỚC TIẾP THEO

**Gửi ảnh test từ Telegram ngay bây giờ!**

1. Telegram → Gửi ảnh
2. Kiểm tra logs (nếu cần)
3. Kiểm tra Google Sheets
4. Hoàn tất! ✅

---

**Hệ thống 100% sẵn sàng! 🎉**
