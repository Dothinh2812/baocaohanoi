# 🧪 Test Tính Năng Gửi Phản Hồi Telegram

## ✅ Sẵn Sàng Kiểm Tra

Hệ thống đã được cập nhật để gửi kết quả phân tích trực tiếp tới Telegram group.

---

## 🎯 Cách Kiểm Tra (3 Cách)

### **Cách 1: Test Nhanh - Gửi Tin Nhắn Test**

```bash
python3 -c "
import requests

BOT_TOKEN = '<your-telegram-bot-token>'
CHAT_ID = '-4863386433'

message = '''<b>🧪 TEST MESSAGE</b>

Nếu bạn thấy tin nhắn này trong Telegram, hệ thống hoạt động! ✅'''

url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage'
payload = {'chat_id': CHAT_ID, 'text': message, 'parse_mode': 'HTML'}

response = requests.post(url, json=payload)
print('✅ Test message sent!' if response.json().get('ok') else '❌ Failed')
"
```

**Kết quả mong đợi:**
- ✅ Nhận tin nhắn trong Telegram group: "🧪 TEST MESSAGE - Nếu bạn thấy tin nhắn này..."

---

### **Cách 2: Test Thực Tế - Gửi Ảnh Thực Tế (Tốt Nhất)**

**Bước 1:** Mở Telegram → Vào group

**Bước 2:** Gửi một bức ảnh có chứa:
- Tên tủ BTS (ví dụ: H-BVI/2024)
- Tọa độ GPS (ví dụ: 21.1234, 105.5678)
- Ngày giờ
- Công suất dBm (ví dụ: -20.13dBm)

**Bước 3:** Chờ 3-5 giây

**Bước 4:** Kiểm tra:
- ✅ Tin nhắn phản hồi trong Telegram (tin nhắn đẹp với emojis)
- ✅ Hàng mới trong Google Sheets
- ✅ Ảnh trong Google Cloud Storage

---

### **Cách 3: Test từ Command Line**

**Gửi ảnh test:**
```bash
curl -X POST https://gsmnvs2.ttvt8.online/webhook \
  -H "Content-Type: application/json" \
  -d '{
    "threadId": "test-123",
    "name": "Nguyễn Văn A",
    "title": "Test image",
    "image_url": "https://example.com/test.jpg"
  }'
```

---

## 📝 Dữ Liệu Sẽ Gửi Tới Telegram

```
📊 KẾT QUẢ PHÂN TÍCH GPS

👤 Người gửi: [Tên người gửi từ Telegram]

🗄️ Tủ BTS: [Tên tủ được OCR]

📅 Ngày: [Ngày từ ảnh]

🕐 Thời gian: [Giờ từ ảnh]

🌍 Tọa độ:
  • Vĩ độ: [Vĩ độ từ ảnh]
  • Kinh độ: [Kinh độ từ ảnh]

📏 Khoảng cách: [Khoảng cách tính được] mét

⚡ Công suất sau S2: [Công suất từ ảnh]

✓ Dữ liệu đã lưu vào hệ thống
```

---

## 🔍 Kiểm Tra Logs

Để xem chi tiết quá trình xử lý, xem logs:

```bash
# Xem background process logs
tail -f /var/log/python_app.log

# Hoặc kiểm tra output trực tiếp
BashOutput bash_id: b42775
```

**Bạn sẽ thấy:**
```
TELEGRAM WEBHOOK RECEIVED
Step 1: Trích xuất metadata từ Telegram
✅ Trích xuất metadata từ Telegram...

Step 2: Upload ảnh lên Google Cloud Storage
✅ Đã upload lên GCS...

Step 3: Processing image with Google Cloud Vision
Vision API - Text detected: ...

Step 4: Analyzing text with OpenAI
OpenAI Chat API output: {...}

Step 5: Saving to Google Sheets
✅ Data saved to Google Sheets successfully

Step 6: Gửi kết quả phân tích tới Telegram
✅ Gửi tin nhắn Telegram thành công
```

---

## 📊 Kiểm Tra Kết Quả

### 1. Google Sheets
- URL: https://docs.google.com/spreadsheets/d/10mFy9EzRNG2VvOOWnZe0Tl8OK2QnN3cJXTJiMaeUK9Q/edit
- Kiểm tra hàng mới ở dưới cùng
- Dữ liệu sẽ có đầy đủ: tên, tủ, tọa độ, khoảng cách, công suất, URL ảnh

### 2. Google Cloud Storage
- URL: https://console.cloud.google.com/storage
- Bucket: bts-telegram-images
- Path: telegram-images/2025/10/20/telegram_XXXX.jpg
- Ảnh sẽ được lưu trữ

### 3. Telegram Group
- Group ID: -4863386433
- Bạn sẽ nhận được tin nhắn phản hồi
- Tin nhắn chứa tất cả kết quả phân tích

---

## ✨ Tính Năng Chi Tiết

### Lợi Ích:
- ✅ User nhận phản hồi **ngay lập tức** trong Telegram
- ✅ Không cần mở Google Sheets để kiểm tra
- ✅ Tin nhắn có format đẹp với emojis
- ✅ Dễ hiểu và dễ chia sẻ

### Format:
- ✅ HTML format với `<b>bold</b>`, `<code>code</code>`
- ✅ Emojis để dễ nhận biết
- ✅ Thông tin được sắp xếp rõ ràng

---

## 🎯 Workflow Hoàn Chỉnh

```
Telegram (Gửi ảnh)
         ↓
Bot nhận ảnh
         ↓
Download từ Telegram
         ↓
Upload GCS
         ↓
Vision API (OCR)
         ↓
OpenAI (Phân tích)
         ↓
Tính toán khoảng cách
         ↓
Lưu Google Sheets
         ↓
Gửi Telegram (RESPONSE) ← NEW
         ↓
Telegram (Nhận phản hồi)
```

---

## 🚀 Bây Giờ Làm Gì?

### Option 1: Test Nhanh (1 phút)
```bash
# Terminal
python3 -c "
import requests
BOT_TOKEN = '<your-telegram-bot-token>'
CHAT_ID = '-4863386433'
msg = '<b>✅ System is working!</b>'
requests.post(f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage',
              json={'chat_id': CHAT_ID, 'text': msg, 'parse_mode': 'HTML'})
"

# Telegram - Kiểm tra nhận tin nhắn
```

### Option 2: Test Thực Tế (3 phút)
1. Mở Telegram
2. Gửi ảnh BTS vào group
3. Chờ phản hồi

### Option 3: Test Chi Tiết (10 phút)
1. Gửi ảnh
2. Kiểm tra Telegram (phản hồi)
3. Kiểm tra Google Sheets (dữ liệu)
4. Kiểm tra GCS (ảnh)
5. Xem logs (chi tiết)

---

## 📞 Nếu Có Vấn Đề

### Lỗi: Không nhận được tin nhắn Telegram

**Kiểm tra:**
```bash
# 1. Bot có hoạt động?
curl https://api.telegram.org/bot<your-telegram-bot-token>/getMe

# 2. Webhook có hoạt động?
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo

# 3. Chat ID đúng không? (phải là negative)
echo "-4863386433"
```

### Lỗi: Server trả về lỗi

**Kiểm tra logs:**
```bash
# Xem logs mới nhất
tail -50 /var/log/python_app.log

# Hoặc kiểm tra background process
ps aux | grep gsmnv.py
```

---

## ✅ Xác Nhận Hoạt Động

**Hệ thống sẵn sàng khi:**
- ✅ `curl https://gsmnvs2.ttvt8.online/` → Returns JSON
- ✅ Bot token hợp lệ
- ✅ Chat ID đúng
- ✅ Server running trên port 8007

**Bạn sẽ biết thành công khi:**
- ✅ Gửi ảnh → Nhận phản hồi trong Telegram trong 3-5 giây
- ✅ Hàng mới xuất hiện trong Google Sheets
- ✅ Ảnh được lưu trong GCS

---

**Hãy test ngay! 🚀**
