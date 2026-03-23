# 📱 Cấu Hình Gửi Phản Hồi tới Telegram

## ✅ Tính Năng Mới

Sau khi OpenAI phân tích ảnh thành công, kết quả sẽ **tự động gửi trả lại Telegram group** dưới dạng một tin nhắn đẹp với:

- 👤 Tên người gửi
- 🗄️ Tên tủ BTS
- 📅 Ngày giờ
- 🌍 Tọa độ GPS (vĩ độ, kinh độ)
- 📏 Khoảng cách tới tủ
- ⚡ Công suất sau S2

---

## 🔧 Cấu Hình

### Telegram Bot Token
```
TELEGRAM_BOT_TOKEN: <your-telegram-bot-token>
```

### Telegram Channel/Group ID
```
TELEGRAM_CHANNEL_ID: -4863386433
```

### Cách Thay Đổi (Nếu Cần)

#### Option 1: Environment Variables
```bash
export TELEGRAM_BOT_TOKEN="your-token-here"
export TELEGRAM_CHANNEL_ID="-123456789"

python3 /home/vtst/s2/gsmnv.py
```

#### Option 2: Sửa Code
Chỉnh sửa trong `gsmnv.py` line 34-35:
```python
TELEGRAM_BOT_TOKEN = "your-token-here"
TELEGRAM_CHANNEL_ID = "-123456789"
```

---

## 💬 Ví Dụ Tin Nhắn Telegram

Khi gửi ảnh, bạn sẽ nhận được tin nhắn như sau:

```
📊 KẾT QUẢ PHÂN TÍCH GPS

👤 Người gửi: Nguyễn Văn A

🗄️ Tủ BTS: H-BVI/2024

📅 Ngày: 2025-10-20

🕐 Thời gian: 14:35:45

🌍 Tọa độ:
  • Vĩ độ: 21.1234
  • Kinh độ: 105.5678

📏 Khoảng cách: 1234 mét

⚡ Công suất sau S2: -20.13dBm

✓ Dữ liệu đã lưu vào hệ thống
```

---

## 🔄 Luồng Xử Lý (Mới)

```
┌─────────────────────────────────────┐
│ Telegram Group                      │
│ (Send Image)                        │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ Bot receives image                  │
│ POST /telegram-webhook              │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ 1. Extract metadata                 │
│ 2. Download from Telegram           │
│ 3. Upload to GCS                    │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ 4. OCR (Vision API)                 │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ 5. AI Analysis (OpenAI)             │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ 6. Distance Calculation             │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ 7. Save to Google Sheets            │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ 8. SEND TELEGRAM MESSAGE ✅ (NEW)   │
└────────────┬────────────────────────┘
             │
             ↓
┌─────────────────────────────────────┐
│ Telegram Group                      │
│ (Receive Response)                  │
└─────────────────────────────────────┘
```

---

## 📝 Hàm Chính

### 1. `send_telegram_message(chat_id, message, parse_mode)`
Gửi tin nhắn tới Telegram

**Parameters:**
- `chat_id`: Chat ID hoặc Group ID (-4863386433)
- `message`: Nội dung tin nhắn (HTML format)
- `parse_mode`: "HTML", "Markdown", "MarkdownV2" (default: "HTML")

**Return:** True/False

### 2. `format_telegram_response(data)`
Format kết quả phân tích thành tin nhắn đẹp

**Parameters:**
- `data`: Dictionary chứa:
  - `cabinet_name`: Tên tủ
  - `lat`: Vĩ độ
  - `long`: Kinh độ
  - `date`: Ngày
  - `time`: Giờ
  - `power_after_s2`: Công suất
  - `distance_m`: Khoảng cách
  - `sender_name`: Tên người gửi

**Return:** String tin nhắn HTML format

---

## 🧪 Test Gửi Tin Nhắn Trực Tiếp

```bash
python3 -c "
import requests
import json

BOT_TOKEN = '<your-telegram-bot-token>'
CHAT_ID = '-4863386433'

message = '''
<b>📊 TEST MESSAGE</b>

<b>👤 Người gửi:</b> Nguyễn Văn A

<b>🗄️ Tủ BTS:</b> <code>H-BVI/2024</code>

<b>🌍 Tọa độ:</b> 21.1234, 105.5678

✓ Test thành công!
'''

url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage'
payload = {
    'chat_id': CHAT_ID,
    'text': message,
    'parse_mode': 'HTML'
}

response = requests.post(url, json=payload)
print(json.dumps(response.json(), indent=2))
"
```

---

## 📊 Lưu Trữ Kết Quả

### Google Sheets
- Vẫn được lưu vào Google Sheets như cũ
- Tất cả dữ liệu được lưu trữ

### Telegram History
- Tin nhắn được lưu trong Telegram group
- Có thể xem lại lịch sử

### Google Cloud Storage
- Ảnh được lưu trữ ở GCS

---

## ⚠️ Troubleshooting

### Lỗi: "Telegram message failed"

**Nguyên nhân:**
1. Bot token không hợp lệ
2. Chat ID không đúng
3. Bot không có quyền gửi message tới group

**Cách khắc phục:**
```bash
# Kiểm tra bot có active không
curl https://api.telegram.org/bot<your-telegram-bot-token>/getMe

# Kiểm tra webhook
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo
```

### Lỗi: "Chat not found"

**Nguyên nhân:**
- Chat ID không hợp lệ
- Bot không có quyền truy cập group

**Cách khắc phục:**
1. Đảm bảo bot đã được thêm vào group
2. Kiểm tra Chat ID là negative (ví dụ: -4863386433, không phải 4863386433)
3. Cấp quyền admin cho bot

### Lỗi: "Unauthorized"

**Nguyên nhân:**
- Token hết hiệu lực
- Token không đúng

**Cách khắc phục:**
1. Tạo token mới từ @BotFather
2. Cập nhật TELEGRAM_BOT_TOKEN

---

## 🎯 Tính Năng Bổ Sung (Có Thể Thêm)

### 1. Gửi Ảnh Trực Tiếp
```python
def send_telegram_photo(chat_id, photo_url, caption=""):
    # Gửi ảnh từ GCS trực tiếp tới Telegram
    pass
```

### 2. Gửi Thông Báo Inline Button
```python
def send_telegram_with_buttons(chat_id, message, buttons):
    # Gửi tin nhắn với button để user bấm
    pass
```

### 3. Reply Thread
```python
def send_telegram_reply(chat_id, message, reply_to_message_id):
    # Reply vào tin nhắn gốc
    pass
```

---

## 📞 Hỗ Trợ

Nếu cần thêm tính năng hoặc gặp lỗi:

1. Kiểm tra logs: `tail -f /var/log/python_app.log`
2. Kiểm tra Telegram API docs: https://core.telegram.org/bots/api
3. Test hàm gửi tin nhắn độc lập

---

## ✅ Tóm Lại

**Hệ thống bây giờ:**

1. ✅ Nhận ảnh từ Telegram
2. ✅ Xử lý OCR + AI
3. ✅ Lưu vào Google Sheets
4. ✅ **Gửi kết quả trả lại Telegram** (NEW)
5. ✅ Gửi thông báo N8N

**Bạn sẽ thấy phản hồi ngay lập tức trong Telegram group!** 🎉
