# 🚀 Quick Start - Gửi Phản Hồi Telegram

## ✨ Tính Năng Mới

**Bot sẽ tự động gửi tin nhắn phản hồi tới Telegram group sau khi xử lý ảnh!**

---

## 🎯 Bắt Đầu Ngay (2 Phút)

### Bước 1: Kiểm Tra Server
```bash
curl https://gsmnvs2.ttvt8.online/
# Kết quả: {"message":"Webhook API is running"}
```

### Bước 2: Mở Telegram

### Bước 3: Gửi Ảnh
1. Vào group Telegram
2. Gửi một bức ảnh (ảnh BTS tốt nhất)

### Bước 4: Nhận Phản Hồi
3-5 giây sau, bot sẽ gửi tin nhắn:
```
📊 KẾT QUẢ PHÂN TÍCH GPS

👤 Người gửi: [Tên của bạn]
🗄️ Tủ BTS: [Tên tủ]
📅 Ngày: [Ngày]
🕐 Thời gian: [Giờ]
🌍 Tọa độ: [GPS]
📏 Khoảng cách: [Meters]
⚡ Công suất: [dBm]

✓ Dữ liệu đã lưu vào hệ thống
```

---

## 🔧 Cấu Hình

### Bot Token
```
<your-telegram-bot-token>
```

### Group ID
```
-4863386433
```

### Thay Đổi (Nếu Cần)

**Option 1: Environment Variables**
```bash
export TELEGRAM_BOT_TOKEN="new-token"
export TELEGRAM_CHANNEL_ID="-123456"
python3 /home/vtst/s2/gsmnv.py
```

**Option 2: Edit Code**
File: `/home/vtst/s2/gsmnv.py` Line 34-35

---

## 📊 Tin Nhắn Ví Dụ

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

## 🧪 Test Nhanh

```bash
# Gửi test message
python3 -c "
import requests
BOT_TOKEN = '<your-telegram-bot-token>'
CHAT_ID = '-4863386433'
msg = '<b>✅ Telegram Phản Hồi Hoạt Động!</b>'
requests.post(f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage',
              json={'chat_id': CHAT_ID, 'text': msg, 'parse_mode': 'HTML'})
print('Test message sent!')
"
```

---

## ✅ Server Status

- **Port:** 8007
- **Domain:** https://gsmnvs2.ttvt8.online
- **Status:** ✅ Running
- **PID:** 306207

---

## 📝 Code Thay Đổi

### Hàm Mới 1: `send_telegram_message()`
Gửi tin nhắn tới Telegram

### Hàm Mới 2: `format_telegram_response()`
Format kết quả phân tích thành tin nhắn đẹp

### Modified: `save_to_google_sheets()`
Thêm code gửi Telegram sau khi lưu Sheets

---

## 🔄 Workflow

```
Gửi ảnh Telegram
       ↓
Download + Upload GCS
       ↓
Vision API
       ↓
OpenAI Analysis
       ↓
Save Google Sheets
       ↓
SEND TELEGRAM RESPONSE ← NEW!
       ↓
Nhận tin nhắn trong Telegram
```

---

## 📞 Troubleshooting

### "Không nhận được tin nhắn"
```bash
# Kiểm tra bot
curl https://api.telegram.org/bot<your-telegram-bot-token>/getMe

# Kiểm tra webhook
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo
```

### "Server error"
```bash
# Kiểm tra logs
tail -50 /var/log/python_app.log

# Restart server
kill 306207
python3 /home/vtst/s2/gsmnv.py &
```

---

## 📚 Documentation

- `TELEGRAM_RESPONSE_CONFIG.md` - Chi tiết cấu hình
- `TEST_TELEGRAM_RESPONSE.md` - Hướng dẫn test
- `TELEGRAM_RESPONSE_SUMMARY.md` - Tóm tắt đầy đủ

---

## ✨ Features

✅ **Instant Response** - Phản hồi trong 3-5 giây
✅ **Beautiful Format** - Tin nhắn đẹp với emojis
✅ **Complete Data** - Tất cả thông tin cần thiết
✅ **Error Handling** - Xử lý lỗi tốt
✅ **No Disturbance** - Không ảnh hưởng tới chức năng cũ

---

## 🚀 Bây Giờ

**Hãy gửi ảnh test tới Telegram! 📸**

Bạn sẽ nhận được phản hồi tự động trong vòng vài giây! ⚡

---

**Status: ✅ PRODUCTION READY**
