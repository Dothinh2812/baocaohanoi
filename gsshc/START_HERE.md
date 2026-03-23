# 🚀 START HERE - Run Application

## ✅ Mọi Thứ Đã Sẵn Sàng!

Tất cả các server cũ đã bị dừng. Bây giờ bạn có thể chạy ứng dụng của riêng mình.

---

## 🎯 Cách Chạy Ứng Dụng

### Option 1: Sử Dụng Script (Dễ Nhất)

```bash
cd /home/vtst/s2
./run.sh
```

**Output sẽ hiển thị:**
```
╔════════════════════════════════════════╗
║  Starting Application                  ║
║  Port: 8007                            ║
╚════════════════════════════════════════╝

INFO:     Started server process [XXXX]
INFO:     Waiting for application startup.
INFO:     Application startup complete.
INFO:     Uvicorn running on http://0.0.0.0:8007
```

### Option 2: Chạy Với Port Khác

```bash
./run.sh 8009
```

### Option 3: Manual Activation

```bash
cd /home/vtst/s2

# Activate virtual environment
source venv/bin/activate

# Chạy ứng dụng
python3 gsmnv.py

# Hoặc với custom port
PORT=8009 python3 gsmnv.py
```

---

## 🧪 Test Application

### 1. Health Check
```bash
curl https://gsmnvs2.ttvt8.online/
```

**Expected Response:**
```json
{"message":"Webhook API is running"}
```

### 2. Send Image to Telegram
- Mở Telegram
- Gửi ảnh vào group
- Chờ 3-5 giây
- Bot sẽ tự động phản hồi ✅

### 3. Check Results
- **Google Sheets:** Dữ liệu được lưu
- **GCS Bucket:** Ảnh được upload
- **Telegram:** Nhận tin nhắn kết quả

---

## 🛑 Stop Application

Khi muốn dừng server:

```bash
# Nhấn Ctrl+C trong terminal

# Hoặc từ terminal khác
pkill -f "python3.*gsmnv"

# Hoặc kill specific port
lsof -i :8007 | grep -v COMMAND | awk '{print $2}' | xargs kill -9
```

---

## 📊 Port Information

| Port | Status | Usage |
|------|--------|-------|
| 8007 | ✅ Free | FastAPI default |
| 8009 | ✅ Free | Custom port option |
| 8000 | ✅ Free | Alternative |

---

## 📁 Project Structure

```
/home/vtst/s2/
├── venv/                    # Virtual environment (263 MB)
├── gsmnv.py                 # Main application
├── telegram_bot.py          # Telegram handler
├── gcs_storage.py           # GCS manager
├── distance_bts.py          # Distance calculation
├── run.sh                   # Run script ✅
├── requirements.txt         # Dependencies
└── [Documentation files]
```

---

## 🔧 Virtual Environment Commands

### Activate
```bash
source venv/bin/activate
```

### Deactivate
```bash
deactivate
```

### Check Python
```bash
python3 --version
which python3
```

### List Packages
```bash
pip list
```

---

## 📞 Features Included

✅ **FastAPI Server** - Port 8007
✅ **Telegram Webhook** - gsmnvs2.ttvt8.online/telegram-webhook
✅ **Google Cloud Vision** - OCR processing
✅ **OpenAI Analysis** - GPT-4o-mini
✅ **Distance Calculation** - Haversine formula
✅ **Google Sheets Integration** - Data storage
✅ **Google Cloud Storage** - Image storage
✅ **Automatic Responses** - Telegram replies

---

## 📖 Documentation

| File | Purpose |
|------|---------|
| `COMPLETE_SETUP.md` | Complete setup guide |
| `VENV_SETUP.md` | Virtual environment guide |
| `VENV_SUMMARY.md` | Quick venv summary |
| `TELEGRAM_RESPONSE_CONFIG.md` | Telegram setup |
| `TEST_TELEGRAM_RESPONSE.md` | Testing guide |
| `CLAUDE.md` | Developer guide |

---

## ⚡ Quick Start (30 seconds)

```bash
cd /home/vtst/s2
./run.sh
```

Then test:
```bash
curl https://gsmnvs2.ttvt8.online/
```

---

## 🎯 What Happens When You Run

1. Activates virtual environment
2. Starts FastAPI server on port 8007
3. Server listens for webhooks:
   - `/webhook` - N8N webhook
   - `/telegram-webhook` - Telegram webhook
   - `/` - Health check

4. Ready to receive images from:
   - Telegram bot
   - N8N workflow

---

## ✨ Ready!

**Bạn đã sẵn sàng để bắt đầu! 🚀**

```bash
./run.sh
```

---

**Status: ✅ READY TO RUN**
