# Server Status - Telegram Bot API

## ✅ Server Status

### API Server
- **Status**: ✅ RUNNING
- **Port**: 8009
- **Host**: 0.0.0.0
- **PID**: 305469

### Local Access
- ✅ `http://localhost:8009/` - **WORKING**
- ✅ Health check endpoint responds

### Telegram Webhook
- ✅ **URL**: https://gsmnvs2.ttvt8.online/telegram-webhook
- ✅ **Status**: Active (set successfully)
- ✅ **Allowed Updates**: ["message"]
- ✅ **Pending Updates**: 0
- ⚠️ **Issue**: Domain returning 502 error

---

## 🔧 Current Issues

### 1. Cloudflare Tunnel 502 Error
```
curl https://gsmnvs2.ttvt8.online/
error code: 502
```

**Possible causes**:
- Cloudflare tunnel not connected to port 8009
- Tunnel configuration pointing to wrong port
- Network connectivity issue

**Fix Options**:
1. Check Cloudflare tunnel config - should point to `http://localhost:8009`
2. Restart Cloudflare tunnel
3. Verify tunnel status: `cloudflared tunnel list` or dashboard

---

## 📋 Endpoints Available

| Endpoint | Method | Status | Purpose |
|----------|--------|--------|---------|
| `/` | GET | ✅ | Health check |
| `/webhook` | POST | ✅ (local) | N8N webhook |
| `/telegram-webhook` | POST | ✅ (local) | Telegram webhook |

---

## 🚀 Quick Commands

### Check Server Status
```bash
curl http://localhost:8009/
```

### Check Telegram Webhook
```bash
curl https://api.telegram.org/bot<your-telegram-bot-token>/getWebhookInfo
```

### View Server Logs
```bash
# Check background process logs
# BashOutput bash_id: ca4d70
```

### Kill Server
```bash
kill 305469
```

---

## 🔍 Debug Steps

### 1. Check Cloudflare Tunnel Connection
```bash
# SSH to server
# Check if tunnel is running
systemctl status cloudflared
# or
ps aux | grep cloudflared

# Check tunnel config
cat /etc/cloudflared/config.yml
```

### 2. Verify Port 8009
```bash
# Check if port 8009 is listening
netstat -tlnp | grep 8009
# or
lsof -i :8009
```

### 3. Test Local Connection
```bash
curl http://localhost:8009/webhook -X POST -H "Content-Type: application/json" -d '{"test":"data"}'
```

### 4. Check Cloudflare Dashboard
- Go to Cloudflare dashboard
- Check tunnel status
- Verify route: `gsmnvs2.ttvt8.online` → `http://localhost:8009`

---

## 📝 Configuration

### Current Setup
```
Telegram Bot → Telegram API
     ↓
https://gsmnvs2.ttvt8.online/telegram-webhook
     ↓ (Cloudflare tunnel - ERROR 502)
http://localhost:8009/telegram-webhook
     ↓
Process image, upload GCS, save to Sheets
```

### Expected Setup
```
Telegram Bot → Telegram API
     ↓
https://gsmnvs2.ttvt8.online/telegram-webhook
     ✅ (should work via Cloudflare tunnel)
     ↓
http://localhost:8009/telegram-webhook
     ↓
Process image, upload GCS, save to Sheets
```

---

## 🎯 Next Steps

**URGENT**: Fix Cloudflare tunnel connection to port 8009

1. SSH to your server
2. Check Cloudflare tunnel status
3. Verify tunnel configuration points to `http://localhost:8009`
4. Test: `curl https://gsmnvs2.ttvt8.online/` should return `{"message":"Webhook API is running"}`

Once tunnel is working:
- Telegram will automatically send images to bot
- Bot will process them through the full pipeline

---

## Server Information

- **Process ID**: 305469
- **Port**: 8009
- **Framework**: FastAPI + Uvicorn
- **Background Process ID**: ca4d70

### How to Stop/Start Server
```bash
# Kill current server
kill 305469

# Start new server
PORT=8009 python3 /home/vtst/s2/gsmnv.py
```

---

**Last Update**: 2025-10-20 07:59:46 UTC
**Status**: Waiting for Cloudflare tunnel fix
