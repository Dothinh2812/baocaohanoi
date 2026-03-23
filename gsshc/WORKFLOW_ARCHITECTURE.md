# 🔄 WORKFLOW ARCHITECTURE - Kiến Trúc Luồng Xử Lý

## ✅ Xác Nhận: Luồng Tuân Thủ Cấu Trúc Gốc

Hệ thống hiện tại **hoàn toàn tuân thủ** cấu trúc luồng xử lý ban đầu:
```
Upload GCS → Vision API OCR → OpenAI Analysis → Google Sheets → Telegram Response
```

---

## 📊 Luồng Chi Tiết (End-to-End)

### **Cấp 1: Telegram Webhook (Điểm Bắt Đầu)**

**Endpoint:** `POST /telegram-webhook`
**File:** `gsmnv.py:541`

User gửi ảnh vào Telegram group → Telegram Bot gọi webhook endpoint

```python
@app.post("/telegram-webhook")
async def telegram_webhook(request: dict):
    # Webhook từ Telegram được gửi tới đây
```

---

### **Cấp 2: Trích Xuất Metadata**

**File:** `telegram_bot.py:26-92`
**Function:** `extract_telegram_metadata(update)`

Trích xuất thông tin từ Telegram message:
- `file_id` - ID ảnh để download từ Telegram server
- `sender_name` - Tên người gửi
- `message_id` - ID tin nhắn
- `chat_id` - ID group

```python
metadata = {
    "file_id": "AgACAgIAAxkBAAIF...",
    "sender_name": "Nguyễn Văn A",
    "message_id": "1001",
    "chat_id": "-4863386433"
}
```

---

### **Cấp 3: Upload ảnh lên Google Cloud Storage (GCS)**

**File:** `telegram_bot.py:95-137`
**Function:** `process_telegram_image(update)` → Bước 2
**File:** `gcs_storage.py:11-80`
**Function:** `upload_image_from_telegram(file_id, bot_token, filename)`

**Luồng chi tiết:**

1. Tải ảnh từ Telegram server:
   ```python
   # https://api.telegram.org/file/botTOKEN/AgACAgIAAxkBAAIF...
   response = requests.get(f"https://api.telegram.org/file/bot{bot_token}/{file_path}")
   ```

2. Upload lên Google Cloud Storage:
   ```python
   bucket.blob(destination_blob_name).upload_from_string(response.content)
   ```

3. Nhận được public URL từ GCS:
   ```
   https://storage.googleapis.com/bts-telegram-images/telegram-images/2025/10/20/file.jpg
   ```

**Output:** `image_url` - URL công khai của ảnh trên GCS

---

### **Cấp 4: Gửi tới Webhook Xử Lý Chính**

**File:** `telegram_bot.py:138-184`
**Function:** `process_telegram_image(update)` → Bước 3-4

Gửi dữ liệu với GCS image URL tới webhook chính:

```python
webhook_payload = {
    "threadId": message_id,
    "name": sender_name,
    "title": caption,
    "image_url": "https://storage.googleapis.com/...",  # ← GCS URL
    "source": "telegram",
    "telegram_metadata": {...}
}

response = requests.post(MAIN_WEBHOOK_URL, json=webhook_payload)
```

**Webhook URL:** `https://your-webhook.example.com/webhook/dhsc-text` hoặc `http://localhost:8007/webhook` (nếu chạy local)

---

### **Cấp 5: Webhook Xử Lý Chính**

**Endpoint:** `POST /webhook`
**File:** `gsmnv.py:485-539`
**Function:** `receive_webhook(data: WebhookData)`

Tiếp nhận payload và bắt đầu xử lý:

```python
@app.post("/webhook")
async def receive_webhook(data: WebhookData):
    # data.image_url = "https://storage.googleapis.com/..."

    # Step 1: Xử lý ảnh với Vision API
    vision_text = process_image_with_vision(data.image_url)

    # Step 2: Phân tích text với OpenAI
    openai_analysis = analyze_text_with_openai(vision_text, data.name)

    # Step 3: Lưu vào Google Sheets
    save_to_google_sheets(...)
```

---

### **Cấp 6: Google Cloud Vision API - OCR**

**File:** `gsmnv.py:139-161`
**Function:** `process_image_with_vision(image_url)`

**Luồng chi tiết:**

1. Tạo đối tượng ảnh từ GCS URL:
   ```python
   image = vision.Image()
   image.source.image_uri = image_url  # ← GCS URL
   ```

2. Gửi yêu cầu OCR text detection tới Google Cloud Vision API:
   ```python
   response = vision_client.text_detection(image=image)
   ```

3. Nhận kết quả text được extract từ ảnh:
   ```
   "H-BVI/2024\n21.1234°N\n105.5678°E\n2025-10-20\n14:30:45\ndBm\n-2013"
   ```

**Input:** GCS image URL
**Output:** `vision_text` - Text extracted từ ảnh

---

### **Cấp 7: OpenAI Analysis - GPT-4o-mini**

**File:** `gsmnv.py:163-262`
**Function:** `analyze_text_with_openai(text, name)`

**Luộc chi tiết:**

1. Gửi extracted text tới OpenAI Chat Completions API:
   ```python
   response = client.chat.completions.create(
       model="gpt-4o-mini",
       messages=[
           {"role": "system", "content": system_prompt},
           {"role": "user", "content": vision_text}  # ← Extracted text từ Vision API
       ],
       response_format={"type": "json_object"}
   )
   ```

2. OpenAI phân tích text theo system prompt và trả về JSON:
   ```json
   {
       "cabinet_name": "H-BVI/2024",
       "lat": "21.1234",
       "long": "105.5678",
       "date": "2025-10-20",
       "time": "14:30:45",
       "power_after_s2": "-20.13dBm"
   }
   ```

**Input:** `vision_text` - Text extracted từ Vision API
**Output:** `openai_analysis` - JSON string với dữ liệu phân tích

---

### **Cấp 8: Lưu vào Google Sheets**

**File:** `gsmnv.py:311-483`
**Function:** `save_to_google_sheets(...)`

**Luồc chi tiết:**

1. Parse JSON từ OpenAI analysis:
   ```python
   parsed_data = json.loads(json_str)
   cabinet_name = parsed_data.get("cabinet_name", "")
   lat = parsed_data.get("lat", "")
   long = parsed_data.get("long", "")
   # ... other fields
   ```

2. Tính toán khoảng cách (nếu có tọa độ):
   ```python
   cabinet_result = find_nearest_bts_station(lat, long, cabinet_name)
   distance_m = round(cabinet_result['distance_km'] * 1000)
   ```

3. Chuẩn bị row data:
   ```python
   row_data = [
       name,              # Người gửi
       thread_id,         # Message ID
       title,             # Caption
       cabinet_name,      # Tủ BTS (từ OpenAI)
       extracted_date,    # Ngày (từ OpenAI)
       extracted_time,    # Thời gian (từ OpenAI)
       lat,              # Latitude (từ OpenAI)
       long,             # Longitude (từ OpenAI)
       distance_m,       # Khoảng cách tính toán
       power_after_s2,   # Công suất (từ OpenAI)
       image_url,        # GCS URL
       created_date,     # Ngày lưu
       created_time      # Thời gian lưu
   ]
   ```

4. Ghi vào Google Sheets:
   ```python
   sheet.append_row(row_data)
   ```

**Input:** `openai_analysis` + các dữ liệu từ webhook
**Output:** Data saved to Google Sheets

---

### **Cấp 9: Gửi Telegram Response**

**File:** `gsmnv.py:97-137` (format_telegram_response)
**File:** `gsmnv.py:63-95` (send_telegram_message)
**Called from:** `gsmnv.py:427-448`

Sau khi lưu thành công vào Sheets, gửi kết quả tới Telegram group:

**Bước 9.1: Format tin nhắn**
```python
telegram_data = {
    "cabinet_name": "H-BVI/2024",
    "lat": "21.1234",
    "long": "105.5678",
    "date": "2025-10-20",
    "time": "14:30:45",
    "power_after_s2": "-20.13dBm",
    "distance_m": "1234",
    "sender_name": "Nguyễn Văn A"
}

telegram_message = format_telegram_response(telegram_data)
```

**Bước 9.2: Gửi tin nhắn**
```python
send_telegram_message(TELEGRAM_CHANNEL_ID, telegram_message)
```

**Output:** Tin nhắn HTML đẹp gửi tới Telegram group

---

## 🔗 Lưu Đồ Luồng Xử Lý

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Telegram Group                                           │
│    User gửi ảnh BTS tới group                               │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. Telegram Bot Webhook                                     │
│    POST /telegram-webhook                                   │
│    telegram_bot.py:95 - process_telegram_image()            │
└────────────────────┬────────────────────────────────────────┘
                     │
        ┌────────────┴────────────┐
        ▼                         ▼
    ┌────────────────┐    ┌─────────────────────┐
    │ Trích xuất     │    │ Extract metadata    │
    │ metadata       │    │ from Telegram msg   │
    └────────┬───────┘    └─────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. Upload ảnh lên GCS                                       │
│    gcs_storage.py:11 - upload_image_from_telegram()         │
│                                                             │
│    • Download ảnh từ Telegram server                        │
│    • Upload lên Google Cloud Storage                        │
│    • Nhận public URL                                        │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼ (GCS public URL)
┌─────────────────────────────────────────────────────────────┐
│ 4. Gửi payload tới webhook chính                            │
│    telegram_bot.py:165                                      │
│    requests.post(MAIN_WEBHOOK_URL)                          │
│                                                             │
│    payload = {                                              │
│        "threadId": message_id,                              │
│        "name": sender_name,                                 │
│        "image_url": "https://storage.googleapis.com/...",   │
│        ...                                                  │
│    }                                                        │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ 5. Webhook xử lý chính                                      │
│    POST /webhook                                            │
│    gsmnv.py:485 - receive_webhook()                         │
└────────────────────┬────────────────────────────────────────┘
                     │
        ┌────────────┴────────────┐
        ▼                         ▼
┌──────────────────────┐  ┌──────────────────────┐
│ STEP 1:              │  │ STEP 2:              │
│ Vision API OCR       │  │ OpenAI Analysis      │
│                      │  │                      │
│ gsmnv.py:139         │  │ gsmnv.py:163         │
│                      │  │                      │
│ Input:               │  │ Input:               │
│ GCS image URL        │  │ Vision extracted     │
│                      │  │ text                 │
│ Output:              │  │                      │
│ vision_text          │  │ Output:              │
│ (extracted text)     │  │ openai_analysis      │
│                      │  │ (JSON with fields)   │
└──────────┬───────────┘  └──────────┬───────────┘
           │                         │
           └─────────────┬───────────┘
                         ▼
┌─────────────────────────────────────────────────────────────┐
│ STEP 3: Lưu vào Google Sheets                               │
│ gsmnv.py:311 - save_to_google_sheets()                      │
│                                                             │
│ • Parse JSON từ OpenAI                                      │
│ • Tính toán khoảng cách                                     │
│ • Chuẩn bị row data                                         │
│ • Ghi vào Google Sheets                                     │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ STEP 4: Gửi Telegram Response                               │
│ gsmnv.py:427-448                                            │
│                                                             │
│ • format_telegram_response() - tạo tin nhắn                │
│ • send_telegram_message() - gửi tới Telegram group          │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ ✅ COMPLETED                                                │
│                                                             │
│ • Google Sheets: Dữ liệu lưu                                │
│ • GCS: Ảnh lưu                                              │
│ • Telegram: Tin nhắn kết quả nhận được                       │
└─────────────────────────────────────────────────────────────┘
```

---

## 📝 Chi Tiết Từng Giai Đoạn

### **Giai Đoạn 1: GCS Upload**

**Mục đích:** Lưu trữ ảnh trên cloud, nhận public URL để Vision API xử lý

**File:** `gcs_storage.py:11-80`

```python
def upload_image_from_telegram(file_id, bot_token, filename):
    """
    1. Tải ảnh từ Telegram server
    2. Upload lên Google Cloud Storage
    3. Trả về public URL
    """

    # Bước 1: Tạo GCS client
    client = get_gcs_client()
    bucket = client.bucket(GCS_BUCKET_NAME)

    # Bước 2: Download ảnh từ Telegram
    telegram_file_url = f"https://api.telegram.org/file/bot{bot_token}/{file_path}"
    response = requests.get(telegram_file_url)

    # Bước 3: Upload lên GCS
    blob = bucket.blob(destination_path)
    blob.upload_from_string(response.content)

    # Bước 4: Trả về public URL
    public_url = f"https://storage.googleapis.com/{GCS_BUCKET_NAME}/{destination_path}"
    return {"public_url": public_url}
```

**Input:**
- `file_id` - Telegram file ID
- `bot_token` - Telegram bot token
- `filename` - Tên file lưu trữ

**Output:**
```json
{
    "public_url": "https://storage.googleapis.com/bts-telegram-images/telegram-images/2025/10/20/telegram_1001_testuser.jpg"
}
```

---

### **Giai Đoạn 2: Vision API OCR**

**Mục đích:** Extract text từ ảnh (OCR)

**File:** `gsmnv.py:139-161`

```python
def process_image_with_vision(image_url):
    """
    1. Tạo Image object từ GCS URL
    2. Gọi Vision API text detection
    3. Trả về extracted text
    """

    # Bước 1: Tạo Image object
    image = vision.Image()
    image.source.image_uri = image_url  # ← GCS URL

    # Bước 2: Gọi Vision API
    response = vision_client.text_detection(image=image)

    # Bước 3: Extract text
    texts = response.text_annotations
    if texts:
        detected_text = texts[0].description
        return detected_text
    else:
        return ""
```

**Input:**
```
image_url = "https://storage.googleapis.com/bts-telegram-images/.../file.jpg"
```

**Output:**
```
"H-BVI/2024
21.1234°N
105.5678°E
Thứ Hai, 13 Tháng 10,2025
11:03
dBm
-2013"
```

---

### **Giai Đoạn 3: OpenAI Analysis**

**Mục đích:** Phân tích extracted text, extract structured data

**File:** `gsmnv.py:163-262`

```python
def analyze_text_with_openai(text, name):
    """
    1. Gửi extracted text + system prompt tới OpenAI
    2. OpenAI phân tích và trả về JSON structured data
    3. Parse và validate JSON
    """

    system_prompt = """Extract information from OCR text:
    - cabinet_name: H-ABC/xxxx
    - lat: latitude coordinate
    - long: longitude coordinate
    - date: YYYY-MM-DD format
    - time: HH:MM:SS format
    - power_after_s2: -XXdBm format
    Return as JSON only."""

    # Gửi tới OpenAI
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text}  # ← Vision extracted text
        ],
        response_format={"type": "json_object"}
    )

    result = response.choices[0].message.content
    return result
```

**Input:**
```
vision_text = """H-BVI/2024
21.1234°N
105.5678°E
Thứ Hai, 13 Tháng 10,2025
11:03
dBm
-2013"""
```

**Output:**
```json
{
    "cabinet_name": "H-BVI/2024",
    "lat": "21.1234",
    "long": "105.5678",
    "date": "2025-10-13",
    "time": "11:03:00",
    "power_after_s2": "-20.13dBm"
}
```

---

### **Giai Đoạn 4: Google Sheets Storage**

**Mục đích:** Lưu dữ liệu phân tích vào Google Sheets

**File:** `gsmnv.py:311-483`

```python
def save_to_google_sheets(thread_id, name, title, image_url, vision_text, openai_analysis):
    """
    1. Parse OpenAI JSON analysis
    2. Tính toán khoảng cách (distance)
    3. Chuẩn bị row data
    4. Ghi vào Google Sheets
    5. Gửi Telegram response
    """

    # Parse OpenAI data
    parsed_data = json.loads(json_str)
    cabinet_name = parsed_data.get("cabinet_name", "")
    lat = parsed_data.get("lat", "")
    long = parsed_data.get("long", "")

    # Tính toán khoảng cách
    cabinet_result = find_nearest_bts_station(lat, long, cabinet_name)
    distance_m = cabinet_result['distance_km'] * 1000

    # Chuẩn bị row
    row_data = [name, thread_id, title, cabinet_name, ...]

    # Ghi vào Sheets
    sheet.append_row(row_data)
```

**Input:**
```
openai_analysis = JSON string
```

**Output:**
```
✅ Data saved to Google Sheets
✅ Row inserted with all fields
```

---

### **Giai Đoạn 5: Telegram Response**

**Mục đích:** Gửi kết quả tổng hợp tới Telegram group

**File:** `gsmnv.py:427-448`

```python
# Chuẩn bị dữ liệu
telegram_data = {
    "cabinet_name": cabinet_name,
    "lat": lat,
    "long": long,
    "date": extracted_date,
    "time": extracted_time,
    "power_after_s2": power_after_s2,
    "distance_m": distance_m,
    "sender_name": name
}

# Format tin nhắn
telegram_message = format_telegram_response(telegram_data)

# Gửi tới Telegram
send_telegram_message(TELEGRAM_CHANNEL_ID, telegram_message)
```

**Output:**
```
📊 KẾT QUẢ PHÂN TÍCH GPS

👤 Người gửi: Nguyễn Văn A

🗄️ Tủ BTS: H-BVI/2024

📅 Ngày: 2025-10-13

🕐 Thời gian: 11:03:00

🌍 Tọa độ:
  • Vĩ độ: 21.1234
  • Kinh độ: 105.5678

📏 Khoảng cách: 1234 mét

⚡ Công suất sau S2: -20.13dBm

✓ Dữ liệu đã lưu vào hệ thống
```

---

## 🔍 Xác Nhận Luồng Cốt Lõi

| Giai Đoạn | Input | Function | Output |
|-----------|-------|----------|--------|
| 1. GCS Upload | Telegram file_id | `upload_image_from_telegram()` | GCS public URL |
| 2. Vision OCR | GCS URL | `process_image_with_vision()` | Extracted text |
| 3. OpenAI Parse | Vision text | `analyze_text_with_openai()` | JSON data |
| 4. Sheets Save | OpenAI JSON | `save_to_google_sheets()` | Row saved |
| 5. Telegram Send | Parsed data | `send_telegram_message()` | Response sent |

---

## ✅ Kết Luận

**Luồng hiện tại HOÀN TOÀN tuân thủ cấu trúc gốc:**

✅ Upload ảnh lên GCS
✅ Gửi GCS URL tới Vision API OCR
✅ Nhận text từ Vision API
✅ Gửi text tới OpenAI API
✅ Nhận kết quả phân tích
✅ Lưu vào Google Sheets
✅ Gửi kết quả tới Telegram

**Không có bước nào bị bỏ lỡ hoặc thay đổi thứ tự.**

---

## 📞 Code References

- **telegram_bot.py:95** - Main Telegram image processing function
- **gcs_storage.py:11** - GCS upload function
- **gsmnv.py:139** - Vision API OCR function
- **gsmnv.py:163** - OpenAI analysis function
- **gsmnv.py:311** - Google Sheets save function
- **gsmnv.py:485** - Main webhook endpoint
- **gsmnv.py:541** - Telegram webhook endpoint

---

**Status: ✅ WORKFLOW ARCHITECTURE VERIFIED**

Cấu trúc luồng xử lý ứng dụng hoàn toàn chính xác và tuân thủ yêu cầu ban đầu.
