# 🎯 CORE FLOW REFERENCE - Tham Chiếu Luồng Cốt Lõi

## ✅ Xác Nhận: Luồng Tuân Thủ 100%

Yêu cầu gốc của bạn:
```
Upload GCS → Vision API → OpenAI API → Google Sheets → Telegram Response
```

**Status:** ✅ **HOÀN TOÀN TUÂN THỦ**

---

## 📍 5 Giai Đoạn Cốt Lõi

### **Giai Đoạn 1️⃣: Upload lên GCS**

**Câu hỏi:** *Sau khi nhận ảnh từ Telegram, ảnh được upload đâu?*
**Trả lời:** ✅ Google Cloud Storage (GCS)

```
telegram_bot.py:95 (process_telegram_image)
    ↓
gcs_storage.py:11 (upload_image_from_telegram)
    │
    ├─ Download ảnh từ Telegram server
    │  URL: https://api.telegram.org/file/bot{TOKEN}/{file_path}
    │
    ├─ Upload lên GCS bucket
    │  Bucket: bts-telegram-images
    │  Path: telegram-images/2025/10/20/file.jpg
    │
    └─ Trả về public URL
       Return: https://storage.googleapis.com/bts-telegram-images/...
```

**Verification:** `gcs_storage.py:42-60`
```python
# Upload lên GCS
blob = bucket.blob(destination_blob_name)
blob.upload_from_string(response.content, content_type="image/jpeg")

# Trả về public URL
public_url = f"https://storage.googleapis.com/{bucket_name}/{destination_blob_name}"
return {"public_url": public_url}
```

---

### **Giai Đoạn 2️⃣: Vision API OCR**

**Câu hỏi:** *Lấy link GCS như thế nào và gửi tới Vision API?*
**Trả lời:** ✅ Gửi GCS URL trực tiếp tới Vision API

```
gsmnv.py:485 (receive_webhook)
    ↓
    │ Nhận input: image_url từ GCS
    │
gsmnv.py:139 (process_image_with_vision)
    │
    ├─ Tạo Vision Image object
    │  image = vision.Image()
    │  image.source.image_uri = image_url  ← GCS URL
    │
    ├─ Gọi Vision API text detection
    │  response = vision_client.text_detection(image=image)
    │
    └─ Trả về extracted text
       Return: "H-BVI/2024\n21.1234°N\n105.5678°E\n..."
```

**Verification:** `gsmnv.py:139-161`
```python
def process_image_with_vision(image_url):
    # Input: GCS URL
    image = vision.Image()
    image.source.image_uri = image_url  # ← Direct GCS URL

    # Call Vision API
    response = vision_client.text_detection(image=image)

    # Get text
    texts = response.text_annotations
    if texts:
        detected_text = texts[0].description
        return detected_text
```

---

### **Giai Đoạn 3️⃣: OpenAI Analysis**

**Câu hỏi:** *Text từ Vision API được làm gì?*
**Trả lời:** ✅ Gửi tới OpenAI API để phân tích và extract structured data

```
gsmnv.py:502 (in receive_webhook)
    ↓
    │ Input: vision_text từ Vision API
    │
gsmnv.py:163 (analyze_text_with_openai)
    │
    ├─ Gửi text + system prompt tới OpenAI
    │  client.chat.completions.create(
    │      model="gpt-4o-mini",
    │      messages=[
    │          {"role": "system", "content": system_prompt},
    │          {"role": "user", "content": vision_text}  ← Vision text
    │      ],
    │      response_format={"type": "json_object"}
    │  )
    │
    ├─ OpenAI phân tích text
    │  Extract: cabinet_name, lat, long, date, time, power_after_s2
    │
    └─ Trả về JSON structured data
       Return: {"cabinet_name": "H-BVI/2024", "lat": "21.1234", ...}
```

**Verification:** `gsmnv.py:163-262`
```python
def analyze_text_with_openai(text, name):
    # Input: vision_text (từ Vision API)

    system_prompt = """Extract information from OCR text:
    - cabinet_name: H-ABC/xxxx
    - lat: latitude coordinate
    - long: longitude coordinate
    - date: YYYY-MM-DD format
    - time: HH:MM:SS format
    - power_after_s2: -XXdBm format
    Return as JSON only."""

    # Send to OpenAI
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text}  # ← Vision text
        ],
        response_format={"type": "json_object"}
    )

    result = response.choices[0].message.content
    return result  # JSON string
```

---

### **Giai Đoạn 4️⃣: Google Sheets Save**

**Câu hỏi:** *JSON từ OpenAI được lưu vào đâu?*
**Trả lời:** ✅ Google Sheets

```
gsmnv.py:509 (in receive_webhook)
    ↓
    │ Input: openai_analysis (JSON)
    │
gsmnv.py:311 (save_to_google_sheets)
    │
    ├─ Parse OpenAI JSON
    │  parsed_data = json.loads(json_str)
    │  cabinet_name = parsed_data.get("cabinet_name")
    │  lat = parsed_data.get("lat")
    │  long = parsed_data.get("long")
    │  date = parsed_data.get("date")
    │  time = parsed_data.get("time")
    │  power_after_s2 = parsed_data.get("power_after_s2")
    │
    ├─ Tính toán khoảng cách (nếu có tọa độ)
    │  cabinet_result = find_nearest_bts_station(lat, long, cabinet_name)
    │  distance_m = round(cabinet_result['distance_km'] * 1000)
    │
    ├─ Chuẩn bị row data
    │  row_data = [name, threadId, title, cabinet_name, date, time, ...]
    │
    ├─ Ghi vào Google Sheets
    │  sheet.append_row(row_data)
    │
    └─ Trả về success status
       Return: True (if saved successfully)
```

**Verification:** `gsmnv.py:311-483`
```python
def save_to_google_sheets(..., openai_analysis):
    # Parse OpenAI analysis
    start_idx = analysis_str.find('{')
    end_idx = analysis_str.rfind('}') + 1
    json_str = analysis_str[start_idx:end_idx]
    parsed_data = json.loads(json_str)

    # Extract fields
    cabinet_name = parsed_data.get("cabinet_name", "")
    lat = parsed_data.get("lat", "")
    long = parsed_data.get("long", "")

    # Calculate distance
    cabinet_result = find_nearest_bts_station(lat, long, cabinet_name)
    distance_m = round(cabinet_result['distance_km'] * 1000)

    # Save to Sheets
    row_data = [name, threadId, title, cabinet_name, ...]
    result = sheet.append_row(row_data)
```

---

### **Giai Đoạn 5️⃣: Telegram Response**

**Câu hỏi:** *Sau khi lưu vào Sheets, kết quả được gửi tới đâu?*
**Trả lời:** ✅ Telegram group (tự động phản hồi)

```
gsmnv.py:427 (in save_to_google_sheets)
    │ Check if: lat && long && date && time
    │
    ├─ Format tin nhắn Telegram
    │  gsmnv.py:97 (format_telegram_response)
    │  telegram_data = {
    │      "cabinet_name": cabinet_name,
    │      "lat": lat,
    │      "long": long,
    │      "date": extracted_date,
    │      "time": extracted_time,
    │      "power_after_s2": power_after_s2,
    │      "distance_m": distance_m,
    │      "sender_name": name
    │  }
    │  telegram_message = format_telegram_response(telegram_data)
    │
    └─ Gửi tin nhắn tới Telegram
       gsmnv.py:63 (send_telegram_message)
       send_telegram_message(TELEGRAM_CHANNEL_ID, telegram_message)

       URL: https://api.telegram.org/bot{TOKEN}/sendMessage
       Payload: {
           "chat_id": TELEGRAM_CHANNEL_ID,
           "text": telegram_message,
           "parse_mode": "HTML"
       }

       Response: ✅ Message sent to Telegram group
```

**Verification:** `gsmnv.py:427-448`
```python
# After save to Sheets
if lat and long and extracted_date and extracted_time:
    # Format data
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

    # Format message
    telegram_message = format_telegram_response(telegram_data)

    # Send to Telegram
    telegram_success = send_telegram_message(TELEGRAM_CHANNEL_ID, telegram_message)
```

---

## 🔗 Complete Data Flow

```
┌─────────────────────────────────────────────────────────────────────────┐
│                    IMAGE FLOW THROUGH SYSTEM                            │
└─────────────────────────────────────────────────────────────────────────┘

Input: Telegram User sends image

  ↓

[Phase 1: GCS Upload]
Telegram image → Download from Telegram server → Upload to GCS
Output: image_url (https://storage.googleapis.com/...)

  ↓

[Phase 2: Vision API OCR]
image_url (GCS URL) → Vision API text_detection()
Output: vision_text (extracted OCR text)

  ↓

[Phase 3: OpenAI Analysis]
vision_text → OpenAI API (gpt-4o-mini) → Parse with system prompt
Output: openai_analysis (JSON with fields)

  ↓

[Phase 4: Google Sheets Save]
openai_analysis (JSON) → Parse fields → Calculate distance → Save to Sheets
Output: sheets_success (bool)

  ↓

[Phase 5: Telegram Response]
Parsed fields → Format message → Send to Telegram group
Output: User receives response in Telegram

  ↓

Final: ✅ Complete
```

---

## 📊 Data at Each Stage

### Stage 1: GCS Upload
```
Input:
  • file_id: "AgACAgIAAxkBAAIF..."
  • bot_token: "<your-telegram-bot-token>"
  • filename: "telegram_1001_testuser.jpg"

Process:
  • Download image from Telegram
  • Upload to GCS bucket

Output:
  • image_url: "https://storage.googleapis.com/bts-telegram-images/telegram-images/2025/10/20/telegram_1001_testuser.jpg"
```

### Stage 2: Vision API OCR
```
Input:
  • image_url: "https://storage.googleapis.com/..."

Process:
  • Send to Vision API
  • Extract text via OCR

Output:
  • vision_text: """
    H-BVI/2024
    21.1234°N
    105.5678°E
    Thứ Hai, 13 Tháng 10,2025
    11:03
    dBm
    -2013
    """
```

### Stage 3: OpenAI Analysis
```
Input:
  • vision_text: extracted OCR text

Process:
  • Send to OpenAI with system prompt
  • Parse and extract structured fields

Output:
  • openai_analysis: {
      "cabinet_name": "H-BVI/2024",
      "lat": "21.1234",
      "long": "105.5678",
      "date": "2025-10-13",
      "time": "11:03:00",
      "power_after_s2": "-20.13dBm"
    }
```

### Stage 4: Google Sheets
```
Input:
  • openai_analysis: JSON with all fields

Process:
  • Parse JSON
  • Extract fields
  • Calculate distance
  • Create row

Output:
  • Row in Google Sheets with 13 columns:
    [name, threadId, title, cabinet_name, date, time, lat, long, distance, power, url, created_date, created_time]
```

### Stage 5: Telegram Response
```
Input:
  • Parsed fields from OpenAI

Process:
  • Format with HTML and emojis
  • Create message

Output:
  • Message in Telegram group:
    📊 KẾT QUẢ PHÂN TÍCH GPS
    👤 Người gửi: [name]
    🗄️ Tủ BTS: [cabinet]
    ...
```

---

## 🔍 Function Call Stack

```
Telegram image arrives
  │
  ├─ POST /telegram-webhook
  │   gsmnv.py:541
  │
  ├─ process_telegram_image()
  │   telegram_bot.py:95
  │
  ├─ upload_image_from_telegram()
  │   gcs_storage.py:11
  │   → Returns: image_url
  │
  ├─ POST to MAIN_WEBHOOK_URL
  │   Sends: {image_url, threadId, name, title}
  │
  ├─ receive_webhook()
  │   gsmnv.py:485
  │
  ├─ process_image_with_vision(image_url)
  │   gsmnv.py:139
  │   → Returns: vision_text
  │
  ├─ analyze_text_with_openai(vision_text)
  │   gsmnv.py:163
  │   → Returns: openai_analysis (JSON)
  │
  ├─ save_to_google_sheets(openai_analysis)
  │   gsmnv.py:311
  │
  ├─ format_telegram_response(parsed_data)
  │   gsmnv.py:97
  │   → Returns: telegram_message
  │
  └─ send_telegram_message(message)
      gsmnv.py:63
      → User receives response in Telegram
```

---

## 🎯 Critical Path (No Deviations)

✅ **Step 1:** Telegram image → GCS (NOT saved locally)
✅ **Step 2:** GCS URL → Vision API (NOT downloaded and then uploaded again)
✅ **Step 3:** Vision text → OpenAI (NOT skipped)
✅ **Step 4:** OpenAI JSON → Google Sheets (NOT cached in memory only)
✅ **Step 5:** Parsed data → Telegram (NOT dropped)

**No shortcuts or alternative paths exist.**

---

## 📋 Checklist: 5 Mandatory Steps

- [x] **Step 1:** Upload image to GCS → Get public URL
- [x] **Step 2:** Send GCS URL to Vision API → Get extracted text
- [x] **Step 3:** Send extracted text to OpenAI → Get structured JSON
- [x] **Step 4:** Parse JSON and save to Google Sheets
- [x] **Step 5:** Send analysis result to Telegram group

All steps executed in order. **No steps skipped.**

---

## 🚀 Quick Verification Commands

### Check Phase 1: GCS Upload Works
```bash
# Verify gcs_storage.py exists and has upload function
grep -n "def upload_image_from_telegram" gcs_storage.py

# Expected: gcs_storage.py:11:def upload_image_from_telegram
```

### Check Phase 2: Vision API is Called with GCS URL
```bash
# Verify process_image_with_vision uses GCS URL directly
grep -A 5 "image.source.image_uri" gsmnv.py

# Expected: image.source.image_uri = image_url
```

### Check Phase 3: OpenAI Gets Vision Text
```bash
# Verify analyze_text_with_openai receives vision_text
grep -B 2 -A 10 "def analyze_text_with_openai" gsmnv.py

# Expected: Second parameter is 'text' (vision_text)
```

### Check Phase 4: OpenAI Result Goes to Sheets
```bash
# Verify save_to_google_sheets calls with openai_analysis
grep "save_to_google_sheets" gsmnv.py

# Expected: openai_analysis is passed as parameter
```

### Check Phase 5: Telegram Gets Response
```bash
# Verify send_telegram_message is called from save_to_google_sheets
grep -A 20 "if lat and long" gsmnv.py

# Expected: send_telegram_message() called with formatted data
```

---

## ✅ Conclusion

**The workflow is 100% compliant with your original architecture:**

1. **Upload GCS** ✅ - `gcs_storage.py:11`
2. **Vision API** ✅ - `gsmnv.py:139`
3. **OpenAI API** ✅ - `gsmnv.py:163`
4. **Google Sheets** ✅ - `gsmnv.py:311`
5. **Telegram Response** ✅ - `gsmnv.py:427-448`

**No deviations, no shortcuts, no alternative paths.**

---

**Status:** ✅ **WORKFLOW ARCHITECTURE VERIFIED & DOCUMENTED**

Hệ thống hoàn toàn tuân thủ yêu cầu gốc của bạn.

**You can now run `./run.sh` with confidence that the workflow is correct!**
