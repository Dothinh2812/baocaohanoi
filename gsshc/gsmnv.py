from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import uvicorn
import os
from google.cloud import vision
from openai import OpenAI
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json
import pandas as pd
import math
import requests

# Import BTS distance calculation function
try:
    from distance_bts import get_nearest_bts_with_distance
except ImportError:
    print("Warning: distance_bts module not found. BTS calculation will be disabled.")
    def get_nearest_bts_with_distance(lat, long, excel_file="map_gps_all_bts.xlsx"):
        return ("BTS calculation unavailable", None)

# Import Telegram bot handler
try:
    from telegram_bot import process_telegram_image
except ImportError:
    print("Warning: telegram_bot module not found. Telegram webhook will be disabled.")
    def process_telegram_image(update):
        return {"error": "Telegram bot module not available"}

app = FastAPI(title="Webhook API", description="API để nhận thông tin từ n8n workflow")

def _read_secret(env_var: str, fallback_file: str = "") -> str:
    value = os.environ.get(env_var, "").strip()
    if value:
        return value
    if fallback_file and os.path.exists(fallback_file):
        with open(fallback_file, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""

# Cấu hình Telegram Bot
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHANNEL_ID = os.environ.get("TELEGRAM_CHANNEL_ID", "")

# Cấu hình Webhook URL để gửi thông báo
WEBHOOK_URL = os.environ.get("MAIN_WEBHOOK_URL", "")

# Cấu hình Google Cloud Vision
google_credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "vision-key.json")
if google_credentials_path:
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = google_credentials_path
try:
    vision_client = vision.ImageAnnotatorClient()
except Exception:
    vision_client = None

# Cấu hình OpenAI
openai_api_key = _read_secret("OPENAI_API_KEY", "openai-vison-key.txt")
client = OpenAI(api_key=openai_api_key) if openai_api_key else None

GOOGLE_SHEETS_CREDENTIALS_FILE = os.environ.get("GOOGLE_SHEETS_CREDENTIALS_FILE", "ggsheet-key.json")

# Cấu hình Google Sheets
def init_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_SHEETS_CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client

# Model cho dữ liệu từ n8n
class WebhookData(BaseModel):
    threadId: str
    name: str
    title: str
    image_url: str

def send_telegram_message(chat_id: str, message: str, parse_mode: str = "HTML") -> bool:
    """
    Gửi tin nhắn tới Telegram chat

    Args:
        chat_id: Chat ID hoặc Group ID
        message: Nội dung tin nhắn
        parse_mode: HTML, Markdown, hoặc MarkdownV2

    Returns:
        True nếu gửi thành công
    """
    try:
        if not TELEGRAM_BOT_TOKEN:
            print("❌ Thiếu TELEGRAM_BOT_TOKEN")
            return False
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            "chat_id": chat_id,
            "text": message,
            "parse_mode": parse_mode
        }

        response = requests.post(url, json=payload, timeout=10)
        result = response.json()

        if result.get("ok"):
            print(f"✅ Gửi tin nhắn Telegram thành công (chat_id: {chat_id})")
            return True
        else:
            print(f"❌ Lỗi gửi tin nhắn Telegram: {result.get('description', 'Unknown error')}")
            return False

    except Exception as e:
        print(f"❌ Lỗi gửi tin nhắn Telegram: {str(e)}")
        return False

def format_telegram_response(data: dict) -> str:
    """
    Format kết quả phân tích thành tin nhắn Telegram đẹp

    Args:
        data: Dictionary chứa kết quả phân tích

    Returns:
        String tin nhắn HTML format
    """
    cabinet_name = data.get("cabinet_name", "N/A")
    lat = data.get("lat", "N/A")
    long = data.get("long", "N/A")
    date = data.get("date", "N/A")
    time = data.get("time", "N/A")
    power_after_s2 = data.get("power_after_s2", "N/A")
    distance_m = data.get("distance_m", "N/A")
    sender_name = data.get("sender_name", "Unknown")

    message = f"""
<b>📊 KẾT QUẢ PHÂN TÍCH GPS</b>
<b>👤 Người gửi:</b> {sender_name}
<b>🗄️ Tủ:</b> <code>{cabinet_name}</code>
<b>📅 Ngày:</b> {date}
<b>🕐 Thời gian:</b> {time}
<b>🌍 Tọa độ:</b>
  • <b>Vĩ độ:</b> {lat}
  • <b>Kinh độ:</b> {long}
<b>📏 Khoảng cách:</b> {distance_m} mét
<b>⚡ Công suất sau S2:</b> {power_after_s2}
<i>✓ Dữ liệu đã lưu vào hệ thống</i>
"""
    return message

def process_image_with_vision(image_url):
    """Xử lý ảnh với Google Cloud Vision API"""
    try:
        if vision_client is None:
            raise RuntimeError("Google Cloud Vision chưa được cấu hình")
        # Tạo đối tượng ảnh từ URL
        image = vision.Image()
        image.source.image_uri = image_url
        
        # Gửi yêu cầu OCR
        response = vision_client.text_detection(image=image)
        
        # Xử lý kết quả
        texts = response.text_annotations
        if texts:
            detected_text = texts[0].description
            print(f"Vision API - Text detected: {detected_text}")
            return detected_text
        else:
            print("Vision API - No text detected")
            return ""
            
    except Exception as e:
        print(f"Error with Vision API: {str(e)}")
        return f"Error: {str(e)}"

def analyze_text_with_openai(text, name):
    """Phân tích text với OpenAI Assistant"""
    try:
        if client is None:
            raise RuntimeError("OpenAI API key chưa được cấu hình")
        # System prompt được nhúng trực tiếp trong code
        system_prompt = """You are a text analysis assistant.
You will receive a text content extracted from a photo via the OCR tool.
Your task is to Analyze information from the text, then return the result in JSON format:
{
  "cabinet_name": "box cabinet name",
  "lat": "coordinates, latitude",
  "long": "long, longtitude",
  "date": "date displayed in text",
  "time": "time displayed in text",
  "power_after_s2": "power after S2"
}
Field Descriptions:
1. cabinet_name - The box cabinet name will be in the form H-ABC/xxxx, or O-ABC/xxxx
   where ABC being 3 alphabetic characters, xxxx being 4 or 5 digits
   Examples: H-STY/7846, H-BVI/5609, H-STY/78463
   Special case: May appear as "O-ABC/xxxx" - normalize to "H-ABC/xxxx" format
   Special case: May appear as just "ABC/xxxx" or "ABC/xxxxx" - add "H-" prefix automatically
   IMPORTANT: Always normalize cabinet names to "H-" prefix regardless of source format
2. "lat": "coordinate, latitude" - is a coordinate component, latitude
   May appear with degree symbol (°) like: 21.156801°N
3. "long": "long, longtitude" - is a coordinate component, longtitude
   May appear with degree symbol (°) like: 105.477859°E
4. date - date displayed in text
   Note: Format date into a unified standard format YYYY-MM-DD
   Examples: "Thứ Hai, 13 Tháng 10,2025" -> "2025-10-13"
5. time - time displayed in text
   Note: Format time into a unified standard format HH:MM:SS
   Examples: "11:03" -> "11:03:00"
6. power_after_s2 - Power after S2 (Công suất sau S2)
   IMPORTANT: This value may appear in different formats in OCR text:
   Format 1: Standard format like "-20.13dBm" or "-17.62 dBm"
   Format 2: Separated on multiple lines like:
      dBm
      -1700
   When you see a negative number near "dBm" text, extract and format it:
   - If value is "-1700" or similar 4-digit number, convert to decimal: "-17.00"
   - If value is "-2013", convert to: "-20.13"
   - If value is already with decimal like "-17.62", keep as is
   - Always add "dBm" unit at the end
   - Final format examples: "-17.00dBm", "-20.13dBm", "-15.45dBm"
   If no power value found in the text, return empty string ""
Important Notes:
- All fields should be extracted from the OCR text
- Dates and times should be standardized to consistent formats
- Cabinet names: ALWAYS normalize to "H-" prefix (convert "O-" to "H-", add "H-" if missing)
- Cabinet names: Support both 4-digit and 5-digit formats after the slash (e.g., H-ABC/7846 or H-ABC/78463)
- Power values: Look for "dBm" keyword and nearby negative numbers (may be on different lines)
- Power values: Convert 4-digit numbers to decimal format (divide by 100)
- Return empty string for any field that cannot be found in the text
- Always return valid JSON format only, no additional text"""

        # Sử dụng Chat Completions API thay vì Responses API
        response = client.chat.completions.create(
            model="gpt-4o-mini",  # Hoặc "gpt-4o" nếu cần chất lượng cao hơn
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text}
            ],
            temperature=0.1,
            response_format={"type": "json_object"}
        )

        # Lấy kết quả
        result = response.choices[0].message.content
        print("OpenAI Chat API output:", result)

        # Kiểm tra có đủ trường JSON không
        required_fields = ["cabinet_name", "lat", "long", "date", "time", "power_after_s2"]
        try:
            start_idx = result.find('{')
            end_idx = result.rfind('}') + 1
            if start_idx != -1 and end_idx > start_idx:
                json_str = result[start_idx:end_idx]
                parsed = json.loads(json_str)
                missing = [field for field in required_fields if field not in parsed]
                if missing:
                    print(f"⚠️ Thiếu trường trong phản hồi: {missing}")
                else:
                    print("✅ Đầy đủ trường JSON yêu cầu.")
        except Exception as e:
            print(f"⚠️ Không thể parse JSON từ phản hồi: {e}")

        return result
    except Exception as e:
        print(f"Lỗi khi phân tích với Responses API: {str(e)}")
        return f"Lỗi khi phân tích: {str(e)}"

def send_analysis_notification(lat, long, date, time, nearest_bts, distance_m, name, thread_id, cabinet_name, power_after_s2):
    """Gửi bản tin kết quả phân tích tới URL endpoint"""
    try:
        print("Step 5: Sending analysis notification...")

        # cabinet_name sẽ được truyền vào hàm này
        # ...existing code...
        message = (
            f"📍 Kết quả phân tích GPS\n"
            f"👤 Người gửi: {name}\n"
            f"🗄️ Tủ: {cabinet_name}\n"
            f"📅 Ngày: {date}\n"
            f"🕐 Thời gian: {time}\n"
            f"🌍 Tọa độ: {lat}, {long}\n"
            # f"📡 Trạm BTS gần nhất: {nearest_bts}\n"
            f"📏 Khoảng cách so với capman: {distance_m} mét\n"
            f"⚡ Công suất sau S2: {power_after_s2}"
        )

        payload = {
            'threadID': thread_id,
            'message': message
        }

        print(f"Sending notification to: {WEBHOOK_URL}")
        print(f"Payload: {payload}")

        response = requests.get(WEBHOOK_URL, json=payload, timeout=10)

        if response.status_code == 200:
            print("✅ Analysis notification sent successfully")
            return True
        else:
            print(f"❌ Failed to send notification. Status code: {response.status_code}")
            print(f"Response: {response.text}")
            return False

    except requests.exceptions.Timeout:
        print("❌ Timeout while sending notification")
        return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Error sending notification: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error sending notification: {str(e)}")
        return False

def save_to_google_sheets(thread_id, name, title, image_url, vision_text, openai_analysis):
    """Lưu kết quả vào Google Sheets"""
    try:
        print("Initializing Google Sheets client...")
        client = init_google_sheets()
        
        print("Opening Google Sheets document...")
        # Mở sheet (cần thay đổi tên sheet phù hợp)
        sheet = client.open_by_key("10mFy9EzRNG2VvOOWnZe0Tl8OK2QnN3cJXTJiMaeUK9Q").worksheet("thang11")
        print(f"Successfully opened worksheet: {sheet.title}")
        
        # Tạo timestamp
        now = datetime.now()
        created_date = now.strftime("%Y-%m-%d")
        created_time = now.strftime("%H:%M:%S")
        
        # Parse JSON từ OpenAI analysis
        cabinet_name = ""
        lat = ""
        long = ""
        extracted_date = ""
        extracted_time = ""
        power_after_s2 = ""
        nearest_bts = ""
        distance_m = ""

        try:
            # Convert openai_analysis to string if it's not already
            analysis_str = str(openai_analysis) if openai_analysis else ""

            # Thử parse JSON từ OpenAI response
            if analysis_str and not analysis_str.startswith("Error"):
                # Tìm JSON trong response (có thể có text khác xung quanh)
                start_idx = analysis_str.find('{')
                end_idx = analysis_str.rfind('}') + 1

                if start_idx != -1 and end_idx > start_idx:
                    json_str = analysis_str[start_idx:end_idx]
                    parsed_data = json.loads(json_str)

                    cabinet_name = parsed_data.get("cabinet_name", "")
                    lat = parsed_data.get("lat", "")
                    long = parsed_data.get("long", "")
                    extracted_date = parsed_data.get("date", "")
                    extracted_time = parsed_data.get("time", "")
                    power_after_s2 = parsed_data.get("power_after_s2", "")

                    print(f"Parsed OpenAI data - Cabinet: {cabinet_name}, Lat: {lat}, Long: {long}, Date: {extracted_date}, Time: {extracted_time}, Power: {power_after_s2}")

                    # Tính toán khoảng cách từ tọa độ nhận được đến hộp cabinet
                    if lat and long and cabinet_name:
                        try:
                            print("Step 4: Calculating distance to cabinet...")

                            # Xử lý tọa độ để loại bỏ ký tự không phải số (như °)
                            def clean_coordinate(coord_str):
                                import re
                                clean_str = re.sub(r'[^0-9.\-]', '', str(coord_str))
                                return float(clean_str)

                            clean_lat = clean_coordinate(lat)
                            clean_long = clean_coordinate(long)

                            print(f"Cleaned coordinates: Lat={clean_lat}, Long={clean_long}")
                            from distance_bts import find_nearest_bts_station
                            cabinet_result = find_nearest_bts_station(clean_lat, clean_long, cabinet_name, excel_file_path="ket_qua_gop.xlsx")
                            if 'error' in cabinet_result:
                                nearest_bts = cabinet_result['error']
                                distance_m = None
                            else:
                                nearest_bts = cabinet_result['cabinet']['cabinet_name']
                                # Chuyển đổi từ km sang mét
                                distance_m = round(cabinet_result['distance_km'] * 1000)
                            print(f"Cabinet: {nearest_bts}, Distance: {distance_m} m")
                        except Exception as bts_error:
                            print(f"Error calculating cabinet distance: {str(bts_error)}")
                            nearest_bts = "Cabinet calculation failed"
                            distance_m = "N/A"

        except (json.JSONDecodeError, KeyError, AttributeError) as e:
            print(f"Error parsing OpenAI JSON: {str(e)}")
            # Nếu không parse được JSON, để trống các trường

        # Chuẩn bị dữ liệu theo format mới
        # name, thread_id, title, cabinet_name, Người gửi, Tủ, Ngày, Thời gian, Tọa độ, Khoảng cách, Công suất sau S2
        # Người gửi: name
        # Tủ: cabinet_name
        # Ngày: extracted_date
        # Thời gian: extracted_time
        # Tọa độ: lat, long
        # Khoảng cách: distance_m (mét)
        # Công suất sau S2: power_after_s2 (dBm)
        row_data = [
            name,           # name
            thread_id,      # threadId
            title,          # title
            cabinet_name,   # cabinet_name
            extracted_date,
            extracted_time,
            lat,
            long,
            distance_m,
            power_after_s2, # power_after_s2
            image_url,       # image_url
            created_date,  # created_date
            created_time,  # created_time
        ]
        
        print(f"Preparing to save row data: {row_data}")
        
        # Ghi vào sheet
        result = sheet.append_row(row_data)
        print(f"Google Sheets append result: {result}")
        print("Data saved to Google Sheets successfully")
        
        # Gửi thông báo kết quả phân tích chỉ khi có đủ dữ liệu bắt buộc
        if lat and long and extracted_date and extracted_time:
            # Chuẩn bị dữ liệu để gửi Telegram
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

            # Gửi Telegram message
            print("\nBước 5: Gửi kết quả phân tích tới Telegram...")
            telegram_message = format_telegram_response(telegram_data)
            telegram_success = send_telegram_message(TELEGRAM_CHANNEL_ID, telegram_message)

            if telegram_success:
                print("✅ Telegram message sent successfully")
            else:
                print("⚠️ Warning: Failed to send Telegram message")

            # Gửi thông báo N8N (webhook cũ)
            if nearest_bts:
                notification_success = send_analysis_notification(
                    lat, long, extracted_date, extracted_time,
                    nearest_bts, distance_m, name, thread_id, cabinet_name, power_after_s2
                )
                if not notification_success:
                    print("⚠️ Warning: Failed to send N8N notification")
            else:
                print("⚠️ Warning: Missing BTS data - skipping N8N notification")
        else:
            missing_data = []
            if not lat: missing_data.append("latitude")
            if not long: missing_data.append("longitude")
            if not extracted_date: missing_data.append("date")
            if not extracted_time: missing_data.append("time")
            print(f"⚠️ Warning: Missing required data ({', '.join(missing_data)}) - skipping notifications")
        
        # Verify by getting the last few rows
        try:
            all_records = sheet.get_all_records()
            print(f"Total rows in sheet: {len(all_records)}")
            if all_records:
                print(f"Last record: {all_records[-1]}")
        except Exception as verify_error:
            print(f"Error verifying data: {verify_error}")
        
        return True
        
    except Exception as e:
        print(f"Error saving to Google Sheets: {str(e)}")
        import traceback
        print(f"Full traceback: {traceback.format_exc()}")
        return False

@app.post("/webhook")
async def receive_webhook(data: WebhookData):
    """
    Endpoint để nhận dữ liệu từ n8n workflow và xử lý đầy đủ
    """
    try:
        print("=" * 50)
        print("WEBHOOK DATA RECEIVED:")
        print("=" * 50)
        print(f"Thread ID: {data.threadId}")
        print(f"Name: {data.name}")
        print(f"Title: {data.title}")
        print(f"Image URL: {data.image_url}")
        print("=" * 50)
        
        # Bước 1: Xử lý ảnh với Google Cloud Vision
        print("Step 1: Processing image with Google Cloud Vision...")
        vision_text = process_image_with_vision(data.image_url)
        
        # Bước 2: Phân tích text với OpenAI
        print("Step 2: Analyzing text with OpenAI...")
        openai_analysis = analyze_text_with_openai(vision_text, data.name)
        
        # Bước 3: Lưu vào Google Sheets
        print("Step 3: Saving to Google Sheets...")
        sheets_success = save_to_google_sheets(
            data.threadId, 
            data.name, 
            data.title,
            data.image_url, 
            vision_text, 
            openai_analysis
        )
        
        print("=" * 50)
        print("PROCESSING COMPLETED")
        print("=" * 50)
        
        return {
            "status": "success",
            "message": "Webhook processed successfully",
            "data": {
                "threadId": data.threadId,
                "name": data.name,
                "title": data.title,
                "image_url": data.image_url,
                "vision_text": vision_text[:200] + "..." if len(vision_text) > 200 else vision_text,
                "openai_analysis": openai_analysis[:200] + "..." if len(openai_analysis) > 200 else openai_analysis,
                "saved_to_sheets": sheets_success
            }
        }
        
    except Exception as e:
        print(f"Error processing webhook: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error processing webhook: {str(e)}")

@app.post("/telegram-webhook")
async def telegram_webhook(request: dict):
    """
    Endpoint để nhận webhook từ Telegram
    Xử lý ảnh từ group Telegram, upload lên GCS, rồi gửi tới webhook xử lý chính
    """
    try:
        print("=" * 60)
        print("TELEGRAM WEBHOOK RECEIVED")
        print("=" * 60)
        print(f"Update data: {json.dumps(request, indent=2, ensure_ascii=False)}")

        # Xử lý ảnh từ Telegram
        result = process_telegram_image(request)

        print("=" * 60)
        print("TELEGRAM PROCESSING COMPLETED")
        print("=" * 60)

        return {
            "status": "success",
            "message": "Telegram image processed successfully",
            "result": result
        }

    except Exception as e:
        print(f"❌ Lỗi xử lý Telegram webhook: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"Error processing Telegram webhook: {str(e)}")


@app.get("/")
async def root():
    """
    Health check endpoint
    """
    return {"message": "Webhook API is running"}

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 8007))  # Default to 8007 for Cloudflare tunnel
    uvicorn.run(app, host="0.0.0.0", port=port)
