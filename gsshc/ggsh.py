import gspread
from oauth2client.service_account import ServiceAccountCredentials
#pip install gspread oauth2client

# B1: Xác thực qua file credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("n8n-doxuanthinh-com-ed80b978b1b3.json", scope)
client = gspread.authorize(creds)

# B2: Mở Google Sheet bằng tên hoặc URL
sheet = client.open("CSKH-sms-agent").sheet1  # hoặc .worksheet("Tên sheet")

# B3: Ghi dữ liệu
# Ghi 1 dòng vào cuối sheet
sheet.append_row(["Tên", "Tuổi", "Thành phố"])

# Đọc dữ liệu từ sheet1 và in ra 5 dòng đầu tiên
rows = sheet.get_all_values()
for row in rows[:5]:
    print(row)