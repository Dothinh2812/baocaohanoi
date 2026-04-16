# Hướng Dẫn Dùng `neo` Trên Windows PowerShell

Tài liệu này gom lại toàn bộ kinh nghiệm thực tế khi dùng `neo` trên Windows, đặc biệt là các lỗi đã gặp với PowerShell, Chrome debug port, extension, `neo connect`, và capture request từ `baocao.hanoi.vnpt.vn`.

Phạm vi:
- cài `neo`
- mở Chrome đúng cách trên PowerShell
- load extension
- sửa lỗi Windows thường gặp
- bắt capture
- cách đi tiếp nếu `neo` trên Windows chỉ xem được summary nhưng không xem được detail/export

## 1. Bối cảnh quan trọng

Trên Windows PowerShell:
- không dùng cú pháp shell Linux kiểu xuống dòng bằng `\`
- không gõ riêng từng option như `--remote-debugging-port=9222`
- phải gọi hẳn executable của Chrome/Edge trong một dòng

Ví dụ sai:

```powershell
chromium \
  --remote-debugging-port=9222 \
  --user-data-dir=/tmp/neo-chrome-profile
```

PowerShell sẽ báo lỗi kiểu:
- `Missing expression after unary operator '--'`
- `chromium : The term 'chromium' is not recognized`

## 2. Chuẩn bị

Cần có:
- `git`
- `node` + `npm`
- `Google Chrome` hoặc `Microsoft Edge`
- một thư mục riêng cho repo `neo`, ví dụ `D:\neo\neo`
- một thư mục temp riêng, ví dụ `D:\tmp`

## 3. Cài `neo`

Ví dụ:

```powershell
cd D:\neo
git clone https://github.com/4ier/neo.git
cd D:\neo\neo
npm install
npm run build
npm link
```

Kiểm tra:

```powershell
neo
```

Nếu cài đúng, CLI sẽ hiện help.

## 4. Mở Chrome đúng cách trên Windows PowerShell

### 4.1. Với Google Chrome

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="D:\neo\chrome-profile"
```

Nếu Chrome nằm ở `Program Files (x86)`:

```powershell
& "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="D:\neo\chrome-profile"
```

### 4.2. Với Microsoft Edge

```powershell
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="D:\neo\edge-profile"
```

## 5. Load extension

Trong Chrome vừa mở:

1. Mở `chrome://extensions`
2. Bật `Developer mode`
3. Chọn `Load unpacked`
4. Trỏ tới thư mục:

```text
D:\neo\neo\extension-dist
```

Lưu ý:
- với repo `neo` đã build, `extension-dist/` là thư mục cần load
- nếu load sai thư mục, `neo doctor` thường không nhìn thấy extension

## 6. Sửa lỗi môi trường Windows cho `neo`

`neo` trên Windows thường giả định môi trường kiểu Linux/macOS. Vì vậy cần set biến môi trường trong cùng cửa sổ PowerShell trước khi dùng.

Chạy:

```powershell
$env:HOME = $env:USERPROFILE
$env:TEMP = "D:\tmp"
$env:TMP = "D:\tmp"
New-Item -ItemType Directory -Force -Path "$env:HOME\.neo" | Out-Null
New-Item -ItemType Directory -Force -Path "$env:HOME\.neo\schemas" | Out-Null
New-Item -ItemType Directory -Force -Path "D:\tmp" | Out-Null
New-Item -ItemType File -Force -Path "D:\tmp\neo-sessions.json" | Out-Null
Set-Content -Path "D:\tmp\neo-sessions.json" -Value "{}"
```

### 6.1. Lỗi `path argument must be of type string`

Lỗi kiểu:

```text
TypeError [ERR_INVALID_ARG_TYPE]: The "path" argument must be of type string. Received undefined
```

Nguyên nhân:
- `process.env.HOME` không có trên PowerShell

Cách sửa:

```powershell
$env:HOME = $env:USERPROFILE
```

### 6.2. Lỗi `ENOENT: no such file or directory, open 'D:\tmp\neo-sessions.json'`

Nguyên nhân:
- `neo` muốn ghi session file nhưng thư mục/file chưa tồn tại

Cách sửa:

```powershell
New-Item -ItemType Directory -Force -Path "D:\tmp" | Out-Null
New-Item -ItemType File -Force -Path "D:\tmp\neo-sessions.json" | Out-Null
Set-Content -Path "D:\tmp\neo-sessions.json" -Value "{}"
```

### 6.3. Lỗi `Schema directory: Missing`

Cách sửa:

```powershell
New-Item -ItemType Directory -Force -Path "$env:HOME\.neo\schemas" | Out-Null
```

## 7. Kiểm tra kết nối `neo`

Sau khi đã mở Chrome với debug port:

```powershell
neo doctor
neo connect
neo tab
```

### 7.1. Kết quả mong đợi

`neo doctor` nên hiện ít nhất:
- Chrome CDP endpoint OK
- Browser tabs OK

`neo connect` nên hiện:

```text
Connected: Chrome/... @ http://localhost:9222
Session: __default__
```

### 7.2. Nếu `neo doctor` báo extension `Not found`

Trường hợp thực tế đã gặp:
- `neo tab` vẫn hiện `service_worker chrome-extension://.../background.js`
- nhưng `neo doctor` vẫn báo `Neo extension service worker: Not found`

Điều này có thể là false negative trên Windows. Khi đó:
- tin vào `neo tab` hơn
- nếu `neo tab` đã thấy `service_worker` của extension thì extension thường đã load

Ví dụ output tốt:

```text
[2] service_worker ... chrome-extension://.../background.js
```

## 8. Bắt capture

Chọn tab web đang thao tác:

```powershell
neo tab
neo tab 0
```

Sau đó trên Chrome:
1. đăng nhập
2. vào đúng màn hình báo cáo
3. thao tác tay:
   - chọn tham số
   - bấm `Báo cáo`
   - bấm `Xuất Excel`

Rồi quay lại PowerShell:

```powershell
neo capture domains
neo capture list --limit 50
```

Ví dụ thực tế đã thấy:
- frontend là `baocao.hanoi.vnpt.vn`
- backend API thật lại nằm ở `baocaobe.myhanoi.vn`

Đó là dấu hiệu rất quan trọng để viết downloader API sau này.

## 9. Các lệnh hay dùng

```powershell
neo doctor
neo connect
neo tab
neo tab 0
neo capture domains
neo capture list --limit 50
neo capture list baocaobe.myhanoi.vn --limit 30
neo reload
```

## 10. Khi `capture detail` bị lỗi `Not found`

Lỗi thực tế đã gặp:

```powershell
neo capture detail 362db0d1
Not found
```

Hoặc:
- `neo capture export ...` trả về `[]`
- `capture list` có summary nhưng không lấy được raw detail

Đây là dấu hiệu `neo` trên Windows đang bắt được summary nhưng không đọc lại được raw capture.

Khi gặp tình huống này, đừng mất thời gian debug quá sâu. Cách đi tiếp ổn hơn là:

### Cách A: dùng Chrome DevTools

1. Nhấn `F12`
2. Tab `Network`
3. Bật `Preserve log`
4. Lọc `baocaobe.myhanoi.vn/report-api`
5. Làm lại thao tác export
6. Chuột phải vào request
7. Chọn `Copy as cURL`

### Cách B: dùng Playwright log network

Đây là cách tốt nếu mục tiêu cuối cùng là viết downloader Python.

Ý tưởng:
- login bằng Playwright
- attach listener `request/response`
- chỉ log request tới `baocaobe.myhanoi.vn/report-api`

Thực tế sau đó đã tạo được các downloader API mới trong `api_transition/`.

Trạng thái hiện tại của nhánh chuyển đổi:
- `api_transition/downloaders.py` đã gom downloader theo helper chung `download_with_recipe()`
- mỗi downloader chỉ còn khai báo recipe và các input override đặc thù
- `api_transition/batch_download.py` đã có runner login 1 lần, reuse session, tự tính date range và chạy tuần tự toàn bộ report đã nối dây

## 11. Kết luận thực dụng về `neo` trên Windows

`neo` trên Windows phù hợp nhất để:
- xác định backend API thật
- xác nhận endpoint export
- biết payload tổng quát

`neo` trên Windows không phải lúc nào cũng đáng tin cho:
- đọc `capture detail`
- export raw capture

Khi bị kẹt ở bước detail/export:
- chuyển sang DevTools
- hoặc Playwright network logging

## 12. Server không GUI

Flow `neo + extension Chrome` không phù hợp để chạy trực tiếp lâu dài trên server không GUI.

Khuyến nghị:
- dùng máy Windows/GUI để bóc API một lần
- sau đó viết downloader Python để chạy trên server headless

Tức là:

```text
Máy GUI:
  dùng neo / DevTools / Playwright để capture

Server không GUI:
  login + gọi API trực tiếp + lưu file
```

## 13. Chuỗi lệnh mẫu đầy đủ trên Windows

### 13.1. Chuẩn bị môi trường

```powershell
$env:HOME = $env:USERPROFILE
$env:TEMP = "D:\tmp"
$env:TMP = "D:\tmp"
New-Item -ItemType Directory -Force -Path "$env:HOME\.neo" | Out-Null
New-Item -ItemType Directory -Force -Path "$env:HOME\.neo\schemas" | Out-Null
New-Item -ItemType Directory -Force -Path "D:\tmp" | Out-Null
New-Item -ItemType File -Force -Path "D:\tmp\neo-sessions.json" | Out-Null
Set-Content -Path "D:\tmp\neo-sessions.json" -Value "{}"
```

### 13.2. Mở Chrome

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="D:\neo\chrome-profile"
```

### 13.3. Kết nối

```powershell
cd D:\neo\neo
neo doctor
neo connect
neo tab
```

### 13.4. Capture

```powershell
neo tab 0
neo capture domains
neo capture list --limit 50
```

## 14. Tóm tắt lỗi thường gặp và cách sửa

| Lỗi | Nguyên nhân thường gặp | Cách xử lý |
|---|---|---|
| `Missing expression after unary operator '--'` | Gõ option Linux trực tiếp trong PowerShell | Gọi full `chrome.exe` trong một dòng |
| `chromium is not recognized` | Máy không có `chromium` trong PATH | Dùng đường dẫn đầy đủ tới `chrome.exe` hoặc `msedge.exe` |
| `path argument must be of type string` | Thiếu `HOME` | `\$env:HOME = \$env:USERPROFILE` |
| `ENOENT ... neo-sessions.json` | Thiếu file session | Tạo `D:\tmp\neo-sessions.json` |
| `Schema directory: Missing` | Thiếu `~\\.neo\\schemas` | Tạo thư mục này |
| `doctor` báo extension `Not found` | false negative hoặc load sai extension | kiểm tra bằng `neo tab`, reload extension, load lại `extension-dist` |
| `capture detail Not found` | giới hạn của `neo` trên Windows | chuyển sang DevTools hoặc Playwright logging |

## 15. File nên đọc tiếp trong repo này

Nếu mục tiêu là thay click UI bằng API downloader, xem thêm:
- [api_transition/README.md](/home/vtst/baocaohanoi/api_transition/README.md)
- [api_transition/MIGRATION_STATUS.md](/home/vtst/baocaohanoi/api_transition/MIGRATION_STATUS.md)
- [api_transition/capture_report_api.py](/home/vtst/baocaohanoi/api_transition/capture_report_api.py)
- [api_transition/capture_with_legacy_flow.py](/home/vtst/baocaohanoi/api_transition/capture_with_legacy_flow.py)
