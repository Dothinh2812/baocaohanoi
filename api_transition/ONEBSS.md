# OneBSS Downloaders

Tài liệu này mô tả chi tiết các module OneBSS trong `api_transition/`, bao gồm:

- [onebss_auth.py](./chua_dung_den/onebss_auth.py)
- [onebss_report_client.py](./chua_dung_den/onebss_report_client.py)
- [onebss_downloaders.py](./chua_dung_den/onebss_downloaders.py)

Mục tiêu của nhóm module này là tự động hóa tải báo cáo từ hệ OneBSS theo 2 luồng khác nhau:

- luồng BI API `api-onebss.vnpt.vn/web-report/report/bi/...`
- luồng `report-onebss.vnpt.vn/ReportViewer.aspx`

Khác với nhóm downloader cũ ở `downloaders.py`, các module OneBSS hiện không dùng recipe JSON. Logic được tổ chức theo hướng:

- `onebss_auth.py`: login và tạo session
- `onebss_report_client.py`: HTTP client và helper dựng payload/report parameter
- `onebss_downloaders.py`: các hàm downloader dùng thật

## 1. Tổng quan kiến trúc

### 1.1 Luồng BI API

Luồng này dùng cho các báo cáo trong giao diện `onebss.vnpt.vn/#/report/bi`.

Chu trình tải:

1. Login OneBSS bằng Playwright.
2. Bắt `Authorization` và các header nghiệp vụ từ request thật.
3. Gọi `GET /web-report/report/bi/parameters?report=...` để lấy metadata tham số.
4. Nếu có dropdown phụ thuộc, gọi `POST /web-report/report/bi/parameters` để refresh tham số.
5. Dựng payload `run_v3`.
6. Gọi `POST /web-report/report/bi/run_v3`.
7. Nhận file Excel dạng binary và lưu vào disk.

### 1.2 Luồng ReportViewer

Luồng này dùng cho các báo cáo render qua `report-onebss.vnpt.vn/ReportViewer.aspx`.

Chu trình tải:

1. Login OneBSS bằng Playwright.
2. Lấy JWT token từ session OneBSS.
3. Mở `https://report-onebss.vnpt.vn/?baocao_id=<id>&token=<jwt>`.
4. Mở `ReportViewer.aspx?baocao_id=...&...tham_số...`.
5. Lấy `guid` từ input ẩn `#ssrBaoCaoCachedReportNameInputId`.
6. Export bằng `GET /ReportViewer.aspx?guid=<guid>&action=export&format=excel`.
7. Dùng `page.expect_download()` để nhận file xuất.

Điểm quan trọng:

- Luồng ReportViewer không dùng `POST JSON`.
- Export thực tế là `GET` với `guid`.
- File export có thể là `.xls` XML Spreadsheet 2003, không phải `.xlsx`.

## 2. Cấu hình và biến môi trường

Các module OneBSS đọc `.env` ở root repo thông qua `load_dotenv(ROOT_DIR / ".env")`.

### 2.1 Biến chính

- `BAOCAO_USERNAME`
- `BAOCAO_PASSWORD`
- `OTP_FILE_PATH`

Đây là 3 biến cốt lõi để login OneBSS.

### 2.2 Biến tùy chọn của OneBSS

- `ONEBSS_BASE_URL`
- `ONEBSS_LOGIN_URL`
- `ONEBSS_TOKEN_PAGE_URL`
- `ONEBSS_REQUEST_HOST_FILTER`
- `ONEBSS_API_BASE_URL`
- `ONEBSS_SELECTED_MENU_ID`
- `ONEBSS_SELECTED_PATH`
- `ONEBSS_MAC_ADDRESS`
- `ONEBSS_TOKEN_ID`
- `ONEBSS_API_KEY`

### 2.3 Biến timeout/OTP

- `OTP_MAX_AGE_SECONDS`
- `ONEBSS_OTP_MAX_RETRIES`
- `ONEBSS_OTP_RETRY_INTERVAL_SECONDS`
- `PAGE_LOAD_TIMEOUT`
- `ACCEPT_DOWNLOADS`

## 3. Module `onebss_auth.py`

Module này chịu trách nhiệm login OneBSS, đọc OTP, bắt token và dựng session tái sử dụng được.

### 3.1 `OneBSSSettings`

Class cấu hình tập trung cho OneBSS.

Các thuộc tính đáng chú ý:

- `USERNAME`, `PASSWORD`
- `BASE_URL`, `LOGIN_URL`
- `API_BASE_URL`
- `DEFAULT_SELECTED_MENU_ID`
- `DEFAULT_SELECTED_PATH`
- `DEFAULT_MAC_ADDRESS`
- `DEFAULT_TOKEN_ID`
- `DEFAULT_API_KEY`
- `OTP_FILE_PATH`
- `PAGE_LOAD_TIMEOUT`

`validate()` sẽ báo lỗi nếu thiếu:

- `BAOCAO_USERNAME`
- `BAOCAO_PASSWORD`
- `OTP_FILE_PATH`

### 3.2 `build_cookie_header(cookies)`

Tạo chuỗi cookie header từ danh sách cookie Playwright.

Kết quả có dạng:

```text
name1=value1; name2=value2
```

Hàm này được dùng khi cần gọi HTTP ngoài browser context.

### 3.3 `read_otp_from_file(...)`

Đọc OTP 6 chữ số từ file OTP.

Hành vi:

- kiểm tra file có tồn tại không
- chỉ chấp nhận OTP mới hơn `OTP_MAX_AGE_SECONDS`
- retry nhiều lần theo `ONEBSS_OTP_MAX_RETRIES`
- mặc định xóa nội dung file sau khi đọc thành công

Đây là primitive nền cho login tự động.

### 3.4 `_try_fill_otp(page)`

Nếu trang login hiện ô OTP:

- đọc OTP từ file
- điền OTP
- bấm xác nhận

Nếu không đọc được OTP tự động:

- script chờ người dùng nhập tay

### 3.5 `login(headless=True)`

Hàm login chính.

Các bước:

1. mở Chromium bằng Playwright
2. vào `ONEBSS_LOGIN_URL`
3. điền username/password
4. check `remember me`
5. click login
6. xử lý OTP nếu có
7. trả về:

```python
(playwright, browser, context, page)
```

Lưu ý:

- mặc định `headless=True`
- phù hợp môi trường server không có GUI

### 3.6 `extract_token_from_storage(page)`

Quét `localStorage` và `sessionStorage` để lấy token nếu request network chưa lộ `Authorization`.

Module hỗ trợ nhận token ở 2 dạng:

- chuỗi `Bearer ...`
- JWT thô
- JSON lồng có khóa như `token`, `access_token`, `accessToken`, `authToken`

### 3.7 `capture_authorization(page, target_url="", timeout_seconds=30, host_filter="")`

Đây là hàm quan trọng nhất của module auth.

Nó:

- lắng nghe request Playwright
- lọc request theo host
- bắt:
  - `authorization`
  - `token-id`
  - `selectedmenuid`
  - `selectedpath`
  - `mac-address`
  - `apikey`
- nếu chưa có `Authorization`, fallback sang đọc token từ storage

Giá trị trả về là `auth_state`:

```python
{
    "authorization": "Bearer ...",
    "token": "...",
    "user_agent": "...",
    "accept": "...",
    "referer": "...",
    "source": "...",
    "token-id": "...",
    "selectedmenuid": "...",
    "selectedpath": "...",
    "mac-address": "...",
    "apikey": "..."
}
```

### 3.8 `capture_token(...)`

Wrapper mỏng trên `capture_authorization()`, chỉ trả về `token`.

### 3.9 `make_common_headers(auth_state, cookies, extra_headers=None)`

Dựng bộ header chuẩn cho OneBSS BI API:

- `Authorization`
- `Accept`
- `Referer`
- `User-Agent`
- `Cookie`
- `token-id`
- `selectedmenuid`
- `selectedpath`
- `mac-address`

### 3.10 `create_session(headed=False, token_page_url="", extra_headers=None)`

Entry point khuyến nghị để khởi tạo session dùng lại.

Các bước:

1. login
2. capture token/header
3. dựng header chung
4. trả về `session`

Shape của `session`:

```python
{
    "token": "...",
    "api_base_url": "https://api-onebss.vnpt.vn",
    "headers": {...},
    "auth_state": {...},
    "playwright": ...,
    "browser": ...,
    "context": ...,
    "page": ...,
}
```

### 3.11 `close_session(session)`

Đóng `browser` và `playwright` an toàn.

Khuyến nghị:

- luôn gọi trong `finally`
- hoặc để downloader generic tự lo khi `session=None`

## 4. Module `onebss_report_client.py`

Module này chứa HTTP client và helper xử lý metadata/payload cho nhóm BI report.

### 4.1 `sanitize_filename(filename)`

Làm sạch tên file:

- thay ký tự cấm bằng `_`
- gom khoảng trắng
- tránh file name rỗng

### 4.2 `build_onebss_headers(session_headers, include_apikey=False, extra_headers=None)`

Clone bộ header từ session và tùy chọn thêm:

- `apikey: x`

`apikey` được dùng ở một số API như:

- `GET /web-report/report/bi/parameters`
- `POST /web-report/report/bi/run_v3`

### 4.3 `http_json_request(...)`

Helper gọi HTTP và parse response JSON.

Đặc điểm:

- hỗ trợ `GET`/`POST`
- nếu có `payload`, tự encode JSON UTF-8
- ném `RuntimeError` nếu:
  - HTTP lỗi
  - lỗi kết nối
  - response không phải JSON

### 4.4 `http_binary_request(...)`

Helper gọi HTTP và nhận response binary.

Trả về:

```python
{
    "body": b"...",
    "headers": {...}
}
```

Hàm này được dùng cho:

- export BI report

Không dùng cho `ReportViewer` mới, vì `ReportViewer` export phải đi qua Playwright download.

### 4.5 `build_report_parameters_url(report_path, api_base_url="")`

Sinh URL:

```text
https://api-onebss.vnpt.vn/web-report/report/bi/parameters?report=<report_path>
```

### 4.6 `get_report_parameters(report_path, headers, api_base_url="", timeout=120)`

Gọi metadata API để lấy danh sách tham số report.

Response thường chứa:

```json
{
  "listOfParamNameValues": {
    "item": [...]
  }
}
```

### 4.7 `refresh_report_parameters(report_path, parameter_items, headers, api_base_url="", timeout=120)`

Dùng khi report có tham số phụ thuộc nhau, ví dụ:

- chọn `TT_ID` xong mới nạp được `DOI_ID`
- chọn batch code xong mới nạp lại danh sách nhân viên

Payload gửi lên:

```json
{
  "report": "...",
  "parameterNameValues": {
    "listOfParamNameValues": {
      "item": [...]
    }
  }
}
```

### 4.8 `prepare_parameter_item(parameter_definition, value)`

Chuyển metadata parameter thành item tối giản cho request refresh.

Support:

- single-value
- multi-value

Ví dụ kết quả:

```python
{
    "defaultValue": "14324",
    "temp": [],
    "name": "TT_ID",
    "values": {"item": "14324"},
}
```

### 4.9 `fill_parameter_values(parameter_items, overrides)`

Điền giá trị cuối cùng vào full danh sách parameter items để chuẩn bị export.

Nó sẽ xử lý:

- `defaultValue`
- `value`
- `temp`
- `mvalue`
- `values`
- `options`

Đây là bước quan trọng để payload `run_v3` khớp với thứ frontend OneBSS thực sự gửi.

### 4.10 `build_run_v3_payload(baocao_id, report_path, items, file_name="", export_type="xlsx", multiselect=1)`

Dựng payload export BI:

```python
{
    "baocao_id": ...,
    "report": "...",
    "type": "xlsx",
    "multiselect": 1,
    "file_name": "...",
    "items": [...],
}
```

### 4.11 `run_report_export(payload, headers, api_base_url="", timeout=120)`

Gọi:

```text
POST /web-report/report/bi/run_v3
```

và trả binary response.

### 4.12 `save_binary_export_file(export_response, output_dir="", output_name="")`

Lưu binary response ra file.

Mặc định lưu dưới:

```text
api_transition/downloads/onebss/
```

## 5. Module `onebss_downloaders.py`

Đây là lớp orchestration và là nơi chứa các hàm downloader công khai.

### 5.1 `group_output_dir(group_name)`

Trả về thư mục output theo nhóm:

```python
group_output_dir("onebss")
```

### 5.2 `extract_parameter_items(parameters_response)`

Rút `item` từ response metadata.

Nếu `item` là object đơn, hàm chuẩn hóa thành list để downstream code luôn xử lý thống nhất.

### 5.3 `build_refresh_items(parameter_items, refresh_values)`

Lọc ra những parameter cần refresh theo map `refresh_values`.

Ví dụ:

```python
refresh_values = {
    "TT_ID": "14324",
    "DOI_ID": "0",
}
```

### 5.4 `_save_playwright_download(download, output_dir, output_name)`

Lưu file download từ Playwright.

Hành vi đặc biệt:

- nếu `suggested_filename` từ server có đuôi khác `output_name`, hàm tự đổi sang đuôi đúng
- dùng để tránh lỗi kiểu file `.xls` bị ép thành `.xlsx`

### 5.5 `download_onebss_reportviewer_report(...)`

Generic downloader cho `report-onebss.vnpt.vn/ReportViewer.aspx`.

Input:

- `report_id`
- `report_query_params`
- `output_dir`
- `output_name`

Luồng:

1. tạo session nếu chưa có
2. lấy JWT token từ session OneBSS
3. mở bootstrap URL với `baocao_id + token`
4. mở `ReportViewer.aspx?...tham_số...`
5. đọc `guid`
6. export bằng `guid`
7. chờ browser download

Đây là luồng chuẩn cho các report kiểu SharpShooter/ReportViewer.

### 5.6 `download_onebss_report(...)`

Generic downloader cho BI report.

Input chính:

- `report_path`
- `baocao_id`
- `overrides`
- `refresh_values`
- `session`

Luồng:

1. tạo session nếu cần
2. lấy metadata tham số
3. refresh dropdown phụ thuộc nếu cần
4. fill giá trị cuối
5. dựng payload `run_v3`
6. export và lưu file

Đây là entrypoint chung cho hầu hết báo cáo BI OneBSS.

### 5.7 Các downloader công khai hiện có

#### `download_hni_pttb_001(...)`

Tải report:

```text
TINH/HANOI/HNI_PTTB_001/RP_HNI_PTTB_001
```

Tham số:

- `unit_id`
- `team_id`
- `service_ids`
- `congnghe_id`

Đây là wrapper nghiệp vụ đầu tiên cho report `40618`.

#### `download_bc_phieu_ton_dv_chi_tiet_hni(...)`

Wrapper cùng report `40618`, đặt tên theo ngữ nghĩa nghiệp vụ:

- báo cáo phiếu tồn dịch vụ chi tiết HNI

Tham số:

- `unit_id`
- `team_id`
- `service_ids`
- `congnghe_id`

#### `download_bc_ton_sua_chua_sontay_2026(...)`

Tải report:

```text
TINH/HANOI/HNI_BHSC_005/RP_HNI_BHSC_005
```

`baocao_id = 40622`

Tham số:

- `unit_id`
- `team_id`
- `service_id`

Đây là báo cáo tồn sửa chữa khu vực Sơn Tây.

#### `download_bc_chi_tiet_ket_qua_cskh_uc3_sontay(...)`

Downloader cho luồng `ReportViewer`.

Tham số:

- `customer_batch_code`
- `start_date`
- `end_date`
- `unit_id`
- `employee_id`
- `region_id`
- `region_text`
- `unit_text`
- `employee_text`

Hàm tự chuẩn hóa `customer_batch_code` để đảm bảo có dấu nháy đơn:

```python
"'UC3_CSKH_042026'"
```

Lưu ý:

- file export thực tế là `.xls`
- nếu server trả `suggested_filename` khác đuôi, helper sẽ tự sửa đuôi

### 5.8 `dump_payload_preview(...)`

Dùng để debug payload BI trước khi export thật.

Kết quả trả về là object payload của `run_v3`.

Rất hữu ích khi:

- muốn kiểm tra mapping `TT_ID`, `DOI_ID`, `DICHVUVT_ID`
- chưa chắc `overrides` đã đúng
- cần so với capture frontend

### 5.9 `payload_to_pretty_json(payload)`

Format payload thành JSON đẹp để log/debug.

## 6. Tham số nghiệp vụ chung

OneBSS BI report hiện lặp lại một số tham số chung:

- `unit_id`
  thường map vào `TT_ID` hoặc field đơn vị tương tự
- `team_id`
  thường map vào `DOI_ID`
- `service_id`
  cho single-select service
- `service_ids`
  cho multi-select service
- `congnghe_id`
- `region_id`
- `employee_id`
- `customer_batch_code`
- `start_date`, `end_date`

Quy tắc thực tế:

- nếu tham số ảnh hưởng dropdown khác, đưa vào `refresh_values`
- nếu tham số là giá trị export cuối, đưa vào `overrides`

Ví dụ:

```python
download_onebss_report(
    report_path="TINH/HANOI/HNI_PTTB_001/RP_HNI_PTTB_001",
    baocao_id=40618,
    refresh_values={
        "TT_ID": "14324",
        "DOI_ID": "0",
    },
    overrides={
        "TT_ID": "14324",
        "DOI_ID": "0",
        "DICHVUVT_ID": ["1", "4", "7"],
        "CONGNGHE_ID": "0",
    },
)
```

## 7. Cách dùng thực tế

### 7.1 Tạo session 1 lần và tái sử dụng

```python
from api_transition.chua_dung_den.onebss_auth import create_session, close_session
from api_transition.chua_dung_den.onebss_downloaders import (
    download_bc_phieu_ton_dv_chi_tiet_hni,
    download_bc_ton_sua_chua_sontay_2026,
)

session = create_session()
try:
    path1 = download_bc_phieu_ton_dv_chi_tiet_hni(session=session)
    path2 = download_bc_ton_sua_chua_sontay_2026(session=session)
    print(path1)
    print(path2)
finally:
    close_session(session)
```

Đây là cách tốt nhất khi chạy nhiều báo cáo liên tiếp.

### 7.2 Chạy một downloader BI đơn lẻ

```bash
cd /home/vtst/baocaohanoi && python3 -c 'from api_transition.chua_dung_den.onebss_downloaders import download_bc_ton_sua_chua_sontay_2026; print(download_bc_ton_sua_chua_sontay_2026())'
```

### 7.3 Chạy downloader `ReportViewer`

```bash
cd /home/vtst/baocaohanoi && python3 -c 'from api_transition.chua_dung_den.onebss_downloaders import download_bc_chi_tiet_ket_qua_cskh_uc3_sontay; print(download_bc_chi_tiet_ket_qua_cskh_uc3_sontay(output_dir="/home/vtst/baocaohanoi/api_transition/downloads/onebss"))'
```

### 7.4 Xem trước payload BI

```python
from api_transition.chua_dung_den.onebss_downloaders import dump_payload_preview, payload_to_pretty_json

payload = dump_payload_preview(
    report_path="TINH/HANOI/HNI_PTTB_001/RP_HNI_PTTB_001",
    baocao_id=40618,
    refresh_values={"TT_ID": "14324", "DOI_ID": "0"},
    overrides={
        "TT_ID": "14324",
        "DOI_ID": "0",
        "DICHVUVT_ID": ["1", "4", "7", "8", "11", "12"],
        "CONGNGHE_ID": "0",
    },
)

print(payload_to_pretty_json(payload))
```

## 8. Lỗi thường gặp và cách xử lý

### 8.1 Không đọc được OTP

Triệu chứng:

- login dừng ở bước OTP

Xử lý:

- kiểm tra `OTP_FILE_PATH`
- kiểm tra file OTP có mã mới không
- kiểm tra `OTP_MAX_AGE_SECONDS`

### 8.2 `401 Unauthorized` khi gọi `report-onebss.vnpt.vn`

Nguyên nhân thường gặp:

- cố gọi `ReportViewer` bằng HTTP client ngoài browser
- chưa bootstrap bằng `token`
- chưa render report để lấy `guid`

Trạng thái hiện tại của code:

- đã xử lý đúng bằng Playwright browser flow

### 8.3 File tải về mở lỗi vì đuôi `.xlsx`

Nguyên nhân:

- server trả file XML Spreadsheet 2003 nhưng tên local bị ép `.xlsx`

Trạng thái hiện tại:

- `_save_playwright_download()` đã tự sửa sang đuôi đúng theo `suggested_filename`

### 8.4 `DOI_ID` hoặc `employee_id` sai sau khi đổi đơn vị

Nguyên nhân:

- dùng id con của trung tâm cũ sau khi đổi `unit_id`

Nguyên tắc:

- đổi `unit_id` xong phải refresh lại dropdown phụ thuộc
- nếu chọn “Tất cả” thì dùng `0`

## 9. Khi nào cần thêm downloader mới

Thêm downloader mới khi:

- anh đã có capture ổn định của report mới
- đã xác định report thuộc BI API hay ReportViewer
- đã biết tham số nào cần refresh
- đã biết output file là binary `.xlsx`, `.xls`, hay định dạng khác

Quy trình khuyến nghị:

1. capture request/response thật
2. xác định report thuộc luồng nào
3. nếu là BI report: thêm wrapper trên `download_onebss_report()`
4. nếu là ReportViewer: thêm wrapper trên `download_onebss_reportviewer_report()`
5. kiểm tra file output thực tế

## 10. Hướng cải tiến

Hiện trạng OneBSS đang dùng code-first wrappers, chưa recipe-driven như `downloaders.py`.

Hướng cải tiến hợp lý trong tương lai:

- bổ sung recipe JSON riêng cho OneBSS BI
- chuẩn hóa schema parameter
- tách rõ report type:
  - BI API
  - ReportViewer
- sinh wrapper mỏng tự động từ capture

Hiện tại, với số lượng báo cáo chưa quá lớn, kiến trúc hiện có là đủ thực dụng:

- auth tập trung
- client tập trung
- downloader wrappers mỏng
