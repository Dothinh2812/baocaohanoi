# API Transition

Thư mục này chứa phiên bản chuyển đổi dần từ luồng tải báo cáo bằng click UI sang gọi API trực tiếp.

Nguyên tắc:
- Không sửa các downloader hiện tại ở thư mục gốc.
- Mỗi báo cáo mới sẽ được thêm riêng trong thư mục này.
- Có thể thử nghiệm song song với pipeline cũ.

## Cấu trúc

- `settings.py`: nạp cấu hình từ `.env`
- `auth.py`: login bằng Playwright + OTP, bắt `Authorization`
- `report_api_client.py`: helper gọi `report-api`
- `catalog.py`: danh mục hàm download cũ và trạng thái migration
- `MIGRATION_STATUS.md`: nhật ký tổng hợp các hàm đã chuyển đổi thành công và các hàm đang lỗi/tạm hoãn
- `capture_report_api.py`: capture request/response và sinh recipe JSON
- `capture_with_legacy_flow.py`: chạy hàm download cũ ở chế độ headless và tự bắt recipe
- `downloaders.py`: các hàm downloader API mới
- `cts_api.py`: module API riêng cho CTS, dùng login trong `cts.py` và tải file binary trực tiếp
- `batch_download.py`: runner batch login 1 lần, tái sử dụng session và tải tuần tự các report đã wired
- `export_from_recipe.py`: CLI generic chạy bằng recipe
- `recipes/`: recipe đã capture và xác nhận
- `downloads/`: thư mục tải file mặc định theo nhóm nghiệp vụ
- `export_c11_api.py`: PoC export C1.1 qua API

## Cấu trúc thư mục tải file

Mặc định các downloader mới sẽ lưu vào `api_transition/downloads/` và tách theo nhóm:

- `api_transition/downloads/chi_tieu_c/`: nhóm chỉ tiêu C
- `api_transition/downloads/chi_tieu_i/`: nhóm chỉ tiêu I
- `api_transition/downloads/cau_hinh_tu_dong/`: nhóm cấu hình tự động
- `api_transition/downloads/ghtt/`: nhóm GHTT
- `api_transition/downloads/cts/`: nhóm báo cáo CTS độc lập
- `api_transition/downloads/xac_minh_tam_dung/`: nhóm xác minh tạm dừng
- `api_transition/downloads/kq_tiep_thi/`: nhóm kết quả tiếp thị
- `api_transition/downloads/vat_tu_thu_hoi/`: nhóm vật tư thu hồi
- `api_transition/downloads/misc/`: nơi rơi mặc định khi chạy generic bằng `--recipe`

Nếu cần, vẫn có thể override bằng `--output-dir`.

## Kiến trúc downloader mới

`downloaders.py` hiện là lớp wrapper chung cho toàn bộ downloader API mới.

Điểm chính:
- `download_with_recipe()` là entrypoint chung: load recipe, login nếu chưa có `session`, resolve `month_id` từ `month_label` khi cần, merge override vào `lstInputParams`, gọi export API rồi lưu file.
- `group_output_dir()` chuẩn hóa thư mục đầu ra theo nhóm nghiệp vụ.
- Mỗi downloader nghiệp vụ chỉ còn khai báo phần khác nhau: `recipe_name`, `output_name`, các tham số override như `ptrungtamid`, `vthoigian`, `vngay_bd`, `vngay_kt`, `pdv`.
- Tất cả downloader đều nhận `session=None`. Nếu truyền `session`, downloader sẽ dùng lại `Authorization` và cookie đã capture, không login lại.

Các downloader hiện đã được nối dây trong file này gồm:
- nhóm Chỉ tiêu C, Chỉ tiêu I, GHTT, xác minh tạm dừng, kết quả tiếp thị, vật tư thu hồi
- cấu hình tự động tổng hợp, thay thế, chi tiết
- `download_quyet_toan_vattu_api()` dùng recipe `quyet_toan_vattu_q2_2026.json`

Lưu ý:
- `quyet_toan_vattu` đã được wired trong `downloaders.py` và `batch_download.py`
- hiện chưa có shortcut CLI riêng trong `export_from_recipe.py`, nên cách chạy trực tiếp thuận tiện nhất là import Python hoặc dùng batch
- `cts_api.py` là module độc lập, không dùng `report-api` của `baocao.hanoi` và không đi qua `download_with_recipe()`

## Yêu cầu

- Dùng chung `.env` hiện tại ở root repo
- Đã cài dependencies của project
- Đã cài Playwright Chromium

## Chạy thử C1.1

Linux/macOS:

```bash
python3 api_transition/export_c11_api.py --headed --month-id 98944548
```

Windows:

```powershell
python api_transition/export_c11_api.py --headed --month-id 98944548
```

Nếu muốn để script tự map kỳ báo cáo theo nhãn:

```bash
python3 api_transition/export_c11_api.py --headed --month-label "Tháng 04/2026"
```

Hoặc dùng downloader generic:

```bash
python3 api_transition/export_from_recipe.py --c11 --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --c12 --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --c13 --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --c14 --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --c14-chi-tiet --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --c11-chi-tiet --headed --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/export_from_recipe.py --c12-chi-tiet-sm1 --headed --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/export_from_recipe.py --c12-chi-tiet-sm2 --headed --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/export_from_recipe.py --i15 --headed --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/export_from_recipe.py --i15-k2 --headed --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/export_from_recipe.py --ghtt-hni --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --ghtt-sontay --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --ghtt-nvktdb --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --xac-minh-tam-dung --headed --start-date "01/04/2026" --end-date "16/04/2026"
python3 api_transition/export_from_recipe.py --kq-tiep-thi --headed --start-date "16/04/2026" --end-date "16/04/2026"
python3 api_transition/export_from_recipe.py --vattu-thuhoi --headed --start-date "24/11/2025" --end-date "16/04/2026"
python3 api_transition/export_from_recipe.py --cau-hinh-tu-dong-ptm --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --cau-hinh-tu-dong-thay-the --headed --month-id 98944548
python3 api_transition/export_from_recipe.py --cau-hinh-tu-dong-chi-tiet --headed --month-id 98944548
```

Các lệnh trên nếu không truyền `--output-dir` sẽ tự lưu vào thư mục nhóm tương ứng dưới `api_transition/downloads/`.

## Batch Download

`batch_download.py` là runner vận hành cho giai đoạn chuyển đổi hiện tại: login 1 lần cho nhóm `baocao.hanoi`, rồi chạy tuần tự toàn bộ downloader đã được nối dây.

Hành vi chính:
- tự tính các mốc ngày từ `REPORT_MONTH` / `REPORT_YEAR`
- ánh xạ tham số theo `params_type`: `month`, `date_range`, `calendar_month`, `t_minus_1`, `date_range_long`
- tái sử dụng cùng một `session` cho các report `baocao.hanoi`
- các report độc lập như `CTS SHC ngày` sẽ tự login session riêng, không dùng shared session ở trên
- retry các lỗi timeout theo `RETRY_TIMEOUTS = [120, 180, 300]`
- hỗ trợ `--only`, `--skip`, `--list`

Ví dụ:

```bash
python3 api_transition/batch_download.py
python3 api_transition/batch_download.py --month 5 --year 2026 --month-id 99001234
python3 api_transition/batch_download.py --only "C1.1" "C1.2"
python3 api_transition/batch_download.py --skip "Vật tư thu hồi"
python3 api_transition/batch_download.py --list
```

Import từ Python:

```python
from api_transition.batch_download import run_batch_download

results = run_batch_download(
    report_month=4,
    report_year=2026,
    month_id="98944548",
)
```

Danh sách report hiện được batch gọi:
- `C1.1`, `C1.2`, `C1.3`, `C1.4`, `C1.4 Chi tiết`
- `C1.1 Chi tiết`, `C1.2 Chi tiết SM1`, `C1.2 Chi tiết SM2`
- `I1.5`, `I1.5 K2`
- `GHTT HNI`, `GHTT Sơn Tây`, `GHTT NVKT DB`
- `Xác minh tạm dừng`, `Kết quả tiếp thị`, `CTS SHC ngày`
- `Vật tư thu hồi`, `Quyết toán vật tư`
- `Cấu hình tự động PTM`, `Cấu hình tự động Thay thế`, `Cấu hình tự động Chi tiết`

`CTS SHC ngày` hiện được batch gọi với tham số ngày `T-1` dưới dạng `report_date`, tức mặc định lấy ngày hôm qua theo định dạng `dd/mm/yyyy`.

`I1.5` và `I1.5 K2` đã được xác nhận tải thành công bằng downloader API mới trên môi trường thực.
Ba recipe `GHTT` đã được capture, đã có downloader API riêng trong `api_transition/`, và đã được xác nhận tải thành công trên môi trường thực.
`xac_minh_tam_dung_download` cũng đã được xác nhận tải thành công bằng downloader API mới trên môi trường thực.
`kq_tiep_thi_download` cũng đã được xác nhận tải thành công bằng downloader API mới trên môi trường thực.
`download_report_vattu_thuhoi` cũng đã được xác nhận tải thành công bằng downloader API mới trên môi trường thực.
Báo cáo `cấu hình tự động` đã được tách rõ thành 2 downloader API riêng trong `api_transition/`:
- `download_cau_hinh_tu_dong_ptm_api()` lưu file `cau_hinh_tu_dong_ptm.xlsx`
- `download_cau_hinh_tu_dong_thay_the_api()` lưu file `cau_hinh_tu_dong_thay_the.xlsx`
Ngoài ra đã có thêm `download_cau_hinh_tu_dong_chi_tiet_api()` lưu file `cau_hinh_tu_dong_chi_tiet.xlsx`.
`download_cau_hinh_tu_dong_chi_tiet_api()` đã được xác nhận tải thành công bằng downloader API mới trên môi trường thực.

Các mục còn lại chưa hoàn tất trong vòng chuyển đổi hiện tại là:
- `download_report_c11_chitiet_SM2`
- `download_report_c15`
- `download_report_c15_chitiet`

## Capture report mới và sinh recipe

Ví dụ capture C1.2:

```bash
python3 api_transition/capture_report_api.py \
  --headed \
  --name c12_q2_2026 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=535021"
```

Sau khi browser mở:

1. Thao tác tay như luồng cũ
2. Bấm `Báo cáo`
3. Bấm `Xuất Excel`
4. Quay lại terminal và nhấn Enter

Script sẽ lưu:
- log JSONL trong `api_transition/captures/`
- recipe trong `api_transition/recipes/`

## Capture trên server không GUI bằng hàm cũ

Nếu server không có X server, không dùng `--headed`. Thay vào đó chạy chính hàm download cũ để nó tự thao tác UI, còn script sẽ capture `report-api` ở nền:

```bash
python3 api_transition/capture_with_legacy_flow.py \
  --name c12_q2_2026 \
  --legacy-func download_report_c12 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=535021" \
  --report-month "Tháng 04/2026"
```

Ví dụ C1.1:

```bash
python3 api_transition/capture_with_legacy_flow.py \
  --name c11_q2_2026_auto \
  --legacy-func download_report_c11 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=534964&menu_id=535020" \
  --report-month "Tháng 04/2026"
```

Nhóm tiếp theo nên capture sau khi tạm bỏ qua C1.5 là 4 báo cáo chi tiết C1.1/C1.2:

```bash
python3 api_transition/capture_with_legacy_flow.py \
  --name c11_chitiet_q2_2026 \
  --legacy-func download_report_c11_chitiet \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"

python3 api_transition/capture_with_legacy_flow.py \
  --name c11_chitiet_sm2_q2_2026 \
  --legacy-func download_report_c11_chitiet_SM2 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"

python3 api_transition/capture_with_legacy_flow.py \
  --name c12_chitiet_sm1_q2_2026 \
  --legacy-func download_report_c12_chitiet_SM1 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"

python3 api_transition/capture_with_legacy_flow.py \
  --name c12_chitiet_sm2_q2_2026 \
  --legacy-func download_report_c12_chitiet_SM2 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"
```

Sau đó export lại bằng recipe:

```bash
python3 api_transition/export_from_recipe.py \
  --headed \
  --recipe c12_q2_2026 \
  --month-id <pthang>
```

## Ghi đè tham số input

Để capture `xac_minh_tam_dung_download`:

```bash
python3 api_transition/capture_with_legacy_flow.py \
  --name xac_minh_tam_dung_q2_2026 \
  --legacy-func xac_minh_tam_dung_download \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267844&menu_id=276199"
```

Nhóm tiếp theo nên capture là `kq_tiep_thi_download`:

```bash
python3 api_transition/capture_with_legacy_flow.py \
  --name kq_tiep_thi_q2_2026 \
  --legacy-func kq_tiep_thi_download \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=257495&menu_id=276101"
```

Sau khi capture xong, chạy bản API mới:

```bash
python3 api_transition/export_from_recipe.py \
  --kq-tiep-thi \
  --start-date "16/04/2026" \
  --end-date "16/04/2026"
```

Mục tiếp theo còn lại là `download_report_vattu_thuhoi`.

`report_url` của luồng cũ đã xác định được là:

```text
https://baocao.hanoi.vnpt.vn/report/report-info?id=270922&menu_id=276242
```

Để capture `download_report_vattu_thuhoi`:

```bash
python3 api_transition/capture_with_legacy_flow.py \
  --name vattu_thuhoi_q2_2026 \
  --legacy-func download_report_vattu_thuhoi \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=270922&menu_id=276242"
```

Sau khi capture xong, chạy bản API mới:

```bash
python3 api_transition/export_from_recipe.py \
  --vattu-thuhoi \
  --start-date "24/11/2025" \
  --end-date "16/04/2026"
```

Đối với báo cáo `cấu hình tự động`, dùng 2 downloader riêng:

PTM:

```bash
python3 api_transition/export_from_recipe.py \
  --cau-hinh-tu-dong-ptm \
  --month-id 98944548
```

Thay Thế:

```bash
python3 api_transition/export_from_recipe.py \
  --cau-hinh-tu-dong-thay-the \
  --month-id 98944548
```

Ánh xạ rõ:
- `--cau-hinh-tu-dong-ptm` gọi `download_cau_hinh_tu_dong_ptm_api()`, ép `pdv=1`, lưu `cau_hinh_tu_dong_ptm.xlsx`
- `--cau-hinh-tu-dong-thay-the` gọi `download_cau_hinh_tu_dong_thay_the_api()`, ép `pdv=13`, lưu `cau_hinh_tu_dong_thay_the.xlsx`

Đối với báo cáo `cấu hình tự động chi tiết`:

```bash
python3 api_transition/export_from_recipe.py \
  --cau-hinh-tu-dong-chi-tiet \
  --month-id 98944548
```

Ánh xạ rõ:
- `--cau-hinh-tu-dong-chi-tiet` gọi `download_cau_hinh_tu_dong_chi_tiet_api()`, dùng `pthang`, lưu `cau_hinh_tu_dong_chi_tiet.xlsx`

Có thể thay tham số trong payload đã capture:

```bash
python3 api_transition/export_from_recipe.py \
  --headed \
  --recipe c11_q2_2026 \
  --set ptrungtamid=14324 \
  --month-id 98944548
```

## Mục tiêu chuyển đổi dần

1. Xác nhận từng báo cáo tải được qua API.
2. So sánh file đầu ra với luồng cũ.
3. Khi ổn định mới thay thế từng downloader ở code chính.
