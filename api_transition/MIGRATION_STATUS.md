# Migration Status

Tài liệu này ghi lại trạng thái chuyển đổi từ downloader cũ dựa trên click UI sang downloader mới trong `api_transition/`.

Nguyên tắc áp dụng:
- Không sửa các downloader đang chạy ở thư mục gốc.
- Mỗi hàm chuyển đổi thành công sẽ có `recipe` riêng trong `api_transition/recipes/`.
- Chỉ khi đã tải thành công bằng API mới coi là `implemented`.

## Đã chuyển đổi thành công

| Hàm cũ | Downloader mới | Recipe | Trạng thái |
|---|---|---|---|
| `download_report_c11` | `download_report_c11_api()` | `c11_q2_2026.json` | OK |
| `download_report_c12` | `download_report_c12_api()` | `c12_q2_2026.json` | OK |
| `download_report_c13` | `download_report_c13_api()` | `c13_q2_2026.json` | OK |
| `download_report_c14` | `download_report_c14_api()` | `c14_q2_2026.json` | OK |
| `download_report_c14_chitiet` | `download_report_c14_chitiet_api()` | `c14_chitiet_q2_2026.json` | OK |
| `download_report_c11_chitiet` | `download_report_c11_chitiet_api()` | `c11_chitiet_q2_2026.json` | OK |
| `download_report_c12_chitiet_SM1` | `download_report_c12_chitiet_sm1_api()` | `c12_chitiet_sm1_q2_2026.json` | OK |
| `download_report_c12_chitiet_SM2` | `download_report_c12_chitiet_sm2_api()` | `c12_chitiet_sm2_q2_2026.json` | OK |
| `download_report_I15` | `download_report_i15_api()` | `i15_q2_2026.json` | OK |
| `download_report_I15_k2` | `download_report_i15_k2_api()` | `i15_k2_q2_2026.json` | OK |
| `download_GHTT_report_HNI` | `download_ghtt_report_hni_api()` | `ghtt_hni_q2_2026.json` | OK |
| `download_GHTT_report_Son_Tay` | `download_ghtt_report_sontay_api()` | `ghtt_sontay_q2_2026.json` | OK |
| `download_GHTT_report_nvktdb` | `download_ghtt_report_nvktdb_api()` | `ghtt_nvktdb_q2_2026.json` | OK |
| `xac_minh_tam_dung_download` | `download_xac_minh_tam_dung_api()` | `xac_minh_tam_dung_q2_2026.json` | OK |
| `kq_tiep_thi_download` | `download_kq_tiep_thi_api()` | `kq_tiep_thi_q2_2026.json` | OK |
| `download_report_vattu_thuhoi` | `download_report_vattu_thuhoi_api()` | `vattu_thuhoi_q2_2026.json` | OK |

Ghi chú xác nhận:
- `I1.5` và `I1.5 K2` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `GHTT HNI`, `GHTT Sơn Tây` và `GHTT NVKT DB` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `xac_minh_tam_dung_download` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `kq_tiep_thi_download` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `download_report_vattu_thuhoi` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `download_cau_hinh_tu_dong_chi_tiet_api()` đã được tải thành công bằng downloader API mới trên môi trường thực.

## Báo cáo bổ sung ngoài downloader cũ

| Báo cáo | Downloader mới | Recipe | File đầu ra | Trạng thái |
|---|---|---|---|
| `cấu hình tự động PTM` | `download_cau_hinh_tu_dong_ptm_api()` | `cau_hinh_tu_dong_q2_2026.json` | `cau_hinh_tu_dong_ptm.xlsx` | Ready |
| `cấu hình tự động Thay Thế` | `download_cau_hinh_tu_dong_thay_the_api()` | `cau_hinh_tu_dong_q2_2026.json` | `cau_hinh_tu_dong_thay_the.xlsx` | Ready |
| `cấu hình tự động chi tiết` | `download_cau_hinh_tu_dong_chi_tiet_api()` | `cau_hinh_tu_dong_chi_tiet_q2_2026.json` | `cau_hinh_tu_dong_chi_tiet.xlsx` | OK |
| `quyết toán vật tư` | `download_quyet_toan_vattu_api()` | `quyet_toan_vattu_q2_2026.json` | `quyet_toan_vat_tu.xlsx` | Implemented trong code / có trong batch |
| `CTS SHC ngày` | `download_cts_gpon_quality_detail_api()` | Không dùng recipe | `cts_shc_ngay.xlsx` | Module độc lập / có trong batch với `T-1` |

## Đang lỗi hoặc tạm hoãn

| Hàm cũ | Tình trạng hiện tại | Ghi chú |
|---|---|---|
| `download_report_c11_chitiet_SM2` | Lỗi ở luồng cũ | Chưa capture được recipe ổn định, nên chưa có downloader API mới |
| `download_report_c15` | Tạm hoãn | Đã có `c15_q2_2026.json` nhưng luồng export hiện đang lỗi |
| `download_report_c15_chitiet` | Tạm hoãn | Luồng gần API sẵn có nhưng đang lỗi, để xử lý sau cùng C1.5 |

## Recipe hiện có

- `c11_q2_2026.json`
- `c11_chitiet_q2_2026.json`
- `c12_q2_2026.json`
- `c12_chitiet_sm1_q2_2026.json`
- `c12_chitiet_sm2_q2_2026.json`
- `c13_q2_2026.json`
- `c14_q2_2026.json`
- `c14_chitiet_q2_2026.json`
- `c15_q2_2026.json`
- `i15_q2_2026.json`
- `i15_k2_q2_2026.json`
- `ghtt_hni_q2_2026.json`
- `ghtt_sontay_q2_2026.json`
- `ghtt_nvktdb_q2_2026.json`
- `xac_minh_tam_dung_q2_2026.json`
- `kq_tiep_thi_q2_2026.json`
- `vattu_thuhoi_q2_2026.json`
- `quyet_toan_vattu_q2_2026.json`
- `cau_hinh_tu_dong_q2_2026.json`
- `cau_hinh_tu_dong_chi_tiet_q2_2026.json`

## Bước kế tiếp đang ưu tiên

Vòng chuyển đổi downloader trong `api_transition/` hiện đã hoàn tất cho toàn bộ các báo cáo capture và xác nhận được.

Các mục còn lại cần xử lý riêng:
- `download_report_c11_chitiet_SM2`
- `download_report_c15`
- `download_report_c15_chitiet`

## Lệnh chạy nhanh cho các downloader mới đã OK

```bash
python3 api_transition/export_from_recipe.py --c11 --month-id 98944548
python3 api_transition/export_from_recipe.py --c12 --month-id 98944548
python3 api_transition/export_from_recipe.py --c13 --month-id 98944548
python3 api_transition/export_from_recipe.py --c14 --month-id 98944548
python3 api_transition/export_from_recipe.py --c14-chi-tiet --month-id 98944548
python3 api_transition/export_from_recipe.py --c11-chi-tiet --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/export_from_recipe.py --c12-chi-tiet-sm1 --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/export_from_recipe.py --c12-chi-tiet-sm2 --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/export_from_recipe.py --i15 --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/export_from_recipe.py --i15-k2 --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/export_from_recipe.py --ghtt-hni --month-id 98944548
python3 api_transition/export_from_recipe.py --ghtt-sontay --month-id 98944548
python3 api_transition/export_from_recipe.py --ghtt-nvktdb --month-id 98944548
python3 api_transition/export_from_recipe.py --xac-minh-tam-dung --start-date "01/04/2026" --end-date "16/04/2026"
python3 api_transition/export_from_recipe.py --kq-tiep-thi --start-date "16/04/2026" --end-date "16/04/2026"
python3 api_transition/export_from_recipe.py --vattu-thuhoi --start-date "24/11/2025" --end-date "16/04/2026"
python3 api_transition/export_from_recipe.py --cau-hinh-tu-dong-ptm --month-id 98944548
python3 api_transition/export_from_recipe.py --cau-hinh-tu-dong-thay-the --month-id 98944548
python3 api_transition/export_from_recipe.py --cau-hinh-tu-dong-chi-tiet --month-id 98944548
```

## Ghi chú vận hành

- `downloaders.py` hiện dùng helper chung `download_with_recipe()` để thống nhất luồng: load recipe, resolve tháng, merge override, export, lưu file.
- Mỗi downloader API mới đều nhận `session`; khi truyền `session`, hàm sẽ dùng lại Authorization/cookie đã capture thay vì login lại.
- `batch_download.py` là runner orchestration cho giai đoạn chuyển đổi hiện tại: `create_session()` login 1 lần, `_build_kwargs()` ánh xạ tham số theo `params_type`, `run_batch_download()` chạy tuần tự và retry timeout.
- `CTS SHC ngày` là ngoại lệ có chủ đích: nó nằm trong `api_transition/cts_api.py`, dùng login từ `cts.py`, gọi endpoint binary của `cts.vnpt.vn`, và khi chạy trong batch sẽ nhận `report_date = T-1`.
- CLI batch hỗ trợ `--only`, `--skip`, `--list`; hiện đây là cách đầy đủ nhất để chạy toàn bộ tập downloader đã được nối dây trong code.
- Mặc định tất cả downloader mới lưu file trong `api_transition/downloads/` và tự tách theo nhóm nghiệp vụ; chỉ khi truyền `--output-dir` mới ghi đè vị trí lưu.
- Nhóm `chi_tieu_c` chứa C1.1, C1.2, C1.3, C1.4 và các báo cáo chi tiết C.
- Nhóm `chi_tieu_i` chứa I1.5 và I1.5 K2.
- Nhóm `cts` chứa downloader CTS độc lập, hiện có `CTS SHC ngày`.
- Nhóm `cau_hinh_tu_dong` chứa các báo cáo cấu hình tự động tổng hợp và chi tiết.
- Nhóm `ghtt`, `xac_minh_tam_dung`, `kq_tiep_thi`, `vat_tu_thu_hoi` lưu tương ứng theo tên nghiệp vụ.
- Nhóm C1.1, C1.2, C1.3 dùng pattern tháng kiểu `pthang`.
- Nhóm C1.4 dùng pattern tháng kiểu `vthoigian`.
- Nhóm chi tiết C1.1/C1.2 hiện dùng cặp ngày `vngay_bd` và `vngay_kt`, không dùng `month-id`.
- Nhóm I1.5 và I1.5 K2 cũng dùng cặp ngày `vngay_bd` và `vngay_kt`; riêng I1.5 có thêm `vdk=0`.
- `CTS SHC ngày` dùng tham số `report_date = T-1` trong batch; module CTS sẽ tự dựng `tuNgay`, `denNgay`, `BeginDate`, `endDate` từ ngày này khi gọi API binary.
- Nhóm GHTT dùng `vthoigian` và `vdonvi`; `vloai=1` cho HNI/Sơn Tây và `vloai=2` cho NVKT DB.
- Báo cáo xác minh tạm dừng dùng `pdonvi_id`, `vngay_bd`, `vngay_kt`, `vloaidv=8,9`, `vloaingay=0`, `vloaibc=0`.
- Báo cáo kết quả tiếp thị dùng `vngay_bd`, `vngay_kt`, `vdonvi_id`.
- Báo cáo vật tư thu hồi dùng `vttvt`, `vtungay`, `vdenngay`, `vdichvuvt_erp`, `vvattu`, cùng các cờ `vloaithu=0`, `vloaibatbuoc=0`, `vloaingay=1`, `vtrangthai=0`.
- Báo cáo quyết toán vật tư dùng `vttvt`, `vtungay`, `vdenngay`, lưu `quyet_toan_vat_tu.xlsx`; hiện đã có recipe và đã được gắn vào batch.
- Báo cáo cấu hình tự động dùng `pthang` và `pdv`; đã tách thành 2 downloader rõ ràng:
- `download_cau_hinh_tu_dong_ptm_api()`: `pdv=1`, file `cau_hinh_tu_dong_ptm.xlsx`
- `download_cau_hinh_tu_dong_thay_the_api()`: `pdv=13`, file `cau_hinh_tu_dong_thay_the.xlsx`
- Báo cáo cấu hình tự động chi tiết dùng `pthang`, file `cau_hinh_tu_dong_chi_tiet.xlsx`.
- Mọi thay đổi mới chỉ nằm trong `api_transition/`, chưa thay thế downloader cũ ở code chính.
