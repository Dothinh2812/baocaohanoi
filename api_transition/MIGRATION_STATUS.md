# Migration Status

Tài liệu này ghi lại trạng thái chuyển đổi từ downloader cũ dựa trên click UI sang downloader mới trong `api_transition/`.

Nguyên tắc áp dụng:
- Không sửa các downloader đang chạy ở thư mục gốc.
- Mỗi hàm chuyển đổi thành công sẽ có `recipe` riêng trong `api_transition/recipes/`.
- Chỉ khi đã tải thành công bằng API mới coi là `implemented`.
- Ngoài downloader, `api_transition/` hiện đã có tầng processor, full pipeline và import SQLite để vận hành end-to-end.
- Luồng vận hành khuyến nghị hiện tại là `multi-instance`: mọi bước nhận `--config` và ghi dữ liệu vào `runtime/<unit>/...`.

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
- Processor `I1.5` / `I1.5 K2` đã được port sang `api_transition` và ghi lịch sử vào `report_history.db` của từng instance qua bảng `i15_*`.
- `GHTT HNI`, `GHTT Sơn Tây` và `GHTT NVKT DB` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `xac_minh_tam_dung_download` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `kq_tiep_thi_download` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `download_report_vattu_thuhoi` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `download_cau_hinh_tu_dong_chi_tiet_api()` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `download_report_c15_api()` đã được tải thành công bằng downloader API mới trên môi trường thực.
- `download_report_c15_chitiet_api()` đã được tải thành công bằng downloader API mới trên môi trường thực ngày `2026-04-21`, nhưng workbook trả về cho `pthoigianid=98944548` hiện đang rỗng.

## Báo cáo bổ sung ngoài downloader cũ

| Báo cáo | Downloader mới | Recipe | File đầu ra | Trạng thái |
|---|---|---|---|
| `cấu hình tự động PTM` | `download_cau_hinh_tu_dong_ptm_api()` | `cau_hinh_tu_dong_q2_2026.json` | `cau_hinh_tu_dong_ptm.xlsx` | Ready |
| `cấu hình tự động Thay Thế` | `download_cau_hinh_tu_dong_thay_the_api()` | `cau_hinh_tu_dong_q2_2026.json` | `cau_hinh_tu_dong_thay_the.xlsx` | Ready |
| `cấu hình tự động chi tiết` | `download_cau_hinh_tu_dong_chi_tiet_api()` | `cau_hinh_tu_dong_chi_tiet_q2_2026.json` | `cau_hinh_tu_dong_chi_tiet.xlsx` | OK |
| `tỷ lệ thiết lập DV BRCĐ` | `download_report_c15_api()` | `c15_q2_2026.json` | `c1.5 report.xlsx` | OK |
| `tỷ lệ thiết lập DV BRCĐ SM2 chi tiết` | `download_report_c15_chitiet_api()` | `c15_chitiet_q2_2026.json` | `c1.5_chitiet_report.xlsx` | OK |
| `CTS SHC ngày` | `download_cts_gpon_quality_detail_api()` | Không dùng recipe | `cts_shc_ngay.xlsx` | Module độc lập / có trong batch với `T-1` |

## Đang lỗi hoặc tạm hoãn

| Hàm cũ | Tình trạng hiện tại | Ghi chú |
|---|---|---|
| `download_report_c11_chitiet_SM2` | Lỗi ở luồng cũ | Chưa capture được recipe ổn định, nên chưa có downloader API mới |

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
- `c15_chitiet_q2_2026.json`
- `i15_q2_2026.json`
- `i15_k2_q2_2026.json`
- `ghtt_hni_q2_2026.json`
- `ghtt_sontay_q2_2026.json`
- `ghtt_nvktdb_q2_2026.json`
- `xac_minh_tam_dung_q2_2026.json`
- `kq_tiep_thi_q2_2026.json`
- `vattu_thuhoi_q2_2026.json`
- `cau_hinh_tu_dong_q2_2026.json`
- `cau_hinh_tu_dong_chi_tiet_q2_2026.json`

## Bước kế tiếp đang ưu tiên

Vòng chuyển đổi downloader trong `api_transition/` hiện đã hoàn tất cho toàn bộ các báo cáo capture và xác nhận được.

Các mục còn lại cần xử lý riêng:
- `download_report_c11_chitiet_SM2`

## Trạng thái Processor và Orchestrator

### Processor runner

- Đã có runner tổng `api_transition/processors/runner.py`
- Entry point Python: `from api_transition.processors import run_all_processors`
- CLI:

```bash
python3 -m api_transition.processors.runner
python3 -m api_transition.processors.runner --config api_transition/configs/units/son_tay.yaml
python3 -m api_transition.processors.runner --overwrite-processed
python3 -m api_transition.processors.runner --group tam_dung_khoi_phuc_dich_vu
python3 -m api_transition.processors.runner --list
```

- Runner hiện quản lý 33 processor đã được port vào `api_transition/processors/`
- Có hỗ trợ `--only`, `--skip`, `--group`, `--stop-on-error`
- Khi truyền `--config`, runner sẽ tự đọc raw từ `runtime/<unit>/downloads` và ghi processed vào `runtime/<unit>/Processed`
- `c15_chitiet` đã được thêm vào runner; trên `son_tay` ngày `2026-04-21` processor chạy thành công và sinh workbook processed chuẩn dù raw workbook hiện rỗng.
- `xac_minh_tam_dung` đã được thêm processor, schema SQLite, importer và view dashboard riêng; workbook processed đã kiểm tra thành công trên `son_tay`.

### Full pipeline

- Đã có entrypoint `api_transition/full_pipeline.py`
- Entry point Python: `from api_transition import run_full_pipeline`
- CLI:

```bash
python3 -m api_transition.full_pipeline
python3 -u -m api_transition.full_pipeline --config api_transition/configs/units/son_tay.yaml --reset-db
python3 -m api_transition.full_pipeline --snapshot-date 2026-04-19
python3 api_transition/full_pipeline.py --snapshot-date 2026-04-19
```

- Luồng hiện tại:
  - download tất cả qua `run_batch_download()`
  - xử lý tất cả qua `run_all_processors()`
  - archive workbook processed thành công sang `ProcessedDaily/<snapshot-date>/...`
  - import vào SQLite của đúng instance
  - apply lại views sau import
- Mặc định pipeline vẫn import phần đã download/process thành công dù một số bước khác lỗi
- Dùng `--strict` nếu muốn dừng trước bước import khi download/process có lỗi
- Khi truyền `--config`, pipeline sẽ tự chạy trên:
  - `runtime/<unit>/downloads`
  - `runtime/<unit>/Processed`
  - `runtime/<unit>/ProcessedDaily`
  - `runtime/<unit>/sqlite_history/report_history.db`
- Đã verify end-to-end thật cho `son_tay` ngày `2026-04-20`:
  - `28` download thành công, `0` lỗi, `5` skip theo config
  - `25` processor thành công, `0` lỗi, `5` skip theo config
  - `28` workbook archive
  - `28` workbook import SQLite, `0` lỗi import

### SQLite history

- Đã có SQLite local cho cả `standalone` và `multi-instance`
- Script import độc lập: `api_transition/sqlite_history/import_processed_to_sqlite.py`
- Admin utility đồng bộ DB instance: `api_transition/sqlite_history/sync_all_instance_dbs.py`
- Nếu chạy qua `full_pipeline.py`, importer chỉ nạp các workbook processed thành công của chính lượt chạy đó
- Nếu chạy import CLI độc lập, script sẽ quét toàn bộ `Processed` của root được truyền vào
- `ProcessedDaily` hiện được tạo ngay sau bước process trong full pipeline, không còn phụ thuộc việc import SQLite có thành công hay không
- `sync_all_instance_dbs.py` nằm ngoài pipeline hằng ngày; dùng để `status`, `apply-views`, `init-if-missing`, `reset-and-init` cho toàn bộ DB instance
- Với `C1.5 chi tiết`, đã thêm schema/import/view SQLite mới và đã verify importer trên fixture có dữ liệu; DB `son_tay` thật ngày `2026-04-21` đã có snapshot import thành công nhưng business rows bằng `0` do workbook raw rỗng.

## Trạng thái Processor nhóm dịch vụ/MyTV

### Đã vận hành được

- `xac_minh_tam_dung`
- `phieu_hoan_cong_dich_vu`
- `tam_dung_khoi_phuc_chi_tiet`
- `tam_dung_khoi_phuc_tong_hop`
- `fiber_thuc_tang`
- `mytv_ngung_psc`
- `mytv_hoan_cong`
- `mytv_thuc_tang`
- `son_tay_mytv_ngung_psc_t_minus_1`
- `son_tay_fiber_ngung_psc_t_minus_1`

### Ghi chú migration MyTV

- `xac_minh_tam_dung` đã có processor riêng, output canonical là `Processed/xac_minh_tam_dung/xac_minh_tam_dung report_processed.xlsx`
- `mytv_ngung_psc` đã bỏ phụ thuộc vào file legacy `mytv_ngung_psc.xlsx`
- Raw path mới:
  - `downloads/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_cap_to.xlsx`
  - `downloads/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_cap_ttvt.xlsx`
- `mytv_hoan_cong` đã bỏ phụ thuộc vào file legacy `mytv_hoan_cong.xlsx`
- `mytv_hoan_cong` và `mytv_thuc_tang` hiện lấy dữ liệu MyTV trực tiếp từ `phieu_hoan_cong_dich_vu_chi_tiet.xlsx`
- `son_tay_mytv_ngung_psc_t_minus_1` hiện đọc từ `downloads/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_sontay.xlsx`
- Output canonical:
  - `Processed/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_cap_to_processed.xlsx`
  - `Processed/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_cap_ttvt_processed.xlsx`
  - `Processed/phieu_hoan_cong_dich_vu/phieu_hoan_cong_dich_vu_chi_tiet_processed.xlsx`
  - `Processed/tam_dung_khoi_phuc_dich_vu/mytv_thuc_tang_processed.xlsx`
- Hạn chế hiện tại: raw API ngưng PSC MyTV mới chưa có đủ chi tiết NVKT, nên `mytv_thuc_tang` hiện mới sinh được nhánh tổng hợp theo tổ/TTVT

## Lệnh chạy nhanh cho các downloader mới đã OK

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py --c11 --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c12 --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c13 --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c14 --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c14-chi-tiet --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c11-chi-tiet --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --c12-chi-tiet-sm1 --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --c12-chi-tiet-sm2 --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --i15 --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --i15-k2 --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --ghtt-hni --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --ghtt-sontay --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --ghtt-nvktdb --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --xac-minh-tam-dung --start-date "01/04/2026" --end-date "16/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --kq-tiep-thi --start-date "16/04/2026" --end-date "16/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --vattu-thuhoi --start-date "24/11/2025" --end-date "16/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --cau-hinh-tu-dong-ptm --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --cau-hinh-tu-dong-thay-the --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --cau-hinh-tu-dong-chi-tiet --month-id 98944548
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
- Báo cáo cấu hình tự động dùng `pthang` và `pdv`; đã tách thành 2 downloader rõ ràng:
- `download_cau_hinh_tu_dong_ptm_api()`: `pdv=1`, file `cau_hinh_tu_dong_ptm.xlsx`
- `download_cau_hinh_tu_dong_thay_the_api()`: `pdv=13`, file `cau_hinh_tu_dong_thay_the.xlsx`
- Báo cáo cấu hình tự động chi tiết dùng `pthang`, file `cau_hinh_tu_dong_chi_tiet.xlsx`.
- Mọi thay đổi mới chỉ nằm trong `api_transition/`, chưa thay thế downloader cũ ở code chính.
