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
- `chua_dung_den/`: công cụ capture/OneBSS/PoC không đi vào full pipeline
- `chua_dung_den/capture_report_api.py`: capture request/response và sinh recipe JSON
- `chua_dung_den/capture_with_legacy_flow.py`: chạy hàm download cũ ở chế độ headless và tự bắt recipe
- `downloaders.py`: các hàm downloader API mới
- `ONEBSS.md`: tài liệu chi tiết cho nhóm downloader OneBSS (`chua_dung_den/onebss_auth.py`, `chua_dung_den/onebss_report_client.py`, `chua_dung_den/onebss_downloaders.py`)
- `cts_api.py`: module API riêng cho CTS, dùng login trong `cts.py` và tải file binary trực tiếp
- `batch_download.py`: runner batch login 1 lần, tái sử dụng session và tải tuần tự các report đã wired
- `runtime_config.py`: nạp YAML config đơn vị và dựng `RuntimeContext` cho multi-instance
- `configs/units/`: bộ config YAML cho từng đơn vị và template chung
- `processors/`: các hàm xử lý workbook raw -> workbook processed
- `full_pipeline.py`: entrypoint chạy full pipeline từ download đến import SQLite
- `PROCESSING_REFACTOR_PLAN.md`: trạng thái refactor processor và orchestrator
- `chua_dung_den/export_from_recipe.py`: CLI generic chạy bằng recipe
- `recipes/`: recipe đã capture và xác nhận
- `downloads/`: thư mục tải file mặc định theo nhóm nghiệp vụ
- `Processed/`: workbook processed chuẩn hóa để import vào SQLite
- `ProcessedDaily/`: snapshot processed theo ngày dữ liệu
- `runtime/`: workspace vận hành multi-instance theo từng đơn vị
- `sqlite_history/sync_all_instance_dbs.py`: admin utility đồng bộ SQLite cho mọi instance
- `chua_dung_den/export_c11_api.py`: PoC export C1.1 qua API

## Cấu trúc thư mục tải file

Có 2 chế độ chạy:

- `standalone`: giữ hành vi cũ, mặc định lưu vào `api_transition/downloads/`
- `multi-instance`: khuyến nghị cho vận hành thật, mỗi đơn vị có workspace riêng dưới `api_transition/runtime/<unit>/`

Ở chế độ `standalone`, các downloader mới sẽ lưu vào `api_transition/downloads/` và tách theo nhóm:

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

Ở chế độ `multi-instance`, toàn bộ dữ liệu của 1 đơn vị sẽ nằm trong:

- `api_transition/runtime/<unit>/downloads/`
- `api_transition/runtime/<unit>/Processed/`
- `api_transition/runtime/<unit>/ProcessedDaily/`
- `api_transition/runtime/<unit>/sqlite_history/report_history.db`

`RuntimeContext` sẽ dựng các đường dẫn này từ file YAML trong `configs/units/`.

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

Lưu ý:
- hiện chưa có shortcut CLI riêng trong `export_from_recipe.py`, nên cách chạy trực tiếp thuận tiện nhất là import Python hoặc dùng batch
- `cts_api.py` là module độc lập, không dùng `report-api` của `baocao.hanoi` và không đi qua `download_with_recipe()`

## Yêu cầu

- Dùng chung `.env` hiện tại ở root repo
- Đã cài dependencies của project
- Đã cài Playwright Chromium

## Chạy thử C1.1

Linux/macOS:

```bash
python3 api_transition/chua_dung_den/export_c11_api.py --headed --month-id 98944548
```

Windows:

```powershell
python api_transition/chua_dung_den/export_c11_api.py --headed --month-id 98944548
```

Nếu muốn để script tự map kỳ báo cáo theo nhãn:

```bash
python3 api_transition/chua_dung_den/export_c11_api.py --headed --month-label "Tháng 04/2026"
```

Hoặc dùng downloader generic:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py --c11 --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c12 --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c13 --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c14 --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c14-chi-tiet --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --c11-chi-tiet --headed --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --c12-chi-tiet-sm1 --headed --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --c12-chi-tiet-sm2 --headed --start-date "26/03/2026" --end-date "25/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --i15 --headed --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --i15-k2 --headed --start-date "14/04/2026" --end-date "14/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --ghtt-hni --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --ghtt-sontay --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --ghtt-nvktdb --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --xac-minh-tam-dung --headed --start-date "01/04/2026" --end-date "16/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --kq-tiep-thi --headed --start-date "16/04/2026" --end-date "16/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --vattu-thuhoi --headed --start-date "24/11/2025" --end-date "16/04/2026"
python3 api_transition/chua_dung_den/export_from_recipe.py --cau-hinh-tu-dong-ptm --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --cau-hinh-tu-dong-thay-the --headed --month-id 98944548
python3 api_transition/chua_dung_den/export_from_recipe.py --cau-hinh-tu-dong-chi-tiet --headed --month-id 98944548
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
- hỗ trợ `--config` để resolve `unit_id`, bật/tắt report, và ghi file vào `runtime/<unit>/downloads`

Ví dụ:

```bash
python3 api_transition/batch_download.py
python3 -m api_transition.batch_download --config api_transition/configs/units/son_tay.yaml
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
- `C1.5`, `C1.5 Chi tiết`
- `C1.1 Chi tiết`, `C1.2 Chi tiết SM1`, `C1.2 Chi tiết SM2`
- `I1.5`, `I1.5 K2`
- `GHTT HNI`, `GHTT Sơn Tây`, `GHTT NVKT DB`
- `Xác minh tạm dừng`, `Kết quả tiếp thị`, `CTS SHC ngày`
- `Tạm dừng, khôi phục DV chi tiết`, `Tạm dừng, khôi phục DV chi tiết - khôi phục`, `Tạm dừng, khôi phục DV tổng hợp`
- `Vật tư thu hồi`
- `Cấu hình tự động PTM`, `Cấu hình tự động Thay thế`, `Cấu hình tự động Chi tiết`

`CTS SHC ngày` hiện được batch gọi với tham số ngày `T-1` dưới dạng `report_date`, tức mặc định lấy ngày hôm qua theo định dạng `dd/mm/yyyy`.

`I1.5` và `I1.5 K2` đã được xác nhận tải thành công bằng downloader API mới trên môi trường thực.
Bộ xử lý của `I1.5` / `I1.5 K2` cũng đã được port sang `api_transition.processors.i15_processors`, và lịch sử được ghi trực tiếp vào `report_history.db` của từng instance qua nhóm bảng `i15_*`.
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

## Processors

`api_transition/processors/` là tầng chuẩn hóa dữ liệu sau download. Ở chế độ `standalone`, processor đọc workbook raw trong `api_transition/downloads/` và ghi vào `api_transition/Processed/`. Ở chế độ `multi-instance`, runner sẽ tự map sang `runtime/<unit>/downloads/` và `runtime/<unit>/Processed/`.

Runner tổng hiện tại là `processors/runner.py`, quản lý 33 processor đã được port vào package.

Ví dụ:

```bash
python3 -m api_transition.processors.runner
python3 -m api_transition.processors.runner --config api_transition/configs/units/son_tay.yaml
python3 -m api_transition.processors.runner --overwrite-processed
python3 -m api_transition.processors.runner --group tam_dung_khoi_phuc_dich_vu
python3 -m api_transition.processors.runner --only mytv_ngung_psc --only mytv_hoan_cong
python3 -m api_transition.processors.runner --list
```

Import từ Python:

```python
from api_transition.processors import run_all_processors

results = run_all_processors(
    overwrite_processed=True,
    groups=["tam_dung_khoi_phuc_dich_vu"],
)
```

Một số lưu ý hiện tại:

- `c15_chitiet` đã có processor riêng, tạo các sheet `KQ_C15_chitiet`, `TH_TTVTST`, `Chi_tiet_TG`, `TH_KIEULD`, `TH_DVVT`, `TH_DVVT_DOI`, `TH_DVVT_TTVT`
- downloader `C1.5 Chi tiết` đã verify tải thật ngày `2026-04-21`, nhưng raw workbook trả về cho `Tháng 04/2026` đang rỗng; processor vẫn sinh được workbook processed chuẩn với các sheet canonical rỗng để pipeline không vỡ
- nhóm MyTV đã được port sang raw schema mới thay cho file legacy
- `xac_minh_tam_dung` đã có processor riêng, tạo workbook processed chuẩn từ `downloads/xac_minh_tam_dung/xac_minh_tam_dung report.xlsx`
- `mytv_ngung_psc` đọc từ `downloads/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_cap_to.xlsx` và `downloads/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_cap_ttvt.xlsx`
- `mytv_hoan_cong` và `mytv_thuc_tang` hiện lấy dữ liệu MyTV trực tiếp từ `phieu_hoan_cong_dich_vu_chi_tiet.xlsx`
- `mytv_thuc_tang` hiện mới sinh được sheet theo tổ/TTVT; raw API hiện chưa có đủ chi tiết NVKT cho nhánh ngưng PSC MyTV

## Full Pipeline

`full_pipeline.py` là entrypoint orchestration cho toàn bộ luồng:

1. login và download tất cả report đã wired
2. chạy toàn bộ processor đã có trong `api_transition/processors`
3. copy các workbook processed thành công sang `ProcessedDaily/<snapshot-date>/...`
4. khởi tạo hoặc tái sử dụng `report_history.db`
5. import vào SQLite chỉ các workbook processed thành công của chính lượt chạy đó
6. apply lại views sau import

Ví dụ:

```bash
python3 -m api_transition.full_pipeline
python3 -u -m api_transition.full_pipeline --config api_transition/configs/units/son_tay.yaml --reset-db
python3 -m api_transition.full_pipeline --snapshot-date 2026-04-19
python3 -m api_transition.full_pipeline --snapshot-date 2026-04-19 --overwrite-processed
python3 -m api_transition.full_pipeline --snapshot-date 2026-04-19 --strict
python3 api_transition/full_pipeline.py --snapshot-date 2026-04-19
```

Import từ Python:

```python
from api_transition import run_full_pipeline

result = run_full_pipeline(overwrite_processed=True)
```

Hành vi chính:

- mặc định pipeline vẫn archive và import phần report/process thành công dù có một số bước khác lỗi
- dùng `--strict` nếu muốn dừng trước bước import khi download/process có lỗi
- `ProcessedDaily` hiện được tạo ngay sau bước process, không còn phụ thuộc vào việc import SQLite thành công hay không
- bước import trong pipeline chỉ nạp các workbook processed thành công của lượt chạy hiện tại, tránh quét lại toàn bộ `Processed/`
- khi truyền `--config`, pipeline sẽ tự dùng `RuntimeContext` để chạy toàn bộ vào `runtime/<unit>/...`

Xác nhận vận hành:

- ngày `2026-04-20`, full pipeline theo config `son_tay` đã được verify end-to-end với kết quả `28` download thành công, `25` processor thành công, `28` workbook archive, `28` workbook import SQLite, `0` lỗi import

## SQLite history

Đã bổ sung khung SQLite local để lưu lịch sử workbook processed theo từng ngày.

- schema SQL: `api_transition/sqlite_history/report_history_schema.sql`
- view SQL: `api_transition/sqlite_history/report_history_views.sql`
- script khởi tạo DB: `api_transition/sqlite_history/init_report_history_db.py`
- script import DB: `api_transition/sqlite_history/import_processed_to_sqlite.py`
- script apply view: `api_transition/sqlite_history/apply_report_history_views.py`
- script đồng bộ toàn bộ DB instance: `api_transition/sqlite_history/sync_all_instance_dbs.py`
- tài liệu bảo trì module: `api_transition/sqlite_history/README.md`
- tài liệu cho app đọc dữ liệu: `api_transition/sqlite_history/CONSUMER_GUIDE.md`
- file DB mặc định standalone: `api_transition/report_history.db`
- file DB khuyến nghị khi chạy multi-instance: `api_transition/runtime/<unit>/sqlite_history/report_history.db`

Nguyên tắc hiện tại:

- mỗi báo cáo chỉ có 1 bản hợp lệ cho mỗi cặp `ma_bao_cao + ngay_du_lieu`
- chạy lại trong cùng ngày sẽ ghi đè dữ liệu ngày đó trong SQLite
- sang ngày mới sẽ tạo snapshot mới
- nếu chạy qua `full_pipeline.py`, chỉ các workbook processed thành công của lượt chạy đó mới được import
- nếu chạy import CLI độc lập, script sẽ quét toàn bộ `api_transition/Processed`
- tên bảng và tên cột dùng tiếng Việt không dấu, bám theo ý nghĩa báo cáo
- khi báo cáo phát sinh cột mới, dữ liệu chưa map chính thức sẽ đi vào `du_lieu_bo_sung_json`

Khởi tạo DB:

```bash
python3 api_transition/sqlite_history/init_report_history_db.py --reset
```

Tạo DB ở đường dẫn khác để thử:

```bash
python3 api_transition/sqlite_history/init_report_history_db.py \
  --db-path /tmp/report_history_test.db \
  --reset
```

Import dữ liệu từ `api_transition/Processed` vào SQLite:

```bash
python3 api_transition/sqlite_history/import_processed_to_sqlite.py \
  --snapshot-date 2026-04-18 \
  --skip-archive
```

Một số lệnh hay dùng:

```bash
python3 api_transition/sqlite_history/import_processed_to_sqlite.py --snapshot-date 2026-04-19
python3 api_transition/sqlite_history/import_processed_to_sqlite.py --snapshot-date 2026-04-19 --dry-run --json
python3 api_transition/sqlite_history/import_processed_to_sqlite.py --snapshot-date 2026-04-19 --path-contains "tam_dung_khoi_phuc_dich_vu"
python3 api_transition/sqlite_history/import_processed_to_sqlite.py --db-path /tmp/report_history_test.db --snapshot-date 2026-04-19 --skip-archive
```

Lưu ý:

- khi chạy script import trực tiếp và không dùng `--skip-archive`, workbook sẽ được copy sang `ProcessedDaily/<snapshot-date>/...`
- khi chạy qua `full_pipeline.py`, archive này thường đã được tạo sẵn ngay sau bước process và importer sẽ tái sử dụng lại bản archive đó
- khi vận hành multi-instance, nên để `full_pipeline.py --config ...` điều phối `db_path`, `processed_root`, `archive_root`
- `sync_all_instance_dbs.py` là admin utility nằm ngoài pipeline hằng ngày; dùng khi cần tạo DB mới, apply lại views, hoặc reset/re-init hàng loạt

Ví dụ admin utility:

```bash
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode status
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode apply-views
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode init-if-missing
python3 -m api_transition.sqlite_history.sync_all_instance_dbs --mode reset-and-init
```

Các bảng nền:

- `danh_muc_bao_cao`
- `bao_cao_ngay`
- `sheet_bao_cao`
- `dong_bao_cao_goc`
- `tep_luu_tru_bao_cao`
- `nhat_ky_nap_bao_cao`
- `danh_muc_don_vi`
- `danh_muc_nhan_vien`

Các bảng nghiệp vụ chính:

- `c11_tong_hop`, `c11_chi_tiet_nvkt`
- `c12_tong_hop`, `c12_hong_lap_lai_nvkt`
- `c13_tong_hop`
- `c14_tong_hop`, `c14_hai_long_nvkt`
- `ghtt_don_vi`, `ghtt_nvkt`
- `kpi_nvkt_c11`, `kpi_nvkt_c12`, `kpi_nvkt_c13`
- `ket_qua_tiep_thi_nv`, `ket_qua_tiep_thi_don_vi`
- `hoan_cong_fiber`, `ngung_psc_fiber`, `khoi_phuc_fiber`
- `hoan_cong_mytv`, `ngung_psc_mytv`
- `thuc_tang_fiber`, `thuc_tang_mytv`
- `xac_minh_chi_tiet`, `xac_minh_tong_hop_nvkt`, `xac_minh_tong_hop_loai_phieu`
- `xac_minh_tam_dung_chi_tiet`, `xac_minh_tam_dung_tong_hop_nvkt`, `xac_minh_tam_dung_tong_hop_dich_vu`, `xac_minh_tam_dung_tong_hop_ly_do_huy`
- `cau_hinh_tu_dong_chi_tiet`, `cau_hinh_tu_dong_tong_hop`, `tong_hop_loi_cau_hinh_tu_dong`
- `vat_tu_thu_hoi`, `chi_tiet_vat_tu_thu_hoi`

## Capture report mới và sinh recipe

Ví dụ capture C1.2:

```bash
python3 api_transition/chua_dung_den/capture_report_api.py \
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
python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name c12_q2_2026 \
  --legacy-func download_report_c12 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=535021" \
  --report-month "Tháng 04/2026"
```

Ví dụ C1.1:

```bash
python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name c11_q2_2026_auto \
  --legacy-func download_report_c11 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=534964&menu_id=535020" \
  --report-month "Tháng 04/2026"
```

Nhóm tiếp theo nên capture sau khi tạm bỏ qua C1.5 là 4 báo cáo chi tiết C1.1/C1.2:

```bash
python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name c11_chitiet_q2_2026 \
  --legacy-func download_report_c11_chitiet \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"

python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name c11_chitiet_sm2_q2_2026 \
  --legacy-func download_report_c11_chitiet_SM2 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"

python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name c12_chitiet_sm1_q2_2026 \
  --legacy-func download_report_c12_chitiet_SM1 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"

python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name c12_chitiet_sm2_q2_2026 \
  --legacy-func download_report_c12_chitiet_SM2 \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194" \
  --start-date "26/03/2026" \
  --end-date "25/04/2026"
```

Sau đó export lại bằng recipe:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py \
  --headed \
  --recipe c12_q2_2026 \
  --month-id <pthang>
```

## Ghi đè tham số input

Để capture `xac_minh_tam_dung_download`:

```bash
python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name xac_minh_tam_dung_q2_2026 \
  --legacy-func xac_minh_tam_dung_download \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=267844&menu_id=276199"
```

Nhóm tiếp theo nên capture là `kq_tiep_thi_download`:

```bash
python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name kq_tiep_thi_q2_2026 \
  --legacy-func kq_tiep_thi_download \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=257495&menu_id=276101"
```

Sau khi capture xong, chạy bản API mới:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py \
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
python3 api_transition/chua_dung_den/capture_with_legacy_flow.py \
  --name vattu_thuhoi_q2_2026 \
  --legacy-func download_report_vattu_thuhoi \
  --report-url "https://baocao.hanoi.vnpt.vn/report/report-info?id=270922&menu_id=276242"
```

Sau khi capture xong, chạy bản API mới:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py \
  --vattu-thuhoi \
  --start-date "24/11/2025" \
  --end-date "16/04/2026"
```

Đối với báo cáo `cấu hình tự động`, dùng 2 downloader riêng:

PTM:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py \
  --cau-hinh-tu-dong-ptm \
  --month-id 98944548
```

Thay Thế:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py \
  --cau-hinh-tu-dong-thay-the \
  --month-id 98944548
```

Ánh xạ rõ:
- `--cau-hinh-tu-dong-ptm` gọi `download_cau_hinh_tu_dong_ptm_api()`, ép `pdv=1`, lưu `cau_hinh_tu_dong_ptm.xlsx`
- `--cau-hinh-tu-dong-thay-the` gọi `download_cau_hinh_tu_dong_thay_the_api()`, ép `pdv=13`, lưu `cau_hinh_tu_dong_thay_the.xlsx`

Đối với báo cáo `cấu hình tự động chi tiết`:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py \
  --cau-hinh-tu-dong-chi-tiet \
  --month-id 98944548
```

Ánh xạ rõ:
- `--cau-hinh-tu-dong-chi-tiet` gọi `download_cau_hinh_tu_dong_chi_tiet_api()`, dùng `pthang`, lưu `cau_hinh_tu_dong_chi_tiet.xlsx`

Có thể thay tham số trong payload đã capture:

```bash
python3 api_transition/chua_dung_den/export_from_recipe.py \
  --headed \
  --recipe c11_q2_2026 \
  --set ptrungtamid=14324 \
  --month-id 98944548
```

## Mục tiêu chuyển đổi dần

1. Xác nhận từng báo cáo tải được qua API.
2. So sánh file đầu ra với luồng cũ.
3. Khi ổn định mới thay thế từng downloader ở code chính.
