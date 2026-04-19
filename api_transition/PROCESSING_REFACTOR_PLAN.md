# Kế hoạch refactor luồng xử lý file tải về trong `api_transition`

## 1. Mục tiêu

- Tách toàn bộ phần xử lý file thô sau download sang namespace `api_transition`.
- Không đụng vào codebase gốc đang chạy.
- Chuẩn hóa cách gọi processor:
  - nhận `input_path` tường minh
  - không hard-code thư mục cũ như `downloads/baocao_hanoi`, `KPI-DOWNLOAD`, `KQ-TIEP-THI`, `PTTB-PSC`
- không ghi đè file raw đã tải
- mọi kết quả xử lý phải đi vào cây thư mục riêng `Processed/`
- Giữ nguyên logic nghiệp vụ cũ khi có thể.
- Với báo cáo mới chưa có hàm cũ tương ứng:
  - nếu dữ liệu ở mức tổ: tạo sheet tổng hợp theo tổ
  - nếu dữ liệu ở mức cá nhân/NVKT: tạo sheet theo cá nhân, theo tổ, và nếu cần file riêng cho từng NVKT

## 1.1. Nguyên tắc lưu file sau xử lý

- Tạo thư mục gốc mới riêng khỏi `downloads`, tên là:
  - `api_transition/Processed/`
- Bên trong `Processed`, giữ nguyên cấu trúc nhóm giống `downloads`:
  - ví dụ:
    - `api_transition/downloads/chi_tieu_c/...`
    - `api_transition/Processed/chi_tieu_c/...`
    - `api_transition/downloads/kpi_nvkt/...`
    - `api_transition/Processed/kpi_nvkt/...`
- Khi xử lý một file raw:
  - copy file gốc sang thư mục `Processed/<group>/`
  - đổi tên file đích bằng cách thêm hậu tố `_processed` trước `.xlsx`
- Tất cả sheet kết quả đều được ghi vào file copy này.
- File raw trong `downloads/` chỉ dùng làm input, không bị chỉnh sửa.

## 1.2. Nguyên tắc mở rộng cho nhiều đơn vị

- Processor phải được viết theo hướng dùng lại cho nhiều đơn vị khác nhau, không chỉ Sơn Tây.
- Không fix cứng:
  - tên tổ
  - tên trung tâm / TTVT
  - tên cá nhân / NVKT
  - danh sách nhân sự đặc thù của một đơn vị
- Mọi grouping và tổng hợp phải ưu tiên trích xuất động từ dữ liệu báo cáo tải về.
- Chỉ dùng file tham chiếu ngoài như `dsnv.xlsx` hoặc `danhba.db` cho:
  - enrich dữ liệu
  - chuẩn hóa mapping
  - bổ sung metadata
- Không dùng dữ liệu tham chiếu để hard-code các tập tên cố định nếu bản thân report đã chứa thông tin đủ để suy luận.

## 2. Các pattern xử lý đã đọc được

### 2.1. Pattern A: Thêm sheet tổng hợp vào chính file raw

Áp dụng trong:
- `c1_process.py`
- `kq_tiep_thi_process.py`
- một phần `i15_process.py`

Đặc điểm:
- đọc file gốc
- lọc/xóa cột/xóa dòng header thừa
- đổi tên cột
- tạo 1 hoặc nhiều sheet kết quả
- trong `api_transition` sẽ đổi thành:
  - copy workbook nguồn sang `Processed/.../*_processed.xlsx`
  - ghi sheet vào workbook copy bằng `ExcelWriter(..., mode='a', if_sheet_exists='replace')`

### 2.2. Pattern B: Chuẩn hóa tên NVKT rồi group theo cá nhân/tổ

Áp dụng trong:
- `kpi_process_from_download_baocaohanoi.py`
- `thuc_tang_process.py`
- `olds_files/KR_process.py`
- `vat_tu_thu_hoi_process.py`
- nhiều hàm chi tiết trong `c1_process.py`

Đặc điểm:
- bóc tên NVKT từ chuỗi dạng `Mã-Tên(...)` hoặc `Khu vực - Tên(...)`
- lookup `đơn vị` từ `dsnv.xlsx`
- tạo báo cáo theo:
  - cá nhân
  - tổ/đội
- đôi khi thêm sheet riêng từng đơn vị
- trong bản refactor mới:
  - việc xác định “tổ”, “trung tâm”, “cá nhân” phải dựa trên cột có trong raw report
  - chỉ fallback sang nguồn ngoài nếu raw thiếu thông tin

### 2.3. Pattern C: Tracking lịch sử theo ngày

Áp dụng trong:
- `i15_process.py`

Đặc điểm:
- cần `danhba.db`
- cần DB history riêng (`suy_hao_history.db`, `suy_hao_history_k2.db`)
- so sánh với ngày trước để sinh:
  - `Tang_moi`
  - `Giam_het`
  - `Van_con`
  - `Bien_dong_tong_hop`

### 2.4. Pattern D: Tạo báo cáo dẫn xuất từ 2 nguồn raw

Áp dụng trong:
- `thuc_tang_process.py`

Đặc điểm:
- đọc 2 file trung gian đã xử lý:
  - ngưng PSC
  - hoàn công
- từ đó tạo báo cáo thực tăng theo tổ và theo NVKT

Lưu ý thêm:
- `thuc_tang_process.py` thực tế đang bao phủ 8 luồng xử lý:
  - `process_ngung_psc_report`
  - `process_hoan_cong_report`
  - `create_thuc_tang_report`
  - `process_mytv_ngung_psc_report`
  - `process_mytv_hoan_cong_report`
  - `create_mytv_thuc_tang_report`
  - `process_son_tay_ngung_psc_report`
  - `process_son_tay_mytv_ngung_psc_report`
- Không phải tất cả các raw tương ứng đã có luồng download API trong `api_transition`.
- Vì vậy phần processor của nhóm này phải được xem là phụ thuộc trực tiếp vào việc có đủ file raw mới trong `api_transition/downloads`.

## 3. Mapping file cũ -> file mới trong `api_transition`

## 3.1. Nhóm đã có mapping rất rõ

| Nghiệp vụ | Hàm cũ | File cũ | File mới trong `api_transition` | Ghi chú |
|---|---|---|---|---|
| C1.1 | `process_c11_report` | `downloads/baocao_hanoi/c1.1 report.xlsx` | `api_transition/downloads/chi_tieu_c/c1.1 report.xlsx` | Port gần như trực tiếp |
| C1.2 | `process_c12_report` | `downloads/baocao_hanoi/c1.2 report.xlsx` | `api_transition/downloads/chi_tieu_c/c1.2 report.xlsx` | Port trực tiếp |
| C1.3 | `process_c13_report` | `downloads/baocao_hanoi/c1.3 report.xlsx` | `api_transition/downloads/chi_tieu_c/c1.3 report.xlsx` | Port trực tiếp |
| C1.4 | `process_c14_report` | `downloads/baocao_hanoi/c1.4 report.xlsx` | `api_transition/downloads/chi_tieu_c/c1.4 report.xlsx` | Port trực tiếp |
| C1.4 chi tiết | `process_c14_chitiet_report` | `downloads/baocao_hanoi/c1.4_chitiet_report.xlsx` | `api_transition/downloads/chi_tieu_c/c1.4_chitiet_report.xlsx` | Port trực tiếp |
| I1.5 | `process_I15_report_with_tracking` | `downloads/baocao_hanoi/I1.5 report.xlsx` | `api_transition/downloads/chi_tieu_i/i1.5 report.xlsx` | Khác chữ hoa/thường, phải xử lý path cẩn thận |
| I1.5 K2 | `process_I15_k2_report_with_tracking` | `downloads/baocao_hanoi/I1.5_k2 report.xlsx` | `api_transition/downloads/chi_tieu_i/i1.5_k2 report.xlsx` | Tương tự |
| KPI NVKT C11 | `c11_process_report_nvkt` | `KPI-DOWNLOAD/c11-nvktdb report.xlsx` | `api_transition/downloads/kpi_nvkt/c11-nvktdb report.xlsx` | Port trực tiếp |
| KPI NVKT C12 | `c12_process_report_nvkt` | `KPI-DOWNLOAD/c12-nvktdb report.xlsx` | `api_transition/downloads/kpi_nvkt/c12-nvktdb report.xlsx` | Port trực tiếp |
| KPI NVKT C13 | chưa có hàm cũ riêng | chưa có | `api_transition/downloads/kpi_nvkt/c13-nvktdb report.xlsx` | Dùng cùng pattern C11/C12 |
| KQ tiếp thị | `process_kq_tiep_thi_report` | `KQ-TIEP-THI/kq_tiep_thi*.xlsx` | `api_transition/downloads/kq_tiep_thi/kq_tiep_thi report.xlsx` | Cần bỏ logic tìm file theo ngày, dùng path tường minh |
| GHTT HNI | chưa có hàm cũ riêng 1-1 | `GHTT/tong_hop_ghtt_hni.xlsx` | `api_transition/downloads/ghtt/ghtt_hni report.xlsx` | Chuẩn hóa summary cấp TTVT |
| GHTT Sơn Tây | chưa có hàm cũ riêng 1-1 | `GHTT/tong_hop_ghtt_sontay.xlsx` | `api_transition/downloads/ghtt/ghtt_sontay report.xlsx` | Chuẩn hóa summary cấp tổ |
| GHTT NVKTDB | `process_GHTT_report_NVKT` | `GHTT/tong_hop_ghtt_nvktdb.xlsx` | `api_transition/downloads/ghtt/ghtt_nvktdb report.xlsx` | Port trực tiếp, có map `đơn vị` |
| Vật tư thu hồi | `vat_tu_thu_hoi_process` | `downloads/baocao_hanoi/bc_thu_hoi_vat_tu.xlsx` | `api_transition/downloads/vat_tu_thu_hoi/bc_thu_hoi_vat_tu.xlsx` | Port trực tiếp |
| Quyết toán vật tư | chưa có hàm cũ riêng | chưa có | `api_transition/downloads/vat_tu_thu_hoi/quyet_toan_vat_tu.xlsx` | Thiết kế mới, tham chiếu pattern vật tư |

## 3.2. Nhóm mới, chưa có processor cũ 1-1

| File mới | Processor cũ gần nhất | Hướng xử lý dự kiến |
|---|---|---|
| `api_transition/downloads/phieu_hoan_cong_dich_vu/phieu_hoan_cong_dich_vu_chi_tiet.xlsx` | `process_hoan_cong_report` trong `thuc_tang_process.py` | Đã triển khai: tạo `Data`, tổng hợp theo NVKT, theo tổ, theo TTVT |
| `api_transition/downloads/tam_dung_khoi_phuc_dich_vu/tam_dung_khoi_phuc_dich_vu_chi_tiet.xlsx` | `process_ngung_psc_report` trong `thuc_tang_process.py` | Đã triển khai: tạo `Data`, tổng hợp theo NVKT, theo tổ, theo TTVT, theo lý do hủy |
| `api_transition/downloads/tam_dung_khoi_phuc_dich_vu/tam_dung_khoi_phuc_dich_vu_tong_hop.xlsx` | phần tổng hợp trong `thuc_tang_process.py` | Đã triển khai bản an toàn: nếu workbook rỗng thì ghi sheet thông báo |
| `api_transition/downloads/ty_le_xac_minh/ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv.xlsx` | `process_KR6_report_NVKT` / `process_KR7_report_NVKT` | Đã triển khai: clean summary cấp tổ |
| `api_transition/downloads/ty_le_xac_minh/ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet.xlsx` | `process_c14_chitiet_report` + `KR_process` | Đã triển khai: tạo `Data`, tổng hợp theo NVKT, theo tổ, theo loại phiếu |

## 4. Hướng tổ chức module mới trong `api_transition`

Đề xuất tạo thư mục:

`api_transition/processors/`

Các module bên trong:

- `common.py`
  - helper đọc/ghi Excel
  - helper replace/add sheet
  - helper auto width, border, header style
  - helper normalize text
  - helper lookup `đơn vị` từ `dsnv.xlsx`
  - helper resolve path trong `api_transition/downloads`

- `c_processors.py`
  - port các hàm tổng hợp C1.1/C1.2/C1.3/C1.4
  - port C1.4 chi tiết
  - giữ nguyên tên sheet nghiệp vụ đang dùng nếu không có lý do đổi

- `i15_processors.py`
  - port toàn bộ logic `i15_process.py`
  - đưa path `danhba.db`, history DB và thư mục output thành tham số/config
  - tách phần tracking DB khỏi phần transform DataFrame nếu có thể

- `kpi_processors.py`
  - `process_kpi_nvkt_c11`
  - `process_kpi_nvkt_c12`
  - `process_kpi_nvkt_c13`
  - dùng chung một core processor vì layout tương tự

- `kq_tiep_thi_processors.py`
  - port `process_kq_tiep_thi_report`
  - bỏ logic tìm file theo ngày
  - thêm sheet chi tiết + tổng hợp như cũ

- `ghtt_processors.py`
  - xử lý 3 nhánh:
    - `ghtt_hni`
    - `ghtt_sontay`
    - `ghtt_nvktdb`
  - với `nvktdb`:
    - bóc tên NVKT
    - map `đơn vị` từ `dsnv.xlsx`
    - giữ lại các cột KPI đã chuẩn hóa

- `service_flow_processors.py`
  - nhóm báo cáo giống `thuc_tang_process.py`
  - gồm:
    - hoàn công chi tiết
    - tạm dừng/khôi phục chi tiết
    - tổng hợp tạm dừng/khôi phục
    - thực tăng dẫn xuất nếu sau này cần
  - điểm chính:
    - normalize NVKT
    - group theo tổ
    - group theo cá nhân

- `cau_hinh_tu_dong_processors.py`
  - xử lý 3 nhánh:
    - `cau_hinh_tu_dong_ptm`
    - `cau_hinh_tu_dong_thay_the`
    - `cau_hinh_tu_dong_chi_tiet`
  - với `PTM` và `Thay thế`:
    - làm sạch HTML trong cột `Đơn vị`
    - bóc dòng `TTVT`
    - tạo summary cấp `TTVT` và cấp `Tổ`
  - với `chi_tiet`:
    - bóc `Mã nhân viên`
    - chuẩn hóa `NVKT`
    - tổng hợp theo tổ, theo NVKT, theo lỗi

- `verification_processors.py`
  - xử lý 2 báo cáo xác minh mới
  - 1 processor cho dạng summary KPI
  - 1 processor cho dạng detail

- `vattu_processors.py`
  - `process_vat_tu_thu_hoi`
  - `process_quyet_toan_vat_tu`
  - chia phần normalize và phần group summary

- `registry.py`
  - map `report_key -> processor function`
  - làm cơ sở cho batch processing sau này

- `batch_process.py`
  - chỉ gọi các processor đã đăng ký
  - không làm logic transform

## 5. Quy tắc refactor

### 5.1. Quy tắc path

- Không dùng path cứng kiểu:
  - `downloads/baocao_hanoi/...`
  - `KPI-DOWNLOAD/...`
  - `KQ-TIEP-THI/...`
  - `PTTB-PSC/...`
- Tất cả processor mới sẽ nhận path explicit hoặc resolve từ:
  - `api_transition/downloads/<group>/<file>`
- Tất cả output sau xử lý sẽ resolve tới:
  - `api_transition/Processed/<group>/<file_processed>.xlsx`

### 5.1.1. Quy tắc tạo file processed

- Với input:
  - `api_transition/downloads/chi_tieu_c/c1.1 report.xlsx`
- Output processed tương ứng:
  - `api_transition/Processed/chi_tieu_c/c1.1 report_processed.xlsx`

- Với input:
  - `api_transition/downloads/kpi_nvkt/c11-nvktdb report.xlsx`
- Output processed tương ứng:
  - `api_transition/Processed/kpi_nvkt/c11-nvktdb report_processed.xlsx`

- Quy tắc tên file:
  - giữ nguyên base name
  - chèn `_processed` trước extension
  - ví dụ:
    - `foo.xlsx` -> `foo_processed.xlsx`
    - `bar report.xlsx` -> `bar report_processed.xlsx`

### 5.2. Quy tắc dependency

Các dependency ngoài report raw cần gom về config:

- `dsnv.xlsx`
- `danhba.db`
- `suy_hao_history.db`
- `suy_hao_history_k2.db`

Không hard-code trực tiếp trong processor; nên để:
- tham số hàm
- hoặc 1 file config nội bộ `api_transition/settings.py`

### 5.2.1. Quy tắc domain logic không hard-code tên riêng

- Không viết điều kiện kiểu:
  - loại bỏ cứng một số tổ theo tên
  - chỉ group theo một danh sách tổ đã biết
  - giả định một tập NVKT cố định
- Nếu cần loại trừ dữ liệu:
  - phải dựa trên rule tổng quát
  - hoặc cấu hình ngoài, không nhúng trực tiếp trong code processor
- Nếu nghiệp vụ yêu cầu loại trừ một số tổ/đơn vị/cá nhân cụ thể:
  - không hard-code trực tiếp trong hàm xử lý
  - phải đi qua rule cấu hình hoặc tham số đầu vào
  - ví dụ:
    - `exclude_unit_patterns`
    - `exclude_team_patterns`
    - `exclude_person_patterns`
- Rule loại trừ phải là tùy chọn:
  - mặc định processor chạy trên toàn bộ dữ liệu raw
  - chỉ khi có cấu hình mới áp dụng loại trừ
- Nếu cần nhận diện cấp dữ liệu:
  - ưu tiên suy ra từ cột hiện có trong raw:
    - `DOIVT`, `TEN_DOI`, `TTVT`, `DONVI`, `Đơn vị`, `NVKT`, `TEN_KV`, `NHANVIEN_KT`, ...
  - không giả định tên cột của riêng một đơn vị nếu có thể dò linh hoạt

### 5.2.2. Quy tắc trích xuất động từ báo cáo

- Processor nên có helper chung để:
  - dò cột ứng viên cho cấp tổ
  - dò cột ứng viên cho cấp trung tâm
  - dò cột ứng viên cho cấp cá nhân
  - chuẩn hóa chuỗi tên từ nhiều format khác nhau
- Ví dụ:
  - cùng một ý nghĩa “tổ” có thể xuất hiện ở các cột:
    - `DOIVT`
    - `TEN_DOI`
    - `Đơn vị`
  - cùng một ý nghĩa “NVKT” có thể xuất hiện ở:
    - `NVKT`
    - `TEN_KV`
    - `NHANVIEN_KT`
    - `Đơn vị/Nhân viên KT`
- Kế hoạch code nên ưu tiên một lớp helper `schema inference` nhẹ thay vì viết processor bám cứng vào một bộ cột duy nhất.

### 5.3. Quy tắc output

- Không ghi sheet vào file raw trong `downloads`.
- Mọi processor đều phải làm việc trên file copy trong `Processed`.
- Dù logic cũ ghi vào chính file nguồn hay tạo file processed riêng:
  - trong `api_transition` sẽ thống nhất một pattern:
    - copy raw
    - ghi kết quả vào file copy có hậu tố `_processed`
- Nếu báo cáo chi tiết tới cá nhân:
  - phải có ít nhất:
    - sheet tổng hợp theo cá nhân
    - sheet tổng hợp theo tổ
- Nếu báo cáo đã là summary cấp tổ:
  - chỉ cần sheet chuẩn hóa tổng hợp

## 6. Danh sách processor cần viết trong `api_transition`

## 6.1. Port trực tiếp từ code cũ

- `process_c11_report_api_output`
- `process_c12_report_api_output`
- `process_c13_report_api_output`
- `process_c14_report_api_output`
- `process_c14_chitiet_report_api_output`
- `process_i15_report_with_tracking_api_output`
- `process_i15_k2_report_with_tracking_api_output`
- `process_kpi_nvkt_c11_api_output`
- `process_kpi_nvkt_c12_api_output`
- `process_kq_tiep_thi_api_output`
- `process_ghtt_nvktdb_api_output`
- `process_vat_tu_thu_hoi_api_output`

## 6.2. Viết mới nhưng tái sử dụng pattern cũ

- `process_kpi_nvkt_c13_api_output`
  - reuse core KPI NVKT

- `process_ghtt_hni_api_output`
  - normalize report tổng hợp cấp TTVT

- `process_ghtt_sontay_api_output`
  - normalize report tổng hợp cấp tổ

- `process_cau_hinh_tu_dong_ptm_api_output`
  - làm sạch report summary
  - bóc `TTVT` cha và tổng hợp cấp tổ

- `process_cau_hinh_tu_dong_thay_the_api_output`
  - cùng pattern với `PTM`

- `process_cau_hinh_tu_dong_chi_tiet_api_output`
  - normalize `NVKT`
  - tổng hợp theo tổ
  - tổng hợp theo cá nhân
  - tổng hợp lỗi cấu hình

- `process_phieu_hoan_cong_dich_vu_chi_tiet_api_output`
  - theo pattern `process_hoan_cong_report`
  - tạo:
    - `Data`
    - tổng hợp theo NVKT
    - tổng hợp theo tổ

- `process_tam_dung_khoi_phuc_dich_vu_chi_tiet_api_output`
  - theo pattern `process_ngung_psc_report`
  - tạo:
    - `Data`
    - tổng hợp theo NVKT
    - tổng hợp theo tổ

- `process_tam_dung_khoi_phuc_dich_vu_tong_hop_api_output`
  - normalize workbook summary
  - nếu raw đã ở mức tổ thì giữ sheet tổng hợp
  - nếu raw có cột cá nhân thì bổ sung sheet theo NVKT

- `process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api_output`
  - nếu dòng dữ liệu là NVKT/cá nhân:
    - chuẩn hóa tên
    - map đơn vị
    - sort theo KPI
    - tạo sheet theo đơn vị
  - nếu chỉ là cấp tổ:
    - tạo summary clean sheet

- `process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api_output`
  - dữ liệu chi tiết -> phải tạo:
    - tổng hợp theo cá nhân
    - tổng hợp theo tổ
    - chi tiết raw chuẩn hóa

- `process_quyet_toan_vat_tu_api_output`
  - chưa có hàm cũ 1-1
  - định hướng:
    - giữ `Data`
    - tạo `Tổng hợp theo loại/SPDV`
    - tạo `Tổng hợp theo mã vật tư`
    - nếu có cột đơn vị hoặc NVKT thì thêm sheet group tương ứng

## 6.3. Helper bắt buộc để hỗ trợ mở rộng đa đơn vị

- `infer_grouping_columns(df)`
  - suy ra các cột đại diện cho:
    - trung tâm
    - tổ/đội
    - cá nhân

- `normalize_person_name(value)`
  - bóc tên người từ các format:
    - `Mã-Tên`
    - `Khu vực - Tên`
    - `Tên(...)`

- `normalize_team_name(value)`
  - chuẩn hóa tên tổ/đội nếu report chứa text thừa

- `build_processed_path(input_path, group_name)`
  - sinh path file `_processed.xlsx`

- `copy_raw_to_processed(input_path, processed_path)`
  - tạo bản copy trước khi ghi sheet

- `append_or_replace_sheet(processed_path, sheet_name, df)`
  - ghi sheet chuẩn hóa vào file processed

## 7. Thứ tự triển khai đề xuất

### Phase 1: Port nhanh, ít rủi ro

- C1.1
- C1.2
- C1.3
- C1.4
- C1.4 chi tiết
- KPI NVKT C11/C12/C13
- KQ tiếp thị
- Vật tư thu hồi

Mục tiêu:
- xác nhận khung `api_transition/processors` và helper path hoạt động đúng
- xác nhận helper `copy_raw_to_processed()` và `build_processed_path()` hoạt động đúng cho mọi group
- xác nhận bộ helper suy luận cột hoạt động được với nhiều layout báo cáo khác nhau

## 8. Trạng thái triển khai hiện tại

### 8.1. Hạ tầng đã hoàn thành

- Đã tạo package:
  - `api_transition/processors/`
- Đã tạo helper chung:
  - `api_transition/processors/common.py`
- Các helper hiện đã có:
  - `build_processed_path()`
  - `copy_raw_to_processed()`
  - `ensure_processed_workbook()`
  - `append_or_replace_sheet()`
- Pattern output đã chạy thật:
  - input đọc từ `api_transition/downloads/<group>/...`
  - output ghi vào `api_transition/Processed/<group>/*_processed.xlsx`
  - file raw không bị sửa

### 8.2. Nhóm `chi_tieu_c` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/c_processors.py`

Các hàm đã có:
- `process_c11_report_api_output`
- `process_c12_report_api_output`
- `process_c13_report_api_output`
- `process_c14_report_api_output`
- `process_c14_chitiet_report_api_output`
- `process_c11_chitiet_report_api_output`
- `process_c12_chitiet_sm1_report_api_output`
- `process_c12_chitiet_sm2_report_api_output`
- `process_c12_chitiet_reports_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/chi_tieu_c/c1.1 report_processed.xlsx`
- `api_transition/Processed/chi_tieu_c/c1.2 report_processed.xlsx`
- `api_transition/Processed/chi_tieu_c/c1.3 report_processed.xlsx`
- `api_transition/Processed/chi_tieu_c/c1.4 report_processed.xlsx`
- `api_transition/Processed/chi_tieu_c/c1.4_chitiet_report_processed.xlsx`
- `api_transition/Processed/chi_tieu_c/c1.1_chitiet_report_processed.xlsx`
- `api_transition/Processed/chi_tieu_c/c1.2_chitiet_sm1_report_processed.xlsx`
- `api_transition/Processed/chi_tieu_c/c1.2_chitiet_sm2_report_processed.xlsx`

### 8.3. Nhóm `kpi_nvkt` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/kpi_processors.py`

Các hàm đã có:
- `process_kpi_nvkt_c11_api_output`
- `process_kpi_nvkt_c12_api_output`
- `process_kpi_nvkt_c13_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/kpi_nvkt/c11-nvktdb report_processed.xlsx`
- `api_transition/Processed/kpi_nvkt/c12-nvktdb report_processed.xlsx`
- `api_transition/Processed/kpi_nvkt/c13-nvktdb report_processed.xlsx`

### 8.4. Nhóm `kq_tiep_thi` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/kq_tiep_thi_processors.py`

Các hàm đã có:
- `process_kq_tiep_thi_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/kq_tiep_thi/kq_tiep_thi report_processed.xlsx`

### 8.5. Nhóm `ghtt` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/ghtt_processors.py`

Các hàm đã có:
- `process_ghtt_hni_api_output`
- `process_ghtt_sontay_api_output`
- `process_ghtt_nvktdb_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/ghtt/ghtt_hni report_processed.xlsx`
- `api_transition/Processed/ghtt/ghtt_sontay report_processed.xlsx`
- `api_transition/Processed/ghtt/ghtt_nvktdb report_processed.xlsx`

### 8.6. Nhóm `cau_hinh_tu_dong` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/cau_hinh_tu_dong_processors.py`

Các hàm đã có:
- `process_cau_hinh_tu_dong_ptm_api_output`
- `process_cau_hinh_tu_dong_thay_the_api_output`
- `process_cau_hinh_tu_dong_chi_tiet_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/cau_hinh_tu_dong/cau_hinh_tu_dong_ptm_processed.xlsx`
- `api_transition/Processed/cau_hinh_tu_dong/cau_hinh_tu_dong_thay_the_processed.xlsx`
- `api_transition/Processed/cau_hinh_tu_dong/cau_hinh_tu_dong_chi_tiet_processed.xlsx`

### 8.7. Nhóm `service_flow` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/service_flow_processors.py`

Các hàm đã có:
- `process_phieu_hoan_cong_dich_vu_chi_tiet_api_output`
- `process_tam_dung_khoi_phuc_dich_vu_chi_tiet_api_output`
- `process_tam_dung_khoi_phuc_dich_vu_tong_hop_api_output`
- `process_mytv_ngung_psc_api_output`
- `process_mytv_hoan_cong_api_output`
- `process_mytv_thuc_tang_api_output`
- `process_son_tay_mytv_ngung_psc_t_minus_1_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/phieu_hoan_cong_dich_vu/phieu_hoan_cong_dich_vu_chi_tiet_processed.xlsx`
- `api_transition/Processed/tam_dung_khoi_phuc_dich_vu/tam_dung_khoi_phuc_dich_vu_chi_tiet_processed.xlsx`
- `api_transition/Processed/tam_dung_khoi_phuc_dich_vu/tam_dung_khoi_phuc_dich_vu_tong_hop_processed.xlsx`
- `api_transition/Processed/mytv_dich_vu/mytv_ngung_psc_processed.xlsx`
- `api_transition/Processed/mytv_dich_vu/mytv_hoan_cong_processed.xlsx`
- `api_transition/Processed/mytv_dich_vu/mytv_thuc_tang_processed.xlsx`
- `api_transition/Processed/mytv_dich_vu/ngung_psc_mytv_thang_t-1_sontay_processed.xlsx`

Ghi chú trạng thái:
- `process_mytv_ngung_psc_api_output` đã chuyển sang đọc raw schema mới:
  - `api_transition/downloads/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_cap_to.xlsx`
  - `api_transition/downloads/mytv_dich_vu/ngung_psc_mytv_thang_t-1_cap_ttvt.xlsx`
- `process_mytv_hoan_cong_api_output` không còn phụ thuộc file legacy `mytv_hoan_cong.xlsx`; dữ liệu MyTV được tách trực tiếp từ `phieu_hoan_cong_dich_vu_chi_tiet.xlsx`.
- `process_mytv_thuc_tang_api_output` hiện tính được cấp tổ từ 2 nguồn mới ở trên.
- `process_son_tay_mytv_ngung_psc_t_minus_1_api_output` đã map đúng sang `api_transition/downloads/tam_dung_khoi_phuc_dich_vu/ngung_psc_mytv_thang_t-1_sontay.xlsx`.
- Đã test lại runner cho 4 task `mytv_ngung_psc`, `mytv_hoan_cong`, `mytv_thuc_tang`, `son_tay_mytv_ngung_psc_t_minus_1` và đều chạy thành công.
- Đã test import SQLite end-to-end trên DB tạm cho 3 workbook:
  - `mytv_ngung_psc_processed.xlsx`
  - `mytv_hoan_cong_processed.xlsx`
  - `mytv_thuc_tang_processed.xlsx`
- Hạn chế hiện còn:
  - raw MyTV ngưng PSC API mới chỉ có cấp tổ/TTVT, chưa có chi tiết NVKT
  - vì vậy `mytv_thuc_tang_processed.xlsx` hiện chỉ sinh `thuc_tang_theo_to`; thay cho sheet NVKT là `thong_bao`

### 8.8. Nhóm `ty_le_xac_minh` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/verification_processors.py`

Các hàm đã có:
- `process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api_output`
- `process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/ty_le_xac_minh/ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_processed.xlsx`
- `api_transition/Processed/ty_le_xac_minh/ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_processed.xlsx`

### 8.9. Nhóm `vat_tu_thu_hoi` đã triển khai

Đã viết và test thành công trong:
- `api_transition/processors/vattu_processors.py`

Các hàm đã có:
- `process_vat_tu_thu_hoi_api_output`
- `process_quyet_toan_vat_tu_api_output`

Các file processed đã tạo và kiểm tra:
- `api_transition/Processed/vat_tu_thu_hoi/bc_thu_hoi_vat_tu_processed.xlsx`
- `api_transition/Processed/vat_tu_thu_hoi/quyet_toan_vat_tu_processed.xlsx`

### 8.10. Sheet đầu ra đã chốt

- `C1.1`:
  - sheet `TH_C1.1`
- `C1.2`:
  - sheet `TH_C1.2`
- `C1.3`:
  - sheet `TH_C1.3`
- `C1.4`:
  - sheet `TH_C1.4`
- `C1.4 chi tiết`:
  - sheet `TH_HL_NVKT`
- `C1.1 chi tiết`:
  - sheet `chi_tiet`
  - sheet `chi_tieu_ko_hen_15h`
  - sheet `chi_tieu_ko_hen_16h`
  - sheet `chi_tieu_ko_hen_17h`
  - sheet `chi_tieu_ko_hen_18h`
  - các sheet `chi_tiet_khong_dat_*h`
- `C1.2 chi tiết`:
  - `TH_phieu_hong_lai_7_ngay` trên file `SM1`
  - `Tong_phieu_bao_hong_thang` trên file `SM2`
  - `TH_SM1C12_HLL_Thang` trên file `SM1`
- `KPI NVKT C11`:
  - sheet `c11 kpi nvkt`
- `KPI NVKT C12`:
  - sheet `c12 kpi nvkt`
- `KPI NVKT C13`:
  - sheet `c13 kpi nvkt`
- `KQ tiếp thị`:
  - sheet `kq_tiep_thi`
  - sheet `kq_th`
- `GHTT HNI`:
  - sheet `kq_hni`
- `GHTT Sơn Tây`:
  - sheet `kq_sontay`
- `GHTT NVKTDB`:
  - sheet `kq_nvktdb`
- `Cấu hình tự động PTM`:
  - sheet `du_lieu_sach`
  - sheet `tong_hop_ttvt`
  - sheet `tong_hop_to`
- `Cấu hình tự động Thay thế`:
  - sheet `du_lieu_sach`
  - sheet `tong_hop_ttvt`
  - sheet `tong_hop_to`
- `Cấu hình tự động Chi tiết`:
  - sheet `chi_tiet`
  - sheet `th_theo_to`
  - sheet `th_theo_nvkt`
  - sheet `tong_hop_loi`
- `Phiếu hoàn công dịch vụ chi tiết`:
  - sheet `Data`
  - sheet `fiber_hoan_cong_thang`
  - sheet `fiber_hoan_cong_thang_theo_to`
  - sheet `fiber_hoan_cong_thang_theo_ttvt`
- `Tạm dừng/khôi phục dịch vụ chi tiết`:
  - sheet `Data`
  - sheet `fiber_ngung_psc_thang`
  - sheet `fiber_ngung_psc_thang_theo_to`
  - sheet `fiber_ngung_psc_thang_theo_ttvt`
  - sheet `tong_hop_ly_do_huy`
- `Tạm dừng/khôi phục dịch vụ tổng hợp`:
  - sheet `thong_bao` nếu workbook nguồn đang rỗng
  - hoặc `du_lieu_sach` nếu raw có dữ liệu
- `MyTV ngưng PSC`:
  - sheet `Data`
  - sheet `mytv_ngung_psc_thang`
  - sheet `mytv_ngung_psc_thang_theo_to`
  - sheet `mytv_ngung_psc_thang_theo_ttvt`
  - sheet `source_cap_to`
  - sheet `source_cap_ttvt`
- `MyTV hoàn công`:
  - sheet `Data`
  - sheet `mytv_hoan_cong_thang`
  - sheet `mytv_hoan_cong_thang_theo_to`
  - sheet `mytv_hoan_cong_thang_theo_ttvt`
- `MyTV thực tăng`:
  - sheet `thuc_tang_theo_to`
  - sheet `thong_bao` nếu nguồn ngưng PSC chưa có chi tiết NVKT
- `MyTV Sơn Tây ngưng PSC tháng T-1`:
  - sheet `TH_ngung_PSC-Thang T-1`
- `Tỷ lệ xác minh TTVTKV`:
  - sheet `tong_hop_ttvtkv`
- `Tỷ lệ xác minh chi tiết`:
  - sheet `Data`
  - sheet `tong_hop_theo_nvkt`
  - sheet `tong_hop_theo_to`
  - sheet `tong_hop_theo_loai_phieu`
- `Vật tư thu hồi`:
  - sheet `Chi tiết`
  - sheet `Chi tiết vật tư`
  - sheet `Tổng hợp`
- `Quyết toán vật tư`:
  - sheet `Data`
  - sheet `Tong_hop_theo_loai`
  - sheet `Tong_hop_theo_SPDV`
  - sheet `Tong_hop_theo_vat_tu`

## 9. Các lưu ý thực tế đã rút ra khi triển khai

### 9.1. Rule loại trừ phải đi qua cấu hình/tham số

- Đã áp dụng thật cho nhóm `C1.1/C1.2/C1.3`.
- Không nhúng logic loại trừ trực tiếp trong thân xử lý nghiệp vụ.
- Hiện tại rule mặc định đang được cấu hình trong module processor cho nhóm `C`:
  - loại `Tổ Kỹ thuật Địa bàn Bất Bạt`
  - loại `Tổ Kỹ thuật Địa bàn Tùng Thiện`
- Về sau cần đẩy rule này ra config chung nếu số lượng rule tăng lên.

### 9.2. Chuẩn hóa tên NVKT là bắt buộc

- Đã gặp lỗi dữ liệu gốc:
  - `Bùi Văn Cường`
  - `Bùi văn Cường`
- Nếu không chuẩn hóa, processor sẽ group thành 2 NVKT khác nhau.
- Vì vậy helper chuẩn hóa tên người phải:
  - bỏ prefix khu vực nếu có
  - bỏ phần trong ngoặc nếu có
  - chuẩn hóa khoảng trắng
  - chuẩn hóa khác biệt hoa/thường
- Trong bản hiện tại, `C1.4 chi tiết` đã áp dụng chuẩn hóa này trước khi group.
- Trong bản hiện tại, các nhóm sau đã áp dụng chuẩn hóa này:
  - `C1.4 chi tiết`
  - `C1.1 chi tiết`
  - `C1.2 chi tiết`
  - `KPI NVKT`
  - `GHTT NVKTDB`
  - `Cấu hình tự động chi tiết`
  - `Phiếu hoàn công dịch vụ chi tiết`
  - `Tạm dừng/khôi phục dịch vụ chi tiết`
  - `Tỷ lệ xác minh chi tiết`
  - `Vật tư thu hồi`
- Kết luận:
  - mọi processor chi tiết theo cá nhân/NVKT về sau phải tái sử dụng chung helper `normalize_person_name()`

### 9.2.1. Chuẩn hóa từ `dsnv.xlsx` vẫn phải chấp nhận dữ liệu không map được

- Dù `dsnv.xlsx` đã được cập nhật, vẫn có thể còn các dòng:
  - dòng tổng (`Tổng`)
  - tên không có trong danh bạ tham chiếu
- Không nên ép map sai chỉ để đạt đủ 100%.
- Processor phải:
  - map được bao nhiêu thì map
  - giữ trống `đơn vị` cho các dòng không match
  - để người vận hành quyết định có bổ sung `dsnv.xlsx` hay không

### 9.3. Ưu tiên group động theo cột có sẵn trong raw

- `C1.4 chi tiết` cho thấy raw report đã có đủ:
  - `TEN_KV`
  - `DOIVT`
  - `TTVT`
- Vì vậy có thể group trực tiếp theo dữ liệu raw mà chưa cần lookup ngoài.
- Chỉ nên fallback sang `dsnv.xlsx` hoặc nguồn khác nếu raw report không đủ metadata.

### 9.4. Processor summary và processor detail phải tách vai trò rõ

- Báo cáo summary như `C1.1/C1.2/C1.3/C1.4`:
  - chủ yếu là chuẩn hóa sheet tổng hợp
- Báo cáo summary như `GHTT HNI`, `GHTT Sơn Tây`:
  - chủ yếu là chuẩn hóa cột và bỏ dòng sub-header
- Báo cáo detail như `C1.4 chi tiết`:
  - phải có bước chuẩn hóa định danh cá nhân
  - sau đó mới group theo cá nhân/tổ
- Báo cáo dạng `NVKTDB` như `KPI NVKT`, `GHTT NVKTDB`:
  - là summary theo cá nhân nhưng vẫn cần bóc tên NVKT và enrich `đơn vị`
- Điều này cần giữ nhất quán cho các nhóm sắp làm:
  - `C1.1 chi tiết`
  - `C1.2 chi tiết`
  - `xác minh`
  - `hoàn công / tạm dừng PSC`

### 9.6. Một số raw report có dòng sub-header nằm trong dữ liệu

- Đã gặp ở nhóm `GHTT`:
  - file Excel tải về không còn header 2 dòng thực sự
  - thay vào đó, dòng đầu của data là dòng mô tả cột con
- Processor phải chủ động:
  - đọc raw bình thường
  - bỏ dòng đầu nếu đó là sub-header lặp lại
  - sau đó mới đổi tên cột chuẩn hóa

### 9.8. Một số raw report summary chứa HTML hoặc markup trong dữ liệu

- Đã gặp ở nhóm `cau_hinh_tu_dong`:
  - dòng tổng `TTVT` được render dưới dạng `<b>TTVT ...</b>`
- Processor phải:
  - strip HTML tags trước khi group
  - dùng chính dòng `TTVT` đã làm sạch để fill xuống cho các dòng `Tổ`
- Không nên giả định raw đã là text thuần.

### 9.9. Sheet tổng hợp theo cá nhân phải loại các dòng thiếu định danh

- Đã gặp ở `cau_hinh_tu_dong_chi_tiet.xlsx`:
  - có các record thiếu `Nhân viên phụ trách`
  - đồng thời thiếu cả `TTVT` và `Đội Viễn thông`
- Nếu group trực tiếp với `dropna=False`, sẽ sinh thêm dòng cá nhân/tổ rỗng.
- Rule đã áp dụng:
  - summary theo `Tổ` hoặc `NVKT` phải `dropna(subset=group_cols)` trước khi group
  - sheet `chi_tiet` raw chuẩn hóa vẫn giữ nguyên các dòng này để không mất dữ liệu nguồn

### 9.10. Với report mới chỉ nên gắn nhãn trung tính nếu chưa chắc semantic

- Đã gặp ở nhóm `ty_le_xac_minh`:
  - file `chi_tiet` có `474` dòng, đúng bằng `Tổng số phiếu giao XM` trong report tổng hợp
  - vì vậy chưa thể khẳng định raw này chỉ gồm các phiếu “đúng hạn”
- Rule đang áp dụng:
  - summary chi tiết dùng nhãn trung tính:
    - `tong_hop_theo_nvkt`
    - `tong_hop_theo_to`
    - `tong_hop_theo_loai_phieu`
  - tránh đặt tên sheet thể hiện kết luận nghiệp vụ nếu chưa xác minh chắc từ nguồn dữ liệu

### 9.11. Workbook nguồn rỗng vẫn phải xử lý mềm để không gãy batch

- Đã gặp ở `tam_dung_khoi_phuc_dich_vu_tong_hop.xlsx`:
  - file tồn tại nhưng không có dữ liệu
- Rule đang áp dụng:
  - vẫn copy sang `Processed/...`
- ghi sheet `thong_bao` nêu rõ workbook nguồn đang rỗng
- không throw exception làm gãy luồng batch

### 9.13. Không đánh dấu hoàn tất nghiệp vụ nếu mới chỉ test bằng file legacy ngoài `api_transition/downloads`

- Đã gặp ở nhóm `MyTV` thuộc `thuc_tang_process.py`.
- Dù processor đã chạy được trên file thật trong thư mục legacy `PTTB-PSC/`, đây vẫn chưa phải trạng thái hoàn tất chuẩn hóa end-to-end.
- Điều kiện để chốt hoàn tất cho nhóm này là:
  - có downloader API tương ứng
  - raw file được lưu đúng vào `api_transition/downloads/...`
  - processor được chạy lại trên raw mới đó
- Trước khi đạt điều kiện trên, chỉ coi đây là:
  - port logic xử lý đã sẵn sàng
  - chưa khóa luồng đầu vào chính thức

### 9.12. `Quyết toán vật tư` là summary thuần, không ép group theo NVKT

- File `quyet_toan_vat_tu.xlsx` hiện không có cột tổ/NVKT.
- Vì vậy processor chỉ tạo các summary đúng semantic của raw:
  - theo `LOAI`
  - theo `MA_SPDV`
  - theo `MA_VT/TEN_VT`
- Không nên cố bịa thêm sheet theo cá nhân hoặc đơn vị nếu nguồn không có thông tin đó.

### 9.7. `KQ tiếp thị` cần giữ cách ghi header riêng

- File `KQ tiếp thị` sử dụng `MultiIndex` header nhưng output mong muốn là một dòng header phẳng, dễ đọc.
- Vì vậy processor hiện dùng cách:
  - ghi dữ liệu từ dòng 2
  - tự dựng dòng header ở dòng 1
  - định dạng border / width trực tiếp qua `openpyxl`
- Đây là ngoại lệ hợp lệ, không nên ép dùng helper ghi sheet đơn giản cho báo cáo này.

### 9.5. Test phải đi cùng từng hàm

- Cách làm hiện tại đã chứng minh ổn:
  - viết 1 hàm
  - chạy ngay trên file raw thật
  - kiểm tra workbook processed
  - kiểm tra shape/sheet/mẫu dữ liệu đầu ra
- Quy trình này nên giữ nguyên cho các processor tiếp theo để tránh port sai logic cũ.

### Phase 2: Port khó hơn do state/history

- I1.5
- I1.5 K2

Mục tiêu:
- tách logic history DB ra khỏi path hard-code

### Phase 3: Báo cáo mới tương tự logic cũ

- Đã hoàn thành:
  - Phiếu hoàn công dịch vụ chi tiết
  - Tạm dừng/khôi phục dịch vụ chi tiết
  - Tạm dừng/khôi phục dịch vụ tổng hợp
  - Port lại logic MyTV theo raw schema mới trong `api_transition/downloads`
  - Tỷ lệ xác minh TTVTKV
  - Tỷ lệ xác minh chi tiết
  - Vật tư thu hồi
  - Quyết toán vật tư

Mục tiêu:
  - đã bao phủ phần lớn raw file mới trong `api_transition`
  - riêng các nhánh còn thiếu raw chi tiết NVKT thực sự sẽ quay lại khi có API/export phù hợp
  - phần còn lại tập trung vào các processor theo dõi lịch sử/state như `I1.5`

### Phase 4: Orchestrator

- Đã triển khai:
  - `api_transition/processors/runner.py`
    - `run_all_processors()`
    - registry 30 processor
  - `api_transition/full_pipeline.py`
    - `run_full_pipeline()`
    - flow `login -> download -> process -> archive ProcessedDaily -> import report_history.db`
- Trạng thái hiện tại:
  - `ProcessedDaily/<snapshot-date>/...` được tạo ngay sau stage process thành công, không còn phụ thuộc việc import SQLite xong mới copy
  - full pipeline mặc định cho phép tiếp tục import DB với các workbook đã process thành công dù còn lỗi ở một số bước khác
  - import SQLite không còn quét bừa toàn bộ `Processed`; chỉ import các workbook success của lượt chạy hiện tại để tránh nạp dữ liệu stale
  - vẫn có chế độ chặt qua `--strict` nếu muốn dừng trước import khi download/process có lỗi

## 8. Các rủi ro cần lưu ý

- `i1.5 report.xlsx` và `i1.5_k2 report.xlsx` trong `api_transition` đang dùng chữ thường; code cũ dùng `I1.5...`. Trên Linux cần map path rõ ràng.
- Nhiều hàm cũ phụ thuộc dữ liệu ngoài:
  - `dsnv.xlsx`
  - `danhba.db`
  - DB history
- Một số báo cáo mới chưa chắc raw layout giống hoàn toàn báo cáo cũ tương tự.
  - cần kiểm tra bằng file mẫu thật trước khi code processor
- `Tỷ lệ xác minh ... chi tiết` có thể không cùng semantic với `KR6/KR7`; chỉ nên tái sử dụng pattern group/normalize, không bê nguyên business formula.

## 9. Kết luận triển khai

Khi bắt đầu code, hướng làm đúng là:

1. Tạo khung `api_transition/processors/common.py`
   - gồm helper path cho `downloads/` và `Processed/`
   - gồm helper copy raw -> processed
   - gồm helper normalize tên và infer cột động
2. Port nhóm dễ trước:
   - C1
   - KPI
   - KQ tiếp thị
   - Vật tư thu hồi
3. Sau đó mới làm I15
4. Cuối cùng mới làm các báo cáo mới theo pattern tương tự

Trạng thái hiện tại không còn là phương án thuần nữa.

Đã có các processor thật đang chạy trong `api_transition/processors/` cho các nhóm:
- `chi_tieu_c`
- `kpi_nvkt`
- `kq_tiep_thi`
- `ghtt`
- `cau_hinh_tu_dong`
- `service_flow`
- `ty_le_xac_minh`
- `vat_tu_thu_hoi`

Các nhóm còn lại vẫn tiếp tục theo đúng thứ tự phase ở trên.
