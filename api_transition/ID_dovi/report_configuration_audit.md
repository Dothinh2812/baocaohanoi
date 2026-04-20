# Báo cáo Rà soát Chi tiết Tham số & Đường dẫn Báo cáo (Deep Audit)

Báo cáo này cung cấp cái nhìn toàn diện về cấu hình của từng hàm tải báo cáo, bao gồm tham số đầu vào, ID hệ thống, và các đường dẫn lưu trữ.

## 1. Giải thích các Cột thông tin

- **Hàm Download**: Tên hàm Python trong `downloaders.py` hoặc `onebss_downloaders.py`.
- **Tham số Đầu vào**: Danh sách các đối số mà hàm nhận.
- **Report ID**: ID định danh báo cáo trên hệ thống Web (từ Recipe JSON).
- **Key Đơn vị / Thời gian**: Tên tham số trong Payload API dùng để truyền ID đơn vị và ID thời gian.
- **Đường dẫn Download**: Thư mục và tên file lưu trữ bản thô (`downloads/`).
- **Đường dẫn Processed**: Thư mục và tên file sau khi xử lý (`Processed/`).
- **URL Gốc**: Link truy cập báo cáo trên trình duyệt.

---

## 2. Bảng Rà soát Chi tiết

### Nhóm Báo cáo Chỉ tiêu C (Kỳ BC: 26 -> 25)

| Hàm Download | Report ID | Key Đơn vị / Thời gian | Folder Download / Tên file | Folder Processed / Tên file | URL Gốc |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `download_report_c11_api` | `534964` | `ptrungtamid` / `pthang` | `chi_tieu_c` / `c1.1 report.xlsx` | `chi_tieu_c` / `c1.1 report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=534964&menu_id=535020) |
| `download_report_c12_api` | `522513` | `ptrungtamid` / `pthang` | `chi_tieu_c` / `c1.2 report.xlsx` | `chi_tieu_c` / `c1.2 report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=535021) |
| `download_report_c13_api` | `522600` | `ptrungtamid` / `pthang` | `chi_tieu_c` / `c1.3 report.xlsx` | `chi_tieu_c` / `c1.3 report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=522600&menu_id=535022) |
| `download_report_c14_api` | `264107` | `vdonvi` / `vthoigian` | `chi_tieu_c` / `c1.4 report.xlsx` | `chi_tieu_c` / `c1.4 report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=264107&menu_id=275688) |
| `download_report_c14_chitiet_api` | `240277` | `vdvvt` / `vthoigian` | `chi_tieu_c` / `c1.4_chitiet_report.xlsx` | `chi_tieu_c` / `c1.4_chitiet_report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=240277&menu_id=275687) |
| `download_report_c11_chitiet_api` | `267215` | `pdonvi_id` / `vngay_bd`-`kt` | `chi_tieu_c` / `c1.1_chitiet_report.xlsx` | `chi_tieu_c` / `c1.1_chitiet_report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194) |
| `download_report_c12_chitiet_sm1_api` | `267215` | `pdonvi_id` / `vngay_bd`-`kt` | `chi_tieu_c` / `c1.2_chitiet_sm1_report.xlsx` | `chi_tieu_c` / `c1.2_chitiet_sm1_report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194) |
| `download_report_c12_chitiet_sm2_api` | `267215` | `pdonvi_id` / `vngay_bd`-`kt` | `chi_tieu_c` / `c1.2_chitiet_sm2_report.xlsx` | `chi_tieu_c` / `c1.2_chitiet_sm2_report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194) |

### Nhóm Báo cáo KPI NVKT

| Hàm Download | Report ID | Key Đơn vị / Thời gian | Folder Download / Tên file | Folder Processed / Tên file | URL Gốc |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `download_kpi_nvkt_c11_api` | `534964` | `ptrungtamid` / `pthang` | `kpi_nvkt` / `c11-nvktdb report.xlsx` | `kpi_nvkt` / `c11-nvktdb report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=534964&menu_id=535020) |
| `download_kpi_nvkt_c12_api` | `522513` | `ptrungtamid` / `pthang` | `kpi_nvkt` / `c12-nvktdb report.xlsx` | `kpi_nvkt` / `c12-nvktdb report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=535021) |
| `download_kpi_nvkt_c13_api` | `522600` | `ptrungtamid` / `pthang` | `kpi_nvkt` / `c13-nvktdb report.xlsx` | `kpi_nvkt` / `c13-nvktdb report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=522600&menu_id=535022) |

### Nhóm Báo cáo Chỉ tiêu I (Kỳ BC: Ngày hôm trước T-1)

| Hàm Download | Report ID | Key Đơn vị / Thời gian | Folder Download / Tên file | Folder Processed / Tên file | URL Gốc |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `download_report_i15_api` | `283632` | `vdv` / `vngay_bd`-`kt` | `chi_tieu_i` / `i1.5 report.xlsx` | (Chưa implement processor riêng) | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=283632&menu_id=283669) |
| `download_report_i15_k2_api` | `290125` | `vdv` / `vngay_bd`-`kt` | `chi_tieu_i` / `i1.5_k2 report.xlsx` | (Chưa implement processor riêng) | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=290125&menu_id=290161) |

### Nhóm Giao hưởng tận tâm (GHTT)

| Hàm Download | Report ID | Key Đơn vị / Thời gian | Folder Download / Tên file | Folder Processed / Tên file | URL Gốc |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `download_ghtt_report_hni_api` | `534220` | `vdonvi` / `vthoigian` | `ghtt` / `ghtt_hni report.xlsx` | `ghtt` / `ghtt_hni report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238) |
| `download_ghtt_report_sontay_api` | `534220` | `vdonvi` / `vthoigian` | `ghtt` / `ghtt_sontay report.xlsx` | `ghtt` / `ghtt_sontay report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238) |
| `download_ghtt_report_nvktdb_api` | `534220` | `vdonvi` / `vthoigian` | `ghtt` / `ghtt_nvktdb report.xlsx` | `ghtt` / `ghtt_nvktdb report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238) |

### Nhóm Nghiệp vụ & Khác

| Hàm Download | Report ID | Key Đơn vị / Thời gian | Folder Download / Tên file | Folder Processed / Tên file | URL Gốc |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `download_xac_minh_tam_dung_api` | `267844` | `pdonvi_id` / `vngay_bd`-`kt` | `xac_minh_tam_dung` / `xac_minh_tam_dung report.xlsx` | `xac_minh_tam_dung` / `xac_minh_tam_dung report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=267844&menu_id=276199) |
| `download_phieu_hoan_cong_dich_vu_chi_tiet_api` | `283737` | `vdv` / `vngay_bd`-`kt` | `phieu_hoan_cong_dich_vu` / `phieu_hoan_cong_dich_vu_chi_tiet.xlsx` | `phieu_hoan_cong_dich_vu` / `phieu_hoan_cong_dich_vu_chi_tiet_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=283737&menu_id=283774) |
| `download_kq_tiep_thi_api` | `257495` | `vdonvi_id` / `vngay_bd`-`kt` | `kq_tiep_thi` / `kq_tiep_thi report.xlsx` | `kq_tiep_thi` / `kq_tiep_thi report_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=257495&menu_id=276101) |
| `download_report_vattu_thuhoi_api` | `270922` | `vttvt` / `vtungay`-`vdenngay` | `vat_tu_thu_hoi` / `bc_thu_hoi_vat_tu.xlsx` | `vat_tu_thu_hoi` / `bc_thu_hoi_vat_tu_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=270922&menu_id=276242) |
| `download_quyet_toan_vattu_api` | `532729` | `vttvt` / `vtungay`-`vdenngay` | `vat_tu_thu_hoi` / `quyet_toan_vat_tu.xlsx` | `vat_tu_thu_hoi` / `quyet_toan_vat_tu_processed.xlsx` | [Link](https://baocao.hanoi.vnpt.vn/report/report-info?id=532729) |

### Nhóm OneBSS (BI Report)

| Hàm Download | Report ID | Key Đơn vị (Refresh/Override) | Folder Download / Tên file | URL Gốc |
| :--- | :--- | :--- | :--- | :--- |
| `download_hni_pttb_001` | `40618` | `TT_ID`, `DOI_ID` | `onebss` / `HNI_PTTB_001.xlsx` | (Đường dẫn BI Nội bộ) |
| `download_bc_phieu_ton_dv_chi_tiet_hni` | `40618` | `TT_ID`, `DOI_ID` | `onebss` / `bc_phieu_ton_dv_chi_tiet_hni.xlsx` | (Đường dẫn BI Nội bộ) |
| `download_bc_ton_sua_chua_sontay_2026` | `40622` | `TT_ID`, `DOI_ID` | `onebss` / `bc_ton_sua_chua_sontay_2026.xlsx` | (Đường dẫn BI Nội bộ) |
| `download_bc_chi_tiet_ket_qua_cskh_uc3_sontay` | `49544` | `vdonvi_id`, `vphanvung_id` | `onebss` / `bc_chi_tiet_ket_qua_cskh_uc3_sontay.xls` | (ReportViewer) |

---

## 3. Tham số đầu vào các hàm (Signature)

Dưới đây là một số mẫu tham số phổ biến:

- **Hệ thống mới (Recipe-based)**:
  - `month_id`, `month_label`, `unit_id`, `headed`, `output_dir`, `session`
  - Hoặc: `start_date`, `end_date`, `unit_id`, `headed`, `output_dir`, `session`

- **Hệ thống OneBSS**:
  - `unit_id`, `team_id`, `service_ids`, `headed`, `output_dir`, `output_name`, `session`

---

> [!TIP]
> Bạn có thể sử dụng các thông tin **Key Đơn vị / Thời gian** này để cấu hình linh hoạt cho từng đơn vị mà không cần sửa code mỗi lần chạy.
