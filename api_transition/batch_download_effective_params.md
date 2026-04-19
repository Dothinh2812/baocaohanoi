# Batch Download Effective Parameters

Tài liệu này mô tả cấu hình thực tế mà [batch_download.py](/home/vtst/baocaohanoi/api_transition/batch_download.py) đang truyền vào từng hàm download với cấu hình hiện tại trong file.

## Cấu hình gốc hiện tại

Từ [batch_download.py](/home/vtst/baocaohanoi/api_transition/batch_download.py:74):

```text
REPORT_MONTH     = 4
REPORT_YEAR      = 2026
MONTH_ID         = "98944548"
MONTH_LABEL      = ""
VATTU_START_DATE = "24/11/2025"
HEADED           = False
RETRY_TIMEOUTS   = [120, 180, 300]
```

## Giá trị batch tự tính

Từ `_compute_dates()` trong [batch_download.py](/home/vtst/baocaohanoi/api_transition/batch_download.py:102), batch sẽ dùng:

```text
month_id         = "98944548"
month_label      = ""
start_date       = "26/03/2026"
end_date         = "25/04/2026"
cal_start_date   = "01/04/2026"
cal_end_date     = "30/04/2026"
t_minus_1        = "17/04/2026"
vattu_start_date = "24/11/2025"
headed           = False
```

Ghi chú:

- Mọi task dùng shared session đều được truyền thêm `session=...`.
- `session["api_timeout"]` được set theo từng lần retry: `120`, `180`, `300`.
- `month_label` đang rỗng nên các task kiểu `month` chỉ nhận `month_id`.

## Quy tắc build kwargs

Từ `_build_kwargs()` trong [batch_download.py](/home/vtst/baocaohanoi/api_transition/batch_download.py:439):

- `month`: truyền `month_id`, và chỉ truyền `month_label` nếu khác rỗng.
- `date_range`: truyền `start_date`, `end_date` theo kỳ `26 -> 25`.
- `calendar_month`: truyền `start_date`, `end_date` theo tháng dương lịch.
- `t_minus_1`: truyền `start_date/end_date` hoặc `report_date` tùy `extra_kwargs`.
- `date_range_long`: dùng `vattu_start_date` làm `start_date`, còn `end_date` vẫn là cuối kỳ `25`.

Trong [downloaders.py](/home/vtst/baocaohanoi/api_transition/downloaders.py:60), `download_with_recipe()` sẽ ghép `overrides` của từng downloader với month key:

- mặc định month key là `pthang`
- một số report dùng `vthoigian`
- một report dùng `pthoigianid`

## Bảng đối chiếu

| Task | Hàm gọi | Kwargs từ batch | Default còn lại của downloader | Overrides cuối cùng xuống payload API |
|---|---|---|---|---|
| C1.1 | `download_report_c11_api` | `month_id="98944548", headed=False, session=...` | `unit_id="14324"`, `output_dir=downloads/chi_tieu_c` | `ptrungtamid="14324", pthang="98944548"` |
| C1.2 | `download_report_c12_api` | `month_id="98944548", headed=False, session=...` | `unit_id="14324"`, `output_dir=downloads/chi_tieu_c` | `ptrungtamid="14324", pthang="98944548"` |
| C1.3 | `download_report_c13_api` | `month_id="98944548", headed=False, session=...` | `unit_id="14324"`, `output_dir=downloads/chi_tieu_c` | `ptrungtamid="14324", pthang="98944548"` |
| KPI NVKT C1.1 | `download_kpi_nvkt_c11_api` | `month_id="98944548", output_dir=downloads/kpi_nvkt, headed=False, session=...` | `unit_id="14324"` | `ptrungtamid="14324", pthang="98944548"` |
| KPI NVKT C1.2 | `download_kpi_nvkt_c12_api` | `month_id="98944548", output_dir=downloads/kpi_nvkt, headed=False, session=...` | `unit_id="14324"` | `ptrungtamid="14324", pthang="98944548"` |
| KPI NVKT C1.3 | `download_kpi_nvkt_c13_api` | `month_id="98944548", output_dir=downloads/kpi_nvkt, headed=False, session=...` | `unit_id="14324"` | `ptrungtamid="14324", pthang="98944548"` |
| C1.4 | `download_report_c14_api` | `month_id="98944548", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/chi_tieu_c` | `vdonvi="284656", vthoigian="98944548"` |
| C1.4 Chi tiết | `download_report_c14_chitiet_api` | `month_id="98944548", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/chi_tieu_c` | `vdvvt="284656", vthoigian="98944548"` |
| C1.1 Chi tiết | `download_report_c11_chitiet_api` | `start_date="26/03/2026", end_date="25/04/2026", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/chi_tieu_c` | `pdonvi_id="284656", vngay_bd="26/03/2026", vngay_kt="25/04/2026"` |
| C1.2 Chi tiết SM1 | `download_report_c12_chitiet_sm1_api` | `start_date="26/03/2026", end_date="25/04/2026", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/chi_tieu_c` | `pdonvi_id="284656", vngay_bd="26/03/2026", vngay_kt="25/04/2026"` |
| C1.2 Chi tiết SM2 | `download_report_c12_chitiet_sm2_api` | `start_date="26/03/2026", end_date="25/04/2026", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/chi_tieu_c` | `pdonvi_id="284656", vngay_bd="26/03/2026", vngay_kt="25/04/2026"` |
| I1.5 | `download_report_i15_api` | `start_date="17/04/2026", end_date="17/04/2026", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/chi_tieu_i` | `vdv="284656", vngay_bd="17/04/2026", vngay_kt="17/04/2026", vdk="0"` |
| I1.5 K2 | `download_report_i15_k2_api` | `start_date="17/04/2026", end_date="17/04/2026", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/chi_tieu_i` | `vdv="284656", vngay_bd="17/04/2026", vngay_kt="17/04/2026"` |
| GHTT HNI | `download_ghtt_report_hni_api` | `month_id="98944548", headed=False, session=...` | `unit_id="284412"`, `output_dir=downloads/ghtt` | `vdonvi="284412", vloai="1", vthoigian="98944548"` |
| GHTT Sơn Tây | `download_ghtt_report_sontay_api` | `month_id="98944548", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/ghtt` | `vdonvi="284656", vloai="1", vthoigian="98944548"` |
| GHTT NVKT DB | `download_ghtt_report_nvktdb_api` | `month_id="98944548", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/ghtt` | `vdonvi="284656", vloai="2", vthoigian="98944548"` |
| Xác minh tạm dừng | `download_xac_minh_tam_dung_api` | `start_date="01/04/2026", end_date="30/04/2026", headed=False, session=...` | `unit_id="284656"`, `service_ids="8,9"`, `output_dir=downloads/xac_minh_tam_dung` | `pdonvi_id="284656", vngay_bd="01/04/2026", vngay_kt="30/04/2026", vloaidv="8,9", vloaingay="0", vloaibc="0"` |
| Phiếu hoàn công DV chi tiết | `download_phieu_hoan_cong_dich_vu_chi_tiet_api` | `start_date="01/04/2026", end_date="30/04/2026", headed=False, session=...` | `unit_id="284656"`, `service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16"`, `customer_type="0"`, `contract_type="0"`, `ticket_type="0"`, `output_dir=downloads/phieu_hoan_cong_dich_vu` | `vdv="284656", vngay_bd="01/04/2026", vngay_kt="30/04/2026", vloaidv="<default_service_ids>", vloaikh="0", vloaihd="0", vphieu="0"` |
| Tạm dừng, khôi phục DV chi tiết | `download_tam_dung_khoi_phuc_dich_vu_chi_tiet_api` | `start_date="01/04/2026", end_date="30/04/2026", headed=False, session=...` | `unit_id="284656"`, `service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16"`, `date_type="1"`, `report_type="0"`, `output_dir=downloads/tam_dung_khoi_phuc_dich_vu` | `pdonvi_id="284656", vngay_bd="01/04/2026", vngay_kt="30/04/2026", vloaidv="<default_service_ids>", vloaingay="1", vloaibc="0"` |
| Tạm dừng, khôi phục DV chi tiết - khôi phục | `download_tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc_api` | `start_date="01/04/2026", end_date="30/04/2026", headed=False, session=...` | `unit_id="284656"`, `service_ids="8,9"`, `date_type="1"`, `report_type="1"`, `ppageindex="1"`, `ppagesize="1000000"`, `output_dir=downloads/tam_dung_khoi_phuc_dich_vu` | `pdonvi_id="284656", vngay_bd="01/04/2026", vngay_kt="30/04/2026", vloaidv="8,9", vloaingay="1", vloaibc="1", ppageindex="1", ppagesize="1000000"` |
| Tạm dừng, khôi phục DV tổng hợp | `download_tam_dung_khoi_phuc_dich_vu_tong_hop_api` | `start_date="01/04/2026", end_date="30/04/2026", headed=False, session=...` | `unit_id="284656"`, `service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16"`, `report_type="0"`, `output_dir=downloads/tam_dung_khoi_phuc_dich_vu` | `vdv="284656", vngay_bd="01/04/2026", vngay_kt="30/04/2026", vloaidv="<default_service_ids>", vloaibc="0"` |
| Ngưng PSC MyTV tháng T-1 cấp TTVT | `download_ngung_psc_mytv_thang_t_1_cap_ttvt_api` | `report_date="17/04/2026", headed=False, session=...` | `unit_id="14316"`, `service_id="8"`, `report_type="1"`, `output_dir=downloads/tam_dung_khoi_phuc_dich_vu` | `vdvvt_id="8", vdenngay="17/04/2026", vdonvi_id="14316", vloai="1"` |
| Ngưng PSC Fiber tháng T-1 cấp TTVT | `download_ngung_psc_fiber_thang_t_1_cap_ttvt_api` | `report_date="17/04/2026", headed=False, session=...` | `unit_id="14316"`, `service_id="9"`, `report_type="1"`, `output_dir=downloads/tam_dung_khoi_phuc_dich_vu` | `vdvvt_id="9", vdenngay="17/04/2026", vdonvi_id="14316", vloai="1"` |
| Ngưng PSC Fiber tháng T-1 cấp Tổ | `download_ngung_psc_fiber_thang_t_1_cap_to_api` | `report_date="17/04/2026", headed=False, session=...` | `unit_id="14324"`, `service_id="9"`, `report_type="2"`, `output_dir=downloads/tam_dung_khoi_phuc_dich_vu` | `vdvvt_id="9", vdenngay="17/04/2026", vdonvi_id="14324", vloai="2"` |
| Ngưng PSC MyTV tháng T-1 cấp Tổ | `download_ngung_psc_mytv_thang_t_1_cap_to_api` | `report_date="17/04/2026", headed=False, session=...` | `unit_id="14324"`, `service_id="8"`, `report_type="1"`, `output_dir=downloads/tam_dung_khoi_phuc_dich_vu` | `vdvvt_id="8", vdenngay="17/04/2026", vdonvi_id="14324", vloai="1"` |
| Tỷ lệ xác minh đúng thời gian quy định - TTVTKV | `download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api` | `month_id="98944548", headed=False, session=...` | `unit_id="284656"`, `report_scope="-1"`, `verification_type="-1"`, `output_dir=downloads/ty_le_xac_minh` | `vdv="284656", vloai="-1", vloaixacminh="-1", vthoigian="98944548"` |
| Tỷ lệ xác minh đúng thời gian quy định chi tiết | `download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api` | `month_id="98944548", headed=False, session=...` | `unit_id="284656"`, `report_type="-1"`, `contract_type="-1"`, `service_type="-1"`, `output_dir=downloads/ty_le_xac_minh` | `ploaibc="-1", pdonvi_id="284656", ploaihd="-1", ploaidv="-1", pthoigianid="98944548"` |
| Kết quả tiếp thị | `download_kq_tiep_thi_api` | `start_date="01/04/2026", end_date="30/04/2026", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/kq_tiep_thi` | `vngay_bd="01/04/2026", vngay_kt="30/04/2026", vdonvi_id="284656"` |
| Vật tư thu hồi | `download_report_vattu_thuhoi_api` | `start_date="24/11/2025", end_date="25/04/2026", headed=False, session=...` | `unit_id="284656"`, `service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16"`, `vat_tu_ids="1,2,3,4,8,6,5"`, `output_dir=downloads/vat_tu_thu_hoi` | `vttvt="284656", vtungay="24/11/2025", vdenngay="25/04/2026", vdichvuvt_erp="<default_service_ids>", vloaithu="0", vloaibatbuoc="0", vvattu="1,2,3,4,8,6,5", vloaingay="1", vtrangthai="0"` |
| Quyết toán vật tư | `download_quyet_toan_vattu_api` | `start_date="01/04/2026", end_date="30/04/2026", headed=False, session=...` | `unit_id="284656"`, `output_dir=downloads/vat_tu_thu_hoi` | `vttvt="284656", vtungay="01/04/2026", vdenngay="30/04/2026"` |
| Cấu hình tự động PTM | `download_cau_hinh_tu_dong_ptm_api` | `month_id="98944548", headed=False, session=...` | `output_dir=downloads/cau_hinh_tu_dong` | `pdv="1", pthang="98944548"` |
| Cấu hình tự động Thay thế | `download_cau_hinh_tu_dong_thay_the_api` | `month_id="98944548", headed=False, session=...` | `output_dir=downloads/cau_hinh_tu_dong` | `pdv="13", pthang="98944548"` |
| Cấu hình tự động Chi tiết | `download_cau_hinh_tu_dong_chi_tiet_api` | `month_id="98944548", headed=False, session=...` | `output_dir=downloads/cau_hinh_tu_dong` | `pthang="98944548"` |

Trong bảng trên:

- `downloads/...` là thư mục con dưới [api_transition/downloads](/home/vtst/baocaohanoi/api_transition/downloads).
- `<default_service_ids>` là chuỗi mặc định:

```text
1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16
```

## CTS SHC ngày

Task `CTS SHC ngày` dùng [download_cts_gpon_quality_detail_api()](api_transition/cts_api.py) chứ không đi qua `download_with_recipe()`.

### Hàm gọi thực tế từ batch

```python
download_cts_gpon_quality_detail_api(
    report_date="17/04/2026",
    headed=False,
)
```

### Default còn lại của downloader CTS

```text
unit_id="87756"
unit_name=""
output_dir=""
output_name="cts_shc_ngay.xlsx"
exclusive_den_ngay=False
session=None
```

### Payload CTS cuối cùng

Theo [cts_api.py](/home/vtst/baocaohanoi/api_transition/cts_api.py:222), payload export được dựng thành:

```json
{
  "searchType": 2,
  "maDonVi": "",
  "tuNgay": "17/04/2026",
  "denNgay": "17/04/2026",
  "tuthang": "3",
  "denthang": "3",
  "UnitID": "87756",
  "Loss_Max": 30,
  "Loss_Ok": 27,
  "ReportType": 1,
  "Quaterly": 0,
  "unitId": "",
  "Year": 2026,
  "province_code": "",
  "nam": 2026,
  "BeginDate": "2026-04-17",
  "endDate": "2026-04-17",
  "ProvinceCode": 1
}
```

Ghi chú:

- Task này có `use_shared_session=False`, nên không dùng session đăng nhập chung của batch.
- Nếu batch không truyền `report_date`, CTS downloader sẽ tự lấy ngày hôm qua; nhưng ở cấu hình hiện tại batch có truyền rõ `report_date="17/04/2026"`.
