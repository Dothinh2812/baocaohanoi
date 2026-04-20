# Draft Unit Configs

Bo file nay la ban nhap config theo tung don vi, duoc dung tu:

- `ID_dovi/C11.txt`
- `ID_dovi/C11 - SM4 2026.txt`
- `ID_dovi/C12 SM1 2026.txt`
- `ID_dovi/C12 cây đơn vị ttvt -tổ kt.txt`
- `ID_dovi/C13.txt`
- `ID_dovi/C14.txt`
- `ID_dovi/C15.txt`
- `ID_dovi/report_configuration_audit.md`

## Mapping report -> ID family

Nhung report dung `ids.center_id_14`:

- `c11`
- `c12`
- `c13`
- `kpi_nvkt_c11`
- `kpi_nvkt_c12`
- `kpi_nvkt_c13`

Nhung report dung `ids.unit_id_28`:

- `c14`
- `c14_chi_tiet`
- `c11_chi_tiet`
- `c12_chi_tiet_sm1`
- `c12_chi_tiet_sm2`
- `i15`
- `i15_k2`
- `ghtt_sontay`
- `ghtt_nvktdb`
- `xac_minh_tam_dung`
- `phieu_hoan_cong_dich_vu_chi_tiet`
- `tam_dung_khoi_phuc_dich_vu_chi_tiet`
- `tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc`
- `tam_dung_khoi_phuc_dich_vu_tong_hop`
- `ty_le_xac_minh_ttvtkv`
- `ty_le_xac_minh_chi_tiet`
- `kq_tiep_thi`
- `vattu_thu_hoi`
- `quyet_toan_vat_tu`

Nhung report khong dung ID don vi:

- `cau_hinh_tu_dong_ptm`
- `cau_hinh_tu_dong_thay_the`
- `cau_hinh_tu_dong_chi_tiet`

Nhung report chua du du lieu de draft an toan cho tat ca 18 don vi:

- `ghtt_hni`
- `ngung_psc_mytv_thang_t_1_cap_ttvt`
- `ngung_psc_fiber_thang_t_1_cap_ttvt`
- `ngung_psc_fiber_thang_t_1_cap_to`
- `ngung_psc_mytv_thang_t_1_cap_to`
- `onebss_hni_pttb_001`
- `onebss_bc_phieu_ton_dv_chi_tiet_hni`
- `onebss_bc_ton_sua_chua_2026`
- `onebss_bc_chi_tiet_ket_qua_cskh_uc3`

Trang thai draft hien tai:

- cac report da xac dinh ID family thi da co du lieu cho 18 don vi
- nhom chua du mapping chac chan duoc de `enabled: false`
- `team_ids.family_14` va `team_ids.family_28` duoc dua vao config de tham chieu ve sau

## Cac file

- `_template.yaml`: template schema va cach resolve report theo ID family
- `*.yaml`: draft config theo tung don vi
