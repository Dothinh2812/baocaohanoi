# Unit Config Draft Validation

Thoi diem kiem tra: 2026-04-20

## Nguon doi chieu

- `ID_dovi/C11.txt`
- `ID_dovi/C11 - SM4 2026.txt`
- `ID_dovi/C12 SM1 2026.txt`
- `ID_dovi/C12 cây đơn vị ttvt -tổ kt.txt`
- `ID_dovi/C13.txt`
- `ID_dovi/C14.txt`
- `ID_dovi/C15.txt`
- `ID_dovi/report_configuration_audit.md`
- `units.py`

## Ket qua tong quan

- So file config don vi da kiem: `18`
- So don vi co `center_id_14` hop le: `18/18`
- So don vi co `unit_id_28` hop le: `18/18`
- So don vi co `onebss_tt_id = center_id_14`: `18/18`
- Ket qua doi chieu ID tong: `PASS`

## Mapping report -> ID family da du source

### Dung `ids.center_id_14`

- `c11`
- `c12`
- `c13`
- `kpi_nvkt_c11`
- `kpi_nvkt_c12`
- `kpi_nvkt_c13`

### Dung `ids.unit_id_28`

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

### Khong dung ID don vi

- `cau_hinh_tu_dong_ptm`
- `cau_hinh_tu_dong_thay_the`
- `cau_hinh_tu_dong_chi_tiet`

## Nhom chua du du lieu de bat len an toan

Nhung report sau hien dang de `enabled: false` trong draft:

- `ghtt_hni`
- `ngung_psc_mytv_thang_t_1_cap_ttvt`
- `ngung_psc_fiber_thang_t_1_cap_ttvt`
- `ngung_psc_fiber_thang_t_1_cap_to`
- `ngung_psc_mytv_thang_t_1_cap_to`
- `onebss_hni_pttb_001`
- `onebss_bc_phieu_ton_dv_chi_tiet_hni`
- `onebss_bc_ton_sua_chua_2026`
- `onebss_bc_chi_tiet_ket_qua_cskh_uc3`

Ly do:

- Chua co bang mapping day du cho tat ca 18 don vi trong bo file `ID_dovi/`.
- Semantics ID cua cac nhom nay khong trung hoan toan voi hai ho `14xxx` va `28xxxx`.

## Ghi chu ve `team_ids`

`team_ids.family_14` va `team_ids.family_28` da duoc dua vao tung file unit config.

Can luu y:

- So team trong `family_14` va `family_28` khong phai luc nao cung bang nhau.
- Day la khac biet den tu du lieu nguon, khong phai loi draft config.

Vi du:

- `gia_lam`: `family_14 = 4`, `family_28 = 2`
- `hoai_duc`: `family_14 = 6`, `family_28 = 4`
- `hoang_mai`: `family_14 = 6`, `family_28 = 3`
- `phu_xuyen`: `family_14 = 6`, `family_28 = 4`
- `son_tay`: `family_14 = 6`, `family_28 = 4`
- `tay_ho`: `family_14 = 5`, `family_28 = 4`
- `thanh_tri`: `family_14 = 7`, `family_28 = 5`

## Ket luan

Bo draft trong `configs/units/` da dat muc tieu cho vong 1:

- du `18` don vi
- co du `center_id_14` va `unit_id_28`
- co mapping report -> ID family o muc draft
- co team map tham chieu theo hai family
- da khoa an toan cac report chua du du lieu xac minh

Buoc tiep theo hop ly:

1. Viet `config loader`.
2. Cho `batch_download.py` resolve `unit_id` theo `ID family`.
3. Truyen `instance_root` vao toan bo runtime roots.
