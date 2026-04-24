# SQLite History Summary Per Sheet

Module nay luu du lieu tu workbook processed vao SQLite local theo mo hinh:

- `1 loai bao cao` co `nhieu bang du lieu`
- `1 sheet tong hop` cua workbook processed = `1 bang du lieu rieng`
- khong import raw sheet vao DB

## Nguyen tac hien tai

- Chi import cac `sheet sau xu ly`, khong phu thuoc ten sheet co chua chu `tong_hop` hay khong.
- Loai bo cac sheet raw/note ro rang nhu:
  - `Sheet`, `Sheet1`
  - `Data`, `Data_combined`, `Data_tam_dung`, `Data_khoi_phuc`
  - `chi tiết vật tư`
  - `chi_tiet_chua_khoi_phuc`
  - `thong_bao`
- Moi bao cao chi co `1 snapshot` hop le cho moi cap:
  - `ma_bao_cao`
  - `ngay_du_lieu`
- Chay lai cung ngay se ghi de snapshot cu trong transaction.
- Luon ghi lai file nguon da dung de import:
  - file trong `Processed`
  - file archive trong `ProcessedDaily` neu co

## Bo bang metadata co dinh

- `cau_hinh_import_tong_hop`
- `danh_muc_bao_cao_tong_hop`
- `danh_muc_bang_du_lieu_bao_cao`
- `bao_cao_tong_hop_ngay`
- `tep_nguon_bao_cao_tong_hop`
- `sheet_bao_cao_tong_hop`
- `nhat_ky_nap_tong_hop`

## Cac bang du lieu dong

Importer tu dong tao them cac bang du lieu dang:

- `bao_cao_<report_code_hash>_<sheet_code_hash>_tong_hop`

Moi bang dong:

- thuoc `1 sheet tong hop`
- co mot so cot metadata he thong:
  - `__snapshot_id`
  - `__sheet_id`
  - `__row_num`
  - `__row_hash`
  - `__imported_at`
- phan con lai la `cac cot that` cua sheet processed

Vi du:

- `bao_cao_xac_minh_tam_dung_9fd8adfd_tong_hop_theo_nvkt_c9178fa2_tong_hop`

Trong bang nay, cac cot nghiep vu duoc luu truc tiep nhu:

- `TTVT`
- `DOIVT`
- `NVKT`
- `SỐ PHIẾU XÁC MINH`

khong di qua JSON payload nghiep vu.

## Y nghia tung bang metadata

- `danh_muc_bao_cao_tong_hop`:
  - danh muc workbook processed theo `ma_bao_cao`
- `danh_muc_bang_du_lieu_bao_cao`:
  - registry `ma_bao_cao + ten_sheet_goc -> ten_bang_du_lieu`
- `bao_cao_tong_hop_ngay`:
  - snapshot theo ngay
  - unique theo `ma_bao_cao + ngay_du_lieu`
- `tep_nguon_bao_cao_tong_hop`:
  - ghi lai file nguon processed/daily
- `sheet_bao_cao_tong_hop`:
  - metadata tung sheet da import trong snapshot
  - luu ten sheet, ten bang du lieu, tong so cot, tong so dong, tong so chi tieu
- `nhat_ky_nap_tong_hop`:
  - log moi lan import

## Views

- `v_bao_cao_tong_hop_moi_nhat_theo_ma_bao_cao`
- `v_bao_cao_tong_hop_moi_nhat_toan_bo`
- `v_nhat_ky_nap_tong_hop_gan_nhat`
- `v_tong_hop_bang_du_lieu_theo_bao_cao`
- `v_tien_do_nap_tong_hop`
- `v_tep_nguon_bao_cao_tong_hop_moi_nhat`
- `v_sheet_bao_cao_tong_hop_moi_nhat`
- `v_danh_muc_bang_du_lieu_bao_cao`

## Van hanh

Khoi tao hoac bo sung schema:

```bash
python3 -m api_transition.sqlite_history.init_report_history_db \
  --db-path /tmp/report_history_summary.db
```

Import tu Processed:

```bash
python3 -m api_transition.sqlite_history.import_processed_to_sqlite \
  --db-path /tmp/report_history_summary.db \
  --processed-root /path/to/Processed \
  --archive-root /path/to/ProcessedDaily \
  --snapshot-date 2026-04-21
```

Neu chay qua `full_pipeline.py`, pipeline se:

1. dam bao schema ton tai
2. import lai theo snapshot ngay
3. apply lai views
