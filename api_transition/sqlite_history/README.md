# SQLite History Module

Module nay dung de luu lich su bao cao processed vao SQLite local trong `api_transition`, thay cho viec chi giu 1 file Excel bi ghi de.

## 1. Muc tieu

- Luu du lieu theo `ngay_du_lieu` de dashboard co the doc lai lich su.
- Cho phep chuong trinh chay nhieu lan trong cung 1 ngay.
- Khi chay lai trong cung ngay, du lieu cung `ma_bao_cao + ngay_du_lieu` se bi ghi de an toan.
- Giu ten bang, ten cot bang tieng Viet khong dau, de hieu theo nghiep vu bao cao.
- Cho phep bao cao them cot moi hoac doi cau truc ma khong vo schema ngay lap tuc.

## 2. Pham vi module

Tat ca phan lien quan SQLite local deu dat trong `api_transition`:

- schema: [report_history_schema.sql](/home/vtst/baocaohanoi/api_transition/sqlite_history/report_history_schema.sql)
- view SQL: [report_history_views.sql](/home/vtst/baocaohanoi/api_transition/sqlite_history/report_history_views.sql)
- script khoi tao: [init_report_history_db.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/init_report_history_db.py)
- script import: [import_processed_to_sqlite.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/import_processed_to_sqlite.py)
- script apply view: [apply_report_history_views.py](/home/vtst/baocaohanoi/api_transition/sqlite_history/apply_report_history_views.py)
- tai lieu cho app doc du lieu: [CONSUMER_GUIDE.md](/home/vtst/baocaohanoi/api_transition/sqlite_history/CONSUMER_GUIDE.md)
- file DB mac dinh: [report_history.db](/home/vtst/baocaohanoi/api_transition/report_history.db)

Module hien tai da co du 3 phan:

- `schema`
- `khoi tao DB`
- `importer` tu `Processed` vao SQLite
- `view layer` cho dashboard va app doc du lieu

## 3. Nhung gi da lam

Da hoan thanh cac buoc sau:

1. Chon huong `1 SQLite DB` thay vi tach nhieu database.
2. Chot quy tac khoa nghiep vu chinh:
   - `ma_bao_cao + ngay_du_lieu` la 1 snapshot hop le.
   - chay lai cung ngay thi ghi de.
   - sang ngay moi thi tao snapshot moi.
3. Tao schema nen:
   - `danh_muc_bao_cao`
   - `bao_cao_ngay`
   - `sheet_bao_cao`
   - `dong_bao_cao_goc`
   - `tep_luu_tru_bao_cao`
   - `nhat_ky_nap_bao_cao`
   - `danh_muc_don_vi`
   - `danh_muc_nhan_vien`
4. Tao schema nghiep vu rieng cho tung nhom bao cao:
   - `c11`, `c12`, `c13`, `c14`
   - `ghtt`
   - `kpi_nvkt`
   - `ket_qua_tiep_thi`
   - `hoan_cong/ngung_psc/khoi_phuc`
   - `thuc_tang`
   - `xac_minh`
   - `cau_hinh_tu_dong`
   - `vat_tu`
5. Bat cac che do SQLite phu hop cho file local:
   - `PRAGMA foreign_keys = ON`
   - `PRAGMA journal_mode = WAL`
   - `PRAGMA synchronous = NORMAL`
6. Viet importer `import_processed_to_sqlite.py`:
   - doc tat ca workbook trong `api_transition/Processed`
   - luu metadata vao `danh_muc_bao_cao`, `bao_cao_ngay`, `sheet_bao_cao`
   - luu tung dong goc vao `dong_bao_cao_goc`
   - map cac sheet processed on dinh vao bang nghiep vu
   - ghi log vao `nhat_ky_nap_bao_cao`
   - ho tro `ghi de cung ngay` theo `ma_bao_cao + ngay_du_lieu`
7. Tao DB that va kiem tra thanh cong:
   - `38` bang
   - `38` index
   - `71` view
8. Da test import full `33` workbook tren DB test.
9. Da xac nhan chay lai cung ngay khong tao duplicate.
10. Da tao bo `view` cho:
   - view quan tri / helper
   - view lich su va moi nhat cho tung nhom nghiep vu
   - view tong hop cho dashboard

## 4. Quy tac ghi de cung ngay

Day la quy tac quan trong nhat cua module:

- Cung `ma_bao_cao`, cung `ngay_du_lieu`: ghi de snapshot cu.
- Khac `ngay_du_lieu`: tao snapshot moi.
- Ghi de chi ap dung theo tung bao cao, khong ghi de tat ca bao cao trong ngay.

Vi du:

- Chay lai `c11` ngay `2026-04-18`: chi du lieu `c11` cua ngay `2026-04-18` bi thay.
- `ghtt` cung ngay van giu nguyen neu khong import lai.

Importer hien tai da thuc hien theo transaction:

1. Xac dinh `bao_cao_ngay` theo `ma_bao_cao + ngay_du_lieu`.
2. Neu da ton tai thi dung lai `id` hien co.
3. Xoa snapshot cu bang cach `DELETE` ban ghi `bao_cao_ngay` trong transaction.
4. Chen lai snapshot moi va toan bo du lieu con.
5. Commit.
6. Neu loi thi rollback, du lieu cu van con.

Khong duoc xoa du lieu cu ngoai transaction.

## 5. Cach to chuc du lieu

Schema duoc tach thanh 3 lop:

### 5.1. Lop quan ly snapshot

- `danh_muc_bao_cao`: danh muc bao cao co dinh.
- `bao_cao_ngay`: moi ngay, moi bao cao 1 dong.
- `nhat_ky_nap_bao_cao`: log moi lan import, ke ca that bai.
- `tep_luu_tru_bao_cao`: thong tin file archive/processsed.

### 5.2. Lop du lieu goc

- `sheet_bao_cao`: metadata tung sheet.
- `dong_bao_cao_goc`: luu tung dong goc dang JSON de tranh vo schema khi bao cao doi cot.

### 5.3. Lop du lieu nghiep vu

Moi nhom bao cao co bang rieng, vi du:

- `c11_tong_hop`
- `ghtt_don_vi`
- `hoan_cong_fiber`
- `xac_minh_chi_tiet`
- `cau_hinh_tu_dong_tong_hop`
- `quyet_toan_vat_tu`

Cac bang nghiep vu deu tham chieu ve `bao_cao_ngay_id`.

## 6. Nguyen tac dat ten bang va cot

- Dung `snake_case`.
- Dung tieng Viet khong dau.
- Uu tien dung tu theo y nghia bao cao, khong dat ten ky thuat mo ho.
- Giu cac viet tat nghiep vu quen thuoc neu can:
  - `ttvt`
  - `nvkt`
  - `psc`
  - `mytv`
  - `bsc`
  - `ghtt`

Vi du:

- `Mã thuê bao` -> `ma_thue_bao`
- `Đội Viễn thông` -> `doi_vien_thong`
- `Tỷ lệ hài lòng kỹ thuật` -> `ty_le_hai_long_ky_thuat`

## 7. Xu ly khi bao cao doi cot hoac them bao cao moi

Day la tinh huong da duoc tinh truoc:

- Neu bao cao them cot moi nhung chua can dung ngay:
  - luu phan chua map vao `du_lieu_bo_sung_json`
- Neu bao cao moi duoc them:
  - bo sung `ma_bao_cao` trong `danh_muc_bao_cao`
  - tao bang nghiep vu moi neu can
  - neu chua on dinh schema, van co the luu tam o `dong_bao_cao_goc`
- Neu cot cu doi ten:
  - giu mapping trong importer
  - khong doi schema vo toc neu dashboard chua can

Nguyen tac bao tri:

- schema chi sua khi can on dinh nghiep vu
- thay doi nho, chua chac chan thi dua vao `du_lieu_bo_sung_json`
- uu tien backward-compatible de dashboard khong vo

## 8. Cac file va lenh van hanh

Khoi tao moi DB:

```bash
python3 api_transition/sqlite_history/init_report_history_db.py --reset
```

Tao DB test o vi tri khac:

```bash
python3 api_transition/sqlite_history/init_report_history_db.py \
  --db-path /tmp/report_history_test.db \
  --reset
```

Import toan bo `Processed` vao DB mac dinh:

```bash
python3 api_transition/sqlite_history/import_processed_to_sqlite.py \
  --snapshot-date 2026-04-18
```

Import nhung khong tao archive `ProcessedDaily`:

```bash
python3 api_transition/sqlite_history/import_processed_to_sqlite.py \
  --snapshot-date 2026-04-18 \
  --skip-archive
```

Import vao DB test:

```bash
python3 api_transition/sqlite_history/import_processed_to_sqlite.py \
  --db-path /tmp/report_history_test.db \
  --snapshot-date 2026-04-18 \
  --skip-archive
```

Dry-run:

```bash
python3 api_transition/sqlite_history/import_processed_to_sqlite.py \
  --snapshot-date 2026-04-18 \
  --dry-run \
  --json
```

`--dry-run` chi parse va thong ke, khong ghi SQLite va khong copy file sang `ProcessedDaily`.

Import lai 1 nhom bao cao:

```bash
python3 api_transition/sqlite_history/import_processed_to_sqlite.py \
  --snapshot-date 2026-04-18 \
  --skip-archive \
  --path-contains "chi_tieu_c/c1.1 report_processed.xlsx"
```

Apply lai bo view cho DB hien co:

```bash
python3 api_transition/sqlite_history/apply_report_history_views.py
```

Apply view cho DB khac:

```bash
python3 api_transition/sqlite_history/apply_report_history_views.py \
  --db-path /tmp/report_history_test.db
```

Kiem tra nhanh so bang va index:

```bash
sqlite3 api_transition/report_history.db \
  "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'; \
   SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name NOT LIKE 'sqlite_%';"
```

## 9. Quy tac bao tri khi sua schema

Khi can them bang hoac cot:

1. Sua [report_history_schema.sql](/home/vtst/baocaohanoi/api_transition/sqlite_history/report_history_schema.sql).
2. Tao lai DB test bang `--db-path /tmp/... --reset`.
3. Kiem tra bang/index tao du.
4. Kiem tra view tao du va query mau chay duoc.
5. Chay dry-run hoac import vao DB test.
6. Neu thay doi lien quan importer/view, cap nhat tai lieu nay.
7. Chi khi xac nhan on dinh moi ap dung vao `api_transition/report_history.db`.

Neu sau nay can migration that su, co the tach them:

- `sqlite_history/migrations/`
- script `migrate_report_history_db.py`

Hien tai chua can, vi module dang o giai doan importer dau tien.

## 10. Viec tiep theo da du kien

Cac buoc se lam tiep theo:

1. Bo sung mapping cho cac workbook hien dang chi luu raw:
   - `tam_dung_khoi_phuc_dich_vu_*_combined`
   - `ty_le_xac_minh_*_ttvtkv`
   - `ngung_psc_*_thang_t-1_*`
2. Them co che archive file theo ngay o luong van hanh that, khong chi o script test.
3. Them view query san cho dashboard local.
4. Neu can, tao API doc tu SQLite de app khac goi.
5. Toi uu them view xu huong 7 ngay / 30 ngay neu dashboard can.

## 11. Ghi chu van hanh

- Khong dung ten file workbook de suy ra ngay, vi nhieu file processed hien dang co ten tinh.
- `ngay_du_lieu` phai duoc truyen ro rang tu tham so chay.
- `report_history.db` la nguon doc cho dashboard local.
- File Excel processed van nen giu de doi soat va debug.
- Mac dinh script se copy file sang `ProcessedDaily/YYYY-MM-DD/...`. Dung `--skip-archive` neu chi muon nap SQLite.
