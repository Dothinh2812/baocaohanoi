# Tai lieu cau truc du lieu `sa_outate_db`

File nay mo ta cau truc du lieu cua DB theo doi su co SA, de mot trang dashboard khac co the doc du lieu va hien thi.

## 1. Tong quan

- DB file: `database/sa_outate_db.db`
- Muc dich:
  - Luu tung lan do/kiem tra SA.
  - Gom cac lan do thanh 1 `incident` de theo doi vong doi su co.
  - Ho tro dashboard theo doi su co dang ton va thong ke theo thang.

DB nay hien duoc sinh va cap nhat boi:
- `sa_outage_db.py`
- `kiem_tra_mat_huong_sa.py`

## 2. Quy tac nghiep vu

### 2.1 Khi nao mo su co

Mot `incident` moi duoc mo khi:
- SA co ket qua vuot nguong canh bao.
- Va hien tai chua co `incident` nao dang `OPEN` cho chinh `sa_code` do.

### 2.2 Khi nao cap nhat su co dang mo

Neu SA van dang vuot nguong va da co `incident` `OPEN`, he thong:
- khong tao incident moi,
- ma cap nhat `last_seen_at`, `latest_off_count`, `max_off_count`, `latest_total_count`.

### 2.3 Khi nao ket thuc su co

Mot `incident` chi duoc dong khi co event phuc hoi ro rang, gom 2 nhom:

- `NORMAL`
  - SA duoc do lai va so thue bao `OFF` da ve duoi nguong.
  - Hoac sau buoc verify thi khong con du dieu kien coi la outage.

- `PHUC HOI`
  - SA khong con xuat hien trong `candidate_sas` cua batch hien tai.
  - He thong tu sinh 1 event `PHUC HOI` roi dong `incident`.

Dashboard nen hieu:
- `OPEN` = su co dang ton.
- `CLOSED` = su co da ket thuc.

### 2.4 Luu y voi du lieu lich su import tu Excel

Du lieu lich su import tu `sa_outage_log.xlsx` chu yeu chi co:
- `CANH BAO`
- `SKIP`

Neu trong log lich su khong co `NORMAL` hoac `PHUC HOI`, incident lich su se van o trang thai `OPEN`.

Vi vay:
- `source = 'excel_import'`
- `confidence = 'inferred'`

can duoc dashboard hien thi nhu du lieu lich su co do tin cay thap hon du lieu runtime.

## 3. Bang du lieu chinh

### 3.1 `sa_outage_batches`

Luu thong tin tung batch kiem tra.

| Cot | Kieu du lieu | Y nghia |
|---|---|---|
| `batch_id` | TEXT, PK | Ma batch, vi du `BATCH_20260406_001` |
| `started_at` | TEXT datetime | Thoi diem bat dau batch |
| `completed_at` | TEXT datetime, nullable | Thoi diem ket thuc batch |
| `threshold` | INTEGER | Nguong OFF dung de xac dinh outage |
| `candidate_sa_count` | INTEGER | So SA duoc dua vao batch de kiem tra |
| `success_count` | INTEGER | So SA kiem tra thanh cong trong batch |
| `warning_count` | INTEGER | So SA duoc xac dinh la outage trong batch |
| `recovered_count` | INTEGER | So incident duoc dong trong batch |
| `error_count` | INTEGER | So SA loi hoac khong xu ly duoc |
| `notes` | TEXT | Ghi chu tong hop batch |

Use case cho dashboard:
- Hien thi lich su cac lan quet.
- Hien thi so luong su co moi, su co dong, so SA loi trong tung batch.

### 3.2 `sa_outage_events`

Luu moi event phat sinh theo thoi gian. Day la bang log chi tiet nhat.

| Cot | Kieu du lieu | Y nghia |
|---|---|---|
| `id` | INTEGER, PK | Khoa chinh tu tang |
| `event_key` | TEXT, UNIQUE | Khoa duy nhat de chong ghi trung |
| `batch_id` | TEXT, nullable | Batch tao ra event |
| `event_at` | TEXT datetime | Thoi diem event |
| `sa_code` | TEXT | Ma SA, vi du `STY.G51_1/10` |
| `doi_vt` | TEXT | Doi/To ky thuat phu trach |
| `status` | TEXT | Trang thai event |
| `off_count` | INTEGER | So thue bao OFF tai thoi diem event |
| `total_count` | INTEGER | Tong so thue bao tren SA |
| `off_percentage` | REAL | Ty le OFF |
| `notes` | TEXT | Ghi chu event |
| `source` | TEXT | `runtime` hoac `excel_import` |
| `created_at` | TEXT datetime | Thoi diem ghi vao DB |

#### Cac gia tri `status` quan trong

- `CANH BAO`
  - SA dang outage va canh bao duoc gui.

- `SKIP`
  - SA van outage nhung khong gui lai canh bao do throttle.

- `NORMAL`
  - SA duoc coi la da phuc hoi sau khi do/verify.

- `PHUC HOI`
  - SA khong con xuat hien trong batch hien tai va duoc danh dau ket thuc.

Dashboard co the dung bang nay de:
- Ve timeline event.
- Hien thi lich su chi tiet cua 1 SA.
- Debug vi sao incident mo/dong.

### 3.3 `sa_outage_incidents`

Day la bang tong hop vong doi su co. Dashboard theo doi su co nen uu tien bang nay.

| Cot | Kieu du lieu | Y nghia |
|---|---|---|
| `id` | INTEGER, PK | Khoa chinh tu tang |
| `incident_code` | TEXT, UNIQUE | Ma su co |
| `sa_code` | TEXT | Ma SA |
| `doi_vt` | TEXT | Doi/To ky thuat |
| `status` | TEXT | `OPEN` hoac `CLOSED` |
| `started_at` | TEXT datetime | Luc bat dau su co |
| `last_seen_at` | TEXT datetime | Lan cuoi incident duoc cap nhat |
| `ended_at` | TEXT datetime, nullable | Luc ket thuc su co |
| `started_batch_id` | TEXT | Batch mo su co |
| `last_batch_id` | TEXT | Batch cap nhat cuoi |
| `ended_batch_id` | TEXT, nullable | Batch dong su co |
| `start_off_count` | INTEGER | So OFF luc bat dau |
| `latest_off_count` | INTEGER | So OFF lan cap nhat cuoi |
| `max_off_count` | INTEGER | So OFF lon nhat trong suot incident |
| `latest_total_count` | INTEGER | Tong thue bao o lan cap nhat cuoi |
| `latest_off_percentage` | REAL | Ty le OFF o lan cap nhat cuoi |
| `duration_minutes` | INTEGER, nullable | Tong thoi gian su co, chi co khi `CLOSED` |
| `recovery_reason` | TEXT, nullable | Ly do ket thuc |
| `notes` | TEXT | Ghi chu cuoi cung cua incident |
| `source` | TEXT | `runtime` hoac `excel_import` |
| `confidence` | TEXT | `exact` hoac `inferred` |
| `created_at` | TEXT datetime | Luc tao record |
| `updated_at` | TEXT datetime | Luc cap nhat record gan nhat |

#### Y nghia `recovery_reason`

Gia tri thuong gap:
- `below_threshold`: SA da xuong duoi nguong qua lan do moi.
- `verified_off_count_is_zero`: verify lai khong con thue bao OFF.
- `verified_off_count_below_threshold`: verify lai con OFF nhung da duoi nguong.
- `missing_from_current_batch`: SA khong con xuat hien trong danh sach SA can kiem tra cua batch hien tai.
- `normal`: du lieu lich su co event `NORMAL`.
- `phục hồi` hoac `phuc hoi`: du lieu lich su co event phuc hoi.

#### Y nghia `confidence`

- `exact`
  - Incident duoc mo/dong boi luong runtime moi.
  - Do tin cay cao.

- `inferred`
  - Incident duoc dung lai tu log lich su import.
  - Do tin cay thap hon do du lieu cu khong day du.

## 4. View danh cho dashboard

### 4.1 `v_sa_outage_monitoring`

View de hien thi man hinh theo doi su co.

Cot chinh:
- `incident_code`
- `sa_code`
- `doi_vt`
- `trang_thai`
- `started_at`
- `last_seen_at`
- `ended_at`
- `thoi_gian_keo_dai_phut`
- `start_off_count`
- `latest_off_count`
- `max_off_count`
- `latest_total_count`
- `latest_off_percentage`
- `recovery_reason`
- `notes`
- `source`
- `confidence`

Y nghia:
- Neu `ended_at IS NULL`, su co dang ton.
- `thoi_gian_keo_dai_phut`:
  - Neu su co dang ton: tinh tu `started_at` den thoi diem hien tai.
  - Neu da ket thuc: dung `duration_minutes`.

### 4.2 `v_sa_outage_monthly_detail`

View chi tiet su co theo thang.

Use case:
- Bang danh sach su co theo thang.
- Drill-down tu tong hop thang xuong tung incident.

### 4.3 `v_sa_outage_monthly_summary`

View tong hop theo thang.

Cot:
- `thang`
- `tong_su_co`
- `su_co_da_ket_thuc`
- `su_co_dang_ton`
- `tg_xu_ly_tb_phut`
- `tg_xu_ly_lon_nhat_phut`
- `tong_tg_xu_ly_phut`

Use case:
- Card KPI theo thang.
- Bieu do xu huong so luong su co.
- Bieu do thoi gian xu ly trung binh.

## 5. Dashboard nen doc bang nao

### 5.1 Man hinh tong quan

Nen dung:
- `v_sa_outage_monitoring`
- `v_sa_outage_monthly_summary`

### 5.2 Man hinh chi tiet mot su co

Nen dung:
- `sa_outage_incidents` de lay thong tin tong quan incident
- `sa_outage_events` de ve timeline event

### 5.3 Man hinh lich su theo SA

Nen dung:
- `sa_outage_events`
- filter theo `sa_code`

## 6. Goi y truy van cho dashboard

### 6.1 Danh sach su co dang ton

```sql
SELECT
  incident_code,
  sa_code,
  doi_vt,
  started_at,
  last_seen_at,
  thoi_gian_keo_dai_phut,
  latest_off_count,
  max_off_count,
  latest_total_count,
  latest_off_percentage
FROM v_sa_outage_monitoring
WHERE trang_thai = 'Đang tồn'
ORDER BY started_at DESC;
```

### 6.2 Top su co keo dai nhat

```sql
SELECT
  incident_code,
  sa_code,
  doi_vt,
  trang_thai,
  thoi_gian_keo_dai_phut
FROM v_sa_outage_monitoring
ORDER BY thoi_gian_keo_dai_phut DESC
LIMIT 20;
```

### 6.3 Lich su event cua 1 SA

```sql
SELECT
  event_at,
  status,
  off_count,
  total_count,
  off_percentage,
  notes,
  batch_id,
  source
FROM sa_outage_events
WHERE sa_code = ?
ORDER BY event_at DESC;
```

### 6.4 Chi tiet su co theo thang

```sql
SELECT
  thang,
  incident_code,
  sa_code,
  doi_vt,
  trang_thai,
  bat_dau,
  ket_thuc,
  thoi_gian_xu_ly_phut
FROM v_sa_outage_monthly_detail
WHERE thang = '2026-04'
ORDER BY bat_dau DESC;
```

### 6.5 Tong hop theo thang

```sql
SELECT *
FROM v_sa_outage_monthly_summary
ORDER BY thang DESC;
```

## 7. Giai thich cac cot dashboard thuong dung

### `off_luc_bat_dau`

So thue bao OFF o lan dau tien mo incident.

### `off_hien_tai_hoac_ket_thuc`

So thue bao OFF o lan cap nhat cuoi cung:
- neu incident dang mo: day la so OFF hien tai,
- neu incident da dong: day la so OFF tai thoi diem ket thuc.

### `off_lon_nhat`

So thue bao OFF lon nhat tung ghi nhan trong suot incident.

### `tong_thue_bao`

Tong so thue bao cua SA tai lan cap nhat cuoi cung.

### `thoi_gian_keo_dai_phut`

- neu incident dang mo: tinh dong theo thoi gian hien tai,
- neu incident da dong: bang tong thoi gian tu `started_at` den `ended_at`.

## 8. Khuyen nghi hien thi tren dashboard

- Nen uu tien `incident_code` lam khoa chinh tren UI.
- Nen cho phep filter theo:
  - `trang_thai`
  - `doi_vt`
  - `sa_code`
  - `source`
  - `confidence`
  - khoang thoi gian `started_at`

- Nen hien thi nhan canh bao:
  - `source = excel_import` + `confidence = inferred`
  - de phan biet voi du lieu runtime moi.

- Nen co man hinh timeline event cho moi incident:
  - `CANH BAO`
  - `SKIP`
  - `NORMAL`
  - `PHUC HOI`

## 9. Ghi chu tich hop

- DB dung SQLite, co the doc truc tiep bang:
  - backend Python
  - Node.js voi sqlite driver
  - hoac bat ky service API nao co kha nang query SQLite

- Tat ca datetime hien dang luu duoi dang chuoi:
  - format: `YYYY-MM-DD HH:MM:SS`

- Neu dashboard can API, khuyen nghi:
  - tao API `/incidents`
  - tao API `/incidents/:incident_code/events`
  - tao API `/monthly-summary`

## 10. Tom tat bang du lieu nen dung

- Theo doi su co hien tai:
  - `v_sa_outage_monitoring`

- Tong hop thang:
  - `v_sa_outage_monthly_summary`

- Chi tiet tung su co theo thang:
  - `v_sa_outage_monthly_detail`

- Timeline va debug nghiep vu:
  - `sa_outage_events`

