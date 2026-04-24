# Hướng dẫn dữ liệu để dựng lại báo cáo "Báo cáo tổng hợp admin"

## 1. Mục tiêu báo cáo

Báo cáo này dùng để tổng hợp số thuê bao `K1`, `K2` mà NVKT xử lý được, theo:

- ngày
- tháng
- từng NVKT
- từng tổ

Nguyên tắc quan trọng nhất:

- Báo cáo **ghi nhận theo NVKT gán trong DB nguồn** (`nvkt_db_normalized`), **không theo người upload**.
- Một bản ghi được tính là "đã xử lý" khi:
  - `k1_status = 'dat'`, hoặc
  - `k2_status = 'dat'`
- Số "cần xử lý" không lấy từ bảng upload nội bộ, mà lấy bằng cách **đếm record trong 2 DB nguồn suy hao K1/K2** theo ngày báo cáo.

## 2. Nguồn dữ liệu cần lấy

### 2.1. DB nội bộ của ứng dụng

File hiện tại:

- `data/nvkt_results.db`

Bảng chính dùng cho báo cáo:

- `processing_results`
- `upload_issue_logs` (chỉ cần nếu muốn dựng thêm phần log lỗi bên dưới báo cáo admin)

Schema tối thiểu của `processing_results`:

| Cột | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `processed_date` | `TEXT` | Ngày ghi nhận kết quả, format `YYYY-MM-DD` |
| `processed_at` | `TEXT` | Thời gian chi tiết |
| `user_name` | `TEXT` | NVKT được ghi nhận để thống kê |
| `uploader_user_name` | `TEXT` | Người upload ảnh, không dùng để cộng số cho báo cáo này |
| `account` | `TEXT` | Account CTS |
| `k1_status` | `TEXT` | Trạng thái K1, giá trị quan trọng là `dat` |
| `k2_status` | `TEXT` | Trạng thái K2, giá trị quan trọng là `dat` |
| `doi_one` | `TEXT` | Tổ của NVKT |
| `nvkt_db_normalized` | `TEXT` | NVKT chuẩn hóa từ DB nguồn |

Lưu ý:

- Khi lưu dữ liệu, app đã cố gắng chuẩn hóa để `user_name` phản ánh `nvkt_db_normalized`.
- Vì vậy khi dựng báo cáo, dùng `user_name` làm khóa NVKT chính.

### 2.2. DB nguồn suy hao K1

File hiện tại:

- `/home/vtst/baocaohanoi/suy_hao_history.db`

Bảng:

- `suy_hao_snapshots`

Các cột cần thiết để tính báo cáo:

| Cột | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `ngay_bao_cao` | `DATE` | Ngày snapshot |
| `account_cts` | `TEXT` | Account CTS |
| `doi_one` | `TEXT` | Tổ |
| `nvkt_db_normalized` | `TEXT` | NVKT chuẩn hóa |

### 2.3. DB nguồn suy hao K2

File hiện tại:

- `/home/vtst/baocaohanoi/suy_hao_history_k2.db`

Bảng:

- `suy_hao_snapshots`

Các cột cần lấy giống K1:

- `ngay_bao_cao`
- `account_cts`
- `doi_one`
- `nvkt_db_normalized`

## 3. Chuẩn hóa dữ liệu trước khi tính

Để ra đúng như màn admin hiện tại, cần áp dụng các quy tắc chuẩn hóa sau:

- Nếu `doi_one` rỗng hoặc chỉ có khoảng trắng thì thay bằng `Chua xac dinh`
- Nếu `user_name` hoặc `nvkt_db_normalized` rỗng hoặc chỉ có khoảng trắng thì thay bằng `Chua xac dinh`
- `report_month` được tách từ ngày theo `SUBSTR(date_value, 1, 7)` tương đương format `YYYY-MM`

## 4. Logic tính đúng như báo cáo admin hiện tại

### 4.1. Chỉ số tổng ngày

Nguồn: `processing_results`

Điều kiện:

- `processed_date = :report_date`

Cách tính:

- `k1_count = SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END)`
- `k2_count = SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END)`
- `total_count = k1_count + k2_count`

### 4.2. Chỉ số tổng tháng

Nguồn: `processing_results`

Điều kiện:

- `SUBSTR(processed_date, 1, 7) = :report_month`

Cách tính giống phần tổng ngày.

### 4.3. Số cần xử lý trong ngày

Nguồn:

- DB K1 `suy_hao_snapshots`
- DB K2 `suy_hao_snapshots`

Điều kiện:

- `ngay_bao_cao = :report_date`

Cách tính:

- DB K1 đếm số record theo ngày để ra `k1_required_count`
- DB K2 đếm số record theo ngày để ra `k2_required_count`
- `total_required_count = k1_required_count + k2_required_count`

### 4.4. Bảng theo cá nhân trong ngày

Nguồn ghép từ 2 phía:

1. Phần "đã xử lý" từ `processing_results`
2. Phần "cần xử lý" từ 2 DB nguồn K1/K2

Khóa ghép:

- `doi_one`
- `user_name`

Điều kiện phần "đã xử lý":

- `processed_date = :report_date`
- chỉ lấy record có `k1_status = 'dat' OR k2_status = 'dat'`

Điều kiện phần "cần xử lý":

- `ngay_bao_cao = :report_date`
- group theo `doi_one`, `nvkt_db_normalized`

Field đầu ra mỗi dòng:

| Field | Ý nghĩa |
| --- | --- |
| `doi_one` | Tổ |
| `user_name` | NVKT |
| `k1_count` | Số K1 đã xử lý |
| `k2_count` | Số K2 đã xử lý |
| `total_count` | `k1_count + k2_count` |
| `k1_required_count` | Tổng K1 cần xử lý từ DB nguồn |
| `k2_required_count` | Tổng K2 cần xử lý từ DB nguồn |
| `total_required_count` | `k1_required_count + k2_required_count` |
| `highlight_warning` | `true` nếu `total_count < 3` và `total_required_count >= 3` |

### 4.5. Bảng theo tổ trong ngày

Nguồn ghép từ 2 phía:

1. Phần "đã xử lý" từ `processing_results`
2. Phần "cần xử lý" từ 2 DB nguồn K1/K2

Khóa ghép:

- `doi_one`

Điều kiện:

- giống phần theo cá nhân, nhưng group theo tổ

Field đầu ra:

- `doi_one`
- `k1_count`
- `k2_count`
- `total_count`
- `k1_required_count`
- `k2_required_count`
- `total_required_count`
- `highlight_warning`

### 4.6. Bảng theo tổ trong tháng

Nguồn:

- `processing_results`

Điều kiện:

- `SUBSTR(processed_date, 1, 7) = :report_month`
- chỉ lấy record có `k1_status = 'dat' OR k2_status = 'dat'`

Group by:

- `doi_one`

Field đầu ra:

- `doi_one`
- `k1_count`
- `k2_count`
- `total_count`

Lưu ý:

- Bảng này hiện tại **không ghép thêm số cần xử lý theo tháng** ở màn admin.

### 4.7. Bảng theo ngày trong tháng

Nguồn:

- `processing_results`

Điều kiện:

- `SUBSTR(processed_date, 1, 7) = :report_month`
- chỉ lấy record có `k1_status = 'dat' OR k2_status = 'dat'`

Group by:

- `processed_date`

Field đầu ra:

- `processed_date`
- `k1_count`
- `k2_count`
- `total_count`

## 5. Team filter đang hoạt động như thế nào

Tham số filter:

- `admin_team_filter`

Giá trị:

- rỗng: tất cả tổ
- có giá trị: chỉ lấy đúng một `doi_one`

Filter này hiện tại chỉ áp vào:

- `daily_user_rows`
- `daily_team_rows`
- `monthly_team_rows`

Filter này **không áp vào**:

- `daily_summary`
- `daily_required_summary`
- `monthly_summary`
- `monthly_day_rows`

Nếu app dashboard mới muốn giống 100% màn admin hiện tại thì giữ nguyên cách này.

## 6. Contract dữ liệu đầu ra nên trả cho dashboard

Nếu làm API riêng để dashboard khác consume, nên trả về đúng shape này:

```json
{
  "selected_date": "2026-03-25",
  "selected_month": "2026-03",
  "selected_team_filter": "",
  "team_options": ["Tổ 1", "Tổ 2"],
  "daily_summary": {
    "k1_count": 0,
    "k2_count": 0,
    "total_count": 0
  },
  "daily_required_summary": {
    "k1_required_count": 0,
    "k2_required_count": 0,
    "total_required_count": 0
  },
  "monthly_summary": {
    "k1_count": 0,
    "k2_count": 0,
    "total_count": 0
  },
  "daily_user_rows": [
    {
      "doi_one": "Tổ 1",
      "user_name": "nguyen van a",
      "k1_count": 1,
      "k2_count": 0,
      "total_count": 1,
      "k1_required_count": 2,
      "k2_required_count": 1,
      "total_required_count": 3,
      "highlight_warning": true
    }
  ],
  "daily_team_rows": [
    {
      "doi_one": "Tổ 1",
      "k1_count": 3,
      "k2_count": 2,
      "total_count": 5,
      "k1_required_count": 6,
      "k2_required_count": 2,
      "total_required_count": 8,
      "highlight_warning": false
    }
  ],
  "monthly_team_rows": [
    {
      "doi_one": "Tổ 1",
      "k1_count": 12,
      "k2_count": 7,
      "total_count": 19
    }
  ],
  "monthly_day_rows": [
    {
      "processed_date": "2026-03-25",
      "k1_count": 3,
      "k2_count": 2,
      "total_count": 5
    }
  ]
}
```

## 7. SQL mẫu để dựng lại báo cáo

### 7.1. Tổng ngày

```sql
SELECT
    COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
    COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
    COALESCE(
        SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
        SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
        0
    ) AS total_count
FROM processing_results
WHERE processed_date = :report_date;
```

### 7.2. Theo cá nhân trong ngày, phần đã xử lý

```sql
SELECT
    COALESCE(NULLIF(TRIM(user_name), ''), 'Chua xac dinh') AS user_name,
    COALESCE(NULLIF(TRIM(doi_one), ''), 'Chua xac dinh') AS doi_one,
    COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
    COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
    COALESCE(
        SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
        SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
        0
    ) AS total_count
FROM processing_results
WHERE processed_date = :report_date
  AND (k1_status = 'dat' OR k2_status = 'dat')
GROUP BY user_name, doi_one;
```

### 7.3. Theo cá nhân trong ngày, phần cần xử lý từ DB nguồn

Chạy ở DB K1:

```sql
SELECT
    COALESCE(NULLIF(TRIM(doi_one), ''), 'Chua xac dinh') AS doi_one,
    COALESCE(NULLIF(TRIM(nvkt_db_normalized), ''), 'Chua xac dinh') AS user_name,
    COUNT(*) AS k1_required_count
FROM suy_hao_snapshots
WHERE ngay_bao_cao = :report_date
GROUP BY doi_one, nvkt_db_normalized;
```

Chạy ở DB K2:

```sql
SELECT
    COALESCE(NULLIF(TRIM(doi_one), ''), 'Chua xac dinh') AS doi_one,
    COALESCE(NULLIF(TRIM(nvkt_db_normalized), ''), 'Chua xac dinh') AS user_name,
    COUNT(*) AS k2_required_count
FROM suy_hao_snapshots
WHERE ngay_bao_cao = :report_date
GROUP BY doi_one, nvkt_db_normalized;
```

Sau đó outer join theo:

- `doi_one`
- `user_name`

## 8. Cách đơn giản nhất nếu app kia không muốn query trực tiếp DB

Ứng dụng hiện tại đã xuất sẵn file:

- `data/nvkt_processing_report.xlsx`

Các sheet hữu ích:

| Sheet | Dùng cho phần nào |
| --- | --- |
| `current_day_by_person` | Bảng theo cá nhân trong ngày |
| `current_day_by_team` | Bảng theo tổ trong ngày |
| `current_month_by_person` | Tổng hợp theo cá nhân trong tháng |
| `current_month_by_team` | Tổng hợp theo tổ trong tháng |
| `daily_by_person` | Dữ liệu theo ngày + cá nhân cho toàn bộ lịch sử |
| `daily_by_team` | Dữ liệu theo ngày + tổ cho toàn bộ lịch sử |
| `monthly_by_person` | Dữ liệu theo tháng + cá nhân |
| `monthly_by_team` | Dữ liệu theo tháng + tổ |
| `success_records` | Record đã xử lý thành công ở mức chi tiết |
| `required_by_person` | Số cần xử lý theo ngày/NVKT/tổ |
| `required_by_team` | Số cần xử lý theo ngày/tháng/tổ |

Header hiện tại của các sheet chính:

- `current_day_by_person`: `processed_date, user_name, doi_one, k1_count, k2_count, total_count, k1_required_count, k2_required_count, total_required_count`
- `current_day_by_team`: `processed_date, doi_one, k1_count, k2_count, total_count, k1_required_count, k2_required_count, total_required_count`
- `current_month_by_team`: `report_month, doi_one, k1_count, k2_count, total_count, k1_required_count, k2_required_count, total_required_count`

Nếu dashboard mới chỉ cần dựng báo cáo nhanh thì đây là phương án tích hợp nhẹ nhất.

## 9. Phần log lỗi nếu muốn dựng full màn admin

Ngoài báo cáo chính, màn admin hiện tại còn có phần log lỗi upload trong ngày.

Nguồn:

- bảng `upload_issue_logs` trong `data/nvkt_results.db`

Summary hiện tại:

- `total_issues`
- `duplicate_count`
- `no_info_count`
- `unmatched_count`
- `needs_review_count`
- `error_count`

Chi tiết mỗi dòng log:

- `created_at`
- `uploader_user_name`
- `attributed_user_name`
- `image_filename`
- `image_gcs_url`
- `processing_status`
- `issue_type`
- `account`
- `detail_message`
- `source_error`

## 10. Kết luận để team dashboard implement

Nếu mục tiêu là dựng lại đúng báo cáo admin hiện tại, app dashboard khác cần lấy tối thiểu 3 nguồn:

1. `processing_results` để biết số K1/K2 đã xử lý.
2. DB nguồn K1 `suy_hao_snapshots` để biết tổng K1 cần xử lý theo ngày.
3. DB nguồn K2 `suy_hao_snapshots` để biết tổng K2 cần xử lý theo ngày.

Khóa thống kê cần thống nhất ở mọi nơi:

- ngày: `processed_date` hoặc `ngay_bao_cao`
- tháng: `YYYY-MM`
- tổ: `doi_one`
- NVKT: `user_name` hoặc `nvkt_db_normalized`

Nếu cần đồng bộ nhanh, có thể bỏ qua query trực tiếp DB và dùng luôn file `data/nvkt_processing_report.xlsx`.
