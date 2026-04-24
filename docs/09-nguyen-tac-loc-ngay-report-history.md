# Nguyên tắc lọc ngày cho `report_history.db`

## Mục tiêu

Chuẩn hóa cách mọi route trong `dashv4` hiển thị dữ liệu theo ngày khi cùng đọc từ:

`/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db`

Mục tiêu cuối cùng:
- mọi trang dùng chung một nguyên tắc lọc ngày
- cùng một trang chỉ dùng một `selected_date`
- không còn tình trạng route này lấy `latest`, route khác lấy snapshot theo ngày mà không công khai

## Phạm vi áp dụng

Áp dụng cho:
- mọi route API/page đọc `report_history.db`
- mọi repository mới viết cho `dashv4`
- mọi trang nhiều bảng nhưng cùng một nguồn dữ liệu theo ngày

Chưa áp dụng bắt buộc cho:
- route còn đọc Excel cũ
- route đọc DB khác ngoài `report_history.db`
- route legacy đang `501`

## Nguồn ngày chuẩn

Ngày chuẩn duy nhất để hiển thị dữ liệu là:

`bao_cao_tong_hop_ngay.ngay_du_lieu`

Không dùng làm ngày hiển thị chính:
- `__imported_at`
- `thoi_gian_tao`
- `thoi_gian_cap_nhat`
- thời điểm request frontend/backend

Giải thích:
- `ngay_du_lieu` là ngày nghiệp vụ của snapshot đã import
- `__imported_at` chỉ là dấu vết kỹ thuật lúc nạp DB

## Mô hình dữ liệu chuẩn để lọc ngày

Mọi bảng raw được import từ `api_transition` đều phải được truy nguyên qua chuỗi:

`bang_raw.__sheet_id -> sheet_bao_cao_tong_hop.id -> bao_cao_tong_hop_ngay.id`

Chuỗi metadata chuẩn:
- `bao_cao_tong_hop_ngay`
  chứa `ma_bao_cao`, `ngay_du_lieu`, `tu_ngay`, `den_ngay`, `thang_bao_cao`, `nam_bao_cao`
- `sheet_bao_cao_tong_hop`
  chứa `ten_sheet`, `ma_sheet`, `ten_bang_du_lieu`
- bảng raw import
  chứa `__snapshot_id`, `__sheet_id`, `__row_num`, `__row_hash`, `__imported_at`

## Nguyên tắc ẩn cột kỹ thuật trên dashboard

Các cột kỹ thuật phục vụ import/snapshot không được hiển thị ra dashboard.

Danh sách cột phải ẩn mặc định:
- `id`
- `__snapshot_id`
- `__sheet_id`
- `__row_num`
- `__row_hash`
- `__imported_at`

Ý nghĩa:
- các cột này chỉ phục vụ truy nguyên dữ liệu, join metadata, debug import
- không phải cột nghiệp vụ để người dùng dashboard theo dõi

Quy tắc thực thi:
- repository vẫn được phép query các cột này khi cần join
- nhưng trước khi trả payload cho frontend phải loại bỏ chúng
- việc ẩn phải được thực hiện tập trung ở tầng serializer/payload builder, không để từng route tự nhớ xóa

## Nguyên tắc truy vấn

### 1. Route có hỗ trợ lọc ngày không được đọc trực tiếp view `..._moi_nhat`

Không dùng:
- `v_*_moi_nhat`
- `v_bao_cao_tong_hop_moi_nhat_theo_ma_bao_cao`
- các view “latest snapshot” khác

trừ khi:
- route được đánh dấu fallback tạm thời
- hoặc dữ liệu nguồn chưa có đủ snapshot lịch sử

### 2. Route hỗ trợ lọc ngày phải query từ bảng raw hoặc view đã gắn ngày

Ưu tiên:
1. bảng raw + join metadata
2. view theo ngày do team tự tạo, nhưng view đó phải lộ rõ `ngay_du_lieu`

### 3. Một trang, một ngày

Nếu một page có nhiều bảng/section:
- tất cả cùng dùng một `selected_date`
- không để mỗi section tự chọn ngày mới nhất riêng

### 4. Không fallback ngầm

Nếu người dùng truyền `date` mà ngày đó không có dữ liệu:
- không tự nhảy sang ngày khác
- phải trả payload rỗng hoặc thông báo rõ
- phải trả danh sách ngày khả dụng để frontend cho phép chọn lại

## SQL mẫu chuẩn

Ví dụ đọc bảng `chi_tieu_c_c1_1_report_th_c1_1` theo ngày:

```sql
SELECT
    t.*,
    b.ngay_du_lieu
FROM chi_tieu_c_c1_1_report_th_c1_1 t
JOIN sheet_bao_cao_tong_hop s
  ON s.id = t.__sheet_id
JOIN bao_cao_tong_hop_ngay b
  ON b.id = s.bao_cao_tong_hop_ngay_id
WHERE b.ma_bao_cao = 'chi_tieu_c_c1_1_report'
  AND s.ten_bang_du_lieu = 'chi_tieu_c_c1_1_report_th_c1_1'
  AND b.ngay_du_lieu = ?
ORDER BY t."Đơn vị"
```

Ví dụ lấy các ngày khả dụng của một `ma_bao_cao`:

```sql
SELECT DISTINCT ngay_du_lieu
FROM bao_cao_tong_hop_ngay
WHERE ma_bao_cao = ?
  AND trang_thai_nap IN ('thanh_cong', 'khong_co_sheet_tong_hop')
ORDER BY ngay_du_lieu DESC
```

Ví dụ lấy ngày mới nhất hợp lệ:

```sql
SELECT MAX(ngay_du_lieu) AS latest_available_date
FROM bao_cao_tong_hop_ngay
WHERE ma_bao_cao = ?
  AND trang_thai_nap IN ('thanh_cong', 'khong_co_sheet_tong_hop')
```

## Contract API chuẩn cho route theo ngày

Mọi route dùng `report_history.db` và hỗ trợ ngày phải nhận:
- query param `date=YYYY-MM-DD`

Mọi payload trả về phải có tối thiểu:
- `selected_date`
- `latest_available_date`
- `available_dates`
- `file_info` nếu phù hợp
- `sheets` hoặc `sheet`

Ví dụ:

```json
{
  "selected_date": "2026-04-22",
  "latest_available_date": "2026-04-22",
  "available_dates": ["2026-04-22", "2026-04-21", "2026-04-20"],
  "file_info": {
    "name": "report_history.db"
  },
  "sheets": {
    "example_table": {
      "columns": [],
      "data": []
    }
  }
}
```

## Contract UI chuẩn cho page theo ngày

Mỗi page dùng dữ liệu theo ngày phải có:
- 1 date picker chung ở đầu page
- hiển thị `selected_date`
- nếu đang xem ngày cũ hơn latest thì phải thấy rõ

Hành vi chuẩn:
- nếu URL không có `?date=...`, backend tự resolve ngày mới nhất và frontend hiển thị lại rõ
- nếu đổi ngày, toàn page reload cùng ngày đó
- nếu chia sẻ URL có `?date=...`, người khác mở ra phải thấy cùng snapshot ngày đó

## Mapping route và báo cáo

Mỗi route phải khai báo tập trung mapping giữa:
- `route/page`
- `ma_bao_cao`
- `ten_bang_du_lieu`
- chiến lược `order_by`

Ví dụ:

```python
REPORT_DATE_BINDINGS = {
    "chatluong": [
        {
            "report_code": "chi_tieu_c_c1_1_report",
            "table_name": "chi_tieu_c_c1_1_report_th_c1_1",
            "order_by": '"Đơn vị"'
        },
        {
            "report_code": "chi_tieu_c_c1_1_chitiet_report",
            "table_name": "chi_tieu_c_c1_1_chitiet_report_chi_tiet",
            "order_by": '"TEN_DOI", "NVKT"'
        }
    ]
}
```

Không cho phép:
- hardcode `ma_bao_cao` rải rác trong nhiều blueprint
- mỗi route tự nghĩ ra cách resolve ngày riêng

## Repository dùng chung cần tạo

Đề xuất thêm module:

`repositories/report_history_by_date.py`

Các hàm chuẩn:
- `get_available_dates(report_codes)`
- `get_latest_available_date(report_codes)`
- `resolve_selected_date(requested_date, report_codes)`
- `load_table_by_date(report_code, table_name, selected_date, order_by=None)`
- `load_many_tables_by_date(bindings, selected_date)`

Nguyên tắc trong module này:
- chỉ đọc SQLite read-only như các repository khác
- kiểm tra format ngày tập trung
- xử lý trường hợp nhiều `report_code` trên cùng một page

## Chiến lược resolve ngày cho page nhiều nguồn

Nếu page dùng nhiều `ma_bao_cao`:
- ngày mặc định phải là ngày lớn nhất mà tất cả nguồn cùng có dữ liệu
- `available_dates` nên là giao của các ngày khả dụng

Không được:
- lấy ngày mới nhất của từng bảng riêng rồi ghép lên cùng một page

## Phân loại route theo mức sẵn sàng

### Mức A: sẵn sàng lọc ngày ngay

Điều kiện:
- bảng raw có `__sheet_id`
- có record tương ứng trong `sheet_bao_cao_tong_hop`
- có `bao_cao_tong_hop_ngay.ngay_du_lieu`

### Mức B: cần view theo ngày hoặc adapter thêm

Điều kiện:
- logic hiện tại đang bám vào view business tổng hợp
- nhưng bảng raw phía dưới vẫn còn snapshot theo ngày

### Mức C: chưa đủ điều kiện

Điều kiện:
- route vẫn dùng Excel/file ngoài
- hoặc bảng đang dùng không truy nguyên được `__sheet_id`

## Thứ tự triển khai đề xuất

### Pha 1

Làm hạ tầng chung:
- `repositories/report_history_by_date.py`
- helper resolve ngày
- helper contract API chuẩn
- partial UI date picker

### Pha 2

Áp dụng cho trang `chatluong`:
- vì mapping `ma_bao_cao -> ten_bang_du_lieu` rõ nhất
- ít phụ thuộc adapter cũ

### Pha 3

Áp dụng cho:
- `tam-dung-khoi-phuc`
- `thuc-tang-ngung-psc`

### Pha 4

Mở rộng sang:
- KPI
- GHTT
- Cấu hình tự động
- các route SQLite còn lại

## Quy tắc cập nhật tài liệu khi triển khai

Đây là quy tắc bắt buộc.

Mỗi lần triển khai xong một page hoặc một nhóm route theo ngày, phải cập nhật doc cùng commit hoặc cùng turn làm việc:

### 1. Cập nhật `docs/08-trang-thai-thuc-thi.md`

Phải bổ sung:
- route nào đã hỗ trợ `date`
- route nào vẫn đang dùng `latest`
- route nào đang fallback

### 2. Cập nhật `docs/04-mapping-route-va-du-lieu.md`

Phải bổ sung:
- route
- `ma_bao_cao`
- `ten_bang_du_lieu`
- đã hỗ trợ `date` hay chưa

### 3. Nếu thay đổi nguyên tắc chung

Phải cập nhật lại chính tài liệu này:

`docs/09-nguyen-tac-loc-ngay-report-history.md`

### 4. Nếu phát hiện DB không đáp ứng giả định

Phải ghi rõ vào doc:
- bảng nào không có `__sheet_id`
- route nào chưa thể hỗ trợ `date`
- cần bổ sung gì ở `api_transition`

## Definition of done cho một route hỗ trợ ngày

Một route chỉ được coi là hoàn tất hỗ trợ ngày khi đủ tất cả điều kiện:
- có nhận `?date=YYYY-MM-DD`
- có resolve ngày mặc định chuẩn
- có trả `selected_date`
- không còn phụ thuộc view `..._moi_nhat`
- có smoke test cơ bản
- đã cập nhật doc trạng thái + mapping route

## Cấm làm

- Không thêm filter ngày riêng từng bảng trên cùng một page.
- Không dùng `__imported_at` thay cho `ngay_du_lieu`.
- Không silently fallback sang latest khi người dùng chọn ngày không có dữ liệu.
- Không triển khai xong code mà bỏ qua cập nhật doc.
