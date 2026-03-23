# XỬ LÝ KHI CHẠY NHIỀU LẦN TRONG 1 NGÀY

## Vấn đề

Khi chương trình `baocaohanoi.py` tải báo cáo I1.5 nhiều lần trong cùng 1 ngày, cần xử lý đúng để:
- Không bị duplicate dữ liệu trong database
- Không tính sai số ngày liên tục (`so_ngay_lien_tuc`)
- Không tính sai biến động (TĂNG/GIẢM/VẪN CÒN)

---

## Giải pháp: Version V2 (Tự động phát hiện)

### Logic hoạt động:

**Lần 1 trong ngày (VD: 8h sáng):**
```
✓ Kiểm tra DB: Chưa có dữ liệu ngày hôm nay
✓ So sánh với ngày hôm qua
✓ Tính biến động: TĂNG/GIẢM/VẪN CÒN
✓ Lưu đầy đủ vào database
✓ Tạo file Excel với các sheet mới
```

**Lần 2+ trong ngày (VD: 14h chiều):**
```
⚠️  Phát hiện: Đã có dữ liệu ngày hôm nay trong DB
✓ BỎ QUA lưu database (tránh ghi đè)
✓ ĐỌC dữ liệu biến động từ DB (lần xử lý đầu tiên)
✓ Chỉ tạo file Excel (dựa trên data từ DB)
✓ Kết quả nhất quán với lần 1
```

---

## Cách sử dụng

### 1. Sử dụng từ baocaohanoi.py (mặc định)

```python
# Trong baocaohanoi.py
from c1_process import process_I15_report

# Chạy bình thường - tự động detect
process_I15_report()
```

**Output lần 1:**
```
================================================================================
BẮT ĐẦU XỬ LÝ BÁO CÁO I1.5 (VỚI TRACKING LỊCH SỬ V2)
================================================================================
...
✓ Đang lưu snapshot ngày 2025-11-05 vào database...
✅ Đã lưu 263 bản ghi vào snapshots
```

**Output lần 2:**
```
================================================================================
BẮT ĐẦU XỬ LÝ BÁO CÁO I1.5 (VỚI TRACKING LỊCH SỬ V2)
================================================================================
...
⚠️  ĐÃ CÓ DỮ LIỆU NGÀY 2025-11-05 TRONG DATABASE (263 bản ghi)
⚠️  BỎ QUA lưu database để tránh trùng lặp và sai số liệu
✓  Chỉ xử lý và tạo file Excel output
```

---

### 2. Force Update (Ghi đè dữ liệu cũ)

Nếu file I1.5 lần đầu bị sai và muốn tải lại:

```python
# Chạy với force_update=True
from c1_process import process_I15_report

process_I15_report(force_update=True)
```

**Output:**
```
⚠️  ĐÃ CÓ DỮ LIỆU NGÀY 2025-11-05 (263 bản ghi)
✓  FORCE_UPDATE=True → Sẽ ghi đè dữ liệu cũ
...
✓ Đang lưu snapshot ngày 2025-11-05 vào database...
✅ Đã lưu 263 bản ghi vào snapshots (đã ghi đè)
```

---

### 3. Xóa dữ liệu ngày cũ (thủ công)

Nếu muốn xóa và import lại:

```python
import sqlite3

# Xóa toàn bộ dữ liệu ngày 05/11/2025
conn = sqlite3.connect('suy_hao_history.db')
cursor = conn.cursor()

ngay_can_xoa = '2025-11-05'

cursor.execute("DELETE FROM suy_hao_snapshots WHERE ngay_bao_cao = ?", (ngay_can_xoa,))
cursor.execute("DELETE FROM suy_hao_daily_changes WHERE ngay_bao_cao = ?", (ngay_can_xoa,))
cursor.execute("DELETE FROM suy_hao_daily_summary WHERE ngay_bao_cao = ?", (ngay_can_xoa,))

# Reset tracking cho các thuê bao của ngày đó
cursor.execute("""
    UPDATE suy_hao_tracking
    SET trang_thai = 'DA_HET_SUY_HAO'
    WHERE ngay_thay_cuoi_cung = ?
""", (ngay_can_xoa,))

conn.commit()
conn.close()

print(f"✅ Đã xóa dữ liệu ngày {ngay_can_xoa}")
print("✓ Bây giờ có thể chạy lại process_I15_report()")
```

---

## So sánh 3 phiên bản

| Tính năng | V1 (cũ) | V2 (hiện tại) |
|-----------|---------|---------------|
| **Chạy lần đầu** | ✅ Lưu DB đầy đủ | ✅ Lưu DB đầy đủ |
| **Chạy lần 2+ cùng ngày** | ⚠️ Ghi đè DB | ✅ Bỏ qua, đọc từ DB |
| **`so_ngay_lien_tuc`** | ❌ Tăng sai | ✅ Chính xác |
| **Biến động** | ❌ Tính lại (giống lần 1) | ✅ Dùng kết quả lần 1 |
| **Force update** | ❌ Không có | ✅ Có tham số `force_update` |
| **Thông báo rõ ràng** | ❌ Không | ✅ Có cảnh báo + gợi ý |

---

## Kịch bản sử dụng thực tế

### Kịch bản 1: Quy trình bình thường

**8h sáng:**
1. Tải I1.5 report từ hệ thống VNPT
2. Chạy `baocaohanoi.py`
3. Hệ thống lưu vào DB + tạo Excel
4. Gửi file Excel cho lãnh đạo

**14h chiều (phát hiện file sáng bị thiếu dữ liệu):**
1. Tải lại I1.5 report (đã đầy đủ)
2. Chạy lại `baocaohanoi.py`
3. Hệ thống BỎ QUA DB, chỉ tạo Excel mới
4. File Excel có dữ liệu mới nhưng biến động vẫn đúng (từ lần 8h sáng)

**Vấn đề:** File Excel lúc 14h có dữ liệu đầy đủ nhưng biến động lại dựa trên file lúc 8h (thiếu)

**Giải pháp:** Sử dụng `force_update=True`

---

### Kịch bản 2: Cần cập nhật lại dữ liệu

**Tình huống:** File I1.5 lúc 8h sáng bị thiếu 50 thuê bao, lúc 14h mới tải được file đầy đủ

**Cách xử lý:**

```python
# Trong baocaohanoi.py, thêm option
from c1_process import process_I15_report

# Lần 1: 8h sáng (file thiếu)
process_I15_report()  # Lưu 200 thuê bao vào DB

# Lần 2: 14h chiều (file đầy đủ) - BẮT BUỘC force update
process_I15_report(force_update=True)  # Ghi đè 250 thuê bao vào DB
```

**Lưu ý:** Khi `force_update=True`, hệ thống sẽ:
- Xóa toàn bộ dữ liệu ngày hôm nay trong DB
- Tính toán lại biến động hoàn toàn
- Lưu lại với số liệu mới

---

### Kịch bản 3: Test/Debug

**Tình huống:** Đang phát triển, cần chạy nhiều lần để test

```python
# Option 1: Luôn force update
process_I15_report(force_update=True)

# Option 2: Xóa DB trước mỗi lần test
import sqlite3
conn = sqlite3.connect('suy_hao_history.db')
conn.execute("DELETE FROM suy_hao_snapshots WHERE ngay_bao_cao = '2025-11-05'")
conn.execute("DELETE FROM suy_hao_daily_changes WHERE ngay_bao_cao = '2025-11-05'")
conn.execute("DELETE FROM suy_hao_daily_summary WHERE ngay_bao_cao = '2025-11-05'")
conn.commit()
conn.close()

process_I15_report()
```

---

## Kiểm tra dữ liệu trong database

```python
import sqlite3
import pandas as pd

conn = sqlite3.connect('suy_hao_history.db')

# Xem các ngày đã có dữ liệu
df = pd.read_sql_query("""
    SELECT ngay_bao_cao, COUNT(*) as so_luong
    FROM suy_hao_snapshots
    GROUP BY ngay_bao_cao
    ORDER BY ngay_bao_cao DESC
    LIMIT 10
""", conn)

print("Dữ liệu 10 ngày gần nhất:")
print(df)

# Xem chi tiết ngày hôm nay
ngay_hom_nay = '2025-11-05'
df_detail = pd.read_sql_query(f"""
    SELECT * FROM suy_hao_daily_summary
    WHERE ngay_bao_cao = '{ngay_hom_nay}'
""", conn)

print(f"\nBiến động ngày {ngay_hom_nay}:")
print(df_detail)

conn.close()
```

---

## Lưu ý quan trọng

1. **Mặc định an toàn:** Version V2 mặc định BỎ QUA lưu DB nếu đã có dữ liệu
   - ✅ Tránh ghi đè nhầm
   - ✅ Tránh tính sai số liệu
   - ✅ Dữ liệu nhất quán

2. **Khi nào dùng force_update:**
   - File lần đầu bị sai/thiếu dữ liệu
   - Cần cập nhật lại hoàn toàn
   - Test/Debug

3. **Backup database:**
   ```bash
   # Backup trước khi force_update
   cp suy_hao_history.db suy_hao_history_backup_$(date +%Y%m%d_%H%M%S).db
   ```

4. **File Excel luôn được tạo mới:**
   - Lần 1: Dữ liệu từ file I1.5 hiện tại + biến động tính toán
   - Lần 2+: Dữ liệu từ file I1.5 hiện tại + biến động từ DB (lần 1)

---

## Troubleshooting

### Vấn đề 1: "Biến động không khớp với dữ liệu trong Excel"

**Nguyên nhân:** File I1.5 lần 1 khác lần 2, nhưng biến động dựa trên lần 1

**Giải pháp:**
```python
process_I15_report(force_update=True)
```

### Vấn đề 2: "Muốn xem biến động của file mới"

**Giải pháp:** So sánh thủ công:
```python
# Đọc file hiện tại
df_current = pd.read_excel("downloads/baocao_hanoi/I1.5 report.xlsx")
current_accounts = set(df_current['ACCOUNT_CTS'].tolist())

# Đọc từ DB (lần trước)
conn = sqlite3.connect('suy_hao_history.db')
df_db = pd.read_sql_query("""
    SELECT DISTINCT account_cts
    FROM suy_hao_snapshots
    WHERE ngay_bao_cao = '2025-11-05'
""", conn)
db_accounts = set(df_db['account_cts'].tolist())

# So sánh
print(f"File hiện tại: {len(current_accounts)} thuê bao")
print(f"Đã lưu trong DB: {len(db_accounts)} thuê bao")
print(f"Chênh lệch: {len(current_accounts) - len(db_accounts)} thuê bao")
print(f"Thuê bao mới trong file: {current_accounts - db_accounts}")
```

### Vấn đề 3: "Nhầm lẫn giữa V1 và V2"

**Kiểm tra version đang dùng:**
```python
# Trong c1_process.py, dòng 409
from i15_process import process_I15_report_with_tracking  # ✅ V2
# from c1_process_enhanced import process_I15_report_with_tracking  # ❌ V1 cũ
```

---

**Phiên bản:** V2
**Ngày tạo:** 05/11/2025
**File liên quan:**
- [i15_process.py](i15_process.py): Logic V2
- [c1_process.py](c1_process.py): Wrapper function
- [HUONG_DAN_SU_DUNG_TRACKING.md](HUONG_DAN_SU_DUNG_TRACKING.md): Hướng dẫn tổng quát
