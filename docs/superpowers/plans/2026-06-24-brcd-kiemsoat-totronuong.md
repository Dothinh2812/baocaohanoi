# BRCD Kiểm soát tổ trưởng Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cho phép tổ trưởng nhập "nội dung kiểm soát" (1 ô tự do) cho từng phiếu tồn BRCD tại `/brcd`, rồi thống kê đã/chưa kiểm soát + thời điểm nhập, có lọc theo trạng thái phiếu và xuất báo cáo Excel.

**Architecture:** Tách 2 lớp giống pattern sa-outage: (1) live tồn vẫn đọc Excel read-only từ `1bss` (`chiTietBrcd5Doi.xlsx`, sheet đầy đủ `ToKT_<doi>` để có `baohong_id`); (2) annotation kiểm soát lưu trong DB SQLite **ghi được** per-instance (`INSTANCE_RUNTIME_DIR/brcd_kiemsoat.db`), keyed bằng `baohong_id`. Thống kê = LEFT JOIN tồn hiện tại với annotation. Annotation giữ lại làm lịch sử khi phiếu rời tồn.

**Tech Stack:** Flask blueprint (`operations`), SQLite ghi được (WAL), pandas, xlsxwriter, vanilla JS, pytest.

**Tiền lệ trong repo:** `blueprints/sa_outage_routes.py` (POST note JSON + `@login_required`, `_write_connection`, schema ensure idempotent) và `tests/test_sa_outage_routes.py` (DB trong `tmp_path` + `monkeypatch`).

## Quy ước / ràng buộc

- Khóa định danh: `baohong_id` (int, duy nhất/phiếu). Sheet `_rut_gon` đang hiển thị **thiếu** `baohong_id` → endpoint mới phải đọc sheet đầy đủ `ToKT_<doi>` (không `_rut_gon`).
- POST dùng `@login_required` + JSON body (không `@csrf_protect`), bám tiền lệ sa-outage. `nguoi_nhap = session['username']`.
- Không ghi ngược vào Excel nguồn (bị `1bss` refresh ghi đè).
- Không thuộc `report_history.db` → date-contract docs/09 không áp dụng.
- Test: `python3 -m pytest tests/test_brcd_kiemsoat.py -q` (system pytest). Verify syntax: `python3 -m py_compile`.
- Không commit trừ khi user yêu cầu.

---

### Task 1: Lớp lưu trữ + config

**Files:**
- Modify: `config.py` (thêm `BRCD_KIEMSOAT_DB_PATH` + vào `DashboardConfig`)
- Modify: `blueprints/operations_routes.py` (import + helper: `_brcd_kiemsoat_write_connection`, `_ensure_brcd_kiemsoat_schema`, `_brcd_kiemsoat_read_connection`, `get_brcd_kiemsoat_map`)
- Test: `tests/test_brcd_kiemsoat.py`

**Interfaces:**
- Produces: `config.BRCD_KIEMSOAT_DB_PATH`; trong `operations_routes.py`: bảng `brcd_kiemsoat(baohong_id INTEGER PK, ma_tb, doi_vt, nvkt, noi_dung_kiem_soat, nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat)`; hàm `get_brcd_kiemsoat_map(baohong_ids: list[int]) -> dict[int, dict]`.

- [ ] Thêm `BRCD_KIEMSOAT_DB_PATH` vào `config.py` (`_first_existing_path(os.getenv('DASHV4_BRCD_KIEMSOAT_DB_PATH'), os.path.join(INSTANCE_RUNTIME_DIR, 'brcd_kiemsoat.db'))`) và vào class `DashboardConfig`.
- [ ] Thêm helper vào `operations_routes.py`: write/read connection (WAL, busy_timeout, Row factory), `_ensure_brcd_kiemsoat_schema()` (CREATE TABLE IF NOT EXISTS, idempotent theo path + Lock), `get_brcd_kiemsoat_map(ids)` (read-only, SELECT ... WHERE baohong_id IN (...)).
- [ ] Test: tạo DB trong `tmp_path`, monkeypatch `operations_routes.BRCD_KIEMSOAT_DB_PATH`, assert `get_brcd_kiemsoat_map` trả đúng map sau khi insert trực tiếp, và trả `{}` khi list rỗng.
- [ ] Chạy `python3 -m pytest tests/test_brcd_kiemsoat.py -q` → pass.

### Task 2: POST lưu nội dung kiểm soát

**Files:**
- Modify: `blueprints/operations_routes.py`
- Test: `tests/test_brcd_kiemsoat.py`

**Interfaces:**
- Produces: `POST /api/brcd-kiemsoat/luu` (endpoint `operations.api_brcd_kiemsoat_luu`), body `{baohong_id, ma_tb, doi_vt, nvkt, noi_dung}`.

- [ ] Thêm route `POST /api/brcd-kiemsoat/luu`, `@login_required`: parse `baohong_id` (int), `noi_dung` (strip, max 2000), `ma_tb/doi_vt/nvkt`. Nếu `noi_dung` rỗng → DELETE row (bỏ kiểm soát). Không thì UPSERT (giữ `thoi_diem_nhap` cũ nếu đã tồn tại, set `thoi_diem_cap_nhat = now`). Trả `{ok, baohong_id, noi_dung, nguoi_nhap}`.
- [ ] Test: POST lưu note mới (assert `nguoi_nhap='test-user'`, có `thoi_diem_nhap`); POST cập nhật (assert `thoi_diem_nhap` giữ nguyên, `thoi_diem_cap_nhat` đổi); POST note rỗng → xóa; POST `baohong_id` sai → 400; POST > 2000 ký tự → 400.
- [ ] Chạy test → pass.

### Task 3: GET detail (join kiemsoat)

**Files:**
- Modify: `blueprints/operations_routes.py`
- Test: `tests/test_brcd_kiemsoat.py`

**Interfaces:**
- Produces: `_load_brcd_kiemsoat_df()` (DataFrame đã join, numeric nguyên vẹn); `GET /api/brcd-kiemsoat/detail` (endpoint `operations.api_brcd_kiemsoat_detail`) → `{sheets: {<sheet>: {columns, data}}, file_info}`.

- [ ] Thêm `BRCD_KIEMSOAT_DISPLAY_COLUMNS` (bao gồm `baohong_id`, `ma_tb`, `NVKT`, `DOI_VT`, `ngay_bh`, `Trạng thái cổng`, `giờ còn lại thực`, `thời gian tồn thực`, `chitieu_tg`, `ttvt_ton`, ...).
- [ ] `_load_brcd_kiemsoat_df()`: đọc sheet đầy đủ `ToKT_<doi>` (bỏ `_rut_gon`), curate cột, ép `baohong_id` + cột giờ về numeric, concat, `get_brcd_kiemsoat_map(all_ids)`, gán `kiemsoat_noi_dung/kiemsoat_nguoi_nhap/kiemsoat_thoi_diem/kiemsoat_da_nhap`. Trả DataFrame hoặc None nếu thiếu file.
- [ ] Route `GET /api/brcd-kiemsoat/detail`: group theo `_sheet` → serialize từng (`serialize_dataframe` + `build_sheet_payload`).
- [ ] Test: dựng Excel giả trong `tmp_path` (pandas `to_excel` sheet `ToKT_SonTay` có `baohong_id` + cột giờ), monkeypatch `BRCD_DETAIL_MAIN_FILE`; insert 1 kiemsoat row; assert detail trả row có `kiemsoat_da_nhap=True` cho phiếu đã nhập và `False` cho phiếu chưa.
- [ ] Chạy test → pass.

### Task 4: GET thống kê (filter + aggregate)

**Files:**
- Modify: `blueprints/operations_routes.py`
- Test: `tests/test_brcd_kiemsoat.py`

**Interfaces:**
- Produces: `_apply_brcd_kiemsoat_filters(df, args)`; `GET /api/brcd-kiemsoat/thongke` (endpoint `operations.api_brcd_kiemsoat_thongke`).

- [ ] `_apply_brcd_kiemsoat_filters(df, args)`: `doi` (DOI_VT), `loaihinh` (LOAIHINH_TB), `nhom` (`qua_gio`→`giờ còn lại thực`≤0 / `trong_gio`→>0), `trangthai` (`da`/`chua` theo `kiemsoat_da_nhap`).
- [ ] Route `GET /api/brcd-kiemsoat/thongke`: load df → filter → trả `{summary: {total, da_kiem_soat, chua, ty_le}, by_doi: [...], by_nvkt: [...], chi_tiet: [...]}`. `chi_tiet` serialize qua `serialize_dataframe`.
- [ ] Test: dữ liệu giả có 1 phiếu quá giờ đã kiểm soát + 1 trong giờ chưa; assert `summary`, `by_doi`, và filter `?nhom=qua_gio` + `?trangthai=chua` đúng.
- [ ] Chạy test → pass.

### Task 5: Xuất báo cáo Excel

**Files:**
- Modify: `blueprints/operations_routes.py`
- Test: `tests/test_brcd_kiemsoat.py`

- [ ] Route `GET /download/brcd-kiemsoat-report` (`@login_required`): load df → filter theo query params → `pd.ExcelWriter(engine='xlsxwriter')` ghi `serialize_dataframe(df)` → `send_file` (BytesIO, `as_attachment`, `download_name=brcd_kiemsoat_YYYYMMDD_HHMM.xlsx`).
- [ ] Test: GET với Excel giả → 200, `content-type` xlsx, body bắt đầu bằng PK; assert áp filter giảm số dòng (qua content-length hoặc đọc lại Excel).
- [ ] Chạy test → pass.

### Task 6: Frontend

**Files:**
- Modify: `templates/pages/brcd.html`
- Modify: `static/js/pages/brcd.js`
- Modify: `static/js/api.js` (thêm `getBrcdKiemSoatDetail`, `getBrcdKiemSoatThongKe`, `saveBrcdKiemSoat`)
- Test: `tests/test_brcd_kiemsoat.py` (HTML assertion)

- [ ] `brcd.html`: thêm section "Kiểm soát tổ trưởng" chứa: card thống kê (`id="kiemsoat-stats"`), bộ lọc (đội, nhóm giờ, trạng thái), và bảng chi tiết có cột "Nội dung kiểm soát" (textarea) + nút Lưu/dòng + badge thời điểm.
- [ ] `api.js`: thêm 3 hàm (`getBrcdKiemSoatDetail`, `getBrcdKiemSoatThongKe` GET; `saveBrcdKiemSoat` POST JSON).
- [ ] `brcd.js`: render bảng chi tiết từ `/api/brcd-kiemsoat/detail` với textarea prefetch + badge "Đã kiểm soát HH:MM dd/MM — <nguoi>"; bind nút Lưu → POST → cập nhật badge; load thongke vào card; bind lọc → gọi lại thongke.
- [ ] Test: GET `/brcd` → 200, HTML chứa `id="kiemsoat-stats"`, "Nội dung kiểm soát", "Lưu".
- [ ] Chạy test → pass.

### Task 7: Doc-sync + verification

**Files:**
- Modify: `docs/08-trang-thai-thuc-thi.md`
- Modify: `docs/04-mapping-route-va-du-lieu.md`

- [ ] `docs/08`: thêm feature "Kiểm soát tổ trưởng tại /brcd" (nguồn: Excel 1bss + DB `brcd_kiemsoat.db` per-instance, không thuộc report_history.db).
- [ ] `docs/04`: thêm mapping cho 4 endpoint mới (route ↔ nguồn ↔ `supports_date=否`).
- [ ] `python3 -m py_compile` cho các file đã sửa.
- [ ] Chạy `python3 -m pytest tests/ -q` → tất cả pass.
