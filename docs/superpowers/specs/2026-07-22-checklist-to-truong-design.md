# Thiết kế: Checklist hằng ngày cho tổ trưởng tổ KTĐB

- **Ngày:** 2026-07-22
- **Liên quan:** `docs/11-kiemsoat-to-truong-van-hanh.md` (kiến trúc tham chiếu), `docs/04-mapping-route-va-du-lieu.md` (doc-sync)
- **Trạng thái:** Đã duyệt qua brainstorm, chờ writing-plans

## Bối cảnh & vấn đề

Hiện dashv4 đã có lớp "kiểm soát tổ trưởng" tại `/brcd`, `/pttb`, `/shc-cts`: tổ trưởng nhập annotation tự do trên từng phiếu tồn, có thống kê, lịch sử, Excel export. Pattern này phù hợp cho việc kiểm soát **đơn vị phiếu**, không phù hợp cho việc chuẩn hóa **công việc hằng ngày của tổ trưởng**.

TTVT cần chuẩn hóa quy trình điều hành của tổ trưởng tổ KTĐB theo 3 giai đoạn:

1. **Đầu giờ sáng (07:15-07:45): Giao việc** — 7 mục
2. **Trong ngày (07:45-18:00): Điều phối** — 4 mục
3. **Cuối ngày (trước 18:00): Đánh giá kỷ luật** — 4 mục

Tổng 15 mục công việc cố định lặp lại mỗi ngày. Tổ trưởng cần:
- Cập nhật trạng thái từng mục (`chua_lam` / `dang_lam` / `xong` / `bo_qua`)
- Ghi chú text lý do/bắt tay việc
- Đính kèm ảnh minh chứng (ATVSLĐ, hiện trường trạm, biên bản vi phạm...)
- Thấy số liệu live từ các báo cáo đã có (SC quá hạn, kiemsoat completion...)
- Phân quyền per-đội: tổ trưởng đội X sửa đội X, xem được các đội khác (read-only), admin sửa tất cả

## Mục tiêu

Xây dựng trang `/checklist` mới (blueprint `checklist_bp`), tách biệt khỏi 3 trang kiemsoat hiện có (giữ nguyên không đụng), với:

- Mẫu template cố định hard-code (3 phases × 15 mục), sẵn sàng nâng cấp lên config-driven sau.
- Data model per-đội: key `(ngay, doi, phase_key, item_key)`.
- Photo upload hạ tầng mới (chưa từng có trong codebase).
- Auto-source metric display: 5 mục có widget hiển thị metric live từ `brcd_kiemsoat.db` / `shc_cts.db`.
- Phân quyền qua role mới `to_truong` + cột `doi` mới trong `username.xlsx`.
- Audit log append-only (hỗ trợ "biên bản vi phạm kỷ luật").
- Excel export đa sheet.
- Multi-instance auto-isolation, lazy migration, không phá instances đang chạy.

## Phạm vi

**Trong phạm vi (MVP):**
- New blueprint `blueprints/checklist_routes.py` + template + JS + CSS mới.
- 3 bảng SQLite mới trong `<INSTANCE_RUNTIME_DIR>/checklist.db` (`checklist_state`, `checklist_photos`, `checklist_audit_log`).
- Photo upload infrastructure (multipart form + CSRF + secure_filename + serving).
- Mở rộng `auth.py` + `username.xlsx`: thêm role `to_truong` + cột `doi`.
- 5 metric fetcher đọc read-only từ DBs có sẵn.
- Excel export 2 sheet.
- ~46 test mới (suite 47 → ~93).
- Docs mới `docs/12-checklist-to-truong.md` + update `docs/00`, `docs/04`, `docs/08`, `AGENTS.md`.

**Ngoài phạm vi (YAGNI / v2):**
- Cấu hình template qua admin UI (hard-code trước, nâng cấp sau).
- Auto-cleanup ảnh cũ (audit/legal value, giữ lại).
- Notification (email/Zalo) khi chưa chốt cuối ngày.
- PWA/mobile app.
- Snapshot cron (không có nguồn Excel ngoài để snapshot, metric đọc live mỗi page load).
- Time-lock phase (07:15-07:45 chỉ là guideline, không cứng).

## Quyết định thiết kế

### Phương án kiến trúc đã chọn: New blueprint `checklist_routes.py` (Approach A)

Loại trừ:
- **Extend `operations_routes.py`** (Approach B): file đã 1947 dòng, trộn 2 luồng khác nhau (annotation theo phiếu vs checklist hằng ngày theo đội).
- **Layered `blueprints/` + `services/` + `repositories/`** (Approach C): codebase chưa có convention này đầy đủ, over-engineering cho MVP. Có thể tách `services/checklist_aggregator.py` ra sau khi file phình (>1000 dòng).

Lý do chọn: tuân thủ convention codebase (1 domain = 1 blueprint file), clean separation, dễ test/review.

### Nguyên lý thiết kế

- **Write-aside**: checklist lưu DB riêng, không ghi ngược vào brcd_kiemsoat.db hay bất kỳ Excel nào.
- **Per-đội scope**: 1 tổ trưởng phụ trách 1 đội, key `(ngay, doi, phase_key, item_key)` cho biết tổ trưởng đội nào đang cập nhật gì hôm nào.
- **View-all, edit-own**: minh bạch chéo, tổ trưởng thấy tiến độ các đội khác nhưng không sửa.
- **Hard-code template, ready nâng cấp**: 3 phases × 15 mục là Python constant; schema thiết kế để bảng `checklist_template` có thể thêm sau mà không phá dữ liệu.
- **Metric display-only**: metric live chỉ hiển thị (badge), không tự động đánh dấu "Xong" — tránh false positive.
- **Lazy migration users.xlsx**: thêm cột `doi`, user cũ giữ nguyên `role=user, doi=None`, không cần script 1-shot.
- **No snapshot cron**: không có nguồn ngoài để snapshot; metric fetcher đọc live mỗi page load.
- **Audit append-only**: hỗ trợ truy vết khi có tranh chấp/khiếu nại; MVP không render trong UI.

## Thiết kế chi tiết

### 1. Cấu trúc file & wiring

**File mới:**
- `blueprints/checklist_routes.py` (~700 dòng) — toàn bộ logic: schema setup idempotent, write/read connection helpers, 3 module constants (3 giai đoạn × items), endpoints, 5 metric fetcher, registry.
- `static/js/pages/checklist.js` — fetch data, render 3 phase cards, save handlers, photo upload với progress.
- `static/css/checklist.css` — status pill styles, accordion, drop-zone.
- `templates/pages/checklist.html` — page shell (theo mẫu `brcd.html`).
- `tests/test_checklist_state.py`, `tests/test_checklist_photos.py`, `tests/test_checklist_metric_fetchers.py`, `tests/test_auth_to_truong.py`, `tests/test_checklist_routes.py`.

**File sửa (wiring):**
- `blueprints/__init__.py` — export `checklist_bp` + `__all__`.
- `dashboard.py` — `app.register_blueprint(checklist_bp)`.
- `route_policy.py` — thêm `'checklist.page_checklist': 'checklist'` vào `PAGE_ACTIVE_KEYS`.
- `templates/base.html` sidebar — thêm 1 menu item `{% if is_endpoint_enabled('checklist.page_checklist') %}` (đặt sau nhóm kiemsoat).
- `config.py` — thêm `CHECKLIST_DB_PATH` + `CHECKLIST_UPLOAD_DIR` + `CHECKLIST_PHOTO_MAX_BYTES` + `CHECKLIST_PHOTO_MAX_PER_ITEM`.
- `auth.py` — mở rộng role enum (`to_truong`), thêm cột `doi`, helper session, decorator `to_truong_or_admin_required`.
- `auth_routes.py` — set `session['doi']`, thêm handler `/admin/set-user-doi`.
- `templates/admin_users.html` — thêm cột Đội + dropdown Role 3 giá trị.
- `static/js/api.js` — thêm 4-5 methods (data/save/upload-photo/delete-photo/download-report).
- `app_helpers.py` — thêm helper `get_csrf_token_for_js` (inject vào meta tag cho JS fetch multipart).

### 2. Endpoints

| Method | URL | Endpoint | Decorator | Mục đích |
|--------|-----|----------|-----------|----------|
| GET | `/checklist` | `checklist.page_checklist` | `@login_required` | Render page shell |
| GET | `/api/checklist/data?date=&doi=` | `checklist.api_checklist_data` | `@login_required` | Trả template (3 phases × items) + state user + metric live + photos + can_edit |
| POST | `/api/checklist/luu` | `checklist.api_checklist_luu` | `@login_required` + `@to_truong_or_admin_required` | JSON UPSERT state (empty → DELETE) + audit log |
| POST | `/api/checklist/upload-photo` | `checklist.api_checklist_upload_photo` | `@login_required` + `@to_truong_or_admin_required` + **`@csrf_protect`** | Multipart form, validate + save file + INSERT row + audit |
| POST | `/api/checklist/delete-photo` | `checklist.api_checklist_delete_photo` | `@login_required` + `@to_truong_or_admin_required` | JSON, ownership check + delete file + row + audit |
| GET | `/checklist/photo/<filename>` | `checklist.serve_photo` | `@login_required` | Serve ảnh với path traversal guard + cache headers |
| GET | `/download/checklist-report?from=&to=&doi=` | `checklist.download_checklist_report` | `@login_required` | Xuất Excel đa sheet (state + photos) |

POST JSON endpoints theo convention kiemsoat: **không dùng `@csrf_protect`** (chỉ form-encoded mới cần). Upload-photo là multipart form → bắt buộc `@csrf_protect`.

### 3. Schema DB — `<INSTANCE_RUNTIME_DIR>/checklist.db`

SQLite WAL + `busy_timeout=5000`. Schema tạo idempotent (`CREATE TABLE IF NOT EXISTS`) trong `_ensure_checklist_schema()`, cache per-worker `_checklist_schema_ready_path` (giống `_brcd_kiemsoat_schema_ready_path`).

#### 3.1. `checklist_state` (state mới nhất per item)

```sql
CREATE TABLE IF NOT EXISTS checklist_state (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay                TEXT NOT NULL,           -- YYYY-MM-DD
    doi                 TEXT NOT NULL,           -- vd "Đội 1", khớp doi_vt trong brcd_phieu
    phase_key           TEXT NOT NULL,           -- vd "sang_giao_viec"
    item_key            TEXT NOT NULL,           -- vd "atvsld_trang_phuc_ccdc"
    trang_thai          TEXT NOT NULL DEFAULT 'chua_lam',  -- chua_lam|dang_lam|xong|bo_qua
    ghi_chu             TEXT,                    -- max 2000 chars (như kiemsoat)
    nguoi_nhap          TEXT,                    -- username từ session
    nguoi_nhap_name     TEXT,                    -- display name cache
    thoi_diem_nhap      TEXT,                    -- ISO timestamp INSERT đầu
    thoi_diem_cap_nhat  TEXT,                    -- ISO timestamp UPDATE cuối
    UNIQUE(ngay, doi, phase_key, item_key)
);
CREATE INDEX IF NOT EXISTS idx_checklist_state_ngay_doi ON checklist_state(ngay, doi);
```

UPSERT key: `(ngay, doi, phase_key, item_key)`. Khi `ghi_chu` rỗng VÀ `trang_thai=chua_lam` → DELETE row (giống kiemsoat "empty content = DELETE").

#### 3.2. `checklist_photos` (N ảnh per item)

```sql
CREATE TABLE IF NOT EXISTS checklist_photos (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay                TEXT NOT NULL,
    doi                 TEXT NOT NULL,
    phase_key           TEXT NOT NULL,
    item_key            TEXT NOT NULL,
    filename_stored     TEXT NOT NULL,           -- <uuid4_hex>.<ext>
    filename_original   TEXT,                    -- đã qua secure_filename, để hiển thị
    mime_type           TEXT,                    -- image/jpeg|png|webp
    size_bytes          INTEGER,
    nguoi_nhap          TEXT,
    thoi_diem_nhap      TEXT                     -- ISO timestamp
);
CREATE INDEX IF NOT EXISTS idx_checklist_photos_item ON checklist_photos(ngay, doi, phase_key, item_key);
```

**Không có FK** tới `checklist_state` — vì có thể upload ảnh trước khi state được tạo. Query JOIN bằng 4 cột composite key.

#### 3.3. `checklist_audit_log` (append-only)

```sql
CREATE TABLE IF NOT EXISTS checklist_audit_log (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    ngay                TEXT NOT NULL,
    doi                 TEXT NOT NULL,
    phase_key           TEXT NOT NULL,
    item_key            TEXT NOT NULL,
    action              TEXT NOT NULL,           -- upsert_state|delete_state|upload_photo|delete_photo
    payload             TEXT,                    -- JSON snapshot before/after
    nguoi_thuc_hien     TEXT,
    thoi_diem           TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_checklist_audit_ngay_doi ON checklist_audit_log(ngay, doi);
```

Append-only, không bao giờ UPDATE/DELETE. MVP không render trong UI; query khi cần truy vết (admin/Giám đốc yêu cầu).

### 4. Mẫu checklist template (hard-code)

```python
CHECKLIST_TEMPLATE = [
    {
        'phase_key': 'sang_giao_viec',
        'phase_name': 'Đầu giờ sáng (07:15-07:45): Giao việc',
        'items': [
            {'item_key': 'atvsld_trang_phuc_ccdc',
             'name': 'Đánh giá ATVSLĐ, trang phục, CCDC toàn tổ', 'metric_source': None},
            {'item_key': 'pho_bien_vb_moi',
             'name': 'Phổ biến văn bản mới VNPT HNi / TTVT', 'metric_source': None},
            {'item_key': 'giai_dap_vuong_mac',
             'name': 'Giải đáp hướng dẫn vướng mắc NVKT', 'metric_source': None},
            {'item_key': 'doichieu_vattu_hoancong',
             'name': 'Đối chiếu vật tư tồn kho, hoàn công/treỏo', 'metric_source': 'inventory'},
            {'item_key': 'giao_chitieu_catondong',
             'name': 'Giao chỉ tiêu xử lý dứt điểm ca tồn đọng', 'metric_source': 'brcd_pttb_overdue'},
            {'item_key': 'phancong_hotro_cheo',
             'name': 'Phân công nhân sự hỗ trợ chéo', 'metric_source': None},
            {'item_key': 'giao_chitieu_nangsuat',
             'name': 'Giao chỉ tiêu năng suất ngày (suy hao/chạm cước)', 'metric_source': None},
        ],
    },
    {
        'phase_key': 'trong_ngay_dieu_phoi',
        'phase_name': 'Trong ngày (07:45-18:00): Điều phối',
        'items': [
            {'item_key': 'giamsat_dieu_huong',
             'name': 'Giám sát điều hướng ưu tiên (sự cố > lắp mới > dịch chuyển > suy hao > gia hạn)',
             'metric_source': 'brcd_pttb_priority'},
            {'item_key': 'capphat_vattu_s2s1',
             'name': 'Cấp phát vật tư bổ sung, hỗ trợ S2/S1', 'metric_source': None},
            {'item_key': 'bori_lich_thucdia',
             'name': 'Bố trí lịch thực địa (1-2 buổi/tuần)', 'metric_source': None},
            {'item_key': 'trieu_tap_dot_xuat',
             'name': 'Triệu tập đột xuất 13h15 NV năng suất thấp',
             'metric_source': 'brcd_nvkt_low_productivity'},
        ],
    },
    {
        'phase_key': 'cuoi_ngay_danh_gia',
        'phase_name': 'Cuối ngày (trước 18:00): Đánh giá kỷ luật',
        'items': [
            {'item_key': 'tiepnhan_thietbi_thuhoi',
             'name': 'Tiếp nhận thiết bị thu hồi, vật tư thừa trả', 'metric_source': None},
            {'item_key': 'rasoat_nguyennhan_ton',
             'name': 'Rà soát 100% NV cập nhật nguyên nhân tồn',
             'metric_source': 'kiemsoat_completion'},
            {'item_key': 'lap_bienban_vipham',
             'name': 'Lập biên bản vi phạm kỷ luật', 'metric_source': None},
            {'item_key': 'tonghop_dutru_vattu',
             'name': 'Tổng hợp dự trù vật tư, kế hoạch nhân sự ngày sau', 'metric_source': None},
        ],
    },
]
```

**15 mục** (7+4+4), **5 mục có metric_source** để hiển thị metric live.

### 5. Auth & role extension

#### 5.1. Thay đổi `username.xlsx`

- **Cột mới `doi`** (TEXT, nullable): giá trị khớp cột `doi_vt` trong brcd_phieu (vd `Đội 1`).
- **Role enum mở rộng**: `admin` / `user` / **`to_truong`** (mới).
- **Lazy migration**: `auth.get_all_users()` đọc `doi` nếu cột tồn tại, fallback `None` nếu thiếu. Users cũ giữ nguyên `role=user, doi=None` → không bị phá.

#### 5.2. Thay đổi `auth.py`

- `get_all_users()` (~line 35-58): thêm `doi` vào DataFrame nếu thiếu, return dict có key `doi`.
- `update_user()` (~line 82-100): chấp nhận `doi=None`, ghi vào cột.
- Validation role (~line 173): `if new_role not in ['admin', 'user', 'to_truong']`.
- Session (`auth_routes.py:38-41`): thêm `session['doi'] = user.get('doi')`.
- `_current_user()`: đảm bảo trả thêm key `doi` từ session.
- Helper response mới `_unauth_response()` / `_forbidden_response()` (return JSON cho `/api/*`, redirect cho page).

#### 5.3. Decorator mới `to_truong_or_admin_required`

```python
def to_truong_or_admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        user = _current_user()
        if not user or 'username' not in session:
            return _unauth_response()
        if user.get('role') == 'admin':
            return f(*args, **kwargs)
        if user.get('role') == 'to_truong' and user.get('doi'):
            return f(*args, **kwargs)
        return _forbidden_response()
    return wrapper
```

#### 5.4. Ownership check trong POST handler

- `api_checklist_luu` / `upload_photo` / `delete_photo`:
  - User `admin`: payload `doi` bất kỳ → OK.
  - User `to_truong`: payload `doi` PHẢI == `session['doi']`, nếu khác → `abort(403, json)`.
  - User `user`: bị chặn ở decorator.

#### 5.5. Quyền xem

- `api_checklist_data`: chỉ cần `@login_required`. Response kèm flag `can_edit: bool` per đội (true nếu user là admin HOẶC `to_truong` của đúng đội). JS dùng flag để bật/tắt UI edit.

#### 5.6. UI admin (`templates/admin_users.html`)

- Thêm cột "Đội": hiển thị `doi` (hoặc "—"); nút "Sửa" mở modal dropdown các đội có sẵn.
- Thêm cột "Role": dropdown 3 giá trị `admin` / `to_truong` / `user`.
- Dropdown đội = `SELECT DISTINCT doi_vt FROM brcd_phieu` (snapshot table), fallback text input nếu rỗng.
- Sau khi đổi: `invalidate_users_cache()` (đã có).

### 6. Photo upload (hạ tầng mới)

#### 6.1. Cấu hình

```python
# config.py
CHECKLIST_DB_PATH = _first_existing_path(
    os.getenv('DASHV4_CHECKLIST_DB_PATH'),
    os.path.join(INSTANCE_RUNTIME_DIR, 'checklist.db'),
)
CHECKLIST_UPLOAD_DIR = _first_existing_path(
    os.getenv('DASHV4_CHECKLIST_UPLOAD_DIR'),
    os.path.join(INSTANCE_RUNTIME_DIR, 'checklist_uploads'),
)
CHECKLIST_PHOTO_MAX_BYTES = int(os.getenv('DASHV4_CHECKLIST_PHOTO_MAX_BYTES', 5 * 1024 * 1024))
CHECKLIST_PHOTO_MAX_PER_ITEM = int(os.getenv('DASHV4_CHECKLIST_PHOTO_MAX_PER_ITEM', 5))
```

#### 6.2. Cấu trúc thư mục

```
<INSTANCE_RUNTIME_DIR>/checklist_uploads/
  2026-07/                    -- partition theo tháng, tránh 1 dir phình to
    <uuid4_hex>.jpg
    <uuid4_hex>.png
```

- **Tên file lưu**: `<uuid.uuid4().hex>.<ext>` — không dấu Việt, không space, không path traversal.
- **Tên gốc**: lưu `filename_original` (qua `secure_filename`) để hiển thị.
- Dir `<YYYY-MM>/` được `os.makedirs(..., exist_ok=True)` trong handler.

#### 6.3. Endpoint `POST /api/checklist/upload-photo` — validation chain

Theo thứ tự:
1. `request.content_length > MAX_BYTES + 1KB overhead` → 413 sớm.
2. `photo = request.files.get('photo')` — None → 400 `"Thiếu file"`.
3. `photo.filename == ''` → 400 `"File rỗng"`.
4. Ext `os.path.splitext(filename)[1].lower()` phải thuộc `{'.jpg', '.jpeg', '.png', '.webp'}`.
5. `photo.mimetype` phải thuộc `{image/jpeg, image/png, image/webp}` — mismatch → 400.
6. `len(photo.read())` > `MAX_BYTES` → 413. **Seek lại 0 sau khi read**.
7. Ownership: form `doi` == `session['doi']` (trừ admin).
8. Per-item limit: `SELECT COUNT(*) ... WHERE ngay,doi,phase_key,item_key` ≥ MAX_PER_ITEM → 400 `"Đạt giới hạn"`. **Dùng transaction `BEGIN IMMEDIATE` để tránh TOCTOU**.

#### 6.4. Lưu file

1. `original_secure = secure_filename(photo.filename)`.
2. `stored_filename = f"{uuid.uuid4().hex}{ext}"`.
3. `month_dir = os.path.join(UPLOAD_DIR, ngay[:7])`; `os.makedirs(month_dir, exist_ok=True)`.
4. `photo.save(os.path.join(month_dir, stored_filename))`.
5. INSERT vào `checklist_photos`.
6. INSERT audit log.
7. Return `{ok: 1, photo: {id, url, thumbnail_url, filename_original, size_bytes}}`.

#### 6.5. Endpoint `POST /api/checklist/delete-photo`

JSON `{photo_id}`. Decorator không cần `@csrf_protect` (JSON convention). Ownership check row's `doi` == `session['doi']`. `os.remove()` trong try/except (FileNotFoundError OK). DELETE row + audit log.

#### 6.6. Endpoint `GET /checklist/photo/<filename>`

- `filename` phải match `^[a-f0-9]{32}\.(jpg|jpeg|png|webp)$` (regex) — chặn path traversal sớm.
- Query DB để lấy `ngay` từ `filename_stored` → derive month dir.
- Dùng `safe_directory_response(month_dir, filename)` từ `app_helpers.py:87-110`.
- Headers: `Cache-Control: public, max-age=86400` (file immutable).

#### 6.7. JS upload flow

```js
async function uploadPhoto(phaseKey, itemKey, fileInput) {
    const formData = new FormData();
    formData.append('csrf_token', getCsrfToken());  // từ meta tag
    formData.append('date', currentDate);
    formData.append('doi', currentDoi);
    formData.append('phase_key', phaseKey);
    formData.append('item_key', itemKey);
    formData.append('photo', fileInput.files[0]);

    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/checklist/upload-photo');
    xhr.upload.onprogress = (e) => showProgress(e.loaded / e.total);
    xhr.send(formData);
}
```

- Drag-drop zone + click để chọn.
- Thumbnail preview ngay sau upload.
- Confirm dialog trước khi delete.

#### 6.8. Multi-worker safety

- Upload file: mỗi upload = uuid duy nhất → không race.
- DB metadata: WAL + busy_timeout + catch `OperationalError: locked` → 503 `{error: 'db_locked'}`, JS retry 1 lần sau 500ms.
- Per-item limit TOCTOU: `BEGIN IMMEDIATE` transaction + re-check count trong transaction.

#### 6.9. Storage estimate

- Trung bình: 5 ảnh/ngày/đội × 5 đội × 1 MB = 25 MB/ngày, ~9 GB/năm.
- Tồi tệ nhất: 5 ảnh/mục × 15 mục × 5 đội × 5 MB = 1.9 GB/ngày — KHÔNG thực tế.
- Soft warning khi `du -sh` vượt 500 MB (trong docs/12 maintenance checklist).

### 7. Metric aggregation (auto-source)

#### 7.1. Nguyên tắc

- **Read-only against existing per-instance DBs**: `brcd_kiemsoat.db`, `shc_cts.db` (mở qua `file:...?mode=ro&immutable=1` URI như `repositories/sqlite_runtime.py`).
- **Fail-safe**: mỗi fetcher wrap try/except, trả `None` khi DB lock/missing/corrupt → UI hiện `"Không có dữ liệu"`.
- **Không gọi HTTP endpoint nội bộ**: gọi thẳng SQL hoặc import `_compute_*` helpers từ `operations_routes.py`.
- **Không cache** trong MVP — fetch mỗi page load. 5 query × ~50ms = ~250ms, acceptable. Nếu chậm, thêm `@lru_cache(maxsize=128, key=(doi, ngay))` + TTL 5 phút sau.

#### 7.2. Registry pattern

```python
METRIC_FETCHERS = {
    'inventory': _fetch_inventory,
    'brcd_pttb_overdue': _fetch_brcd_pttb_overdue,
    'brcd_pttb_priority': _fetch_brcd_pttb_priority,
    'brcd_nvkt_low_productivity': _fetch_brcd_nvkt_low_productivity,
    'kiemsoat_completion': _fetch_kiemsoat_completion,
}

def _safe_fetch(source_key, doi, ngay):
    fn = METRIC_FETCHERS.get(source_key)
    if not fn:
        return None
    try:
        return fn(doi, ngay)
    except sqlite3.OperationalError as e:
        if 'locked' in str(e).lower() or 'busy' in str(e).lower():
            app.logger.warning(f"checklist metric {source_key} db_locked")
            return None
        raise
    except Exception as e:
        app.logger.exception(f"checklist metric {source_key} failed")
        return None
```

#### 7.3. Output contract

```python
{
    'label': 'SC/PT quá hạn',
    'value': '5 phiếu',
    'detail': {'sc': 3, 'pt': 2},
    'fetched_at': '2026-07-22T07:30:00',
    'source_url': '/brcd',
    'source_label': 'Xem chi tiết',
}
```

UI render: badge nhỏ cạnh tên mục, click mở `source_url` tab mới.

#### 7.4. Chi tiết từng fetcher

**`_fetch_brcd_pttb_overdue(doi, ngay)`** — query `brcd_phieu` + `pttb_phieu` từ `brcd_kiemsoat.db`:
- WHERE `doi_vt = ?` AND `gio_con_lai_thuc < 0`.
- Return `{sc: N, pt: M, tong: N+M, source_url: '/brcd'}`.

**`_fetch_brcd_pttb_priority(doi, ngay)`** — query `brcd_phieu` group theo `trang_thai_cong` hoặc `sa`:
- Map raw values sang 5 hạng: Sự cố > Lắp mới > Dịch chuyển > Suy hao > Gia hạn.
- Return `{su_co: N, lap_moi: M, ...}`.

**`_fetch_brcd_nvkt_low_productivity(doi, ngay)`** — query `brcd_phieu` filter `doi_vt = ?` AND `date(first_seen) = ?`:
- Group by `nvkt`, count `trang_thai_cong = 'Đã xong'` vs total.
- Threshold: `completion_rate < 0.5` trước 13h15 → flag.
- Return `{low_nvkt: [{'nvkt': 'A', 'total': 10, 'done': 2, 'rate': 0.2}, ...]}`.

**`_fetch_kiemsoat_completion(doi, ngay)`** — `brcd_phieu` LEFT JOIN `brcd_kiemsoat`:
- WHERE `doi_vt = ?` AND `brcd_kiemsoat.noi_dung_kiem_soat IS NOT NULL`.
- Return `{total_phieu: N, da_ks: M, chua_ks: N-M, rate: M/N, source_url: '/brcd'}`.

**`_fetch_inventory(doi, ngay)`** — placeholder MVP:
- Investigate `blueprints/inventory_routes.py` để biết data source.
- Nếu phức tạp, return `None` + log warning → UI hiện "Không có dữ liệu". Cải thiện v2.

#### 7.5. Cần verify khi implement

- `brcd_phieu.doi_vt` case-sensitive; `pttb_phieu` dùng `DOI_VT` chữ hoa.
- `gio_con_lai_thuc` (BRCD) vs `gio_conlai` (PTTB) — khác format.
- `trang_thai_cong` enum values: chạy `SELECT DISTINCT trang_thai_cong FROM brcd_phieu LIMIT 20` trước.
- Mỗi query SELECT cột tối thiểu (vd `doi_vt, trang_thai_cong, gio_con_lai_thuc`).

### 8. UI layout

#### 8.1. Layout chọn: 3 section dọc + accordion per item

Loại trừ tabs (giấu overview), 3 cột (quá hẹp cho notes+photos).

```
┌────────────────────────────────────────────────────────────┐
│ 📋 Checklist tổ trưởng tổ KTĐB                    [Hôm nay]  │
│ Đơn vị: Sơn Tây                                            │
└────────────────────────────────────────────────────────────┘
┌────────────────────────────────────────────────────────────┐
│ Ngày: [< Hôm qua]  [2026-07-22 ▼]  [Ngày mai >]           │
│ Đội:  [Đội 1 ▼]              Tiến độ: ██████ 8/15 (53%)    │
└────────────────────────────────────────────────────────────┘
┌─── Phase 1: Đầu giờ sáng (07:15-07:45): Giao việc ───────┐
│ ⏰ 07:15-07:45           Tiến độ: 5/7                     │
├──────────────────────────────────────────────────────────┤
│ [Xong ▼]  Đánh giá ATVSLĐ, trang phục, CCDC toàn tổ     │
│            [badge: SC quá hạn: 5 → /brcd]                │
│           Ghi chú: [textarea]            [💾 Lưu]         │
│           Ảnh: [thumb][thumb][+Thêm]                     │
│           Cập nhật 07:32 bởi thangvv.st                  │
└──────────────────────────────────────────────────────────┘
[Phase 2 ...]
[Phase 3 ...]
```

#### 8.2. Status pill CSS

| Trạng thái | Màu | Class |
|-----------|-----|-------|
| Chưa làm | xám | `.status-chua-lam` |
| Đang làm | xanh dương | `.status-dang-lam` |
| Xong | xanh lá | `.status-xong` |
| Bỏ qua | cam | `.status-bo-qua` |

Pill là `<select>` trong edit mode, `<span>` trong view-only.

#### 8.3. Behavior

- **Đổi trạng thái**: click pill → dropdown → chọn → auto-save (gọi `api_checklist_luu` ngay, không cần nút Lưu). Toast "Đã lưu ✓".
- **Ghi chú**: `<textarea>` autosize, save khi blur HOẶC click nút 💾. Tránh auto-save mỗi ký tự.
- **Photo**: drag-drop zone + click. Thumbnail preview sau upload. Click thumb → modal zoom. "✕" delete với confirm.
- **Ngày/đội nav**: date picker + prev/next/today + đội dropdown (admin đổi được, to_truong khóa về đội mình). URL chứa `?date=&doi=` để bookmark/share.
- **View-only mode (xem đội khác)**: pill = `<span>`, textarea disable, ẩn drop zone + ✕ ảnh. Banner nhẹ "Chỉ xem — bạn không phải tổ trưởng đội này".
- **Empty state**: tất cả mục mặc định `chua_lam`, hint "Chọn trạng thái + thêm ghi chú để bắt đầu checklist hôm nay".

#### 8.4. Excel export

Button "📥 Xuất Excel" → `GET /download/checklist-report?from=&to=&doi=`.
- Sheet 1: state (date × đội × phase × item × status × note × người × thời điểm).
- Sheet 2: photos (date × đội × item × filename × url).
- Theo pattern `quality_routes.py:502-549` (multi-sheet `openpyxl`).

#### 8.5. Sidebar menu item (thêm vào `base.html`)

```html
{% if is_endpoint_enabled('checklist.page_checklist') %}
<a href="{{ url_for('checklist.page_checklist') }}"
   class="menu-item {% if active_page == 'checklist' %}active{% endif %}">
  <i class="fas fa-clipboard-check"></i>
  <span>14. Checklist tổ trưởng</span>
</a>
{% endif %}
```

Sidebar hiện dùng số thứ tự không liên tục (cao nhất là `13`). Đặt số `14` cho checklist, vị trí sau nhóm kiemsoat hiện tại (`/brcd`, `/pttb`, `/shc-cts`).

## Multi-instance deploy

**Mặc định bật** (không cần khai báo trong `units.yaml`):
- Endpoint `checklist.page_checklist` không nằm trong `DISABLED_PAGE_ENDPOINTS` → tất cả instance tự động có trang.
- DB + uploads dir tự tạo trong `_ensure_checklist_schema()` lần đầu truy cập.

**Tắt cho instance không dùng** (opt-out): thêm vào `disabled_endpoints` trong `deploy/units.yaml`:

```yaml
- code: hoai_duc
  disabled_endpoints:
    - checklist.page_checklist
    - checklist.api_checklist_data
    - checklist.api_checklist_luu
    - checklist.api_checklist_upload_photo
    - checklist.api_checklist_delete_photo
    - checklist.download_checklist_report
```

Chạy `python3 scripts/generate_instances.py --units-file deploy/units.yaml --output-dir deploy/generated` rồi restart.

**Users migration per-instance**: admin của mỗi TTVT đăng nhập `/admin/users`, gán `role=to_truong` + `doi=Đội X` cho các tổ trưởng.

**Systemd**: KHÔNG thay đổi template `dashv4@<unit>.service` — không có cron mới, không có env mới bắt buộc.

## Cập nhật docs (bắt buộc theo doc-sync rule)

**Sửa:**
- `docs/00-doc-index.md` — thêm dòng reference tới `docs/12-checklist-to-truong.md` trong section "Khi cần làm gì thì đọc gì" + mục "Kiểm soát tổ trưởng / checklist hằng ngày".
- `docs/04-mapping-route-va-du-lieu.md` — thêm 6 dòng (page_checklist, api_checklist_data, api_checklist_luu, api_checklist_upload_photo, api_checklist_delete_photo, download_checklist_report) với `supports_date=n/a`, `status=migrated`, `ten_bang_du_lieu=checklist.db::checklist_state/photos/audit_log`.
- `docs/08-trang-thai-thuc-thi.md` — thêm **section 9**: "KT tổ trưởng — Checklist hằng ngày" (mô tả endpoint, trạng thái migrated, mẫu template cố định, photo upload mới).
- `AGENTS.md` — cập nhật đoạn "Architecture" thêm dòng mô tả `checklist_routes.py` (blueprint thứ 10) + paragraph về photo upload hạ tầng mới.

**Tạo mới:**
- `docs/12-checklist-to-truong.md` (~250 dòng, mirror cấu trúc `docs/11`):
  - §1 Tổng quan (mục đích, khác gì với kiemsoat)
  - §2 Nguyên lý thiết kế
  - §3 CSDL (3 bảng) + đường dẫn
  - §4 Biến môi trường
  - §5 Nguồn metric (5 fetcher + SQL tham chiếu)
  - §6 Photo upload (validation chain, storage, serving, multi-worker)
  - §7 API endpoints
  - §8 Backup & Restore (checklist.db `.backup` + `checklist_uploads/` rsync)
  - §9 Troubleshooting
  - §10 Cạm bẫy (case-sensitivity `doi_vt`, lazy migration users.xlsx, max photo limit)
  - §11 Test
  - §12 File tham chiếu nhanh
  - §13 Checklist bảo dưỡng định kỳ

## Tests

| File test | Số test ước tính | Phủ |
|-----------|------------------|-----|
| `tests/test_checklist_state.py` | ~12 | UPSERT state, DELETE khi empty, audit log append, SELECT by (ngay,doi) |
| `tests/test_checklist_photos.py` | ~10 | Validation chain (ext/mime/size), per-item limit, ownership check, delete |
| `tests/test_checklist_metric_fetchers.py` | ~8 | Mỗi fetcher với mock SQLite (brcd_phieu sample), fail-safe (None on locked), column case |
| `tests/test_auth_to_truong.py` | ~6 | Decorator chặn đúng role, payload `doi` mismatch → 403, admin bypass, session `doi` |
| `tests/test_checklist_routes.py` | ~10 | GET data có đủ 3 phases, POST save thành công, view-only mode (can_edit=false), download report |
| **Tổng cộng** | **~46 test mới** | suite 47 → ~93 |

**Pattern test** (theo `tests/test_brcd_kiemsoat.py`):
- Build SQLite trong `tmp_path`, monkeypatch `config.CHECKLIST_DB_PATH` + `CHECKLIST_UPLOAD_DIR`.
- Reset `_checklist_schema_ready_path = None` ở teardown.
- Mock `brcd_kiemsoat.db` sample data cho metric fetcher test.
- Flask test client cho endpoint test, login giả bằng `client.session_transaction()` set session.

**Verification command:** `python3 -m pytest tests/test_checklist_*.py tests/test_auth_to_truong.py -v` (riêng) hoặc `python3 -m pytest tests/ -q` (toàn bộ).

## Verification checklist trước khi merge

- [ ] `python3 -m py_compile blueprints/checklist_routes.py` (syntax check)
- [ ] `python3 -m pytest tests/ -q` (all green, ~93 tests)
- [ ] `python3 dashboard.py` start không lỗi (import blueprint OK)
- [ ] Manual smoke: mở `/checklist`, đổi trạng thái 1 mục, thêm 1 ghi chú, upload 1 ảnh, reload page → state persist.
- [ ] Manual smoke quyền: đăng nhập user `role=user` → không thấy nút Save; đăng nhập `to_truong` đội khác → xem được nhưng không sửa.
- [ ] Manual smoke export: click Xuất Excel → file có 2 sheet (state + photos).

## Rollout đề nghị

1. **Pilot 1 tuần ở `son_tay`**: deploy lên instance `dashv4@son_tay`, admin gán to_truong + doi cho 1-2 tổ trưởng, theo dõi log + DB size.
2. **Sau 1 tuần ổn định**: chạy `generate_instances.py` + restart tất cả 18 instance. Mỗi TTVT tự gán role cho tổ trưởng.
3. **Sau 2 tuần**: review audit log, thống kê adoption, thu thập feedback tổ trưởng.
4. **v2 candidates** (không trong scope MVP): admin UI sửa template, auto-cleanup ảnh, notification, PWA.

## Scope MVP (chốt lại)

✅ Trang `/checklist` mới, 3 phases × 15 mục hard-code
✅ Status (4 giá trị) + ghi chú text + photo upload (max 5 ảnh/item, 5MB/ảnh)
✅ Per-đội scope, role `to_truong` mới, view-all-edit-own
✅ 5 metric fetcher hiển thị số liệu live
✅ Audit log append-only
✅ Excel export 2 sheet
✅ Multi-instance auto-isolation
✅ Tests ~46 mới, docs/12 mới

❌ Out of scope (v2): admin UI sửa template, auto-cleanup ảnh, notification, PWA.

## Tham chiếu

- `docs/11-kiemsoat-to-truong-van-hanh.md` — pattern tham chiếu (schema setup, write/read conn, audit trail concept)
- `blueprints/operations_routes.py:115-232` — write/read connection helpers + idempotent schema setup
- `app_helpers.py:61-73` — `csrf_protect` decorator
- `app_helpers.py:87-110` — `safe_file_response` / `safe_directory_response` cho serving ảnh
- `auth.py:195-226` — pattern decorator `login_required` / `admin_required`
- `repositories/sqlite_runtime.py` — pattern read-only connection URI
- `quality_routes.py:502-549` — pattern multi-sheet Excel export với `openpyxl`
