# Changelog

Tất cả các thay đổi quan trọng của Dashboard Application sẽ được ghi lại trong file này.

Định dạng dựa trên [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [2.2.0] - 2026-04-21

### Added
- I1.5 và I1.5 K2 chuyển sang đọc trực tiếp từ `report_history.db`.
- Bổ sung loader SQLite cho các view I1.5 trong `repositories/dashboard_views.py`.
- Mở lại hai page `/i15` và `/i15k2` cùng các API liên quan.

### Fixed
- `repositories/sqlite_runtime.py` mở DB read-only bằng `mode=ro&immutable=1` để tránh lỗi journal/WAL khi dashboard query view.
- `quality_routes.py` không còn phụ thuộc Excel cũ cho I1.5 và SHC variation.

## [2.1.0] - 2025-10-28

### Added
- **2 biểu đồ NVKT mới** cho section "Thực tăng Fiber, MyTV":
  - Fiber NVKT: `fiber_thuctang_nvkt.png` (472KB)
    - Thống kê thực tăng Fiber theo Nhân viên kinh tế
  - MyTV NVKT: `mytv_thuctang_nvkt.png` (471KB)
    - Thống kê thực tăng MyTV theo Nhân viên kinh tế
- Cấu trúc mới cho section "Thực tăng Fiber, MyTV":
  - Fiber: 2 biểu đồ (PTTB + NVKT)
  - MyTV: 2 biểu đồ (PTTB + NVKT)

### Updated
- **dashboard.html**:
  - Thêm `<div class="chart-card">` cho fiber_thuctang_nvkt.png (lines 1469-1476)
  - Thêm `<div class="chart-card">` cho mytv_thuctang_nvkt.png (lines 1493-1500)
  - Spacing: `margin-top: 20px` giữa 2 biểu đồ của mỗi dịch vụ

### Technical Details
- **File paths**:
  - `/baocaohanoi/chart/thuc_tang_fiber/fiber_thuctang_nvkt.png`
  - `/baocaohanoi/chart/thuc_tang_mytv/mytv_thuctang_nvkt.png`
- **Endpoints**:
  - `/baocaohanoi/chart/thuc_tang_fiber/fiber_thuctang_nvkt.png`
  - `/baocaohanoi/chart/thuc_tang_mytv/mytv_thuctang_nvkt.png`
- **Features**:
  - Click để zoom biểu đồ (onclick="zoomImage(this.src)")
  - Responsive design (max-width: 100%; height: auto)

### Breaking Changes
- None

### Known Issues
- None

---

## [2.0.0] - 2025-10-28

### Added
- **Mục menu mới "Thực tăng Fiber, MyTV"** (Menu item #3)
  - Thay thế mục "Thống kê Ticket" cũ
  - Hiển thị 2 biểu đồ thực tăng thuê bao:
    - Thực tăng Fiber (thuc_tang_fiber_pttb.png)
    - Thực tăng MyTV (thuc_tang_mytv_pttb.png)
  - Icon: fa-chart-area
  - Section ID: `thuctang-section`

- **Route mới trong dashboard.py**:
  ```python
  @app.route('/baocaohanoi/chart/<path:filepath>')
  def serve_baocaohanoi_chart(filepath):
  ```
  - Phục vụ file chart từ `/home/vtst/baocaohanoi/chart/`
  - Hỗ trợ nested path với `<path:filepath>`
  - Cache headers: no-cache, no-store, must-revalidate

- **Data sources mới**:
  - `/home/vtst/baocaohanoi/chart/thuc_tang_fiber/thuc_tang_fiber_pttb.png` (210KB)
  - `/home/vtst/baocaohanoi/chart/thuc_tang_mytv/thuc_tang_mytv_pttb.png` (222KB)

- **Documentation**:
  - Cập nhật README.md với thông tin đầy đủ về tính năng mới
  - Thêm CHANGELOG.md để tracking thay đổi
  - Thêm phần troubleshooting cho "Thực tăng" section

### Fixed
- **Lỗi load chart "Tồn sửa chữa theo địa bàn"** trong BRCD section
  - **Root cause**: Tên file chart không khớp với team ID trong team_config.py
  - **Changes**:
    - Cập nhật `data-chart` attribute từ tên ngắn sang team ID format
    - Cập nhật thứ tự hiển thị: Phúc Thọ → Sơn Tây → Quảng Oai → Suối Hai
    - Cập nhật default image từ Quảng Oai sang Phúc Thọ
  - **Details**:
    | Trước | Sau |
    |-------|-----|
    | `data-chart="quangoai"` | `data-chart="ToKT_QuangOai"` |
    | `data-chart="sontay"` | `data-chart="ToKT_SonTay"` |
    | `data-chart="suoihai"` | `data-chart="ToKT_SuoiHai"` |
    | `data-chart="phuctho"` | `data-chart="ToKT_PhucTho"` |
    | `/chart/Chart_diaban_quangoai.png` | `/chart/Chart_diaban_ToKT_QuangOai.png` |
  - **Affected files**:
    - `/home/vtst/dash/dashboard.html` (lines 1620-1629)
    - JavaScript chart switcher (line 1771) unchanged (already correct)

### Changed
- **Menu sidebar structure**:
  - Item #3 changed from "Thống kê Ticket" to "Thực tăng Fiber, MyTV"
  - Tooltip updated: "Thống kê Ticket" → "Thực tăng Fiber, MyTV"
  - Icon changed: fa-chart-bar → fa-chart-area

- **Dashboard section IDs**:
  - `thongke-section` content moved/replaced by `thuctang-section`
  - Original "Thống kê Ticket" section still exists but not linked in menu

- **File locations**:
  - Dashboard app: `/home/vtst/dash/` (primary)
  - Synced copy: `/home/vtst/onev2/` (for backward compatibility)

### Technical Details

#### Modified Files
1. **dashboard.py**:
   - Line 57-67: New route `serve_baocaohanoi_chart()`
   - Uses `os.path.join('/home/vtst/baocaohanoi/chart', filepath)`

2. **dashboard.html**:
   - Line 1204-1207: Menu item update
   - Line 1442-1486: New section content
   - Line 1620-1629: Fixed chart data attributes

#### New Dependencies
- No new Python packages required
- Requires read access to `/home/vtst/baocaohanoi/chart/`

#### Configuration Changes
- No changes to PORT (still 5009)
- No changes to BASE_DATA_PATH (still `/home/vtst/onev2/`)
- Added access to `/home/vtst/baocaohanoi/` (hardcoded path)

### Migration Notes

**For existing users**:
1. No database migration required
2. No configuration changes needed
3. Simply restart dashboard to apply changes:
   ```bash
   cd /home/vtst/dash
   ./start_dashboard.sh
   ```
4. Clear browser cache if menu not updating (Ctrl+Shift+R)

**Rollback procedure** (if needed):
1. Copy old files from git/backup
2. Restart dashboard
3. No data loss (only UI changes)

### Known Issues
- None at this time

### Security
- No security updates in this release
- Route `/baocaohanoi/chart/` uses hardcoded path (not configurable)
- File access limited to PNG files via browser request

---

## [1.0.0] - 2025-10-27

### Initial Release
- Dashboard application separated from main codebase
- BRCD section: Điều hành sửa chữa
- PTTB section: Phát triển thuê bao
- Thống kê Ticket section
- Ticket tracking functionality
- Multi-data source support (onev2)
- Port: 5009
- Base path: `/home/vtst/onev2/`
