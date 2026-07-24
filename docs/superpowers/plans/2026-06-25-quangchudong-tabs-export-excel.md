# Quang Chủ Động Tabs Export Excel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two derived Quang chủ động tabs and a downloadable Excel report with one sheet per tab.

**Architecture:** Keep `/api/quangchudong/dashboard` as the data source. Add small pure helpers in `blueprints/quangchudong_routes.py` for server-side row classification and Excel sheet generation, and mirror the same classification in `templates/quangchudong.html` for client-side rendering. Do not change the existing `Cảnh báo DOWN`, `Sự cố diện rộng`, `Pattern exclusion`, or `Port OLT Down` behavior.

**Tech Stack:** Flask blueprint routes, pandas `ExcelWriter` with `xlsxwriter`, in-memory `io.BytesIO`, existing login/session auth, existing client-side HTML/JS tab pattern.

## Global Constraints

- Run commands from repo root: `/home/vtst/dashv4`.
- Use system pytest: `python3 -m pytest tests/`; do not use `venv/bin/pytest`.
- Preserve `runtime_limits.py` import ordering; this task does not add entrypoints.
- No new dependencies; `pandas`, `xlsxwriter`, and `openpyxl` are already present in runtime.
- Keep UI copy Vietnamese.
- Keep existing tab `Cảnh báo DOWN` unchanged.
- New data derives from `payload.active.alerts`, not `active.excluded`, `wide_area_groups`, `pattern_exclusions`, or `port_down_groups`.
- Excel export must include all tab data as separate sheets.
- Do not update date-filtering docs unless this task changes `report_history.db` mapping; it does not.

---

## File Structure

- Modify `blueprints/quangchudong_routes.py`: add reusable classification/export helpers and new authenticated download route.
- Modify `templates/quangchudong.html`: add two tabs, two table panes, a download button, and client-side render/filter helpers.
- Modify `tests/test_quangchudong_routes.py`: add tests for tab HTML, export route, sheet names, row classification, and auth behavior.

---

### Task 1: Backend Helpers For OFF-In-Day And Sleeping Rows

**Files:**
- Modify: `blueprints/quangchudong_routes.py`
- Test: `tests/test_quangchudong_routes.py`

**Interfaces:**
- Consumes: rows from `get_quangchudong_cache().get_dashboard_payload()['active']['alerts']`.
- Produces:
  - `_get_event_time(row: dict) -> datetime | None`
  - `_filter_off_today_rows(rows: list[dict], cutoff_time: datetime) -> list[dict]`
  - `_with_sleep_days(row: dict, now: datetime) -> dict`
  - `_filter_sleeping_rows(rows: list[dict], cutoff_time: datetime, now: datetime) -> list[dict]`

- [ ] **Step 1: Add failing tests for row classification**

Append to `tests/test_quangchudong_routes.py`:

```python
def test_quangchudong_helpers_split_off_today_and_sleeping_rows():
    rows = [
        {'ma_tb': 'AFTER', 'first_off_time': '2026-05-13T06:00:00'},
        {'ma_tb': 'ALERT_FALLBACK', 'first_off_time': '', 'alert_time': '2026-05-13 07:15:00'},
        {'ma_tb': 'BEFORE', 'first_off_time': '2026-05-12 05:30:00'},
        {'ma_tb': 'BAD_TIME', 'first_off_time': 'not-a-date'},
    ]
    cutoff = datetime(2026, 5, 13, 6, 0, 0)
    now = datetime(2026, 5, 15, 8, 30, 0)

    off_today = quangchudong_routes._filter_off_today_rows(rows, cutoff)
    sleeping = quangchudong_routes._filter_sleeping_rows(rows, cutoff, now)

    assert [row['ma_tb'] for row in off_today] == ['ALERT_FALLBACK', 'AFTER']
    assert [row['ma_tb'] for row in sleeping] == ['BEFORE']
    assert sleeping[0]['thoi_gian_ngu'] == 3
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_helpers_split_off_today_and_sleeping_rows -q`

Expected: FAIL because helper functions are not defined.

- [ ] **Step 3: Implement minimal backend helpers**

Add near the existing `_parse_alert_time` helper in `blueprints/quangchudong_routes.py`:

```python
def _get_event_time(row):
    return _parse_alert_time(row.get('first_off_time') or row.get('alert_time'))


def _filter_off_today_rows(rows, cutoff_time):
    filtered = []
    for row in rows:
        event_time = _get_event_time(row)
        if event_time is not None and event_time >= cutoff_time:
            filtered.append(row)
    filtered.sort(key=lambda row: _get_event_time(row) or datetime.min, reverse=True)
    return filtered


def _with_sleep_days(row, now):
    enriched = dict(row)
    event_time = _get_event_time(row)
    if event_time is None:
        enriched['thoi_gian_ngu'] = ''
        return enriched
    diff_days = max(1, int((now - event_time).total_seconds() // 86400))
    enriched['thoi_gian_ngu'] = diff_days
    return enriched


def _filter_sleeping_rows(rows, cutoff_time, now):
    filtered = []
    for row in rows:
        event_time = _get_event_time(row)
        if event_time is not None and event_time < cutoff_time:
            filtered.append(_with_sleep_days(row, now))
    filtered.sort(key=lambda row: row.get('thoi_gian_ngu') or 0, reverse=True)
    return filtered
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_helpers_split_off_today_and_sleeping_rows -q`

Expected: PASS.

---

### Task 2: Excel Export Route With One Sheet Per Tab

**Files:**
- Modify: `blueprints/quangchudong_routes.py`
- Test: `tests/test_quangchudong_routes.py`

**Interfaces:**
- Consumes helpers from Task 1.
- Produces authenticated route `GET /download/quangchudong-report` mapped to endpoint `quangchudong.download_quangchudong_report`.

- [ ] **Step 1: Add failing export test**

Append imports near the top of `tests/test_quangchudong_routes.py`:

```python
import io
import pandas as pd
```

Append this fake cache and test:

```python
class _FakeExportQuangChuDongCache:
    def get_dashboard_payload(self):
        return {
            'active': {
                'alerts': [
                    {
                        'ma_tb': 'TODAY',
                        'ten_tb': 'Trong ngày',
                        'first_off_time': '2026-05-13 07:00:00',
                        'doi_vt': 'Tổ A',
                    },
                    {
                        'ma_tb': 'SLEEP',
                        'ten_tb': 'Ngủ dài',
                        'first_off_time': '2026-05-10 05:00:00',
                        'doi_vt': 'Tổ B',
                    },
                ],
                'excluded': [],
                'sources': [],
                'cache': {'refreshed_at': '13/05/2026 08:00:00'},
            },
            'wide_area_groups': [{'parent_port_key': 'PARENT1', 'down_count': 3}],
            'pattern_exclusions': [{'ma_tb': 'PATTERN1', 'exclude_reason': 'Trong danh sách tắt chủ động'}],
            'port_down_groups': [{'parent_port_key': 'PORT1', 'down_count': 5}],
        }


def test_quangchudong_download_excel_contains_all_tab_sheets(monkeypatch):
    monkeypatch.setattr(
        quangchudong_routes,
        'get_quangchudong_cache',
        lambda: _FakeExportQuangChuDongCache(),
    )
    monkeypatch.setattr(quangchudong_routes, 'datetime', _FixedDateTime)

    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/download/quangchudong-report')

    assert response.status_code == 200
    assert response.mimetype == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    assert response.headers['Content-Disposition'].startswith('attachment;')

    workbook = pd.ExcelFile(io.BytesIO(response.data))
    assert workbook.sheet_names == [
        'Canh_bao_DOWN',
        'OFF_trong_ngay',
        'Thue_bao_ngu',
        'Su_co_dien_rong',
        'Pattern_exclusion',
        'Port_OLT_Down',
    ]

    off_today = pd.read_excel(workbook, sheet_name='OFF_trong_ngay')
    sleeping = pd.read_excel(workbook, sheet_name='Thue_bao_ngu')
    assert off_today['Mã TB'].tolist() == ['TODAY']
    assert sleeping['Mã TB'].tolist() == ['SLEEP']
    assert sleeping['Thời gian ngủ'].tolist() == [3]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_download_excel_contains_all_tab_sheets -q`

Expected: FAIL with 404 for missing route.

- [ ] **Step 3: Implement Excel helpers and route**

Add imports at the top of `blueprints/quangchudong_routes.py`:

```python
import io
import pandas as pd
from flask import Blueprint, jsonify, render_template, send_file, session
```

Add helpers below Task 1 helpers:

```python
_SUBSCRIBER_EXPORT_COLUMNS = [
    ('ma_tb', 'Mã TB'),
    ('ten_tb', 'Tên TB'),
    ('diachi_ld', 'Địa chỉ'),
    ('dienthoai_lh', 'SĐT'),
    ('ten_nvkt_db', 'NVKT'),
    ('doi_vt', 'Đội VT'),
    ('olt_name', 'OLT'),
    ('sa', 'SA'),
    ('first_off_time', 'Thời gian OFF'),
    ('duration_text', 'Thời lượng'),
    ('ngay_bh', 'Phiếu BH Lúc'),
]

_GROUP_EXPORT_COLUMNS = [
    ('parent_port_key', 'Parent port'),
    ('olt_name', 'OLT'),
    ('sa', 'SA'),
    ('down_count', 'Số thuê bao'),
    ('start_time', 'Thời điểm bắt đầu'),
]


def _rows_to_export_frame(rows, columns):
    records = []
    for row in rows:
        records.append({label: row.get(key, '') for key, label in columns})
    return pd.DataFrame(records, columns=[label for _, label in columns])


def _subscriber_export_frame(rows, *, include_sleep_days=False):
    columns = list(_SUBSCRIBER_EXPORT_COLUMNS)
    if include_sleep_days:
        columns.append(('thoi_gian_ngu', 'Thời gian ngủ'))
    return _rows_to_export_frame(rows, columns)
```

Add route near other `/api/quangchudong/*` routes:

```python
@quangchudong_bp.route('/download/quangchudong-report')
@login_required
def download_quangchudong_report():
    payload = get_quangchudong_cache().get_dashboard_payload()
    active = payload.get('active') or {}
    alerts = active.get('alerts') or []
    cutoff_time = _today_six_am()
    now = datetime.now()

    sheets = [
        ('Canh_bao_DOWN', _subscriber_export_frame(alerts)),
        ('OFF_trong_ngay', _subscriber_export_frame(_filter_off_today_rows(alerts, cutoff_time))),
        ('Thue_bao_ngu', _subscriber_export_frame(
            _filter_sleeping_rows(alerts, cutoff_time, now),
            include_sleep_days=True,
        )),
        ('Su_co_dien_rong', _rows_to_export_frame(payload.get('wide_area_groups') or [], _GROUP_EXPORT_COLUMNS)),
        ('Pattern_exclusion', _subscriber_export_frame(payload.get('pattern_exclusions') or [])),
        ('Port_OLT_Down', _rows_to_export_frame(payload.get('port_down_groups') or [], _GROUP_EXPORT_COLUMNS)),
    ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, frame in sheets:
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f'quang_chu_dong_{datetime.now():%Y%m%d_%H%M}.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
```

- [ ] **Step 4: Run export test to verify it passes**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_download_excel_contains_all_tab_sheets -q`

Expected: PASS.

---

### Task 3: Add Two New UI Tabs

**Files:**
- Modify: `templates/quangchudong.html`
- Test: `tests/test_quangchudong_routes.py`

**Interfaces:**
- Consumes existing `activeData.alerts` in `loadData()`.
- Produces DOM tables with IDs:
  - `off-today-table`
  - `sleeping-table`

- [ ] **Step 1: Add failing HTML test for new tabs**

Append to `tests/test_quangchudong_routes.py`:

```python
def test_quangchudong_page_has_new_off_today_and_sleeping_tabs():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/quangchudong')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    assert 'Cảnh báo OFF trong ngày' in html
    assert 'Thuê bao ngủ theo ngày' in html
    assert 'id="off-today-table"' in html
    assert 'id="sleeping-table"' in html
    assert 'Thời gian ngủ' in html
    assert 'renderOffTodayAlerts' in html
    assert 'renderSleepingSubscribers' in html
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_page_has_new_off_today_and_sleeping_tabs -q`

Expected: FAIL because HTML does not contain the new tabs.

- [ ] **Step 3: Add tab buttons**

In `templates/quangchudong.html`, after the `Cảnh báo DOWN` button, add:

```html
<button class="nav-link" type="button" data-tab="off-today">
    Cảnh báo OFF trong ngày <span class="badge badge-danger" id="off-today-tab-count">-</span>
</button>
<button class="nav-link" type="button" data-tab="sleeping">
    Thuê bao ngủ theo ngày <span class="badge badge-secondary" id="sleeping-tab-count">-</span>
</button>
```

- [ ] **Step 4: Add tab panes and tables**

In `templates/quangchudong.html`, after `alerts-pane`, add two panes. The first should mirror `alert-table` columns and use `off-today-table`; the second should mirror `alert-table` and add a final `Thời gian ngủ` column.

Required IDs and filter targets:

```html
<div class="tab-pane" id="off-today-pane">
    <div class="table-filters" id="off-today-filters">
        <input type="text" class="filter-input" data-table="off-today-table" data-column="0" placeholder="Mã TB">
        <input type="text" class="filter-input" data-table="off-today-table" data-column="1" placeholder="Tên TB">
        <input type="text" class="filter-input" data-table="off-today-table" data-column="4" placeholder="NVKT">
        <input type="text" class="filter-input" data-table="off-today-table" data-column="5" placeholder="Đội VT">
        <input type="text" class="filter-input" data-table="off-today-table" data-column="6" placeholder="OLT">
    </div>
    <table class="data-table" id="off-today-table" style="table-layout: fixed; width: 100%;">
        <thead><tr><th>Mã TB</th><th>Tên TB</th><th>Địa chỉ</th><th>SĐT</th><th>NVKT</th><th>Đội VT</th><th>OLT</th><th>SA</th><th>Thời gian OFF</th><th>Thời lượng</th><th>Phiếu BH Lúc</th></tr></thead>
        <tbody></tbody>
    </table>
</div>

<div class="tab-pane" id="sleeping-pane">
    <div class="table-filters" id="sleeping-filters">
        <input type="text" class="filter-input" data-table="sleeping-table" data-column="0" placeholder="Mã TB">
        <input type="text" class="filter-input" data-table="sleeping-table" data-column="1" placeholder="Tên TB">
        <input type="text" class="filter-input" data-table="sleeping-table" data-column="4" placeholder="NVKT">
        <input type="text" class="filter-input" data-table="sleeping-table" data-column="5" placeholder="Đội VT">
        <input type="text" class="filter-input" data-table="sleeping-table" data-column="11" placeholder="Thời gian ngủ">
    </div>
    <table class="data-table" id="sleeping-table" style="table-layout: fixed; width: 100%;">
        <thead><tr><th>Mã TB</th><th>Tên TB</th><th>Địa chỉ</th><th>SĐT</th><th>NVKT</th><th>Đội VT</th><th>OLT</th><th>SA</th><th>Thời gian OFF</th><th>Thời lượng</th><th>Phiếu BH Lúc</th><th>Thời gian ngủ</th></tr></thead>
        <tbody></tbody>
    </table>
</div>
```

- [ ] **Step 5: Add client-side classification/render helpers**

Add below `getTodaySixAm()`:

```javascript
function getEventTime(row) {
    return row.first_off_time || row.alert_time || '';
}

function isOffToday(row) {
    const eventTime = parseLocalDateTime(getEventTime(row));
    return eventTime && eventTime >= getTodaySixAm();
}

function getSleepDays(row) {
    const eventTime = parseLocalDateTime(getEventTime(row));
    if (!eventTime) return '';
    const diffDays = Math.floor((new Date() - eventTime) / 86400000);
    return Math.max(1, diffDays);
}

function isSleepingSubscriber(row) {
    const eventTime = parseLocalDateTime(getEventTime(row));
    return eventTime && eventTime < getTodaySixAm();
}
```

Add render functions near `renderAlerts`:

```javascript
function renderOffTodayAlerts(alerts) {
    const rows = alerts.filter(isOffToday).sort((a, b) => parseLocalDateTime(getEventTime(b)) - parseLocalDateTime(getEventTime(a)));
    const tbody = document.querySelector('#off-today-table tbody');
    document.getElementById('off-today-tab-count').textContent = rows.length;
    if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="11" class="no-data">Không có cảnh báo OFF trong ngày từ 06:00</td></tr>';
        return;
    }
    tbody.innerHTML = rows.map(row => `
        <tr data-duration-minutes="${getDurationMinutes(getEventTime(row)) ?? ''}" data-event-time="${getEventTime(row)}">
            <td>${row.ma_tb || 'N/A'}</td><td>${row.ten_tb || 'N/A'}</td><td title="${row.diachi_ld || ''}">${(row.diachi_ld || '').substring(0, 40)}</td><td>${row.dienthoai_lh || ''}</td><td>${row.ten_nvkt_db || ''}</td><td>${row.doi_vt || ''}</td><td>${row.olt_name || ''}</td><td>${row.sa || ''}</td><td>${formatDateTime(getEventTime(row))}</td><td class="duration-badge">${formatDuration(getEventTime(row))}</td><td>${row.ngay_bh || ''}</td>
        </tr>
    `).join('');
}

function renderSleepingSubscribers(alerts) {
    const rows = alerts.filter(isSleepingSubscriber).sort((a, b) => getSleepDays(b) - getSleepDays(a));
    const tbody = document.querySelector('#sleeping-table tbody');
    document.getElementById('sleeping-tab-count').textContent = rows.length;
    if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="12" class="no-data">Không có thuê bao ngủ theo ngày</td></tr>';
        return;
    }
    tbody.innerHTML = rows.map(row => `
        <tr data-duration-minutes="${getDurationMinutes(getEventTime(row)) ?? ''}" data-event-time="${getEventTime(row)}">
            <td>${row.ma_tb || 'N/A'}</td><td>${row.ten_tb || 'N/A'}</td><td title="${row.diachi_ld || ''}">${(row.diachi_ld || '').substring(0, 40)}</td><td>${row.dienthoai_lh || ''}</td><td>${row.ten_nvkt_db || ''}</td><td>${row.doi_vt || ''}</td><td>${row.olt_name || ''}</td><td>${row.sa || ''}</td><td>${formatDateTime(getEventTime(row))}</td><td class="duration-badge">${formatDuration(getEventTime(row))}</td><td>${row.ngay_bh || ''}</td><td>${getSleepDays(row)} ngày</td>
        </tr>
    `).join('');
}
```

- [ ] **Step 6: Wire render calls and filters**

In `loadData()`, after `renderAlerts(activeData.alerts || []);`, add:

```javascript
renderOffTodayAlerts(activeData.alerts || []);
renderSleepingSubscribers(activeData.alerts || []);
```

Update filter list:

```javascript
['alert-table', 'off-today-table', 'sleeping-table', 'wide-area-table', 'pattern-table', 'port-down-table'].forEach(filterTable);
```

- [ ] **Step 7: Run HTML test to verify it passes**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_page_has_new_off_today_and_sleeping_tabs -q`

Expected: PASS.

---

### Task 4: Add Excel Download Button In UI

**Files:**
- Modify: `templates/quangchudong.html`
- Test: `tests/test_quangchudong_routes.py`

**Interfaces:**
- Consumes route `quangchudong.download_quangchudong_report` from Task 2.
- Produces visible download link on `/quangchudong`.

- [ ] **Step 1: Add failing HTML test for download link**

Append to `tests/test_quangchudong_routes.py`:

```python
def test_quangchudong_page_has_excel_download_link():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/quangchudong')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    assert '/download/quangchudong-report' in html
    assert 'Kết xuất Excel' in html
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_page_has_excel_download_link -q`

Expected: FAIL because link does not exist.

- [ ] **Step 3: Add download link**

In `templates/quangchudong.html`, near the page header or before the tabs, add:

```html
<div style="display: flex; justify-content: flex-end; margin-bottom: 1rem;">
    <a class="btn btn-success" href="{{ url_for('quangchudong.download_quangchudong_report') }}">
        <i class="fas fa-file-excel"></i> Kết xuất Excel
    </a>
</div>
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_quangchudong_routes.py::test_quangchudong_page_has_excel_download_link -q`

Expected: PASS.

---

### Task 5: Auth, Policy, And Regression Verification

**Files:**
- Modify: `tests/test_quangchudong_routes.py`

**Interfaces:**
- Consumes export route from Task 2.
- Produces confidence that downloads are protected by existing login policy.

- [ ] **Step 1: Add unauthenticated download test**

Append to `tests/test_quangchudong_routes.py`:

```python
def test_quangchudong_download_requires_login():
    client = app.test_client()

    response = client.get('/download/quangchudong-report')

    assert response.status_code in {302, 401}
```

- [ ] **Step 2: Run targeted quangchudong tests**

Run: `python3 -m pytest tests/test_quangchudong_routes.py -q`

Expected: all tests pass.

- [ ] **Step 3: Compile changed Python files**

Run: `python3 -m py_compile blueprints/quangchudong_routes.py tests/test_quangchudong_routes.py`

Expected: no output and exit code 0.

- [ ] **Step 4: Run full test suite**

Run: `python3 -m pytest tests/`

Expected: all tests pass.

---

## Self-Review

- Spec coverage: The plan adds the two requested tabs, keeps `Cảnh báo DOWN`, and adds a one-file Excel export with sheets for all tabs.
- Placeholder scan: No TBD/TODO placeholders remain.
- Type consistency: Backend helper names are used consistently across tests and route implementation.
- Scope check: This is one coherent feature because the UI tabs and export share row classification logic.
