# Hướng dẫn Tích hợp Trang Giám sát Cảnh báo Quang Chủ động

## Tổng quan

Tài liệu này hướng dẫn cách thêm trang giám sát cảnh báo thuê bao DOWN quang chủ động vào dự án dashboard hiện tại.

---

## 1. Nguồn dữ liệu

### Database: `subscriber_history.db`

**Đường dẫn:** `/home/vtst/do-kiem-chu-dong-v0/subscriber_history.db`

| Bảng | Mục đích | Các cột quan trọng |
|------|----------|-------------------|
| `outage_alerts` | Cảnh báo thuê bao DOWN | `id`, `port_id`, `ma_tb`, `ten_tb`, `olt_name`, `alert_time`, `first_on_time`, `notification_sent`, `diachi_ld`, `dienthoai_lh`, `ten_nvkt_db`, `doi_vt` |
| `recovery_alerts` | Thuê bao đã phục hồi | `id`, `port_id`, `ma_tb`, `ten_tb`, `olt_name`, `outage_time`, `recovery_time`, `outage_duration_minutes`, `notification_sent`, `diachi_ld`, `dienthoai_lh`, `ten_nvkt_db`, `doi_vt` |
| `subscriber_status_history` | Trạng thái hiện tại | `port_id`, `current_state`, `last_status`, `last_check_time`, `consecutive_on_count`, `consecutive_off_count` |

---

## 2. Cấu trúc Routes cần thêm

### 2.1 Route trang chính

```python
@app.route('/quangchudong')
def page_quangchudong():
    """
    Trang Giám sát Cảnh báo Quang Chủ động
    """
    return render_template('quangchudong.html')
```

### 2.2 API - Lấy danh sách thuê bao đang DOWN

```python
@app.route('/api/quangchudong/active')
def get_active_outages():
    """
    API lấy danh sách thuê bao đang DOWN (chưa có recovery)
    
    Logic:
    - Lấy từ outage_alerts
    - Loại trừ các port_id đã có trong recovery_alerts với recovery_time > alert_time
    
    Returns:
        JSON: Danh sách thuê bao đang DOWN
    """
    import sqlite3
    db_path = '/home/vtst/do-kiem-chu-dong-v0/subscriber_history.db'
    
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    query = """
        SELECT o.* 
        FROM outage_alerts o
        WHERE NOT EXISTS (
            SELECT 1 FROM recovery_alerts r 
            WHERE r.port_id = o.port_id 
            AND r.recovery_time > o.alert_time
        )
        ORDER BY o.alert_time DESC
    """
    cursor.execute(query)
    rows = cursor.fetchall()
    conn.close()
    
    return jsonify([dict(row) for row in rows])
```

### 2.3 API - Lấy danh sách thuê bao đã phục hồi

```python
@app.route('/api/quangchudong/recovered')
def get_recovered_alerts():
    """
    API lấy danh sách thuê bao đã phục hồi trong 24h gần nhất
    
    Returns:
        JSON: Danh sách thuê bao đã phục hồi
    """
    import sqlite3
    from datetime import datetime, timedelta
    
    db_path = '/home/vtst/do-kiem-chu-dong-v0/subscriber_history.db'
    cutoff_time = (datetime.now() - timedelta(hours=24)).strftime('%Y-%m-%d %H:%M:%S')
    
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    query = """
        SELECT * FROM recovery_alerts 
        WHERE recovery_time > ?
        ORDER BY recovery_time DESC
    """
    cursor.execute(query, (cutoff_time,))
    rows = cursor.fetchall()
    conn.close()
    
    return jsonify([dict(row) for row in rows])
```

### 2.4 API - Thống kê theo Đội VT

```python
@app.route('/api/quangchudong/stats')
def get_outage_stats():
    """
    API thống kê cảnh báo theo Đội VT
    
    Returns:
        JSON: Thống kê số lượng đang DOWN và đã phục hồi theo đội
    """
    import sqlite3
    db_path = '/home/vtst/do-kiem-chu-dong-v0/subscriber_history.db'
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Thống kê đang DOWN theo đội
    cursor.execute("""
        SELECT doi_vt, COUNT(*) as count
        FROM outage_alerts o
        WHERE NOT EXISTS (
            SELECT 1 FROM recovery_alerts r 
            WHERE r.port_id = o.port_id 
            AND r.recovery_time > o.alert_time
        )
        GROUP BY doi_vt
    """)
    active_stats = {row[0]: row[1] for row in cursor.fetchall()}
    
    # Thống kê phục hồi trong 24h
    cursor.execute("""
        SELECT doi_vt, COUNT(*) as count
        FROM recovery_alerts
        WHERE recovery_time > datetime('now', '-24 hours')
        GROUP BY doi_vt
    """)
    recovered_stats = {row[0]: row[1] for row in cursor.fetchall()}
    
    conn.close()
    
    return jsonify({
        'active_by_doi_vt': active_stats,
        'recovered_24h_by_doi_vt': recovered_stats
    })
```

---

## 3. Template HTML

### File: `templates/quangchudong.html`

```html
{% extends "base.html" %}

{% block title %}Giám sát Quang Chủ động{% endblock %}

{% block content %}
<div class="container-fluid">
    <h2 class="mb-4">
        <i class="fas fa-satellite-dish"></i> Giám sát Cảnh báo Quang Chủ động
    </h2>
    
    <!-- Thống kê tổng quan -->
    <div class="row mb-4">
        <div class="col-md-4">
            <div class="card bg-danger text-white">
                <div class="card-body">
                    <h5>Đang DOWN</h5>
                    <h2 id="active-count">-</h2>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card bg-success text-white">
                <div class="card-body">
                    <h5>Đã phục hồi (24h)</h5>
                    <h2 id="recovered-count">-</h2>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card bg-info text-white">
                <div class="card-body">
                    <h5>Cập nhật lúc</h5>
                    <h4 id="last-update">-</h4>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Tabs -->
    <ul class="nav nav-tabs" id="alertTabs">
        <li class="nav-item">
            <a class="nav-link active" data-bs-toggle="tab" href="#active">
                🔴 Đang DOWN
            </a>
        </li>
        <li class="nav-item">
            <a class="nav-link" data-bs-toggle="tab" href="#recovered">
                🟢 Đã phục hồi
            </a>
        </li>
    </ul>
    
    <div class="tab-content">
        <!-- Tab Đang DOWN -->
        <div class="tab-pane fade show active" id="active">
            <table class="table table-striped" id="active-table">
                <thead>
                    <tr>
                        <th>Mã TB</th>
                        <th>Tên TB</th>
                        <th>SĐT</th>
                        <th>NVKT</th>
                        <th>Đội VT</th>
                        <th>OLT</th>
                        <th>Thời gian DOWN</th>
                        <th>Thời lượng</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>
        
        <!-- Tab Đã phục hồi -->
        <div class="tab-pane fade" id="recovered">
            <table class="table table-striped" id="recovered-table">
                <thead>
                    <tr>
                        <th>Mã TB</th>
                        <th>Tên TB</th>
                        <th>NVKT</th>
                        <th>Đội VT</th>
                        <th>Thời gian DOWN</th>
                        <th>Thời gian phục hồi</th>
                        <th>Thời lượng mất</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>
    </div>
</div>

<script>
// Auto-refresh mỗi 30 giây
const REFRESH_INTERVAL = 30000;

function formatDuration(startTime) {
    const start = new Date(startTime);
    const now = new Date();
    const diff = Math.floor((now - start) / 60000); // phút
    
    if (diff < 60) return `${diff} phút`;
    if (diff < 1440) return `${Math.floor(diff/60)} giờ ${diff%60} phút`;
    return `${Math.floor(diff/1440)} ngày ${Math.floor((diff%1440)/60)} giờ`;
}

function loadData() {
    // Load active outages
    fetch('/api/quangchudong/active')
        .then(res => res.json())
        .then(data => {
            document.getElementById('active-count').textContent = data.length;
            
            const tbody = document.querySelector('#active-table tbody');
            tbody.innerHTML = data.map(row => `
                <tr>
                    <td>${row.ma_tb || 'N/A'}</td>
                    <td>${row.ten_tb || 'N/A'}</td>
                    <td>${row.dienthoai_lh || ''}</td>
                    <td>${row.ten_nvkt_db || ''}</td>
                    <td>${row.doi_vt || ''}</td>
                    <td>${row.olt_name || ''}</td>
                    <td>${row.alert_time || ''}</td>
                    <td class="text-danger fw-bold">${formatDuration(row.alert_time)}</td>
                </tr>
            `).join('');
        });
    
    // Load recovered
    fetch('/api/quangchudong/recovered')
        .then(res => res.json())
        .then(data => {
            document.getElementById('recovered-count').textContent = data.length;
            
            const tbody = document.querySelector('#recovered-table tbody');
            tbody.innerHTML = data.map(row => `
                <tr>
                    <td>${row.ma_tb || 'N/A'}</td>
                    <td>${row.ten_tb || 'N/A'}</td>
                    <td>${row.ten_nvkt_db || ''}</td>
                    <td>${row.doi_vt || ''}</td>
                    <td>${row.outage_time || ''}</td>
                    <td>${row.recovery_time || ''}</td>
                    <td>${row.outage_duration_minutes || 0} phút</td>
                </tr>
            `).join('');
        });
    
    document.getElementById('last-update').textContent = 
        new Date().toLocaleTimeString('vi-VN');
}

// Load ban đầu
loadData();

// Auto-refresh
setInterval(loadData, REFRESH_INTERVAL);
</script>
{% endblock %}
```

---

## 4. Cập nhật Navigation Menu

Thêm link vào menu navigation (thường trong `base.html` hoặc sidebar):

```html
<li class="nav-item">
    <a class="nav-link" href="/quangchudong">
        <i class="fas fa-satellite-dish"></i>
        Quang Chủ động
    </a>
</li>
```

---

## 5. Checklist triển khai

- [ ] Copy file `subscriber_history.db` sang server dashboard (hoặc cấu hình đường dẫn)
- [ ] Thêm 4 routes vào `dashboard.py`:
  - [ ] `/quangchudong` - Trang chính
  - [ ] `/api/quangchudong/active` - API active outages
  - [ ] `/api/quangchudong/recovered` - API recovered
  - [ ] `/api/quangchudong/stats` - API thống kê
- [ ] Tạo file `templates/quangchudong.html`
- [ ] Thêm link vào navigation menu
- [ ] Restart dashboard server

---

## 6. Đồng bộ dữ liệu

### Option A: Shared Database
Nếu dashboard và batch_measure chạy cùng server, cấu hình chung đường dẫn DB.

### Option B: Sync định kỳ
Tạo cron job copy `subscriber_history.db` mỗi phút:
```bash
* * * * * cp /home/vtst/do-kiem-chu-dong-v0/subscriber_history.db /path/to/dashboard/data/
```

### Option C: Remote Database
Sử dụng PostgreSQL/MySQL thay vì SQLite nếu cần truy cập từ xa.

---

## 7. Mở rộng (tùy chọn)

1. **Filter theo Đội VT**: Thêm dropdown lọc theo đội
2. **Export Excel**: Nút tải danh sách ra Excel
3. **Biểu đồ xu hướng**: Chart số lượng DOWN theo thời gian
4. **WebSocket**: Real-time update không cần refresh
