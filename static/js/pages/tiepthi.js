document.addEventListener('DOMContentLoaded', async function () {
    bindDateFilter();
    await loadTiepThiData();
});

function getRequestedDate() {
    const params = new URLSearchParams(window.location.search);
    return params.get('date') || '';
}

function bindDateFilter() {
    const input = document.getElementById('tiepthi-date-input');
    const applyButton = document.getElementById('tiepthi-date-apply');
    if (!input || !applyButton) {
        return;
    }

    input.value = getRequestedDate();

    function applyDateFilter() {
        const params = new URLSearchParams(window.location.search);
        const nextDate = input.value.trim();
        if (nextDate) {
            params.set('date', nextDate);
        } else {
            params.delete('date');
        }
        const nextQuery = params.toString();
        window.location.href = nextQuery
            ? `${window.location.pathname}?${nextQuery}`
            : window.location.pathname;
    }

    applyButton.addEventListener('click', applyDateFilter);
    input.addEventListener('keydown', function(event) {
        if (event.key === 'Enter') {
            applyDateFilter();
        }
    });
}

function syncDateFilterState(data) {
    const input = document.getElementById('tiepthi-date-input');
    if (input) {
        input.value = data.selected_date || getRequestedDate();
    }

    const meta = document.getElementById('tiepthi-date-meta');
    if (!meta) {
        return;
    }

    if (!data.selected_date) {
        meta.textContent = 'Chưa xác định được ngày dữ liệu.';
        return;
    }

    if (data.date_has_data) {
        meta.textContent = `Đang xem dữ liệu ngày ${data.selected_date}. Ngày mới nhất hiện có: ${data.latest_available_date || data.selected_date}.`;
        return;
    }

    meta.textContent = `Ngày ${data.selected_date} hiện chưa có dữ liệu. Ngày mới nhất hiện có: ${data.latest_available_date || 'không xác định'}.`;
}

async function loadTiepThiData() {
    try {
        const endpoint = getRequestedDate()
            ? `/api/tiepthi-data?date=${encodeURIComponent(getRequestedDate())}`
            : '/api/tiepthi-data';
        const data = await API.fetchData(endpoint);
        syncDateFilterState(data);

        renderTiepThiSummary(data.tong_hop, data.selected_date);

        if (data && data.sheets) {
            const fileInfo = data.file_info || null;
            const sheetNames = Object.keys(data.sheets);

            renderTiepThiTabs('tiepthi-tabs', sheetNames, (sheetName) => {
                renderTiepThiTable('tiepthi-tables-container', data.sheets[sheetName], sheetName, fileInfo, data.selected_date);
            });

            if (sheetNames.length > 0) {
                renderTiepThiTable('tiepthi-tables-container', data.sheets[sheetNames[0]], sheetNames[0], fileInfo, data.selected_date);
            }
        }
    } catch (error) {
        const container = document.getElementById('tiepthi-tables-container');
        if (container) {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Không thể tải dữ liệu Tiếp thị</p></div>';
        }
    }
}

function renderTiepThiSummary(sheetData, selectedDate) {
    const table = document.getElementById('tiepthi-summary-table');
    if (!table) return;

    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    if (!thead || !tbody) return;

    if (!sheetData || !Array.isArray(sheetData.columns) || !Array.isArray(sheetData.data) || sheetData.data.length === 0) {
        thead.innerHTML = '';
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:16px;">Không có dữ liệu${selectedDate ? ` ngày ${selectedDate}` : ''}</td></tr>`;
        return;
    }

    thead.innerHTML = `<tr>${sheetData.columns.map(col => `<th>${escapeHtml(col)}</th>`).join('')}</tr>`;
    tbody.innerHTML = sheetData.data.map(row => {
        const isTotal = row['Đơn vị'] === 'TỔNG CỘNG';
        const rowHtml = sheetData.columns.map(col => {
            const value = row[col] ?? '';
            if (col === 'STT') {
                return `<td style="text-align: center; width: 60px;">${escapeHtml(value)}</td>`;
            }
            if (col === 'Đơn vị') {
                return `<td style="width: 160px;"><strong>${escapeHtml(value)}</strong></td>`;
            }
            if (col === 'Tổng') {
                return `<td style="text-align: center; font-weight: 600;${!isTotal ? ' color: #1a5089;' : ''}">${escapeHtml(value)}</td>`;
            }
            return `<td style="text-align: center;">${escapeHtml(value)}</td>`;
        }).join('');
        return `<tr${isTotal ? ' style="background-color: #6b9bc3; color: #2c3e50; font-weight: 700;"' : ''}>${rowHtml}</tr>`;
    }).join('');
}

function renderTiepThiTabs(tabsId, sheetNames, onTabClick) {
    const tabsContainer = document.getElementById(tabsId);
    if (!tabsContainer) return;

    let html = '';
    sheetNames.forEach((name, index) => {
        const activeClass = index === 0 ? 'active' : '';
        html += `<li class="excel-tab ${activeClass}" data-sheet="${name}">${name}</li>`;
    });

    tabsContainer.innerHTML = html;

    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(t => t.classList.remove('active'));
            this.classList.add('active');
            onTabClick(this.dataset.sheet);
        });
    });
}

function renderTiepThiTable(containerId, sheetData, sheetName, fileInfo = null, selectedDate = '') {
    const container = document.getElementById(containerId);
    if (!container) return;

    if (!sheetData || !Array.isArray(sheetData.columns) || !Array.isArray(sheetData.data) || sheetData.data.length === 0) {
        container.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu${selectedDate ? ` ngày ${selectedDate}` : ''}</p></div>`;
        return;
    }

    const filtered = {
        columns: sheetData.columns.filter(col => col !== 'STT'),
        data: sheetData.data
    };

    const html = createExcelTable(filtered, sheetName, {
        showRowNumbers: true,
        maxHeight: 'none',
        fileInfo: fileInfo
    });

    container.innerHTML = html;
}

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
