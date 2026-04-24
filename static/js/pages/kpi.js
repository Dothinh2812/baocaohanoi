document.addEventListener('DOMContentLoaded', async function () {
    bindDateFilter();
    await initKPIPage();
});

window.kpiData = {
    tomtat: null,
    chitiet: null,
    fileInfo: null,
    selectedDate: '',
    latestAvailableDate: '',
    dateHasData: false
};

function getRequestedDate() {
    const params = new URLSearchParams(window.location.search);
    return params.get('date') || '';
}

function bindDateFilter() {
    const input = document.getElementById('kpi-date-input');
    const applyButton = document.getElementById('kpi-date-apply');
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
    const input = document.getElementById('kpi-date-input');
    if (input) {
        input.value = data.selected_date || getRequestedDate();
    }

    const meta = document.getElementById('kpi-date-meta');
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

async function initKPIPage() {
    await loadKPIData();
    initKPIDonviTabs();
}

async function loadKPIData() {
    try {
        const endpoint = getRequestedDate()
            ? `/api/kpi-data?date=${encodeURIComponent(getRequestedDate())}`
            : '/api/kpi-data';
        const response = await fetch(endpoint);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        if (data.error) {
            throw new Error(data.error);
        }

        syncDateFilterState(data);

        window.kpiData.fileInfo = data.file_info || null;
        window.kpiData.selectedDate = data.selected_date || '';
        window.kpiData.latestAvailableDate = data.latest_available_date || '';
        window.kpiData.dateHasData = Boolean(data.date_has_data);
        window.kpiData.tomtat = data.sheets && data.sheets['tomtat'] ? data.sheets['tomtat'] : null;
        window.kpiData.chitiet = data.sheets && data.sheets['chitiet'] ? data.sheets['chitiet'] : null;

        if (window.kpiData.fileInfo) {
            const tomtatTs = document.getElementById('kpi-tomtat-timestamp');
            const chitietTs = document.getElementById('kpi-chitiet-timestamp');

            if (tomtatTs && window.kpiData.fileInfo.tomtat_modified) {
                tomtatTs.textContent = `(Cập nhật: ${window.kpiData.fileInfo.tomtat_modified})`;
            }
            if (chitietTs && window.kpiData.fileInfo.chitiet_modified) {
                chitietTs.textContent = `(Cập nhật: ${window.kpiData.fileInfo.chitiet_modified})`;
            }
        }

        filterKPIByDonvi('all');
    } catch (error) {
        console.error('Error loading KPI data:', error);
        showKPIError('kpi-tomtat-table', error.message);
        showKPIError('kpi-chitiet-table', error.message);
    }
}

function initKPIDonviTabs() {
    const tabs = document.querySelectorAll('#kpi-donvi-tabs .excel-tab');

    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            tabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const donvi = this.dataset.donvi;
            filterKPIByDonvi(donvi);
        });
    });
}

function filterKPIByDonvi(donvi) {
    if (window.kpiData.tomtat) {
        const filteredTomtat = filterSheetByDonvi(window.kpiData.tomtat, donvi);
        const tomtatFileInfo = window.kpiData.fileInfo && window.kpiData.fileInfo.tomtat_modified
            ? { modified: window.kpiData.fileInfo.tomtat_modified }
            : null;
        renderKPITable('kpi-tomtat-table', filteredTomtat, tomtatFileInfo, window.kpiData.selectedDate);
    }

    if (window.kpiData.chitiet) {
        const filteredChitiet = filterSheetByDonvi(window.kpiData.chitiet, donvi);
        const chitietFileInfo = window.kpiData.fileInfo && window.kpiData.fileInfo.chitiet_modified
            ? { modified: window.kpiData.fileInfo.chitiet_modified }
            : null;
        renderKPITable('kpi-chitiet-table', filteredChitiet, chitietFileInfo, window.kpiData.selectedDate);
    }
}

function filterSheetByDonvi(sheetData, donvi) {
    if (donvi === 'all') {
        return sheetData;
    }

    const filteredData = sheetData.data.filter(row => row['don_vi'] === donvi);
    return {
        columns: sheetData.columns,
        data: filteredData
    };
}

function renderKPITable(containerId, sheetData, fileInfo = null, selectedDate = '') {
    const container = document.getElementById(containerId);
    if (!container) {
        console.error(`Container not found: ${containerId}`);
        return;
    }

    container.innerHTML = '';

    if (!sheetData || !sheetData.columns || !sheetData.data || sheetData.data.length === 0) {
        container.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu${selectedDate ? ` ngày ${selectedDate}` : ''}</p></div>`;
        return;
    }

    const tableHtml = createExcelTable(sheetData, '', {
        showRowNumbers: true,
        tableClass: 'excel-table',
        maxHeight: '600px',
        fileInfo: fileInfo
    });

    container.innerHTML = tableHtml;
}

function showKPIError(containerId, message) {
    const container = document.getElementById(containerId);
    if (!container) return;

    container.innerHTML = `
        <div class="loading">
            <i class="fas fa-exclamation-triangle"></i>
            <br>Lỗi khi tải dữ liệu: ${message}
        </div>
    `;
}
