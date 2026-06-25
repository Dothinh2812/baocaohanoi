/* ========================================
   SHC CTS PAGE JAVASCRIPT
   - Load SHC CTS comparison data from Excel
   - Render summary and detail tabs by Don vi
   ======================================== */

let shcCTSProgressChart = null;

document.addEventListener('DOMContentLoaded', async function () {
    await initSHCCTSDetailDownloader();
    await loadSHCCTSData();
    initShcCtsKiemSoat();
});

async function loadSHCCTSData() {
    try {
        const data = await API.fetchData('/api/shc-cts-data');
        const fileInfo = data.file_info || null;

        if (data.tong_hop) {
            document.getElementById('shc-cts-summary-container').innerHTML = createExcelTable(
                data.tong_hop,
                'Tổng hợp SHC CTS theo tổ',
                {
                    showRowNumbers: true,
                    maxHeight: '500px',
                    fileInfo: fileInfo
                }
            );
        }

        if (data.don_vi && Object.keys(data.don_vi).length > 0) {
            const donViNames = Object.keys(data.don_vi);
            window.shcCTSDetailData = data.don_vi;
            window.shcCTSFileInfo = fileInfo;
            renderSHCCTSDetailTabs(donViNames);
            renderSHCCTSDetailTable(donViNames[0]);
        } else {
            showEmptyState('Không có dữ liệu chi tiết theo tổ', 'shc-cts-detail-container');
        }

        if (data.tien_do_theo_don_vi && Object.keys(data.tien_do_theo_don_vi).length > 0) {
            const progressDonViNames = Object.keys(data.tien_do_theo_don_vi);
            window.shcCTSProgressData = data.tien_do_theo_don_vi;
            window.shcCTSProgressFileInfo = data.tien_do_file_info || fileInfo;
            renderSHCCTSProgressTabs(progressDonViNames);
            renderSHCCTSProgressTable(progressDonViNames[0]);
        } else {
            showEmptyState('Không có dữ liệu tiến độ theo đơn vị', 'shc-cts-progress-container');
        }
    } catch (error) {
        console.error('Error loading SHC CTS data:', error);
        showError('Không thể tải dữ liệu SHC CTS: ' + error.message, 'shc-cts-summary-container');
        showError('Không thể tải dữ liệu SHC CTS: ' + error.message, 'shc-cts-detail-container');
        showError('Không thể tải dữ liệu SHC CTS: ' + error.message, 'shc-cts-progress-container');
    }
}

function renderSHCCTSDetailTabs(donViNames) {
    const tabsContainer = document.getElementById('shc-cts-don-vi-tabs');
    if (!tabsContainer) return;

    tabsContainer.innerHTML = donViNames
        .map((name, index) => {
            const activeClass = index === 0 ? 'active' : '';
            const shortName = name.replace('Tổ Kỹ thuật Địa bàn ', '');
            return `<li class="excel-tab ${activeClass}" data-don-vi="${escapeSHCCTSHtml(name)}">${escapeSHCCTSHtml(shortName)}</li>`;
        })
        .join('');

    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(item => item.classList.remove('active'));
            this.classList.add('active');
            renderSHCCTSDetailTable(this.dataset.donVi);
        });
    });
}

function renderSHCCTSDetailTable(donViName) {
    const sheetData = window.shcCTSDetailData ? window.shcCTSDetailData[donViName] : null;
    document.getElementById('shc-cts-detail-container').innerHTML = createExcelTable(
        sheetData,
        `Tổng hợp SHC CTS - ${donViName}`,
        {
            showRowNumbers: true,
            maxHeight: '600px',
            fileInfo: window.shcCTSFileInfo || null
        }
    );
}

function renderSHCCTSProgressTabs(donViNames) {
    const tabsContainer = document.getElementById('shc-cts-progress-tabs');
    if (!tabsContainer) return;

    tabsContainer.innerHTML = donViNames
        .map((name, index) => {
            const activeClass = index === 0 ? 'active' : '';
            const shortName = name.replace('Tổ Kỹ thuật Địa bàn ', '');
            return `<li class="excel-tab ${activeClass}" data-don-vi="${escapeSHCCTSHtml(name)}">${escapeSHCCTSHtml(shortName)}</li>`;
        })
        .join('');

    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(item => item.classList.remove('active'));
            this.classList.add('active');
            renderSHCCTSProgressTable(this.dataset.donVi);
        });
    });
}

function renderSHCCTSProgressTable(donViName) {
    const sheetData = window.shcCTSProgressData ? window.shcCTSProgressData[donViName] : null;
    const container = document.getElementById('shc-cts-progress-container');
    if (!container) return;

    const chartId = `shc-cts-progress-chart-${Math.random().toString(36).slice(2, 9)}`;
    const sourceTimestamp = getSHCCTSProgressSourceTimestamp(sheetData);
    const tableHtml = createExcelTable(
        sheetData,
        `Tiến độ xử lý shc trong ngày - ${donViName}`,
        {
            showRowNumbers: true,
            maxHeight: '600px',
            fileInfo: window.shcCTSProgressFileInfo || null
        }
    );

    container.innerHTML = `
        <div class="shc-cts-progress-chart-card">
            <div class="shc-cts-progress-chart-head">
                <div>
                    <div class="shc-cts-progress-chart-title">Biểu đồ tiến độ xử lý - ${escapeSHCCTSHtml(donViName)}</div>
                    <div class="shc-cts-progress-chart-subtitle">
                        Mỗi NVKT một cột, stacked theo trạng thái đạt/chưa đạt
                        ${sourceTimestamp ? `<br><span class="shc-cts-progress-measured-at">Thời điểm đo: ${escapeSHCCTSHtml(sourceTimestamp)}</span>` : ''}
                    </div>
                </div>
                <div class="shc-cts-progress-chart-legend">
                    <span><span class="shc-cts-progress-chart-dot" style="background:#198754;"></span>Đã đạt</span>
                    <span><span class="shc-cts-progress-chart-dot" style="background:#dc3545;"></span>Chưa đạt</span>
                    <span><span class="shc-cts-progress-chart-dot" style="background:#fd7e14;"></span>Lỗi đo/OFF</span>
                </div>
            </div>
            <div class="shc-cts-progress-chart-canvas">
                <canvas id="${chartId}"></canvas>
            </div>
        </div>
        ${tableHtml}
    `;

    renderSHCCTSProgressChart(chartId, sheetData);
}

function getSHCCTSProgressSourceTimestamp(sheetData) {
    if (sheetData && Array.isArray(sheetData.data) && sheetData.data.length > 0) {
        const timestamp = sheetData.data[0].Timestamp;
        if (timestamp) return String(timestamp);
    }

    const fileInfo = window.shcCTSProgressFileInfo || null;
    return fileInfo && fileInfo.modified ? fileInfo.modified : '';
}

function downloadExcelSHCCTS() {
    window.location.href = '/download/excel-shc-cts';
}

function renderSHCCTSProgressChart(canvasId, sheetData) {
    if (shcCTSProgressChart) {
        shcCTSProgressChart.destroy();
        shcCTSProgressChart = null;
    }

    if (typeof Chart === 'undefined' || !sheetData || !Array.isArray(sheetData.data) || sheetData.data.length === 0) {
        return;
    }

    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    const rows = sheetData.data.filter(row => {
        const name = String(row['NVKT_DB'] || row['NVKT'] || '').trim();
        return name && name.toUpperCase() !== 'TỔNG';
    });
    if (rows.length === 0) {
        return;
    }

    const labels = rows.map(row => String(row['NVKT_DB'] || row['NVKT'] || '').trim());
    const achieved = rows.map(row => parseSHCCTSNumber(row['Tổng đã đạt']));
    const notAchieved = rows.map(row => parseSHCCTSNumber(row['Chưa đạt']));
    const measurementErrors = rows.map(row => parseSHCCTSNumber(getSHCCTSProgressValue(row, ['OFF/Lỗi', 'ONU OFF/Lỗi đo', 'Lỗi đo/OFF'])));

    shcCTSProgressChart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Đã đạt',
                    data: achieved,
                    backgroundColor: '#198754',
                    borderColor: '#146c43',
                    borderWidth: 1,
                    borderRadius: 5,
                    barPercentage: 0.72,
                    categoryPercentage: 0.72,
                    maxBarThickness: 42,
                    stack: 'progress',
                },
                {
                    label: 'Chưa đạt',
                    data: notAchieved,
                    backgroundColor: '#dc3545',
                    borderColor: '#b02a37',
                    borderWidth: 1,
                    borderRadius: 5,
                    barPercentage: 0.72,
                    categoryPercentage: 0.72,
                    maxBarThickness: 42,
                    stack: 'progress',
                },
                {
                    label: 'Lỗi đo/OFF',
                    data: measurementErrors,
                    backgroundColor: '#fd7e14',
                    borderColor: '#c85f0d',
                    borderWidth: 1,
                    borderRadius: 5,
                    barPercentage: 0.72,
                    categoryPercentage: 0.72,
                    maxBarThickness: 42,
                    stack: 'progress',
                },
            ],
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                intersect: false,
                mode: 'index',
            },
            plugins: {
                legend: {
                    display: false,
                },
                tooltip: {
                    callbacks: {
                        footer: function (items) {
                            const total = items.reduce((sum, item) => sum + Number(item.parsed.y || 0), 0);
                            return `Tổng: ${total}`;
                        },
                    },
                },
            },
            scales: {
                x: {
                    stacked: true,
                    ticks: {
                        color: '#415848',
                        maxRotation: 45,
                        minRotation: 0,
                    },
                    grid: {
                        display: false,
                    },
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        precision: 0,
                        color: '#415848',
                    },
                    grid: {
                        color: 'rgba(65, 88, 72, 0.12)',
                    },
                },
            },
        },
    });
}

async function initSHCCTSDetailDownloader() {
    const teamSelect = document.getElementById('shc-cts-detail-team');
    const fileSelect = document.getElementById('shc-cts-detail-file');
    const downloadBtn = document.getElementById('download-shc-cts-detail-btn');
    const previewBtn = document.getElementById('preview-shc-cts-detail-btn');

    if (!teamSelect || !fileSelect || !downloadBtn || !previewBtn) {
        return;
    }

    try {
        setSHCCTSDetailStatus('Đang tải danh sách tổ...');
        const options = await API.fetchData('/api/shc-cts-nvkt-detail/options');
        const teams = options.teams || [];

        if (teams.length === 0) {
            teamSelect.innerHTML = '<option value="">Không có tổ</option>';
            fileSelect.innerHTML = '<option value="">Không có cá nhân</option>';
            syncSHCCTSDetailActionButtons();
            setSHCCTSDetailStatus('Không tìm thấy tổ trong thư mục chi tiết SHC NVKT K1.', true);
            return;
        }

        teamSelect.innerHTML = teams
            .map(item => `<option value="${escapeSHCCTSHtml(item.key)}">${escapeSHCCTSHtml(item.label)}</option>`)
            .join('');

        teamSelect.addEventListener('change', async () => {
            clearSHCCTSDetailPreview();
            await loadSHCCTSDetailFiles();
        });
        fileSelect.addEventListener('change', () => {
            syncSHCCTSDetailActionButtons();
            clearSHCCTSDetailPreview();
        });
        downloadBtn.addEventListener('click', downloadSelectedSHCCTSDetailFile);
        previewBtn.addEventListener('click', previewSelectedSHCCTSDetailFile);

        await loadSHCCTSDetailFiles();
    } catch (error) {
        teamSelect.innerHTML = '<option value="">Không thể tải tổ</option>';
        fileSelect.innerHTML = '<option value="">Không thể tải cá nhân</option>';
        syncSHCCTSDetailActionButtons();
        setSHCCTSDetailStatus(`Không thể tải danh mục file chi tiết: ${error.message}`, true);
    }
}

async function loadSHCCTSDetailFiles() {
    const teamSelect = document.getElementById('shc-cts-detail-team');
    const fileSelect = document.getElementById('shc-cts-detail-file');
    const team = teamSelect ? teamSelect.value : '';

    if (!team) {
        fileSelect.innerHTML = '<option value="">Chưa chọn tổ</option>';
        syncSHCCTSDetailActionButtons();
        return;
    }

    try {
        setSHCCTSDetailStatus('Đang tải danh sách cá nhân...');
        const data = await API.fetchData(`/api/shc-cts-nvkt-detail/files?team=${encodeURIComponent(team)}`);
        const files = data.files || [];

        if (files.length === 0) {
            fileSelect.innerHTML = '<option value="">Không có cá nhân phù hợp</option>';
            syncSHCCTSDetailActionButtons();
            setSHCCTSDetailStatus('Không tìm thấy file Excel cá nhân trong tổ đã chọn.');
            return;
        }

        fileSelect.innerHTML = files
            .map(file => `<option value="${escapeSHCCTSHtml(file.name)}">${escapeSHCCTSHtml(file.display_name)}</option>`)
            .join('');

        syncSHCCTSDetailActionButtons();
        const sourceLabel = data.source_name ? ` từ ${data.source_name}` : '';
        setSHCCTSDetailStatus(`Tìm thấy ${files.length} cá nhân chi tiết${sourceLabel}.`);
    } catch (error) {
        fileSelect.innerHTML = '<option value="">Không thể tải danh sách cá nhân</option>';
        syncSHCCTSDetailActionButtons();
        setSHCCTSDetailStatus(`Không thể tải danh sách cá nhân: ${error.message}`, true);
    }
}

function downloadSelectedSHCCTSDetailFile() {
    const team = document.getElementById('shc-cts-detail-team').value;
    const fileName = document.getElementById('shc-cts-detail-file').value;

    if (!team || !fileName) {
        setSHCCTSDetailStatus('Vui lòng chọn đủ tổ và cá nhân.', true);
        return;
    }

    setSHCCTSDetailStatus(`Đang tải file cá nhân: ${fileName}`);
    window.location.href = `/download/shc-cts-nvkt-detail/${encodeURIComponent(team)}/${encodeURIComponent(fileName)}`;
}

async function previewSelectedSHCCTSDetailFile() {
    const team = document.getElementById('shc-cts-detail-team').value;
    const fileName = document.getElementById('shc-cts-detail-file').value;
    const previewSection = document.getElementById('shc-cts-detail-preview-section');

    if (!team || !fileName) {
        setSHCCTSDetailStatus('Vui lòng chọn đủ tổ và cá nhân.', true);
        return;
    }

    previewSection.style.display = 'block';
    showLoading('shc-cts-detail-preview-container');
    document.getElementById('shc-cts-detail-preview-tabs').innerHTML = '';
    document.getElementById('shc-cts-detail-preview-meta').textContent = 'Đang tải nội dung file...';
    setSHCCTSDetailStatus(`Đang tải xem trực tiếp: ${fileName}`);

    try {
        const query = new URLSearchParams({ team: team, file_name: fileName });
        const data = await API.fetchData(`/api/shc-cts-nvkt-detail/preview?${query.toString()}`);
        const sheets = data.sheets || {};
        const sheetNames = Object.keys(sheets);

        if (sheetNames.length === 0) {
            showEmptyState('File Excel không có dữ liệu để hiển thị', 'shc-cts-detail-preview-container');
            document.getElementById('shc-cts-detail-preview-meta').textContent = fileName;
            setSHCCTSDetailStatus(`Không có dữ liệu trong file: ${fileName}`, true);
            return;
        }

        window.shcCTSDetailPreviewData = data;
        renderSHCCTSDetailPreviewTabs(sheetNames);
        renderSHCCTSDetailPreviewSheet(sheetNames[0]);

        const fileLabel = data.file_info && data.file_info.name ? data.file_info.name : fileName;
        const modifiedLabel = data.file_info && data.file_info.modified ? ` | Cập nhật: ${data.file_info.modified}` : '';
        const sourceLabel = data.source_name ? ` | Nguồn: ${data.source_name}` : '';
        document.getElementById('shc-cts-detail-preview-meta').textContent = `${fileLabel}${modifiedLabel}${sourceLabel}`;
        setSHCCTSDetailStatus(`Đang hiển thị nội dung file: ${fileLabel}`);
        document.getElementById('shc-cts-detail-preview-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) {
        showError(`Không thể xem trực tiếp file: ${error.message}`, 'shc-cts-detail-preview-container');
        document.getElementById('shc-cts-detail-preview-meta').textContent = fileName;
        setSHCCTSDetailStatus(`Không thể xem trực tiếp file: ${error.message}`, true);
    }
}

function renderSHCCTSDetailPreviewTabs(sheetNames) {
    const tabsContainer = document.getElementById('shc-cts-detail-preview-tabs');
    if (!tabsContainer) return;

    tabsContainer.innerHTML = sheetNames
        .map((name, index) => `<li class="excel-tab ${index === 0 ? 'active' : ''}" data-sheet="${escapeSHCCTSHtml(name)}">${escapeSHCCTSHtml(name)}</li>`)
        .join('');

    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(item => item.classList.remove('active'));
            this.classList.add('active');
            renderSHCCTSDetailPreviewSheet(this.dataset.sheet);
        });
    });
}

function renderSHCCTSDetailPreviewSheet(sheetName) {
    const previewData = window.shcCTSDetailPreviewData;
    if (!previewData || !previewData.sheets || !previewData.sheets[sheetName]) {
        showEmptyState('Không có dữ liệu sheet để hiển thị', 'shc-cts-detail-preview-container');
        return;
    }

    const sheetData = previewData.sheets[sheetName];
    if (sheetData.error) {
        showError(sheetData.error, 'shc-cts-detail-preview-container');
        return;
    }

    document.getElementById('shc-cts-detail-preview-container').innerHTML = createExcelTable(
        sheetData,
        `Chi tiết cá nhân - ${sheetName}`,
        {
            showRowNumbers: true,
            maxHeight: '600px',
            fileInfo: previewData.file_info || null
        }
    );
}

function syncSHCCTSDetailActionButtons() {
    const fileSelect = document.getElementById('shc-cts-detail-file');
    const hasValue = Boolean(fileSelect && fileSelect.value);
    const downloadBtn = document.getElementById('download-shc-cts-detail-btn');
    const previewBtn = document.getElementById('preview-shc-cts-detail-btn');

    if (downloadBtn) {
        downloadBtn.disabled = !hasValue;
    }
    if (previewBtn) {
        previewBtn.disabled = !hasValue;
    }
}

function clearSHCCTSDetailPreview() {
    const previewSection = document.getElementById('shc-cts-detail-preview-section');
    const previewTabs = document.getElementById('shc-cts-detail-preview-tabs');
    const previewContainer = document.getElementById('shc-cts-detail-preview-container');
    const previewMeta = document.getElementById('shc-cts-detail-preview-meta');

    if (previewSection) {
        previewSection.style.display = 'none';
    }
    if (previewTabs) {
        previewTabs.innerHTML = '';
    }
    if (previewContainer) {
        previewContainer.innerHTML = '';
    }
    if (previewMeta) {
        previewMeta.textContent = '';
    }

    window.shcCTSDetailPreviewData = null;
}

function setSHCCTSDetailStatus(message, isError = false) {
    const statusElement = document.getElementById('shc-cts-detail-status');
    if (!statusElement) return;

    statusElement.textContent = message;
    statusElement.style.color = isError ? '#d32f2f' : '#495057';
}

function escapeSHCCTSHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function getSHCCTSProgressValue(row, candidateColumns) {
    for (const column of candidateColumns) {
        if (Object.prototype.hasOwnProperty.call(row, column)) {
            return row[column];
        }
    }
    return 0;
}

function parseSHCCTSNumber(value) {
    if (value === null || value === undefined || value === '') {
        return 0;
    }
    if (typeof value === 'number') {
        return Number.isFinite(value) ? value : 0;
    }
    let normalized = String(value)
        .replace('%', '')
        .trim();
    if (normalized.includes(',')) {
        normalized = normalized.replace(/\./g, '').replace(',', '.');
    }
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
}

/* ========================================
   LỊCH SỬ TIẾN ĐỘ & KIỂM SOÁT TỔ TRƯỞNG SHC CTS
   - Datepicker chọn ngày (hôm nay = live Excel, cũ = DB)
   - Bảng chi tiết theo NVKT + cột nhập kiểm soát
   ======================================== */

const SHC_CTS_KS_DISPLAY_COLS = [
    'NVKT_DB', 'Tổng số', 'Đạt baseline', 'Đã xử lý trong ngày',
    'Tổng đã đạt', 'Chưa đạt', 'OFF/Lỗi', '% đạt',
];
const SHC_CTS_KS_DISPLAY_LABELS = {
    'NVKT_DB': 'NVKT', 'Tổng số': 'Tổng', 'Đạt baseline': 'Đạt BL',
    'Đã xử lý trong ngày': 'XL trong ngày', 'Tổng đã đạt': 'Đã đạt',
    'Chưa đạt': 'Chưa đạt', 'OFF/Lỗi': 'OFF/Lỗi', '% đạt': '%',
};
let _shcCtsKsCurrent = { date: '', donVi: '', sheets: {}, sourcePill: '' };

function _shcCtsKsEscape(v) {
    return String(v == null ? '' : v)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function _shcCtsKsThoiDiem(value) {
    if (!value) return '';
    const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/);
    return m ? `${m[4]}:${m[5]} ${m[3]}/${m[2]}` : value;
}

function _shcCtsKsBadge(row) {
    if (row.kiemsoat_da_nhap) {
        const when = _shcCtsKsThoiDiem(row.kiemsoat_thoi_diem);
        const who = row.kiemsoat_nguoi_nhap || '';
        return `<span class="shc-cts-ks-badge shc-cts-ks-da">Đã KS${when ? ' ' + when : ''}${who ? ' — ' + _shcCtsKsEscape(who) : ''}</span>`;
    }
    return `<span class="shc-cts-ks-badge shc-cts-ks-chua">Chưa</span>`;
}

function initShcCtsKiemSoat() {
    const dateInput = document.getElementById('shc-cts-ks-date');
    const donViSelect = document.getElementById('shc-cts-ks-don-vi');
    if (!dateInput) return;
    dateInput.addEventListener('change', () => { _shcCtsKsCurrent.date = dateInput.value; loadShcCtsKiemSoatDetail(); });
    donViSelect.addEventListener('change', () => { _shcCtsKsCurrent.donVi = donViSelect.value; loadShcCtsKiemSoatThongKe(); });
    loadShcCtsKiemSoatDetail();
}
window.initShcCtsKiemSoat = initShcCtsKiemSoat;

async function loadShcCtsKiemSoatDetail() {
    const dateInput = document.getElementById('shc-cts-ks-date');
    const donViSelect = document.getElementById('shc-cts-ks-don-vi');
    const chiTietEl = document.getElementById('shc-cts-ks-chitiet');
    const pillEl = document.getElementById('shc-cts-ks-source-pill');
    if (!chiTietEl) return;
    chiTietEl.innerHTML = '<div class="loading"><i class="fas fa-spinner"></i><br>Đang tải...</div>';

    const params = new URLSearchParams();
    if (_shcCtsKsCurrent.date) params.set('date', _shcCtsKsCurrent.date);
    try {
        const data = await API.getShcCtsKiemSoatDetail(params.toString());
        _shcCtsKsCurrent.sheets = data.sheets || {};

        if (!dateInput.value) dateInput.value = data.selected_date;
        _shcCtsKsCurrent.date = data.selected_date;

        const donViNames = Object.keys(_shcCtsKsCurrent.sheets);
        const prevDonVi = donViSelect.value;
        donViSelect.innerHTML = '<option value="">Tất cả</option>' +
            donViNames.map(n => `<option value="${_shcCtsKsEscape(n)}"${n === prevDonVi ? ' selected' : ''}>${_shcCtsKsEscape(n)}</option>`).join('');
        if (prevDonVi && donViNames.includes(prevDonVi)) {
            donViSelect.value = prevDonVi;
            _shcCtsKsCurrent.donVi = prevDonVi;
        }

        pillEl.innerHTML = data.is_today_live
            ? '<span class="shc-cts-ks-live-pill">Hôm nay (Excel trực tiếp)</span>'
            : '<span class="shc-cts-ks-live-pill db">Lịch sử (DB)</span>';

        renderShcCtsKiemSoatDetail();
        loadShcCtsKiemSoatThongKe();
    } catch (error) {
        chiTietEl.innerHTML = `<div class="error">Không tải được dữ liệu: ${_shcCtsKsEscape(error.message)}</div>`;
    }
}

function renderShcCtsKiemSoatDetail() {
    const container = document.getElementById('shc-cts-ks-chitiet');
    if (!container) return;
    const sheets = _shcCtsKsCurrent.sheets || {};
    const donViFilter = _shcCtsKsCurrent.donVi;
    const targetSheets = donViFilter ? { [donViFilter]: sheets[donViFilter] } : sheets;

    const parts = [];
    for (const [donVi, sheet] of Object.entries(targetSheets)) {
        const rows = (sheet && sheet.data) || [];
        const headers = SHC_CTS_KS_DISPLAY_COLS.map(c => `<th>${SHC_CTS_KS_DISPLAY_LABELS[c] || c}</th>`).join('');
        const body = rows.map(row => {
            const cells = SHC_CTS_KS_DISPLAY_COLS.map(c => `<td>${row[c] != null ? row[c] : ''}</td>`).join('');
            const nvkt = _shcCtsKsEscape(row.NVKT_DB);
            const noiDung = _shcCtsKsEscape(row.kiemsoat_noi_dung || '');
            return `<tr>${cells}<td class="shc-cts-ks-cell">
                <textarea class="shc-cts-ks-input" rows="2" data-nvkt="${nvkt}" data-don_vi="${_shcCtsKsEscape(row['Đơn vị'] || donVi)}">${noiDung}</textarea>
                <button class="shc-cts-ks-save-btn" onclick="saveShcCtsKiemSoatRow('${nvkt}')">Lưu</button>
                <span class="shc-cts-ks-badge" id="shc-cts-ks-status-${nvkt}">${_shcCtsKsBadge(row)}</span>
            </td></tr>`;
        }).join('');
        parts.push(`<h5>${_shcCtsKsEscape(donVi)}</h5>
            <div class="excel-table-card"><div class="excel-table-body" style="max-height:520px;overflow:auto;">
            <table class="excel-table"><thead><tr>${headers}<th>Kiểm soát</th></tr></thead>
            <tbody>${body || '<tr><td colspan="99">Không có NVKT.</td></tr>'}</tbody></table>
            </div></div>`);
    }
    container.innerHTML = parts.join('') || '<div>Không có dữ liệu.</div>';
}

async function saveShcCtsKiemSoatRow(nvktDb) {
    const textarea = document.querySelector(`.shc-cts-ks-input[data-nvkt="${CSS.escape(nvktDb)}"]`);
    if (!textarea) return;
    const statusEl = document.getElementById(`shc-cts-ks-status-${nvktDb}`);
    try {
        const result = await API.saveShcCtsKiemSoat({
            ngay_xu_ly: _shcCtsKsCurrent.date,
            nvkt_db: nvktDb,
            don_vi: textarea.dataset.don_vi,
            noi_dung: textarea.value,
        });
        if (!result || result.ok === false) throw new Error((result && result.error) || 'Lỗi');
        if (statusEl) {
            statusEl.innerHTML = _shcCtsKsBadge({
                kiemsoat_da_nhap: !!result.noi_dung,
                kiemsoat_thoi_diem: new Date().toISOString().replace('T', ' ').substring(0, 19),
                kiemsoat_nguoi_nhap: result.nguoi_nhap,
            });
        }
        loadShcCtsKiemSoatThongKe();
    } catch (error) {
        alert('Không lưu được: ' + error.message);
    }
}
window.saveShcCtsKiemSoatRow = saveShcCtsKiemSoatRow;

async function loadShcCtsKiemSoatThongKe() {
    const statsEl = document.getElementById('shc-cts-ks-stats');
    const lichSuEl = document.getElementById('shc-cts-ks-lich-su');
    if (!statsEl) return;
    const params = new URLSearchParams();
    if (_shcCtsKsCurrent.date) params.set('date', _shcCtsKsCurrent.date);
    if (_shcCtsKsCurrent.donVi) params.set('don_vi', _shcCtsKsCurrent.donVi);
    try {
        const data = await API.getShcCtsKiemSoatThongKe(params.toString());
        const s = data.summary || {};
        const card = (v, l) => `<div class="shc-cts-ks-card"><div class="v">${v}</div><div class="l">${l}</div></div>`;
        statsEl.innerHTML =
            card(s.tong_so ?? 0, 'Tổng số') +
            card(s.tong_dat ?? 0, 'Tổng đã đạt') +
            card(s.chua_dat ?? 0, 'Chưa đạt') +
            card((s.ty_le_dat ?? 0) + '%', 'Tỷ lệ đạt') +
            card(s.da_ks ?? 0, 'Đã kiểm soát') +
            card(s.chua_ks ?? 0, 'Chưa KS');

        const lichSu = data.lich_su || [];
        lichSuEl.innerHTML = lichSu.length
            ? `<div class="excel-table-card"><div class="excel-table-body" style="max-height:320px;overflow:auto;">
               <table class="excel-table"><thead><tr><th>Ngày</th><th>Tổng số</th><th>Tổng đã đạt</th><th>Đã KS</th></tr></thead>
               <tbody>${lichSu.map(r => `<tr><td>${r.ngay_xu_ly}</td><td>${r.tong_so}</td><td>${r.tong_dat}</td><td>${r.da_ks}</td></tr>`).join('')}</tbody>
               </table></div></div>`
            : '<div>Chưa có lịch sử.</div>';
    } catch (error) {
        statsEl.innerHTML = `<div class="error">${_shcCtsKsEscape(error.message)}</div>`;
    }
}

function exportShcCtsKiemSoat() {
    const params = new URLSearchParams();
    if (_shcCtsKsCurrent.date) params.set('date', _shcCtsKsCurrent.date);
    if (_shcCtsKsCurrent.donVi) params.set('don_vi', _shcCtsKsCurrent.donVi);
    window.location.href = '/download/shc-cts-kiemsoat-report' + (params.toString() ? '?' + params.toString() : '');
}
window.exportShcCtsKiemSoat = exportShcCtsKiemSoat;
