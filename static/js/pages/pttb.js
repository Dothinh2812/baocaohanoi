/* ========================================
   PTTB PAGE JAVASCRIPT
   - Load PTTB data from 3 endpoints
   - Handle chart tabs and data tabs
   ======================================== */

document.addEventListener('DOMContentLoaded', async function () {
    await initPTTBPage();
});

async function initPTTBPage() {
    // Initialize chart tabs
    initPTTBChartTabs();

    // Load all data in parallel
    await Promise.all([
        loadPTTBSummaryData(),
        loadPTTBDetailData(),
        loadPTTBPendingData(),
        loadPTTBChitietToData()
    ]);

    // Initialize detail tabs after data is loaded
    initPTTBDiabanTabs();
    initPTTBDetailTabs();
    initPTTBPendingTabs();
    initPTTBChitietToTabs();
}

/**
 * Initialize chart tabs for PTTB 4 tổ
 */
function initPTTBChartTabs() {
    const tabs = document.querySelectorAll('#pttb-3to-tabs .image-tab');
    const chartDisplay = document.getElementById('pttb-chart-display');

    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            tabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const chartName = this.dataset.chart;
            const chartPath = `/pttb_3_to/chart_pttb_3to_tokt_${chartName}.png`;
            chartDisplay.src = chartPath;
            chartDisplay.alt = `Biểu đồ PTTB ${this.textContent}`;
        });
    });
}

/**
 * Load PTTB Summary data (3 tables)
 */
async function loadPTTBSummaryData() {
    try {
        const response = await fetch('/api/pttb-data-summary');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        if (data.error) {
            throw new Error(data.error);
        }

        const fileInfo = data.file_info || null;

        // Render TK_TongHop_DonGian
        if (data.sheets && data.sheets['TK_TongHop_DonGian']) {
            renderPTTBTable('pttb-summary-simple-table', data.sheets['TK_TongHop_DonGian'], fileInfo);
        }

        // Render tong_hop_5doi
        if (data.sheets && data.sheets['tong_hop_5doi']) {
            // Rename column "Số phiếu sắp quá giờ" -> "Số phiếu sắp quá giờ (t/g còn lại <8h)"
            const sheetData = data.sheets['tong_hop_5doi'];
            const renamedColumns = sheetData.columns.map(col =>
                col === 'Số phiếu sắp quá giờ' ? 'Số phiếu sắp quá giờ (t/g còn lại <8h)' : col
            );
            const renamedSheetData = {
                columns: renamedColumns,
                data: sheetData.data.map(row => {
                    const newRow = { ...row };
                    if ('Số phiếu sắp quá giờ' in newRow) {
                        newRow['Số phiếu sắp quá giờ (t/g còn lại <8h)'] = newRow['Số phiếu sắp quá giờ'];
                        delete newRow['Số phiếu sắp quá giờ'];
                    }
                    return newRow;
                })
            };
            renderPTTBTable('pttb-summary-5doi-table', renamedSheetData, fileInfo);
        }

        // Render TK_TongHop_TrangThai
        if (data.sheets && data.sheets['TK_TongHop_TrangThai']) {
            renderPTTBTable('pttb-summary-status-table', data.sheets['TK_TongHop_TrangThai'], fileInfo);
        }
    } catch (error) {
        console.error('Error loading PTTB summary:', error);
        showPTTBError('pttb-summary-simple-table', error.message);
        showPTTBError('pttb-summary-5doi-table', error.message);
        showPTTBError('pttb-summary-status-table', error.message);
    }
}

/**
 * Load PTTB Detail data (tong_hop_dia_ban + 4 tổ)
 */
async function loadPTTBDetailData() {
    try {
        const response = await fetch('/api/pttb-data-detail');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        console.log('PTTB Detail Data:', data);

        if (data.error) {
            throw new Error(data.error);
        }

        // Store data globally for tab switching
        window.pttbDetailData = data.sheets;
        window.pttbDetailFileInfo = data.file_info || null;
        console.log('Available sheets:', Object.keys(data.sheets || {}));

        // Store tong_hop_dia_ban data for tab filtering
        if (data.sheets && data.sheets['tong_hop_dia_ban']) {
            window.pttbDiabanData = data.sheets['tong_hop_dia_ban'];
            // Render Sơn Tây data by default (first tab)
            filterDiabanByDoiVT('ToKT_SonTay');
        }

        // Render first tab (TK_ToKT_SonTay) by default
        if (data.sheets && data.sheets['TK_ToKT_SonTay']) {
            console.log('Rendering TK_ToKT_SonTay data...');
            renderPTTBTable('pttb-detail-tabs-content', data.sheets['TK_ToKT_SonTay'], window.pttbDetailFileInfo);
        } else {
            console.error('TK_ToKT_SonTay sheet not found. Available sheets:', Object.keys(data.sheets || {}));
        }
    } catch (error) {
        console.error('Error loading PTTB detail:', error);
        showPTTBError('pttb-detail-diaban-table', error.message);
        showPTTBError('pttb-detail-tabs-content', error.message);
    }
}

/**
 * Load PTTB Pending data (chua_co_lydoton, phieu_qua_gio)
 */
async function loadPTTBPendingData() {
    try {
        const response = await fetch('/api/pttb-data-pending');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        if (data.error) {
            throw new Error(data.error);
        }

        // Store data globally for tab switching
        window.pttbPendingData = data.sheets;
        window.pttbPendingFileInfo = data.file_info || null;

        // Render first tab (chua_co_lydoton) by default
        if (data.sheets && data.sheets['chua_co_lydoton']) {
            renderPTTBTable('pttb-pending-tabs-content', data.sheets['chua_co_lydoton'], window.pttbPendingFileInfo);
        }
    } catch (error) {
        console.error('Error loading PTTB pending:', error);
        showPTTBError('pttb-pending-tabs-content', error.message);
    }
}

/**
 * Initialize dia ban tabs (filter by DOI_VT)
 */
function initPTTBDiabanTabs() {
    const tabs = document.querySelectorAll('#pttb-diaban-tabs .excel-tab');

    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            tabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const doiVT = this.dataset.doi;
            filterDiabanByDoiVT(doiVT);
        });
    });
}

/**
 * Filter dia ban data by DOI_VT
 */
function filterDiabanByDoiVT(doiVT) {
    if (!window.pttbDiabanData) {
        console.error('No diaban data available');
        return;
    }

    const originalData = window.pttbDiabanData;
    let filteredSheetData;

    if (doiVT === 'all') {
        // Show all data
        filteredSheetData = originalData;
    } else {
        // Filter by DOI_VT
        const filteredRows = originalData.data.filter(row => row['DOI_VT'] === doiVT);
        filteredSheetData = {
            columns: originalData.columns,
            data: filteredRows
        };
    }

    renderPTTBTable('pttb-detail-diaban-table', filteredSheetData, window.pttbDetailFileInfo);
}

/**
 * Initialize detail tabs (4 tổ)
 */
function initPTTBDetailTabs() {
    const tabs = document.querySelectorAll('#pttb-detail-tabs .excel-tab');

    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            tabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const sheetName = this.dataset.sheet;
            if (window.pttbDetailData && window.pttbDetailData[sheetName]) {
                renderPTTBTable('pttb-detail-tabs-content', window.pttbDetailData[sheetName], window.pttbDetailFileInfo);
            }
        });
    });
}

/**
 * Initialize pending tabs (2 tabs)
 */
function initPTTBPendingTabs() {
    const tabs = document.querySelectorAll('#pttb-pending-tabs .excel-tab');

    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            tabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const sheetName = this.dataset.sheet;
            if (window.pttbPendingData && window.pttbPendingData[sheetName]) {
                renderPTTBTable('pttb-pending-tabs-content', window.pttbPendingData[sheetName], window.pttbPendingFileInfo);
            }
        });
    });
}

/**
 * Render PTTB table
 */
function renderPTTBTable(containerId, sheetData, fileInfo = null) {
    const container = document.getElementById(containerId);
    if (!container) {
        console.error(`Container not found: ${containerId}`);
        return;
    }

    console.log(`Rendering table in ${containerId}, data:`, sheetData);

    container.innerHTML = '';

    if (!sheetData || !sheetData.columns || !sheetData.data) {
        console.error(`Invalid sheet data for ${containerId}:`, sheetData);
        container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
        return;
    }

    const { columns, data } = sheetData;

    if (data.length === 0) {
        container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
        return;
    }

    console.log(`Creating table with ${data.length} rows, ${columns.length} columns`);

    // Create Excel table using the helper function
    const tableMaxHeight = containerId === 'pttb-chitiet-to-content' ? '820px' : '600px';

    const tableHtml = createExcelTable(sheetData, '', {
        showRowNumbers: true,
        tableClass: 'excel-table',
        maxHeight: tableMaxHeight,
        fileInfo: fileInfo
    });

    container.innerHTML = tableHtml;
    console.log(`Table rendered successfully in ${containerId}`);
}

/**
 * Show error message
 */
function showPTTBError(containerId, message) {
    const container = document.getElementById(containerId);
    if (!container) return;

    container.innerHTML = `
        <div class="loading">
            <i class="fas fa-exclamation-triangle"></i>
            <br>Lỗi khi tải dữ liệu: ${message}
        </div>
    `;
}

/**
 * Download Excel PTTB
 */
function downloadExcelPTTB() {
    window.location.href = '/download/excel-pttb';
}

/**
 * Load PTTB Chi tiết tồn các tổ data
 */
async function loadPTTBChitietToData() {
    try {
        const response = await fetch('/api/pttb-data-chitiet-to');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        if (data.error) {
            throw new Error(data.error);
        }

        // Store data globally for tab switching
        window.pttbChitietToData = data.sheets;
        window.pttbChitietToFileInfo = data.file_info || null;

        // Render first tab (ToKT_SonTay) by default
        if (data.sheets && data.sheets['ToKT_SonTay']) {
            renderPTTBTable('pttb-chitiet-to-content', data.sheets['ToKT_SonTay'], window.pttbChitietToFileInfo);
        }
    } catch (error) {
        console.error('Error loading PTTB Chitiet To:', error);
        showPTTBError('pttb-chitiet-to-content', error.message);
    }
}

/**
 * Initialize Chi tiết tồn các tổ tabs (4 tabs)
 */
function initPTTBChitietToTabs() {
    const tabs = document.querySelectorAll('#pttb-chitiet-to-tabs .excel-tab');

    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            tabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const sheetName = this.dataset.sheet;
            if (window.pttbChitietToData && window.pttbChitietToData[sheetName]) {
                renderPTTBTable('pttb-chitiet-to-content', window.pttbChitietToData[sheetName], window.pttbChitietToFileInfo);
            }
        });
    });
}
