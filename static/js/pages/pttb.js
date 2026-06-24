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
        loadPTTBChitietToData(),
        loadPttbKiemSoatThongKe(),
        loadPttbKiemSoatDetail(),
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

/* ========================================
   KIỂM SOÁT TỔ TRƯỞNG PTTB
   - Inline edit nội dung kiểm soát / phiếu
   - Thống kê đã/chưa + thời điểm nhập
   ======================================== */

const PTTB_KS_DISPLAY_COLS = [
    'ma_thue_bao', 'ten_thuebao', 'diachi_lapdat', 'loaihinh_tb',
    'nhanvien_tiepthi', 'doi_vt', 'ten_kv', 'ngayhen_den',
    'noidung_hen', 'chitieu_tg', 'gio_conlai', 'trang_thai',
];
const PTTB_KS_DISPLAY_LABELS = {
    'ma_thue_bao': 'Mã TB', 'ten_thuebao': 'Khách hàng', 'diachi_lapdat': 'Địa chỉ',
    'loaihinh_tb': 'Loại', 'nhanvien_tiepthi': 'NVTT', 'doi_vt': 'Tổ',
    'ten_kv': 'Khu vực', 'ngayhen_den': 'Ngày hẹn', 'noidung_hen': 'Nội dung hẹn',
    'chitieu_tg': 'Chỉ tiêu', 'gio_conlai': 'Giờ còn lại', 'trang_thai': 'Trạng thái',
};
const PTTB_KS_COL_WIDTHS = {
    'ma_thue_bao': '7%', 'ten_thuebao': '10%', 'diachi_lapdat': '12%',
    'loaihinh_tb': '6%', 'nhanvien_tiepthi': '7%', 'doi_vt': '7%',
    'ten_kv': '7%', 'ngayhen_den': '8%', 'noidung_hen': '10%',
    'chitieu_tg': '5%', 'gio_conlai': '6%', 'trang_thai': '7%',
};

function _formatPttbKsThoiDiem(value) {
    if (!value) return '';
    const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/);
    if (!m) return value;
    return `${m[4]}:${m[5]} ${m[3]}/${m[2]}`;
}

function _buildPttbKsBadge(row) {
    if (row.kiemsoat_da_nhap) {
        const when = _formatPttbKsThoiDiem(row.kiemsoat_thoi_diem);
        const who = row.kiemsoat_nguoi_nhap || '';
        return `<span class="pttb-ks-badge pttb-ks-da">Đã kiểm soát${when ? ' ' + when : ''}${who ? ' &mdash; ' + who : ''}</span>`;
    }
    return `<span class="pttb-ks-badge pttb-ks-chua">Chưa</span>`;
}

function _pttbKiemSoatQuery() {
    const params = new URLSearchParams();
    const nhom = document.getElementById('pttb-kiemsoat-filter-nhom');
    const trangthai = document.getElementById('pttb-kiemsoat-filter-trangthai');
    const to = document.getElementById('pttb-kiemsoat-filter-to');
    const khoang = document.getElementById('pttb-kiemsoat-filter-khoang');
    if (nhom && nhom.value) params.set('nhom', nhom.value);
    if (trangthai && trangthai.value) params.set('trangthai', trangthai.value);
    if (to && to.value) params.set('doi', to.value);
    if (khoang && khoang.value) params.set('khoang', khoang.value);
    return params.toString();
}

async function loadPttbKiemSoatDetail() {
    try {
        const data = await API.getPttbKiemSoatDetail();
        renderPttbKiemSoatDetail(data);
    } catch (error) {
        const container = document.getElementById('pttb-kiemsoat-chitiet-container');
        if (container) container.innerHTML = '<p>Không thể tải dữ liệu kiểm soát PTTB.</p>';
    }
}

function renderPttbKiemSoatDetail(data) {
    const container = document.getElementById('pttb-kiemsoat-chitiet-container');
    if (!container) return;
    if (!data || !data.sheets) {
        container.innerHTML = '<p>Không có dữ liệu phiếu tồn.</p>';
        return;
    }

    const teamNames = Object.keys(data.sheets);
    const toSelect = document.getElementById('pttb-kiemsoat-filter-to');
    if (toSelect) {
        const current = toSelect.value;
        const displayNameMap = { 'PhucTho': 'Phúc Thọ', 'SonTay': 'Sơn Tây', 'QuangOai': 'Quảng Oai', 'SuoiHai': 'Suối Hai' };
        toSelect.innerHTML = '<option value="">Tất cả</option>' +
            teamNames.map(t => {
                const m = t.match(/ToKT_(\w+?)(?:_rut_gon)?$/);
                const key = m ? m[1] : t;
                const display = displayNameMap[key] || key;
                return `<option value="${t}">${display}</option>`;
            }).join('');
        if (current && teamNames.includes(current)) toSelect.value = current;
    }

    let allRows = [];
    teamNames.forEach(name => {
        const sheet = data.sheets[name];
        if (sheet && sheet.data) {
            allRows = allRows.concat(sheet.data);
        }
    });

    if (!allRows.length) {
        container.innerHTML = '<p>Không có phiếu tồn.</p>';
        return;
    }

    const colgroup = '<colgroup>' +
        PTTB_KS_DISPLAY_COLS.map(c => `<col style="width:${PTTB_KS_COL_WIDTHS[c] || 'auto'}">`).join('') +
        '<col style="width:18%">' +
        '</colgroup>';

    const headers = PTTB_KS_DISPLAY_COLS.map(c => `<th>${PTTB_KS_DISPLAY_LABELS[c] || c}</th>`).join('');
    const body = allRows.map(row => {
        const cells = PTTB_KS_DISPLAY_COLS.map(c => `<td>${row[c] != null ? row[c] : ''}</td>`).join('');
        const maTb = row.ma_thue_bao;
        const noiDung = (row.kiemsoat_noi_dung || '').toString()
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return `
            <tr>
                ${cells}
                <td class="pttb-ks-cell">
                    <textarea class="pttb-ks-input" rows="2" data-ma_tb="${maTb}"
                        data-loai="${row.loaihinh_tb || ''}" data-doi="${row.DOI_VT || ''}" data-nvtt="${row.nhanvien_tiepthi || ''}">${noiDung}</textarea>
                    <button class="pttb-ks-save-btn" onclick="savePttbKiemSoat('${maTb}')">Lưu</button>
                    <span class="pttb-ks-status" id="pttb-ks-status-${maTb}">${_buildPttbKsBadge(row)}</span>
                </td>
            </tr>`;
    }).join('');

    container.innerHTML = `
        <div class="excel-table-card">
            <div class="excel-table-body" style="max-height:600px;overflow:auto;">
                <table class="excel-table pttb-ks-detail-table">
                    ${colgroup}
                    <thead><tr>${headers}<th>Nội dung kiểm soát</th></tr></thead>
                    <tbody>${body}</tbody>
                </table>
            </div>
        </div>`;
}

async function savePttbKiemSoat(maTb) {
    const textarea = document.querySelector(`.pttb-ks-input[data-ma_tb="${maTb}"]`);
    if (!textarea) return;
    const statusEl = document.getElementById(`pttb-ks-status-${maTb}`);
    try {
        const result = await API.savePttbKiemSoat({
            ma_thue_bao: maTb,
            loaihinh_tb: textarea.dataset.loai,
            doi_vt: textarea.dataset.doi,
            nhanvien_tiepthi: textarea.dataset.nvtt,
            noi_dung: textarea.value,
        });
        if (!result || result.ok === false) {
            throw new Error((result && result.error) || 'Lỗi không xác định');
        }
        if (statusEl) {
            const fakeRow = {
                kiemsoat_da_nhap: !!result.noi_dung,
                kiemsoat_thoi_diem: new Date().toISOString().replace('T', ' ').substring(0, 19),
                kiemsoat_nguoi_nhap: result.nguoi_nhap,
            };
            statusEl.innerHTML = _buildPttbKsBadge(fakeRow);
        }
        await loadPttbKiemSoatThongKe();
    } catch (error) {
        alert('Không lưu được: ' + error.message);
    }
}
window.savePttbKiemSoat = savePttbKiemSoat;

async function loadPttbKiemSoatThongKe() {
    const statsEl = document.getElementById('pttb-kiemsoat-stats');
    try {
        const data = await API.getPttbKiemSoatThongKe(_pttbKiemSoatQuery());
        renderPttbKiemSoatStats(data, statsEl);
    } catch (error) {
        if (statsEl) statsEl.innerHTML = '<p class="error">Không tải được thống kê.</p>';
    }
}

function renderPttbKiemSoatStats(data, statsEl) {
    if (!statsEl) return;
    const s = data.summary || {};
    const ls = data.lich_su || {};
    const card = (num, label, cls) => `<div class="pttb-ks-card ${cls || ''}"><div class="pttb-ks-num">${num}</div><div class="pttb-ks-label">${label}</div></div>`;
    statsEl.innerHTML = `
        <div class="pttb-ks-cards">
            ${card(s.total ?? 0, 'Tổng tồn', 'pttb-ks-total')}
            ${card(s.da_kiem_soat ?? 0, 'Đã KS', 'pttb-ks-da-card')}
            ${card(s.chua ?? 0, 'Chưa KS', 'pttb-ks-chua-card')}
            ${card((s.ty_le ?? 0) + '%', 'Tỉ lệ', 'pttb-ks-ty-le')}
            ${card(ls.roi_da_ks ?? 0, 'Rời tồn đã KS', 'pttb-ks-roi-da')}
            ${card(ls.roi_chua_ks ?? 0, 'Rời tồn chưa KS', 'pttb-ks-roi-chua')}
        </div>`;
}

async function reloadPttbKiemSoatThongKe() {
    await loadPttbKiemSoatThongKe();
    await loadPttbKiemSoatDetail();
}
window.reloadPttbKiemSoatThongKe = reloadPttbKiemSoatThongKe;

function exportPttbKiemSoat() {
    const q = _pttbKiemSoatQuery();
    window.location.href = '/download/pttb-kiemsoat-report' + (q ? '?' + q : '');
}
window.exportPttbKiemSoat = exportPttbKiemSoat;
