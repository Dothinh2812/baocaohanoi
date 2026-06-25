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
    ]);

    // Initialize detail tabs after data is loaded
    initPTTBDiabanTabs();
    initPTTBDetailTabs();
    initPTTBPendingTabs();
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
 * Load PTTB Chi tiết tồn các tổ data (từ kiêm soát endpoint, có cột kiểm soát inline)
 */
let _pttbKsActiveTeam = null;

async function loadPTTBChitietToData() {
    try {
        const data = await API.getPttbKiemSoatDetail();
        renderPttbKiemSoatMain(data);
    } catch (error) {
        console.error('Error loading PTTB Chitiet To:', error);
        showPTTBError('pttb-chitiet-to-content', error.message);
    }
}

function renderPttbKiemSoatMain(data) {
    const container = document.getElementById('pttb-chitiet-to-content');
    const tabsEl = document.getElementById('pttb-chitiet-to-tabs');
    if (!data || !data.sheets || !container || !tabsEl) {
        if (container) container.innerHTML = '<p>Không có dữ liệu phiếu tồn.</p>';
        return;
    }

    const teamNames = Object.keys(data.sheets);
    if (teamNames.length === 0) {
        tabsEl.innerHTML = '';
        container.innerHTML = '<p>Không có phiếu tồn.</p>';
        return;
    }

    // populate bộ lọc tổ (thống kê) cùng lúc
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

    if (!_pttbKsActiveTeam || !teamNames.includes(_pttbKsActiveTeam)) {
        _pttbKsActiveTeam = teamNames[0];
    }

    tabsEl.innerHTML = teamNames.map(name => {
        const active = name === _pttbKsActiveTeam ? 'active' : '';
        const m = name.match(/ToKT_(\w+?)(?:_rut_gon)?$/);
        const key = m ? m[1] : name;
        const displayNameMap = { 'PhucTho': 'Phúc Thọ', 'SonTay': 'Sơn Tây', 'QuangOai': 'Quảng Oai', 'SuoiHai': 'Suối Hai' };
        const display = displayNameMap[key] || key;
        return `<li class="excel-tab ${active}" data-sheet="${name}">${display}</li>`;
    }).join('');

    tabsEl.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            _pttbKsActiveTeam = this.dataset.sheet;
            renderPttbKiemSoatMain(data);
        });
    });

    renderPttbKiemSoatTable(_pttbKsActiveTeam, data.sheets[_pttbKsActiveTeam], container, data.file_info);
}

function renderPttbKiemSoatTable(team, sheetData, container, fileInfo) {
    if (!container) return;
    const rows = (sheetData && sheetData.data) || [];

    const colgroup = '<colgroup>' +
        PTTB_KS_DISPLAY_COLS.map(c => `<col style="width:${PTTB_KS_COL_WIDTHS[c] || 'auto'}">`).join('') +
        '<col style="width:18%">' +
        '</colgroup>';

    const headers = PTTB_KS_DISPLAY_COLS.map(c => `<th>${PTTB_KS_DISPLAY_LABELS[c] || c}</th>`).join('');
    const body = rows.map(row => {
        const cells = PTTB_KS_DISPLAY_COLS.map(c => `<td>${row[c] != null ? row[c] : ''}</td>`).join('');
        const maTb = row.MA_THUE_BAO;
        const noiDung = (row.kiemsoat_noi_dung || '').toString()
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return `
            <tr>
                ${cells}
                <td class="pttb-ks-cell">
                    <textarea class="pttb-ks-input" rows="2" data-ma_tb="${maTb}"
                        data-loai="${row.LOAIHINH_TB || ''}" data-doi="${row.DOI_VT || ''}" data-nvtt="${row.NHANVIEN_TIEPTHI || ''}">${noiDung}</textarea>
                    <button class="pttb-ks-save-btn" onclick="savePttbKiemSoat('${maTb}')">Lưu</button>
                    <span class="pttb-ks-status" id="pttb-ks-status-${maTb}">${_buildPttbKsBadge(row)}</span>
                </td>
            </tr>`;
    }).join('');

    const ts = (fileInfo && fileInfo.modified)
        ? `<div style="color:#d32f2f;font-weight:600;font-size:0.85rem;margin-bottom:6px;"><i class="fas fa-clock"></i> Dữ liệu báo cáo cập nhật: ${fileInfo.modified}</div>`
        : '';

    container.innerHTML = `
        ${ts}
        <div class="excel-table-card">
            <div class="excel-table-body" style="max-height:820px;overflow:auto;">
                <table class="excel-table pttb-ks-detail-table">
                    ${colgroup}
                    <thead><tr>${headers}<th>Nội dung kiểm soát</th></tr></thead>
                    <tbody>${body || '<tr><td colspan="99">Không có phiếu.</td></tr>'}</tbody>
                </table>
            </div>
        </div>`;
}

/* ========================================
   KIỂM SOÁT TỔ TRƯỞNG PTTB
   - Inline edit nội dung kiểm soát / phiếu
   - Thống kê đã/chưa + thời điểm nhập
   ======================================== */

const PTTB_KS_DISPLAY_COLS = [
    'MA_THUE_BAO', 'TEN_THUEBAO', 'DIACHI_LAPDAT', 'LOAIHINH_TB',
    'NHANVIEN_TIEPTHI', 'DOI_VT', 'TEN_KV', 'NGAYHEN_DEN',
    'NOIDUNG_HEN', 'chitieu_tg', 'gio_conlai', 'trang_thai',
];
const PTTB_KS_DISPLAY_LABELS = {
    'MA_THUE_BAO': 'Mã TB', 'TEN_THUEBAO': 'Khách hàng', 'DIACHI_LAPDAT': 'Địa chỉ',
    'LOAIHINH_TB': 'Loại', 'NHANVIEN_TIEPTHI': 'NVTT', 'DOI_VT': 'Tổ',
    'TEN_KV': 'Khu vực', 'NGAYHEN_DEN': 'Ngày hẹn', 'NOIDUNG_HEN': 'Nội dung hẹn',
    'chitieu_tg': 'Chỉ tiêu', 'gio_conlai': 'Giờ còn lại', 'trang_thai': 'Trạng thái',
};
const PTTB_KS_COL_WIDTHS = {
    'MA_THUE_BAO': '7%', 'TEN_THUEBAO': '10%', 'DIACHI_LAPDAT': '12%',
    'LOAIHINH_TB': '6%', 'NHANVIEN_TIEPTHI': '7%', 'DOI_VT': '7%',
    'TEN_KV': '7%', 'NGAYHEN_DEN': '8%', 'NOIDUNG_HEN': '10%',
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
    const chiTietEl = document.getElementById('pttb-kiemsoat-chitiet-container');
    try {
        const data = await API.getPttbKiemSoatThongKe(_pttbKiemSoatQuery());
        renderPttbKiemSoatStats(data, statsEl);
        renderPttbKiemSoatChiTiet(data.chi_tiet || [], chiTietEl);
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

function renderPttbKiemSoatChiTiet(rows, container) {
    if (!container) return;
    if (!rows.length) {
        container.innerHTML = '<p>Không có phiếu khớp bộ lọc.</p>';
        return;
    }
    const cols = ['MA_THUE_BAO', 'TEN_THUEBAO', 'LOAIHINH_TB', 'NHANVIEN_TIEPTHI', 'DOI_VT', 'gio_conlai', 'trang_thai', 'kiemsoat_noi_dung', 'kiemsoat_thoi_diem', 'kiemsoat_nguoi_nhap'];
    const labels = {
        'MA_THUE_BAO': 'Mã TB', 'TEN_THUEBAO': 'Khách hàng', 'LOAIHINH_TB': 'Loại',
        'NHANVIEN_TIEPTHI': 'NVTT', 'DOI_VT': 'Tổ', 'gio_conlai': 'Giờ còn lại',
        'trang_thai': 'Trạng thái', 'kiemsoat_noi_dung': 'Nội dung KS',
        'kiemsoat_thoi_diem': 'Thời điểm nhập', 'kiemsoat_nguoi_nhap': 'Người nhập',
    };
    const headers = cols.map(c => `<th>${labels[c] || c}</th>`).join('');
    const body = rows.map(r => `<tr>${cols.map(c => `<td>${r[c] != null ? r[c] : ''}</td>`).join('')}</tr>`).join('');
    container.innerHTML = `
        <div class="excel-table-card">
            <div class="excel-table-body">
                <table class="excel-table summary-table">
                    <thead><tr>${headers}</tr></thead>
                    <tbody>${body}</tbody>
                </table>
            </div>
        </div>`;
}

async function reloadPttbKiemSoatThongKe() {
    await loadPttbKiemSoatThongKe();
    await loadPTTBChitietToData();
}
window.reloadPttbKiemSoatThongKe = reloadPttbKiemSoatThongKe;

function exportPttbKiemSoat() {
    const q = _pttbKiemSoatQuery();
    window.location.href = '/download/pttb-kiemsoat-report' + (q ? '?' + q : '');
}
window.exportPttbKiemSoat = exportPttbKiemSoat;
