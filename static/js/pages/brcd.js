/* ========================================
   BRCD PAGE JAVASCRIPT
   - Load BRCD data
   - Render charts and tables
   - Handle tabs for địa bàn
   ======================================== */

document.addEventListener('DOMContentLoaded', async function () {
    await initBRCDPage();
});

async function initBRCDPage() {
    // Initialize địa bàn tabs
    initDiabanTabs();

    // Initialize NVKT tabs
    initNVKTTabs();

    // Load all BRCD data
    await Promise.all([
        loadBRCDData(),
        loadBRCDMainData(),
        loadBRCDPendingData(),
        loadKiemSoatThongKe(),
    ]);
}

/**
 * Initialize địa bàn (4 tổ) tabs
 */
function initDiabanTabs() {
    const tabs = document.querySelectorAll('#diaban-3to-tabs .image-tab');
    const chartDisplay = document.getElementById('diaban-chart-display');

    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            // Remove active class from all tabs
            tabs.forEach(t => t.classList.remove('active'));
            // Add active class to clicked tab
            this.classList.add('active');

            // Get chart name from data attribute
            const chartName = this.getAttribute('data-chart');
            // Update image source
            chartDisplay.src = `/chart/Chart_diaban_${chartName}.png`;
            chartDisplay.alt = `Biểu đồ địa bàn ${this.textContent}`;
        });
    });
}

/**
 * Initialize NVKT tabs - filter table by team
 */
function initNVKTTabs() {
    const tabs = document.querySelectorAll('#nvkt-tabs .image-tab');
    const rows = document.querySelectorAll('.nvkt-row');

    // Filter function
    function filterByTeam(teamName) {
        rows.forEach(row => {
            const rowTeam = row.getAttribute('data-team');
            if (rowTeam === teamName) {
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
        });
    }

    // Tab click handler
    tabs.forEach(tab => {
        tab.addEventListener('click', function () {
            // Remove active class from all tabs
            tabs.forEach(t => t.classList.remove('active'));
            // Add active class to clicked tab
            this.classList.add('active');

            // Filter table by team
            const teamName = this.getAttribute('data-team');
            filterByTeam(teamName);
        });
    });

    // Initialize with first tab (PhucTho)
    filterByTeam('PhucTho');
}

/**
 * Load BRCD down port data
 */
async function loadBRCDData() {
    try {
        const data = await API.getBRCDData();

        if (data && data.sheets) {
            const fileInfo = data.file_info || null;

            renderExcelTabs('excel-tabs', Object.keys(data.sheets), (sheetName) => {
                renderExcelTable('excel-tables-container', data.sheets[sheetName], sheetName, fileInfo);
            });

            // Render first sheet by default
            const firstSheet = Object.keys(data.sheets)[0];
            renderExcelTable('excel-tables-container', data.sheets[firstSheet], firstSheet, fileInfo);
        }
    } catch (error) {
        showError('Không thể tải dữ liệu BRCD down port', 'excel-tables-container');
    }
}

/**
 * Load BRCD main data (4 teams) — chi tiết phiếu tồn + cột kiểm soát tổ trưởng
 * Đọc sheet đầy đủ ToKT_<doi> (qua /api/brcd-kiemsoat/detail) để có baohong_id làm khóa.
 */
async function loadBRCDMainData() {
    try {
        const data = await API.getBrcdKiemSoatDetail();
        renderKiemSoatMain(data);
    } catch (error) {
        showError('Không thể tải dữ liệu BRCD 4 tổ', 'excel-tables-container-main');
    }
}

/**
 * Load BRCD pending tickets
 */
async function loadBRCDPendingData() {
    try {
        const data = await API.getBRCDPendingData();

        if (data && data.sheets) {
            const fileInfo = data.file_info || null;

            renderExcelTabs('excel-tabs-pending', Object.keys(data.sheets), (sheetName) => {
                renderExcelTable('excel-tables-container-pending', data.sheets[sheetName], sheetName, fileInfo);
            });

            // Render first sheet by default
            const firstSheet = Object.keys(data.sheets)[0];
            renderExcelTable('excel-tables-container-pending', data.sheets[firstSheet], firstSheet, fileInfo);
        }
    } catch (error) {
        showError('Không thể tải dữ liệu phiếu pending', 'excel-tables-container-pending');
    }
}

/**
 * Convert sheet name to display name
 * e.g., "ToKT_PhucTho_rut_gon" -> "Phúc Thọ"
 */
function formatTabDisplayName(sheetName) {
    // Mapping for friendly display names
    const displayNameMap = {
        'PhucTho': 'Phúc Thọ',
        'SonTay': 'Sơn Tây',
        'QuangOai': 'Quảng Oai',
        'SuoiHai': 'Suối Hai'
    };

    // Extract the team name from patterns like "ToKT_PhucTho_rut_gon" or "ToKT_PhucTho"
    const match = sheetName.match(/ToKT_(\w+?)(?:_rut_gon)?$/);
    if (match && match[1]) {
        const teamKey = match[1];
        return displayNameMap[teamKey] || teamKey;
    }

    // Fallback: remove common prefixes/suffixes for cleaner display
    return sheetName
        .replace(/^ToKT_/, '')
        .replace(/_rut_gon$/, '')
        .replace(/_/g, ' ');
}

/**
 * Render Excel tabs
 */
function renderExcelTabs(tabsId, sheetNames, onTabClick) {
    const tabsContainer = document.getElementById(tabsId);
    if (!tabsContainer) return;

    let html = '';
    sheetNames.forEach((name, index) => {
        const activeClass = index === 0 ? 'active' : '';
        const displayName = formatTabDisplayName(name);
        html += `<li class="excel-tab ${activeClass}" data-sheet="${name}">${displayName}</li>`;
    });

    tabsContainer.innerHTML = html;

    // Add click handlers
    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(t => t.classList.remove('active'));
            this.classList.add('active');
            onTabClick(this.dataset.sheet);
        });
    });
}

/**
 * Render Excel table
 */
function renderExcelTable(containerId, sheetData, sheetName, fileInfo = null) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const html = createExcelTable(sheetData, sheetName, {
        showRowNumbers: true,
        maxHeight: 'none',
        fileInfo: fileInfo
    });

    container.innerHTML = html;
}

/**
 * Download Excel handlers
 */
function downloadExcel() {
    window.location.href = '/download/excel';
}

function downloadExcelMain() {
    window.location.href = '/download/excel-main';
}

// Export functions to global scope
window.downloadExcel = downloadExcel;
window.downloadExcelMain = downloadExcelMain;


/* ========================================
   KIỂM SOÁT TỔ TRƯỞNG
   - Inline edit nội dung kiểm soát / phiếu
   - Thống kê đã/chưa + thời điểm nhập
   ======================================== */

const KIEMSOAT_DISPLAY_COLS = [
    'ma_tb', 'TEN_TB', 'DIACHI_LD', 'LOAIHINH_TB', 'GHICHU_HONG',
    'NVKT', 'ngay_bh', 'Trạng thái cổng', 'giờ còn lại thực', 'ttvt_ton', 'SA',
];
const KIEMSOAT_DISPLAY_LABELS = {
    'ma_tb': 'Mã TB', 'TEN_TB': 'Khách hàng', 'LOAIHINH_TB': 'Loại', 'NVKT': 'NVKT',
    'ngay_bh': 'Báo hỏng', 'Trạng thái cổng': 'Cổng', 'giờ còn lại thực': 'Giờ còn lại',
    'ttvt_ton': 'Lý do tồn', 'DIACHI_LD': 'Địa chỉ', 'GHICHU_HONG': 'Nội dung báo', 'SA': 'SA',
};
// Width tính bằng % (table-layout: fixed, width: 100% → co về vừa container, không scroll ngang).
// 3 cột dài (Địa chỉ, Nội dung báo, Lý do tồn) được để rộng + wrap.
const KIEMSOAT_COL_WIDTHS = {
    'ma_tb': '7%', 'TEN_TB': '11%', 'DIACHI_LD': '15%', 'LOAIHINH_TB': '6%',
    'GHICHU_HONG': '14%', 'NVKT': '9%', 'ngay_bh': '10%', 'Trạng thái cổng': '6%',
    'giờ còn lại thực': '6%', 'ttvt_ton': '14%', 'SA': '8%',
};
let _kiemsoatActiveTeam = null;

function _formatKiemSoatThoiDiem(value) {
    if (!value) return '';
    // value dạng 'YYYY-MM-DD HH:MM:SS' -> 'HH:MM dd/MM'
    const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/);
    if (!m) return value;
    return `${m[4]}:${m[5]} ${m[3]}/${m[2]}`;
}

function _buildKiemSoatBadge(row) {
    if (row.kiemsoat_da_nhap) {
        const when = _formatKiemSoatThoiDiem(row.kiemsoat_thoi_diem);
        const who = row.kiemsoat_nguoi_nhap || '';
        return `<span class="ks-badge ks-da">Đã kiểm soát${when ? ' ' + when : ''}${who ? ' &mdash; ' + who : ''}</span>`;
    }
    return `<span class="ks-badge ks-chua">Chưa</span>`;
}

function renderKiemSoatMain(data) {
    const container = document.getElementById('excel-tables-container-main');
    const tabsEl = document.getElementById('excel-tabs-main');
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

    // populate bộ lọc đội (thống kê) cùng lúc
    // value phải là tên sheet gốc (= giá trị cột DOI_VT) để khớp với backend filter
    const doiSelect = document.getElementById('kiemsoat-filter-doi');
    if (doiSelect) {
        const current = doiSelect.value;
        doiSelect.innerHTML = '<option value="">Tất cả</option>' +
            teamNames.map(t => `<option value="${t}">${formatTabDisplayName(t)}</option>`).join('');
        if (current && teamNames.includes(current)) doiSelect.value = current;
    }

    if (!_kiemsoatActiveTeam || !teamNames.includes(_kiemsoatActiveTeam)) {
        _kiemsoatActiveTeam = teamNames[0];
    }

    tabsEl.innerHTML = teamNames.map(name => {
        const active = name === _kiemsoatActiveTeam ? 'active' : '';
        return `<li class="excel-tab ${active}" data-sheet="${name}">${formatTabDisplayName(name)}</li>`;
    }).join('');
    tabsEl.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            _kiemsoatActiveTeam = this.dataset.sheet;
            renderKiemSoatMain(data);
        });
    });

    renderKiemSoatTable(_kiemsoatActiveTeam, data.sheets[_kiemsoatActiveTeam], container);
}

function renderKiemSoatTable(team, sheetData, container) {
    if (!container) return;
    const rows = (sheetData && sheetData.data) || [];

    const colgroup = '<colgroup>' +
        KIEMSOAT_DISPLAY_COLS.map(c => `<col style="width:${KIEMSOAT_COL_WIDTHS[c] || 'auto'}">`).join('') +
        '<col style="width:18%">' +  // cột Nội dung kiểm soát
        '</colgroup>';

    const headers = KIEMSOAT_DISPLAY_COLS.map(c => `<th>${KIEMSOAT_DISPLAY_LABELS[c] || c}</th>`).join('');
    const body = rows.map(row => {
        const cells = KIEMSOAT_DISPLAY_COLS.map(c => `<td>${row[c] != null ? row[c] : ''}</td>`).join('');
        const baohong = row.baohong_id;
        const noiDung = (row.kiemsoat_noi_dung || '').toString()
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return `
            <tr>
                ${cells}
                <td class="ks-cell">
                    <textarea class="ks-input" rows="2" data-baohong="${baohong}"
                        data-ma_tb="${row.ma_tb || ''}" data-doi="${row.DOI_VT || ''}" data-nvkt="${row.NVKT || ''}">${noiDung}</textarea>
                    <button class="ks-save-btn" onclick="saveKiemSoat(${baohong})">Lưu</button>
                    <span class="ks-status" id="ks-status-${baohong}">${_buildKiemSoatBadge(row)}</span>
                </td>
            </tr>`;
    }).join('');

    container.innerHTML = `
        <div class="excel-table-card">
            <div class="excel-table-body">
                <table class="excel-table brcd-detail-table">
                    ${colgroup}
                    <thead><tr>${headers}<th>Nội dung kiểm soát</th></tr></thead>
                    <tbody>${body || '<tr><td colspan="99">Không có phiếu.</td></tr>'}</tbody>
                </table>
            </div>
        </div>`;
}

async function saveKiemSoat(baohongId) {
    const textarea = document.querySelector(`.ks-input[data-baohong="${baohongId}"]`);
    if (!textarea) return;
    const statusEl = document.getElementById(`ks-status-${baohongId}`);
    try {
        const result = await API.saveBrcdKiemSoat({
            baohong_id: baohongId,
            ma_tb: textarea.dataset.ma_tb,
            doi_vt: textarea.dataset.doi,
            nvkt: textarea.dataset.nvkt,
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
            statusEl.innerHTML = _buildKiemSoatBadge(fakeRow);
        }
        await loadKiemSoatThongKe();
    } catch (error) {
        alert('Không lưu được: ' + error.message);
    }
}
window.saveKiemSoat = saveKiemSoat;

function _kiemSoatQuery() {
    const params = new URLSearchParams();
    const nhom = document.getElementById('kiemsoat-filter-nhom');
    const trangthai = document.getElementById('kiemsoat-filter-trangthai');
    const doi = document.getElementById('kiemsoat-filter-doi');
    const khoang = document.getElementById('kiemsoat-filter-khoang');
    if (nhom && nhom.value) params.set('nhom', nhom.value);
    if (trangthai && trangthai.value) params.set('trangthai', trangthai.value);
    if (doi && doi.value) params.set('doi', doi.value);
    if (khoang && khoang.value) params.set('khoang', khoang.value);
    return params.toString();
}

async function loadKiemSoatThongKe() {
    const statsEl = document.getElementById('kiemsoat-stats');
    const byDoiEl = document.getElementById('kiemsoat-by-doi');
    const chiTietEl = document.getElementById('kiemsoat-chitiet-container');
    try {
        const data = await API.getBrcdKiemSoatThongKe(_kiemSoatQuery());
        renderKiemSoatStats(data, statsEl, byDoiEl);
        renderKiemSoatChiTiet(data.chi_tiet || [], chiTietEl);
    } catch (error) {
        if (statsEl) statsEl.innerHTML = '<p class="error">Không tải được thống kê.</p>';
    }
}

function renderKiemSoatStats(data, statsEl, byDoiEl) {
    if (!statsEl) return;
    const s = data.summary || {};
    const ls = data.lich_su || {};
    const card = (num, label, cls) => `<div class="ks-card ${cls || ''}"><div class="ks-num">${num}</div><div class="ks-label">${label}</div></div>`;
    statsEl.innerHTML = `
        <div class="kiemsoat-stats-cards">
            ${card(s.total ?? 0, 'Tổng tồn', 'ks-total')}
            ${card(s.da_kiem_soat ?? 0, 'Đã kiểm soát', 'ks-da-card')}
            ${card(s.chua ?? 0, 'Chưa kiểm soát', 'ks-chua-card')}
            ${card((s.ty_le ?? 0) + '%', 'Tỉ lệ', 'ks-ty-le')}
            ${card(ls.roi_da_ks ?? 0, 'Rời tồn đã KS', 'ks-roi-da')}
            ${card(ls.roi_chua_ks ?? 0, 'Rời tồn chưa KS', 'ks-roi-chua')}
        </div>`;

    if (byDoiEl) {
        const rows = (data.by_doi || []);
        if (rows.length === 0) { byDoiEl.innerHTML = ''; return; }
        byDoiEl.innerHTML = `
            <table class="excel-table summary-table">
                <thead><tr><th>Đội</th><th>Tổng</th><th>Đã KS</th><th>Chưa</th><th>Tỉ lệ</th></tr></thead>
                <tbody>${rows.map(r => `<tr><td>${r.DOI_VT || ''}</td><td>${r.total}</td><td>${r.da_kiem_soat}</td><td>${r.chua}</td><td>${r.ty_le}%</td></tr>`).join('')}</tbody>
            </table>`;
    }
}

function renderKiemSoatChiTiet(rows, container) {
    if (!container) return;
    if (!rows.length) { container.innerHTML = '<p>Không có phiếu khớp bộ lọc.</p>'; return; }
    const cols = ['ma_tb', 'NVKT', 'DOI_VT', 'Trạng thái cổng', 'giờ còn lại thực', 'kiemsoat_noi_dung', 'kiemsoat_thoi_diem', 'kiemsoat_nguoi_nhap'];
    const labels = { 'ma_tb': 'Mã TB', 'Trạng thái cổng': 'Cổng', 'giờ còn lại thực': 'Giờ còn lại', 'kiemsoat_noi_dung': 'Nội dung KS', 'kiemsoat_thoi_diem': 'Thời điểm nhập', 'kiemsoat_nguoi_nhap': 'Người nhập' };
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

async function reloadKiemSoatThongKe() {
    await loadKiemSoatThongKe();
}
window.reloadKiemSoatThongKe = reloadKiemSoatThongKe;

function exportBrcdKiemSoat() {
    const q = _kiemSoatQuery();
    window.location.href = '/download/brcd-kiemsoat-report' + (q ? '?' + q : '');
}
window.exportBrcdKiemSoat = exportBrcdKiemSoat;
