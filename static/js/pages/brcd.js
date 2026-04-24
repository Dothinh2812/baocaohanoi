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
        loadBRCDPendingData()
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
 * Load BRCD main data (4 teams)
 */
async function loadBRCDMainData() {
    try {
        const data = await API.getBRCDMainData();

        if (data && data.sheets) {
            const fileInfo = data.file_info || null;

            // Chỉ hiển thị các sheet có tên chứa '_rut_gon'
            const filteredSheetNames = Object.keys(data.sheets).filter(name => name.includes('_rut_gon'));

            if (filteredSheetNames.length > 0) {
                renderExcelTabs('excel-tabs-main', filteredSheetNames, (sheetName) => {
                    renderExcelTable('excel-tables-container-main', data.sheets[sheetName], sheetName, fileInfo);
                });

                // Render first filtered sheet by default
                const firstSheet = filteredSheetNames[0];
                renderExcelTable('excel-tables-container-main', data.sheets[firstSheet], firstSheet, fileInfo);
            }
        }
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
