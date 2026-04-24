/* ========================================
   I1.5 K2 PAGE JAVASCRIPT
   - Load network quality data (Kỳ 2)
   - Render SHC tables
   ======================================== */

document.addEventListener('DOMContentLoaded', async function () {
    bindDateFilter();
    await initI15K2Page();
});

async function initI15K2Page() {
    await initSHCDetailDownloader('k2');
    await loadI15K2Data();
    await loadSHCVariationK2Data();
}

function getRequestedDate() {
    const params = new URLSearchParams(window.location.search);
    return params.get('date') || '';
}

function bindDateFilter() {
    const input = document.getElementById('i15k2-date-input');
    const applyButton = document.getElementById('i15k2-date-apply');
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
    input.addEventListener('keydown', function (event) {
        if (event.key === 'Enter') {
            applyDateFilter();
        }
    });
}

function syncI15K2DateState(data) {
    const input = document.getElementById('i15k2-date-input');
    if (input) {
        input.value = data.selected_date || getRequestedDate();
    }

    const meta = document.getElementById('i15k2-date-meta');
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

/**
 * Load I1.5 K2 data
 */
async function loadI15K2Data() {
    try {
        const endpoint = getRequestedDate()
            ? `/api/i15k2-data?date=${encodeURIComponent(getRequestedDate())}`
            : '/api/i15k2-data';
        const data = await API.fetchData(endpoint);

        if (data) {
            const fileInfo = data.file_info || null;
            syncI15K2DateState(data);

            // Bảng 1: Tổng hợp SHC theo Tổ
            if (data.tong_hop) {
                const summaryHtml = createExcelTable(data.tong_hop, 'Tổng hợp SHC theo Tổ', {
                    showRowNumbers: false,
                    maxHeight: '600px',
                    fileInfo: fileInfo,
                    columnLabels: {
                        stt: 'STT',
                        tt: 'TT',
                        don_vi: 'Đơn vị',
                        so_tb_suy_hao_k1: 'Số TB suy hao K1',
                        so_tb_quan_ly: 'Số TB quản lý',
                        ti_le_shc: 'Tỷ lệ SHC (%)'
                    }
                });
                document.getElementById('i15k2-table-container').innerHTML = summaryHtml;
            }

            // Bảng 2-5: Chi tiết SHC theo Đơn vị
            if (data.don_vi && Object.keys(data.don_vi).length > 0) {
                const donViNames = Object.keys(data.don_vi);
                renderI15K2DetailTabs(donViNames);
                renderI15K2DetailTable(data.don_vi[donViNames[0]], donViNames[0], fileInfo);

                // Store data for tab switching
                window.i15k2DetailData = data;
                window.i15k2FileInfo = fileInfo;
            }

            // Bảng 6: TH SHC theo SA
            if (data.shc_theo_sa) {
                const saHtml = createExcelTable(data.shc_theo_sa, 'TH SHC theo SA', {
                    showRowNumbers: true,
                    maxHeight: '600px',
                    fileInfo: fileInfo
                });
                document.getElementById('i15k2-sa-table-container').innerHTML = saHtml;
            }
        }
    } catch (error) {
        console.error('Error loading I1.5 K2 data:', error);
        showError('Không thể tải dữ liệu I1.5 K2: ' + error.message, 'i15k2-table-container');
    }
}

/**
 * Render I1.5 K2 detail tabs
 */
function renderI15K2DetailTabs(sheetNames) {
    const tabsContainer = document.getElementById('i15k2-don-vi-tabs');
    if (!tabsContainer) return;

    let html = '';
    sheetNames.forEach((name, index) => {
        const activeClass = index === 0 ? 'active' : '';
        html += `<li class="excel-tab ${activeClass}" data-sheet="${name}">${name}</li>`;
    });

    tabsContainer.innerHTML = html;

    // Add click handlers
    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const donViName = this.dataset.sheet;
            renderI15K2DetailTable(window.i15k2DetailData.don_vi[donViName], donViName, window.i15k2FileInfo);
        });
    });
}

/**
 * Render I1.5 K2 detail table
 */
function renderI15K2DetailTable(sheetData, sheetName, fileInfo = null) {
    const html = createExcelTable(sheetData, `Chi tiết SHC - ${sheetName}`, {
        showRowNumbers: true,
        maxHeight: '600px',
        fileInfo: fileInfo
    });

    document.getElementById('i15k2-don-vi-tables-container').innerHTML = html;
}

/**
 * Download I1.5 K2 Excel file
 */
function downloadExcelI15K2() {
    const requestedDate = getRequestedDate();
    const query = requestedDate ? `?date=${encodeURIComponent(requestedDate)}` : '';
    window.location.href = `/download/excel-i15k2${query}`;
}

/**
 * Load SHC K2 Variation data (T-1 comparison)
 */
async function loadSHCVariationK2Data() {
    try {
        const endpoint = getRequestedDate()
            ? `/api/shc-variation-k2-data?date=${encodeURIComponent(getRequestedDate())}`
            : '/api/shc-variation-k2-data';
        const data = await API.fetchData(endpoint);

        if (data) {
            const fileInfo = data.file_info || null;

            // Biến động SHC theo đơn vị
            if (data.theo_don_vi) {
                const donViHtml = createExcelTable(data.theo_don_vi, 'Biến động SHC theo đơn vị', {
                    showRowNumbers: true,
                    maxHeight: '400px',
                    fileInfo: fileInfo
                });
                document.getElementById('shc-variation-don-vi-k2-container').innerHTML = donViHtml;
            }

            // Biến động SHC theo NVKT - with tabs by Đơn vị
            if (data.chi_tiet_nvkt && Object.keys(data.chi_tiet_nvkt).length > 0) {
                const donViNames = Object.keys(data.chi_tiet_nvkt);
                renderSHCNVKTK2Tabs(donViNames);
                renderSHCNVKTK2Table(data.chi_tiet_nvkt[donViNames[0]], donViNames[0], fileInfo);

                // Store data for tab switching
                window.shcNvktK2Data = data.chi_tiet_nvkt;
                window.shcNvktK2FileInfo = fileInfo;
            }
        }
    } catch (error) {
        console.error('Error loading SHC K2 Variation data:', error);
        // Show error in containers if they exist
        const donViContainer = document.getElementById('shc-variation-don-vi-k2-container');
        const nvktContainer = document.getElementById('shc-variation-nvkt-k2-container');
        if (donViContainer) {
            donViContainer.innerHTML = '<div class="error-message">Không thể tải dữ liệu biến động SHC theo đơn vị</div>';
        }
        if (nvktContainer) {
            nvktContainer.innerHTML = '<div class="error-message">Không thể tải dữ liệu biến động SHC theo NVKT</div>';
        }
    }
}

/**
 * Render SHC NVKT K2 tabs
 */
function renderSHCNVKTK2Tabs(donViNames) {
    const tabsContainer = document.getElementById('shc-nvkt-k2-tabs');
    if (!tabsContainer) return;

    let html = '';
    donViNames.forEach((name, index) => {
        const activeClass = index === 0 ? 'active' : '';
        // Shorten long names for tabs
        const shortName = name.replace('Tổ Kỹ thuật Địa bàn ', '');
        html += `<li class="excel-tab ${activeClass}" data-don-vi="${name}">${shortName}</li>`;
    });

    tabsContainer.innerHTML = html;

    // Add click handlers
    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(t => t.classList.remove('active'));
            this.classList.add('active');

            const donViName = this.dataset.donVi;
            renderSHCNVKTK2Table(window.shcNvktK2Data[donViName], donViName, window.shcNvktK2FileInfo);
        });
    });
}

/**
 * Render SHC NVKT K2 table for a specific Đơn vị
 */
function renderSHCNVKTK2Table(sheetData, donViName, fileInfo = null) {
    const html = createExcelTable(sheetData, `Biến động SHC - ${donViName}`, {
        showRowNumbers: true,
        maxHeight: 'none',
        fileInfo: fileInfo
    });

    document.getElementById('shc-variation-nvkt-k2-container').innerHTML = html;
}

async function initSHCDetailDownloader(defaultReportType = 'k2') {
    const reportTypeSelect = document.getElementById('shc-detail-report-type');
    const teamSelect = document.getElementById('shc-detail-team');
    const fileSelect = document.getElementById('shc-detail-file');
    const downloadBtn = document.getElementById('download-shc-detail-btn');
    const previewBtn = document.getElementById('preview-shc-detail-btn');

    if (!reportTypeSelect || !teamSelect || !fileSelect || !downloadBtn || !previewBtn) {
        return;
    }

    try {
        const requestedDate = getRequestedDate();
        const options = typeof API.getSHCDetailOptions === 'function'
            ? await API.getSHCDetailOptions(defaultReportType, requestedDate)
            : await API.fetchData(`/api/shc-nvkt-detail/options?report_type=${encodeURIComponent(defaultReportType)}${requestedDate ? `&date=${encodeURIComponent(requestedDate)}` : ''}`);
        const reportTypes = options.report_types || [];
        const teams = options.teams || [];

        reportTypeSelect.innerHTML = reportTypes
            .map(item => `<option value="${item.key}">${item.label}</option>`)
            .join('');

        teamSelect.innerHTML = teams
            .map(item => `<option value="${item.key}">${item.label}</option>`)
            .join('');

        if (reportTypes.some(item => item.key === defaultReportType)) {
            reportTypeSelect.value = defaultReportType;
        }

        reportTypeSelect.addEventListener('change', async () => {
            clearSHCDetailPreview();
            await loadSHCDetailFiles();
        });
        teamSelect.addEventListener('change', async () => {
            clearSHCDetailPreview();
            await loadSHCDetailFiles();
        });
        fileSelect.addEventListener('change', () => {
            syncSHCDetailActionButtons();
            clearSHCDetailPreview();
        });
        downloadBtn.addEventListener('click', downloadSelectedSHCDetailFile);
        previewBtn.addEventListener('click', previewSelectedSHCDetailFile);

        await loadSHCDetailFiles();
    } catch (error) {
        setSHCDetailStatus(`Không thể tải danh mục file chi tiết: ${error.message}`, true);
        downloadBtn.disabled = true;
        previewBtn.disabled = true;
    }
}

async function loadSHCDetailFiles() {
    const reportTypeSelect = document.getElementById('shc-detail-report-type');
    const teamSelect = document.getElementById('shc-detail-team');
    const fileSelect = document.getElementById('shc-detail-file');

    const reportType = reportTypeSelect.value;
    const team = teamSelect.value;

    if (!reportType || !team) {
        fileSelect.innerHTML = '<option value="">Chưa có dữ liệu</option>';
        syncSHCDetailActionButtons();
        return;
    }

    try {
        setSHCDetailStatus('Đang tải danh sách cá nhân...');
        const requestedDate = getRequestedDate();
        const data = typeof API.getSHCDetailFiles === 'function'
            ? await API.getSHCDetailFiles(reportType, team, requestedDate)
            : await API.fetchData(`/api/shc-nvkt-detail/files?report_type=${encodeURIComponent(reportType)}&team=${encodeURIComponent(team)}${requestedDate ? `&date=${encodeURIComponent(requestedDate)}` : ''}`);
        const files = data.files || [];

        if (files.length === 0) {
            fileSelect.innerHTML = '<option value="">Không có cá nhân phù hợp</option>';
            syncSHCDetailActionButtons();
            setSHCDetailStatus('Không tìm thấy cá nhân phù hợp cho lựa chọn hiện tại.');
            return;
        }

        fileSelect.innerHTML = files
            .map(file => `<option value="${file.name}">${file.display_name}</option>`)
            .join('');

        syncSHCDetailActionButtons();
        setSHCDetailStatus(`Tìm thấy ${files.length} cá nhân chi tiết.`);
    } catch (error) {
        fileSelect.innerHTML = '<option value="">Không thể tải danh sách cá nhân</option>';
        syncSHCDetailActionButtons();
        setSHCDetailStatus(`Không thể tải danh sách cá nhân: ${error.message}`, true);
    }
}

function downloadSelectedSHCDetailFile() {
    const reportType = document.getElementById('shc-detail-report-type').value;
    const team = document.getElementById('shc-detail-team').value;
    const fileName = document.getElementById('shc-detail-file').value;

    if (!reportType || !team || !fileName) {
        setSHCDetailStatus('Vui lòng chọn đủ loại báo cáo, tổ và file.', true);
        return;
    }

    setSHCDetailStatus(`Đang tải dữ liệu cá nhân: ${fileName}`);
    const requestedDate = getRequestedDate();
    const query = requestedDate ? `?date=${encodeURIComponent(requestedDate)}` : '';
    window.location.href = `/download/shc-nvkt-detail/${encodeURIComponent(reportType)}/${encodeURIComponent(team)}/${encodeURIComponent(fileName)}${query}`;
}

async function previewSelectedSHCDetailFile() {
    const reportType = document.getElementById('shc-detail-report-type').value;
    const team = document.getElementById('shc-detail-team').value;
    const fileName = document.getElementById('shc-detail-file').value;
    const previewSection = document.getElementById('shc-detail-preview-section');
    const previewContainer = document.getElementById('shc-detail-preview-container');

    if (!reportType || !team || !fileName) {
        setSHCDetailStatus('Vui lòng chọn đủ loại báo cáo, tổ và file.', true);
        return;
    }

    previewSection.style.display = 'block';
    showLoading('shc-detail-preview-container');
    document.getElementById('shc-detail-preview-tabs').innerHTML = '';
    document.getElementById('shc-detail-preview-meta').textContent = 'Đang tải nội dung file...';
    setSHCDetailStatus(`Đang tải xem trực tiếp: ${fileName}`);

    try {
        const requestedDate = getRequestedDate();
        const data = typeof API.getSHCDetailPreview === 'function'
            ? await API.getSHCDetailPreview(reportType, team, fileName, requestedDate)
            : await API.fetchData(`/api/shc-nvkt-detail/preview?report_type=${encodeURIComponent(reportType)}&team=${encodeURIComponent(team)}&file_name=${encodeURIComponent(fileName)}${requestedDate ? `&date=${encodeURIComponent(requestedDate)}` : ''}`);
        const sheets = data.sheets || {};
        const sheetNames = Object.keys(sheets);

        if (sheetNames.length === 0) {
            showEmptyState('File Excel không có dữ liệu để hiển thị', 'shc-detail-preview-container');
            document.getElementById('shc-detail-preview-meta').textContent = fileName;
            setSHCDetailStatus(`Không có dữ liệu trong file: ${fileName}`, true);
            return;
        }

        window.shcDetailPreviewData = data;
        window.shcDetailPreviewFileInfo = data.file_info || null;

        renderSHCDetailPreviewTabs(sheetNames);
        renderSHCDetailPreviewSheet(sheetNames[0]);

        const fileLabel = data.file_info && data.file_info.name ? data.file_info.name : fileName;
        const modifiedLabel = data.file_info && data.file_info.modified ? ` | Cập nhật: ${data.file_info.modified}` : '';
        const sheetLabel = sheetNames.length > 1 ? ` | ${sheetNames.length} sheet` : '';
        document.getElementById('shc-detail-preview-meta').textContent = `${fileLabel}${modifiedLabel}${sheetLabel}`;
        setSHCDetailStatus(`Đang hiển thị nội dung file: ${fileLabel}`);
        previewContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) {
        showError(`Không thể xem trực tiếp file: ${error.message}`, 'shc-detail-preview-container');
        document.getElementById('shc-detail-preview-meta').textContent = fileName;
        setSHCDetailStatus(`Không thể xem trực tiếp file: ${error.message}`, true);
    }
}

function renderSHCDetailPreviewTabs(sheetNames) {
    const tabsContainer = document.getElementById('shc-detail-preview-tabs');
    if (!tabsContainer) return;

    tabsContainer.innerHTML = sheetNames
        .map((name, index) => `<li class="excel-tab ${index === 0 ? 'active' : ''}" data-sheet="${name}">${name}</li>`)
        .join('');

    tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            tabsContainer.querySelectorAll('.excel-tab').forEach(item => item.classList.remove('active'));
            this.classList.add('active');
            renderSHCDetailPreviewSheet(this.dataset.sheet);
        });
    });
}

function renderSHCDetailPreviewSheet(sheetName) {
    const previewData = window.shcDetailPreviewData;
    const previewContainer = document.getElementById('shc-detail-preview-container');

    if (!previewData || !previewData.sheets || !previewData.sheets[sheetName]) {
        showEmptyState('Không có dữ liệu sheet để hiển thị', 'shc-detail-preview-container');
        return;
    }

    const sheetData = previewData.sheets[sheetName];
    if (sheetData.error) {
        showError(sheetData.error, 'shc-detail-preview-container');
        return;
    }

    previewContainer.innerHTML = createExcelTable(sheetData, `Chi tiết cá nhân - ${sheetName}`, {
        showRowNumbers: true,
        maxHeight: '600px',
        fileInfo: previewData.file_info || null
    });
}

function syncSHCDetailActionButtons() {
    const fileSelect = document.getElementById('shc-detail-file');
    const downloadBtn = document.getElementById('download-shc-detail-btn');
    const previewBtn = document.getElementById('preview-shc-detail-btn');
    const hasValue = Boolean(fileSelect && fileSelect.value);

    if (downloadBtn) {
        downloadBtn.disabled = !hasValue;
    }
    if (previewBtn) {
        previewBtn.disabled = !hasValue;
    }
}

function clearSHCDetailPreview() {
    const previewSection = document.getElementById('shc-detail-preview-section');
    const previewTabs = document.getElementById('shc-detail-preview-tabs');
    const previewContainer = document.getElementById('shc-detail-preview-container');
    const previewMeta = document.getElementById('shc-detail-preview-meta');

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

    window.shcDetailPreviewData = null;
    window.shcDetailPreviewFileInfo = null;
}

function setSHCDetailStatus(message, isError = false) {
    const statusElement = document.getElementById('shc-detail-status');
    if (!statusElement) return;

    statusElement.textContent = message;
    statusElement.style.color = isError ? '#d32f2f' : '#495057';
}
