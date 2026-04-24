document.addEventListener('DOMContentLoaded', async function () {
    await initNvktChiTietPage();
});

async function initNvktChiTietPage() {
    const root = document.getElementById('nvkt-chi-tiet-section');
    const tableContainer = document.getElementById('nvkt-detail-table');
    const chartContainer = document.getElementById('nvkt-detail-chart');
    if (!root || !tableContainer || !chartContainer) return;

    const nvktSlug = root.dataset.nvktSlug || '';
    if (!nvktSlug) return;

    try {
        const data = await API.fetchData(`/api/nvkt-chi-tiet/${encodeURIComponent(nvktSlug)}`);
        if (data.error) {
            throw new Error(data.error);
        }

        const tsEl = document.getElementById('nvkt-detail-timestamp');
        if (tsEl && data.file_modified) {
            tsEl.textContent = `(Cập nhật: ${data.file_modified})`;
        }

        if (!data.sheet || !Array.isArray(data.sheet.columns) || !Array.isArray(data.sheet.data) || data.sheet.data.length === 0) {
            tableContainer.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
            chartContainer.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu biểu đồ</p></div>';
            return;
        }

        tableContainer.innerHTML = renderNvktDetailTable(data.sheet.columns, data.sheet.data, data.file_modified);
        renderNvktDetailChart(data.sheet.data[0], chartContainer);
    } catch (error) {
        console.error('Error loading NVKT detail:', error);
        tableContainer.innerHTML = `<div class="loading"><i class="fas fa-exclamation-triangle"></i><br>Lỗi khi tải dữ liệu: ${error.message}</div>`;
        chartContainer.innerHTML = `<div class="loading"><i class="fas fa-exclamation-triangle"></i><br>Lỗi khi tải biểu đồ: ${error.message}</div>`;
    }
}

function renderNvktDetailTable(columns, rows, fileModified) {
    const timestampHtml = fileModified
        ? `<div class="excel-table-timestamp">Cập nhật: ${fileModified}</div>`
        : '';

    const headerHtml = columns.map((column, index) => {
        const stickyStyle = detailStickyStyle(index, true);
        return `<th style="${detailHeaderStyle()}${stickyStyle}">${detailEscapeHtml(column)}</th>`;
    }).join('');

    const bodyHtml = rows.map((row, rowIndex) => {
        const rowHtml = columns.map((column, colIndex) => {
            const rawValue = row[column];
            const displayValue = detailFormatCell(rawValue);
            const align = typeof rawValue === 'number' && !Number.isNaN(rawValue) ? 'right' : 'left';
            const stickyStyle = detailStickyStyle(colIndex, false, rowIndex);
            return `<td style="${detailCellStyle(align)}${stickyStyle}">${detailEscapeHtml(displayValue)}</td>`;
        }).join('');
        return `<tr>${rowHtml}</tr>`;
    }).join('');

    return `
        <div class="excel-table-card">
            <div class="excel-table-header">
                <h3 class="excel-table-title">Raw view của NVKT</h3>
                <span class="excel-table-count">${rows.length} bản ghi</span>
                ${timestampHtml}
            </div>
            <div class="excel-table-body" style="max-height: 700px; overflow: auto;">
                <table class="excel-table" style="border-collapse: separate; border-spacing: 0;">
                    <thead><tr class="header-row">${headerHtml}</tr></thead>
                    <tbody>${bodyHtml}</tbody>
                </table>
            </div>
        </div>
    `;
}

function renderNvktDetailChart(row, container) {
    const subtitle = `${detailEscapeHtml(row['to_doi_hoac_don_vi'] || '')}${row['ttvt'] ? ` | ${detailEscapeHtml(row['ttvt'])}` : ''}`;
    container.innerHTML = `
        <div class="nvkt-chart-head">
            <div class="nvkt-chart-title">${detailEscapeHtml(row['nvkt_hoac_ten_nv'] || '')}</div>
            <div class="nvkt-chart-subtitle">${subtitle}</div>
        </div>
        <div class="nvkt-chart-canvas">
            <canvas id="nvkt-detail-chart-canvas"></canvas>
        </div>
    `;

    const canvas = document.getElementById('nvkt-detail-chart-canvas');
    if (!canvas || typeof Chart === 'undefined') return;

    const volumeValues = [
        parseRawMetric(row['c11_tong_phieu']),
        parseRawMetric(row['c12_sm1_so_phieu_hll']),
        parseRawMetric(row['kqtt_tong']),
    ];

    const rateValues = [
        parseRawMetric(row['c14_ty_le_hl']),
        parseRawMetric(row['i15_k1_ty_le_shc']),
        parseRawMetric(row['i15_k2_ty_le_shc']),
        parseRawMetric(row['ghtt_ty_le_tong']),
        parseRawMetric(row['kpi_c11_chi_tieu_bsc']),
        parseRawMetric(row['kpi_c12_chi_tieu_bsc']),
    ];

    new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels: ['C11', 'C12 HLL', 'KQTT', 'C14', 'I1.5 K1', 'I1.5 K2', 'GHTT', 'KPI C11', 'KPI C12'],
            datasets: [
                {
                    type: 'bar',
                    label: 'Sản lượng',
                    data: [volumeValues[0], volumeValues[1], volumeValues[2], null, null, null, null, null, null],
                    backgroundColor: ['#215f8b', '#4c8cc2', '#8ab6d6', 'transparent', 'transparent', 'transparent', 'transparent', 'transparent', 'transparent'],
                    borderRadius: 6,
                    yAxisID: 'y',
                },
                {
                    type: 'line',
                    label: 'Tỷ lệ / điểm',
                    data: [null, null, null, rateValues[0], rateValues[1], rateValues[2], rateValues[3], rateValues[4], rateValues[5]],
                    borderColor: '#d65a31',
                    backgroundColor: 'rgba(214, 90, 49, 0.15)',
                    pointBackgroundColor: '#d65a31',
                    pointRadius: 4,
                    tension: 0.28,
                    spanGaps: true,
                    yAxisID: 'y1',
                },
            ],
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false,
            },
            plugins: {
                legend: {
                    position: 'bottom',
                },
                tooltip: {
                    callbacks: {
                        label(context) {
                            const value = context.raw;
                            if (value === null || value === undefined || Number.isNaN(value)) {
                                return `${context.dataset.label}: `;
                            }
                            return `${context.dataset.label}: ${detailFormatCell(value)}`;
                        },
                    },
                },
            },
            scales: {
                x: {
                    ticks: {
                        maxRotation: 0,
                        autoSkip: false,
                        font: { size: 10 },
                    },
                    grid: { display: false },
                },
                y: {
                    beginAtZero: true,
                    position: 'left',
                    title: { display: true, text: 'Sản lượng' },
                },
                y1: {
                    beginAtZero: true,
                    position: 'right',
                    grid: { drawOnChartArea: false },
                    title: { display: true, text: 'Tỷ lệ / điểm' },
                },
            },
        },
    });
}

function detailHeaderStyle() {
    return 'position: sticky; top: 0; background: #1a5089; color: #fff; border: 1px solid #d7e3f1; padding: 10px 8px; white-space: nowrap; text-align: center; z-index: 20;';
}

function detailCellStyle(align) {
    return `border: 1px solid #e3eaf3; padding: 8px 10px; white-space: nowrap; text-align: ${align}; background: #fff;`;
}

function detailStickyStyle(colIndex, isHeader, rowIndex = 0) {
    const widths = [180, 240, 140];
    if (colIndex >= widths.length) return '';
    const left = widths.slice(0, colIndex).reduce((sum, width) => sum + width, 0);
    const zIndex = isHeader ? 30 + (widths.length - colIndex) : 10 + (widths.length - colIndex);
    const background = isHeader ? '#1a5089' : (rowIndex % 2 === 0 ? '#ffffff' : '#f8fbff');
    const shadow = colIndex === widths.length - 1 ? 'box-shadow: 2px 0 6px rgba(0,0,0,0.08);' : '';
    return `position: sticky; left: ${left}px; min-width: ${widths[colIndex]}px; max-width: ${widths[colIndex]}px; background: ${background}; z-index: ${zIndex}; ${shadow}`;
}

function parseRawMetric(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') return Number.isNaN(value) ? null : value;
    const normalized = String(value).replace('%', '').replace(',', '.').trim();
    if (!normalized) return null;
    const numeric = Number(normalized);
    return Number.isNaN(numeric) ? null : numeric;
}

function detailFormatCell(value) {
    if (value === null || value === undefined || value === '' || (typeof value === 'number' && Number.isNaN(value))) {
        return '';
    }
    if (typeof value === 'number') {
        return Number.isInteger(value)
            ? value.toLocaleString('vi-VN')
            : value.toLocaleString('vi-VN', { maximumFractionDigits: 2 });
    }
    return String(value);
}

function detailEscapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
