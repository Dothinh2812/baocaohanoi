document.addEventListener('DOMContentLoaded', async function () {
    await initKpiBchnPage();
});

async function initKpiBchnPage() {
    const container = document.getElementById('kpi-bchn-table');
    const chartsContainer = document.getElementById('kpi-bchn-charts');

    try {
        const data = await API.fetchData('/api/kpi-nvkt-bchn-data');
        if (data.error) {
            throw new Error(data.error);
        }

        const tsEl = document.getElementById('kpi-bchn-timestamp');
        if (tsEl && data.file_modified) {
            tsEl.textContent = `(Cập nhật: ${data.file_modified})`;
        }

        if (!data.sheet || !Array.isArray(data.sheet.columns) || !Array.isArray(data.sheet.data) || data.sheet.data.length === 0) {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
            if (chartsContainer) {
                chartsContainer.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu biểu đồ</p></div>';
            }
            return;
        }

        container.innerHTML = renderFlatKpiTable(data.sheet.columns, data.sheet.data, data.file_modified);
        renderNvktCharts(data.sheet.data, chartsContainer);
    } catch (error) {
        console.error('Error loading KPI BCHN data:', error);
        container.innerHTML = `
            <div class="loading">
                <i class="fas fa-exclamation-triangle"></i>
                <br>Lỗi khi tải dữ liệu: ${error.message}
            </div>`;
        if (chartsContainer) {
            chartsContainer.innerHTML = `
                <div class="loading">
                    <i class="fas fa-exclamation-triangle"></i>
                    <br>Lỗi khi tải biểu đồ: ${error.message}
                </div>`;
        }
    }
}

function renderFlatKpiTable(columns, rows, fileModified) {
    const timestampHtml = fileModified
        ? `<div class="excel-table-timestamp">Cập nhật: ${fileModified}</div>`
        : '';

    const headerHtml = columns.map((column, index) => {
        const stickyStyle = getStickyStyle(index, true);
        return `<th style="${baseHeaderStyle()}${stickyStyle}">${escapeHtml(column)}</th>`;
    }).join('');

    const bodyHtml = rows.map((row, rowIndex) => {
        const rowHtml = columns.map((column, colIndex) => {
            const rawValue = row[column];
            const displayValue = formatCellValue(rawValue);
            const alignment = isNumericValue(rawValue) ? 'right' : 'left';
            const stickyStyle = getStickyStyle(colIndex, false, rowIndex);
            return `<td style="${baseCellStyle(alignment)}${stickyStyle}">${escapeHtml(displayValue)}</td>`;
        }).join('');
        return `<tr>${rowHtml}</tr>`;
    }).join('');

    return `
        <div class="excel-table-card">
            <div class="excel-table-header">
                <h3 class="excel-table-title">NVKT tổng hợp đa nguồn</h3>
                <span class="excel-table-count">${rows.length} bản ghi</span>
                ${timestampHtml}
            </div>
            <div class="excel-table-body" style="max-height: 700px; overflow: auto;">
                <table class="excel-table" style="border-collapse: separate; border-spacing: 0;">
                    <thead>
                        <tr class="header-row">${headerHtml}</tr>
                    </thead>
                    <tbody>${bodyHtml}</tbody>
                </table>
            </div>
        </div>
    `;
}

function baseHeaderStyle() {
    return 'position: sticky; top: 0; background: #1a5089; color: #fff; border: 1px solid #d7e3f1; padding: 10px 8px; white-space: nowrap; text-align: center; z-index: 20;';
}

function baseCellStyle(alignment) {
    return `border: 1px solid #e3eaf3; padding: 8px 10px; white-space: nowrap; text-align: ${alignment}; background: #fff;`;
}

function getStickyStyle(colIndex, isHeader, rowIndex = 0) {
    const widths = [180, 250, 140];
    if (colIndex >= widths.length) {
        return '';
    }

    const left = widths.slice(0, colIndex).reduce((sum, width) => sum + width, 0);
    const zIndex = isHeader ? 30 + (widths.length - colIndex) : 10 + (widths.length - colIndex);
    const background = isHeader ? '#1a5089' : (rowIndex % 2 === 0 ? '#ffffff' : '#f8fbff');
    const shadow = colIndex === widths.length - 1 ? 'box-shadow: 2px 0 6px rgba(0,0,0,0.08);' : '';

    return `position: sticky; left: ${left}px; min-width: ${widths[colIndex]}px; max-width: ${widths[colIndex]}px; background: ${background}; z-index: ${zIndex}; ${shadow}`;
}

function isNumericValue(value) {
    return typeof value === 'number' && !Number.isNaN(value);
}

function formatCellValue(value) {
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

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function renderNvktCharts(rows, container) {
    if (!container) return;

    const nvktRows = (rows || []).filter(row => {
        const name = String(row['nvkt_hoac_ten_nv'] || '').trim();
        return name && name !== 'TỔNG CỘNG';
    });

    if (nvktRows.length === 0) {
        container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu biểu đồ</p></div>';
        return;
    }

    container.innerHTML = `<div class="kpi-bchn-chart-grid">${nvktRows.map((row, index) => `
        <div class="kpi-bchn-chart-card">
            <div class="kpi-bchn-chart-head">
                <div class="kpi-bchn-chart-title">${escapeHtml(row['nvkt_hoac_ten_nv'] || '')}</div>
                <div class="kpi-bchn-chart-subtitle">${escapeHtml(row['to_doi_hoac_don_vi'] || '')}${row['ttvt'] ? ` | ${escapeHtml(row['ttvt'])}` : ''}</div>
            </div>
            <div class="kpi-bchn-chart-canvas">
                <canvas id="kpi-bchn-chart-${index}"></canvas>
            </div>
        </div>
    `).join('')}</div>`;

    if (typeof Chart === 'undefined') {
        return;
    }

    nvktRows.forEach((row, index) => {
        const canvas = document.getElementById(`kpi-bchn-chart-${index}`);
        if (!canvas) return;

        const rateMetrics = [
            parsePercentLike(row['c14_ty_le_hl']),
            parsePercentLike(row['i15_k1_ty_le_shc']),
            parsePercentLike(row['i15_k2_ty_le_shc']),
            parsePercentLike(row['ghtt_ty_le_tong']),
            parsePercentLike(row['kpi_c11_chi_tieu_bsc']),
            parsePercentLike(row['kpi_c12_chi_tieu_bsc']),
        ];

        const volumeMetrics = [
            parsePercentLike(row['c11_tong_phieu']),
            parsePercentLike(row['c12_sm1_so_phieu_hll']),
            parsePercentLike(row['kqtt_tong']),
        ];

        new Chart(canvas.getContext('2d'), {
            type: 'bar',
            data: {
                labels: ['C11', 'C12 HLL', 'KQTT', 'C14', 'I1.5 K1', 'I1.5 K2', 'GHTT', 'KPI C11', 'KPI C12'],
                datasets: [
                    {
                        type: 'bar',
                        label: 'Sản lượng',
                        data: [
                            volumeMetrics[0],
                            volumeMetrics[1],
                            volumeMetrics[2],
                            null,
                            null,
                            null,
                            null,
                            null,
                            null,
                        ],
                        backgroundColor: ['#1f6aa5', '#4d8fc6', '#80b3d8', 'transparent', 'transparent', 'transparent', 'transparent', 'transparent', 'transparent'],
                        borderRadius: 6,
                        yAxisID: 'y',
                    },
                    {
                        type: 'line',
                        label: 'Tỷ lệ / điểm',
                        data: [
                            null,
                            null,
                            null,
                            rateMetrics[0],
                            rateMetrics[1],
                            rateMetrics[2],
                            rateMetrics[3],
                            rateMetrics[4],
                            rateMetrics[5],
                        ],
                        borderColor: '#d65a31',
                        backgroundColor: 'rgba(214, 90, 49, 0.18)',
                        pointBackgroundColor: '#d65a31',
                        pointRadius: 4,
                        pointHoverRadius: 5,
                        tension: 0.28,
                        spanGaps: true,
                        yAxisID: 'y1',
                    },
                ],
            },
            options: {
                maintainAspectRatio: false,
                responsive: true,
                interaction: {
                    mode: 'index',
                    intersect: false,
                },
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 12,
                            usePointStyle: true,
                        },
                    },
                    tooltip: {
                        callbacks: {
                            label(context) {
                                const value = context.raw;
                                if (value === null || value === undefined || Number.isNaN(value)) {
                                    return `${context.dataset.label}: `;
                                }
                                return `${context.dataset.label}: ${formatChartNumber(value)}`;
                            },
                        },
                    },
                },
                scales: {
                    x: {
                        ticks: {
                            maxRotation: 0,
                            autoSkip: false,
                            font: {
                                size: 10,
                            },
                        },
                        grid: {
                            display: false,
                        },
                    },
                    y: {
                        position: 'left',
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Sản lượng',
                        },
                    },
                    y1: {
                        position: 'right',
                        beginAtZero: true,
                        grid: {
                            drawOnChartArea: false,
                        },
                        title: {
                            display: true,
                            text: 'Tỷ lệ / điểm',
                        },
                    },
                },
            },
        });
    });
}

function parsePercentLike(value) {
    if (value === null || value === undefined || value === '') {
        return null;
    }

    if (typeof value === 'number') {
        return Number.isNaN(value) ? null : value;
    }

    const normalized = String(value).replace('%', '').replace(',', '.').trim();
    if (!normalized) return null;
    const numeric = Number(normalized);
    return Number.isNaN(numeric) ? null : numeric;
}

function formatChartNumber(value) {
    return Number.isInteger(value)
        ? value.toLocaleString('vi-VN')
        : value.toLocaleString('vi-VN', { maximumFractionDigits: 2 });
}
