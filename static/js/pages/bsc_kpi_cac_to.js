document.addEventListener('DOMContentLoaded', async function () {
    await initBscKpiCacToPage();
});

async function initBscKpiCacToPage() {
    const container = document.getElementById('bsc-kpi-cac-to-table');
    if (!container) return;

    try {
        const data = await API.fetchData('/api/bsc-kpi-cac-to-data');
        if (data.error) {
            throw new Error(data.error);
        }

        const tsEl = document.getElementById('bsc-kpi-cac-to-timestamp');
        if (tsEl && data.file_modified) {
            tsEl.textContent = `(Cập nhật: ${data.file_modified})`;
        }

        if (!data.sheet || !Array.isArray(data.sheet.columns) || !Array.isArray(data.sheet.data) || data.sheet.data.length === 0) {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
            return;
        }

        container.innerHTML = renderBscKpiCacToTable(data.sheet.columns, data.sheet.data, data.file_modified);
    } catch (error) {
        console.error('Error loading bsc-kpi-cac-to data:', error);
        container.innerHTML = `
            <div class="loading">
                <i class="fas fa-exclamation-triangle"></i>
                <br>Lỗi khi tải dữ liệu: ${error.message}
            </div>`;
    }
}

function renderBscKpiCacToTable(columns, rows, fileModified) {
    const timestampHtml = fileModified
        ? `<div class="excel-table-timestamp">Cập nhật: ${fileModified}</div>`
        : '';

    const headerHtml = columns.map((column, index) => {
        const stickyStyle = bscKpiCacToStickyStyle(index, true);
        return `<th style="${bscKpiCacToHeaderStyle()}${stickyStyle}">${bscKpiCacToEscapeHtml(column)}</th>`;
    }).join('');

    const bodyHtml = rows.map((row, rowIndex) => {
        const rowHtml = columns.map((column, colIndex) => {
            const rawValue = row[column];
            const displayValue = bscKpiCacToFormatCell(rawValue);
            const align = typeof rawValue === 'number' && !Number.isNaN(rawValue) ? 'right' : 'left';
            const stickyStyle = bscKpiCacToStickyStyle(colIndex, false, rowIndex);
            return `<td style="${bscKpiCacToCellStyle(align)}${stickyStyle}">${bscKpiCacToEscapeHtml(displayValue)}</td>`;
        }).join('');
        return `<tr>${rowHtml}</tr>`;
    }).join('');

    return `
        <div class="excel-table-card">
            <div class="excel-table-header">
                <h3 class="excel-table-title">v_chi_tieu_bsc_kpi_cac_to</h3>
                <span class="excel-table-count">${rows.length} bản ghi</span>
                ${timestampHtml}
            </div>
            <div class="excel-table-body" style="max-height: 760px; overflow: auto;">
                <table class="excel-table" style="border-collapse: separate; border-spacing: 0;">
                    <thead><tr class="header-row">${headerHtml}</tr></thead>
                    <tbody>${bodyHtml}</tbody>
                </table>
            </div>
        </div>
    `;
}

function bscKpiCacToHeaderStyle() {
    return 'position: sticky; top: 0; background: #1a5089; color: #fff; border: 1px solid #d7e3f1; padding: 10px 8px; white-space: nowrap; text-align: center; z-index: 20;';
}

function bscKpiCacToCellStyle(align) {
    return `border: 1px solid #e3eaf3; padding: 8px 10px; white-space: nowrap; text-align: ${align}; background: #fff;`;
}

function bscKpiCacToStickyStyle(colIndex, isHeader, rowIndex = 0) {
    const widths = [240];
    if (colIndex >= widths.length) return '';
    const left = 0;
    const zIndex = isHeader ? 30 : 10;
    const background = isHeader ? '#1a5089' : (rowIndex % 2 === 0 ? '#ffffff' : '#f8fbff');
    return `position: sticky; left: ${left}px; min-width: ${widths[colIndex]}px; max-width: ${widths[colIndex]}px; background: ${background}; z-index: ${zIndex}; box-shadow: 2px 0 6px rgba(0,0,0,0.08);`;
}

function bscKpiCacToFormatCell(value) {
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

function bscKpiCacToEscapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
