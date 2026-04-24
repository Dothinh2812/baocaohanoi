document.addEventListener('DOMContentLoaded', async function() {
    bindDateFilter();
    await initGiaHanPage();
});

function getRequestedDate() {
    const params = new URLSearchParams(window.location.search);
    return params.get('date') || '';
}

function bindDateFilter() {
    const input = document.getElementById('giahan-date-input');
    const applyButton = document.getElementById('giahan-date-apply');
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
    input.addEventListener('keydown', function(event) {
        if (event.key === 'Enter') {
            applyDateFilter();
        }
    });
}

function syncDateFilterState(data) {
    const input = document.getElementById('giahan-date-input');
    if (input) {
        input.value = data.selected_date || getRequestedDate();
    }

    const meta = document.getElementById('giahan-date-meta');
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

function buildEndpoint(path) {
    const requestedDate = getRequestedDate();
    return requestedDate ? `${path}?date=${encodeURIComponent(requestedDate)}` : path;
}

async function fetchGiaHanData(path) {
    const response = await fetch(buildEndpoint(path));
    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    if (data.error) {
        throw new Error(data.error);
    }
    syncDateFilterState(data);
    return data;
}

async function initGiaHanPage() {
    await Promise.all([
        loadGHTTHNIData(),
        loadGHTTSTYData(),
        loadGHTTNVKTDBData()
    ]);
}

async function loadGHTTHNIData() {
    const container = document.getElementById('ghtt-hni-container');

    try {
        const data = await fetchGiaHanData('/api/giahan-ghtt-hni');
        renderGHTTHNITable(data, container, 'KQ GHTT HNI', 'ghtt-hni-table');
    } catch (error) {
        console.error('Error loading GHTT HNI data:', error);
        container.innerHTML = `
            <div class="loading">
                <i class="fas fa-exclamation-triangle"></i>
                <br>Lỗi khi tải dữ liệu: ${error.message}
            </div>
        `;
    }
}

function renderGHTTHNITable(data, container, title = 'KQ GHTT HNI', tableId = 'ghtt-hni-table', downloadFile = '', redTextColumnIndices = []) {
    const { header1, header2, merges, data: rows, file_info, selected_date } = data;
    const selectedDateLabel = selected_date ? ` ngày ${selected_date}` : '';

    if (!rows || rows.length === 0) {
        container.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu ${title}${selectedDateLabel}</p></div>`;
        return;
    }

    const mergeLookup = {};
    merges.forEach(m => {
        const colspan = m.max_col - m.min_col + 1;
        const rowspan = m.max_row - m.min_row + 1;
        mergeLookup[`${m.min_row}_${m.min_col}`] = { colspan, rowspan };
        for (let r = m.min_row; r <= m.max_row; r++) {
            for (let c = m.min_col; c <= m.max_col; c++) {
                if (r !== m.min_row || c !== m.min_col) {
                    mergeLookup[`${r}_${c}`] = { hidden: true };
                }
            }
        }
    });

    let timestampHtml = '';
    if (file_info && file_info.modified) {
        timestampHtml = `<div class="excel-table-timestamp">Cập nhật: ${file_info.modified}</div>`;
    }

    let html = `
        <div class="excel-table-card">
            <div class="excel-table-header">
                <h3 class="excel-table-title">${title}</h3>
                <span class="excel-table-count">${rows.length} bản ghi</span>
                ${timestampHtml}
                ${downloadFile ? `<a href="/api/giahan-ghtt-download/${downloadFile}" class="download-btn" style="padding: 6px 14px; font-size: 13px; border-radius: 6px;"><i class="fas fa-download"></i> Tải file gốc</a>` : ''}
            </div>
            <div class="excel-table-body" style="overflow: auto;">
                <table class="excel-table" id="${tableId}">
                    <thead>
                        <tr class="header-row">`;

    for (let c = 1; c <= header1.length; c++) {
        const key = `1_${c}`;
        const m = mergeLookup[key];
        if (m && m.hidden) continue;
        const colspan = (m && m.colspan) ? ` colspan="${m.colspan}"` : '';
        const rowspan = (m && m.rowspan) ? ` rowspan="${m.rowspan}"` : '';
        const val = header1[c - 1] || '';
        html += `<th${colspan}${rowspan}>${val}</th>`;
    }

    html += `</tr><tr class="header-row">`;

    for (let c = 1; c <= header2.length; c++) {
        const key = `2_${c}`;
        const m = mergeLookup[key];
        if (m && m.hidden) continue;
        const colspan = (m && m.colspan) ? ` colspan="${m.colspan}"` : '';
        const rowspan = (m && m.rowspan) ? ` rowspan="${m.rowspan}"` : '';
        const val = header2[c - 1] || '';
        html += `<th${colspan}${rowspan}>${val}</th>`;
    }

    html += `</tr></thead><tbody>`;

    rows.forEach(row => {
        const isSonTay = tableId === 'ghtt-hni-table' && row[0] && row[0].includes('Sơn Tây');
        const rowStyle = isSonTay ? ' style="background-color: #ffcccc; font-weight: bold; color: #cc0000;"' : '';
        html += `<tr${rowStyle}>`;
        row.forEach((val, idx) => {
            const isRedCol = redTextColumnIndices.includes(idx);
            const redStyle = isRedCol ? ' color: #d32f2f; font-weight: 600;' : '';
            if (idx === 0) {
                html += `<td style="font-weight: bold; white-space: nowrap;${isSonTay ? ' color: #cc0000;' : ''}">${val}</td>`;
            } else {
                html += `<td style="text-align: center;${redStyle}">${val}</td>`;
            }
        });
        html += '</tr>';
    });

    html += '</tbody></table></div></div>';
    container.innerHTML = html;
}

async function loadGHTTSTYData() {
    const container = document.getElementById('ghtt-sty-container');

    try {
        const data = await fetchGiaHanData('/api/giahan-ghtt-sty');
        renderGHTTHNITable(data, container, 'KQ GHTT TTVT STY', 'ghtt-sty-table', 'tong_hop_ghtt_sontay.xlsx', [3, 7, 10]);
    } catch (error) {
        console.error('Error loading GHTT STY data:', error);
        container.innerHTML = `
            <div class="loading">
                <i class="fas fa-exclamation-triangle"></i>
                <br>Lỗi khi tải dữ liệu: ${error.message}
            </div>
        `;
    }
}

async function loadGHTTNVKTDBData() {
    const tabsContainer = document.getElementById('ghtt-nvktdb-tabs');
    const tablesContainer = document.getElementById('ghtt-nvktdb-tables-container');

    try {
        const data = await fetchGiaHanData('/api/giahan-ghtt-nvktdb');

        if (data && data.sheets && Object.keys(data.sheets).length > 0) {
            const unitNames = Object.keys(data.sheets);
            loadGHTTNVKTDBTabs(unitNames, data.sheets, data.file_info, tabsContainer, tablesContainer);
        } else {
            const selectedDateLabel = data.selected_date ? ` ngày ${data.selected_date}` : '';
            tabsContainer.innerHTML = '';
            tablesContainer.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu NVKTĐB${selectedDateLabel}</p></div>`;
        }
    } catch (error) {
        console.error('Error loading GHTT NVKTDB data:', error);
        tablesContainer.innerHTML = `
            <div class="loading">
                <i class="fas fa-exclamation-triangle"></i>
                <br>Lỗi khi tải dữ liệu: ${error.message}
            </div>
        `;
    }
}

function loadGHTTNVKTDBTabs(unitNames, unitsData, fileInfo, tabsContainer, tablesContainer) {
    tabsContainer.innerHTML = '';
    tablesContainer.innerHTML = '';

    unitNames.forEach((unitName, index) => {
        const li = document.createElement('li');
        li.className = `excel-tab ${index === 0 ? 'active' : ''}`;
        li.textContent = unitName;
        li.dataset.unit = unitName;

        li.addEventListener('click', function() {
            document.querySelectorAll('#ghtt-nvktdb-tabs .excel-tab').forEach(tab => {
                tab.classList.remove('active');
            });
            this.classList.add('active');

            document.querySelectorAll('#ghtt-nvktdb-tables-container > div').forEach(table => {
                table.style.display = 'none';
            });
            document.getElementById(`ghtt-nvktdb-table-${unitName}`).style.display = 'block';
        });

        tabsContainer.appendChild(li);
    });

    unitNames.forEach((unitName, index) => {
        const tableDiv = document.createElement('div');
        tableDiv.id = `ghtt-nvktdb-table-${unitName}`;
        tableDiv.style.display = index === 0 ? 'block' : 'none';

        const tableHtml = createExcelTable(unitsData[unitName], unitName, {
            showRowNumbers: true,
            tableClass: 'excel-table',
            maxHeight: 'none',
            fileInfo: fileInfo,
            redTextColumns: ['Tỷ lệ T', 'Tỷ lệ T+1', 'Tỷ lệ chạm quá 72H Tháng T']
        });
        tableDiv.innerHTML = tableHtml;

        tablesContainer.appendChild(tableDiv);
    });
}
