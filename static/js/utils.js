/* ========================================
   UTILITY FUNCTIONS
   - Table creation
   - Data formatting
   - Helper functions
   ======================================== */

/**
 * Format date to DD/MM/YYYY
 */
function formatDate(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

/**
 * Format number with thousand separators
 */
function formatNumber(number) {
    if (number === null || number === undefined || number === '') return '';
    return Number(number).toLocaleString('vi-VN');
}

/**
 * Show loading state
 */
function showLoading(containerId = null) {
    const loadingHTML = `
        <div class="loading">
            <i class="fas fa-spinner"></i>
            <p>Đang tải dữ liệu...</p>
        </div>
    `;

    if (containerId) {
        const container = document.getElementById(containerId);
        if (container) {
            container.innerHTML = loadingHTML;
        }
    }
}

/**
 * Hide loading state
 */
function hideLoading(containerId = null) {
    if (containerId) {
        const container = document.getElementById(containerId);
        if (container) {
            const loading = container.querySelector('.loading');
            if (loading) {
                loading.remove();
            }
        }
    }
}

/**
 * Show error message
 */
function showError(message, containerId = null) {
    const errorHTML = `
        <div class="empty-table">
            <i class="fas fa-exclamation-triangle"></i>
            <p>${message}</p>
        </div>
    `;

    if (containerId) {
        const container = document.getElementById(containerId);
        if (container) {
            container.innerHTML = errorHTML;
        }
    } else {
        console.error(message);
    }
}

/**
 * Show empty state
 */
function showEmptyState(message = 'Không có dữ liệu', containerId = null) {
    const emptyHTML = `
        <div class="empty-table">
            <i class="fas fa-inbox"></i>
            <p>${message}</p>
        </div>
    `;

    if (containerId) {
        const container = document.getElementById(containerId);
        if (container) {
            container.innerHTML = emptyHTML;
        }
    }
}

/**
 * Create Excel table from data
 */
function createExcelTable(data, sheetName, options = {}) {
    if (!data || !data.columns || !data.data || data.data.length === 0) {
        return '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
    }

    const {
        showRowNumbers = true,
        tableClass = 'excel-table',
        maxHeight = '600px',
        fileInfo = null,
        enableFilter = true,
        redTextColumns = [],
        frozenColumns = 0,
        columnLabels = {}
    } = options;

    const hasVerticalScroll = maxHeight && maxHeight !== 'none';
    const freezeCount = Math.max(0, parseInt(frozenColumns, 10) || 0);
    const tableId = 'table-' + Math.random().toString(36).substr(2, 9);
    const freezeGroupId = freezeCount > 0 ? 'freeze-' + Math.random().toString(36).substr(2, 9) : '';
    const bodyClasses = ['excel-table-body'];

    if (freezeCount > 0) {
        bodyClasses.push('has-frozen-columns');
    }

    const tableBodyStyle = [
        maxHeight && maxHeight !== 'none' ? `max-height: ${maxHeight};` : '',
        `overflow-y: ${hasVerticalScroll ? 'auto' : 'visible'};`,
        freezeCount > 0 ? 'overflow-x: visible;' : ''
    ].filter(Boolean).join(' ');

    let timestampHtml = '';
    if (fileInfo && fileInfo.modified) {
        timestampHtml = `<div class="excel-table-timestamp">Cập nhật: ${fileInfo.modified}</div>`;
    }

    const columnDefs = [];
    if (showRowNumbers) {
        columnDefs.push({
            key: '__row_number__',
            label: 'STT',
            originalIndex: -1,
            isRowNumber: true
        });
    }

    data.columns.forEach((col, idx) => {
        columnDefs.push({
            key: col,
            dataKey: col,
            label: columnLabels[col] || col,
            originalIndex: idx,
            isRowNumber: false
        });
    });

    const uniqueValues = {};
    if (enableFilter) {
        data.columns.forEach((col, idx) => {
            const values = new Set();
            data.data.forEach(row => {
                const val = row[col] !== null && row[col] !== undefined ? String(row[col]).trim() : '';
                if (val) values.add(val);
            });
            uniqueValues[idx] = Array.from(values).sort((a, b) => a.localeCompare(b, 'vi'));
        });
    }

    const gioConLaiColIdx = data.columns.findIndex(col => col.toLowerCase() === 'giờ còn lại thực');
    const tgThiCongColIdx = data.columns.findIndex(col => col.toLowerCase() === 'tg_thicong_h');

    function getRowStyle(row) {
        let rowStyle = '';

        if (gioConLaiColIdx !== -1) {
            const gioConLai = parseFloat(row[data.columns[gioConLaiColIdx]]);
            if (!isNaN(gioConLai)) {
                if (gioConLai < 0) {
                    rowStyle = 'color: #d32f2f; font-weight: 600;';
                } else if (gioConLai < 1) {
                    rowStyle = 'background-color: #ffcdd2; color: #b71c1c;';
                } else if (gioConLai < 2) {
                    rowStyle = 'background-color: #ffe0b2; color: #e65100;';
                } else if (gioConLai < 3) {
                    rowStyle = 'background-color: #fff9c4; color: #f57f17;';
                }
            }
        }

        if (tgThiCongColIdx !== -1) {
            const tgThiCongRaw = row[data.columns[tgThiCongColIdx]];
            const tgThiCong = parseFloat(String(tgThiCongRaw).replace(',', '.'));
            if (!isNaN(tgThiCong)) {
                if (tgThiCong >= 24) {
                    rowStyle = 'color: #d32f2f; font-weight: 600;';
                } else if (tgThiCong > 16 && tgThiCong < 24) {
                    rowStyle = 'background-color: #ffe0b2;';
                }
            }
        }

        return rowStyle;
    }

    function escapeHtml(value) {
        return String(value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function classifyHeaderLabel(label) {
        const normalized = String(label || '').replace(/\s+/g, ' ').trim();
        const length = normalized.length;
        const wordCount = normalized ? normalized.split(' ').length : 0;

        if (length >= 90 || wordCount >= 12) {
            return 'is-very-long';
        }

        if (length >= 48 || wordCount >= 7) {
            return 'is-long';
        }

        return '';
    }

    function formatHeaderLabel(label) {
        const escaped = escapeHtml(label).replace(/\s+/g, ' ').trim();
        return escaped
            .replace(/,\s+/g, ',<wbr> ')
            .replace(/;\s+/g, ';<wbr> ')
            .replace(/:\s+/g, ':<wbr> ')
            .replace(/\)\s+/g, ')<wbr> ')
            .replace(/\s+\(/g, ' <wbr>(')
            .replace(/\s*\/\s*/g, ' /<wbr> ')
            .replace(/\s*-\s*/g, ' - <wbr>');
    }

    function buildHeaderCell(def) {
        if (def.isRowNumber) {
            return '<th style="width: 50px; text-align: center;">STT</th>';
        }

        const headerClass = classifyHeaderLabel(def.label);
        const classAttr = headerClass ? ` class="excel-header-wrap ${headerClass}"` : '';
        const titleAttr = escapeHtml(def.label);

        return `<th${classAttr} data-column="${def.dataKey}" data-column-idx="${def.originalIndex}" title="${titleAttr}"><span class="excel-header-text">${formatHeaderLabel(def.label)}</span></th>`;
    }

    function buildFilterCell(def) {
        if (def.isRowNumber) {
            return '<th style="padding: 4px;"></th>';
        }

        const col = def.dataKey;
        const idx = def.originalIndex;
        const hasNumeric = data.data.some(row => {
            const v = row[col];
            return v !== null && v !== undefined && v !== '' && !isNaN(Number(v));
        });

        let dropdownItemsHtml = '<div class="filter-dropdown-item" data-value="">-- Tất cả --</div>';
        if (hasNumeric) {
            dropdownItemsHtml += '<div class="filter-dropdown-item filter-operator" data-value="!=0">&#8800; 0 (Khác 0)</div>';
            dropdownItemsHtml += '<div class="filter-dropdown-item filter-operator" data-value=">0">&gt; 0 (Lớn hơn 0)</div>';
        }

        (uniqueValues[idx] || []).forEach(val => {
            const escaped = val.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            dropdownItemsHtml += `<div class="filter-dropdown-item" data-value="${escaped}">${escaped}</div>`;
        });

        return `<th style="padding: 4px;">
            <div class="filter-wrapper">
                <input type="text" class="column-filter" data-column-idx="${idx}" placeholder="Lọc..." autocomplete="off" />
                <span class="filter-dropdown-toggle" data-column-idx="${idx}"><i class="fas fa-caret-down"></i></span>
                <div class="filter-dropdown-list" data-column-idx="${idx}">
                    ${dropdownItemsHtml}
                </div>
            </div>
        </th>`;
    }

    function buildDataCell(def, row, rowIndex) {
        if (def.isRowNumber) {
            return `<td data-column-idx="-1" style="text-align: center; font-weight: 600; color: #6c757d;">${rowIndex + 1}</td>`;
        }

        const col = def.dataKey;
        let cellValue = row[col] !== null && row[col] !== undefined ? row[col] : '';
        let cellClass = '';
        let cellStyle = '';

        if (col.includes('Trạng thái') && cellValue.toString().toLowerCase().includes('down')) {
            cellClass = 'status-down';
        } else if (col.includes('Trạng thái') && cellValue.toString().toLowerCase().includes('up')) {
            cellClass = 'status-up';
        }

        if (redTextColumns.length > 0 && redTextColumns.some(rc => col.trim() === rc.trim())) {
            cellStyle = 'color: #d32f2f; font-weight: 600;';
        }

        const normalizedCol = col.trim().toUpperCase();
        if (normalizedCol === 'LYDOTON' || normalizedCol === 'GHICHU_TON') {
            cellStyle = `${cellStyle} background-color: #ffd54f; font-weight: 600;`;
        }

        if (normalizedCol === 'TG_THICONG_H') {
            cellStyle = `${cellStyle} background-color: #fff8cc; font-weight: 600;`;
        }

        return `<td data-column="${col}" data-column-idx="${def.originalIndex}" class="${cellClass}" style="${cellStyle}">${cellValue}</td>`;
    }

    function buildTable(defs, id, extraClasses = '') {
        let tableHtml = `<table class="${tableClass} ${extraClasses} filterable-table" id="${id}" data-show-row-numbers="${showRowNumbers}"`;
        if (freezeGroupId) {
            tableHtml += ` data-freeze-group="${freezeGroupId}"`;
        }
        tableHtml += '>';
        tableHtml += '<thead><tr class="header-row">';
        defs.forEach(def => {
            tableHtml += buildHeaderCell(def);
        });
        tableHtml += '</tr>';

        if (enableFilter) {
            tableHtml += '<tr class="filter-row">';
            defs.forEach(def => {
                tableHtml += buildFilterCell(def);
            });
            tableHtml += '</tr>';
        }

        tableHtml += '</thead><tbody>';
        data.data.forEach((row, index) => {
            tableHtml += `<tr class="data-row" data-row-index="${index}" style="${getRowStyle(row)}">`;
            defs.forEach(def => {
                tableHtml += buildDataCell(def, row, index);
            });
            tableHtml += '</tr>';
        });
        tableHtml += '</tbody></table>';
        return tableHtml;
    }

    const canFreeze = freezeCount > 0 && freezeCount < columnDefs.length;
    let tableContentHtml = '';

    if (canFreeze) {
        const frozenDefs = columnDefs.slice(0, freezeCount);
        const scrollDefs = columnDefs.slice(freezeCount);

        tableContentHtml = `
            <div class="excel-freeze-layout" data-freeze-group="${freezeGroupId}">
                <div class="excel-frozen-pane">
                    ${buildTable(frozenDefs, `${tableId}-frozen`, 'excel-table-frozen')}
                </div>
                <div class="excel-scroll-pane">
                    ${buildTable(scrollDefs, `${tableId}-scroll`, 'excel-table-scroll')}
                </div>
            </div>
        `;
    } else {
        tableContentHtml = buildTable(columnDefs, tableId);
    }

    return `
        <div class="excel-table-card">
            <div class="excel-table-header">
                <h3 class="excel-table-title">${sheetName}</h3>
                <span class="excel-table-count">${data.data.length} bản ghi</span>
                ${timestampHtml}
            </div>
            <div class="${bodyClasses.join(' ')}" style="${tableBodyStyle}">
                ${tableContentHtml}
            </div>
        </div>
    `;
}

function syncFrozenTableLayout(layout) {
    if (!layout) return;

    const frozenTable = layout.querySelector('.excel-table-frozen');
    const scrollTable = layout.querySelector('.excel-table-scroll');
    if (!frozenTable || !scrollTable) return;

    const frozenRows = Array.from(layout.querySelectorAll('.excel-frozen-pane thead tr, .excel-frozen-pane tbody tr.data-row'));
    const scrollRows = Array.from(layout.querySelectorAll('.excel-scroll-pane thead tr, .excel-scroll-pane tbody tr.data-row'));
    const rowCount = Math.min(frozenRows.length, scrollRows.length);

    frozenRows.forEach(row => {
        row.style.height = '';
    });
    scrollRows.forEach(row => {
        row.style.height = '';
    });

    for (let index = 0; index < rowCount; index += 1) {
        const frozenRow = frozenRows[index];
        const scrollRow = scrollRows[index];
        const height = Math.max(frozenRow.getBoundingClientRect().height, scrollRow.getBoundingClientRect().height);

        if (height > 0) {
            const normalizedHeight = `${Math.ceil(height)}px`;
            frozenRow.style.height = normalizedHeight;
            scrollRow.style.height = normalizedHeight;
        }
    }
}

function initializeFrozenTableLayouts(scope = document) {
    if (!scope) return;

    const layouts = scope.classList && scope.classList.contains('excel-freeze-layout')
        ? [scope]
        : Array.from(scope.querySelectorAll('.excel-freeze-layout'));

    if (layouts.length === 0) return;

    requestAnimationFrame(() => {
        layouts.forEach(syncFrozenTableLayout);
    });
}

window.addEventListener('resize', function () {
    syncStickyHeaderOffsets(document);
    initializeFrozenTableLayouts(document);
});

function syncStickyHeaderOffsets(scope = document) {
    if (!scope) return;

    const tables = scope.classList && scope.classList.contains('excel-table')
        ? [scope]
        : Array.from(scope.querySelectorAll('.excel-table'));

    tables.forEach(table => {
        const headerRow = table.querySelector('thead tr.header-row');
        const filterRow = table.querySelector('thead tr.filter-row');
        if (!headerRow || !filterRow) return;

        const headerHeight = Math.ceil(headerRow.getBoundingClientRect().height);
        if (headerHeight > 0) {
            table.style.setProperty('--excel-filter-row-top', `${headerHeight}px`);
        }
    });
}

function scheduleStickyHeaderSync(scope = document) {
    requestAnimationFrame(() => {
        syncStickyHeaderOffsets(scope);
        initializeFrozenTableLayouts(scope);
    });
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () {
            scheduleStickyHeaderSync(document);
        });
    } else {
        scheduleStickyHeaderSync(document);
    }

    if (typeof MutationObserver !== 'undefined') {
        const tableObserver = new MutationObserver(mutations => {
            for (const mutation of mutations) {
                for (const node of mutation.addedNodes) {
                    if (!(node instanceof Element)) continue;
                    if (node.matches('.excel-table-card, .excel-table, .excel-freeze-layout') || node.querySelector('.excel-table')) {
                        scheduleStickyHeaderSync(node);
                    }
                }
            }
        });

        if (document.body) {
            tableObserver.observe(document.body, { childList: true, subtree: true });
        }
    }
}

// Parse filter expression into {type, value} object
function parseFilterExpression(filterValue) {
    // Match operators: !=, <>, >=, <=, >, <, =
    const match = filterValue.match(/^(!=|<>|>=|<=|>|<|=)\s*(.+)$/);
    if (match) {
        const op = match[1] === '<>' ? '!=' : match[1];
        const val = parseFloat(match[2]);
        if (!isNaN(val)) {
            return { type: 'operator', op: op, value: val };
        }
    }
    return { type: 'text', value: filterValue.toLowerCase() };
}

// Check if a cell value matches a filter expression
function matchesFilter(cellText, filter) {
    if (filter.type === 'text') {
        return cellText.toLowerCase().includes(filter.value);
    }

    // Operator filter - parse cell as number
    const cellNum = parseFloat(cellText.replace(/\./g, '').replace(/,/g, '.'));
    const isEmpty = cellText.trim() === '';

    switch (filter.op) {
        case '!=':
            // Exclude empty cells and cells equal to the value
            if (isEmpty) return false;
            return isNaN(cellNum) ? true : cellNum !== filter.value;
        case '>':
            if (isEmpty || isNaN(cellNum)) return false;
            return cellNum > filter.value;
        case '<':
            if (isEmpty || isNaN(cellNum)) return false;
            return cellNum < filter.value;
        case '>=':
            if (isEmpty || isNaN(cellNum)) return false;
            return cellNum >= filter.value;
        case '<=':
            if (isEmpty || isNaN(cellNum)) return false;
            return cellNum <= filter.value;
        case '=':
            if (isEmpty || isNaN(cellNum)) return false;
            return cellNum === filter.value;
        default:
            return cellText.toLowerCase().includes(filter.value.toString());
    }
}

// Apply column filters on a table
function applyColumnFilters(table) {
    const freezeGroup = table.dataset.freezeGroup;
    const card = table.closest('.excel-table-card');

    if (freezeGroup) {
        const groupedTables = Array.from(document.querySelectorAll(`.filterable-table[data-freeze-group="${freezeGroup}"]`));
        const filterInputs = groupedTables.flatMap(groupedTable => Array.from(groupedTable.querySelectorAll('.column-filter')));
        const filters = {};

        filterInputs.forEach(inp => {
            const colIdx = parseInt(inp.dataset.columnIdx, 10);
            const value = inp.value.trim();
            if (value) {
                filters[colIdx] = parseFilterExpression(value);
            }
        });

        const rowGroups = new Map();
        groupedTables.forEach(groupedTable => {
            groupedTable.querySelectorAll('tbody tr.data-row').forEach(row => {
                const rowIndex = row.dataset.rowIndex;
                if (!rowGroups.has(rowIndex)) {
                    rowGroups.set(rowIndex, []);
                }
                rowGroups.get(rowIndex).push(row);
            });
        });

        let visibleCount = 0;
        rowGroups.forEach(rows => {
            let show = true;

            for (const [colIdx, filter] of Object.entries(filters)) {
                let cellText = '';
                for (const row of rows) {
                    const cell = row.querySelector(`td[data-column-idx="${colIdx}"]`);
                    if (cell) {
                        cellText = cell.textContent;
                        break;
                    }
                }

                if (!matchesFilter(cellText, filter)) {
                    show = false;
                    break;
                }
            }

            rows.forEach(row => {
                row.style.display = show ? '' : 'none';
            });

            if (show) {
                visibleCount += 1;
            }
        });

        if (card) {
            const countEl = card.querySelector('.excel-table-count');
            if (countEl) {
                const total = rowGroups.size;
                countEl.textContent = visibleCount < total ? `${visibleCount}/${total} bản ghi` : `${total} bản ghi`;
            }
        }

        const layout = card ? card.querySelector('.excel-freeze-layout') : null;
        if (layout) {
            initializeFrozenTableLayouts(layout);
        }
        return;
    }

    const filterInputs = table.querySelectorAll('.column-filter');
    const tbody = table.querySelector('tbody');
    const rows = tbody.querySelectorAll('tr.data-row');
    const showRowNumbers = table.dataset.showRowNumbers === 'true';
    const filters = {};

    filterInputs.forEach(inp => {
        const colIdx = parseInt(inp.dataset.columnIdx, 10);
        const value = inp.value.trim();
        if (value) {
            filters[colIdx] = parseFilterExpression(value);
        }
    });

    let visibleCount = 0;
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        let show = true;

        for (const [colIdx, filter] of Object.entries(filters)) {
            const cellIdx = showRowNumbers ? parseInt(colIdx, 10) + 1 : parseInt(colIdx, 10);
            const cell = cells[cellIdx];
            if (cell) {
                const cellText = cell.textContent;
                if (!matchesFilter(cellText, filter)) {
                    show = false;
                    break;
                }
            }
        }

        row.style.display = show ? '' : 'none';
        if (show) visibleCount++;
    });

    if (card) {
        const countEl = card.querySelector('.excel-table-count');
        if (countEl) {
            const total = rows.length;
            if (visibleCount < total) {
                countEl.textContent = visibleCount + '/' + total + ' bản ghi';
            } else {
                countEl.textContent = total + ' bản ghi';
            }
        }
    }
}

// Initialize table filters using event delegation (runs once on page load)
(function initTableFilters() {
    // Text input filter
    document.addEventListener('input', function (e) {
        if (!e.target.classList.contains('column-filter')) return;

        const table = e.target.closest('.filterable-table');
        if (!table) return;

        applyColumnFilters(table);

        // Also filter dropdown items to match typed text
        const wrapper = e.target.closest('.filter-wrapper');
        if (wrapper) {
            const dropdownList = wrapper.querySelector('.filter-dropdown-list');
            if (dropdownList) {
                const typedValue = e.target.value.toLowerCase().trim();
                const items = dropdownList.querySelectorAll('.filter-dropdown-item');
                items.forEach(item => {
                    const itemValue = item.dataset.value.toLowerCase();
                    if (!typedValue || itemValue === '' || itemValue.includes(typedValue)) {
                        item.style.display = '';
                    } else {
                        item.style.display = 'none';
                    }
                });
            }
        }
    });

    // Dropdown toggle click
    document.addEventListener('click', function (e) {
        const toggle = e.target.closest('.filter-dropdown-toggle');
        if (toggle) {
            e.stopPropagation();
            const wrapper = toggle.closest('.filter-wrapper');
            const dropdownList = wrapper.querySelector('.filter-dropdown-list');

            // Close all other open dropdowns
            document.querySelectorAll('.filter-dropdown-list.open').forEach(dl => {
                if (dl !== dropdownList) dl.classList.remove('open');
            });

            // Toggle this dropdown
            dropdownList.classList.toggle('open');

            // Reset visibility of all items when opening
            if (dropdownList.classList.contains('open')) {
                const input = wrapper.querySelector('.column-filter');
                const typedValue = input ? input.value.toLowerCase().trim() : '';
                const items = dropdownList.querySelectorAll('.filter-dropdown-item');
                items.forEach(item => {
                    const itemValue = item.dataset.value.toLowerCase();
                    if (!typedValue || itemValue === '' || itemValue.includes(typedValue)) {
                        item.style.display = '';
                    } else {
                        item.style.display = 'none';
                    }
                });
            }
            return;
        }

        // Dropdown item click
        const dropdownItem = e.target.closest('.filter-dropdown-item');
        if (dropdownItem) {
            const dropdownList = dropdownItem.closest('.filter-dropdown-list');
            const wrapper = dropdownItem.closest('.filter-wrapper');
            const input = wrapper.querySelector('.column-filter');
            const table = wrapper.closest('.filterable-table');

            // Set input value from dropdown selection
            input.value = dropdownItem.dataset.value;

            // Mark active item
            dropdownList.querySelectorAll('.filter-dropdown-item').forEach(item => {
                item.classList.remove('active');
            });
            dropdownItem.classList.add('active');

            // Close dropdown
            dropdownList.classList.remove('open');

            // Apply filters
            if (table) {
                applyColumnFilters(table);
            }
            return;
        }

        // Click outside - close all dropdowns
        document.querySelectorAll('.filter-dropdown-list.open').forEach(dl => {
            dl.classList.remove('open');
        });
    });
})();

/**
 * Download Excel file
 */
function downloadExcel(endpoint, filename) {
    window.location.href = endpoint;
}

/**
 * Search/filter table
 */
function filterTable(tableId, searchTerm) {
    const table = document.getElementById(tableId);
    if (!table) return;

    const rows = table.querySelectorAll('tbody tr');
    const lowerSearchTerm = searchTerm.toLowerCase();

    rows.forEach(row => {
        const text = row.textContent.toLowerCase();
        if (text.includes(lowerSearchTerm)) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}

/**
 * Debounce function
 */
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

/**
 * Get file timestamp from path
 */
function getFileTimestamp(filePath) {
    // This would typically come from the API
    return new Date().toLocaleString('vi-VN');
}

// Export functions for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        formatDate,
        formatNumber,
        showLoading,
        hideLoading,
        showError,
        showEmptyState,
        createExcelTable,
        downloadExcel,
        filterTable,
        debounce,
        getFileTimestamp
    };
}
