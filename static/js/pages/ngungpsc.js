document.addEventListener('DOMContentLoaded', async function() {
    bindDateFilter();
    await Promise.all([
        loadNgungPscFiberData(),
        loadNgungPscMytvData()
    ]);
});

function getRequestedDate() {
    const params = new URLSearchParams(window.location.search);
    return params.get('date') || '';
}

function bindDateFilter() {
    const input = document.getElementById('ngungpsc-date-input');
    const applyButton = document.getElementById('ngungpsc-date-apply');
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
    const input = document.getElementById('ngungpsc-date-input');
    if (input) {
        input.value = data.selected_date || getRequestedDate();
    }

    const meta = document.getElementById('ngungpsc-date-meta');
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

function updateTimestamp(fileInfo) {
    const timestampEl = document.getElementById('ngungpsc-timestamp');
    if (timestampEl && fileInfo) {
        timestampEl.textContent = `(${fileInfo.name} - Cập nhật: ${fileInfo.modified})`;
    }
}

async function loadNgungPscFiberData() {
    const container = document.getElementById('ngungpsc-container');

    try {
        const requestedDate = getRequestedDate();
        const endpoint = requestedDate
            ? `/api/ngungpsc-data?date=${encodeURIComponent(requestedDate)}`
            : '/api/ngungpsc-data';
        const response = await fetch(endpoint);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const data = await response.json();
        if (data.error) throw new Error(data.error);

        syncDateFilterState(data);
        updateTimestamp(data.file_info);

        const selectedDateLabel = data.selected_date ? ` ngày ${data.selected_date}` : '';
        const sheetData = data.sheets['v_ngung_psc_fiber_thang_t_1_cap_ttvt'];
        if (!sheetData) throw new Error('Không tìm thấy sheet v_ngung_psc_fiber_thang_t_1_cap_ttvt');

        if (!sheetData.data?.length) {
            container.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu Fiber${selectedDateLabel}</p></div>`;
            return;
        }

        const html = createExcelTable(sheetData, 'v_ngung_psc_fiber_thang_t_1_cap_ttvt', {
            showRowNumbers: true,
            maxHeight: 'none',
            fileInfo: data.file_info,
            enableFilter: true
        });

        container.innerHTML = html;
    } catch (error) {
        console.error('Error loading Ngưng PSC Fiber data:', error);
        container.innerHTML = `
            <div class="empty-table">
                <i class="fas fa-exclamation-triangle"></i>
                <p>Lỗi khi tải dữ liệu: ${error.message}</p>
            </div>
        `;
    }
}

async function loadNgungPscMytvData() {
    const container = document.getElementById('ngungpsc-mytv-container');

    try {
        const requestedDate = getRequestedDate();
        const endpoint = requestedDate
            ? `/api/ngungpsc-mytv-data?date=${encodeURIComponent(requestedDate)}`
            : '/api/ngungpsc-mytv-data';
        const response = await fetch(endpoint);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const data = await response.json();
        if (data.error) throw new Error(data.error);

        syncDateFilterState(data);

        const selectedDateLabel = data.selected_date ? ` ngày ${data.selected_date}` : '';
        const sheetData = data.sheets['v_ngung_psc_mytv_thang_t_1_cap_ttvt'];
        if (!sheetData) throw new Error('Không tìm thấy sheet v_ngung_psc_mytv_thang_t_1_cap_ttvt');

        if (!sheetData.data?.length) {
            container.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu MyTV${selectedDateLabel}</p></div>`;
            return;
        }

        const html = createExcelTable(sheetData, 'v_ngung_psc_mytv_thang_t_1_cap_ttvt', {
            showRowNumbers: true,
            maxHeight: 'none',
            fileInfo: data.file_info,
            enableFilter: true
        });

        container.innerHTML = html;
    } catch (error) {
        console.error('Error loading Ngưng PSC MyTV data:', error);
        container.innerHTML = `
            <div class="empty-table">
                <i class="fas fa-exclamation-triangle"></i>
                <p>Lỗi khi tải dữ liệu: ${error.message}</p>
            </div>
        `;
    }
}
