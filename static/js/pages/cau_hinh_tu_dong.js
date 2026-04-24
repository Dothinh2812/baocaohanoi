document.addEventListener('DOMContentLoaded', async function () {
    bindDateFilter();
    await initCauHinhTuDongPage();
});

function getRequestedDate() {
    const params = new URLSearchParams(window.location.search);
    return params.get('date') || '';
}

function bindDateFilter() {
    const input = document.getElementById('cau-hinh-tu-dong-date-input');
    const applyButton = document.getElementById('cau-hinh-tu-dong-date-apply');
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
    const input = document.getElementById('cau-hinh-tu-dong-date-input');
    if (input) {
        input.value = data.selected_date || getRequestedDate();
    }

    const meta = document.getElementById('cau-hinh-tu-dong-date-meta');
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

async function initCauHinhTuDongPage() {
    try {
        const endpoint = getRequestedDate()
            ? `/api/cau-hinh-tu-dong/son-tay?date=${encodeURIComponent(getRequestedDate())}`
            : '/api/cau-hinh-tu-dong/son-tay';
        const data = await API.fetchData(endpoint);

        if (data.error) {
            throw new Error(data.error);
        }

        syncDateFilterState(data);
        renderDataTable(
            'cau-hinh-tu-dong-tong-hop',
            data.tong_hop,
            'Tổng hợp cấu hình tự động theo tổ - TTVT Sơn Tây',
            data.file_info,
            data.selected_date
        );
        renderSummaryCards(data.summary, data.file_info, data.selected_date);
        renderDataTable(
            'cau-hinh-tu-dong-team-summary',
            data.team_summary,
            'Tổng hợp cấu hình tự động theo NVKT - TTVT Sơn Tây',
            data.file_info,
            data.selected_date
        );
    } catch (error) {
        showError(`Không thể tải dữ liệu cấu hình tự động: ${error.message}`, 'cau-hinh-tu-dong-tong-hop');
        showError(`Không thể tải dữ liệu cấu hình tự động: ${error.message}`, 'cau-hinh-tu-dong-summary');
        showError(`Không thể tải dữ liệu cấu hình tự động: ${error.message}`, 'cau-hinh-tu-dong-team-summary');
    }
}

function renderSummaryCards(summary, fileInfo, selectedDate) {
    const container = document.getElementById('cau-hinh-tu-dong-summary');
    if (!container) return;

    const cards = [
        { label: 'Tổng bản ghi', value: summary.tong_so },
        { label: 'Thành công', value: summary.thanh_cong, note: `Tỷ lệ ${summary.ty_le_thanh_cong}%` },
        { label: 'Thất bại', value: summary.that_bai },
        { label: 'Chưa có trạng thái', value: summary.chua_co_trang_thai },
        { label: 'Lắp mới', value: summary.lap_moi },
        { label: 'Thay thế', value: summary.thay_the },
        { label: 'Cấu hình WAN', value: summary.cau_hinh_wan },
        { label: 'Cấu hình WiFi', value: summary.cau_hinh_wifi },
    ];

    const timestampHtml = fileInfo && fileInfo.modified
        ? `<div class="summary-card"><div class="summary-card-label">Cập nhật file</div><div class="summary-card-value" style="font-size:1.05rem;">${fileInfo.modified}</div><div class="summary-card-note">${selectedDate ? `Dữ liệu ngày ${selectedDate}` : 'Nguồn chi tiết Sơn Tây'}</div></div>`
        : '';

    container.innerHTML = cards.map(card => `
        <div class="summary-card">
            <div class="summary-card-label">${card.label}</div>
            <div class="summary-card-value">${formatNumber(card.value)}</div>
            <div class="summary-card-note">${card.note || '&nbsp;'}</div>
        </div>
    `).join('') + timestampHtml;
}

function renderDataTable(containerId, sheetData, title, fileInfo, selectedDate) {
    if (!sheetData || !sheetData.columns || !sheetData.data || sheetData.data.length === 0) {
        showEmptyState(`Không có dữ liệu${selectedDate ? ` ngày ${selectedDate}` : ''}`, containerId);
        return;
    }

    const container = document.getElementById(containerId);
    if (!container) return;

    container.innerHTML = createExcelTable(sheetData, title, {
        showRowNumbers: false,
        maxHeight: 'none',
        fileInfo: fileInfo,
    });
}
