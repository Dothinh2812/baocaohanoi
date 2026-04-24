document.addEventListener('DOMContentLoaded', async function () {
    bindDateFilter();
    await loadChatLuongData();
});

const CHAT_LUONG_VIEW_ORDER = [
    'v_chi_tieu_c_c1_1_report_th_c1_1',
    'v_chi_tieu_c_c1_1_chitiet_report_chi_tiet',
    'v_chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h',
    'v_chi_tieu_c_c1_2_report_th_c1_2',
    'v_chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang',
    'v_chi_tieu_c_c1_3_report_th_c1_3',
    'v_chi_tieu_c_c1_4_report_th_c1_4',
    'v_chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt',
    'v_chi_tieu_c_c1_5_report_th_c1_5'
];

async function loadChatLuongData() {
    try {
        const query = new URLSearchParams(window.location.search);
        const selectedDate = query.get('date');
        const endpoint = selectedDate ? `/api/c1-chat-luong-data?date=${encodeURIComponent(selectedDate)}` : '/api/c1-chat-luong-data';
        const response = await fetch(endpoint);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

        const data = await response.json();
        if (data.error) throw new Error(data.error);

        syncDateFilterState(data);

        CHAT_LUONG_VIEW_ORDER.forEach((viewId) => {
            const container = document.getElementById(`${viewId}-container`);
            if (!container) return;

            const sheetData = data.sheets?.[viewId];
            if (!sheetData?.data?.length) {
                container.innerHTML = `
                    <div class="empty-table">
                        <i class="fas fa-inbox"></i>
                        <p>Không có dữ liệu view ${viewId}${data.selected_date ? ` cho ngày ${data.selected_date}` : ''}</p>
                    </div>
                `;
                return;
            }

            container.innerHTML = createExcelTable(sheetData, viewId, {
                showRowNumbers: true,
                maxHeight: 'none',
                fileInfo: data.file_info || null,
                enableFilter: true
            });
        });
    } catch (error) {
        console.error('Error loading chat luong data:', error);
        const errorHtml = `
            <div class="empty-table">
                <i class="fas fa-exclamation-triangle"></i>
                <p>Lỗi khi tải dữ liệu: ${error.message}</p>
            </div>
        `;
        CHAT_LUONG_VIEW_ORDER.forEach((viewId) => {
            const container = document.getElementById(`${viewId}-container`);
            if (container) container.innerHTML = errorHtml;
        });
    }
}

function bindDateFilter() {
    const input = document.getElementById('chatluong-date-input');
    const button = document.getElementById('chatluong-date-apply');
    if (!input || !button) return;

    button.addEventListener('click', function () {
        const query = new URLSearchParams(window.location.search);
        const selectedDate = String(input.value || '').trim();
        if (selectedDate) {
            query.set('date', selectedDate);
        } else {
            query.delete('date');
        }
        const suffix = query.toString() ? `?${query.toString()}` : '';
        window.location.href = `${window.location.pathname}${suffix}`;
    });
}

function syncDateFilterState(data) {
    const input = document.getElementById('chatluong-date-input');
    const meta = document.getElementById('chatluong-date-meta');
    if (input && data.selected_date) {
        input.value = data.selected_date;
    }
    if (!meta) return;

    const latestText = data.latest_available_date ? `Ngày mới nhất: ${data.latest_available_date}` : 'Chưa có ngày dữ liệu khả dụng';
    if (data.date_has_data === false && data.selected_date) {
        meta.textContent = `Không có dữ liệu cho ngày ${data.selected_date}. ${latestText}`;
        return;
    }
    if (data.selected_date) {
        meta.textContent = `Đang xem ngày ${data.selected_date}. ${latestText}`;
        return;
    }
    meta.textContent = latestText;
}
