document.addEventListener('DOMContentLoaded', async function() {
    bindDateFilter();
    await loadTamDungKhoiPhucData();
});

function getRequestedDate() {
    const params = new URLSearchParams(window.location.search);
    return params.get('date') || '';
}

function bindDateFilter() {
    const input = document.getElementById('tdkp-date-input');
    const applyButton = document.getElementById('tdkp-date-apply');
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
    const input = document.getElementById('tdkp-date-input');
    if (input) {
        input.value = data.selected_date || getRequestedDate();
    }

    const meta = document.getElementById('tdkp-date-meta');
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

async function loadTamDungKhoiPhucData() {
    const theoToContainer = document.getElementById('tdkp-theo-to-container');
    const theoNvktContainer = document.getElementById('tdkp-theo-nvkt-container');

    try {
        const requestedDate = getRequestedDate();
        const endpoint = requestedDate
            ? `/api/tam-dung-khoi-phuc-data?date=${encodeURIComponent(requestedDate)}`
            : '/api/tam-dung-khoi-phuc-data';
        const response = await fetch(endpoint);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

        const data = await response.json();
        if (data.error) throw new Error(data.error);

        syncDateFilterState(data);

        const timestampEl = document.getElementById('tdkp-timestamp');
        if (timestampEl && data.file_info) {
            timestampEl.textContent = `(${data.file_info.name} - Cập nhật: ${data.file_info.modified})`;
        }

        const selectedDateLabel = data.selected_date ? ` ngày ${data.selected_date}` : '';
        const theoToData = data.sheets?.v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to;
        const theoNvktData = data.sheets?.v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt;

        if (theoToData?.data?.length) {
            theoToContainer.innerHTML = createExcelTable(theoToData, 'v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to', {
                showRowNumbers: true,
                maxHeight: 'none',
                fileInfo: data.file_info,
                enableFilter: true
            });
        } else {
            theoToContainer.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu bảng tổng hợp theo tổ${selectedDateLabel}</p></div>`;
        }

        if (theoNvktData?.data?.length) {
            theoNvktContainer.innerHTML = createExcelTable(theoNvktData, 'v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt', {
                showRowNumbers: true,
                maxHeight: 'none',
                fileInfo: data.file_info,
                enableFilter: true
            });
        } else {
            theoNvktContainer.innerHTML = `<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu bảng tổng hợp theo NVKT${selectedDateLabel}</p></div>`;
        }
    } catch (error) {
        console.error('Error loading tạm dừng khôi phục data:', error);
        const errorHtml = `
            <div class="empty-table">
                <i class="fas fa-exclamation-triangle"></i>
                <p>Lỗi khi tải dữ liệu: ${error.message}</p>
            </div>
        `;
        theoToContainer.innerHTML = errorHtml;
        theoNvktContainer.innerHTML = errorHtml;
    }
}
