document.addEventListener('DOMContentLoaded', async function () {
    bindDateFilter();
    await loadTTVTSonTayTongHop();
});

async function loadTTVTSonTayTongHop() {
    try {
        const payload = await API.getTTVTSonTayTongHop(getRequestedDate());
        if (payload.error) {
            throw new Error(payload.error);
        }

        syncDateFilterState(payload);
        renderChatLuongCTable(payload.chat_luong_c, payload.file_info);
    } catch (error) {
        const message = `Không thể tải chỉ tiêu chất lượng C: ${error.message}`;
        showError(message, 'ttvt-son-tay-chat-luong-c');
    }
}

function renderChatLuongCTable(tableData, fileInfo) {
    const container = document.getElementById('ttvt-son-tay-chat-luong-c');
    if (!container) return;

    if (!tableData || !tableData.columns) {
        showEmptyState('Không có dữ liệu', 'ttvt-son-tay-chat-luong-c');
        return;
    }

    container.innerHTML = createExcelTable(tableData, 'Chỉ tiêu chất lượng C', {
        showRowNumbers: true,
        maxHeight: 'none',
        fileInfo,
        frozenColumns: 1,
        enableFilter: true,
    });
}

function bindDateFilter() {
    const input = document.getElementById('ttvt-son-tay-date-input');
    const button = document.getElementById('ttvt-son-tay-date-apply');
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

function getRequestedDate() {
    const query = new URLSearchParams(window.location.search);
    return String(query.get('date') || '').trim();
}

function syncDateFilterState(data) {
    const input = document.getElementById('ttvt-son-tay-date-input');
    const meta = document.getElementById('ttvt-son-tay-date-meta');

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
