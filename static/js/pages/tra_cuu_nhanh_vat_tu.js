/* ========================================
   TRA CUU NHANH VAT TU PAGE JAVASCRIPT
   ======================================== */

document.addEventListener('DOMContentLoaded', function () {
    loadTraCuuNhanhVatTu();
});

async function loadTraCuuNhanhVatTu() {
    var container = document.getElementById('tra-cuu-table-container');
    if (!container) return;

    try {
        var response = await fetch('/api/tra-cuu-nhanh-vat-tu');
        if (!response.ok) {
            throw new Error('HTTP error ' + response.status);
        }

        var data = await response.json();

        if (data.error) {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>' + data.error + '</p></div>';
            return;
        }

        if (data.columns && data.data) {
            container.innerHTML = createExcelTable(data, 'Chi tiết tồn kho vật tư', {
                showRowNumbers: false,
                maxHeight: 'none',
                fileInfo: data.file_info,
                enableFilter: true
            });
        } else {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
        }
    } catch (error) {
        console.error('Error loading tra cuu nhanh vat tu:', error);
        container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Lỗi khi tải dữ liệu: ' + error.message + '</p></div>';
    }
}
