/* ========================================
   TON KHO VAT TU PAGE JAVASCRIPT
   - Load data from API
   - Render table with column filters using createExcelTable
   ======================================== */

document.addEventListener('DOMContentLoaded', function () {
    loadTonKhoVatTu();
    loadTonThuongDung();
});

/**
 * Load dữ liệu tồn kho vật tư từ API và render table
 */
/**
 * Load dữ liệu tổng hợp tồn vật tư thường dùng (tốt) theo đơn vị
 */
async function loadTonThuongDung() {
    var container = document.getElementById('ton-thuong-dung-table-container');
    if (!container) return;

    try {
        var response = await fetch('/api/ton-kho-vat-tu-tot-thuong-dung');
        if (!response.ok) {
            throw new Error('HTTP error ' + response.status);
        }

        var data = await response.json();

        if (data.error) {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>' + data.error + '</p></div>';
            return;
        }

        if (data.columns && data.data) {
            container.innerHTML = createExcelTable(data, 'Tổng hợp tồn vật tư thường dùng (tốt) theo đơn vị', {
                showRowNumbers: false,
                maxHeight: 'none',
                fileInfo: data.file_info,
                enableFilter: true,
                frozenColumns: 3
            });
            initializeFrozenTableLayouts(container);
        } else {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
        }
    } catch (error) {
        console.error('Error loading ton thuong dung:', error);
        container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Lỗi khi tải dữ liệu: ' + error.message + '</p></div>';
    }
}

async function loadTonKhoVatTu() {
    var container1 = document.getElementById('ton-kho-table-container');
    var container2 = document.getElementById('ton-theo-loai-table-container');

    try {
        var response = await fetch('/api/ton-kho-vat-tu');
        if (!response.ok) {
            throw new Error('HTTP error ' + response.status);
        }

        var data = await response.json();

        if (data.error) {
            var errorHtml = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>' + data.error + '</p></div>';
            if (container1) container1.innerHTML = errorHtml;
            if (container2) container2.innerHTML = errorHtml;
            return;
        }

        if (data.sheets) {
            // Table 1: Tổng hợp tồn vật tư
            if (container1) {
                var sheetData1 = data.sheets['Tổng hợp'];
                if (sheetData1) {
                    container1.innerHTML = createExcelTable(sheetData1, 'Tổng hợp tồn vật tư', {
                        showRowNumbers: false,
                        maxHeight: 'none',
                        fileInfo: data.file_info,
                        enableFilter: true,
                        frozenColumns: 3
                    });
                    initializeFrozenTableLayouts(container1);
                } else {
                    container1.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
                }
            }

            // Table 2: Tồn theo loại vật tư
            if (container2) {
                var sheetData2 = data.sheets['thong-ke-theo-vat-tu'];
                if (sheetData2) {
                    container2.innerHTML = createExcelTable(sheetData2, 'Tồn theo loại vật tư', {
                        showRowNumbers: false,
                        maxHeight: 'none',
                        fileInfo: data.file_info,
                        enableFilter: true,
                        frozenColumns: 3
                    });
                    initializeFrozenTableLayouts(container2);
                } else {
                    container2.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
                }
            }
        }
    } catch (error) {
        console.error('Error loading ton kho vat tu:', error);
        var errHtml = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Lỗi khi tải dữ liệu: ' + error.message + '</p></div>';
        if (container1) container1.innerHTML = errHtml;
        if (container2) container2.innerHTML = errHtml;
    }
}
