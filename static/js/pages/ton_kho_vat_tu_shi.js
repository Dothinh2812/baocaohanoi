/* ========================================
   TON KHO VAT TU SUOI HAI PAGE JAVASCRIPT
   ======================================== */

document.addEventListener('DOMContentLoaded', function () {
    loadTonKhoVatTuSHI();
    loadTonThuongDungSHI();
});

async function loadTonThuongDungSHI() {
    var container = document.getElementById('ton-thuong-dung-shi-container');
    if (!container) return;

    try {
        var response = await fetch('/api/ton-kho-vat-tu-shi-tot-thuong-dung');
        if (!response.ok) {
            throw new Error('HTTP error ' + response.status);
        }

        var data = await response.json();

        if (data.error) {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>' + data.error + '</p></div>';
            return;
        }

        if (data.columns && data.data) {
            container.innerHTML = createExcelTable(data, 'Tổng hợp tồn vật tư thường dùng (tốt) - Suối Hai', {
                showRowNumbers: false,
                maxHeight: 'none',
                fileInfo: data.file_info,
                enableFilter: true
            });
        } else {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
        }
    } catch (error) {
        console.error('Error loading ton thuong dung SHI:', error);
        container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Lỗi khi tải dữ liệu: ' + error.message + '</p></div>';
    }
}

async function loadTonKhoVatTuSHI() {
    var container1 = document.getElementById('ton-kho-nvkt-container');
    var container2 = document.getElementById('ton-theo-loai-shi-container');

    try {
        var response = await fetch('/api/ton-kho-vat-tu-shi');
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
            // Table 1: Tồn kho theo NVKT
            if (container1) {
                var sheetData1 = data.sheets['ton-kho-nvkt'];
                if (sheetData1) {
                    container1.innerHTML = createExcelTable(sheetData1, 'Tồn kho Suối Hai theo NVKT', {
                        showRowNumbers: false,
                        maxHeight: 'none',
                        fileInfo: data.file_info,
                        enableFilter: true
                    });
                } else {
                    container1.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
                }
            }

            // Table 2: Tồn theo chủng loại
            if (container2) {
                var sheetData2 = data.sheets['ton-theo-loai'];
                if (sheetData2) {
                    container2.innerHTML = createExcelTable(sheetData2, 'Tồn kho vật tư theo chủng loại', {
                        showRowNumbers: false,
                        maxHeight: 'none',
                        fileInfo: data.file_info,
                        enableFilter: true
                    });
                } else {
                    container2.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
                }
            }
        }
    } catch (error) {
        console.error('Error loading ton kho vat tu SHI:', error);
        var errHtml = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Lỗi khi tải dữ liệu: ' + error.message + '</p></div>';
        if (container1) container1.innerHTML = errHtml;
        if (container2) container2.innerHTML = errHtml;
    }
}
