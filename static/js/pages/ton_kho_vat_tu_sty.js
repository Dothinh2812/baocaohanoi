/* ========================================
   TON KHO VAT TU SON TAY PAGE JAVASCRIPT
   ======================================== */

document.addEventListener('DOMContentLoaded', function () {
    loadTonKhoVatTuSTY();
    loadTonThuongDungSTY();
});

async function loadTonThuongDungSTY() {
    var container = document.getElementById('ton-thuong-dung-sty-container');
    if (!container) return;

    try {
        var response = await fetch('/api/ton-kho-vat-tu-sty-tot-thuong-dung');
        if (!response.ok) {
            throw new Error('HTTP error ' + response.status);
        }

        var data = await response.json();

        if (data.error) {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>' + data.error + '</p></div>';
            return;
        }

        if (data.columns && data.data) {
            container.innerHTML = createExcelTable(data, 'Tổng hợp tồn vật tư thường dùng (tốt) - Sơn Tây', {
                showRowNumbers: false,
                maxHeight: 'none',
                fileInfo: data.file_info,
                enableFilter: true
            });
        } else {
            container.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
        }
    } catch (error) {
        console.error('Error loading ton thuong dung STY:', error);
        container.innerHTML = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Lỗi khi tải dữ liệu: ' + error.message + '</p></div>';
    }
}

async function loadTonKhoVatTuSTY() {
    var container1 = document.getElementById('ton-kho-nvkt-container');
    var container2 = document.getElementById('ton-theo-loai-sty-container');

    try {
        var response = await fetch('/api/ton-kho-vat-tu-sty');
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
                    container1.innerHTML = createExcelTable(sheetData1, 'Tồn kho Sơn Tây theo NVKT', {
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
        console.error('Error loading ton kho vat tu STY:', error);
        var errHtml = '<div class="empty-table"><i class="fas fa-exclamation-triangle"></i><p>Lỗi khi tải dữ liệu: ' + error.message + '</p></div>';
        if (container1) container1.innerHTML = errHtml;
        if (container2) container2.innerHTML = errHtml;
    }
}
