/* ========================================
   THU HỒI TBĐC PAGE JAVASCRIPT
   - Load equipment recovery data
   ======================================== */

document.addEventListener('DOMContentLoaded', async function() {
    await initThuHoiPage();
});

async function initThuHoiPage() {
    await loadThuHoiData();
}

/**
 * Load equipment recovery data
 */
async function loadThuHoiData() {
    const summaryContainer = document.getElementById('thuhoi-summary-container');
    const detailContainer = document.getElementById('thuhoi-detail-container');

    try {
        const data = await API.getThuHoiData();

        if (data.error) {
            throw new Error(data.error);
        }

        // API returns: {tong_hop, chi_tiet, file_info}
        if (data) {
            const fileInfo = data.file_info || null;

            // Render Tổng hợp table
            if (data.tong_hop) {
                summaryContainer.innerHTML = '';
                const summaryHtml = createExcelTable(data.tong_hop, 'Tổng hợp Thu hồi TBĐC', {
                    showRowNumbers: true,
                    tableClass: 'excel-table',
                    maxHeight: '600px',
                    fileInfo: fileInfo
                });
                summaryContainer.innerHTML = summaryHtml;
            }

            // Render Chi tiết vật tư table
            if (data.chi_tiet) {
                detailContainer.innerHTML = '';
                const detailHtml = createExcelTable(data.chi_tiet, 'Chi tiết vật tư thu hồi', {
                    showRowNumbers: true,
                    tableClass: 'excel-table',
                    maxHeight: '600px',
                    fileInfo: fileInfo
                });
                detailContainer.innerHTML = detailHtml;
            }
        }
    } catch (error) {
        console.error('Error loading Thu hoi data:', error);
        const errorHtml = `
            <div class="loading">
                <i class="fas fa-exclamation-triangle"></i>
                <br>Lỗi khi tải dữ liệu: ${error.message}
            </div>
        `;
        if (summaryContainer) summaryContainer.innerHTML = errorHtml;
        if (detailContainer) detailContainer.innerHTML = errorHtml;
    }
}
