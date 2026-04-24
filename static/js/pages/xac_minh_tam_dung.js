/* ========================================
   XÁC MINH TẠM DỪNG PAGE JAVASCRIPT
   ======================================== */

document.addEventListener('DOMContentLoaded', async function() {
    await loadXacMinhTamDungData();
});

let xmtdData = null;

async function loadXacMinhTamDungData() {
    const tabsContainer = document.getElementById('xmtd-tabs');
    const fiberContainer = document.getElementById('xmtd-fiber-container');
    const mytvContainer = document.getElementById('xmtd-mytv-container');

    try {
        const response = await fetch('/api/xac-minh-tam-dung-data');
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const data = await response.json();
        if (data.error) throw new Error(data.error);

        xmtdData = data;

        const tsEl = document.getElementById('xmtd-timestamp');
        if (tsEl && data.file_info) {
            tsEl.textContent = `(${data.file_info.name} - Cập nhật: ${data.file_info.modified})`;
        }

        const sheetNames = Object.keys(data.sheets);
        if (sheetNames.length === 0) {
            tabsContainer.innerHTML = '';
            fiberContainer.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
            mytvContainer.innerHTML = '';
            return;
        }

        let tabsHtml = '';
        sheetNames.forEach((name, index) => {
            tabsHtml += `<li class="excel-tab${index === 0 ? ' active' : ''}" data-sheet="${name}">${name}</li>`;
        });
        tabsContainer.innerHTML = tabsHtml;

        renderTabContent(sheetNames[0], data);

        tabsContainer.querySelectorAll('.excel-tab').forEach(tab => {
            tab.addEventListener('click', function() {
                tabsContainer.querySelectorAll('.excel-tab').forEach(t => t.classList.remove('active'));
                this.classList.add('active');
                renderTabContent(this.dataset.sheet, data);
            });
        });

    } catch (error) {
        console.error('Error loading xác minh tạm dừng data:', error);
        tabsContainer.innerHTML = '';
        fiberContainer.innerHTML = `
            <div class="empty-table">
                <i class="fas fa-exclamation-triangle"></i>
                <p>Lỗi khi tải dữ liệu: ${error.message}</p>
            </div>
        `;
        mytvContainer.innerHTML = '';
    }
}

function renderTabContent(sheetName, data) {
    const fiberContainer = document.getElementById('xmtd-fiber-container');
    const mytvContainer = document.getElementById('xmtd-mytv-container');
    const sheetData = data.sheets[sheetName];

    if (!sheetData) {
        fiberContainer.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu</p></div>';
        mytvContainer.innerHTML = '';
        return;
    }

    // Fiber table
    if (sheetData.fiber && sheetData.fiber.data.length > 0) {
        fiberContainer.innerHTML = createExcelTable(sheetData.fiber, 'KQ xác minh tạm dừng Fiber - ' + sheetName, {
            showRowNumbers: true,
            maxHeight: 'none',
            fileInfo: data.file_info,
            enableFilter: true
        });
    } else {
        fiberContainer.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu Fiber cho ' + sheetName + '</p></div>';
    }

    // MyTV table
    if (sheetData.mytv && sheetData.mytv.data.length > 0) {
        mytvContainer.innerHTML = createExcelTable(sheetData.mytv, 'KQ xác minh tạm dừng MyTV - ' + sheetName, {
            showRowNumbers: true,
            maxHeight: 'none',
            fileInfo: data.file_info,
            enableFilter: true
        });
    } else {
        mytvContainer.innerHTML = '<div class="empty-table"><i class="fas fa-inbox"></i><p>Không có dữ liệu MyTV cho ' + sheetName + '</p></div>';
    }
}
