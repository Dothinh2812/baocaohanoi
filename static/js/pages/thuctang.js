/* ========================================
   THỰC TĂNG PAGE JAVASCRIPT
   - Display Fiber and MyTV growth charts
   - Load chart timestamps
   ======================================== */

document.addEventListener('DOMContentLoaded', async function() {
    await loadChartTimestamps();
});

/**
 * Load chart timestamps from API and display them
 */
async function loadChartTimestamps() {
    try {
        const response = await fetch('/api/file-info');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();

        if (data && data.baocaohanoi_charts) {
            // Map of chart paths to timestamp element IDs
            const chartMappings = [
                { path: 'thuc_tang_fiber/thuc_tang_fiber_pttb.png', id: 'timestamp-thuc_tang_fiber-thuc_tang_fiber_pttb-png' },
                { path: 'thuc_tang_fiber/fiber_thuctang_nvkt.png', id: 'timestamp-thuc_tang_fiber-fiber_thuctang_nvkt-png' },
                { path: 'thuc_tang_mytv/thuc_tang_mytv_pttb.png', id: 'timestamp-thuc_tang_mytv-thuc_tang_mytv_pttb-png' },
                { path: 'thuc_tang_mytv/mytv_thuctang_nvkt.png', id: 'timestamp-thuc_tang_mytv-mytv_thuctang_nvkt-png' }
            ];

            chartMappings.forEach(mapping => {
                const chartInfo = data.baocaohanoi_charts[mapping.path];
                if (chartInfo && chartInfo.created) {
                    const element = document.getElementById(mapping.id);
                    if (element) {
                        element.textContent = `Cập nhật: ${chartInfo.created}`;
                    }
                }
            });
        }
    } catch (error) {
        console.error('Error loading chart timestamps:', error);
    }
}
