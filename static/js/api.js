/* ========================================
   API MODULE
   - Centralized API calls
   - Error handling
   - Data fetching
   ======================================== */

const API = {
    /**
     * Base fetch function with error handling
     */
    async fetchData(endpoint) {
        try {
            const response = await fetch(endpoint, {
                method: 'GET',
                headers: {
                    'Cache-Control': 'no-cache, no-store, must-revalidate',
                    'Pragma': 'no-cache',
                    'Expires': '0'
                }
            });

            if (response.status === 401) {
                window.location.href = '/login';
                throw new Error('Phiên đăng nhập đã hết hạn');
            }

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            return data;
        } catch (error) {
            console.error(`Error fetching ${endpoint}:`, error);
            throw error;
        }
    },

    /* ========================================
       BRCD APIs (Điều hành sửa chữa)
       ======================================== */

    /**
     * Get BRCD down port data
     */
    getBRCDData() {
        return this.fetchData('/api/excel-data');
    },

    /**
     * Get BRCD 4 teams data
     */
    getBRCDMainData() {
        return this.fetchData('/api/excel-data-main');
    },

    /**
     * Get BRCD pending tickets
     */
    getBRCDPendingData() {
        return this.fetchData('/api/excel-data-pending');
    },

    /* ========================================
       PTTB APIs (Phát triển thuê bao)
       ======================================== */

    /**
     * Get PTTB summary data
     */
    getPTTBSummary() {
        return this.fetchData('/api/pttb-data-summary');
    },

    /**
     * Get PTTB detail data
     */
    getPTTBDetail() {
        return this.fetchData('/api/pttb-data-detail');
    },

    /**
     * Get PTTB pending tickets
     */
    getPTTBPending() {
        return this.fetchData('/api/pttb-data-pending');
    },

    /**
     * Get auto-configuration detail data for TTVT Son Tay
     */
    getCauHinhTuDongSonTay() {
        return this.fetchData('/api/cau-hinh-tu-dong/son-tay');
    },

    /**
     * Get consolidated dashboard data for TTVT Son Tay
     */
    getTTVTSonTayTongHop(date = '') {
        const query = new URLSearchParams();
        if (date) {
            query.set('date', date);
        }
        const suffix = query.toString() ? `?${query.toString()}` : '';
        return this.fetchData(`/api/ttvt-son-tay-tong-hop${suffix}`);
    },

    /* ========================================
       Gia hạn TTTC APIs
       ======================================== */

    /**
     * Get renewal data (KR6)
     */
    getGiaHanData() {
        return this.fetchData('/api/giahan-data');
    },

    /**
     * Get renewal data by team (KR6)
     */
    getGiaHanDataTo() {
        return this.fetchData('/api/giahan-data-to');
    },

    /**
     * Get renewal KR7 data
     */
    getGiaHanKR7() {
        return this.fetchData('/api/giahan-kr7-data');
    },

    /**
     * Get renewal KR7 data by team
     */
    getGiaHanKR7To() {
        return this.fetchData('/api/giahan-kr7-data-to');
    },

    /* ========================================
       Thu hồi TBĐC APIs
       ======================================== */

    /**
     * Get equipment recovery data
     */
    getThuHoiData() {
        return this.fetchData('/api/thu-hoi-data');
    },

    /* ========================================
       C1 - Chất lượng APIs
       ======================================== */

    /**
     * Get C1 quality data
     */
    getChatLuongData() {
        return this.fetchData('/api/c1-chat-luong-data');
    },

    /* ========================================
       I1.5 - Chất lượng mạng APIs
       ======================================== */

    /**
     * Get I1.5 network quality data
     */
    getI15Data() {
        return this.fetchData('/api/i15-data');
    },

    /* ========================================
       File Info API
       ======================================== */

    /**
     * Get file metadata (charts, images)
     */
    getFileInfo() {
        return this.fetchData('/api/file-info');
    },

    /* ========================================
       SHC Variation APIs
       ======================================== */

    /**
     * Get SHC variation data (T-1 comparison)
     */
    getSHCVariationData() {
        return this.fetchData('/api/shc-variation-data');
    },

    /* ========================================
       I1.5 K2 - Chất lượng mạng APIs (Kỳ 2)
       ======================================== */

    /**
     * Get I1.5 K2 network quality data
     */
    getI15K2Data() {
        return this.fetchData('/api/i15k2-data');
    },

    /**
     * Get SHC K2 variation data (T-1 comparison)
     */
    getSHCVariationK2Data() {
        return this.fetchData('/api/shc-variation-k2-data');
    },

    /**
     * Get SHC processing report data
     */
    getSHCProcessingReport(reportDate = '', teamFilter = '') {
        const query = new URLSearchParams();
        if (reportDate) {
            query.set('report_date', reportDate);
        }
        if (teamFilter) {
            query.set('team_filter', teamFilter);
        }
        const suffix = query.toString() ? `?${query.toString()}` : '';
        return this.fetchData(`/api/shc-processing-report${suffix}`);
    },

    /**
     * Get SHC NVKT detail download options
     */
    getSHCDetailOptions(reportType = '', date = '') {
        const query = new URLSearchParams();
        if (reportType) {
            query.set('report_type', reportType);
        }
        if (date) {
            query.set('date', date);
        }
        const suffix = query.toString() ? `?${query.toString()}` : '';
        return this.fetchData(`/api/shc-nvkt-detail/options${suffix}`);
    },

    /**
     * Get SHC NVKT detail files by report type and team
     */
    getSHCDetailFiles(reportType, team, date = '') {
        const query = new URLSearchParams({
            report_type: reportType,
            team: team
        });
        if (date) {
            query.set('date', date);
        }
        return this.fetchData(`/api/shc-nvkt-detail/files?${query.toString()}`);
    },

    /**
     * Preview SHC NVKT detail file content
     */
    getSHCDetailPreview(reportType, team, fileName, date = '') {
        const query = new URLSearchParams({
            report_type: reportType,
            team: team,
            file_name: fileName
        });
        if (date) {
            query.set('date', date);
        }
        return this.fetchData(`/api/shc-nvkt-detail/preview?${query.toString()}`);
    },

    /* ========================================
       Tiếp thị APIs
       ======================================== */

    /**
     * Get marketing results data
     */
    getTiepThiData() {
        return this.fetchData('/api/tiepthi-data');
    }
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = API;
}
