document.addEventListener('DOMContentLoaded', async function () {
    await initSHCProcessingReportPage();
});

window.shcProcessingReportState = {
    defaultDate: '',
    report: null
};

async function initSHCProcessingReportPage() {
    bindSHCProcessingReportEvents();

    const params = new URLSearchParams(window.location.search);
    const initialDate = params.get('report_date') || '';
    const initialTeam = params.get('team_filter') || '';

    await loadSHCProcessingReport(initialDate, initialTeam);
}

function bindSHCProcessingReportEvents() {
    const applyButton = document.getElementById('shc-report-apply');
    const resetButton = document.getElementById('shc-report-reset');
    const dateInput = document.getElementById('shc-report-date');
    const teamSelect = document.getElementById('shc-report-team');

    if (applyButton) {
        applyButton.addEventListener('click', async function () {
            await loadSHCProcessingReport(dateInput.value, teamSelect.value);
        });
    }

    if (resetButton) {
        resetButton.addEventListener('click', async function () {
            const resetDate = window.shcProcessingReportState.defaultDate || '';
            dateInput.value = resetDate;
            teamSelect.value = '';
            await loadSHCProcessingReport(resetDate, '');
        });
    }

    if (dateInput) {
        dateInput.addEventListener('keydown', async function (event) {
            if (event.key === 'Enter') {
                await loadSHCProcessingReport(dateInput.value, teamSelect.value);
            }
        });
    }

    if (teamSelect) {
        teamSelect.addEventListener('keydown', async function (event) {
            if (event.key === 'Enter') {
                await loadSHCProcessingReport(dateInput.value, teamSelect.value);
            }
        });
    }
}

async function loadSHCProcessingReport(reportDate = '', teamFilter = '') {
    showLoadingState();

    try {
        const report = await API.getSHCProcessingReport(reportDate, teamFilter);

        if (report.error) {
            throw new Error(report.error);
        }

        window.shcProcessingReportState.report = report;
        if (!window.shcProcessingReportState.defaultDate) {
            window.shcProcessingReportState.defaultDate = report.selected_date;
        }

        syncSHCFilterControls(report);
        updateSHCProcessingReportUrl(report.selected_date, report.selected_team_filter);
        renderSHCProcessingState(report);
        renderSHCKpiCards(report);
        renderSHCDailyUserTable(report.daily_user_rows || []);
        renderSHCDailyTeamTable(report.daily_team_rows || []);
        renderSHCMonthlyTeamTable(report.monthly_team_rows || []);
        renderSHCMonthlyDayTable(report.monthly_day_rows || []);
    } catch (error) {
        console.error('Error loading SHC processing report:', error);
        showSHCReportError(error.message);
    }
}

function showLoadingState() {
    const loadingHtml = `
        <div class="loading">
            <i class="fas fa-spinner"></i>
            <br>Đang tải dữ liệu...
        </div>
    `;

    ['shc-kpi-grid', 'shc-daily-user-table', 'shc-daily-team-table', 'shc-monthly-team-table', 'shc-monthly-day-table']
        .forEach(containerId => {
            const container = document.getElementById(containerId);
            if (container) {
                container.innerHTML = loadingHtml;
            }
        });
}

function showSHCReportError(message) {
    const errorHtml = `
        <div class="empty-table">
            <i class="fas fa-exclamation-triangle"></i>
            <p>Lỗi khi tải báo cáo: ${message}</p>
        </div>
    `;

    const state = document.getElementById('shc-report-state');
    if (state) {
        state.innerHTML = `
            <div class="shc-state-pill" style="background: #fdecea; color: #c62828;">
                <i class="fas fa-circle-exclamation"></i>
                Không thể tải báo cáo
            </div>
        `;
    }

    ['shc-kpi-grid', 'shc-daily-user-table', 'shc-daily-team-table', 'shc-monthly-team-table', 'shc-monthly-day-table']
        .forEach(containerId => {
            const container = document.getElementById(containerId);
            if (container) {
                container.innerHTML = errorHtml;
            }
        });
}

function syncSHCFilterControls(report) {
    const dateInput = document.getElementById('shc-report-date');
    const teamSelect = document.getElementById('shc-report-team');

    if (dateInput) {
        dateInput.value = report.selected_date || '';
    }

    if (teamSelect) {
        const teamOptions = report.team_options || [];
        teamSelect.innerHTML = '<option value="">Tất cả tổ</option>' + teamOptions
            .map(team => `<option value="${escapeHtml(team)}">${escapeHtml(team)}</option>`)
            .join('');
        teamSelect.value = report.selected_team_filter || '';
    }
}

function updateSHCProcessingReportUrl(reportDate, teamFilter) {
    const params = new URLSearchParams();
    if (reportDate) {
        params.set('report_date', reportDate);
    }
    if (teamFilter) {
        params.set('team_filter', teamFilter);
    }

    const queryString = params.toString();
    const nextUrl = queryString ? `${window.location.pathname}?${queryString}` : window.location.pathname;
    window.history.replaceState({}, '', nextUrl);
}

function renderSHCProcessingState(report) {
    const state = document.getElementById('shc-report-state');
    if (!state) {
        return;
    }

    const selectedTeam = report.selected_team_filter || 'Tất cả tổ';
    const dailyProcessed = report.daily_summary?.total_count || 0;
    const dailyRequired = report.daily_required_summary?.total_required_count || 0;

    state.innerHTML = `
        <div class="shc-state-pill">
            <i class="fas fa-calendar-day"></i>
            Ngày: ${escapeHtml(report.selected_date || '')}
        </div>
        <div class="shc-state-pill">
            <i class="fas fa-calendar-alt"></i>
            Tháng: ${escapeHtml(report.selected_month || '')}
        </div>
        <div class="shc-state-pill">
            <i class="fas fa-people-group"></i>
            Bộ lọc tổ: ${escapeHtml(selectedTeam)}
        </div>
        <div class="shc-state-pill">
            <i class="fas fa-gauge-high"></i>
            Hôm nay: ${formatNumber(dailyProcessed)}/${formatNumber(dailyRequired)}
        </div>
    `;
}

function renderSHCKpiCards(report) {
    const container = document.getElementById('shc-kpi-grid');
    if (!container) {
        return;
    }

    const dailySummary = report.daily_summary || {};
    const requiredSummary = report.daily_required_summary || {};
    const monthlySummary = report.monthly_summary || {};
    const completionRate = calculateRate(dailySummary.total_count, requiredSummary.total_required_count);

    container.innerHTML = `
        <div class="shc-kpi-card">
            <div class="shc-kpi-label">Đã xử lý trong ngày</div>
            <div class="shc-kpi-value">${formatNumber(dailySummary.total_count || 0)}</div>
            <div class="shc-kpi-breakdown">
                <span>K1: ${formatNumber(dailySummary.k1_count || 0)}</span>
                <span>K2: ${formatNumber(dailySummary.k2_count || 0)}</span>
                <span>Tỷ lệ: ${completionRate}</span>
            </div>
        </div>
        <div class="shc-kpi-card required">
            <div class="shc-kpi-label">Cần xử lý trong ngày</div>
            <div class="shc-kpi-value">${formatNumber(requiredSummary.total_required_count || 0)}</div>
            <div class="shc-kpi-breakdown">
                <span>K1: ${formatNumber(requiredSummary.k1_required_count || 0)}</span>
                <span>K2: ${formatNumber(requiredSummary.k2_required_count || 0)}</span>
                <span>Chênh lệch: ${formatNumber((requiredSummary.total_required_count || 0) - (dailySummary.total_count || 0))}</span>
            </div>
        </div>
        <div class="shc-kpi-card monthly">
            <div class="shc-kpi-label">Đã xử lý trong tháng</div>
            <div class="shc-kpi-value">${formatNumber(monthlySummary.total_count || 0)}</div>
            <div class="shc-kpi-breakdown">
                <span>K1: ${formatNumber(monthlySummary.k1_count || 0)}</span>
                <span>K2: ${formatNumber(monthlySummary.k2_count || 0)}</span>
                <span>Tháng: ${escapeHtml(report.selected_month || '')}</span>
            </div>
        </div>
    `;
}

function renderSHCDailyUserTable(rows) {
    renderSHCReportTable(
        'shc-daily-user-table',
        'Theo cá nhân trong ngày',
        [
            { key: 'doi_one', label: 'Tổ' },
            { key: 'user_name', label: 'NVKT' },
            { key: 'k1_count', label: 'K1 đã xử lý' },
            { key: 'k1_required_count', label: 'Tổng K1' },
            { key: 'k2_count', label: 'K2 đã xử lý' },
            { key: 'k2_required_count', label: 'Tổng K2' },
            { key: 'total_count', label: 'Tổng đã xử lý' },
            { key: 'total_required_count', label: 'Tổng cần xử lý' }
        ],
        rows
    );
}

function renderSHCDailyTeamTable(rows) {
    renderSHCReportTable(
        'shc-daily-team-table',
        'Theo tổ trong ngày',
        [
            { key: 'doi_one', label: 'Tổ' },
            { key: 'k1_count', label: 'K1 đã xử lý' },
            { key: 'k1_required_count', label: 'Tổng K1' },
            { key: 'k2_count', label: 'K2 đã xử lý' },
            { key: 'k2_required_count', label: 'Tổng K2' },
            { key: 'total_count', label: 'Tổng đã xử lý' },
            { key: 'total_required_count', label: 'Tổng cần xử lý' }
        ],
        rows
    );
}

function renderSHCMonthlyTeamTable(rows) {
    renderSHCReportTable(
        'shc-monthly-team-table',
        'Theo tổ trong tháng',
        [
            { key: 'doi_one', label: 'Tổ' },
            { key: 'k1_count', label: 'K1' },
            { key: 'k2_count', label: 'K2' },
            { key: 'total_count', label: 'Tổng' }
        ],
        rows
    );
}

function renderSHCMonthlyDayTable(rows) {
    renderSHCReportTable(
        'shc-monthly-day-table',
        'Theo ngày trong tháng',
        [
            { key: 'processed_date', label: 'Ngày' },
            { key: 'k1_count', label: 'K1' },
            { key: 'k2_count', label: 'K2' },
            { key: 'total_count', label: 'Tổng' }
        ],
        rows
    );
}

function renderSHCReportTable(containerId, title, columnDefs, rows) {
    const container = document.getElementById(containerId);
    if (!container) {
        return;
    }

    const tableRows = (rows || []).map(row => {
        const mapped = {};
        columnDefs.forEach(column => {
            mapped[column.label] = row[column.key];
        });
        return mapped;
    });

    const sheetData = {
        columns: columnDefs.map(column => column.label),
        data: tableRows
    };

    container.innerHTML = createExcelTable(sheetData, title, {
        showRowNumbers: true,
        maxHeight: '520px'
    });

    applyWarningStyles(container, rows || []);
}

function applyWarningStyles(container, rows) {
    const bodyRows = container.querySelectorAll('tbody tr.data-row');
    bodyRows.forEach((element, index) => {
        if (rows[index] && rows[index].highlight_warning) {
            element.classList.add('shc-warning-row');
        }
    });
}

function calculateRate(processed, required) {
    const processedValue = Number(processed || 0);
    const requiredValue = Number(required || 0);
    if (!requiredValue) {
        return '0%';
    }
    return `${((processedValue / requiredValue) * 100).toFixed(1)}%`;
}

function escapeHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
