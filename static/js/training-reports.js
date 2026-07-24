(function (window, document) {
    'use strict';

    var panel = document.getElementById('training-report');
    if (!panel || !window.TrainingUI) return;

    var errors = document.getElementById('report-errors');
    var listView = document.getElementById('report-list-view');
    var detail = document.getElementById('report-detail');
    var listBody = document.getElementById('report-list');
    var pagination = document.getElementById('report-pagination');

    var listPage = 1;
    var listPageSize = 25;
    var loaded = false;

    function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); }

    function renderErrors(el, error) {
        clear(el);
        if (!error) return;
        var heading = document.createElement('strong');
        heading.textContent = error.message || 'Không thể hoàn tất yêu cầu.';
        el.appendChild(heading);
        var messages = error.details && error.details.errors;
        if (Array.isArray(messages) && messages.length) {
            var ul = document.createElement('ul');
            messages.forEach(function (m) {
                var li = document.createElement('li');
                li.textContent = m;
                ul.appendChild(li);
            });
            el.appendChild(ul);
        }
    }

    function statusBadge(s, finalizedAtMs) {
        var span = document.createElement('span');
        var cls = s;
        var text = { draft: 'Nháp', ready: 'Sẵn sàng', open: 'Đang mở', closed: 'Đã đóng', cancelled: 'Đã hủy' }[s] || s;
        if (s === 'closed' && finalizedAtMs) { cls = 'finalized'; text = 'Đã chốt'; }
        span.className = 'training-status-badge training-status-' + cls;
        span.textContent = text;
        return span;
    }

    function appendText(parent, tag, label, value) {
        var row = document.createElement(tag || 'p');
        if (label) {
            var title = document.createElement('strong');
            title.textContent = label + ': ';
            row.appendChild(title);
        }
        row.appendChild(document.createTextNode(
            value === null || value === undefined || value === '' ? 'Chưa có' : String(value)
        ));
        parent.appendChild(row);
    }

    function summaryCard(label, value) {
        var card = document.createElement('div');
        card.className = 'training-report-summary-card';
        var num = document.createElement('strong');
        num.textContent = String(value);
        var lbl = document.createElement('span');
        lbl.textContent = label;
        card.append(num, lbl);
        return card;
    }

    function loadOnce() {
        if (loaded) return;
        loaded = true;
        loadList();
    }

    function loadList() {
        renderErrors(errors, null);
        detail.hidden = true;
        listView.hidden = false;
        TrainingUI.setLoading(listBody, true);
        var params = new URLSearchParams();
        params.set('page', String(listPage));
        params.set('page_size', String(listPageSize));
        params.set('status', 'closed');
        TrainingUI.fetchJson('/api/training/exams?' + params.toString())
            .then(renderList)
            .catch(function (e) { renderErrors(errors, e); })
            .finally(function () { TrainingUI.setLoading(listBody, false); });
    }

    function renderList(payload) {
        clear(listBody);
        clear(pagination);
        var items = (payload.items || []).filter(function (e) {
            return e.finalized_at_ms;
        });
        if (!items.length) {
            var empty = document.createElement('div');
            empty.className = 'training-report-empty';
            var icon = document.createElement('i');
            icon.className = 'fas fa-chart-bar';
            empty.appendChild(icon);
            var msg = document.createElement('p');
            msg.textContent = 'Chưa có kỳ thi nào đã chốt báo cáo.';
            empty.appendChild(msg);
            listBody.appendChild(empty);
            return;
        }
        items.forEach(function (e) {
            var card = document.createElement('div');
            card.className = 'training-report-exam-card';
            card.setAttribute('role', 'button');
            card.setAttribute('tabindex', '0');
            card.addEventListener('click', function () { loadReport(e.id); });
            card.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); loadReport(e.id); }
            });
            var info = document.createElement('div');
            info.className = 'training-report-exam-info';
            var titleEl = document.createElement('strong');
            titleEl.textContent = e.code + ' — ' + e.title;
            info.appendChild(titleEl);
            var meta = document.createElement('span');
            meta.textContent = 'Chốt: ' + TrainingUI.formatTime(e.finalized_at_ms);
            info.appendChild(meta);
            card.appendChild(info);
            card.appendChild(statusBadge(e.status, e.finalized_at_ms));
            listBody.appendChild(card);
        });
        renderPagination(items);
    }

    function renderPagination(items) {
        var totalPages = Math.max(1, Math.ceil(items.length / listPageSize));
        if (totalPages <= 1) return;
        var prev = document.createElement('button');
        prev.type = 'button';
        prev.textContent = 'Trang trước';
        prev.disabled = listPage <= 1;
        prev.addEventListener('click', function () { listPage -= 1; loadList(); });
        var cur = document.createElement('span');
        cur.textContent = 'Trang ' + listPage + ' / ' + totalPages;
        var next = document.createElement('button');
        next.type = 'button';
        next.textContent = 'Trang sau';
        next.disabled = listPage >= totalPages;
        next.addEventListener('click', function () { listPage += 1; loadList(); });
        pagination.append(prev, cur, next);
    }

    function loadReport(examId) {
        renderErrors(errors, null);
        TrainingUI.setLoading(detail, true);
        detail.hidden = false;
        listView.hidden = true;
        TrainingUI.fetchJson('/api/training/exams/' + encodeURIComponent(examId))
            .then(function (exam) {
                return TrainingUI.fetchJson('/api/training/exams/' + encodeURIComponent(examId) + '/report')
                    .then(function (report) { renderReportDetail(exam, report); })
                    .catch(function (e) {
                        if (e.status === 404) {
                            renderReportDetail(exam, null);
                        } else {
                            throw e;
                        }
                    });
            })
            .catch(function (e) { renderErrors(errors, e); })
            .finally(function () { TrainingUI.setLoading(detail, false); });
    }

    function renderReportDetail(exam, report) {
        clear(detail);
        var backBtn = document.createElement('button');
        backBtn.type = 'button';
        backBtn.className = 'training-action training-action-secondary';
        backBtn.textContent = 'Quay lại danh sách';
        backBtn.addEventListener('click', function () {
            detail.hidden = true;
            listView.hidden = false;
            loadList();
        });
        detail.appendChild(backBtn);

        var title = document.createElement('h4');
        title.textContent = 'Báo cáo kỳ thi';
        detail.appendChild(title);

        appendText(detail, 'p', 'Mã kỳ thi', exam.code);
        appendText(detail, 'p', 'Tiêu đề', exam.title);

        var statusRow = document.createElement('p');
        var statusStrong = document.createElement('strong');
        statusStrong.textContent = 'Trạng thái: ';
        statusRow.appendChild(statusStrong);
        statusRow.appendChild(statusBadge(exam.status, exam.finalized_at_ms));
        detail.appendChild(statusRow);

        if (exam.template) {
            appendText(detail, 'p', 'Mẫu đề', exam.template.code + ' — ' + exam.template.title);
        }
        appendText(detail, 'p', 'Nhóm đối tượng', exam.target_audience_code);

        if (!report) {
            var noReport = document.createElement('div');
            noReport.className = 'training-report-empty';
            var icon = document.createElement('i');
            icon.className = 'fas fa-exclamation-triangle';
            noReport.appendChild(icon);
            var msg = document.createElement('p');
            msg.textContent = 'Kỳ thi chưa được chốt báo cáo.';
            noReport.appendChild(msg);

            var finalizeRow = document.createElement('div');
            finalizeRow.className = 'training-action-row';
            var finalizeBtn = document.createElement('button');
            finalizeBtn.type = 'button';
            finalizeBtn.className = 'training-action';
            finalizeBtn.textContent = 'Chốt báo cáo';
            finalizeBtn.addEventListener('click', function () { runFinalize(exam, finalizeBtn); });
            finalizeRow.appendChild(finalizeBtn);
            noReport.appendChild(finalizeRow);
            detail.appendChild(noReport);
            return;
        }

        var payload = report.payload || {};
        var summary = payload.summary || {};

        var meta = document.createElement('div');
        meta.className = 'training-report-meta';
        meta.innerHTML = '<strong>Phiên bản:</strong> ' + report.revision +
            ' &nbsp;|&nbsp; <strong>Chốt lúc:</strong> ' + TrainingUI.formatTime(exam.finalized_at_ms) +
            ' &nbsp;|&nbsp; <strong>Bởi:</strong> ' + (exam.finalized_by || '—');
        detail.appendChild(meta);

        var grid = document.createElement('div');
        grid.className = 'training-report-summary-grid';
        grid.appendChild(summaryCard('Được giao', summary.assigned || 0));
        grid.appendChild(summaryCard('Đã làm', summary.completed || 0));
        grid.appendChild(summaryCard('Chưa làm', (summary.assigned || 0) - (summary.completed || 0) - (summary.expired || 0)));
        grid.appendChild(summaryCard('Hết hạn', summary.expired || 0));
        grid.appendChild(summaryCard('Đạt', summary.passed || 0));
        grid.appendChild(summaryCard('Không đạt', summary.failed || 0));
        if (summary.completed > 0) {
            var avgPercent = Math.round(
                (payload.individual || []).reduce(function (sum, row) {
                    return sum + (row.result ? row.result.percent : 0);
                }, 0) / summary.completed
            );
            grid.appendChild(summaryCard('Điểm TB', avgPercent + '%'));
        }
        detail.appendChild(grid);

        var exportRow = document.createElement('div');
        exportRow.className = 'training-action-row';
        var exportBtn = document.createElement('button');
        exportBtn.type = 'button';
        exportBtn.className = 'training-action';
        exportBtn.textContent = 'Tải Excel';
        exportBtn.addEventListener('click', function () {
            window.location.href = '/download/training/exams/' + encodeURIComponent(exam.id) + '/report.xlsx';
        });
        exportRow.appendChild(exportBtn);
        detail.appendChild(exportRow);

        var individual = payload.individual || [];
        if (individual.length) {
            var tableHeading = document.createElement('h5');
            tableHeading.textContent = 'Kết quả từng người';
            detail.appendChild(tableHeading);

            var table = document.createElement('table');
            table.className = 'training-report-table';
            var thead = document.createElement('thead');
            thead.innerHTML = '<tr><th>Tài khoản</th><th>Họ tên</th><th>Điểm</th><th>Tỷ lệ</th><th>Trạng thái</th></tr>';
            table.appendChild(thead);
            var tbody = document.createElement('tbody');
            individual.forEach(function (row) {
                var assignment = row.assignment || {};
                var result = row.result;
                var tr = document.createElement('tr');
                var cells = [
                    { label: 'Tài khoản', value: assignment.username },
                    { label: 'Họ tên', value: assignment.display_name },
                    { label: 'Điểm', value: result ? (result.raw_score + '/' + result.maximum_score) : '—' },
                    { label: 'Tỷ lệ', value: result ? (result.percent + '%') : '—' },
                    { label: 'Trạng thái', value: '' },
                ];
                cells.forEach(function (c, idx) {
                    var td = document.createElement('td');
                    td.dataset.label = c.label;
                    if (idx === 4 && result) {
                        var passSpan = document.createElement('span');
                        passSpan.className = result.passed ? 'training-pass' : 'training-fail';
                        passSpan.textContent = result.passed ? 'Đạt' : 'Không đạt';
                        td.appendChild(passSpan);
                    } else {
                        td.textContent = c.value || '—';
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
            table.appendChild(tbody);
            detail.appendChild(table);
        }

        detail.focus();
    }

    function runFinalize(exam, btn) {
        TrainingUI.confirm('Chốt kỳ thi này? Sau khi chốt sẽ tạo báo cáo không thể sửa.').then(function (ok) {
            if (!ok) return;
            btn.disabled = true;
            TrainingUI.setLoading(btn, true);
            var url = '/api/training/exams/' + encodeURIComponent(exam.id) + '/finalize';
            TrainingUI.fetchJson(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })
                .then(function (resp) {
                    var note = 'Đã chốt kỳ thi';
                    if (resp.revision) note += ' — Phiên bản ' + resp.revision;
                    TrainingUI.toast(note + '.', 'success');
                    loadReport(exam.id);
                })
                .catch(function (e) {
                    if (e.details && e.details.blocking_attempts) {
                        renderErrors(errors, {
                            message: 'Không thể chốt báo cáo khi còn bài làm lỗi toàn vẹn.',
                            details: { errors: e.details.blocking_attempts.map(function (a) {
                                return 'Bài ' + a.attempt_id + ' — lỗi ' + a.error_code;
                            }) },
                        });
                    } else {
                        renderErrors(errors, e);
                    }
                })
                .finally(function () {
                    btn.disabled = false;
                    TrainingUI.setLoading(btn, false);
                });
        });
    }

    panel.addEventListener('training:panel-show', loadOnce);
    if (!panel.hidden) loadOnce();
}(window, document));
