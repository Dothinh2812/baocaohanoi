(function (window, document) {
    'use strict';

    var panel = document.getElementById('training-exams');
    if (!panel || !window.TrainingUI) return;

    var audiences = [];
    try { audiences = JSON.parse(panel.dataset.audiences || '[]'); } catch (e) { audiences = []; }

    var createBtn = document.getElementById('exam-create-btn');
    var listBtn = document.getElementById('exam-list-btn');
    var createForm = document.getElementById('exam-create-form');
    var listView = document.getElementById('exam-list-view');
    var detail = document.getElementById('exam-detail');
    var errors = document.getElementById('exam-errors');
    var listBody = document.getElementById('exam-list-body');
    var listPagination = document.getElementById('exam-pagination');
    var templateSelect = document.getElementById('exam-template-select');
    var audienceSelect = document.getElementById('exam-audience');

    var listPage = 1;
    var listPageSize = 25;
    var loaded = false;
    var currentTemplate = null;
    var currentExamId = null;
    var selectedUsers = {};

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

    function audienceLabel(code) {
        for (var i = 0; i < audiences.length; i += 1) {
            if (audiences[i].code === code) return audiences[i].name;
        }
        return code;
    }

    function statusLabel(s) {
        return {
            draft: 'Nháp', ready: 'Sẵn sàng', open: 'Đang mở',
            closed: 'Đã đóng', cancelled: 'Đã hủy',
        }[s] || s;
    }

    function statusBadge(s, finalizedAtMs) {
        var span = document.createElement('span');
        var cls = s;
        var text = statusLabel(s);
        if (s === 'closed' && finalizedAtMs) {
            cls = 'finalized';
            text = 'Đã chốt';
        }
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

    function populateAudiences() {
        clear(audienceSelect);
        audiences.forEach(function (a) {
            var opt = document.createElement('option');
            opt.value = a.code;
            opt.textContent = a.name;
            audienceSelect.appendChild(opt);
        });
    }

    function loadOnce() {
        if (loaded) return;
        loaded = true;
        populateAudiences();
        loadTemplates();
        loadList();
    }

    function loadTemplates() {
        TrainingUI.fetchJson('/api/training/templates?page=1&page_size=100')
            .then(function (res) {
                var first = templateSelect.options[0] ? templateSelect.options[0].cloneNode(true) : null;
                clear(templateSelect);
                if (first) templateSelect.appendChild(first);
                (res.items || []).forEach(function (t) {
                    var opt = document.createElement('option');
                    opt.value = t.id;
                    opt.textContent = t.code + ' — ' + t.title;
                    templateSelect.appendChild(opt);
                });
            })
            .catch(function (e) { renderErrors(errors, e); });
    }

    function onTemplateChange() {
        var id = templateSelect.value;
        if (!id) {
            currentTemplate = null;
            audienceSelect.value = '';
            return;
        }
        TrainingUI.fetchJson('/api/training/templates/' + encodeURIComponent(id))
            .then(function (t) {
                currentTemplate = t;
                audienceSelect.value = t.target_audience_code;
                document.getElementById('exam-duration').value =
                    Math.round((t.duration_seconds || 0) / 60) || 30;
                document.getElementById('exam-pass-score').value = t.pass_score_percent;
            })
            .catch(function (e) { renderErrors(errors, e); });
    }

    function showCreate() {
        renderErrors(errors, null);
        createForm.hidden = false;
        listView.hidden = true;
        detail.hidden = true;
        if (audienceSelect.options.length === 0) populateAudiences();
        if (templateSelect.options.length <= 1) loadTemplates();
    }

    function showList() {
        renderErrors(errors, null);
        createForm.hidden = true;
        detail.hidden = true;
        listView.hidden = false;
        listPage = 1;
        loadList();
    }

    function renderList(payload) {
        clear(listBody);
        clear(listPagination);
        if (!payload.items.length) {
            var row = document.createElement('tr');
            var td = document.createElement('td');
            td.colSpan = 7;
            td.className = 'training-empty-state';
            td.textContent = 'Chưa có kỳ thi nào.';
            row.appendChild(td);
            listBody.appendChild(row);
            return;
        }
        payload.items.forEach(function (e) {
            var tr = document.createElement('tr');
            var codeCell = document.createElement('td');
            codeCell.dataset.label = 'Mã';
            codeCell.textContent = e.code;
            tr.appendChild(codeCell);
            var titleCell = document.createElement('td');
            titleCell.dataset.label = 'Tiêu đề';
            titleCell.textContent = e.title;
            tr.appendChild(titleCell);
            var statusCell = document.createElement('td');
            statusCell.dataset.label = 'Trạng thái';
            statusCell.appendChild(statusBadge(e.status, e.finalized_at_ms));
            tr.appendChild(statusCell);
            var audCell = document.createElement('td');
            audCell.dataset.label = 'Nhóm đối tượng';
            audCell.textContent = audienceLabel(e.target_audience_code);
            tr.appendChild(audCell);
            var startCell = document.createElement('td');
            startCell.dataset.label = 'Bắt đầu';
            startCell.textContent = TrainingUI.formatTime(e.start_at_ms);
            tr.appendChild(startCell);
            var endCell = document.createElement('td');
            endCell.dataset.label = 'Kết thúc';
            endCell.textContent = TrainingUI.formatTime(e.end_at_ms);
            tr.appendChild(endCell);
            var action = document.createElement('td');
            action.dataset.label = 'Thao tác';
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'training-action training-action-secondary';
            btn.textContent = 'Xem';
            btn.addEventListener('click', function () { loadDetail(e.id); });
            action.appendChild(btn);
            tr.appendChild(action);
            listBody.appendChild(tr);
        });
        renderListPagination(payload);
    }

    function renderListPagination(payload) {
        var totalPages = Math.max(1, Math.ceil(payload.total / payload.page_size));
        if (totalPages === 1) return;
        var prev = document.createElement('button');
        prev.type = 'button';
        prev.textContent = 'Trang trước';
        prev.disabled = payload.page <= 1;
        prev.addEventListener('click', function () { listPage -= 1; loadList(); });
        var cur = document.createElement('span');
        cur.textContent = 'Trang ' + payload.page + ' / ' + totalPages;
        var next = document.createElement('button');
        next.type = 'button';
        next.textContent = 'Trang sau';
        next.disabled = payload.page >= totalPages;
        next.addEventListener('click', function () { listPage += 1; loadList(); });
        listPagination.append(prev, cur, next);
    }

    function loadList() {
        renderErrors(errors, null);
        TrainingUI.setLoading(listBody, true);
        var params = new URLSearchParams();
        params.set('page', String(listPage));
        params.set('page_size', String(listPageSize));
        TrainingUI.fetchJson('/api/training/exams?' + params.toString())
            .then(renderList)
            .catch(function (e) { renderErrors(errors, e); })
            .finally(function () { TrainingUI.setLoading(listBody, false); });
    }

    function submitCreate() {
        renderErrors(errors, null);
        var code = document.getElementById('exam-code').value.trim();
        var title = document.getElementById('exam-title').value.trim();
        var templateId = templateSelect.value;
        var audience = currentTemplate ? currentTemplate.target_audience_code : '';
        var description = document.getElementById('exam-description').value.trim();
        var durationMinutes = parseFloat(document.getElementById('exam-duration').value);
        var passScore = parseFloat(document.getElementById('exam-pass-score').value);
        var startLocal = document.getElementById('exam-start').value;
        var endLocal = document.getElementById('exam-end').value;
        var reveal = document.getElementById('exam-reveal').checked;

        var localErrors = [];
        if (!code) localErrors.push('Mã kỳ thi là bắt buộc.');
        if (!title) localErrors.push('Tiêu đề là bắt buộc.');
        if (!templateId) localErrors.push('Phải chọn mẫu đề.');
        if (!audience) localErrors.push('Mẫu đề chưa xác định nhóm đối tượng.');
        if (!startLocal) localErrors.push('Thời gian bắt đầu là bắt buộc.');
        if (!endLocal) localErrors.push('Thời gian kết thúc là bắt buộc.');
        var startMs = startLocal ? new Date(startLocal).getTime() : NaN;
        var endMs = endLocal ? new Date(endLocal).getTime() : NaN;
        if (!isNaN(startMs) && !isNaN(endMs) && !(startMs < endMs)) {
            localErrors.push('Thời gian kết thúc phải sau thời gian bắt đầu.');
        }
        if (localErrors.length) {
            renderErrors(errors, { message: 'Vui lòng kiểm tra lại.', details: { errors: localErrors } });
            return;
        }

        var body = {
            code: code,
            title: title,
            template_id: templateId,
            target_audience_code: audience,
            start_at_ms: startMs,
            end_at_ms: endMs,
            duration_seconds: Math.max(1, Math.round((isNaN(durationMinutes) ? 30 : durationMinutes) * 60)),
            pass_score_percent: isNaN(passScore) ? 80 : passScore,
            description: description || null,
            reveal_answers_after_finalize: reveal,
        };
        var submitBtnEl = document.getElementById('exam-submit-btn');
        TrainingUI.setLoading(submitBtnEl, true);
        submitBtnEl.disabled = true;
        TrainingUI.fetchJson('/api/training/exams', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        }).then(function () {
            TrainingUI.toast('Đã tạo kỳ thi.', 'success');
            createForm.hidden = true;
            listView.hidden = false;
            listPage = 1;
            loadList();
        }).catch(function (e) { renderErrors(errors, e); })
          .finally(function () {
              TrainingUI.setLoading(submitBtnEl, false);
              submitBtnEl.disabled = false;
          });
    }

    function loadDetail(id) {
        currentExamId = id;
        selectedUsers = {};
        renderErrors(errors, null);
        TrainingUI.setLoading(detail, true);
        Promise.all([
            TrainingUI.fetchJson('/api/training/exams/' + encodeURIComponent(id)),
            TrainingUI.fetchJson('/api/training/exams/' + encodeURIComponent(id) + '/assignments'),
        ]).then(function (results) {
            renderDetail(results[0], results[1].items || []);
        }).catch(function (e) { renderErrors(errors, e); })
          .finally(function () { TrainingUI.setLoading(detail, false); });
    }

    function summaryCard(label, value) {
        var card = document.createElement('div');
        card.className = 'training-summary-card';
        var num = document.createElement('strong');
        num.textContent = String(value);
        var lbl = document.createElement('span');
        lbl.textContent = label;
        card.append(num, lbl);
        return card;
    }

    function actionButton(label, handler) {
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'training-action';
        btn.textContent = label;
        btn.addEventListener('click', function () { handler(btn); });
        return btn;
    }

    function runTransition(exam, action, confirmMsg, btn) {
        TrainingUI.confirm(confirmMsg).then(function (ok) {
            if (!ok) return;
            performTransition(exam, action, btn);
        });
    }

    function performTransition(exam, action, btn) {
        if (btn) { btn.disabled = true; TrainingUI.setLoading(btn, true); }
        var url = '/api/training/exams/' + encodeURIComponent(exam.id) + '/' + action;
        TrainingUI.fetchJson(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })
            .then(function (resp) {
                if (action === 'close' && resp.recovery_summary) {
                    renderRecovery(resp.recovery_summary);
                }
                TrainingUI.toast('Đã cập nhật kỳ thi.', 'success');
                listPage = 1;
                loadList();
            })
            .catch(function (e) {
                if (e.details && e.details.blocking_attempts) {
                    renderBlocking(e.details);
                } else {
                    renderErrors(errors, e);
                }
            })
            .finally(function () {
                loadDetail(exam.id);
                if (btn) { btn.disabled = false; TrainingUI.setLoading(btn, false); }
            });
    }

    function runFinalize(exam, btn) {
        TrainingUI.confirm('Chốt kỳ thi này? Sau khi chốt sẽ tạo báo cáo không thể sửa.').then(function (ok) {
            if (!ok) return;
            if (btn) { btn.disabled = true; TrainingUI.setLoading(btn, true); }
            var url = '/api/training/exams/' + encodeURIComponent(exam.id) + '/finalize';
            TrainingUI.fetchJson(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })
                .then(function (resp) {
                    var note = 'Đã chốt kỳ thi';
                    if (resp.revision) note += ' — Phiên bản ' + resp.revision;
                    TrainingUI.toast(note + '.', 'success');
                    listPage = 1;
                    loadList();
                })
                .catch(function (e) {
                    if (e.details && e.details.blocking_attempts) {
                        renderBlocking(e.details);
                    } else {
                        renderErrors(errors, e);
                    }
                })
                .finally(function () {
                    loadDetail(exam.id);
                    if (btn) { btn.disabled = false; TrainingUI.setLoading(btn, false); }
                });
        });
    }

    function renderRecovery(summary) {
        var box = document.createElement('div');
        box.className = 'training-form-errors';
        box.style.color = '#0b5394';
        box.style.background = '#e8f0fe';
        box.style.borderColor = '#c2dbff';
        var head = document.createElement('strong');
        head.textContent = 'Kết quả xử lý khi đóng kỳ thi';
        box.appendChild(head);
        appendText(box, 'p', 'Đã xử lý', (summary.processed_attempt_ids || []).length + ' bài');
        appendText(box, 'p', 'Đã hoàn tất trước đó', (summary.already_completed_ids || []).length + ' bài');
        appendText(box, 'p', 'Xử lý thất bại', (summary.failed_attempts || []).length + ' bài');
        detail.insertBefore(box, detail.firstChild);
    }

    function renderBlocking(details) {
        renderErrors(errors, {
            message: 'Không thể chốt báo cáo khi còn bài làm lỗi toàn vẹn.',
            details: { errors: (details.blocking_attempts || []).map(function (a) {
                return 'Bài ' + a.attempt_id + ' — lỗi ' + a.error_code;
            }) },
        });
        detail.hidden = false;
    }

    function renderLifecycle(exam, container) {
        var finalized = exam.status === 'closed' && exam.finalized_at_ms;
        if (finalized) {
            var note = document.createElement('p');
            note.className = 'training-empty-state';
            note.textContent = 'Kỳ thi đã chốt';
            container.appendChild(note);
            if (exam.finalized_at_ms) {
                TrainingUI.fetchJson('/api/training/exams/' + encodeURIComponent(exam.id) + '/report')
                    .then(function (rep) {
                        if (rep && rep.revision) {
                            note.textContent = 'Đã chốt — Phiên bản ' + rep.revision;
                        }
                    })
                    .catch(function () { /* chưa có báo cáo — giữ nhãn mặc định */ });
            }
            return;
        }
        if (exam.status === 'cancelled') {
            var cancelled = document.createElement('p');
            cancelled.className = 'training-empty-state';
            cancelled.textContent = 'Kỳ thi đã hủy';
            container.appendChild(cancelled);
            return;
        }
        if (exam.status === 'draft') {
            container.appendChild(actionButton('Sẵn sàng', function (btn) {
                runTransition(exam, 'ready', 'Chuyển kỳ thi sang trạng thái Sẵn sàng?', btn);
            }));
            container.appendChild(actionButton('Hủy', function (btn) {
                runTransition(exam, 'cancel', 'Hủy kỳ thi này?', btn);
            }));
        } else if (exam.status === 'ready') {
            container.appendChild(actionButton('Mở', function (btn) {
                runTransition(exam, 'open', 'Mở kỳ thi để người học làm bài?', btn);
            }));
            container.appendChild(actionButton('Hủy', function (btn) {
                runTransition(exam, 'cancel', 'Hủy kỳ thi này?', btn);
            }));
        } else if (exam.status === 'open') {
            container.appendChild(actionButton('Đóng', function (btn) {
                runTransition(exam, 'close',
                    'Đóng kỳ thi? Các bài đang làm sẽ được nộp tự động.', btn);
            }));
        } else if (exam.status === 'closed') {
            container.appendChild(actionButton('Chốt', function (btn) { runFinalize(exam, btn); }));
        }
    }

    function renderAssignmentTable(assignments) {
        if (!assignments.length) {
            var empty = document.createElement('p');
            empty.className = 'training-empty-state';
            empty.textContent = 'Chưa có người được giao.';
            return empty;
        }
        var table = document.createElement('table');
        table.className = 'training-question-table';
        var head = document.createElement('thead');
        head.innerHTML = '<tr><th>Tài khoản</th><th>Họ tên</th><th>Đội</th><th>Trạng thái giao</th><th>Bài làm</th></tr>';
        table.appendChild(head);
        var body = document.createElement('tbody');
        assignments.forEach(function (a) {
            var tr = document.createElement('tr');
            var cells = [
                { label: 'Tài khoản', value: a.username },
                { label: 'Họ tên', value: a.display_name },
                { label: 'Đội', value: a.team_name || a.team_code || '' },
                { label: 'Trạng thái giao', value: a.status },
                { label: 'Bài làm', value: a.attempt_status || 'Chưa làm' },
            ];
            cells.forEach(function (c) {
                var td = document.createElement('td');
                td.dataset.label = c.label;
                td.textContent = c.value || '—';
                tr.appendChild(td);
            });
            body.appendChild(tr);
        });
        table.appendChild(body);
        return table;
    }

    function renderAssignmentSection(exam) {
        if (exam.status !== 'draft' && exam.status !== 'ready') return null;
        var section = document.createElement('div');
        section.className = 'training-question-detail';
        var heading = document.createElement('h4');
        heading.textContent = 'Giao bài';
        section.appendChild(heading);

        var searchLabel = document.createElement('label');
        searchLabel.textContent = 'Tìm người dùng ';
        var searchInput = document.createElement('input');
        searchInput.type = 'search';
        searchInput.placeholder = 'Nhập tài khoản hoặc họ tên...';
        searchInput.autocomplete = 'off';
        searchLabel.appendChild(searchInput);
        section.appendChild(searchLabel);

        var results = document.createElement('div');
        results.className = 'training-user-results';
        section.appendChild(results);

        var countPara = document.createElement('p');
        countPara.className = 'training-selected-count';
        countPara.textContent = 'Đã chọn: 0 người';
        section.appendChild(countPara);

        var assignBtn = document.createElement('button');
        assignBtn.type = 'button';
        assignBtn.className = 'training-action';
        assignBtn.textContent = 'Giao bài';
        section.appendChild(assignBtn);

        var debounce = null;
        function runSearch() {
            var q = searchInput.value.trim();
            TrainingUI.fetchJson('/api/training/users?q=' + encodeURIComponent(q))
                .then(function (res) { renderUserResults(res.items || [], results, countPara); })
                .catch(function (e) { renderErrors(errors, e); });
        }
        searchInput.addEventListener('input', function () {
            if (debounce) window.clearTimeout(debounce);
            debounce = window.setTimeout(runSearch, 300);
        });
        assignBtn.addEventListener('click', function () {
            submitAssignments(exam, assignBtn);
        });
        runSearch();
        return section;
    }

    function renderUserResults(users, container, countPara) {
        clear(container);
        if (!users.length) {
            var empty = document.createElement('p');
            empty.className = 'training-empty-state';
            empty.textContent = 'Không tìm thấy người dùng.';
            container.appendChild(empty);
            return;
        }
        var listEl = document.createElement('ul');
        listEl.className = 'training-user-list';
        users.forEach(function (u) {
            var li = document.createElement('li');
            var label = document.createElement('label');
            var cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = u.username;
            cb.checked = Object.prototype.hasOwnProperty.call(selectedUsers, u.username);
            cb.addEventListener('change', function () {
                if (cb.checked) selectedUsers[u.username] = u.display_name;
                else delete selectedUsers[u.username];
                countPara.textContent = 'Đã chọn: ' + Object.keys(selectedUsers).length + ' người';
            });
            label.appendChild(cb);
            label.appendChild(document.createTextNode(' ' + u.username + ' — ' + u.display_name));
            li.appendChild(label);
            listEl.appendChild(li);
        });
        container.appendChild(listEl);
        countPara.textContent = 'Đã chọn: ' + Object.keys(selectedUsers).length + ' người';
    }

    function submitAssignments(exam, btn) {
        var usernames = Object.keys(selectedUsers);
        if (!usernames.length) {
            renderErrors(errors, { message: 'Chưa chọn người để giao bài.' });
            return;
        }
        var users = usernames.map(function (un) {
            return { username: un, display_name: selectedUsers[un] };
        });
        TrainingUI.setLoading(btn, true);
        btn.disabled = true;
        TrainingUI.fetchJson('/api/training/exams/' + encodeURIComponent(exam.id) + '/assignments', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ users: users, audience_code: exam.target_audience_code }),
        }).then(function (resp) {
            TrainingUI.toast('Đã giao ' + (resp.assignment_ids || []).length + ' bài.', 'success');
            selectedUsers = {};
            loadDetail(exam.id);
        }).catch(function (e) {
            renderErrors(errors, e);
        }).finally(function () {
            TrainingUI.setLoading(btn, false);
            btn.disabled = false;
        });
    }

    function renderDetail(exam, assignments) {
        clear(detail);
        detail.hidden = false;
        listView.hidden = true;
        createForm.hidden = true;

        var title = document.createElement('h4');
        title.textContent = 'Chi tiết kỳ thi';
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
        appendText(detail, 'p', 'Nhóm đối tượng', audienceLabel(exam.target_audience_code));
        appendText(detail, 'p', 'Thời gian làm bài',
            TrainingUI.formatTime(exam.start_at_ms) + ' → ' + TrainingUI.formatTime(exam.end_at_ms));
        appendText(detail, 'p', 'Thời lượng', Math.round((exam.duration_seconds || 0) / 60) + ' phút');
        appendText(detail, 'p', 'Điểm đạt', exam.pass_score_percent + '%');
        appendText(detail, 'p', 'Công bố đáp án sau chốt', exam.reveal_answers_after_finalize ? 'Có' : 'Không');

        var summaryHeading = document.createElement('h5');
        summaryHeading.textContent = 'Tổng quan giao bài';
        detail.appendChild(summaryHeading);
        var grid = document.createElement('div');
        grid.className = 'training-summary-grid';
        var s = exam.assignment_summary || {};
        grid.appendChild(summaryCard('Tổng', s.total || 0));
        grid.appendChild(summaryCard('Đã giao', s.assigned || 0));
        grid.appendChild(summaryCard('Đang làm', s.in_progress || 0));
        grid.appendChild(summaryCard('Hoàn thành', s.completed || 0));
        grid.appendChild(summaryCard('Hết hạn', s.expired || 0));
        grid.appendChild(summaryCard('Đã hủy', s.cancelled || 0));
        detail.appendChild(grid);

        var lifecycle = document.createElement('div');
        lifecycle.className = 'training-action-row';
        renderLifecycle(exam, lifecycle);
        if (lifecycle.childNodes.length) detail.appendChild(lifecycle);

        var assignmentsHeading = document.createElement('h5');
        assignmentsHeading.textContent = 'Danh sách người được giao';
        detail.appendChild(assignmentsHeading);
        detail.appendChild(renderAssignmentTable(assignments));

        var assignSection = renderAssignmentSection(exam);
        if (assignSection) detail.appendChild(assignSection);

        var back = document.createElement('button');
        back.type = 'button';
        back.className = 'training-action training-action-secondary';
        back.textContent = 'Quay lại danh sách';
        back.addEventListener('click', showList);
        detail.appendChild(back);
        detail.focus();
    }

    createBtn.addEventListener('click', showCreate);
    listBtn.addEventListener('click', showList);
    document.getElementById('exam-submit-btn').addEventListener('click', submitCreate);
    document.getElementById('exam-cancel-btn').addEventListener('click', showList);
    templateSelect.addEventListener('change', onTemplateChange);

    panel.addEventListener('training:panel-show', loadOnce);
    if (!panel.hidden) loadOnce();
}(window, document));
