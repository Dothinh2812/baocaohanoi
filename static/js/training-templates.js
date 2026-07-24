(function (window, document) {
    'use strict';

    var panel = document.getElementById('training-templates');
    if (!panel || !window.TrainingUI) return;

    var canEdit = panel.dataset.canEdit === 'true';
    var audiences = [];
    var domains = [];
    try { audiences = JSON.parse(panel.dataset.audiences || '[]'); } catch (e) { audiences = []; }
    try { domains = JSON.parse(panel.dataset.domains || '[]'); } catch (e) { domains = []; }

    var createBtn = document.getElementById('template-create-btn');
    var listBtn = document.getElementById('template-list-btn');
    var createForm = document.getElementById('template-create-form');
    var listView = document.getElementById('template-list-view');
    var detail = document.getElementById('template-detail');
    var errors = document.getElementById('template-errors');
    var questionList = document.getElementById('template-question-list');
    var questionPagination = document.getElementById('template-question-pagination');
    var questionSearch = document.getElementById('template-question-search');
    var audienceSelect = document.getElementById('template-audience');
    var domainSelect = document.getElementById('template-question-domain');
    var selectedCount = document.getElementById('template-selected-count');
    var submitBtn = document.getElementById('template-submit-btn');
    var cancelBtn = document.getElementById('template-cancel-btn');
    var listBody = document.getElementById('template-list-body');
    var listPagination = document.getElementById('template-pagination');

    var selectedQuestions = new Set();
    var questionPage = 1;
    var questionPageSize = 25;
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

    function audienceLabel(code) {
        var found = null;
        for (var i = 0; i < audiences.length; i += 1) {
            if (audiences[i].code === code) { found = audiences[i]; break; }
        }
        return found ? found.name : code;
    }

    function populateSelects() {
        clear(audienceSelect);
        audiences.forEach(function (a) {
            var opt = document.createElement('option');
            opt.value = a.code;
            opt.textContent = a.name;
            audienceSelect.appendChild(opt);
        });
        clear(domainSelect);
        var allOpt = document.createElement('option');
        allOpt.value = '';
        allOpt.textContent = 'Tất cả';
        domainSelect.appendChild(allOpt);
        domains.forEach(function (d) {
            var opt = document.createElement('option');
            opt.value = d.code;
            opt.textContent = d.name;
            domainSelect.appendChild(opt);
        });
    }

    function updateSelectedCount() {
        selectedCount.textContent = 'Đã chọn: ' + selectedQuestions.size + ' câu';
    }

    function showCreate() {
        renderErrors(errors, null);
        createForm.hidden = false;
        listView.hidden = true;
        detail.hidden = true;
        if (audienceSelect.options.length === 0) populateSelects();
        questionPage = 1;
        loadQuestions();
    }

    function showList() {
        renderErrors(errors, null);
        createForm.hidden = true;
        detail.hidden = true;
        listView.hidden = false;
        listPage = 1;
        loadList();
    }

    function questionParams() {
        var params = new URLSearchParams();
        params.set('status', 'published');
        params.set('page', String(questionPage));
        params.set('page_size', String(questionPageSize));
        if (audienceSelect.value) params.set('audience', audienceSelect.value);
        var q = document.getElementById('template-question-q').value.trim();
        if (q) params.set('q', q);
        if (domainSelect.value) params.set('domain', domainSelect.value);
        var topic = document.getElementById('template-question-topic').value.trim();
        if (topic) params.set('topic', topic);
        return params;
    }

    function renderQuestionList(payload) {
        clear(questionList);
        clear(questionPagination);
        if (!payload.items.length) {
            var empty = document.createElement('p');
            empty.className = 'training-empty-state';
            empty.textContent = 'Không có câu hỏi đã phát hành phù hợp.';
            questionList.appendChild(empty);
            updateSelectedCount();
            return;
        }
        var table = document.createElement('table');
        table.className = 'training-question-table';
        var head = document.createElement('thead');
        head.innerHTML = '<tr><th>Chọn</th><th>Câu hỏi</th><th>Loại</th><th>Độ khó</th><th>Đối tượng</th><th>Chủ đề</th></tr>';
        table.appendChild(head);
        var body = document.createElement('tbody');
        payload.items.forEach(function (item) {
            var tr = document.createElement('tr');
            var checkCell = document.createElement('td');
            checkCell.dataset.label = 'Chọn';
            var cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = item.id;
            cb.checked = selectedQuestions.has(item.id);
            cb.addEventListener('change', function () {
                if (cb.checked) selectedQuestions.add(item.id);
                else selectedQuestions.delete(item.id);
                updateSelectedCount();
            });
            checkCell.appendChild(cb);
            tr.appendChild(checkCell);
            var values = [item.stem, item.type, item.difficulty,
                          (item.audience || []).join(', '), (item.topic || []).join(', ')];
            var labels = ['Câu hỏi', 'Loại', 'Độ khó', 'Đối tượng', 'Chủ đề'];
            values.forEach(function (v, i) {
                var td = document.createElement('td');
                td.dataset.label = labels[i];
                td.textContent = v || 'Chưa phân loại';
                tr.appendChild(td);
            });
            body.appendChild(tr);
        });
        table.appendChild(body);
        questionList.appendChild(table);
        renderQuestionPagination(payload);
        updateSelectedCount();
    }

    function renderQuestionPagination(payload) {
        var totalPages = Math.max(1, Math.ceil(payload.total / payload.page_size));
        if (totalPages === 1) return;
        var prev = document.createElement('button');
        prev.type = 'button';
        prev.textContent = 'Trang trước';
        prev.disabled = payload.page <= 1;
        prev.addEventListener('click', function () { questionPage -= 1; loadQuestions(); });
        var cur = document.createElement('span');
        cur.textContent = 'Trang ' + payload.page + ' / ' + totalPages;
        var next = document.createElement('button');
        next.type = 'button';
        next.textContent = 'Trang sau';
        next.disabled = payload.page >= totalPages;
        next.addEventListener('click', function () { questionPage += 1; loadQuestions(); });
        questionPagination.append(prev, cur, next);
    }

    function loadQuestions() {
        renderErrors(errors, null);
        TrainingUI.setLoading(questionList, true);
        TrainingUI.fetchJson('/api/training/questions?' + questionParams().toString())
            .then(renderQuestionList)
            .catch(function (e) { renderErrors(errors, e); })
            .finally(function () { TrainingUI.setLoading(questionList, false); });
    }

    function renderList(payload) {
        clear(listBody);
        clear(listPagination);
        if (!payload.items.length) {
            var row = document.createElement('tr');
            var td = document.createElement('td');
            td.colSpan = 7;
            td.className = 'training-empty-state';
            td.textContent = 'Chưa có mẫu đề nào.';
            row.appendChild(td);
            listBody.appendChild(row);
            return;
        }
        payload.items.forEach(function (t) {
            var tr = document.createElement('tr');
            var cells = [
                { label: 'Mã', value: t.code },
                { label: 'Tiêu đề', value: t.title },
                { label: 'Nhóm đối tượng', value: audienceLabel(t.target_audience_code) },
                { label: 'Số câu', value: t.total_questions },
                { label: 'Thời lượng', value: Math.round((t.duration_seconds || 0) / 60) + ' phút' },
                { label: 'Trạng thái', value: t.locked ? 'Đã khóa' : 'Có thể sửa' },
            ];
            cells.forEach(function (c) {
                var td = document.createElement('td');
                td.dataset.label = c.label;
                td.textContent = c.value;
                tr.appendChild(td);
            });
            var action = document.createElement('td');
            action.dataset.label = 'Thao tác';
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'training-action training-action-secondary';
            btn.textContent = 'Xem';
            btn.addEventListener('click', function () { loadDetail(t.id); });
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
        TrainingUI.fetchJson('/api/training/templates?' + params.toString())
            .then(renderList)
            .catch(function (e) { renderErrors(errors, e); })
            .finally(function () { TrainingUI.setLoading(listBody, false); });
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

    function loadDetail(id) {
        renderErrors(errors, null);
        TrainingUI.setLoading(detail, true);
        TrainingUI.fetchJson('/api/training/templates/' + encodeURIComponent(id))
            .then(function (t) {
                clear(detail);
                detail.hidden = false;
                listView.hidden = true;
                createForm.hidden = true;
                var title = document.createElement('h4');
                title.textContent = 'Chi tiết mẫu đề';
                detail.appendChild(title);
                appendText(detail, 'p', 'Mã đề', t.code);
                appendText(detail, 'p', 'Tiêu đề', t.title);
                appendText(detail, 'p', 'Nhóm đối tượng', audienceLabel(t.target_audience_code));
                appendText(detail, 'p', 'Số câu hỏi', t.total_questions);
                appendText(detail, 'p', 'Thời lượng', Math.round((t.duration_seconds || 0) / 60) + ' phút');
                appendText(detail, 'p', 'Điểm đạt', t.pass_score_percent + '%');
                appendText(detail, 'p', 'Trộn câu hỏi / đáp án',
                    (t.shuffle_questions ? 'Có' : 'Không') + ' / ' + (t.shuffle_options ? 'Có' : 'Không'));
                appendText(detail, 'p', 'Trạng thái', t.locked ? 'Đã khóa' : 'Có thể sửa');
                appendText(detail, 'p', 'Người tạo', t.created_by);
                appendText(detail, 'p', 'Thời gian tạo', TrainingUI.formatTime(t.created_at_ms));
                var heading = document.createElement('h5');
                heading.textContent = 'Danh sách câu hỏi';
                detail.appendChild(heading);
                var listEl = document.createElement('ol');
                listEl.className = 'training-detail-list';
                (t.items || []).forEach(function (it) {
                    var li = document.createElement('li');
                    li.textContent = '[' + it.sequence_number + '] ' + it.stem +
                        ' (' + it.type + ' / ' + it.difficulty + ')';
                    listEl.appendChild(li);
                });
                detail.appendChild(listEl);
                var back = document.createElement('button');
                back.type = 'button';
                back.className = 'training-action training-action-secondary';
                back.textContent = 'Quay lại danh sách';
                back.addEventListener('click', showList);
                detail.appendChild(back);
                detail.focus();
            })
            .catch(function (e) { renderErrors(errors, e); })
            .finally(function () { TrainingUI.setLoading(detail, false); });
    }

    function submitCreate() {
        renderErrors(errors, null);
        var code = document.getElementById('template-code').value.trim();
        var title = document.getElementById('template-title').value.trim();
        var audience = audienceSelect.value;
        var durationMinutes = parseFloat(document.getElementById('template-duration').value);
        var passScore = parseFloat(document.getElementById('template-pass-score').value);

        var localErrors = [];
        if (!code) localErrors.push('Mã đề là bắt buộc.');
        if (!title) localErrors.push('Tiêu đề là bắt buộc.');
        if (!audience) localErrors.push('Nhóm đối tượng là bắt buộc.');
        if (!selectedQuestions.size) localErrors.push('Cần chọn ít nhất một câu hỏi.');
        if (localErrors.length) {
            renderErrors(errors, { message: 'Vui lòng kiểm tra lại.', details: { errors: localErrors } });
            return;
        }

        var body = {
            code: code,
            title: title,
            target_audience_code: audience,
            question_version_ids: Array.from(selectedQuestions),
            duration_seconds: Math.max(1, Math.round((isNaN(durationMinutes) ? 30 : durationMinutes) * 60)),
            pass_score_percent: isNaN(passScore) ? 80 : passScore,
            shuffle_questions: document.getElementById('template-shuffle-questions').checked,
            shuffle_options: document.getElementById('template-shuffle-options').checked,
        };
        TrainingUI.setLoading(submitBtn, true);
        submitBtn.disabled = true;
        TrainingUI.fetchJson('/api/training/templates', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        }).then(function () {
            TrainingUI.toast('Đã tạo mẫu đề.', 'success');
            selectedQuestions.clear();
            updateSelectedCount();
            createForm.hidden = true;
            listView.hidden = false;
            listPage = 1;
            loadList();
        }).catch(function (e) { renderErrors(errors, e); })
          .finally(function () {
              TrainingUI.setLoading(submitBtn, false);
              submitBtn.disabled = false;
          });
    }

    if (canEdit) {
        createBtn.addEventListener('click', showCreate);
        submitBtn.addEventListener('click', submitCreate);
        cancelBtn.addEventListener('click', showList);
        questionSearch.addEventListener('click', function () { questionPage = 1; loadQuestions(); });
        audienceSelect.addEventListener('change', function () { questionPage = 1; loadQuestions(); });
    } else {
        createBtn.disabled = true;
        submitBtn.disabled = true;
    }
    listBtn.addEventListener('click', showList);

    function loadOnce() {
        if (loaded) return;
        loaded = true;
        populateSelects();
        loadList();
    }
    panel.addEventListener('training:panel-show', loadOnce);
    if (!panel.hidden) loadOnce();
}(window, document));
