(function (window, document) {
    'use strict';

    var panel = document.getElementById('training-question-bank');
    if (!panel || !window.TrainingUI) return;

    var filters = document.getElementById('question-bank-filters');
    var list = document.getElementById('question-bank-list');
    var pagination = document.getElementById('question-bank-pagination');
    var detail = document.getElementById('question-bank-detail');
    var errors = document.getElementById('question-bank-errors');
    var textarea = document.getElementById('question-batch-json');
    var batchErrors = document.getElementById('question-batch-errors');
    var clearFiltersButton = document.getElementById('question-bank-clear-filters');
    var canReview = panel.dataset.canReview === 'true';
    var page = 1;
    var pageSize = 25;

    function clear(element) {
        while (element.firstChild) element.removeChild(element.firstChild);
    }

    function renderErrors(element, error) {
        clear(element);
        if (!error) return;
        var heading = document.createElement('strong');
        heading.textContent = error.message || 'Không thể hoàn tất yêu cầu.';
        element.appendChild(heading);
        var messages = error.details && error.details.errors;
        if (Array.isArray(messages) && messages.length) {
            var listElement = document.createElement('ul');
            messages.forEach(function (message) {
                var item = document.createElement('li');
                item.textContent = message;
                listElement.appendChild(item);
            });
            element.appendChild(listElement);
        }
    }

    function appendText(parent, tag, label, value) {
        var row = document.createElement(tag || 'p');
        if (label) {
            var title = document.createElement('strong');
            title.textContent = label + ': ';
            row.appendChild(title);
        }
        row.appendChild(document.createTextNode(value === null || value === undefined || value === '' ? 'Chưa có' : String(value)));
        parent.appendChild(row);
    }

    function selectedFilters() {
        var values = new URLSearchParams();
        Array.prototype.forEach.call(filters.elements, function (input) {
            if (input.name && input.value.trim()) values.set(input.name, input.value.trim());
        });
        values.set('page', String(page));
        values.set('page_size', String(pageSize));
        return values;
    }

    function resetFilters() {
        Array.prototype.forEach.call(filters.elements, function (input) {
            if (input.name) input.value = '';
        });
        page = 1;
    }

    function statusLabel(status) {
        return { draft: 'Nháp', approved: 'Đã duyệt', published: 'Đã phát hành', rejected: 'Từ chối' }[status] || status;
    }

    function renderList(payload) {
        clear(list);
        clear(pagination);
        if (!payload.items.length) {
            var empty = document.createElement('p');
            empty.className = 'training-empty-state';
            empty.textContent = 'Không có câu hỏi phù hợp.';
            list.appendChild(empty);
            return;
        }
        var table = document.createElement('table');
        table.className = 'training-question-table';
        var head = document.createElement('thead');
        head.innerHTML = '<tr><th>Câu hỏi</th><th>Loại</th><th>Độ khó</th><th>Đối tượng</th><th>Chủ đề</th><th>Trạng thái</th><th><span class="sr-only">Chi tiết</span></th></tr>';
        table.appendChild(head);
        var body = document.createElement('tbody');
        payload.items.forEach(function (item) {
            var row = document.createElement('tr');
            [item.stem, item.type, item.difficulty, item.audience.join(', '), item.topic.join(', '), statusLabel(item.status)].forEach(function (value, index) {
                var cell = document.createElement('td');
                cell.dataset.label = ['Câu hỏi', 'Loại', 'Độ khó', 'Đối tượng', 'Chủ đề', 'Trạng thái'][index];
                cell.textContent = value || 'Chưa phân loại';
                row.appendChild(cell);
            });
            var action = document.createElement('td');
            action.dataset.label = 'Thao tác';
            var button = document.createElement('button');
            button.type = 'button';
            button.className = 'training-action training-action-secondary';
            button.textContent = 'Xem';
            button.addEventListener('click', function () { loadDetail(item.id); });
            action.appendChild(button);
            row.appendChild(action);
            body.appendChild(row);
        });
        table.appendChild(body);
        list.appendChild(table);
        renderPagination(payload);
    }

    function renderPagination(payload) {
        var totalPages = Math.max(1, Math.ceil(payload.total / payload.page_size));
        if (totalPages === 1) return;
        var previous = document.createElement('button');
        previous.type = 'button';
        previous.textContent = 'Trang trước';
        previous.disabled = payload.page <= 1;
        previous.addEventListener('click', function () { page -= 1; loadList(); });
        var current = document.createElement('span');
        current.textContent = 'Trang ' + payload.page + ' / ' + totalPages;
        var next = document.createElement('button');
        next.type = 'button';
        next.textContent = 'Trang sau';
        next.disabled = payload.page >= totalPages;
        next.addEventListener('click', function () { page += 1; loadList(); });
        pagination.append(previous, current, next);
    }

    function loadList() {
        renderErrors(errors, null);
        TrainingUI.setLoading(list, true);
        TrainingUI.fetchJson('/api/training/questions?' + selectedFilters().toString())
            .then(renderList)
            .catch(function (error) { renderErrors(errors, error); })
            .finally(function () { TrainingUI.setLoading(list, false); });
    }

    function renderDetail(question) {
        clear(detail);
        detail.hidden = false;
        var title = document.createElement('h4');
        title.textContent = 'Chi tiết câu hỏi';
        detail.appendChild(title);
        appendText(detail, 'p', 'Câu hỏi', question.stem);
        appendText(detail, 'p', 'Tình huống / ngữ cảnh', question.stimulus);
        appendText(detail, 'p', 'Ngôn ngữ', question.language);
        appendText(detail, 'p', 'Phiên bản', question.version);
        appendText(detail, 'p', 'Loại / độ khó', question.type + ' / ' + question.difficulty);
        appendText(detail, 'p', 'Trạng thái', statusLabel(question.publication.status === 'published' ? 'published' : question.review_status));
        appendText(detail, 'p', 'Giải thích', question.explanation);
        appendText(detail, 'p', 'Đáp án đúng', question.correct_option_ids.join(', '));
        appendText(detail, 'p', 'Phân loại', 'Lĩnh vực: ' + question.classification.domain_codes.join(', ') + '; Đối tượng: ' + question.classification.audience_codes.join(', ') + '; Chủ đề: ' + question.classification.topic_codes.join(', ') + '; Chỉ tiêu: ' + question.classification.indicator_codes.join(', '));
        appendText(detail, 'p', 'Nhận thức / mức độ quan trọng', (question.cognitive_level || 'Chưa có') + ' / ' + (question.criticality || 'Chưa có'));
        appendText(detail, 'p', 'Thời gian ước tính', question.estimated_seconds ? question.estimated_seconds + ' giây' : null);
        appendText(detail, 'p', 'Điểm tối đa', question.max_score);
        appendText(detail, 'p', 'Chính sách chấm', question.scoring_policy ? JSON.stringify(question.scoring_policy) : null);
        appendText(detail, 'p', 'Lý do phương án nhiễu', JSON.stringify(question.distractor_rationales));
        appendText(detail, 'p', 'Người tạo / thời gian tạo', question.created_by + ' / ' + TrainingUI.formatTime(question.created_at_ms));
        appendText(detail, 'p', 'Người duyệt / thời gian duyệt', (question.publication.approved_by || 'Chưa có') + ' / ' + TrainingUI.formatTime(question.publication.approved_at_ms));
        appendText(detail, 'p', 'Trạng thái phát hành', question.publication.status);

        var options = document.createElement('ol');
        options.className = 'training-detail-list';
        question.options.forEach(function (option) {
            var item = document.createElement('li');
            item.textContent = option.id + ': ' + option.text;
            options.appendChild(item);
        });
        detail.appendChild(options);

        var evidenceTitle = document.createElement('h5');
        evidenceTitle.textContent = 'Dẫn chứng';
        detail.appendChild(evidenceTitle);
        question.evidence.forEach(function (evidence) {
            appendText(detail, 'p', 'Tài liệu ' + evidence.document_version_id + ' / ' + evidence.block_id, evidence.quoted_text || evidence.supports);
        });
        appendText(detail, 'p', 'Lịch sử duyệt', question.review_history.map(function (item) {
            return item.action + ' - ' + item.reviewer + ' - ' + TrainingUI.formatTime(item.created_at_ms) + (item.comment ? ': ' + item.comment : '');
        }).join('\n') || 'Chưa có');
        appendText(detail, 'p', 'Lịch sử phát hành', question.publication_history.map(function (item) {
            return item.action + ' - ' + item.actor + ' - ' + TrainingUI.formatTime(item.created_at_ms);
        }).join('\n') || 'Chưa phát hành');
        renderReviewActions(question);
        detail.focus();
    }

    function actionButton(label, handler) {
        var button = document.createElement('button');
        button.type = 'button';
        button.className = 'training-action';
        button.textContent = label;
        button.addEventListener('click', handler);
        return button;
    }

    function runReviewAction(question, action, body) {
        TrainingUI.fetchJson('/api/training/questions/' + encodeURIComponent(question.id) + '/' + action, {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {})
        }).then(function () {
            TrainingUI.toast('Đã cập nhật câu hỏi.', 'success');
            // Sau phát hành, luôn quay lại danh sách đầy đủ. Nhờ vậy không
            // để lại bộ lọc "Đã phát hành" khiến các câu nháp tưởng như mất.
            if (action === 'publish') resetFilters();
            loadList();
            loadDetail(question.id);
        }).catch(function (error) { renderErrors(errors, error); });
    }

    function renderReviewActions(question) {
        if (!canReview || question.publication.status === 'published') return;
        var controls = document.createElement('div');
        controls.className = 'training-action-row';
        if (question.review_status !== 'approved') {
            controls.appendChild(actionButton('Duyệt', function () {
                TrainingUI.confirm('Duyệt câu hỏi này?').then(function (confirmed) {
                    if (confirmed) runReviewAction(question, 'approve');
                });
            }));
        }
        if (question.review_status !== 'rejected') {
            controls.appendChild(actionButton('Từ chối', function () {
                var comment = window.prompt('Nhận xét từ chối (không bắt buộc):', '');
                if (comment !== null) {
                    TrainingUI.confirm('Từ chối câu hỏi này?').then(function (confirmed) {
                        if (confirmed) runReviewAction(question, 'reject', { comment: comment });
                    });
                }
            }));
        }
        if (question.review_status === 'approved') {
            controls.appendChild(actionButton('Phát hành', function () {
                TrainingUI.confirm('Phát hành câu hỏi này? Câu đã phát hành không được sửa trực tiếp.').then(function (confirmed) {
                    if (confirmed) runReviewAction(question, 'publish');
                });
            }));
        }
        detail.appendChild(controls);
    }

    function loadDetail(versionId) {
        renderErrors(errors, null);
        TrainingUI.fetchJson('/api/training/questions/' + encodeURIComponent(versionId))
            .then(renderDetail)
            .catch(function (error) { renderErrors(errors, error); });
    }

    function parseBatch() {
        try {
            return JSON.parse(textarea.value);
        } catch (exception) {
            renderErrors(batchErrors, { message: 'JSON không hợp lệ.', details: { errors: [exception.message] } });
            return null;
        }
    }

    function sendBatch(endpoint, success) {
        var payload = parseBatch();
        if (!payload) return;
        renderErrors(batchErrors, null);
        TrainingUI.fetchJson(endpoint, {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
        }).then(success).catch(function (error) { renderErrors(batchErrors, error); });
    }

    filters.addEventListener('submit', function (event) { event.preventDefault(); page = 1; loadList(); });
    if (clearFiltersButton) {
        clearFiltersButton.addEventListener('click', function () {
            resetFilters();
            loadList();
        });
    }
    if (textarea) {
        document.getElementById('question-batch-validate').addEventListener('click', function () {
            sendBatch('/api/training/questions/validate', function () { TrainingUI.toast('JSON hợp lệ.', 'success'); });
        });
        document.getElementById('question-batch-import').addEventListener('click', function () {
            sendBatch('/api/training/questions/import', function (result) {
                TrainingUI.toast('Đã tạo ' + result.version_ids.length + ' câu hỏi nháp.', 'success');
                page = 1;
                loadList();
            });
        });
    }
    loadList();
}(window, document));
