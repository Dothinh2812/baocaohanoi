(function (window, document) {
    'use strict';

    var workspace = document.getElementById('training-attempt-workspace');
    if (!workspace || !window.TrainingUI) return;

    var attemptId = null;
    var attempt = null;
    var _revisions = {};
    var _debounceTimers = {};
    var _saveStatuses = {};
    var _countdownTimer = null;
    var _locked = false;

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

    function formatCountdown(ms) {
        if (ms <= 0) return '00:00:00';
        var total = Math.floor(ms / 1000);
        var h = Math.floor(total / 3600);
        var m = Math.floor((total % 3600) / 60);
        var s = total % 60;
        return (h < 10 ? '0' : '') + h + ':' +
               (m < 10 ? '0' : '') + m + ':' +
               (s < 10 ? '0' : '') + s;
    }

    function startCountdown() {
        if (_countdownTimer) window.clearInterval(_countdownTimer);
        var clockEl = workspace.querySelector('.attempt-clock');
        if (!clockEl) return;

        function tick() {
            if (!attempt || !attempt.deadline_at_ms) return;
            var remaining = attempt.deadline_at_ms - Date.now();
            clockEl.textContent = 'Thời gian còn lại: ' + formatCountdown(remaining);
            if (remaining < 5 * 60 * 1000) {
                clockEl.classList.add('attempt-clock-warning');
            }
            if (remaining <= 0) {
                window.clearInterval(_countdownTimer);
                _countdownTimer = null;
                lockUI();
                loadAttemptFromServer();
            }
        }

        tick();
        _countdownTimer = window.setInterval(tick, 1000);
    }

    function lockUI() {
        _locked = true;
        var inputs = workspace.querySelectorAll('input[type="radio"], input[type="checkbox"]');
        for (var i = 0; i < inputs.length; i++) {
            inputs[i].disabled = true;
        }
        var submitBtn = workspace.querySelector('.attempt-submit-btn');
        if (submitBtn) submitBtn.disabled = true;
    }

    function renderTerminalState(status) {
        if (_countdownTimer) { window.clearInterval(_countdownTimer); _countdownTimer = null; }
        clear(workspace);
        var container = document.createElement('div');
        container.className = 'attempt-completed';
        var heading = document.createElement('h3');
        heading.textContent = 'Bài làm đã kết thúc';
        container.appendChild(heading);
        var statusLabels = {
            submitted: 'Đã nộp bài',
            timed_out: 'Hết thời gian',
            cancelled: 'Đã hủy',
        };
        var label = document.createElement('p');
        label.textContent = 'Trạng thái: ' + (statusLabels[status] || status);
        container.appendChild(label);
        var back = document.createElement('button');
        back.type = 'button';
        back.className = 'training-action training-action-secondary';
        back.textContent = 'Về danh sách bài thi';
        back.addEventListener('click', function () {
            window.location.href = '/training/exams';
        });
        container.appendChild(back);
        workspace.appendChild(container);
    }

    function renderResultScreen(submitResp) {
        if (_countdownTimer) { window.clearInterval(_countdownTimer); _countdownTimer = null; }
        clear(workspace);
        var container = document.createElement('div');
        container.className = 'attempt-result';
        var heading = document.createElement('h3');
        heading.textContent = 'Kết quả bài thi';
        container.appendChild(heading);

        var result = submitResp.result || {};
        var scorePara = document.createElement('p');
        var scoreStrong = document.createElement('strong');
        scoreStrong.textContent = 'Điểm: ';
        scorePara.appendChild(scoreStrong);
        scorePara.appendChild(document.createTextNode(
            (result.score != null ? result.score : '?') + ' / ' +
            (result.maximum_score != null ? result.maximum_score : '?')
        ));
        container.appendChild(scorePara);

        var percentPara = document.createElement('p');
        var percentStrong = document.createElement('strong');
        percentStrong.textContent = 'Tỷ lệ: ';
        percentPara.appendChild(percentStrong);
        percentPara.appendChild(document.createTextNode(
            result.percent != null ? result.percent + '%' : '?'
        ));
        container.appendChild(percentPara);

        var passedPara = document.createElement('p');
        var passedStrong = document.createElement('strong');
        passedStrong.textContent = 'Kết quả: ';
        passedPara.appendChild(passedStrong);
        passedPara.appendChild(document.createTextNode(result.passed ? 'Đạt' : 'Chưa đạt'));
        container.appendChild(passedPara);

        if (submitResp.answers_released === false) {
            var note = document.createElement('p');
            note.className = 'attempt-result-note';
            note.textContent = 'Đáp án sẽ được công bố sau khi kỳ thi được chốt.';
            container.appendChild(note);
        }

        var back = document.createElement('button');
        back.type = 'button';
        back.className = 'training-action training-action-secondary';
        back.textContent = 'Về danh sách bài thi';
        back.addEventListener('click', function () {
            window.location.href = '/training/exams';
        });
        container.appendChild(back);
        workspace.appendChild(container);
    }

    function renderWorkspace(data) {
        if (data.status !== 'active') {
            renderTerminalState(data.status);
            return;
        }

        attempt = data;
        _revisions = {};
        _saveStatuses = {};
        _locked = false;
        clear(workspace);

        var header = document.createElement('div');
        header.className = 'attempt-header';
        var clock = document.createElement('div');
        clock.className = 'attempt-clock';
        header.appendChild(clock);
        workspace.appendChild(header);

        var nav = document.createElement('div');
        nav.className = 'attempt-nav';
        nav.setAttribute('aria-label', 'Đi đến câu hỏi');
        workspace.appendChild(nav);

        var content = document.createElement('div');
        content.className = 'attempt-content';
        workspace.appendChild(content);

        var items = data.items || [];
        items.forEach(function (item, idx) {
            var navBtn = document.createElement('button');
            navBtn.type = 'button';
            navBtn.className = 'attempt-nav-btn';
            navBtn.textContent = String(idx + 1);
            navBtn.setAttribute('data-item-index', idx);
            navBtn.addEventListener('click', function () {
                var target = content.querySelector('[data-item-id="' + item.item_id + '"]');
                if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
            });
            nav.appendChild(navBtn);

            var card = document.createElement('div');
            card.className = 'attempt-question-card';
            card.setAttribute('data-item-id', item.item_id);

            var stemRow = document.createElement('div');
            stemRow.className = 'attempt-question-stem';
            var numStrong = document.createElement('strong');
            numStrong.textContent = 'Câu ' + (idx + 1) + '. ';
            stemRow.appendChild(numStrong);
            var stemSpan = document.createElement('span');
            stemSpan.textContent = item.stem;
            stemRow.appendChild(stemSpan);
            card.appendChild(stemRow);

            if (item.stimulus) {
                var stimDiv = document.createElement('div');
                stimDiv.className = 'attempt-question-stimulus';
                stimDiv.textContent = item.stimulus;
                card.appendChild(stimDiv);
            }

            var isMultiple = item.type === 'multiple_choice';
            var inputType = isMultiple ? 'checkbox' : 'radio';

            var optionsDiv = document.createElement('div');
            optionsDiv.className = 'attempt-options';
            (item.options || []).forEach(function (opt) {
                var label = document.createElement('label');
                label.className = 'attempt-option';
                var input = document.createElement('input');
                input.type = inputType;
                input.name = 'item_' + item.item_id;
                input.value = opt.id;
                if (!isMultiple) {
                    input.setAttribute('data-group', item.item_id);
                }
                var selected = item.response && item.response.selected_option_ids &&
                    item.response.selected_option_ids.indexOf(opt.id) !== -1;
                input.checked = selected;
                label.appendChild(input);
                var optText = document.createTextNode(' ' + opt.id + '. ' + opt.text);
                label.appendChild(optText);
                optionsDiv.appendChild(label);
            });
            card.appendChild(optionsDiv);

            var statusDiv = document.createElement('div');
            statusDiv.className = 'attempt-save-status';
            statusDiv.setAttribute('data-save-status', item.item_id);
            card.appendChild(statusDiv);

            if (item.response && item.response.client_revision != null) {
                _revisions[item.item_id] = item.response.client_revision;
            }

            content.appendChild(card);
        });

        var submitRow = document.createElement('div');
        submitRow.className = 'attempt-submit-row';
        var submitBtn = document.createElement('button');
        submitBtn.type = 'button';
        submitBtn.className = 'training-action attempt-submit-btn';
        submitBtn.textContent = 'Nộp bài';
        submitBtn.addEventListener('click', function () { handleSubmit(submitBtn); });
        submitRow.appendChild(submitBtn);
        workspace.appendChild(submitRow);

        attachOptionListeners();
        updateNavAnswered();
        startCountdown();
    }

    function attachOptionListeners() {
        var inputs = workspace.querySelectorAll('.attempt-options input');
        for (var i = 0; i < inputs.length; i++) {
            inputs[i].addEventListener('change', onOptionChange);
        }
    }

    function onOptionChange(e) {
        if (_locked) return;
        var card = e.target.closest('.attempt-question-card');
        if (!card) return;
        var itemId = card.getAttribute('data-item-id');
        updateNavAnswered();
        scheduleSave(itemId, card);
    }

    function updateNavAnswered() {
        var cards = workspace.querySelectorAll('.attempt-question-card');
        var navBtns = workspace.querySelectorAll('.attempt-nav-btn');
        for (var i = 0; i < cards.length; i++) {
            var inputs = cards[i].querySelectorAll('.attempt-options input:checked');
            var answered = inputs.length > 0;
            if (navBtns[i]) {
                navBtns[i].classList.toggle('attempt-nav-answered', answered);
            }
        }
    }

    function scheduleSave(itemId, card) {
        if (_debounceTimers[itemId]) window.clearTimeout(_debounceTimers[itemId]);
        _debounceTimers[itemId] = window.setTimeout(function () {
            performSave(itemId, card);
        }, 400);
    }

    function performSave(itemId, card) {
        var inputs = card.querySelectorAll('.attempt-options input');
        var selected = [];
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].checked) selected.push(inputs[i].value);
        }

        var currentRev = (_revisions[itemId] != null) ? _revisions[itemId] : 0;
        var nextRev = currentRev + 1;

        _saveStatuses[itemId] = 'saving';
        setSaveStatusText(itemId, 'Đang lưu...');

        _revisions[itemId] = nextRev;

        var url = '/api/training/attempts/' + encodeURIComponent(attemptId) +
                  '/responses/' + encodeURIComponent(itemId);

        TrainingUI.fetchJson(url, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                selected_option_ids: selected,
                client_revision: nextRev
            })
        }).then(function (resp) {
            _saveStatuses[itemId] = 'saved';
            setSaveStatusText(itemId, 'Đã lưu');
            if (!resp.accepted && resp.stored_response) {
                _revisions[itemId] = resp.stored_response.client_revision;
                syncUIFromServer(itemId, card, resp.stored_response.selected_option_ids);
            }
        }).catch(function () {
            _saveStatuses[itemId] = 'error';
            _revisions[itemId] = currentRev;
            setSaveStatusText(itemId, 'Lỗi lưu — sẽ thử lại');
            scheduleSave(itemId, card);
        });
    }

    function syncUIFromServer(itemId, card, selectedIds) {
        var inputs = card.querySelectorAll('.attempt-options input');
        for (var i = 0; i < inputs.length; i++) {
            inputs[i].checked = selectedIds.indexOf(inputs[i].value) !== -1;
        }
        updateNavAnswered();
    }

    function setSaveStatusText(itemId, text) {
        var el = workspace.querySelector('[data-save-status="' + itemId + '"]');
        if (!el) return;
        clear(el);
        el.textContent = text;
    }

    function handleSubmit(btn) {
        if (_locked) return;
        var items = attempt.items || [];
        var unanswered = 0;
        items.forEach(function (item) {
            var card = workspace.querySelector('[data-item-id="' + item.item_id + '"]');
            if (!card) { unanswered++; return; }
            var checked = card.querySelectorAll('.attempt-options input:checked');
            if (checked.length === 0) unanswered++;
        });

        var msg = unanswered > 0
            ? 'Bạn còn ' + unanswered + ' câu chưa trả lời. Nộp bài ngay?'
            : 'Xác nhận nộp bài?';

        TrainingUI.confirm(msg).then(function (ok) {
            if (!ok) return;
            performSubmit(btn);
        });
    }

    function performSubmit(btn) {
        btn.disabled = true;
        TrainingUI.setLoading(btn, true);

        var url = '/api/training/attempts/' + encodeURIComponent(attemptId) + '/submit';
        TrainingUI.fetchJson(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: '{}'
        }).then(function (resp) {
            if (resp.status === 'submitted' && resp.result) {
                renderResultScreen(resp);
            } else {
                renderTerminalState(resp.status);
            }
        }).catch(function (e) {
            var code = e && e.code;
            if (code === 'ATTEMPT_EXPIRED' || code === 'ATTEMPT_ALREADY_COMPLETED' || code === 'EXAM_NOT_OPEN') {
                lockUI();
                loadAttemptFromServer();
            } else {
                btn.disabled = false;
                TrainingUI.setLoading(btn, false);
                var errBox = workspace.querySelector('.attempt-errors');
                if (!errBox) {
                    errBox = document.createElement('div');
                    errBox.className = 'attempt-errors';
                    workspace.insertBefore(errBox, workspace.firstChild);
                }
                renderErrors(errBox, e);
            }
        });
    }

    function loadAttempt(id) {
        attemptId = id;
        loadAttemptFromServer();
    }

    function loadAttemptFromServer() {
        clear(workspace);
        var loadingEl = document.createElement('div');
        loadingEl.className = 'attempt-loading';
        loadingEl.textContent = 'Đang tải bài làm...';
        workspace.appendChild(loadingEl);

        var url = '/api/training/attempts/' + encodeURIComponent(attemptId);
        TrainingUI.fetchJson(url)
            .then(renderWorkspace)
            .catch(function (e) {
                clear(workspace);
                var errBox = document.createElement('div');
                errBox.className = 'attempt-errors';
                renderErrors(errBox, e);
                workspace.appendChild(errBox);
            });
    }

    workspace.addEventListener('training:attempt-load', function (e) {
        var detail = e.detail || {};
        if (detail.attemptId) loadAttempt(detail.attemptId);
    });

    window.TrainingAttempt = {
        load: loadAttempt,
    };
}(window, document));
