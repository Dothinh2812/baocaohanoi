(function (window, document) {
    'use strict';

    var panel = document.getElementById('training-my-exams');
    if (!panel || !window.TrainingUI) return;

    var container = panel.querySelector('.training-my-exams-body');
    var loaded = false;

    function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); }

    function renderError(msg) {
        clear(container);
        var p = document.createElement('p');
        p.className = 'training-form-errors';
        var strong = document.createElement('strong');
        strong.textContent = msg || 'Không thể tải danh sách bài thi.';
        p.appendChild(strong);
        container.appendChild(p);
    }

    function categorize(items) {
        var result = { in_progress: [], not_started: [], completed: [], expired: [] };
        items.forEach(function (a) {
            var s = a.assignment_status;
            if (s === 'expired' || s === 'cancelled') {
                result.expired.push(a);
            } else if (s === 'completed') {
                result.completed.push(a);
            } else if (a.attempt_id && (a.attempt_status === 'active' || a.attempt_status === 'created')) {
                result.in_progress.push(a);
            } else {
                result.not_started.push(a);
            }
        });
        return result;
    }

    function categoryLabel(key) {
        return {
            in_progress: 'Đang làm',
            not_started: 'Chưa bắt đầu',
            completed: 'Đã hoàn thành',
            expired: 'Hết hạn',
        }[key] || key;
    }

    function formatDuration(seconds) {
        if (!seconds) return '—';
        var m = Math.round(seconds / 60);
        return m + ' phút';
    }

    function assignmentCard(a) {
        var card = document.createElement('div');
        card.className = 'training-assignment-card';

        var title = document.createElement('p');
        var strong = document.createElement('strong');
        strong.textContent = (a.exam_code || '') + ' — ' + (a.exam_title || '');
        title.appendChild(strong);
        card.appendChild(title);

        var timeRow = document.createElement('p');
        timeRow.textContent = 'Thời gian: ' + TrainingUI.formatTime(a.start_at_ms) + ' — ' + TrainingUI.formatTime(a.end_at_ms);
        card.appendChild(timeRow);

        var durRow = document.createElement('p');
        durRow.textContent = 'Thời lượng: ' + formatDuration(a.duration_seconds);
        card.appendChild(durRow);

        var passRow = document.createElement('p');
        passRow.textContent = 'Điểm đạt: ' + (a.pass_score_percent != null ? a.pass_score_percent + '%' : '—');
        card.appendChild(passRow);

        if (a.attempt_status) {
            var attemptRow = document.createElement('p');
            attemptRow.textContent = 'Trạng thái bài làm: ' + a.attempt_status;
            card.appendChild(attemptRow);
        }

        if (a.attempt_id && a.deadline_at_ms) {
            var deadlineRow = document.createElement('p');
            deadlineRow.textContent = 'Hạn nộp: ' + TrainingUI.formatTime(a.deadline_at_ms);
            card.appendChild(deadlineRow);
        }

        var actionRow = document.createElement('div');
        actionRow.className = 'training-action-row';

        var canStart = a.assignment_status === 'assigned' && a.exam_status === 'open' && !a.attempt_id;
        var canContinue = a.attempt_id && (a.attempt_status === 'active' || a.attempt_status === 'created');
        var isCompleted = a.assignment_status === 'completed';

        if (canStart) {
            var startBtn = document.createElement('button');
            startBtn.type = 'button';
            startBtn.className = 'training-action';
            startBtn.textContent = 'Bắt đầu làm bài';
            startBtn.addEventListener('click', function () { handleStart(a, startBtn); });
            actionRow.appendChild(startBtn);
        } else if (canContinue) {
            var contBtn = document.createElement('button');
            contBtn.type = 'button';
            contBtn.className = 'training-action';
            contBtn.textContent = 'Tiếp tục làm bài';
            contBtn.addEventListener('click', function () { handleContinue(a); });
            actionRow.appendChild(contBtn);
        } else if (isCompleted) {
            var viewBtn = document.createElement('button');
            viewBtn.type = 'button';
            viewBtn.className = 'training-action';
            viewBtn.disabled = true;
            viewBtn.textContent = 'Xem kết quả';
            actionRow.appendChild(viewBtn);
        }

        if (actionRow.childNodes.length) card.appendChild(actionRow);

        return card;
    }

    function renderAssignments(items) {
        clear(container);
        if (!items.length) {
            var empty = document.createElement('p');
            empty.className = 'training-empty-state';
            empty.textContent = 'Chưa có bài thi nào được giao.';
            container.appendChild(empty);
            return;
        }

        var groups = categorize(items);
        var keys = ['in_progress', 'not_started', 'completed', 'expired'];

        keys.forEach(function (key) {
            var list = groups[key];
            if (!list.length) return;

            var heading = document.createElement('h4');
            heading.textContent = categoryLabel(key) + ' (' + list.length + ')';
            container.appendChild(heading);

            list.forEach(function (a) {
                container.appendChild(assignmentCard(a));
            });
        });
    }

    function handleStart(assignment, btn) {
        TrainingUI.confirm('Bắt đầu làm bài thi này?').then(function (ok) {
            if (!ok) return;
            btn.disabled = true;
            TrainingUI.setLoading(btn, true);
            TrainingUI.fetchJson('/api/training/assignments/' + encodeURIComponent(assignment.assignment_id) + '/attempts', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: '{}',
            }).then(function (resp) {
                TrainingUI.toast('Đã bắt đầu bài thi.', 'success');
                window.location.hash = 'training-attempt';
                var event;
                try {
                    event = new CustomEvent('training:attempt-load', { detail: { attemptId: resp.attempt_id } });
                } catch (_e) {
                    event = document.createEvent('CustomEvent');
                    event.initCustomEvent('training:attempt-load', true, true, { attemptId: resp.attempt_id });
                }
                panel.dispatchEvent(event);
            }).catch(function (e) {
                renderError(e.message || 'Không thể bắt đầu bài thi.');
            }).finally(function () {
                btn.disabled = false;
                TrainingUI.setLoading(btn, false);
            });
        });
    }

    function handleContinue(assignment) {
        window.location.hash = 'training-attempt';
        var event;
        try {
            event = new CustomEvent('training:attempt-load', { detail: { attemptId: assignment.attempt_id } });
        } catch (_e) {
            event = document.createEvent('CustomEvent');
            event.initCustomEvent('training:attempt-load', true, true, { attemptId: assignment.attempt_id });
        }
        panel.dispatchEvent(event);
    }

    function loadAssignments() {
        if (loaded) return;
        loaded = true;
        TrainingUI.setLoading(container, true);
        TrainingUI.fetchJson('/api/training/my-assignments', { method: 'GET' })
            .then(function (res) {
                renderAssignments(res.items || []);
            })
            .catch(function (e) {
                renderError(e.message || 'Không thể tải danh sách bài thi.');
            })
            .finally(function () {
                TrainingUI.setLoading(container, false);
            });
    }

    panel.addEventListener('training:panel-show', function () {
        if (!loaded) loadAssignments();
    });

    if (window.location.hash === '#training-my-exams' && !panel.hidden) {
        loadAssignments();
    }
}(window, document));
