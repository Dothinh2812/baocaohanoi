(function (window, document) {
    'use strict';

    function csrfToken() {
        var token = document.querySelector('meta[name="csrf-token"]');
        return token ? token.getAttribute('content') : '';
    }

    function parseError(payload, fallback) {
        var error = payload && payload.error;
        if (error && typeof error === 'object') {
            return {
                code: error.code || 'REQUEST_FAILED',
                message: error.message || fallback,
                details: error.details || {}
            };
        }
        return {
            code: payload && payload.code ? payload.code : 'REQUEST_FAILED',
            message: typeof error === 'string' ? error : (payload && payload.message) || fallback,
            details: payload && payload.details ? payload.details : {}
        };
    }

    function noStoreOptions(options) {
        var result = Object.assign({}, options || {});
        result.cache = 'no-store';
        result.headers = Object.assign({
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache'
        }, result.headers || {});
        return result;
    }

    function withCsrf(options) {
        var result = noStoreOptions(options);
        var method = (result.method || 'GET').toUpperCase();
        var token = csrfToken();
        if (token && !['GET', 'HEAD', 'OPTIONS'].includes(method)) {
            result.headers['X-CSRF-Token'] = token;
        }
        return result;
    }

    async function fetchJson(url, options) {
        var response = await fetch(url, withCsrf(options));
        var payload = await response.json().catch(function () { return {}; });
        if (!response.ok) {
            var error = parseError(payload, 'Không thể hoàn tất yêu cầu.');
            error.status = response.status;
            throw error;
        }
        return payload;
    }

    function setLoading(element, loading) {
        if (!element) return;
        element.classList.toggle('is-loading', Boolean(loading));
        element.setAttribute('aria-busy', loading ? 'true' : 'false');
    }

    function toast(message, type) {
        var item = document.createElement('div');
        item.className = 'training-toast training-toast-' + (type || 'info');
        item.setAttribute('role', 'status');
        item.textContent = message;
        document.body.appendChild(item);
        window.setTimeout(function () { item.remove(); }, 4000);
    }

    function confirm(message) {
        return new Promise(function (resolve) {
            var overlay = document.createElement('div');
            overlay.className = 'training-confirm-overlay';
            var dialog = document.createElement('div');
            dialog.className = 'training-confirm-dialog';
            var text = document.createElement('p');
            text.textContent = message;
            var cancel = document.createElement('button');
            cancel.type = 'button';
            cancel.textContent = 'Hủy';
            var accept = document.createElement('button');
            accept.type = 'button';
            accept.textContent = 'Xác nhận';
            function close(value) { overlay.remove(); resolve(value); }
            cancel.addEventListener('click', function () { close(false); });
            accept.addEventListener('click', function () { close(true); });
            dialog.append(text, cancel, accept);
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            accept.focus();
        });
    }

    function formatTime(value) {
        if (value === null || value === undefined || value === '') return '';
        var date = value instanceof Date ? value : new Date(value);
        return Number.isNaN(date.getTime()) ? '' : new Intl.DateTimeFormat('vi-VN', {
            dateStyle: 'short', timeStyle: 'short'
        }).format(date);
    }

    // Giữ điều hướng workspace trong file JS ngoài. Một số môi trường áp CSP
    // chặn script inline; khi đó menu Mẫu đề/Kỳ thi đổi hash nhưng panel vẫn
    // hidden nếu điều hướng chỉ nằm trong template.
    function initWorkspaceNavigation() {
        var panels = document.querySelectorAll('.training-workspace-grid > .training-workspace-panel');
        if (!panels.length) return;

        function hasWorkspacePanel(panelId) {
            return Array.prototype.some.call(panels, function (panel) { return panel.id === panelId; });
        }
        function showPanel(panelId) {
            if (!hasWorkspacePanel(panelId)) return;
            panels.forEach(function (panel) { panel.hidden = (panel.id !== panelId); });
            document.querySelectorAll('.training-workspace-nav a[data-panel]').forEach(function (link) {
                link.classList.toggle('active', link.getAttribute('data-panel') === panelId);
            });
            var shown = document.getElementById(panelId);
            if (shown && typeof CustomEvent === 'function') {
                shown.dispatchEvent(new CustomEvent('training:panel-show'));
            }
        }

        document.querySelectorAll('.training-workspace-nav a[data-panel]').forEach(function (link) {
            link.addEventListener('click', function (event) {
                event.preventDefault();
                var target = link.getAttribute('data-panel');
                if (window.history && window.history.pushState) {
                    window.history.pushState(null, '', '#' + target);
                } else {
                    window.location.hash = target;
                }
                showPanel(target);
            });
        });
        window.addEventListener('hashchange', function () {
            var hash = (window.location.hash || '').replace('#', '');
            if (hash && hasWorkspacePanel(hash)) showPanel(hash);
        });
        var initial = (window.location.hash || '').replace('#', '');
        if (initial && hasWorkspacePanel(initial)) showPanel(initial);
    }

    window.TrainingUI = {
        csrfToken: csrfToken,
        withCsrf: withCsrf,
        fetchJson: fetchJson,
        parseError: parseError,
        setLoading: setLoading,
        toast: toast,
        confirm: confirm,
        formatTime: formatTime,
        noStoreOptions: noStoreOptions
    };
    initWorkspaceNavigation();
}(window, document));
