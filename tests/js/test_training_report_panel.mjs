// Client/UI test cho training-reports.js.
//
// Mô phỏng DOM tối thiểu, nạp file JS production, rồi quan sát
// report panel hiển thị danh sách kỳ thi, chi tiết báo cáo, và control.
//
// Chạy: node tests/js/test_training_report_panel.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '..', '..');
const scriptPath = path.join(repoRoot, 'static', 'js', 'training-reports.js');

class El {
  constructor(tag) {
    this.tagName = (tag || 'div').toUpperCase();
    this.id = '';
    this.children = [];
    this.firstChild = null;
    this._textContent = undefined;
    this.className = '';
    this.hidden = false;
    this.type = '';
    this.value = '';
    this.disabled = false;
    this.innerHTML = '';
    this.dataset = {};
    this._listeners = {};
    this._attrs = {};
    this.style = {};
    this.role = '';
  }
  appendChild(child) {
    this.children.push(child);
    child.parentNode = this;
    this.firstChild = this.children[0] || null;
    return child;
  }
  removeChild(child) {
    const i = this.children.indexOf(child);
    if (i >= 0) this.children.splice(i, 1);
    this.firstChild = this.children[0] || null;
    return child;
  }
  append(...kids) { kids.forEach((k) => this.appendChild(k)); }
  get childNodes() { return this.children; }
  get textContent() {
    if (this._textContent !== undefined) return this._textContent;
    return this.children.map(c => c.textContent || '').join('');
  }
  set textContent(v) { this._textContent = v; }
  addEventListener(type, fn) {
    (this._listeners[type] = this._listeners[type] || []).push(fn);
  }
  click() { (this._listeners.click || []).forEach((fn) => fn.call(this)); }
  focus() {}
  cloneNode() { return new El(this.tagName); }
  setAttribute(name, val) { this._attrs[name] = val; }
  getAttribute(name) { return this._attrs[name] || null; }
  get classList() {
    const el = this;
    return {
      add(cls) { if (!String(el.className).split(/\s+/).includes(cls)) el.className = (el.className ? el.className + ' ' : '') + cls; },
      remove(cls) { el.className = String(el.className).split(/\s+/).filter(c => c !== cls).join(' '); },
      toggle(cls, force) {
        if (force === undefined) {
          if (String(el.className).split(/\s+/).includes(cls)) { this.remove(cls); return false; }
          else { this.add(cls); return true; }
        }
        if (force) { this.add(cls); return true; }
        else { this.remove(cls); return false; }
      },
      contains(cls) { return String(el.className).split(/\s+/).includes(cls); },
    };
  }
  querySelectorAll(selector) {
    const out = [];
    const walk = (el) => {
      for (const child of el.children) {
        if (matches(child, selector)) out.push(child);
        walk(child);
      }
    };
    walk(this);
    return out;
  }
  querySelector(selector) {
    return this.querySelectorAll(selector)[0] || null;
  }
}

function matches(el, selector) {
  if (selector.startsWith('.')) {
    return String(el.className).split(/\s+/).includes(selector.slice(1));
  }
  if (selector.startsWith('#')) {
    return el.id === selector.slice(1);
  }
  const attrMatch = selector.match(/^\[([^\]=]+)(?:="([^"]+)")?\]$/);
  if (attrMatch) {
    const [, attr, val] = attrMatch;
    const actual = el._attrs[attr] || el.dataset[attr];
    if (val === undefined) return actual != null;
    return actual === val;
  }
  if (selector.includes(' ')) {
    const parts = selector.split(/\s+/);
    const last = parts[parts.length - 1];
    return matches(el, last);
  }
  if (selector === 'button') return el.tagName === 'BUTTON';
  if (selector === 'input') return el.tagName === 'INPUT';
  if (selector === 'table') return el.tagName === 'TABLE';
  if (selector === 'thead') return el.tagName === 'THEAD';
  if (selector === 'tbody') return el.tagName === 'TBODY';
  if (selector === 'tr') return el.tagName === 'TR';
  if (selector === 'td') return el.tagName === 'TD';
  if (selector === 'th') return el.tagName === 'TH';
  if (selector === 'div') return el.tagName === 'DIV';
  if (selector === 'span') return el.tagName === 'SPAN';
  if (selector === 'p') return el.tagName === 'P';
  if (selector === 'h4') return el.tagName === 'H4';
  if (selector === 'h5') return el.tagName === 'H5';
  return el.tagName === selector.toUpperCase();
}

function makeExamListPayload() {
  return {
    items: [
      { id: 'exam-1', code: 'EX01', title: 'Kỳ thi mẫu', status: 'closed', finalized_at_ms: 1700000000000, target_audience_code: 'nvkt', start_at_ms: 1699900000000, end_at_ms: 1700000000000, duration_seconds: 600, pass_score_percent: 80 },
      { id: 'exam-2', code: 'EX02', title: 'Kỳ thi Draft', status: 'draft', finalized_at_ms: null, target_audience_code: 'nvkt', start_at_ms: 1699900000000, end_at_ms: 1700000000000, duration_seconds: 600, pass_score_percent: 80 },
    ],
    total: 2,
    page: 1,
    page_size: 25,
  };
}

function makeExamDetail() {
  return {
    id: 'exam-1', code: 'EX01', title: 'Kỳ thi mẫu', status: 'closed',
    finalized_at_ms: 1700000000000, finalized_by: 'mgr',
    template: { code: 'T1', title: 'Mẫu đề' }, target_audience_code: 'nvkt',
    start_at_ms: 1699900000000, end_at_ms: 1700000000000,
    duration_seconds: 600, pass_score_percent: 80, reveal_answers_after_finalize: 1,
    assignment_summary: { total: 5, assigned: 0, completed: 4, expired: 1, in_progress: 0, cancelled: 0 },
  };
}

function makeReportPayload() {
  return {
    revision: 1,
    payload: {
      summary: { assigned: 5, completed: 4, expired: 1, passed: 3, failed: 1 },
      individual: [
        { assignment: { username: 'u1', display_name: 'User 1' }, result: { raw_score: 18, maximum_score: 20, percent: 90, passed: true } },
        { assignment: { username: 'u2', display_name: 'User 2' }, result: { raw_score: 10, maximum_score: 20, percent: 50, passed: false } },
      ],
    },
  };
}

async function loadReportPanel(mockFetchJson) {
  const registry = {};
  const make = (id, tag) => {
    const node = new El(tag);
    node.id = id;
    registry[id] = node;
    return node;
  };

  make('training-report');
  make('report-errors');
  make('report-list-view');
  make('report-detail');
  make('report-list');
  make('report-pagination');
  make('report-search');

  globalThis.window = {
    TrainingUI: {
      csrfToken: () => 'test-csrf',
      fetchJson: mockFetchJson,
      setLoading: () => {},
      toast: () => {},
      confirm: () => Promise.resolve(true),
      formatTime: (ms) => ms ? '01/01/2026' : '',
      noStoreOptions: (o) => o || {},
      withCsrf: (o) => o || {},
      parseError: (e) => ({ code: 'ERR', message: 'err', details: {} }),
    },
    history: { pushState() {} },
    location: { hash: '', href: '' },
    setTimeout: (...a) => setTimeout(...a),
  };
  globalThis.TrainingUI = globalThis.window.TrainingUI;
  globalThis.document = {
    getElementById: (id) => registry[id] || null,
    createElement: (tag) => new El(tag),
    createTextNode: (text) => {
      const node = new El('#text');
      node.textContent = text === null || text === undefined ? '' : String(text);
      return node;
    },
    body: { appendChild() {} },
    querySelector: () => null,
    querySelectorAll: () => [],
  };

  const source = fs.readFileSync(scriptPath, 'utf8');
  (0, eval)(source);
  return registry;
}

// ---- Test 1: empty state when no finalized exams ----
{
  const fetchCalls = [];
  const registry = await loadReportPanel(async (url) => {
    fetchCalls.push(url);
    return { items: [], total: 0, page: 1, page_size: 25 };
  });

  // Trigger panel-show
  const panel = registry['training-report'];
  const listeners = panel._listeners['training:panel-show'] || [];
  listeners.forEach((fn) => fn.call(panel));
  await new Promise((r) => setTimeout(r, 50));

  const listBody = registry['report-list'];
  const emptyEl = listBody.querySelector('.training-report-empty');

  let ok = true;
  function check(condition, msg) {
    if (!condition) { ok = false; console.error('  FAIL: ' + msg); }
  }

  check(emptyEl !== null, 'should render empty state div');
  check(emptyEl && String(emptyEl.textContent).indexOf('Chưa có') >= 0,
    'empty state should contain "Chưa có", got "' + (emptyEl ? emptyEl.textContent : '') + '"');

  if (ok) console.log('ok   [report panel shows empty state when no finalized exams]');
  else { console.error('FAIL [report panel shows empty state when no finalized exams]'); process.exit(1); }
}

// ---- Test 2: renders exam cards from server data ----
{
  const registry = await loadReportPanel(async (url) => {
    return makeExamListPayload();
  });

  const panel = registry['training-report'];
  const listeners = panel._listeners['training:panel-show'] || [];
  listeners.forEach((fn) => fn.call(panel));
  await new Promise((r) => setTimeout(r, 50));

  const listBody = registry['report-list'];
  const cards = listBody.querySelectorAll('.training-report-exam-card');

  let ok = true;
  function check(condition, msg) {
    if (!condition) { ok = false; console.error('  FAIL: ' + msg); }
  }

  check(cards.length === 1, 'should render 1 finalized exam card (draft filtered out), got ' + cards.length);
  check(cards[0] && String(cards[0].textContent).indexOf('EX01') >= 0,
    'card should contain exam code "EX01"');
  check(cards[0] && String(cards[0].textContent).indexOf('Kỳ thi mẫu') >= 0,
    'card should contain exam title');

  if (ok) console.log('ok   [report panel renders exam cards from server data]');
  else { console.error('FAIL [report panel renders exam cards from server data]'); process.exit(1); }
}

// ---- Test 3: report detail renders summary cards and individual table ----
{
  const registry = await loadReportPanel(async (url) => {
    if (url.includes('/report') && !url.includes('page=')) return makeReportPayload();
    if (url.includes('/exam-1/report')) return makeReportPayload();
    if (url.includes('page=')) return makeExamListPayload();
    return makeExamDetail();
  });

  const panel = registry['training-report'];
  const listeners = panel._listeners['training:panel-show'] || [];
  listeners.forEach((fn) => fn.call(panel));
  await new Promise((r) => setTimeout(r, 50));

  // Click on the first exam card to load report detail
  const listBody = registry['report-list'];
  const cards = listBody.querySelectorAll('.training-report-exam-card');
  if (cards.length > 0) {
    (cards[0]._listeners.click || []).forEach((fn) => fn.call(cards[0]));
    await new Promise((r) => setTimeout(r, 50));
  }

  const detail = registry['report-detail'];
  const summaryCards = detail.querySelectorAll('.training-report-summary-card');
  const table = detail.querySelector('.training-report-table');
  const actionBtns = detail.querySelectorAll('.training-action');
  const exportBtn = actionBtns.length > 1 ? actionBtns[actionBtns.length - 1] : null;

  let ok = true;
  function check(condition, msg) {
    if (!condition) { ok = false; console.error('  FAIL: ' + msg); }
  }

  check(summaryCards.length >= 6, 'should render at least 6 summary cards, got ' + summaryCards.length);
  check(table !== null, 'should render individual results table');
  check(exportBtn !== null, 'should render export button');

  if (table) {
    const rows = table.querySelectorAll('tbody tr');
    check(rows.length === 2, 'should render 2 individual result rows, got ' + rows.length);
  }

  if (exportBtn) {
    check(String(exportBtn.textContent).indexOf('Excel') >= 0,
      'export button should contain "Excel"');
  }

  if (ok) console.log('ok   [report detail renders summary cards and individual table]');
  else { console.error('FAIL [report detail renders summary cards and individual table]'); process.exit(1); }
}

console.log('OK: report panel tests passed.');
