// Client/UI test cho training-exams.js.
//
// Mô phỏng DOM tối thiểu, nạp đúng file JS production, rồi quan sát các nút
// thao tác vòng đời được render cho từng trạng thái kỳ thi (draft/ready/open/
// closed/cancelled). Các nút phải khớp với transition hợp lệ ở backend
// (services/training_exam_service.py):
//   - draft  -> "Sẵn sàng" + "Hủy"
//   - ready  -> "Mở" + "Hủy"
//   - open   -> "Đóng"
//   - closed -> "Chốt" (NHƯNG không hiện nếu finalized_at_ms đã đặt)
//   - cancelled -> không có nút hành động
//
// Chạy: node tests/js/test_training_exams_lifecycle.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '..', '..');
const scriptPath = path.join(repoRoot, 'static', 'js', 'training-exams.js');

class El {
  constructor(tag) {
    this.tagName = (tag || 'div').toUpperCase();
    this.children = [];
    this.firstChild = null;
    this.textContent = '';
    this.className = '';
    this.hidden = false;
    this.type = '';
    this.value = '';
    this.checked = false;
    this.disabled = false;
    this.innerHTML = '';
    this.dataset = {};
    this.options = [];
    this._listeners = {};
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
  addEventListener(type, fn) {
    (this._listeners[type] = this._listeners[type] || []).push(fn);
  }
  click() { (this._listeners.click || []).forEach((fn) => fn.call(this)); }
  focus() {}
  cloneNode() { const node = new El(this.tagName); return node; }
  insertBefore(node, ref) {
    const idx = this.children.indexOf(ref);
    if (idx >= 0) this.children.splice(idx, 0, node);
    else this.children.push(node);
    node.parentNode = this;
    this.firstChild = this.children[0] || null;
    return node;
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
}

function matches(el, selector) {
  if (selector === 'button') return el.tagName === 'BUTTON';
  if (selector.startsWith('.')) return String(el.className).split(/\s+/).includes(selector.slice(1));
  return el.tagName === selector.toUpperCase();
}

let registry;
let list;
let detail;
let currentExam;

function resetDom() {
  registry = {};
  const make = (id, tag) => {
    const node = new El(tag);
    node.id = id;
    registry[id] = node;
    return node;
  };
  const panel = make('training-exams');
  panel.dataset.audiences = JSON.stringify([{ code: 'a', name: 'Audience A' }]);
  make('exam-create-btn', 'button');
  make('exam-list-btn', 'button');
  make('exam-create-form');
  make('exam-list-view');
  detail = make('exam-detail');
  make('exam-errors');
  list = make('exam-list-body');
  make('exam-pagination');
  make('exam-template-select', 'select');
  make('exam-audience', 'select');
  make('exam-submit-btn', 'button');
  make('exam-cancel-btn', 'button');
  return panel;
}

function makeExam(status, finalizedAtMs) {
  return {
    id: 'exam-1',
    code: 'E1',
    title: 'Kỳ thi mẫu',
    status: status,
    finalized_at_ms: finalizedAtMs || null,
    template: { code: 'TPL', title: 'Mẫu đề' },
    target_audience_code: 'a',
    start_at_ms: 0,
    end_at_ms: 0,
    duration_seconds: 600,
    pass_score_percent: 80,
    reveal_answers_after_finalize: true,
    assignment_summary: {
      total: 0, assigned: 0, in_progress: 0, completed: 0, expired: 0, cancelled: 0,
    },
  };
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

function buildWindow() {
  const win = {};
  win.setTimeout = (...args) => setTimeout(...args);
  win.clearTimeout = (...args) => clearTimeout(...args);
  win.TrainingUI = {
    setLoading() {},
    formatTime: (ms) => String(ms || 0),
    toast() {},
    confirm: () => Promise.resolve(true),
    fetchJson(url) {
      if (url.indexOf('/templates') >= 0) {
        return Promise.resolve({ items: [], total: 0, page: 1, page_size: 100 });
      }
      if (url.indexOf('/exams?') >= 0) {
        return Promise.resolve({
          items: [{
            id: 'exam-1', code: 'E1', title: 'Kỳ thi mẫu', status: currentExam.status,
            target_audience_code: 'a', start_at_ms: 0, end_at_ms: 0,
            finalized_at_ms: currentExam.finalized_at_ms,
          }],
          total: 1, page: 1, page_size: 25,
        });
      }
      if (url.indexOf('/assignments') >= 0) {
        return Promise.resolve({ items: [] });
      }
      if (url.indexOf('/users') >= 0) {
        return Promise.resolve({ items: [] });
      }
      if (url.indexOf('/report') >= 0) {
        return Promise.resolve({ revision: 1 });
      }
      return Promise.resolve(currentExam);
    },
  };
  win.prompt = () => null;
  return win;
}

function lifecycleButtonLabels() {
  const containers = detail.querySelectorAll('.training-action-row');
  if (!containers.length) return [];
  const buttons = containers[0].querySelectorAll('button');
  return buttons.map((b) => String(b.textContent).trim()).filter((l) => l.length > 0);
}

function sameSet(actual, expected) {
  const a = actual.slice().sort();
  const e = expected.slice().sort();
  return a.length === e.length && a.every((value, i) => value === e[i]);
}

async function renderLifecycleFor(status, finalizedAtMs) {
  resetDom();
  const win = buildWindow();
  globalThis.window = win;
  globalThis.TrainingUI = win.TrainingUI;
  globalThis.document = {
    getElementById: (id) => registry[id] || null,
    createElement: (tag) => new El(tag),
    createTextNode: (text) => {
      const node = new El('#text');
      node.textContent = text === null || text === undefined ? '' : String(text);
      return node;
    },
  };
  currentExam = makeExam(status, finalizedAtMs);
  const source = fs.readFileSync(scriptPath, 'utf8');
  // eslint-disable-next-line no-eval
  (0, eval)(source); // IIFE tự chạy -> loadOnce() vẽ danh sách có nút "Xem".
  await flush();
  await flush();
  const xemButton = list.querySelectorAll('button')[0];
  if (!xemButton) throw new Error('Không render được nút "Xem" ở danh sách kỳ thi.');
  xemButton.click(); // -> loadDetail -> renderDetail -> renderLifecycle.
  await flush();
  await flush();
  return lifecycleButtonLabels();
}

const CASES = [
  { status: 'draft', finalizedAtMs: null, expect: ['Sẵn sàng', 'Hủy'] },
  { status: 'ready', finalizedAtMs: null, expect: ['Mở', 'Hủy'] },
  { status: 'open', finalizedAtMs: null, expect: ['Đóng'] },
  { status: 'closed', finalizedAtMs: null, expect: ['Chốt'] },
  { status: 'cancelled', finalizedAtMs: null, expect: [] },
  { status: 'closed', finalizedAtMs: 1234567890, expect: [] },
];

let failed = 0;
for (const testCase of CASES) {
  const got = await renderLifecycleFor(testCase.status, testCase.finalizedAtMs);
  const label = testCase.status + (testCase.finalizedAtMs ? '+finalized' : '');
  if (!sameSet(got, testCase.expect)) {
    failed += 1;
    console.error(
      `FAIL [${label}]: expected [${testCase.expect.join(', ')}], got [${got.join(', ')}]`,
    );
  } else {
    console.log(`ok   [${label}] -> [${got.join(', ')}]`);
  }
}

if (failed > 0) {
  console.error(`${failed} case(s) failed.`);
  process.exit(1);
}
console.log('OK: exam lifecycle controls match legal transitions for all statuses.');
