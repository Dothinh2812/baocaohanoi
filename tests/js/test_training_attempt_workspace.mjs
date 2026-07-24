// Client/UI test cho training-attempt.js.
//
// Mô phỏng DOM tối thiểu, nạp đúng file JS production, rồi quan sát
// workspace hiển thị câu hỏi, nav, submit, và khóa khi trạng thái terminal.
//
// Chạy: node tests/js/test_training_attempt_workspace.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '..', '..');
const scriptPath = path.join(repoRoot, 'static', 'js', 'training-attempt.js');

class El {
  constructor(tag) {
    this.tagName = (tag || 'div').toUpperCase();
    this.children = [];
    this.firstChild = null;
    this._textContent = undefined;
    this.className = '';
    this.hidden = false;
    this.type = '';
    this.value = '';
    this.checked = false;
    this.disabled = false;
    this.innerHTML = '';
    this.dataset = {};
    this._listeners = {};
    this._attrs = {};
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
  scrollIntoView() {}
}

function matches(el, selector) {
  if (selector.startsWith('.')) {
    return String(el.className).split(/\s+/).includes(selector.slice(1));
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
  return el.tagName === selector.toUpperCase();
}

function makeAttemptData(status, items) {
  return {
    attempt_id: 'att-test',
    exam_id: 'exam-1',
    status: status || 'active',
    deadline_at_ms: Date.now() + 600000,
    items: items || [
      {
        item_id: 'item-1',
        question_id: 'q-1',
        sequence: 1,
        stem: 'Câu hỏi mẫu 1',
        type: 'single_choice',
        options: [
          { id: 'A', text: 'Đáp án A' },
          { id: 'B', text: 'Đáp án B' },
          { id: 'C', text: 'Đáp án C' },
          { id: 'D', text: 'Đáp án D' },
        ],
        response: null,
      },
      {
        item_id: 'item-2',
        question_id: 'q-2',
        sequence: 2,
        stem: 'Câu hỏi mẫu 2',
        stimulus: 'Đoạn văn mẫu',
        type: 'multiple_choice',
        options: [
          { id: 'A', text: 'Đáp án A' },
          { id: 'B', text: 'Đáp án B' },
        ],
        response: { selected_option_ids: ['A'], client_revision: 3 },
      },
    ],
  };
}

async function loadAttemptWithMock(attemptData) {
  const registry = {};
  const make = (id, tag) => {
    const node = new El(tag);
    node.id = id;
    registry[id] = node;
    return node;
  };
  const ws = make('training-attempt-workspace');

  const mockData = attemptData || makeAttemptData('active');

  globalThis.window = {
    setTimeout: (...a) => setTimeout(...a),
    clearTimeout: (...a) => clearTimeout(...a),
    setInterval: () => 1,
    clearInterval: () => {},
    location: { href: '/training/exams' },
    TrainingUI: {
      setLoading() {},
      formatTime: (ms) => String(ms || 0),
      toast() {},
      confirm: () => Promise.resolve(true),
      fetchJson: () => Promise.resolve(mockData),
    },
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
  };

  const source = fs.readFileSync(scriptPath, 'utf8');
  (0, eval)(source);

  if (!window.TrainingAttempt) {
    throw new Error('TrainingAttempt not exposed');
  }

  window.TrainingAttempt.load('att-test');
  await new Promise((r) => setTimeout(r, 100));
  return ws;
}

// ---- Test 1: render questions from server data ----
{
  const ws = await loadAttemptWithMock(makeAttemptData('active'));

  const cards = ws.querySelectorAll('.attempt-question-card');
  const navBtns = ws.querySelectorAll('.attempt-nav-btn');
  const submitRow = ws.querySelector('.attempt-submit-row');
  const submitBtn = submitRow ? submitRow.querySelector('button') : null;

  let ok = true;
  function check(condition, msg) {
    if (!condition) { ok = false; console.error('  FAIL: ' + msg); }
  }

  check(cards.length === 2, 'expected 2 question cards, got ' + cards.length);
  check(navBtns.length === 2, 'expected 2 nav buttons, got ' + navBtns.length);
  check(submitBtn !== null, 'submit button should exist');
  check(submitBtn && String(submitBtn.textContent).trim() === 'Nộp bài',
    'submit button text should be "Nộp bài", got "' + (submitBtn ? submitBtn.textContent : '') + '"');

  // Verify first question stem
  const firstStem = cards[0] ? cards[0].querySelector('.attempt-question-stem') : null;
  check(firstStem !== null, 'first card should have stem');
  check(firstStem && String(firstStem.textContent).indexOf('Câu hỏi mẫu 1') >= 0,
    'first stem should contain "Câu hỏi mẫu 1", got "' + (firstStem ? firstStem.textContent : '') + '"');

  // Verify second question has stimulus
  const secondStim = cards[1] ? cards[1].querySelector('.attempt-question-stimulus') : null;
  check(secondStim !== null, 'second card should have stimulus');
  check(secondStim && String(secondStim.textContent).indexOf('Đoạn văn mẫu') >= 0,
    'second stimulus should contain "Đoạn văn mẫu"');

  // Verify first question has 4 options
  const firstOpts = cards[0] ? cards[0].querySelectorAll('.attempt-option') : [];
  check(firstOpts.length === 4, 'first card should have 4 options, got ' + firstOpts.length);

  // Verify second question has 2 options
  const secondOpts = cards[1] ? cards[1].querySelectorAll('.attempt-option') : [];
  check(secondOpts.length === 2, 'second card should have 2 options, got ' + secondOpts.length);

  if (ok) console.log('ok   [attempt renders questions from server data]');
  else { console.error('FAIL [attempt renders questions from server data]'); process.exit(1); }
}

// ---- Test 2: lock UI on terminal status ----
{
  const ws = await loadAttemptWithMock(makeAttemptData('submitted'));

  const completed = ws.querySelector('.attempt-completed');
  const hasCompletedText = completed &&
    String(completed.textContent).indexOf('kết thúc') >= 0;

  let ok = true;
  function check(condition, msg) {
    if (!condition) { ok = false; console.error('  FAIL: ' + msg); }
  }

  check(completed !== null, 'terminal state should render .attempt-completed div');
  check(hasCompletedText, 'completed div should contain "kết thúc"');

  if (ok) console.log('ok   [attempt locks UI on terminal status]');
  else { console.error('FAIL [attempt locks UI on terminal status]'); process.exit(1); }
}

console.log('OK: attempt workspace tests passed.');
