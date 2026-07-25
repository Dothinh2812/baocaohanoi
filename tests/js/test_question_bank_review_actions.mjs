// Client/UI test cho ngân hàng câu hỏi (training-question-bank.js).
//
// Mô phỏng DOM tối thiểu, nạp đúng file JS production, rồi quan sát các nút
// thao tác duyệt được render cho từng trạng thái review/publication.
// Các nút phải khớp với transition hợp lệ ở backend
// (services/training_question_service.py::add_review_action):
//   - rejected KHÔNG được hiện nút "Từ chối" (transition bất hợp lệ).
//
// Chạy: node tests/js/test_question_bank_review_actions.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '..', '..');
const scriptPath = path.join(repoRoot, 'static', 'js', 'training-question-bank.js');

class El {
  constructor(tag) {
    this.tagName = (tag || 'div').toUpperCase();
    this.children = [];
    this.firstChild = null;
    this.textContent = '';
    this.className = '';
    this.hidden = false;
    this.type = '';
    this.innerHTML = '';
    this.dataset = {};
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
  addEventListener(type, fn) {
    (this._listeners[type] = this._listeners[type] || []).push(fn);
  }
  click() { (this._listeners.click || []).forEach((fn) => fn.call(this)); }
  focus() {}
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

function resetDom() {
  registry = {};
  const make = (id, tag) => {
    const node = new El(tag);
    node.id = id;
    registry[id] = node;
    return node;
  };
  const panel = make('training-question-bank');
  panel.dataset.canReview = 'true';
  const filters = make('question-bank-filters');
  filters.elements = [];
  make('question-bank-clear-filters', 'button');
  list = make('question-bank-list');
  make('question-bank-pagination');
  detail = make('question-bank-detail');
  make('question-bank-errors');
  // question-batch-json cố tình không đăng ký -> textarea === null -> bỏ qua batch.
  return { panel, filters };
}

let currentQuestion;

function makeQuestion(reviewStatus, publicationStatus) {
  return {
    id: 'qv-1',
    stem: 'Câu hỏi mẫu?',
    stimulus: '',
    language: 'vi',
    version: 1,
    type: 'single_choice',
    difficulty: 'easy',
    explanation: 'Giải thích mẫu',
    correct_option_ids: ['o1'],
    classification: {
      domain_codes: ['d'],
      audience_codes: ['a'],
      topic_codes: ['t'],
      indicator_codes: ['i'],
    },
    cognitive_level: 'remember',
    criticality: 'low',
    estimated_seconds: 60,
    max_score: 1,
    scoring_policy: null,
    distractor_rationales: [],
    created_by: 'seed',
    created_at_ms: 0,
    publication: { status: publicationStatus, approved_by: '', approved_at_ms: 0 },
    options: [{ id: 'o1', text: 'Đáp án 1' }],
    evidence: [{ document_version_id: 'dv1', block_id: 'b1', quoted_text: 'q', supports: 'supports' }],
    review_history: [],
    publication_history: [],
    review_status: reviewStatus,
  };
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

function buildWindow() {
  const win = {};
  win.TrainingUI = {
    setLoading() {},
    formatTime: (ms) => String(ms || 0),
    toast() {},
    confirm: () => Promise.resolve(true),
    fetchJson(url) {
      if (String(url).indexOf('?') >= 0) {
        return Promise.resolve({
          items: [
            {
              id: 'qv-1',
              stem: 'Câu hỏi mẫu?',
              type: 'single_choice',
              difficulty: 'easy',
              audience: ['a'],
              topic: ['t'],
              status: 'rejected',
            },
          ],
          total: 1,
          page: 1,
          page_size: 25,
        });
      }
      return Promise.resolve(currentQuestion);
    },
  };
  win.prompt = () => null;
  return win;
}

function actionLabels() {
  return detail
    .querySelectorAll('button')
    .map((button) => String(button.textContent).trim())
    .filter((label) => label.length > 0)
    .sort();
}

function sameSet(actual, expected) {
  const a = actual.slice().sort();
  const e = expected.slice().sort();
  return a.length === e.length && a.every((value, i) => value === e[i]);
}

async function renderActionsFor(reviewStatus, publicationStatus) {
  resetDom();
  const win = buildWindow();
  globalThis.window = win;
  globalThis.TrainingUI = win.TrainingUI; // file JS dùng cả `TrainingUI` trần và `window.TrainingUI`.
  globalThis.document = {
    getElementById: (id) => registry[id] || null,
    createElement: (tag) => new El(tag),
    createTextNode: (text) => {
      const node = new El('#text');
      node.textContent = text === null || text === undefined ? '' : String(text);
      return node;
    },
  };
  currentQuestion = makeQuestion(reviewStatus, publicationStatus);
  const source = fs.readFileSync(scriptPath, 'utf8');
  // eslint-disable-next-line no-eval
  (0, eval)(source); // IIFE tự chạy -> loadList() vẽ 1 dòng có nút "Xem".
  await flush();
  const xemButton = list.querySelectorAll('button')[0];
  if (!xemButton) throw new Error('Không render được nút "Xem" ở danh sách.');
  xemButton.click(); // -> loadDetail -> renderDetail -> renderReviewActions.
  await flush();
  await flush();
  return actionLabels();
}

const CASES = [
  { review: 'draft', expect: ['Duyệt', 'Từ chối'] },
  { review: 'needs_review', expect: ['Duyệt', 'Từ chối'] },
  { review: 'approved', expect: ['Phát hành', 'Từ chối'] },
  { review: 'rejected', expect: ['Duyệt'] },
];

let failed = 0;
for (const testCase of CASES) {
  const got = await renderActionsFor(testCase.review, 'unpublished');
  if (!sameSet(got, testCase.expect)) {
    failed += 1;
    console.error(
      `FAIL [${testCase.review}]: expected [${testCase.expect.join(', ')}], got [${got.join(', ')}]`,
    );
  } else {
    console.log(`ok   [${testCase.review}] -> [${got.join(', ')}]`);
  }
}

// Bộ lọc "Đã phát hành" dễ làm người vận hành tưởng những câu nháp đã mất.
// Nút xóa lọc phải thực sự đưa toàn bộ input về rỗng trước khi tải lại list.
resetDom();
registry['question-bank-filters'].elements = [
  { name: 'status', value: 'published' },
  { name: 'topic', value: 'brcd_repair' },
];
const clearWindow = buildWindow();
globalThis.window = clearWindow;
globalThis.TrainingUI = clearWindow.TrainingUI;
globalThis.document = {
  getElementById: (id) => registry[id] || null,
  createElement: (tag) => new El(tag),
  createTextNode: (text) => {
    const node = new El('#text');
    node.textContent = text === null || text === undefined ? '' : String(text);
    return node;
  },
};
(0, eval)(fs.readFileSync(scriptPath, 'utf8'));
registry['question-bank-clear-filters'].click();
await flush();
if (registry['question-bank-filters'].elements.some((input) => input.value !== '')) {
  failed += 1;
  console.error('FAIL: nút xóa lọc không trả bộ lọc câu hỏi về trạng thái tất cả.');
} else {
  console.log('ok   [clear filters] -> all question filters reset');
}

if (failed > 0) {
  console.error(`${failed} case(s) failed.`);
  process.exit(1);
}
console.log('OK: review/publication actions match legal transitions for all statuses.');
