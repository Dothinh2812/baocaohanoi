# UI-2 Question Bank Design

## Scope

Add the operator question-bank vertical slice to `/dao-tao-sat-hach`. The slice is limited to question list/detail, JSON batch validation/import, and manager review actions. It does not add learner question APIs, schema migrations, frontend dependencies, or UI-3 screens.

## Access and API

- Editor and exam-manager roles may read the question list and editor/manager detail DTOs. A dashboard admin retains the existing module-role bypass.
- Editors may validate/import drafts. Exam managers may approve, reject with an optional comment, and publish.
- All writes use the existing CSRF protection and all expected business failures use the existing `{"error": {"code", "message", "details"}}` envelope.
- `GET /api/training/questions` is paginated (default 25, maximum 100) and accepts `status`, `audience`, `domain`, `topic`, and `q` filters.
- Status maps existing columns without inventing state: `draft` is unpublished with review `draft` or `needs_review`; `approved` is unpublished with review `approved`; `published` is publication `published`.
- `GET /api/training/questions/<version_id>` returns a dedicated management DTO, including answer and evidence fields. No route returns raw SQLite rows.

## Service and Classification

`training_question_service` remains the sole business and validation source. It creates explicit list and management-detail DTOs.

- Topic and audience derive from the existing `question_topics` and `question_audiences` mappings.
- Domain derives through the existing evidence mapping: `question_sources` joins `knowledge_blocks` by document version, extraction revision, and block identifier. A question without a matching stored block has no derived domain and does not match a domain filter.
- Evidence, review history, and publication information are assembled in service DTOs. The learner attempt service remains the only learner read path and continues to exclude correct answers, explanation, distractor rationales, evidence, and scoring metadata.

## UI

The existing role-aware workspace gains an operator-only question-bank panel and a small dedicated static JS module using `TrainingUI` helpers. It provides filters, accessible pagination, a detail panel, JSON textarea validate/import workflow, field/question error display without clearing input, manager review controls, and a confirmation dialog before publish. Role-specific controls are hidden in the template and enforced by the server.

## Verification

Each behavior is written and run RED before its implementation, then GREEN. Tests cover role denial for all question-bank APIs, draft-only valid import, invalid import envelope, filters/pagination, management detail, reject/approve/publish states and expected failures, plus the real learner attempt route's no-secret payload and `Cache-Control: no-store`. Final verification runs the full suite, Python compilation for changed Python files, diff inspection, and a focused commit/push only to `origin/feat/dao-tao-sat-hach-mvp`.
