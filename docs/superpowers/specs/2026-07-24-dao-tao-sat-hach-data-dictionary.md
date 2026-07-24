# Data dictionary — Module Đào tạo & sát hạch (MVP)

- Ngày lập: `2026-07-24`
- Spec nguồn: `docs/superpowers/specs/2026-07-24-dao-tao-sat-hach-design.md` (1.1)
- DB: `training.db` per-instance (`runtime_app/<unit_code>/training.db`), SQLite 3.37+
- Quy ước timestamp: **`INTEGER` UTC epoch milliseconds** (xem `training/time_policy.py`).
- Quy ước ID: chuỗi `TEXT` do server sinh (uuid4 hex hoặc `<prefix>_<ulid-ish>`), ổn định và không đoán được.
- Delete policy: dữ liệu đã snapshot dùng trong attempt/result/report **không xóa cứng** (FK `ON DELETE RESTRICT` hoặc không khai báo FK trực tiếp, kiểm tra nghiệp vụ). Bản ghi draft/loại bỏ được soft-retire hoặc giữ làm lịch sử.
- Mọi cột JSON lưu là chuỗi JSON hợp lệ; không lưu null thay cho `{}`/`[]` khi có ngữ nghĩa danh sách rỗng.

## Chú pháp cột

- `PK` = primary key; `FK` = foreign key; `UQ` = unique; `IDX` = index thường.
- `immutable` = không bao giờ UPDATE sau INSERT (trừ khi bỏ DB).
- `gen` = do server sinh; `input` = do client/worker cung cấp.

---

## 1. Instance & migration

### training_instance_metadata
Đúng 1 dòng/DB. Kiểm tra `unit_code` khớp `DASHV4_UNIT_CODE` lúc migrate.

| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| unit_code | TEXT | no | — | PK | phải khớp env |
| schema_version | INTEGER | no | — | | phiên bản schema cao nhất |
| created_at_ms | INTEGER | no | — | | UTC epoch ms |

### training_schema_migrations
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| version | INTEGER | no | — | PK | |
| applied_at_ms | INTEGER | no | utc_now_ms | | |

---

## 2. Danh mục & phân quyền

### training_domains
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| code | TEXT | no | PK, UQ(code) | vd `quality` |
| name | TEXT | no | | |
| sort_order | INTEGER | no | 0 | |

### training_categories
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| code | TEXT | no | PK | vd `repair_quality` |
| domain_code | TEXT | no | FK domains(code) | |
| name | TEXT | no | | |
| IDX | (domain_code) | | | |

### training_indicators
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| code | TEXT | no | PK | vd `C1.1` |
| name | TEXT | no | | |

### training_topics
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| code | TEXT | no | PK, UQ(code) | |
| name | TEXT | no | | |

### training_competencies
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| code | TEXT | no | PK | |
| name | TEXT | no | | |

### training_audiences
Nhóm đối tượng nghiệp vụ (to_truong, nvkt, b2a, ...). Có mô tả trách nhiệm.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| code | TEXT | no | PK | |
| name | TEXT | no | | |
| responsibility_text | TEXT | yes | | AI không tự suy đoán nếu trống |

### training_services
| code TEXT PK | name TEXT |

### training_tags
| code TEXT PK, UQ(code) | name TEXT |

### training_user_roles
Quyền module, độc lập cột `role` của `users.xlsx`.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| username | TEXT | no | | snapshot từ users.xlsx |
| role_code | TEXT | no | | `learner`/`editor`/`exam_manager`/`admin` |
| granted_at_ms | INTEGER | no | | |
| UQ | (username, role_code) | | | |
| IDX | (username) | | | |

### training_user_audiences
Một người thuộc 0..n nhóm đối tượng nghiệp vụ.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| username | TEXT | no | | |
| audience_code | TEXT | no | FK audiences(code) | |
| UQ | (username, audience_code) | | | |

---

## 3. Kho tri thức

### knowledge_documents
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| document_code | TEXT | no | — | IDX | mã nghiệp vụ, vd `C1.1` |
| title | TEXT | no | — | | |
| document_type | TEXT | no | 'kpi_definition' | | |
| issuer | TEXT | yes | | | |
| review_status | TEXT | no | 'draft' | | `draft/analyzed/approved/needs_confirmation/rejected` |
| created_at_ms | INTEGER | no | | | |
| updated_at_ms | INTEGER | no | | | |
| IDX | (document_code), (review_status) | | | | |

### knowledge_document_versions
Phiên bản bất biến nội dung tài liệu.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen, immutable | |
| document_id | TEXT | no | FK documents(id) | |
| version_number | INTEGER | no | | tăng đơn điệu |
| content_text | TEXT | no | | nội dung paste gốc |
| content_sha256 | TEXT | no | | hash nội dung |
| effective_from_ms | INTEGER | yes | | |
| effective_to_ms | INTEGER | yes | | |
| review_status | TEXT | no | 'draft' | |
| created_at_ms | INTEGER | no | immutable | |
| created_by | TEXT | no | | username |
| UQ | (document_id, version_number) | | | |

### knowledge_document_topics / knowledge_document_audiences
Mapping nhiều-nhiều.
| document_version_id TEXT FK / document_id TEXT FK | topic_code TEXT FK / audience_code TEXT FK |
UQ (document_version_id, topic_code) tương ứng.

### knowledge_blocks
Đoạn tri thức có mã ổn định trong `(document_version_id, extraction_revision)`.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| document_version_id | TEXT | no | FK versions(id) | |
| extraction_revision | INTEGER | no | | bump khi re-chunk |
| block_id | TEXT | no | | vd `C1.1-B007` |
| heading_path_json | TEXT | yes | | JSON list |
| char_start | INTEGER | yes | | |
| char_end | INTEGER | yes | | |
| content | TEXT | no | | |
| content_sha256 | TEXT | no | | |
| domain_code | TEXT | yes | FK | |
| category_code | TEXT | yes | FK | |
| created_at_ms | INTEGER | no | immutable | |
| UQ | (document_version_id, extraction_revision, block_id) | | | |
| IDX | (document_version_id, extraction_revision) | | | |

### knowledge_block_topics
| block_id TEXT FK | topic_code TEXT FK | UQ(block_id, topic_code) |

### knowledge_issues
Vấn đề cần xác nhận.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| document_version_id | TEXT | no | FK | |
| block_id | TEXT | yes | | |
| severity | TEXT | no | | `low/medium/high` |
| issue_type | TEXT | no | | vd `ambiguous`,`conflict` |
| description | TEXT | no | | |
| status | TEXT | no | 'open' | `open/confirmed/excluded/resolved_by_new_version` |
| resolved_by | TEXT | yes | | |
| resolved_at_ms | INTEGER | yes | | |
| created_at_ms | INTEGER | no | | |
| IDX | (document_version_id, status) | | | |

### knowledge_rules
Công thức/SLA trích xuất (MVP: tối thiểu, dùng cho calculation rule registry).
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| document_version_id | TEXT | no | FK | |
| rule_code | TEXT | no | | `repair_duration_v1` |
| rule_text | TEXT | no | | mô tả tự nhiên |
| created_at_ms | INTEGER | no | | |
| IDX | (rule_code) | | | |

---

## 4. Kho câu hỏi

### question_items
Thực thể logic một câu hỏi.
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| code | TEXT | yes | — | IDX | mã nghiệp vụ tuỳ chọn |
| current_version_id | TEXT | yes | — | | trỏ version mới nhất (denormalized) |
| source_changed | INTEGER | no | 0 | | 0/1 |
| created_at_ms | INTEGER | no | | | |
| IDX | (code) | | | | |

### question_versions
Nội dung bất biến của một câu hỏi tại một thời điểm.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen, immutable | |
| question_item_id | TEXT | no | FK items(id) | |
| version_number | INTEGER | no | | |
| type | TEXT | no | | `single_choice` MVP; khác = sau MVP |
| stem | TEXT | no | | |
| stimulus | TEXT | yes | | |
| language | TEXT | no | 'vi' | |
| correct_option_ids_json | TEXT | no | | JSON list, server-only |
| explanation | TEXT | yes | | server-only |
| distractor_rationales_json | TEXT | yes | | server-only |
| difficulty | TEXT | no | | `easy/medium/hard` |
| cognitive_level | TEXT | yes | | |
| criticality | TEXT | yes | | |
| estimated_seconds | INTEGER | yes | | |
| max_score | REAL | no | 1.0 | |
| scoring_policy_json | TEXT | yes | | |
| normalized_stem_hash | TEXT | yes | | để chặn duplicate publish |
| review_status | TEXT | no | 'draft' | `draft/needs_review/approved/rejected` |
| publication_status | TEXT | no | 'unpublished' | `unpublished/published/retired_for_new_exams` |
| created_by | TEXT | no | | |
| created_at_ms | INTEGER | no | immutable | |
| approved_by | TEXT | yes | | |
| approved_at_ms | INTEGER | yes | | |
| UQ | (question_item_id, version_number) | | | |
| IDX | (publication_status), (review_status) | | | |

### question_options
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| question_version_id | TEXT | no | FK versions(id) | |
| option_code | TEXT | no | | `A`,`B`,... |
| option_text | TEXT | no | | |
| display_order | INTEGER | no | | |
| UQ | (question_version_id, option_code) | | | |

### question_sources
Liên kết câu hỏi ↔ khối tri thức (evidence).
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| question_version_id | TEXT | no | FK | |
| document_version_id | TEXT | no | FK | |
| block_id | TEXT | no | | |
| extraction_revision | INTEGER | no | | |
| quoted_text | TEXT | yes | | |
| quote_start | INTEGER | yes | | server xác minh |
| quote_end | INTEGER | yes | | server xác minh |
| supports | TEXT | yes | | `correct_answer` |
| IDX | (document_version_id, block_id) | | | |
| IDX | (question_version_id) | | | |

### question_topics / question_indicators / question_audiences / question_competencies / question_tags
Mapping nhiều-nhiều của question_version.
| question_version_id TEXT FK | <dim>_code TEXT FK | UQ(question_version_id, <dim>_code) |

### question_reviews
Lịch sử duyệt/sửa/từ chối.
| id TEXT PK, gen | question_version_id TEXT FK | action TEXT | reviewer TEXT | comment TEXT | created_at_ms INTEGER |
IDX (question_version_id).

### ai_generation_jobs
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| status | TEXT | no | 'pending' | | `pending/running/completed/failed/cancelled` |
| idempotency_key | TEXT | yes | — | UQ | tránh tạo trùng khi refresh |
| request_payload_json | TEXT | no | | | |
| source_document_version_ids_json | TEXT | no | | | |
| target_audience_codes_json | TEXT | no | | | |
| requested_count | INTEGER | no | | | |
| created_by | TEXT | no | | | |
| created_at_ms | INTEGER | no | | | |
| claimed_by | TEXT | yes | | | worker id |
| lease_expires_at_ms | INTEGER | yes | | | |
| heartbeat_at_ms | INTEGER | yes | | | |
| retry_count | INTEGER | no | 0 | | |
| max_retries | INTEGER | no | | | |
| error_code | TEXT | yes | | | |
| error_detail | TEXT | yes | | | |
| completed_at_ms | INTEGER | yes | | | |
| IDX | (status, lease_expires_at_ms) | | | | |

### ai_generation_batches
Raw + processed batch của một job.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| job_id | TEXT | no | FK jobs(id) | |
| schema_version | TEXT | no | | |
| provider | TEXT | no | | `fake`/`openai` |
| model | TEXT | yes | | |
| prompt_version | TEXT | yes | | |
| usage_json | TEXT | yes | | |
| raw_response_text | TEXT | yes | | audit, không dùng làm dữ liệu duyệt |
| questions_payload_json | TEXT | yes | | parsed batch |
| created_at_ms | INTEGER | no | | |
| UQ | (job_id) | | | 1 batch / job |

---

## 5. Đề và thi

### exam_templates
Danh sách cố định question versions đã publish (MVP fixed template).
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| code | TEXT | no | — | IDX | |
| title | TEXT | no | — | | |
| target_audience_code | TEXT | no | — | FK audiences(code) | một audience/template (MVP) |
| total_questions | INTEGER | no | — | | |
| duration_seconds | INTEGER | no | — | | |
| pass_score_percent | REAL | no | 80.0 | | |
| shuffle_questions | INTEGER | no | 0 | | |
| shuffle_options | INTEGER | no | 0 | | |
| locked | INTEGER | no | 0 | | chốt trước ready |
| created_by | TEXT | no | | | |
| created_at_ms | INTEGER | no | | | |

### exam_template_items
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| template_id | TEXT | no | FK templates(id) | |
| sequence_number | INTEGER | no | | |
| question_version_id | TEXT | no | FK versions(id) | chỉ published |
| section_label | TEXT | yes | | |
| points | REAL | no | 1.0 | |
| UQ | (template_id, sequence_number) | | | |
| IDX | (template_id, question_version_id) | | | |

### exam_events (kỳ thi)
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| code | TEXT | no | — | IDX | |
| title | TEXT | no | — | | |
| description | TEXT | yes | | | |
| template_id | TEXT | no | — | FK templates(id) | |
| target_audience_code | TEXT | no | — | FK | |
| status | TEXT | no | 'draft' | | `draft/ready/open/closed/cancelled` |
| start_at_ms | INTEGER | no | | | UTC ms |
| end_at_ms | INTEGER | no | | | UTC ms |
| duration_seconds | INTEGER | no | | | giới hạn cá nhân |
| pass_score_percent | REAL | no | | | |
| reveal_answers_after_finalize | INTEGER | no | 1 | | |
| created_by | TEXT | no | | | |
| created_at_ms | INTEGER | no | | | |
| finalized_at_ms | INTEGER | yes | | | |
| finalized_by | TEXT | yes | | | |
| IDX | (status), (start_at_ms) | | | | |

### exam_assignments
Quan hệ người ↔ kỳ thi + snapshot.
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| exam_event_id | TEXT | no | — | FK events(id) | |
| username | TEXT | no | — | | snapshot |
| display_name | TEXT | no | | | snapshot |
| team_code | TEXT | yes | | | snapshot, để trống nếu thiếu |
| team_name | TEXT | yes | | | snapshot |
| organization_code | TEXT | yes | | | snapshot |
| organization_name | TEXT | yes | | | snapshot |
| audience_code | TEXT | no | | | snapshot |
| status | TEXT | no | 'assigned' | | `assigned/completed/expired/cancelled` |
| duration_seconds | INTEGER | no | | | override cá nhân nếu có |
| retake_of_assignment_id | TEXT | yes | | | liên kết thi lại |
| assigned_at_ms | INTEGER | no | | | |
| user_source_updated_at_ms | INTEGER | yes | | | |
| UQ | (exam_event_id, username, audience_code) | | | | |
| IDX | (exam_event_id, username, status) | | | | |

### exam_attempts
Một attempt/assignment (MVP). Bảng này là bảo vệ race cuối cùng.
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| assignment_id | TEXT | no | — | FK assignments(id) | |
| status | TEXT | no | 'created' | | `created/active/submitted/timed_out/administratively_submitted/invalidated` |
| random_seed | INTEGER | no | | | server-side |
| shuffle_algorithm_version | TEXT | no | | | |
| snapshot_checksum | TEXT | no | | | |
| started_at_ms | INTEGER | yes | | | UTC ms, khi `active` |
| deadline_at_ms | INTEGER | yes | | | UTC ms |
| submitted_at_ms | INTEGER | yes | | | |
| ended_reason | TEXT | yes | | | `submit/timeout/admin_close` |
| created_at_ms | INTEGER | no | immutable | | |
| UQ | (assignment_id) | | | | 1 attempt/assignment (MVP) |
| IDX | (assignment_id, status) | | | | |

### exam_attempt_items
Snapshot bất biến câu hỏi của attempt.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen, immutable | |
| attempt_id | TEXT | no | FK attempts(id) | |
| sequence_number | INTEGER | no | | |
| question_version_id | TEXT | no | | snapshot, không JOIN lại kho |
| type | TEXT | no | | |
| stem | TEXT | no | | snapshot |
| stimulus | TEXT | yes | | |
| language | TEXT | no | | |
| correct_option_ids_json | TEXT | no | | server-only |
| explanation | TEXT | yes | | server-only |
| distractor_rationales_json | TEXT | yes | | server-only |
| difficulty | TEXT | yes | | |
| section_label | TEXT | yes | | |
| topic_codes_json | TEXT | yes | | |
| competency_codes_json | TEXT | yes | | |
| points | REAL | no | | |
| max_score | REAL | no | | |
| scoring_policy_json | TEXT | yes | | |
| evidence_json | TEXT | yes | | snapshot sources |
| UQ | (attempt_id, sequence_number) | | | |

### exam_attempt_options
Snapshot bất biến phương án của từng item.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen, immutable | |
| attempt_item_id | TEXT | no | FK attempt_items(id) | |
| option_code | TEXT | no | | |
| option_text | TEXT | no | | |
| display_order | INTEGER | no | | snapshot đã shuffle |
| UQ | (attempt_item_id, option_code) | | | |

### exam_responses
Đáp án learner đã chọn + revision.
| column | type | null | default | key | note |
| --- | --- | --- | --- | --- | --- |
| id | TEXT | no | — | PK, gen | |
| attempt_id | TEXT | no | — | FK attempts(id) | |
| attempt_item_id | TEXT | no | — | FK attempt_items(id) | |
| selected_option_ids_json | TEXT | no | '[]' | | JSON list |
| client_revision | INTEGER | no | 0 | | tăng đơn điệu |
| answered_at_ms | INTEGER | yes | | | UTC ms |
| UQ | (attempt_id, attempt_item_id) | | | | |
| IDX | (attempt_id) | | | | |

### exam_results
Một result gốc/attempt (MVP).
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| attempt_id | TEXT | no | FK attempts(id) | |
| raw_score | REAL | no | | |
| maximum_score | REAL | no | | |
| percent | REAL | no | | |
| passed | INTEGER | no | | 0/1 |
| topic_breakdown_json | TEXT | no | | thống kê theo topic |
| scored_at_ms | INTEGER | no | immutable | |
| status | TEXT | no | 'scored' | `scored/invalidated` |
| UQ | (attempt_id) | | | đúng một result gốc |
| IDX | (attempt_id) | | | |

### exam_result_topics
Chi tiết theo topic cho một result.
| result_id TEXT FK | topic_code TEXT | correct_count INTEGER | total_count INTEGER | points REAL | max_points REAL |
UQ(result_id, topic_code).

### exam_score_adjustments
Điều chỉnh điểm append-only.
| id TEXT PK, gen | result_id TEXT FK results(id) | delta_score REAL | reason TEXT | adjusted_by TEXT | created_at_ms INTEGER |
IDX(result_id).

### exam_report_snapshots
Báo cáo đã chốt, không ghi đè.
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| exam_event_id | TEXT | no | FK events(id) | |
| revision | INTEGER | no | | tăng dần |
| schema_version | TEXT | no | | |
| report_algorithm_version | TEXT | no | | |
| payload_json | TEXT | no | | |
| checksum | TEXT | no | | |
| previous_revision | INTEGER | yes | | |
| created_by | TEXT | no | | |
| created_at_ms | INTEGER | no | immutable | |
| UQ | (exam_event_id, revision) | | | |

---

## 6. Audit

### training_audit_log
| column | type | null | key | note |
| --- | --- | --- | --- | --- |
| id | TEXT | no | PK, gen | |
| actor | TEXT | no | | username |
| unit_code | TEXT | no | | |
| action | TEXT | no | | |
| entity_type | TEXT | no | | |
| entity_id | TEXT | no | | |
| before_json | TEXT | yes | | |
| after_json | TEXT | yes | | |
| request_id | TEXT | yes | | |
| ip | TEXT | yes | | |
| user_agent | TEXT | yes | | |
| created_at_ms | INTEGER | no | | |
| IDX | (entity_type, entity_id, created_at_ms) | | | |

---

## 7. Invariant → constraint/test map

| Invariant (spec §6.3) | Bảo vệ DB | Bảo vệ service/test |
| --- | --- | --- |
| Một attempt chỉ thuộc một assignment | FK attempt_items→attempts | test ownership |
| Một assignment ≤ 1 attempt record | `UQ exam_attempts(assignment_id)` | test hai start đồng thời |
| Attempt kết thúc không quay `active` | — | service compare-and-set + test |
| Attempt item/option bất biến sau `active` | không có UPDATE path | test no-mutate |
| Question version trong attempt không sửa/xóa | không xóa cứng | test snapshot ổn định |
| Result chỉ từ snapshot attempt | chấm đọc attempt_items, không JOIN question_versions | test chấm độc lập |
| ≤ 1 result gốc/attempt | `UQ exam_results(attempt_id)` | test burst submit |
| Report snapshot không ghi đè | `UQ report_snapshots(event_id, revision)` | test finalize lặp |
| Score adjustment append-only | không UPDATE path | test |
| Mỗi write kiểm tra quyền/ownership/status/revision | — | route + service tests |

---

## 8. Transaction boundaries

| Thao tác | Transaction | Ghi chú |
| --- | --- | --- |
| start attempt | `BEGIN IMMEDIATE` ngắn: kiểm tra assignment+exam+time+attempt hiện có → tạo attempt+items+options → commit | unique `(assignment_id)` là lớp cuối |
| autosave response | 1 `INSERT ... ON CONFLICT` hoặc `UPDATE ... WHERE client_revision < ?` ngắn | compare-and-set revision |
| submit/score | compare-and-set `status=active→submitted` + tạo result ngắn | retry trả result hiện có |
| close | compare-and-set event `open→closed` trước, rồi từng active attempt 1 transaction riêng | không giữ 1 transaction lớn |
| finalize | compare-and-set event → recover từng attempt dở → expire assignment → tạo report snapshot `UQ(event,revision)` | retry trả revision hiện hành |
| AI claim job | `BEGIN IMMEDIATE` + `UPDATE ... WHERE status='pending'` | test hai worker |
