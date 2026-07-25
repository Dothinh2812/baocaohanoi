# CLI-first: kho tri thức và AI sinh câu hỏi

## Mục tiêu

Cho phép vận hành nội dung thực tế qua terminal trước khi có UI kho tri thức:

```text
Tài liệu TXT/DOCX
-> document/version/blocks/classification trong training.db
-> AI generation job
-> draft câu hỏi có evidence
-> reviewer approve/publish
-> UI hiện có tạo mẫu đề, kỳ thi, giao bài và báo cáo
```

Pilot đầu tiên là chỉ tiêu `C1.1 - Chất lượng sửa chữa thuê bao BRCĐ`.

## Phạm vi và nguyên tắc

- Làm trên `feat/dao-tao-sat-hach-mvp`; không sửa production worktree hay service `dashv4@son_tay`.
- Không ghi SQL trực tiếp vào `training.db`. CLI phải gọi service/validator đã có để giữ unit-code guard, transaction, audit và invariant snapshot.
- Mọi lệnh mutating nhận `--actor`; kiểm tra role module phù hợp và ghi audit.
- AI chỉ tạo `draft`; không tự approve hoặc publish.
- Evidence phải neo bằng `document_version_id`, `block_id`, offsets và content hash; quote text chỉ phục vụ hiển thị/audit.
- MVP import `TXT`, paste text và `DOCX`; chưa thêm PDF/Excel/QTI.
- OpenAI adapter lazy import. Không có SDK/API key thì app, migration và CLI không-AI vẫn chạy.
- API key chỉ đọc từ environment/secret store; tuyệt đối không ghi vào DB, generated env hay audit log.

## Pha 0 - Contract, configuration và test foundation

1. Review schema/service/job queue/validator hiện có, tránh duplicate logic.
2. Chốt argparse contract, stable output JSON/human-readable và stable error codes.
3. Bổ sung `requirements-training.txt` hoặc cơ chế pin dependency phù hợp repo cho `jsonschema`, OpenAI SDK và parser DOCX; kiểm tra runtime system Python và venv.
4. Adapter OpenAI dùng Responses API + Structured Outputs, model/config từ `DASHV4_TRAINING_AI_*`; default chỉ là cấu hình có thể thay đổi.
5. Test missing key, disabled provider, lazy import và mock transport.

**Gate:** CLI help hoạt động, migration/app không bị yêu cầu cài OpenAI SDK khi AI disabled.

## Pha 1 - CLI nhập kho tri thức

Lệnh cần có:

```bash
python3 -m training.cli knowledge-import --file c1_1.docx \
  --title "C1.1 Chất lượng sửa chữa thuê bao BRCĐ" \
  --domain quality --topics brcd_repair \
  --audiences nvkt,to_truong,b2a --actor thinhdx.hni

python3 -m training.cli knowledge-import --paste-file c1_1.txt ...
python3 -m training.cli knowledge-list
python3 -m training.cli knowledge-show --document-version-id <id>
python3 -m training.cli knowledge-issues --document-version-id <id>
```

Yêu cầu:

- Validate extension, MIME/signature, max bytes và DOCX/ZIP safety trước parse.
- Tạo document, immutable version, extraction revision và blocks deterministic trong phạm vi version/revision.
- Gán domain/topic/audience. Nếu topic chưa tồn tại, lệnh catalog riêng phải tạo có audit; không tự tạo silently từ AI.
- Report `document_id`, `document_version_id`, extraction revision, block count và issues.
- Re-import cùng nội dung idempotent theo checksum hoặc trả conflict có hướng dẫn rõ ràng.

**Tests:** TXT/DOCX valid, file giả MIME, max-size, bad DOCX, duplicate, unit-code mismatch, block stability và issue blocking.

## Pha 2 - CLI tạo và quản lý generation job

```bash
python3 -m training.cli generate-create \
  --document-version-ids <ver_id_1>,<ver_id_2> \
  --audiences nvkt \
  --count 15 \
  --actor thinhdx.hni

python3 -m training.cli generate-list
python3 -m training.cli generate-show --job-id <id>
python3 -m training.cli generate-cancel --job-id <id> --actor thinhdx.hni
```

Yêu cầu:

- Job snapshot selection: document versions, block IDs, classification, audience profile, schema/prompt version.
- Idempotency key deterministic hoặc `--idempotency-key`; retry không tạo job/batch trùng.
- Chặn job khi document có blocking issue.
- Sinh batch riêng cho `nvkt`, `to_truong`, `b2a`; không dùng chung prompt/level cho mọi audience.
- Chỉ batch đã validate mới được import draft atomically.

## Pha 3 - Worker và OpenAI adapter production

```bash
# Bắt buộc: API key từ secret store, KHÔNG ghi vào DB/log/commit
export OPENAI_API_KEY='sk-...'
export DASHV4_TRAINING_AI_ENABLED=1

# Tuỳ chọn: model, timeout
export DASHV4_TRAINING_AI_MODEL=gpt-4o-mini
export DASHV4_TRAINING_GENERATION_TIMEOUT_SECONDS=120

# Provider chọn ở worker, KHÔNG phải env var
python3 -m training.cli worker --provider openai --once
python3 -m training.cli worker --provider openai --poll-interval 5
```

Yêu cầu:

- Provider nhận chỉ blocks được phép, audience profile, prompt/schema version và requested count.
- Structured output phải qua JSON Schema và semantic validation nội bộ; không tin output AI trực tiếp.
- Chuẩn hóa timeout, refusal, malformed/truncated output, provider error, retry/lease/cancel/stale worker ownership.
- Lưu metadata provider/model/prompt/schema/usage/raw response có giới hạn và không chứa secret.
- Test mock SDK/transport; gọi API thật chỉ là smoke tách riêng, không chạy trong test suite.

## Pha 4 - CLI review và publish ngân hàng câu hỏi

```bash
python3 -m training.cli questions-list --status draft --topic brcd_repair
python3 -m training.cli questions-show --version-id <id>
python3 -m training.cli questions-approve --version-id <id> --actor thinhdx.hni
python3 -m training.cli questions-publish --version-id <id> --actor thinhdx.hni
```

`questions-show` phải hiển thị stem/options/correct answer/evidence/warning duplicate cho reviewer có quyền. State transition/CAS, version bất biến và audit dùng cùng service/UI hiện có.

## Pha 5 - Pilot C1.1

1. Chuẩn hóa nội dung C1.1 thành TXT hoặc DOCX.
2. Import domain `quality`, topic `brcd_repair`, audiences `nvkt,to_truong,b2a`.
3. Kiểm tra block, issue và evidence sample trước khi gọi AI.
4. Tạo ba generation job: 15 draft cho từng audience.
5. Review/sửa/approve/publish tối thiểu 20 câu chất lượng.
6. Dùng UI hiện có tạo fixed template, kỳ thi thử và giao 2-3 tài khoản test.
7. Làm bài, finalize và tải Excel; xác minh report không đổi khi question bank về sau thay đổi.

## Pha 6 - Nghiệm thu CLI và handoff UI

- Targeted unit/route/worker tests, full `pytest`, `py_compile`, `git diff --check`.
- Manual runbook: migrate -> import -> inspect -> generate -> worker -> review -> publish -> template/exam/report.
- Ghi retention cho file gốc/raw response/audit, backup note và systemd worker là hạng mục vận hành tiếp theo.
- UI kho tri thức/UI enqueue chỉ bắt đầu sau khi pilot C1.1 chạy ổn định qua CLI.

## Out of scope

- UI kho tri thức/UI AI job.
- Blueprint/matrix adaptive, multi-attempt, QTI, proctoring.
- PDF/Excel ingestion, score adjustment UI, retake automation.

