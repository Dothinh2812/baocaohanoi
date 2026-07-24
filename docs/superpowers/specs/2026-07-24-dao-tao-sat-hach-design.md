# Đặc tả module Đào tạo & sát hạch

## 1. Thông tin tài liệu

- Tên: Module Đào tạo & sát hạch tích hợp `dashv4`
- Trạng thái: đặc tả thiết kế, engine fixed-template đã triển khai một phần theo các contract cập nhật bên dưới
- Phiên bản đặc tả: `1.2`
- Ngày lập: `2026-07-24`
- Phạm vi triển khai: từng instance TTVT
- Quy mô mục tiêu: tối đa khoảng 40–50 người dùng mỗi instance

## 2. Các quyết định đã chốt

1. Module được tích hợp vào Flask app hiện tại, không tách microservice.
2. Mỗi instance dùng một SQLite ghi được riêng: `runtime_app/<unit_code>/training.db`.
3. Không ghi dữ liệu đào tạo vào `report_history.db`.
4. Không cần PostgreSQL, Redis, Celery hoặc message queue trong phạm vi này.
5. Người vận hành chủ động tạo kỳ thi, chọn người thi, mở/đóng kỳ thi và chỉ định thời gian.
6. Không có cron tự giao bài, không tự tạo kỳ thi theo tuần và không tự mở kỳ thi chỉ vì đã tới giờ.
7. Hệ thống tự chấm khi nộp bài và tự tổng hợp báo cáo từ kết quả đã chấm.
8. AI chỉ sinh tri thức có cấu trúc và câu hỏi ở trạng thái nháp; AI không tự phát hành câu hỏi hoặc kỳ thi.
9. JSON theo JSON Schema là hợp đồng dữ liệu chuẩn giữa AI, backend, CLI và import/export.
10. PDF/DOCX/XLSX/văn bản dán là nguồn tri thức; Excel là kênh import/export; Word/PDF là định dạng trình bày, không phải nguồn chuẩn cho câu hỏi.
11. Một kho tri thức và một kho câu hỏi chung trong mỗi instance, được phân loại nhiều chiều theo lĩnh vực, chỉ tiêu, chủ đề, đối tượng và năng lực.
12. UI, API nội bộ và CLI phải gọi chung service backend và chung bộ validator.
13. MVP dùng một assignment–một attempt hợp lệ; thi lại được tổ chức thành kỳ thi/assignment mới.
14. MVP không có trạng thái tạm dừng kỳ thi và không dừng đồng hồ bài làm.
15. Deadline bài làm là thời điểm sớm hơn giữa giờ kết thúc kỳ thi và thời lượng cá nhân tính từ lúc bắt đầu.
16. MVP triển khai template câu hỏi cố định trước; blueprint chọn ngẫu nhiên là pha mở rộng sau khi engine thi ổn định.
17. AI chạy qua job worker đọc SQLite, không chạy lâu trong Flask web worker.

## 3. Mục tiêu

Module phải cho phép:

- xây dựng kho tài liệu quy trình, quy định, chỉ tiêu và hướng dẫn nghiệp vụ;
- giữ phiên bản, hiệu lực và nguồn gốc của tri thức;
- dùng AI phân tích tài liệu và sinh câu hỏi có cấu trúc, có dẫn chứng;
- tổ chức ngân hàng câu hỏi theo lĩnh vực, nhóm nghiệp vụ, chỉ tiêu, chủ đề và nhóm đối tượng;
- tạo đề cố định hoặc đề theo ma trận;
- tổ chức một kỳ thi cho một hoặc nhiều nhóm đối tượng;
- trộn câu hỏi và phương án nhưng vẫn tái hiện chính xác đề đã giao;
- chấm tự động và tổng hợp báo cáo sau thi;
- xác định nội dung yếu và hỗ trợ người vận hành tạo đề thi lại cho cá nhân/nhóm;
- thao tác đầy đủ qua UI và hỗ trợ sinh/kiểm tra JSON qua terminal.

## 4. Ngoài phạm vi

- Thi tập trung liên TTVT hoặc kho dữ liệu dùng chung giữa các instance.
- Đồng bộ tự động tri thức, câu hỏi hoặc kết quả giữa các TTVT.
- Lập lịch tự giao bài hoặc tự sinh kỳ thi định kỳ.
- Giám sát thi bằng webcam, nhận diện khuôn mặt hoặc khóa trình duyệt.
- Chấm tự động câu tự luận trong MVP.
- Chuẩn QTI trong MVP; chỉ giữ khả năng bổ sung bộ chuyển đổi sau này.
- AI tự quyết định văn bản còn hiệu lực, tự giải quyết mâu thuẫn hoặc tự duyệt nội dung.
- Dùng tín hiệu đổi tab, IP hoặc tốc độ trả lời làm bằng chứng gian lận duy nhất.

### 4.1. Đường cắt MVP dễ triển khai

| Có trong MVP | Để sau MVP |
| --- | --- |
| Dán văn bản vào kho tri thức | Parser PDF/DOCX/XLSX và OCR |
| AI worker sinh câu `single_choice` có evidence | Nhiều loại câu và duplicate bằng embedding |
| Review và publish question version | Workflow duyệt nhiều cấp |
| Template câu hỏi cố định | Blueprint chọn ngẫu nhiên và supply validation |
| Một audience cho mỗi kỳ thi/template | Nhiều variant trong cùng kỳ thi |
| Một assignment–một attempt | Nhiều lượt làm trong cùng assignment |
| Autosave revision, deadline, submit và chấm | Partial credit và pause/dừng đồng hồ |
| Report cơ bản, finalize, Excel | PDF mẫu chính thức và analytics nâng cao |
| Thi lại bằng kỳ thi/assignment mới | Adaptive retake tự động |

MVP vẫn giữ schema/versioning/snapshot đủ an toàn để mở rộng, nhưng không triển khai trước những nhánh nghiệp vụ chưa cần.

## 5. Thuật ngữ

| Thuật ngữ | Ý nghĩa |
| --- | --- |
| Kho tri thức | Tài liệu nguồn, phiên bản tài liệu, khối tri thức và quy tắc đã trích xuất |
| Khối tri thức | Đoạn tài liệu có mã ổn định, metadata và vị trí nguồn |
| Câu hỏi | Thực thể logic trong ngân hàng câu hỏi |
| Phiên bản câu hỏi | Nội dung bất biến của một câu hỏi tại một thời điểm |
| Blueprint | Quy tắc chọn số câu theo lĩnh vực/chủ đề/độ khó/đối tượng |
| Template | Danh sách phiên bản câu hỏi cố định đã được duyệt |
| Variant | Template/blueprint áp dụng cho một nhóm đối tượng trong một kỳ thi |
| Kỳ thi | Sự kiện có người tham gia, thời gian, chính sách và đề/ma trận đề |
| Lượt được giao | Quan hệ giữa một người và một kỳ thi, kèm nhóm đối tượng/biến thể đề |
| Bài làm | Một lần người dùng bắt đầu và nộp bài |
| Snapshot bài làm | Đề cụ thể và toàn bộ dữ liệu chấm/báo cáo đã đóng băng cho một lần làm |
| Đối tượng | Nhóm nghiệp vụ như tổ trưởng, NVKT, nhân viên B2A |
| Vai trò hệ thống | Quyền dùng module như learner, editor, exam_manager, admin |

## 6. Kiến trúc tổng thể

```text
Browser
  -> Flask blueprint đào tạo
  -> service nghiệp vụ
  -> repository
  -> training.db (ghi được, per-instance)

AI provider
  <- question_generation_service
  <- JSON Schema + tài liệu/khối tri thức đã được phép sử dụng

CLI
  -> cùng service + validator như UI
```

### 6.1. Vị trí trong ứng dụng

Sidebar có một mục cấp cao:

```text
Đào tạo & sát hạch
```

Các chức năng bên trong:

```text
Đào tạo & sát hạch
├── Tổng quan
├── Kho tri thức
├── AI soạn câu hỏi
├── Kho câu hỏi
├── Kho đề và ma trận đề
├── Quản lý kỳ thi
├── Bài thi của tôi
├── Kết quả
└── Báo cáo và tạo đề thi lại
```

### 6.2. Cấu trúc code mục tiêu

Không bắt buộc tách thành package lớn ngay trong commit đầu, nhưng ranh giới lớp phải tương đương:

```text
blueprints/
  training_routes.py

repositories/
  training_repository.py

services/
  knowledge_service.py
  question_generation_service.py
  question_validation_service.py
  exam_service.py
  scoring_service.py
  report_service.py
  review_service.py

schemas/
  knowledge_manifest.schema.json
  question_batch.schema.json
  exam_blueprint.schema.json

templates/pages/training/
static/js/pages/training/
training/
  __init__.py
  cli.py
```

Nếu route trở nên quá lớn, tách thành các blueprint `training_knowledge`, `training_questions`, `training_exams`, `training_attempts`, `training_reports` mà không đổi URL công khai.

### 6.3. Domain invariants

Các invariant sau phải được bảo vệ bằng cả service và constraint DB khi có thể:

- Một attempt chỉ thuộc một assignment.
- Một assignment MVP có tối đa một attempt record, kể cả attempt bị invalidated.
- Attempt đã `submitted`, `timed_out`, `administratively_submitted` hoặc `invalidated` không quay lại `active`.
- Attempt item/option bất biến sau khi attempt chuyển sang `active`.
- Question version đã xuất hiện trong attempt không được sửa hoặc xóa cứng.
- Result chỉ được tạo từ snapshot của attempt, không đọc đáp án hiện hành từ kho câu hỏi.
- Mỗi attempt có tối đa một result gốc; điều chỉnh điểm được ghi append-only.
- Report snapshot đã chốt không bị ghi đè; thay đổi sau chốt tạo revision mới.
- Mỗi write request phải kiểm tra quyền, ownership, trạng thái hiện tại và version/revision mong đợi.

## 7. Phân quyền

### 7.1. Vai trò hệ thống

| Vai trò | Quyền chính |
| --- | --- |
| `learner` | Làm bài, xem kết quả được phép công bố, xem ôn tập cá nhân |
| `editor` | Thêm tài liệu, dùng AI, soạn và sửa câu hỏi nháp |
| `exam_manager` | Duyệt/phát hành câu hỏi, tạo đề, tổ chức kỳ thi, xem báo cáo theo phạm vi |
| `admin` | Toàn quyền và quản lý danh mục/phân quyền |

- Một người có thể có nhiều vai trò module.
- Admin chung của dashboard được quyền quản trị module mặc định, trừ khi cấu hình sau này quy định khác.
- Quyền module lưu trong `training_user_roles`, không mở rộng ý nghĩa cột `role` của `users.xlsx`.

### 7.2. Nhóm đối tượng nghiệp vụ

Nhóm đối tượng độc lập với vai trò hệ thống:

- `to_truong`: Tổ trưởng;
- `nvkt`: Nhân viên kỹ thuật;
- `b2a`: Nhân viên B2A;
- các nhóm mới do admin cấu hình.

Mỗi nhóm phải có mô tả trách nhiệm. AI không được tự suy đoán trách nhiệm của một mã nhóm chưa được mô tả.

Một người có thể thuộc một hoặc nhiều nhóm đối tượng. Khi giao thi, người vận hành chọn nhóm áp dụng cụ thể cho lượt được giao.

## 8. Phân loại kho tri thức và câu hỏi

### 8.1. Cây phân loại chính

```text
Lĩnh vực
  -> Nhóm nghiệp vụ
    -> Chỉ tiêu
      -> Chủ đề
```

Danh mục gợi ý ban đầu:

```text
Chất lượng dịch vụ
├── Chất lượng sửa chữa
├── Chất lượng mạng
├── Báo hỏng lặp lại
└── Chất lượng phục vụ

Phát triển thuê bao
├── FiberVNN
├── MyTV
├── PTTB đúng hẹn
└── Thực tăng thuê bao

Cấu hình tự động
├── Điều kiện cấu hình
├── Kích hoạt
├── Xử lý lỗi
└── Tỷ lệ thành công

Tiếp thị thuê bao
├── Quy trình tiếp thị
├── Mã tiếp thị
├── Kết quả tiếp thị
└── Chăm sóc sau tiếp thị

Gia hạn và duy trì thuê bao
Điều hành kỹ thuật
Vật tư và tài sản
```

Cây được quản lý trong DB, không hardcode trong Python/template.

### 8.2. Các chiều phân loại

Mỗi tài liệu, khối tri thức hoặc câu hỏi có thể mang:

- một lĩnh vực chính;
- một nhóm nghiệp vụ chính;
- không, một hoặc nhiều chỉ tiêu;
- một hoặc nhiều chủ đề;
- một hoặc nhiều năng lực;
- một hoặc nhiều nhóm đối tượng;
- một hoặc nhiều dịch vụ;
- loại kiến thức: định nghĩa, công thức, quy trình, ngoại lệ, SLA, thang điểm;
- độ quan trọng;
- tags tự do có kiểm soát.

### 8.3. Nguyên tắc kế thừa

```text
Phân loại đã xác nhận của tài liệu
  -> khối tri thức
    -> câu hỏi AI sinh
      -> đề/ma trận đề
```

- AI có thể đề xuất phân loại và tags.
- Lĩnh vực, chỉ tiêu và đối tượng do người vận hành chọn là ràng buộc; AI không được tự thay đổi.
- Người duyệt có thể sửa phân loại trước khi phát hành câu hỏi.

## 9. Kho tri thức

### 9.1. Loại đầu vào

- MVP: văn bản dán trực tiếp.
- Pha mở rộng: PDF có text và DOCX.
- Pha mở rộng: XLSX đối với bảng chỉ tiêu/quy định.
- Ngoài MVP: PDF scan sau OCR.

### 9.2. Vòng đời tài liệu

Trạng thái kiểm duyệt và hiệu lực nghiệp vụ được tách riêng:

```text
review_status:
draft -> analyzed -> approved
                  -> needs_confirmation -> approved/rejected
```

Hiệu lực được suy ra từ `effective_from`, `effective_to`, `retired_at` và trạng thái tham khảo; không ghép `effective` vào `review_status`.

- Tài liệu gốc không bị sửa bởi kết quả AI.
- Thay nội dung đã duyệt tạo phiên bản tài liệu mới.
- Mỗi file lưu SHA-256 để nhận diện và audit.
- Phiên bản hết hiệu lực vẫn được giữ để tái hiện lịch sử kỳ thi.

### 9.3. Phân đoạn

Mỗi đoạn có mã ổn định, ví dụ `C1.1-B007`, và metadata:

```json
{
  "block_id": "C1.1-B007",
  "document_version_id": "docver-...",
  "extraction_revision": 1,
  "heading_path": ["C1.1", "Cách tính thời gian hoàn thành sửa chữa"],
  "page_start": null,
  "page_end": null,
  "char_start": 4231,
  "char_end": 4672,
  "content_sha256": "...",
  "content": "...",
  "classification": {
    "domain_code": "quality",
    "category_code": "repair_quality",
    "indicator_codes": ["C1.1"],
    "topic_codes": ["repair_time_calculation"]
  }
}
```

Với văn bản dán, dùng `char_start`/`char_end`; với PDF/DOCX ưu tiên giữ trang, heading và offset nếu parser cung cấp được.

- Block ID chỉ bất biến trong phạm vi một `document_version_id + extraction_revision`.
- Reprocess không được ghi đè âm thầm blocks đã được câu hỏi tham chiếu.
- Parser/chunking thay đổi phải tạo extraction revision mới.
- Phiên bản câu hỏi cũ tiếp tục trỏ tới revision cũ để bảo toàn evidence.

### 9.4. Vấn đề cần xác nhận

AI phải phát hiện nhưng không tự giải quyết:

- câu có nhiều cách hiểu;
- mốc thời gian không nhất quán;
- tên đơn vị/thuật ngữ có alias chưa được xác nhận;
- công thức thiếu biến hoặc khoảng điều kiện;
- hai quy định có khả năng mâu thuẫn;
- văn bản không nêu ngày hiệu lực;
- tham chiếu tới phụ lục/tài liệu chưa được nhập.

Mỗi vấn đề có trạng thái `open`, `confirmed`, `excluded`, `resolved_by_new_version` và lưu người/thời điểm xử lý.

Nội dung `open` mức nghiêm trọng cao không được dùng để sinh câu hỏi có một đáp án chính thức.

## 10. Hợp đồng JSON

### 10.1. Nguyên tắc

- Mọi payload có `schema_version`.
- Bộ schema nội bộ dùng JSON Schema Draft 2020-12; adapter gửi AI có thể chuyển sang tập con mà provider hỗ trợ.
- JSON được validate trước khi ghi DB.
- ID chính và trạng thái phát hành do server tạo; AI chỉ dùng `local_ref`.
- Schema dùng `additionalProperties: false` ở các object nghiệp vụ quan trọng.
- Raw response của AI được lưu để audit, nhưng không được coi là dữ liệu đã duyệt.

### 10.2. Knowledge Document Manifest

```json
{
  "schema_version": "1.0",
  "document_code": "C1.1",
  "title": "Chất lượng sửa chữa thuê bao BRCĐ",
  "document_type": "kpi_definition",
  "document_number": null,
  "issuer": "VNPT Hà Nội",
  "issued_date": null,
  "effective_from": null,
  "effective_to": null,
  "applicable_units": ["ttvt_kv", "tt_cskh", "ttht"],
  "classification": {
    "domain_code": "quality",
    "category_code": "repair_quality",
    "indicator_codes": ["C1.1"],
    "topic_codes": ["brcd_repair", "kpi_scoring"]
  },
  "audience_codes": ["to_truong", "nvkt", "b2a"],
  "source_file": {
    "filename": null,
    "sha256": null
  }
}
```

### 10.3. AI Question Batch

```json
{
  "schema_version": "1.0",
  "batch": {
    "title": "Câu hỏi C1.1 cho NVKT",
    "language": "vi",
    "source_document_version_ids": ["docver-..."],
    "target_audience_codes": ["nvkt"],
    "requested_count": 20
  },
  "questions": [
    {
      "local_ref": "C1.1-AI-001",
      "type": "scenario_single_choice",
      "stem": "Một thuê bao ngoài CCCO báo hỏng lúc 22h30 và được sửa xong lúc 07h30 sáng hôm sau. Thời gian sửa chữa được tính là bao nhiêu?",
      "stimulus": null,
      "options": [
        {"id": "A", "text": "9 giờ"},
        {"id": "B", "text": "1 giờ"},
        {"id": "C", "text": "30 phút"},
        {"id": "D", "text": "0 giờ"}
      ],
      "correct_option_ids": ["D"],
      "explanation": "Phiếu ngoài CCCO báo sau 22h được tính từ 8h hôm sau; hoàn thành trước 8h được tính bằng 0.",
      "distractor_rationales": {
        "A": "Không áp dụng quy tắc loại trừ thời gian qua đêm.",
        "B": "Không có căn cứ tính một giờ.",
        "C": "Không có căn cứ tính 30 phút.",
        "D": "Đáp án đúng."
      },
      "classification": {
        "domain_code": "quality",
        "category_code": "repair_quality",
        "indicator_codes": ["C1.1"],
        "topic_codes": ["repair_time_calculation"],
        "competency_codes": ["apply_repair_time_rule"],
        "audience_codes": ["nvkt"],
        "service_codes": ["FiberVNN", "MyTV"],
        "tags": ["outside_ccco", "after_22h"]
      },
      "difficulty": "medium",
      "cognitive_level": "apply",
      "criticality": "important",
      "estimated_seconds": 60,
      "evidence": [
        {
          "document_version_id": "docver-...",
          "block_id": "C1.1-B007",
          "extraction_revision": 1,
          "quote_start": null,
          "quote_end": null,
          "quoted_text": "Đối với các thuê bao ngoài CCCO báo hỏng sau 22h...",
          "supports": "correct_answer"
        }
      ]
    }
  ]
}
```

### 10.4. Loại câu hỏi trong contract 1.0

- `single_choice`
- `multiple_choice`
- `true_false`
- `scenario_single_choice`
- `scenario_multiple_choice`

Quy tắc:

- `single_choice`, `true_false`, `scenario_single_choice`: mảng `correct_option_ids` có đúng một phần tử.
- Các loại multiple: từ hai đáp án đúng trở lên và ít nhất một đáp án sai.
- Mỗi câu có ít nhất hai phương án; mặc định UI/AI dùng bốn phương án.
- Đáp án tham chiếu bằng ID, không bằng vị trí mảng.

MVP bắt buộc hoàn thiện `single_choice` trước. `scenario_single_choice` dùng cùng scoring contract và có thể bật trong cùng pha; các loại multiple/true-false chỉ bật khi có test đầy đủ, không chặn nghiệm thu engine thi cơ bản.

### 10.5. Exam Blueprint

```json
{
  "schema_version": "1.0",
  "blueprint_code": "NVKT-TONG-HOP-01",
  "title": "Kiểm tra nghiệp vụ tổng hợp NVKT",
  "target_audience_codes": ["nvkt"],
  "total_questions": 20,
  "duration_minutes": 25,
  "pass_score_percent": 80,
  "sections": [
    {
      "domain_code": "quality",
      "category_code": "repair_quality",
      "question_count": 8,
      "difficulty_distribution": {"easy": 2, "medium": 4, "hard": 2}
    },
    {
      "domain_code": "subscriber_growth",
      "question_count": 4
    },
    {
      "domain_code": "automatic_configuration",
      "question_count": 4
    },
    {
      "domain_code": "subscriber_marketing",
      "question_count": 4
    }
  ],
  "shuffle_questions": true,
  "shuffle_options": true
}
```

## 11. Luồng AI

### 11.1. Các pha độc lập

1. **Ingest**: nhận tài liệu, lưu bản gốc, hash và metadata.
2. **Extract**: phân đoạn và trích xuất công thức, định nghĩa, SLA, ngoại lệ.
3. **Audit**: phát hiện nội dung mơ hồ/mâu thuẫn/thiếu căn cứ.
4. **Confirm**: người vận hành xác nhận hoặc loại khỏi phạm vi sinh câu hỏi.
5. **Plan**: tạo mục tiêu kiến thức và ma trận câu hỏi theo đối tượng.
6. **Generate**: sinh Question Batch JSON theo schema.
7. **Validate**: kiểm tra cấu trúc, dẫn chứng, công thức, trùng lặp và điều kiện biên.
8. **Review**: người vận hành duyệt/sửa/từ chối/yêu cầu viết lại.
9. **Publish**: server tạo phiên bản câu hỏi bất biến.

Không gộp tất cả các pha thành một prompt duy nhất.

### 11.2. Ràng buộc tạo câu hỏi

- Chỉ dùng tài liệu và phiên bản đã được chọn.
- Không dùng kiến thức bên ngoài nếu yêu cầu là câu hỏi bám văn bản.
- Mỗi câu phải có ít nhất một evidence hỗ trợ đáp án.
- Câu hỏi không được dựa trên issue nghiêm trọng đang `open`.
- AI phải giải thích đáp án và lý do sai của distractor.
- Không sinh câu hỏi về thông tin không cần đánh giá như số điện thoại, trừ khi người vận hành yêu cầu rõ.
- Câu tính toán phải có input cấu trúc để backend tính lại độc lập.
- Câu tính toán chỉ tham chiếu `rule_code` trong registry do code quản lý; không thực thi biểu thức do AI cung cấp.
- Câu theo đối tượng phải dựa trên hồ sơ trách nhiệm của đối tượng.

Ví dụ payload tính toán:

```json
{
  "calculation": {
    "rule_code": "repair_duration_v1",
    "inputs": {
      "reported_at": "2026-07-24T15:30:00Z",
      "completed_at": "2026-07-25T00:30:00Z",
      "is_outside_ccco": true
    },
    "expected": {
      "value": 0,
      "unit": "minute"
    }
  }
}
```

### 11.3. Kiểm tra tự động

- JSON đúng schema.
- Số câu đúng yêu cầu.
- ID phương án duy nhất và đáp án tồn tại.
- Loại câu phù hợp số đáp án đúng.
- Evidence tồn tại trong đúng phiên bản và extraction revision của tài liệu.
- AI trả `block_id` và quote; backend tự tìm/xác minh span, lưu `quote_start`, `quote_end` và content hash. Không tin offset AI nếu chưa kiểm chứng.
- Trong raw AI batch, offsets có thể `null`; question version không được publish cho tới khi backend đã điền span xác minh duy nhất.
- Không dùng tài liệu hết hiệu lực cho câu mới, trừ chế độ lịch sử có cảnh báo.
- Exact normalized hash trùng thì chặn publish; token/Jaccard similarity cao chỉ tạo cảnh báo cho reviewer trong MVP.
- Không lộ đáp án trong stem.
- Không có hai phương án trùng nghĩa rõ ràng.
- Công thức và đáp án số được backend tính lại.
- Kiểm tra các điểm biên `>`, `>=`, `<`, `<=`.

### 11.4. Trạng thái generation job

```text
pending -> running -> completed
                   -> failed
                   -> cancelled
running -> pending (lease hết hạn và còn lượt retry)
```

MVP dùng một CLI worker polling bảng job; không dùng Redis/Celery và không chạy tác vụ AI lâu trong web request. Worker claim job bằng transaction, lưu `claimed_by`, `lease_expires_at`, `heartbeat_at`, `retry_count`, `error_code` và `error_detail`. Job lease hết hạn có thể được claim lại nếu chưa vượt số lần retry.

## 12. Kho câu hỏi

### 12.1. Vòng đời

```text
question_version.review_status:
draft -> needs_review -> approved/rejected

question_version.publication_status:
unpublished -> published -> retired_for_new_exams
```

- `question_item` giữ lifecycle chung và cờ như `source_changed`.
- Sửa câu `published` luôn tạo question version mới.
- Không xóa cứng câu đã xuất hiện trong bài thi.
- Tài liệu nguồn đổi phiên bản sẽ gắn `source_changed` lên item và tạo review task; version cũ vẫn giữ trạng thái lịch sử, không đổi ngược thành `needs_review`.

### 12.2. Chức năng

- Tìm kiếm và lọc nhiều chiều.
- Xem câu hỏi và chứng cứ song song.
- Soạn/sửa thủ công.
- Duyệt, từ chối, viết lại bằng AI.
- Import/export Excel và JSON.
- Phát hiện câu thiếu nguồn, trùng lặp hoặc cần rà soát.
- Thống kê độ khó thực nghiệm và tỷ lệ chọn từng phương án sau khi có kết quả thi.

## 13. Kho đề và ma trận đề

### 13.1. Template cố định — phạm vi MVP

Người vận hành chọn cụ thể các phiên bản câu hỏi. Khi kỳ thi được khóa, danh sách không đổi.

### 13.2. Blueprint — pha mở rộng

Hệ thống chọn câu tại thời điểm tạo lượt/bắt đầu bài theo:

- nhóm đối tượng;
- lĩnh vực/nhóm/chỉ tiêu/chủ đề;
- năng lực;
- độ khó;
- loại câu;
- trạng thái `published`;
- phiên bản tài liệu còn hiệu lực.

Nếu kho không đủ câu, không được tự nới điều kiện âm thầm. UI phải báo thiếu bao nhiêu câu và cho phép người vận hành sửa ma trận hoặc sinh bổ sung.

### 13.3. Nhiều nhóm đối tượng

Một kỳ thi có thể có nhiều biến thể:

```text
Kỳ thi C1.1
├── Biến thể tổ trưởng
├── Biến thể NVKT
└── Biến thể B2A
```

Mỗi lượt được giao lưu `audience_code` và `exam_variant_id`.

Đề có thể gồm phần chung và phần riêng theo vai trò, ví dụ 30% chung và 70% nghiệp vụ.

MVP có thể tạo các kỳ thi/template riêng cho từng nhóm trước. Một kỳ thi nhiều variant và blueprint tự chọn câu chỉ triển khai sau khi fixed-template engine đã đạt tiêu chí tải và snapshot.

### 13.4. Quy tắc randomization

- Mỗi attempt có `random_seed` do server sinh.
- Thuật toán shuffle và `shuffle_algorithm_version` được snapshot.
- Cùng seed, thuật toán và input phải tái tạo cùng thứ tự.
- Việc trộn chỉ thay thứ tự; ID câu/phương án và đáp án đúng không đổi.
- Không dùng seed do client cung cấp.

## 14. Tổ chức kỳ thi

### 14.1. Người vận hành cấu hình

- Tên, mô tả kỳ thi.
- Template; blueprint/variant là phần mở rộng.
- Danh sách người thi.
- `start_at`, `end_at`, thời lượng cá nhân.
- Điểm đạt tổng và ngưỡng bắt buộc theo chủ đề nếu có.
- Chính sách hiển thị điểm/đáp án.
- Chính sách bài chưa nộp khi hết giờ.
- Trộn câu, trộn phương án.

### 14.2. Trạng thái kỳ thi

```text
draft -> ready -> open -> closed
              \-> cancelled
```

Điều kiện được bắt đầu bài:

```text
exam.status == open
AND start_at <= server_time < end_at
AND assignment.status cho phép
```

Tới `start_at` không tự chuyển `ready` thành `open`. Người vận hành phải bấm mở.

MVP không có `paused`. Nếu cần ngăn người mới bắt đầu, người vận hành đóng kỳ thi; các bài active được xử lý theo chính sách đóng ở mục 14.7.

### 14.3. Vòng đời assignment, attempt và result

#### Assignment

```text
assigned -> completed
         -> expired
         -> cancelled
```

`in_progress` được suy ra khi assignment có attempt `active`, không lưu thành trạng thái độc lập để tránh lệch dữ liệu.

| Từ | Đến | Actor | Điều kiện và side effect |
| --- | --- | --- | --- |
| tạo mới | `assigned` | exam_manager | Snapshot username, tên, tổ, đơn vị, audience và variant/template |
| `assigned` | `completed` | system | Attempt kết thúc hợp lệ và có result |
| `assigned` | `expired` | system/finalize | Hết kỳ thi mà chưa có attempt |
| `assigned` | `cancelled` | exam_manager/admin | Chưa có attempt active, hoặc attempt đã bị admin invalidated; có lý do và audit |

Nếu attempt bị admin invalidated, assignment tương ứng được chuyển `cancelled`; kỳ thi lại dùng assignment mới có liên kết nguồn.

#### Attempt

```text
created -> active -> submitted
                  -> timed_out
                  -> administratively_submitted
created/active -> invalidated
```

| Từ | Đến | Actor | Điều kiện và side effect |
| --- | --- | --- | --- |
| không có | `created` | learner/system | Assignment hợp lệ; tạo snapshot trong transaction |
| `created` | `active` | system | Snapshot hoàn chỉnh; ghi `started_at`, `deadline_at` |
| `active` | `submitted` | learner/system | Compare-and-set; chấm và tạo result một lần |
| `active` | `timed_out` | system | Quá deadline; chấm các response đã nhận đúng hạn |
| `active` | `administratively_submitted` | exam_manager/finalize | Đóng/chốt theo policy; ghi lý do và audit |
| `created`/`active` | `invalidated` | admin | Sự cố hoặc sai nghiệp vụ; result nếu có không còn hiệu lực |

MVP không cho attempt đã kết thúc quay lại active. Thi lại tạo exam/assignment mới, lưu `retake_of_assignment_id` nếu có.

#### Result

```text
scored -> invalidated
```

- Result gốc được tạo đồng thời với transition kết thúc attempt.
- Điều chỉnh điểm không sửa result gốc; ghi `exam_score_adjustments` append-only và tạo report revision mới nếu kỳ thi đã chốt.
- Result hiệu lực bằng result gốc cộng các adjustment hợp lệ.

### 14.4. Chính sách thời gian

Mọi timestamp trong DB lưu dưới UTC (`INTEGER` epoch milliseconds hoặc RFC 3339 kết thúc bằng `Z`; schema phải chọn một cách duy nhất). UI hiển thị bằng `DASHV4_TIMEZONE`, mặc định `Asia/Ho_Chi_Minh`.

```text
deadline_at = min(
    exam.end_at,
    attempt.started_at + assignment.duration_seconds
)
```

- Server time là nguồn chuẩn; client chỉ hiển thị countdown từ `deadline_at` server trả về.
- Độ chính xác tính theo giây; so sánh bằng timestamp tuyệt đối, không làm tròn phút.
- Autosave chỉ được nhận khi server nhận request trước deadline và revision hợp lệ; tại đúng deadline bị từ chối.
- Submit đến sau deadline vẫn được phép gọi idempotently để chuyển bài sang `timed_out`, nhưng không nhận đáp án mới; chỉ chấm response server đã lưu đúng hạn.
- MVP không có grace period thay đổi đáp án. Retry submit không làm mất các response đã autosave.
- Nhiều Gunicorn worker cùng dùng clock hệ điều hành của cùng host; vận hành phải đồng bộ thời gian host bằng NTP/systemd-timesyncd.

### 14.5. Đóng băng snapshot bài làm

Khi bắt đầu bài:

- chọn câu theo template; blueprint là phần mở rộng;
- lưu seed ngẫu nhiên;
- lưu `question_version_id`, loại câu, ngôn ngữ, section và classification;
- lưu nội dung câu, stimulus, phương án, thứ tự và option IDs;
- lưu đáp án đúng, điểm tối đa, scoring policy và normalization rule;
- lưu topic/competency weights, pass threshold liên quan, explanation và evidence;
- lưu schema/algorithm version và checksum snapshot;
- giữ đáp án/scoring metadata ở vùng server không trả cho learner;
- không truy vấn lại kho câu hỏi để chấm.

`exam_attempt_items` và `exam_attempt_options` bất biến từ khi attempt `active`. Thay đổi snapshot chỉ được phép khi attempt còn `created` và chưa từng hiển thị cho learner.

### 14.6. Lưu bài và autosave

- Mỗi lần chọn đáp án có thể autosave qua API.
- Mỗi response có `client_revision` tăng đơn điệu theo item.
- Server chỉ nhận revision lớn hơn revision đã lưu; request cũ đến muộn trả trạng thái hiện hành mà không ghi đè.
- Transaction lưu đáp án phải ngắn.
- Không giữ transaction trong thời gian người dùng làm bài.
- Không lưu toàn bộ bài trong Flask session.
- Session chỉ cần danh tính và `attempt_id` đang thao tác.

Payload tối thiểu:

```json
{
  "selected_option_ids": ["B"],
  "client_revision": 7
}
```

### 14.7. Hết giờ, đóng và chốt

- Client tự nộp khi hết thời lượng.
- Server là nguồn thời gian chuẩn.
- Nếu client đóng/mất mạng, lần truy cập tiếp theo hoặc finalize chuyển attempt quá hạn sang `timed_out` và chấm response đã lưu.
- `close` thủ công kết thúc kỳ thi ngay: ngăn attempt/response mới và chuyển attempt đang active thành `administratively_submitted` từ các response đã lưu; không tạo report snapshot.
- Nếu chỉ hết `end_at` mà chưa bấm close/finalize, API vẫn từ chối attempt/response mới theo time policy; finalize sẽ chuyển attempt active thành `timed_out`.
- `finalize` phục hồi mọi transition còn dang dở, chuyển assignment chưa bắt đầu thành `expired`, bảo đảm attempt đã kết thúc/chấm, rồi tạo report snapshot.
- Mặc định không mở lại kỳ thi đã `closed`; nếu có sai sót vận hành, admin tạo kỳ thi mới hoặc quy trình correction có audit, không đổi ngược trạng thái.

## 15. Chấm điểm

### 15.1. Nguyên tắc

- Chấm dựa trên snapshot đề.
- Submit idempotent: gửi lặp không tạo hai kết quả.
- Điểm lưu cả raw score, maximum score và percent.
- Kết quả lưu theo chủ đề/năng lực để phục vụ phân tích và thi lại.
- Một assignment MVP chỉ có một result hiệu lực; không có chính sách điểm lần đầu/cuối/cao nhất.

### 15.2. Chính sách MVP

- Một đáp án: đúng nhận toàn bộ điểm, sai/không trả lời nhận 0.
- Nhiều đáp án: mặc định chỉ nhận điểm khi chọn đúng toàn bộ và không chọn phương án sai.
- Không trừ điểm âm trong MVP.
- Chính sách partial credit có thể bổ sung, phải là cấu hình rõ trên đề và snapshot.
- Đáp án/giải thích mặc định chỉ được công bố sau khi kỳ thi đã finalize; không phụ thuộc việc cá nhân đã nộp sớm.

### 15.3. Kết quả đạt

```text
passed = total_percent >= pass_score_percent
         AND mọi ngưỡng chủ đề bắt buộc đều đạt
```

## 16. Báo cáo

### 16.1. Báo cáo trực tuyến

Tổng hợp trực tiếp từ DB, cập nhật khi có bài nộp:

- số được chỉ định;
- chưa bắt đầu, đang làm, đã nộp, quá hạn;
- đạt/chưa đạt;
- điểm trung bình/cao nhất/thấp nhất;
- thời gian làm trung bình;
- kết quả theo nhóm đối tượng, tổ, cá nhân;
- kết quả theo lĩnh vực/chỉ tiêu/chủ đề/năng lực;
- câu có tỷ lệ sai cao;
- phân bố lựa chọn distractor;
- danh sách cần thi lại.

### 16.2. Chốt kỳ thi

Nút `Chốt kỳ thi`:

1. đóng kỳ thi;
2. xử lý bài đang dở theo chính sách;
3. bảo đảm mọi bài đã nộp được chấm;
4. tạo toàn bộ report payload JSON chính thức, gồm cấu hình kỳ thi, snapshot người thi/tổ, kết quả cá nhân, thống kê chủ đề/câu hỏi và danh sách chưa thi;
5. ghi người và thời điểm chốt.

Report snapshot có `schema_version`, `report_algorithm_version`, `revision`, checksum và payload JSON. Mở báo cáo sau `end_at` có thể tổng hợp động, nhưng báo cáo chính thức chỉ bất biến sau thao tác chốt.

Điều chỉnh điểm sau chốt không ghi đè snapshot. Hệ thống tạo revision mới, liên kết revision trước, ghi lý do, người duyệt và danh sách thay đổi.

### 16.3. Xuất báo cáo

- Excel: tổng quan, kết quả cá nhân, theo chủ đề, phân tích câu hỏi, danh sách thi lại.
- PDF/Word: bổ sung sau nếu có mẫu báo cáo chính thức.
- File được tạo theo yêu cầu; không cần job định kỳ.
- Excel được dựng từ report snapshot/revision được chọn; mọi text bắt đầu bằng `=`, `+`, `-`, `@` phải được escape để chống formula injection.

## 17. Tạo đề thi lại

Hệ thống không tự tổ chức kỳ thi lại. Người vận hành:

1. mở báo cáo;
2. chọn cá nhân/nhóm chưa đạt;
3. bấm `Đề xuất đề thi lại`;
4. xem phân tích nội dung yếu;
5. duyệt/sửa ma trận hoặc câu hỏi;
6. tạo một kỳ thi mới và chỉ định thời gian.

Gợi ý mặc định có thể gồm:

- 50% câu từ chủ đề yếu;
- 20% câu tương đương câu đã sai;
- 20% câu ôn lại nội dung đã đúng;
- 10% câu mới/ngẫu nhiên.

Tỷ lệ chỉ là mặc định UI, người vận hành được sửa. “Câu tương đương” trong MVP nghĩa là cùng topic và competency, độ khó gần tương đương, khác question item và không trùng normalized stem; không chỉ đổi vài từ của câu cũ.

## 18. Giao diện

### 18.1. Kho tri thức

- Cây lĩnh vực/nhóm/chỉ tiêu/chủ đề ở bên trái.
- Danh sách tài liệu và bộ lọc ở bên phải.
- Thêm bằng paste/file.
- Xem phiên bản, hiệu lực, phân đoạn và câu hỏi liên quan.

### 18.2. Phân tích tri thức

Hai cột:

```text
Văn bản nguồn | Tri thức/công thức/quy tắc AI trích xuất
```

Hiển thị tổng số công thức, biến, SLA, ngoại lệ và issue cần xác nhận.

### 18.3. Cấu hình sinh câu hỏi

- Nguồn tài liệu.
- Lĩnh vực/chỉ tiêu/chủ đề.
- Nhóm đối tượng và hồ sơ trách nhiệm.
- Số câu, loại câu, độ khó, mức độ nhận thức.
- Tùy chọn bắt buộc evidence/giải thích/kiểm tra công thức.

### 18.4. Duyệt câu hỏi

Hai cột:

```text
Câu hỏi, đáp án, giải thích | Chứng cứ được tô sáng trong tài liệu nguồn
```

Thao tác: duyệt, sửa, từ chối, viết lại bằng AI, chuyển chủ đề.

### 18.5. Tạo đề

- Chọn đề cố định hoặc ma trận.
- Hiển thị số câu khả dụng theo từng điều kiện.
- Cảnh báo thiếu câu, không tự nới lọc.
- Xem trước cơ cấu và đề mẫu.

### 18.6. Làm bài

- Đồng hồ dựa trên deadline server.
- Autosave và chỉ báo đã lưu.
- Điều hướng câu, đánh dấu xem lại.
- Không gửi đáp án đúng/giải thích trong payload trước khi được phép công bố.

### 18.7. Báo cáo

- Cards tổng quan.
- Bộ lọc đối tượng/tổ/chủ đề.
- Bảng cá nhân.
- Phân tích câu hỏi.
- Chọn người và tạo đề thi lại.
- Xuất Excel.

## 19. API dự kiến

Tên endpoint Flask là đề xuất và phải được thêm vào route policy khi triển khai.

### 19.1. Kho tri thức

```text
GET    /api/training/knowledge
POST   /api/training/knowledge
GET    /api/training/knowledge/<id>
POST   /api/training/knowledge/<id>/versions
POST   /api/training/knowledge/<id>/analyze
GET    /api/training/knowledge/<id>/issues
POST   /api/training/knowledge/issues/<id>/resolve
```

### 19.2. Câu hỏi và AI

```text
POST   /api/training/generation-jobs
GET    /api/training/generation-jobs/<id>
GET    /api/training/questions
POST   /api/training/questions
PATCH  /api/training/questions/<id>
POST   /api/training/questions/<id>/approve
POST   /api/training/questions/<id>/publish
POST   /api/training/questions/<id>/rewrite
```

### 19.3. Đề và kỳ thi

```text
GET    /api/training/blueprints
POST   /api/training/blueprints
POST   /api/training/blueprints/<id>/validate-supply
GET    /api/training/exams
POST   /api/training/exams
POST   /api/training/exams/<id>/assignments
POST   /api/training/exams/<id>/open
POST   /api/training/exams/<id>/close
POST   /api/training/exams/<id>/finalize
```

Ba endpoint blueprint là pha mở rộng; MVP bắt đầu bằng template cố định.

### 19.4. Bài làm và báo cáo

```text
POST   /api/training/assignments/<id>/attempts
GET    /api/training/attempts/<id>
PUT    /api/training/attempts/<id>/responses/<item_id>
POST   /api/training/attempts/<id>/submit
GET    /api/training/exams/<id>/report
GET    /download/training/exams/<id>/report.xlsx
POST   /api/training/exams/<id>/retake-proposal
```

### 19.5. HTTP và bảo vệ ghi dữ liệu

- Tất cả route không public; auth chung của dashboard bảo vệ.
- Mọi POST/PUT/PATCH/DELETE dùng `csrf_protect`.
- API kiểm tra quyền module ở server.
- Validate content type, kích thước file và JSON.
- Các thao tác open/close/finalize/submit phải idempotent.

### 19.6. Contract danh sách, version và lỗi

Endpoint danh sách dùng contract chung:

```text
?page=1&page_size=25&sort=-created_at&status=published
&domain_code=quality&topic_code=repair_time&q=CCCO
```

- `page_size` mặc định 25, tối đa 100.
- Sort chỉ chấp nhận allowlist field.
- `PATCH` dùng `expected_version` hoặc `If-Match`; version sai trả `409 VERSION_CONFLICT`.
- Không trả toàn bộ ngân hàng câu hỏi nếu thiếu pagination.

Lỗi nghiệp vụ có mã ổn định:

```json
{
  "error": {
    "code": "ATTEMPT_ALREADY_ACTIVE",
    "message": "Bài làm đã được bắt đầu.",
    "details": {}
  }
}
```

Các mã tối thiểu:

- `ASSIGNMENT_NOT_FOUND`
- `ATTEMPT_ALREADY_ACTIVE`
- `ATTEMPT_ALREADY_COMPLETED`
- `ATTEMPT_EXPIRED`
- `EXAM_NOT_OPEN`
- `QUESTION_SUPPLY_INSUFFICIENT`
- `VERSION_CONFLICT`
- `DOCUMENT_HAS_BLOCKING_ISSUES`
- `PERMISSION_SCOPE_DENIED`

### 19.7. DTO theo đối tượng sử dụng

Phải có serializer/DTO riêng:

```text
AttemptItemLearnerView
AttemptItemReviewView
AttemptItemManagerView
```

Learner DTO không chứa `correct_option_ids`, explanation, distractor rationales, evidence hoặc scoring metadata có thể suy ra đáp án. Không dùng cách serialize entity đầy đủ rồi xóa field trước khi trả.

### 19.8. Contract tối thiểu của luồng làm bài

Start attempt thành công hoặc retry cùng assignment trả cùng attempt đang active:

```json
{
  "attempt_id": "att-...",
  "status": "active",
  "started_at": "2026-07-24T01:00:00Z",
  "deadline_at": "2026-07-24T01:25:00Z",
  "server_now": "2026-07-24T01:00:01Z",
  "attempt_version": 1
}
```

Load attempt trả learner DTO, thứ tự đã snapshot và response/revision hiện hành, không trả dữ liệu chấm:

```json
{
  "attempt_id": "att-...",
  "status": "active",
  "deadline_at": "2026-07-24T01:25:00Z",
  "server_now": "2026-07-24T01:05:00Z",
  "items": [
    {
      "item_id": "atti-...",
      "sequence": 1,
      "type": "single_choice",
      "stem": "...",
      "options": [
        {"id": "A", "text": "..."},
        {"id": "B", "text": "..."}
      ],
      "response": {
        "selected_option_ids": ["B"],
        "client_revision": 7
      }
    }
  ]
}
```

Autosave trả revision thực tế trên server; nếu request cũ đến muộn, `accepted = false` nhưng không coi là lỗi mất bài:

```json
{
  "accepted": false,
  "stored_response": {
    "selected_option_ids": ["B"],
    "client_revision": 8
  }
}
```

Submit/retry trả cùng result gốc:

```json
{
  "attempt_id": "att-...",
  "status": "submitted",
  "submitted_at": "2026-07-24T01:20:00Z",
  "result": {
    "score": 17,
    "maximum_score": 20,
    "percent": 85,
    "passed": true
  },
  "answers_released": false
}
```

## 20. CLI

CLI dùng cùng service và schema với UI:

```bash
python3 -m training.cli knowledge-import \
  --input input/c1_1.txt \
  --code C1.1 \
  --title "Chất lượng sửa chữa thuê bao BRCĐ"

python3 -m training.cli knowledge-analyze \
  --code C1.1 \
  --output /tmp/c1_1_knowledge.json

python3 -m training.cli questions-generate \
  --knowledge-code C1.1 \
  --audience nvkt \
  --count 20 \
  --generation-plan configs/quiz/c1_1_nvkt_generation.json \
  --output /tmp/c1_1_questions.json

python3 -m training.cli questions-validate /tmp/c1_1_questions.json

python3 -m training.cli questions-import \
  /tmp/c1_1_questions.json \
  --status draft
```

Chế độ bắt buộc:

- `--output-only`: không ghi DB;
- `--import-draft`: nhập nhưng không phát hành;
- không có tùy chọn AI sinh xong tự publish trong MVP.

Entry point mới phải `import runtime_limits` trước pandas/numpy hoặc thư viện có thể kéo chúng vào.

## 21. Mô hình dữ liệu

### 21.1. Danh mục và phân quyền

```text
training_instance_metadata
training_domains
training_categories
training_indicators
training_topics
training_competencies
training_audiences
training_services
training_tags
training_user_roles
training_user_audiences
```

### 21.2. Kho tri thức

```text
knowledge_documents
knowledge_document_versions
knowledge_blocks
knowledge_issues
knowledge_rules
knowledge_document_topics
knowledge_document_audiences
knowledge_block_topics
```

### 21.3. Kho câu hỏi

```text
question_items
question_versions
question_options
question_sources
question_reviews
question_topics
question_indicators
question_audiences
question_competencies
question_tags
ai_generation_jobs
ai_generation_batches
```

### 21.4. Đề và thi

```text
exam_blueprints
exam_blueprint_sections
exam_templates
exam_template_items
exam_events
exam_variants
exam_assignments
exam_attempts
exam_attempt_items
exam_attempt_options
exam_responses
exam_results
exam_result_topics
exam_score_adjustments
exam_report_snapshots
```

### 21.5. Audit

```text
training_audit_log
```

Audit tối thiểu:

- actor username và unit code;
- action;
- entity type/id;
- before/after JSON khi phù hợp;
- timestamp;
- request ID/IP/user agent ở các thao tác quản trị quan trọng.

### 21.6. Ràng buộc và index tối thiểu

Mỗi `training.db` thuộc đúng một instance. Bảng `training_instance_metadata` có một dòng chứa `unit_code`, schema version và thời điểm tạo; startup/migration từ chối chạy nếu `unit_code` không khớp `DASHV4_UNIT_CODE`.

Unique constraint/index tối thiểu:

```text
training_domains(code) UNIQUE
training_topics(code) UNIQUE
knowledge_document_versions(document_id, version_number) UNIQUE
knowledge_blocks(document_version_id, extraction_revision, block_id) UNIQUE
question_versions(question_item_id, version_number) UNIQUE
question_options(question_version_id, option_code) UNIQUE
exam_assignments(exam_event_id, username, audience_code) UNIQUE
exam_attempts(assignment_id) UNIQUE
exam_attempt_items(attempt_id, sequence_number) UNIQUE
exam_responses(attempt_id, attempt_item_id) UNIQUE
exam_results(attempt_id) UNIQUE
exam_report_snapshots(exam_event_id, revision) UNIQUE
```

Vì MVP chỉ có một attempt record cho mỗi assignment, attempt `invalidated` không được thay thế trong cùng assignment. Người vận hành tạo assignment thi lại mới và có thể lưu `retake_of_assignment_id`.

Index truy vấn tối thiểu:

- assignment theo `(exam_event_id, username, status)`;
- attempt theo `(assignment_id, status)`;
- response theo `attempt_id`;
- question version theo publication status và classification;
- question source theo `(document_version_id, block_id)`;
- report snapshot theo `(exam_event_id, revision)`;
- audit theo `(entity_type, entity_id, created_at)`.

### 21.7. Snapshot người được giao

`exam_assignments` giữ snapshot tối thiểu:

```text
username
display_name
team_code
team_name
organization_code/name
audience_code
user_source_updated_at
assigned_at
```

Báo cáo lịch sử dùng snapshot này, không join động hoàn toàn với `users.xlsx` hiện tại.

## 22. SQLite và vận hành

### 22.1. Kết nối

- Write connection bật `PRAGMA journal_mode=WAL`.
- `PRAGMA busy_timeout=5000`.
- Foreign keys bật cho mọi connection ghi.
- Transaction ngắn; không gọi AI trong transaction.
- Read connection có `row_factory = sqlite3.Row`.
- Mỗi request/worker dùng connection riêng; không dùng global connection chung giữa thread/process.

### 22.2. Schema migration

Không đặt toàn bộ `CREATE TABLE` trong route. Cần một cơ chế migration có version, tối thiểu:

```text
training_schema_migrations(version, applied_at)
```

Startup chỉ kiểm tra/migrate có khóa tiến trình phù hợp; CLI có lệnh migrate chủ động.

- Migration dùng file lock per-instance để nhiều Gunicorn worker không migrate đồng thời.
- Migration phải kiểm tra `training_instance_metadata.unit_code` trước khi thay schema.

### 22.3. Backup

- Backup `training.db` bằng SQLite backup API hoặc lệnh `.backup`.
- Backup kèm thư mục tài liệu nguồn.
- Không copy file DB đang ghi theo cách thô nếu không xử lý WAL.
- Kiểm tra phục hồi định kỳ.

### 22.4. Concurrency và idempotency

#### Tạo attempt

- Dùng `BEGIN IMMEDIATE` ngắn.
- Kiểm tra assignment, exam status/time và attempt hiện hữu trong cùng transaction.
- Tạo attempt/snapshot rồi commit; unique `exam_attempts(assignment_id)` là lớp bảo vệ race cuối cùng.
- Hai request start đồng thời: một request tạo thành công, request còn lại trả attempt hiện hành với `ATTEMPT_ALREADY_ACTIVE` hoặc response idempotent tương đương.

#### Autosave

Update theo compare-and-set revision:

```sql
UPDATE exam_responses
SET selected_option_ids_json = ?,
    client_revision = ?,
    answered_at = ?
WHERE attempt_id = ?
  AND attempt_item_id = ?
  AND client_revision < ?;
```

Insert lần đầu dùng unique `(attempt_id, attempt_item_id)` và xử lý conflict bằng cùng quy tắc revision. Request revision cũ không ghi đè đáp án mới.

#### Submit và scoring

- Transition attempt dùng compare-and-set `WHERE status = 'active'`.
- Chấm từ snapshot và tạo result trong cùng transaction ngắn.
- Request submit lặp trả result hiện có.
- Không cập nhật counter tổng hợp dùng chung theo từng submit; báo cáo dùng aggregate query hoặc cache có thể tái tạo.

#### State machine và thời gian đã triển khai

- Kỳ thi chuyển chính xác `draft -> ready -> open -> closed`; `draft`/`ready -> cancelled`. `closed` và `cancelled` không mở lại, close chỉ nhận `open` (retry close khi đã `closed` chỉ chạy recovery).
- Attempt terminal không trở lại `active`. Close trước deadline dùng `administratively_submitted`/`exam_closed`; close tại hoặc sau deadline dùng `timed_out`/`timeout`. Khi finalize close kỳ thi open đã hết `end_at`, attempt active cũng là `timed_out`, kể cả deadline cá nhân muộn hơn.
- Autosave dùng thứ tự precedence: attempt không active -> `ATTEMPT_ALREADY_COMPLETED`; exam không open hoặc `now >= end_at` -> `EXAM_NOT_OPEN`; sau đó `now >= deadline_at` -> `ATTEMPT_EXPIRED`. Điều này không nhận response tại đúng deadline.
- Submit sau deadline không nhận đáp án mới và chuyển `timed_out`, chấm response đã persist. Submit retry trả result gốc trừ attempt đã đóng hành chính, khi đó learner nhận `ATTEMPT_ALREADY_COMPLETED`.

#### Audience chain đã triển khai

- Question version có mapping audience chỉ được đưa vào template cùng audience; template, exam và assignment phải cùng audience.
- User có ít nhất một `training_user_audiences` chỉ được giao khi có mapping khớp; user không có mapping vẫn hợp lệ để hỗ trợ dữ liệu người dùng chưa cấu hình, và assignment snapshot audience được chọn.

#### Finalize/report

- Compare-and-set trạng thái kỳ thi.
- `close` compare-and-set kỳ thi trước để chặn ghi mới, sau đó kết thúc từng active attempt bằng các transaction ngắn, idempotent; không giữ một transaction cho toàn bộ kỳ thi.
- Unique `(exam_event_id, revision)` ngăn snapshot trùng.
- Finalize lặp trả revision đã tạo nếu không có adjustment mới.

#### Recovery và chặn report đã triển khai

- Close commit trạng thái kỳ thi trước, rồi xử lý từng attempt trong transaction độc lập. Summary gồm `processed_attempt_ids`, `already_completed_ids`, `failed_attempts`; retry close/finalize có thể tiếp tục các attempt còn active.
- Lỗi toàn vẹn snapshot của một attempt được ghi audit và cô lập, không rollback attempt khác. Finalize không tạo report khi summary còn `failed_attempts`; lỗi trả `blocking_attempts` và `recovery_summary` để vận hành sửa dữ liệu trước khi chốt lại.

#### Retry SQLite

- Chỉ retry lỗi `SQLITE_BUSY/locked` cho thao tác idempotent hoặc có idempotency key/revision.
- Retry tối đa theo cấu hình với backoff ngắn có jitter; hết retry trả mã lỗi ổn định và không báo thành công giả.
- Không giữ transaction trong lúc gọi AI, render Excel hoặc thực hiện I/O file lớn.

## 23. Cấu hình

Các biến dự kiến:

```text
DASHV4_TRAINING_DB_PATH
DASHV4_TRAINING_FILES_DIR
DASHV4_TRAINING_EXPORT_DIR
DASHV4_TRAINING_AI_ENABLED
DASHV4_TRAINING_AI_PROVIDER
DASHV4_TRAINING_AI_MODEL
DASHV4_TRAINING_MAX_UPLOAD_MB
DASHV4_TRAINING_GENERATION_TIMEOUT_SECONDS
```

- API key lấy từ biến môi trường/secret store, không lưu trong DB hoặc commit.
- Không trả API key xuống browser.
- Model/provider metadata được lưu theo generation batch để audit.

## 24. Bảo mật và toàn vẹn

- Không trả `correct_option_ids`, scoring rule hoặc explanation trước thời điểm được công bố.
- Trang/API làm bài và response chứa đề dùng `Cache-Control: no-store`.
- Không tin `score`, `elapsed_seconds` hoặc deadline từ client.
- Server kiểm tra người dùng có assignment tương ứng với attempt.
- MVP quy định `exam_manager` có phạm vi toàn instance TTVT; learner chỉ xem dữ liệu cá nhân. Scope theo tổ chỉ bổ sung khi có yêu cầu nghiệp vụ chính thức.
- JSON API dùng session cookie cùng origin, CSRF token trong header `X-CSRF-Token`; không bật CORS cross-origin trong MVP.
- Giữ cookie `HttpOnly`, `SameSite` theo cấu hình app và bật `Secure` tại môi trường HTTPS production.
- File upload đổi tên nội bộ, không dùng trực tiếp tên file làm path.
- Chặn path traversal và loại file không cho phép.
- Kiểm tra MIME/signature, giới hạn dung lượng và giới hạn độ phức tạp khi đọc DOCX/XLSX/PDF để tránh file giả mạo hoặc zip bomb.
- Nội dung HTML từ tài liệu/AI phải sanitize trước khi render.
- Câu hỏi đã dùng trong thi không bị xóa cứng.
- Mọi điều chỉnh điểm, invalidation hoặc thay trạng thái kỳ thi đều có audit.
- Dữ liệu cá nhân trong report chỉ hiển thị theo phạm vi quyền.
- Endpoint download kiểm tra auth/scope giống API report; file export không đặt trong static/public và dùng tên nội bộ không đoán được.
- Mọi text từ AI/tài liệu/người dùng được escape trước khi xuất Excel nếu bắt đầu bằng `=`, `+`, `-`, `@`.

## 25. Yêu cầu phi chức năng

- 40–50 người có thể làm/nộp bài trong cùng kỳ thi trên một instance.
- Autosave phản hồi bình thường trong điều kiện mạng nội bộ/Internet hợp lý.
- Submit lặp do retry không tạo kết quả trùng.
- Sau khi nộp hoặc hết giờ, client không thể sửa đáp án.
- Một đề đã giao có thể tái hiện chính xác sau này.
- Một báo cáo đã chốt không thay đổi khi câu hỏi/tài liệu có phiên bản mới.
- UI dùng tiếng Việt và responsive theo shell hiện tại.
- Không phụ thuộc `report_history.db` hoặc contract lọc ngày của docs 09.

## 26. Tiêu chí nghiệm thu MVP

### 26.1. Kho tri thức và AI

- Dán được văn bản và lưu tài liệu nháp.
- Phân loại theo lĩnh vực/chỉ tiêu/chủ đề/đối tượng.
- Phân đoạn có ID ổn định trong `document_version + extraction_revision`.
- Worker AI trả Question Batch JSON `single_choice` đúng schema.
- Hiển thị và xử lý issue chưa rõ.
- Câu hỏi có evidence và không được publish khi thiếu điều kiện duyệt.
- Ít nhất một calculation rule mẫu được backend xác minh lại qua registry, không thực thi expression AI.

### 26.2. Kho câu hỏi và đề

- Soạn/sửa/duyệt/phát hành câu hỏi.
- Lọc theo nhiều chiều.
- Tạo template cố định từ question versions đã publish.
- Template/kỳ thi ghi rõ một nhóm đối tượng; có thể tạo kỳ thi riêng cho các nhóm khác nhau.
- Blueprint, supply validation và nhiều variant trong một kỳ thi không bắt buộc ở MVP.

### 26.3. Tổ chức thi

- Người vận hành chọn danh sách người thi và thời gian.
- Tới giờ nhưng chưa bấm mở thì không thi được.
- Người không được giao không mở được bài.
- Trộn câu/phương án và lưu snapshot.
- Một assignment chỉ tạo được một attempt hợp lệ.
- Autosave revision, submit idempotent và deadline hoạt động theo server time.
- Chấm đúng `single_choice`; `multiple_choice` có thể bổ sung ngay sau khi engine một đáp án ổn định.

### 26.4. Báo cáo

- Có tổng quan, danh sách cá nhân, theo nhóm đối tượng/chủ đề và câu sai nhiều.
- Chốt kỳ thi tạo snapshot báo cáo.
- Xuất Excel.
- Chọn cá nhân chưa đạt và tạo kỳ thi/assignment thi lại mới; không tái sử dụng assignment cũ và không tự mở kỳ thi.

### 26.5. CLI

- Import tài liệu, phân tích, sinh JSON, validate và import draft.
- UI và CLI tạo dữ liệu cùng schema và qua cùng validator.

### 26.6. Nghiệm thu tải và phục hồi MVP

Trên cấu hình production mục tiêu được ghi trong biên bản test:

- Smoke hiện có chỉ là local in-process với SQLite connection độc lập theo thread: 50 autosave đồng thời trên một item giữ revision cao nhất, 50 submit đồng thời trên một attempt tạo đúng một result/audit terminal, và race close/autosave/submit giữ đúng một result terminal. Không suy diễn kết quả này thành kiểm thử scheduler đa process hoặc tải production.
- 50 người autosave mỗi 10 giây trong 10 phút, không mất hoặc đảo revision hợp lệ, vẫn là tiêu chí nghiệm thu production cần chạy và ghi biên bản riêng.
- 50 request submit trong vòng 5 giây, mỗi attempt có đúng một trạng thái kết thúc và một result gốc, vẫn cần đo trên cấu hình production mục tiêu.
- Không có result/report snapshot trùng khi retry hoặc hai worker xử lý đồng thời.
- Restart web app khi attempt đang active không làm mất response đã autosave.
- Worker AI bị dừng giữa job có thể được claim lại sau lease timeout hoặc chuyển failed có thể retry.
- Khôi phục từ backup tái hiện được template, attempt snapshot, result và report revision đã chốt.

Mục tiêu latency cụ thể chỉ chốt sau khi ghi rõ CPU/RAM/ổ đĩa, số Gunicorn worker và reverse proxy; không dùng một con số P95 không gắn môi trường.

## 27. Kiểm thử bắt buộc

- Unit test JSON Schema và semantic validator.
- Unit test công thức/điểm biên của câu tính toán.
- Unit test chọn câu theo ma trận và báo thiếu nguồn cung.
- Unit test state transition tài liệu/câu hỏi/kỳ thi/assignment/attempt/result/report.
- Unit test scoring và ngưỡng theo chủ đề.
- Route test auth, CSRF, permission và assignment ownership.
- Test idempotency cho autosave, submit, open/close/finalize.
- Test hai worker cùng submit/chốt với SQLite WAL.
- Test không lộ đáp án trong API làm bài.
- Test snapshot không thay đổi khi câu hỏi nguồn có phiên bản mới.
- Test export Excel.

Chạy toàn bộ bằng lệnh chuẩn của repo:

```bash
python3 -m pytest tests/
```

## 28. Lộ trình triển khai đề xuất

### Pha 0 — Khóa thiết kế kỹ thuật tối thiểu

- Domain invariants và state transition tests.
- Time policy, snapshot schema và error contract.
- Data dictionary/ERD với PK/FK/unique/index/immutable fields.
- Transaction boundary và idempotency design.

### Pha 1 — Nền, kho tri thức và AI vertical slice

- DB/migration/config.
- Danh mục, đối tượng, quyền module.
- Dán văn bản, phân loại, phiên bản, extraction revision và blocks.
- AI worker sinh `single_choice` có evidence.
- Review/publish câu hỏi và JSON/CLI validate.

### Pha 2 — Engine thi cố định

- Fixed template cho một audience mỗi kỳ thi.
- Assignment snapshot người dùng, một assignment–một attempt.
- Attempt snapshot, randomization, autosave revision, deadline, submit và scoring single-choice.
- Learner DTO không lộ đáp án.

### Pha 3 — Báo cáo và thi lại thủ công

- Dashboard kết quả cơ bản.
- Finalize, report snapshot/revision và Excel an toàn.
- Tạo kỳ thi/assignment thi lại mới từ danh sách chưa đạt.

### Pha 4 — Mở rộng sau MVP

- Blueprint và supply validation.
- Nhiều variant/audience trong cùng kỳ thi.
- Multiple-choice, topic threshold và retake proposal nâng cao.
- Import PDF/DOCX/XLSX, duplicate similarity nâng cao và phân tích chất lượng câu hỏi.

## 29. Các quyết định cần xác nhận trước khi code

1. Nguồn ánh xạ `username -> họ tên -> tổ -> nhóm đối tượng`: mở rộng import Excel hay có nguồn chính thức khác.
2. Mô tả trách nhiệm chính thức của nhóm B2A và các nhóm ngoài NVKT/tổ trưởng.
3. Nhà cung cấp/model AI, chính sách gửi tài liệu nội bộ ra API và nơi giữ API key.
4. Có lưu file gốc lâu dài trong instance hay chỉ lưu nội dung đã trích xuất kèm bản backup ngoài.
5. Chính sách công bố điểm/đáp án mặc định sau thi.
6. Mẫu Excel/PDF báo cáo chính thức nếu đơn vị đã có biểu mẫu.
7. Cách tính điểm câu nhiều đáp án nếu triển khai sau MVP.
8. Thời hạn lưu file gốc, raw AI response, audit, attempt snapshot và report revision.
9. Chính sách backup ngoài host và người chịu trách nhiệm kiểm tra restore.
10. Có yêu cầu giới hạn `exam_manager` theo tổ hay mặc định toàn TTVT như MVP.

Các quyết định 1–5 phải chốt trước tích hợp AI/thi production. Các quyết định còn lại có mặc định MVP trong spec nhưng nên được đơn vị xác nhận trước nghiệm thu.

## 30. Đồng bộ tài liệu khi triển khai

Khi bắt đầu đưa code vào vận hành phải cập nhật tối thiểu:

- `docs/00-doc-index.md`;
- `docs/04-mapping-route-va-du-lieu.md` với nguồn `training.db`, `supports_date = n/a`;
- `docs/08-trang-thai-thuc-thi.md`;
- `route_policy.py` và `deploy/units.yaml` nếu cần bật/tắt theo instance;
- tài liệu vận hành backup/migration/khôi phục `training.db`.

Module không đọc `report_history.db`, vì vậy không áp dụng date-filter contract của `docs/09-nguyen-tac-loc-ngay-report-history.md`.
