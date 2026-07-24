# Đào tạo & sát hạch - Delta thiết kế tính toàn vẹn engine thi

## Phạm vi

Tài liệu này bổ sung cho `2026-07-24-dao-tao-sat-hach-design.md` phiên bản 1.1. Phạm vi chỉ gồm tính toàn vẹn engine thi và RBAC MVP sau commit `4b858ca`; không mở rộng blueprint, QTI, proctoring hoặc nhiều attempt.

## Trộn đề tái lập được

- Khi khởi tạo attempt, server tạo và lưu `random_seed`; client không được cung cấp seed.
- Thứ tự câu hỏi và phương án được tạo bằng RNG cục bộ khởi tạo từ seed snapshot. Không dùng RNG dùng chung của process.
- `shuffle_questions` chỉ thay đổi thứ tự câu nếu template/exam bật cờ; `shuffle_options` chỉ thay đổi thứ tự phương án nếu bật cờ.
- Mỗi câu dùng stream RNG xác định theo seed attempt và định danh snapshot câu để việc trộn phương án không phụ thuộc thứ tự thực thi ngoài ý muốn.
- Snapshot vẫn lưu thứ tự đã giao. Cùng seed, version thuật toán và input phải tái tạo đúng thứ tự này.

## Đóng và chốt kỳ thi

- Close đổi exam sang `closed` trước khi xử lý attempt, để chặn start attempt và autosave mới.
- Mỗi attempt `active` được kết thúc trong transaction ngắn từ các response đã lưu: trạng thái `administratively_submitted`, `ended_reason = exam_closed_by_manager`, result chấm từ snapshot, và audit actor/thời điểm.
- Close lặp là idempotent; không chấm lại result đã tồn tại.
- Finalize chỉ hoàn tất transition còn dở, expire assignment chưa bắt đầu, và tạo hoặc trả report snapshot revision đã chốt. Không ghi đè report snapshot hiện có.

## Autosave an toàn

- `attempt_item_id` phải thuộc `attempt_id` đang ghi; khác attempt trả `ATTEMPT_ITEM_NOT_FOUND`.
- Mọi `selected_option_ids` phải tồn tại trong option snapshot của đúng attempt item; option lạ trả `INVALID_OPTION_SELECTION`.
- Với `single_choice`, payload có quá một option trả `SINGLE_CHOICE_REQUIRES_ONE_OPTION`.
- Validation xảy ra trước upsert revision. Input sai không được tạo hoặc thay đổi `exam_responses`.
- Revision cũ vẫn trả trạng thái response hiện có với `accepted = false`, không phải lỗi validation.

## Tính toàn vẹn template

- Template chỉ chứa question version tồn tại, `published`, không lặp, và khớp audience template nếu version khai báo audience.
- Service validation áp dụng cho toàn bộ đường tạo template. Database dùng foreign key đến question version và unique `(template_id, question_version_id)` nơi migration tương thích.
- Số lượng câu template được suy ra từ items đã lưu; không tin một số lượng do client khai báo.

## RBAC MVP

- `editor`: tạo tài liệu, enqueue generation job, import hoặc tạo question draft.
- `exam_manager`: approve/publish question; tạo template, exam, assignment; ready/open/close/finalize và xem report.
- `learner`: chỉ liệt kê assignment và đọc/ghi attempt thuộc username phiên hiện tại.
- `admin`: toàn quyền module.
- Permission kiểm tra ở route và service. DTO learner không chứa đáp án, explanation, evidence hoặc scoring metadata.

## Kiểm thử bắt buộc

- Reproducibility cho seed snapshot; hai attempt tuân thủ cờ trộn câu/phương án.
- Concurrent start và burst submit chỉ tạo một attempt/result.
- Close active attempt tạo result hành chính, chặn start/autosave, finalize giữ report snapshot bất biến.
- Autosave cross-attempt, option không thuộc snapshot và nhiều option cho single choice bị từ chối không ghi response.
- Route tests cho quyền được cấp/từ chối của bốn vai trò.
