# Mục tiêu và phạm vi

## Mục tiêu

Xây một dashboard mới tên `dashv4`, độc lập với `dashv3`, nhưng vẫn giữ tinh thần sử dụng của dashboard cũ:
- Giữ lại các route chính đang dùng trong `dashv3`.
- Giữ logic hiển thị và ý nghĩa nghiệp vụ của dữ liệu.
- Đổi lớp nguồn dữ liệu sang `SQLite` của dự án mới `baocaohanoi/api_transition`.
- Chuẩn bị sẵn nền để mở rộng dần các màn còn thiếu khi DB mới bổ sung đủ dữ liệu.

## Mục tiêu kỹ thuật

- Tách app khỏi phụ thuộc đọc file Excel ở request time.
- Tách app khỏi việc đọc trực tiếp dữ liệu từ nhiều repo anh em.
- Chuẩn hóa truy vấn qua một lớp repository/service đọc từ `report_history.db`.
- Chỉ dùng các `view` tiêu thụ ổn định khi có thể.
- Có khả năng chạy độc lập bằng cấu hình riêng và DB path riêng.

## Phạm vi của đợt 1

Đợt 1 nên tập trung các màn đã có contract dữ liệu khá rõ trong DB runtime Sơn Tây:
- `/cau-hinh-tu-dong`
- `/ttvt-son-tay-tong-hop`
- `/chatluong`
- `/giahan`
- `/tiepthi`
- `/thuc-tang-ngung-psc`
- `/xac-minh-tam-dung`
- `/thuhoi`
- một phần `/kpi` và `/kpi-nvkt-bchn`

## Ngoài phạm vi của đợt 1

Các màn sau chưa nên cam kết trong vòng đầu vì hiện chưa có contract dữ liệu tương ứng trong `api_transition` runtime Sơn Tây hoặc còn phụ thuộc nguồn khác:
- `/brcd`
- `/pttb`
- `/i15`
- `/i15k2`
- `/shc-processing`
- `/quangchudong`
- `/su_co_sa`
- `/ton-kho-vat-tu`
- `/Tong_hop_tien`
- `/tra-cuu-nhanh-vat-tu`
- các API thống kê ticket

## Giả định đang dùng cho tài liệu này

- Đơn vị triển khai mục tiêu trước mắt là Sơn Tây.
- Nguồn DB chính là:
  `/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db`
- Thời điểm kiểm tra dữ liệu thực tế là `2026-04-20`.
- `dashv4` có thể tiếp tục dùng Flask để giảm chi phí chuyển đổi vì `dashv3` đang là Flask app server-rendered shell + page JS.

## Tiêu chí hoàn thành

Một bản `dashv4` đạt yêu cầu nền tảng khi:
- Chạy độc lập, không cần import `dashv3`.
- Các route đã chọn vẫn truy cập bằng URL cũ.
- Dữ liệu hiển thị lấy từ DB mới, không đọc Excel trực tiếp.
- Các màn chưa sẵn sàng được khóa hoặc gắn nhãn rõ ràng, không trả dữ liệu sai.
