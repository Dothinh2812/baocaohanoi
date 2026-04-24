# Kế hoạch triển khai

## Nguyên tắc thực hiện

- Làm theo lát cắt dọc: một route hoàn chỉnh từ query đến UI.
- Không di cư toàn bộ dashboard trong một lần.
- Mỗi route chỉ bật khi đã có dữ liệu thật trong DB mới.
- Dùng chiến lược fork có kiểm soát: copy `dashv3` trước, rồi thay backend theo từng route.

## Pha 0: fork codebase

Mục tiêu:
- copy codebase `dashv3` sang `dashv4`
- giữ nguyên shell vận hành ban đầu
- loại trừ artifact runtime không cần thiết
- giữ lại `docs` đã viết cho `dashv4`

Đầu ra:
- `dashv4` có khung Flask đầy đủ ngay từ đầu
- có thể sửa trên codebase thật thay vì tiếp tục tài liệu hóa trên repo trống

## Pha 1: cấu hình và data layer nền

Mục tiêu:
- đổi nhận diện cấu hình sang `dashv4`
- thêm cấu hình DB mới
- thêm helper kết nối SQLite read-only
- tạo lớp repository/service ban đầu

Đầu ra:
- app có thể kết nối được tới DB runtime mới
- chưa cần xong toàn bộ route

## Pha 2: các route gần như sẵn sàng

Ưu tiên theo thứ tự:
1. `/ttvt-son-tay-tong-hop`
2. `/cau-hinh-tu-dong`
3. `/tiepthi`
4. `/thuhoi`

Lý do:
- contract dữ liệu rõ
- scope hẹp
- giá trị xác nhận kiến trúc cao

## Pha 3: các route cần adapter

Ưu tiên:
1. `/chatluong`
2. `/giahan`
3. `/thuc-tang-ngung-psc`
4. `/kpi`
5. `/kpi-nvkt-bchn`

Việc cần làm:
- đọc shape JSON cũ của `dashv3`
- viết serializer để tái tạo shape đó từ SQLite views
- xác nhận chênh lệch do nguồn mới

Trạng thái cập nhật:
- pha này đã được thực hiện phần lớn cho `chatluong`, `giahan`, `thuc-tang-ngung-psc`, `kpi`, `kpi-nvkt-bchn`
- phần còn lại là gom adapter rải trong blueprint về lớp `services/serializers`

## Pha 4: route còn thiếu contract

Ưu tiên nghiên cứu sau:
- `/tam-dung-khoi-phuc`

Điều kiện vào pha này:
- xác định được có nên tạo `view` mới trong `api_transition`
- hoặc chấp nhận query từ view thấp hơn mà vẫn đủ ổn định

Trạng thái cập nhật:
- `/xac-minh-tam-dung` đã chuyển được từ DB mới
- `/tam-dung-khoi-phuc` đã bị ngắt khỏi Excel cũ và đổi sang trạng thái chờ contract dữ liệu

## Pha 5: route disabled pending data

Giữ danh sách treo:
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
- thống kê ticket

Nguyên tắc bắt buộc của pha này:
- không để API cũ tiếp tục đọc Excel/file riêng trong khi page đã pending
- mỗi endpoint bị ngắt phải trả `required_display_contract` đủ cụ thể để làm backlog dữ liệu

## Thứ tự kỹ thuật cho từng route

1. xác nhận view nguồn
2. viết query read-only
3. viết service chuẩn hóa
4. viết serializer hoặc adapter
5. dựng API JSON
6. gắn page template và JS
7. so đối chiếu với `dashv3`
8. thêm smoke test
9. nếu không đủ dữ liệu DB thì khóa route cũ và ghi contract hiển thị cần bổ sung

## Những việc không nên làm

- Không copy `dashv3` một cách mù quáng rồi đổi path hàng loạt.
- Không mở route bằng dữ liệu giả.
- Không query bảng raw nếu đã có consumer view.
- Không trộn chung logic query, format và HTML trong một route handler.
- Không giữ endpoint Excel cũ “cho tạm dùng” nếu mục tiêu của route đã chuyển sang DB mới.

## Điều kiện sẵn sàng để bắt đầu code

- DB path chuẩn cho môi trường chạy
- auth hiện tại có được giữ cho đợt này hay không
- danh sách route đủ dữ liệu để bật từ DB
- danh sách route phải trả `pending/501` cùng `required_display_contract`
