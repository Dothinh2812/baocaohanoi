# Chỉ mục tài liệu

Tài liệu này là điểm vào nhanh cho toàn bộ hệ thống doc của `dashv4`.

Nếu vào repo ở phiên sau, hoặc cần sửa/thêm tính năng, bắt đầu từ đây thay vì mở ngẫu nhiên nhiều file.

## Đường đọc ngắn nhất

Đọc theo thứ tự này:

1. [08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
   Dùng để biết repo đang ở trạng thái nào ngay bây giờ.
2. [04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
   Dùng để xác định route/page đang đọc nguồn nào, trạng thái `supports_date`, và mức tương thích hiện tại.
3. [09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md)
   Dùng khi route đọc `report_history.db` và cần hiển thị theo ngày.
4. [03-hop-dong-du-lieu-api-transition.md](/home/vtst/dashv4/docs/03-hop-dong-du-lieu-api-transition.md)
   Dùng để hiểu cấu trúc DB runtime, metadata import, và contract dữ liệu từ `api_transition`.

## Khi cần làm gì thì đọc gì

### Muốn biết nên bắt đầu từ đâu

- [08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
- [04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)

### Muốn sửa một route/page cụ thể

- [04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
- [08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)

### Muốn thêm hỗ trợ hiển thị theo ngày

- [09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md)
- [03-hop-dong-du-lieu-api-transition.md](/home/vtst/dashv4/docs/03-hop-dong-du-lieu-api-transition.md)
- [04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)

### Muốn hiểu DB `report_history.db`

- [03-hop-dong-du-lieu-api-transition.md](/home/vtst/dashv4/docs/03-hop-dong-du-lieu-api-transition.md)
- [05-kien-truc-de-xuat.md](/home/vtst/dashv4/docs/05-kien-truc-de-xuat.md)

### Muốn lên kế hoạch triển khai tiếp

- [06-ke-hoach-trien-khai.md](/home/vtst/dashv4/docs/06-ke-hoach-trien-khai.md)
- [08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
- [04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)

### Muốn triển khai nhiều instance song song

- [10-trien-khai-multi-instance.md](/home/vtst/dashv4/docs/10-trien-khai-multi-instance.md)
- [07-trien-khai-van-hanh.md](/home/vtst/dashv4/docs/07-trien-khai-van-hanh.md)

## Danh sách tài liệu

### Tài liệu định hướng và phạm vi

- [01-muc-tieu-va-pham-vi.md](/home/vtst/dashv4/docs/01-muc-tieu-va-pham-vi.md)
- [02-hien-trang-dashv3.md](/home/vtst/dashv4/docs/02-hien-trang-dashv3.md)

### Tài liệu dữ liệu và mapping

- [03-hop-dong-du-lieu-api-transition.md](/home/vtst/dashv4/docs/03-hop-dong-du-lieu-api-transition.md)
- [04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
- [09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md)

### Tài liệu kiến trúc và triển khai

- [05-kien-truc-de-xuat.md](/home/vtst/dashv4/docs/05-kien-truc-de-xuat.md)
- [06-ke-hoach-trien-khai.md](/home/vtst/dashv4/docs/06-ke-hoach-trien-khai.md)
- [07-trien-khai-van-hanh.md](/home/vtst/dashv4/docs/07-trien-khai-van-hanh.md)
- [08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
- [10-trien-khai-multi-instance.md](/home/vtst/dashv4/docs/10-trien-khai-multi-instance.md)

## Quy tắc đồng bộ tài liệu

Mọi thay đổi liên quan tới:
- nguồn dữ liệu
- mapping route
- trạng thái hỗ trợ `date`
- nguyên tắc đọc `report_history.db`

thì phải cập nhật tối thiểu:
- [04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
- [08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)

Nếu thay đổi chuẩn chung của kiến trúc lọc ngày, cập nhật thêm:
- [09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md)
