# dashv4

`dashv4` là repo khởi tạo cho dashboard mới, độc lập với `dashv3`, nhưng được triển khai theo chiến lược fork có kiểm soát từ `dashv3`.

Mục tiêu của app mới:
- Giữ lại các route và cách hiển thị dữ liệu quan trọng từ `dashv3`.
- Dùng lại codebase `dashv3` làm khung ban đầu để đi nhanh.
- Đổi nguồn đọc dữ liệu sang SQLite của `baocaohanoi/api_transition`.
- Tách dần khỏi mô hình phụ thuộc file Excel trực tiếp và phụ thuộc chéo nhiều repo.

Trạng thái hiện tại:
- Đã đọc `dashv3` để nắm inventory route, page, API và nguồn dữ liệu cũ.
- Đã đọc `baocaohanoi/api_transition`, đặc biệt runtime Sơn Tây và `report_history.db`.
- Đã chốt chiến lược triển khai: fork có kiểm soát từ `dashv3`.
- Trong quá trình triển khai, `dashv4` sẽ được copy codebase từ `dashv3` rồi thay backend theo từng route.

## Bắt đầu từ đâu

Nếu vào repo ở phiên sau hoặc cần sửa/thêm tính năng, đọc theo đúng thứ tự này:

1. [docs/08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
   Đây là ảnh chụp trạng thái mới nhất của repo: route nào đã chuyển, route nào còn `501`, route nào đang dùng DB mới.
2. [docs/04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
   Dùng để biết route/page đang map vào nguồn dữ liệu nào, `supports_date` đang ở mức nào, và status hiện tại là gì.
3. [docs/09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md)
   Đây là chuẩn kỹ thuật bắt buộc cho mọi route đọc `report_history.db` nếu cần hiển thị theo ngày.
4. [docs/03-hop-dong-du-lieu-api-transition.md](/home/vtst/dashv4/docs/03-hop-dong-du-lieu-api-transition.md)
   Dùng để hiểu cấu trúc DB runtime, metadata import, view consumer, và contract dữ liệu từ `api_transition`.
5. [docs/05-kien-truc-de-xuat.md](/home/vtst/dashv4/docs/05-kien-truc-de-xuat.md)
   Dùng khi cần quyết định hướng kiến trúc hoặc refactor.
6. [docs/06-ke-hoach-trien-khai.md](/home/vtst/dashv4/docs/06-ke-hoach-trien-khai.md)
   Dùng khi cần lên batch triển khai tiếp theo.

## Khi sửa hoặc thêm tính năng

Checklist ngắn trước khi code:

1. Xác định route/page cần sửa trong [docs/04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md).
2. Kiểm tra route đó đang dùng `report_history.db` hay nguồn khác.
3. Nếu dùng `report_history.db`, bắt buộc đối chiếu [docs/09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md).
4. Nếu route đã hỗ trợ hoặc chuẩn bị hỗ trợ `date`, không được tự ý dùng `view ... moi_nhat` làm nguồn chính.
5. Sau khi triển khai, cập nhật doc cùng lúc:
   - [docs/08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
   - [docs/04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
   - [docs/09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md) nếu thay đổi chuẩn chung

## Tài liệu chính

- [docs/01-muc-tieu-va-pham-vi.md](/home/vtst/dashv4/docs/01-muc-tieu-va-pham-vi.md)
- [docs/02-hien-trang-dashv3.md](/home/vtst/dashv4/docs/02-hien-trang-dashv3.md)
- [docs/03-hop-dong-du-lieu-api-transition.md](/home/vtst/dashv4/docs/03-hop-dong-du-lieu-api-transition.md)
- [docs/04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
- [docs/05-kien-truc-de-xuat.md](/home/vtst/dashv4/docs/05-kien-truc-de-xuat.md)
- [docs/06-ke-hoach-trien-khai.md](/home/vtst/dashv4/docs/06-ke-hoach-trien-khai.md)
- [docs/07-trien-khai-van-hanh.md](/home/vtst/dashv4/docs/07-trien-khai-van-hanh.md)
- [docs/08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
- [docs/09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md)

## Tài liệu theo mục đích

- Muốn biết repo đang ở trạng thái nào: [docs/08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
- Muốn biết route nào đang dùng nguồn gì: [docs/04-mapping-route-va-du-lieu.md](/home/vtst/dashv4/docs/04-mapping-route-va-du-lieu.md)
- Muốn triển khai lọc ngày chuẩn cho dashboard: [docs/09-nguyen-tac-loc-ngay-report-history.md](/home/vtst/dashv4/docs/09-nguyen-tac-loc-ngay-report-history.md)
- Muốn hiểu DB `report_history.db`: [docs/03-hop-dong-du-lieu-api-transition.md](/home/vtst/dashv4/docs/03-hop-dong-du-lieu-api-transition.md)
- Muốn xem kiến trúc tổng thể và hướng refactor: [docs/05-kien-truc-de-xuat.md](/home/vtst/dashv4/docs/05-kien-truc-de-xuat.md)
- Muốn xem backlog triển khai tiếp: [docs/06-ke-hoach-trien-khai.md](/home/vtst/dashv4/docs/06-ke-hoach-trien-khai.md)

Nguyên tắc cho `dashv4`:
- Dùng `dashv3` làm shell khởi đầu, không viết lại toàn bộ ngay từ đầu.
- Không đọc Excel trực tiếp trong request path nếu dữ liệu đã có trong DB mới.
- Ưu tiên đọc `view` consumer thay vì bảng thô trong SQLite.
- Giữ tương thích route-level với `dashv3` ở nơi dữ liệu đã sẵn sàng.
- Những màn chưa có contract dữ liệu trong DB mới phải đánh dấu rõ là chưa triển khai, không dựng tạm bằng dữ liệu chắp vá.

Quy tắc đồng bộ tài liệu:
- Mọi thay đổi có ảnh hưởng tới nguồn dữ liệu, mapping route, hoặc nguyên tắc lọc ngày phải cập nhật doc trong cùng nhịp triển khai.
- Không để code đi trước tài liệu quá một phiên làm việc.
