# Hiện trạng dashv3

## Kiến trúc hiện tại

`dashv3` là Flask monolith:
- bootstrap app ở `dashboard.py`
- tách nghiệp vụ qua blueprint
- giao diện theo kiểu `base.html` + page template + page script
- nhiều route API đọc Excel, ảnh, SQLite, hoặc dữ liệu từ repo khác

Các nhóm blueprint chính:
- `operations`
- `quality`
- `growth`
- `retention`
- `inventory`
- `quangchudong`
- `sa_outage`
- `statistics`
- `auth`

## Cách render giao diện

Luồng chung của một màn `dashv3`:
1. route HTML render page template
2. page JS gọi `/api/...`
3. API đọc file hoặc DB
4. JS dựng bảng, chart, badge, summary

Điểm cần kế thừa cho `dashv4`:
- cấu trúc route và trải nghiệm truy cập
- cách chia trang theo nghiệp vụ
- shape JSON mà frontend đang kỳ vọng nếu muốn tái sử dụng JS cũ

## Inventory route/page quan trọng

### Operations

| Route | Vai trò hiện tại |
| --- | --- |
| `/brcd` | báo cáo BRCD |
| `/pttb` | báo cáo PTTB |
| `/cau-hinh-tu-dong` | cấu hình tự động Sơn Tây |
| `/tong-hop-bsc-kpi` | dashboard tổng hợp BSC KPI theo đơn vị hiện tại |
| `/thuctang` | trang ảnh/thông tin thực tăng |
| `/kpi` | KPI tổng hợp |
| `/kpi-nvkt-bchn` | KPI NVKT BCHN |

API liên quan nổi bật:
- `/api/excel-data`
- `/api/excel-data-main`
- `/api/excel-data-pending`
- `/api/pttb-data-summary`
- `/api/pttb-data-detail`
- `/api/pttb-data-pending`
- `/api/pttb-data-chitiet-to`
- `/api/cau-hinh-tu-dong/son-tay`
- `/api/tong-hop-bsc-kpi`
- `/api/kpi-data`
- `/api/kpi-nvkt-bchn-data`

### Quality

| Route | Vai trò hiện tại |
| --- | --- |
| `/chatluong` | dashboard C1 |
| `/i15` | I1.5 |
| `/i15k2` | I1.5 K2 |
| `/shc-processing` | tổng hợp xử lý SHC |

API liên quan nổi bật:
- `/api/c1-chat-luong-data`
- `/api/c11-chi-tiet-data`
- `/api/c11-chi-tiet-15h-data`
- `/api/c11-chi-tiet-16h-data`
- `/api/c11-chi-tiet-17h-data`
- `/api/c11-chi-tiet-18h-data`
- `/api/c11-chi-tiet-quality-data`
- `/api/c12-repeat-failure-data`
- `/api/c14-chi-tiet-data`
- `/api/c15-chi-tiet-data`
- `/api/i15-data`
- `/api/i15k2-data`

### Growth và Retention

| Route | Vai trò hiện tại |
| --- | --- |
| `/tiepthi` | kết quả tiếp thị |
| `/thuc-tang-ngung-psc` | tăng trưởng/ngừng PSC |
| `/giahan` | GHTT |
| `/thuhoi` | vật tư thu hồi |

API liên quan nổi bật:
- `/api/tiepthi-data`
- `/api/ngungpsc-data`
- `/api/ngungpsc-mytv-data`
- `/api/ngungpsc-fiber-to-data`
- `/api/ngungpsc-mytv-to-data`
- `/api/ngungpsc-fiber-nvkt-data`
- `/api/ngungpsc-mytv-nvkt-data`
- `/api/giahan-ghtt-hni`
- `/api/giahan-ghtt-sty`
- `/api/giahan-ghtt-nvktdb`
- `/api/thu-hoi-data`

### Inventory

| Route | Vai trò hiện tại |
| --- | --- |
| `/ton-kho-vat-tu` | tồn kho tổng |
| `/Tong_hop_tien` | tổng hợp tiền vật tư |
| `/tra-cuu-nhanh-vat-tu` | tra cứu nhanh |
| `/xac-minh-tam-dung` | xác minh tạm dừng |
| `/tam-dung-khoi-phuc` | tạm dừng/khôi phục |

API liên quan nổi bật:
- `/api/ton-kho-vat-tu`
- `/api/ton-kho-vat-tu-tot-thuong-dung`
- `/api/tra-cuu-nhanh-vat-tu`
- `/api/xac-minh-tam-dung-data`
- `/api/tam-dung-khoi-phuc-data`

### Nhóm khác

| Route | Vai trò hiện tại |
| --- | --- |
| `/quangchudong` | quang chủ động |
| `/su_co_sa` | sự cố SA |
| `/api/ticket-statistics` | thống kê ticket |

## Phụ thuộc dữ liệu hiện tại

`dashv3` không phải app tự chứa dữ liệu. Nó đang đọc từ nhiều nơi:
- Excel nội bộ trong repo hoặc repo anh em
- DB lịch sử ticket riêng
- DB/cached service của quang chủ động
- SQLite hoặc file từ các quy trình khác
- ảnh PNG/chart render sẵn

Hệ quả:
- route giống nhau nhưng nguồn dữ liệu rất không đồng nhất
- một số màn khó tái sử dụng nguyên xi nếu chỉ đổi sang SQLite mới

## Hệ quả đối với dashv4

`dashv4` không nên sao chép toàn bộ `dashv3` một cách cơ học.

Nên tách 3 nhóm:
- nhóm giữ route + giữ UI + chỉ đổi backend query
- nhóm giữ route nhưng phải chỉnh shape dữ liệu một ít
- nhóm chưa triển khai vì DB mới chưa đủ nguồn
