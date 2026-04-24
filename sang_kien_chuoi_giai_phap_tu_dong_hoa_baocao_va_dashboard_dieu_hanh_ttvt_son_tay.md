# YÊU CẦU CÔNG NHẬN SÁNG KIẾN CẤP CƠ SỞ

**Kính gửi:** Hội đồng sáng kiến Trung tâm Viễn thông Sơn Tây

## Bảng tác giả

| TT | Họ tên tác giả | Nam/Nữ | Trình độ chuyên môn | Chức vụ, đơn vị công tác | Chủ trì sáng kiến | Tỉ lệ đóng góp | Ký tên |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | [Điền họ tên tác giả 1] | [Điền] | [Điền] | Trung tâm Viễn thông Sơn Tây | X | [Điền %] | |
| 2 | [Điền họ tên tác giả 2 nếu có] | [Điền] | [Điền] | Trung tâm Viễn thông Sơn Tây |  | [Điền %] | |

**Điện thoại:** [Điền số điện thoại]  
**Email:** [Điền email]  
**Địa chỉ bưu điện:** Trung tâm Viễn thông Sơn Tây, VNPT Hà Nội

Căn cứ quy định hiện hành của Tập đoàn về hoạt động sáng kiến, sáng tạo, chúng tôi đề nghị xét công nhận sáng kiến cấp cơ sở với nội dung sau:

## Tên sáng kiến

**Xây dựng chuỗi giải pháp tự động hóa thu thập, xử lý báo cáo và dashboard điều hành giám sát chỉ tiêu chất lượng phục vụ công tác quản trị tại Trung tâm Viễn thông Sơn Tây**

## Thời điểm bắt đầu áp dụng

[Điền thời điểm áp dụng]

## Địa điểm áp dụng

Trung tâm Viễn thông Sơn Tây, VNPT Hà Nội

## Mô tả sáng kiến

### 1. Mục tiêu

Mục tiêu của sáng kiến là xây dựng một chuỗi giải pháp phần mềm liên thông từ khâu thu thập báo cáo nguồn, xử lý và chuẩn hóa dữ liệu, đến khâu hiển thị dashboard điều hành nhằm phục vụ theo dõi, giám sát và quản trị các chỉ tiêu chất lượng tại Trung tâm Viễn thông Sơn Tây. Thay vì chỉ cải tiến một công đoạn riêng lẻ, sáng kiến hướng tới số hóa trọn vẹn luồng dữ liệu từ lúc báo cáo phát sinh trên các hệ thống nguồn đến khi trở thành thông tin điều hành trực quan cho lãnh đạo và cán bộ nghiệp vụ.

Chuỗi giải pháp này hướng tới bốn kết quả chính: giảm mạnh thao tác thủ công khi khai thác báo cáo từ `baocao.hanoi.vnpt.vn` và các quy trình điều hành kỹ thuật dựa trên OneBSS; chuẩn hóa dữ liệu đầu vào theo NVKT, tổ địa bàn, chỉ tiêu và kỳ báo cáo; hình thành dashboard tập trung để giám sát tức thời các chỉ tiêu chất lượng, phiếu tồn, tiến độ xử lý và cảnh báo tác nghiệp; và tạo nền tảng dữ liệu có thể kiểm tra vết, đối chiếu và khai thác lại cho công tác điều hành hằng ngày, hằng tuần, hằng tháng.

### 2. Lý do đề xuất sáng kiến

Trong thực tế, bài toán quản trị chỉ tiêu chất lượng tại đơn vị không chỉ nằm ở việc “xem dashboard”, mà còn bắt đầu từ khâu lấy dữ liệu đầu vào. Trước đây, cán bộ phụ trách phải đăng nhập thủ công vào hệ thống báo cáo, nhập hoặc chờ OTP, truy cập từng biểu mẫu, chọn đơn vị, chọn thời gian, tải từng file Excel, đổi tên tệp, xử lý dữ liệu, làm sạch tên NVKT, tổng hợp lại theo tổ và theo cá nhân, rồi mới có thể dùng tiếp cho báo cáo điều hành. Với các nghiệp vụ kỹ thuật trên OneBSS, cán bộ còn phải lặp lại các bước tải báo cáo BRCD/PTTB, đo kiểm theo file, phân loại phiếu, theo dõi phiếu mới, phiếu quá hạn, phiếu sắp quá hạn hoặc nguy cơ bất thường trên cùng hướng mạng. Quy trình này tốn nhiều thời gian, dễ sai lệch, phụ thuộc kinh nghiệm cá nhân và khó duy trì tính nhất quán giữa các kỳ báo cáo.

Sau khi có dữ liệu gốc, đơn vị vẫn phải tiếp tục xử lý nhiều lớp nghiệp vụ như tính lại số liệu sau giảm trừ, đối chiếu dữ liệu trước và sau xử lý, tổng hợp các nhóm C1.x, I1.5, SHC, KPI, thực tăng, tiếp thị và các báo cáo liên quan. Với nhóm điều hành kỹ thuật, còn phát sinh nhu cầu snapshot vòng đời phiếu, phát hiện phiếu mới nhận, phiếu xử lý xong, cảnh báo khách hàng ưu tiên, cảnh báo phiếu kéo dài và bóc tách nhanh các trường hợp cần can thiệp theo đội kỹ thuật. Nếu thiếu một chuỗi giải pháp liên thông, dữ liệu đầu vào và dữ liệu dùng để điều hành rất dễ bị đứt đoạn, khiến dashboard nếu có cũng chỉ là nơi hiển thị lại các file đã được làm thủ công, chưa tạo ra bước chuyển thực sự về năng lực quản trị.

Xuất phát từ yêu cầu đó, cần có một sáng kiến theo hướng gộp mạnh hơn: tự động hóa báo cáo nguồn ở lớp đầu, chuẩn hóa và kết xuất dữ liệu ở lớp giữa, sau đó dùng dashboard nội bộ ở lớp cuối để giám sát, phân tích và điều hành trên cùng một chuỗi dữ liệu thống nhất.

### 3. Nội dung sáng kiến

Sáng kiến là việc hình thành một chuỗi giải pháp phần mềm gồm hai lớp liên thông chặt chẽ.

Lớp thứ nhất là hệ thống tự động hóa khai thác báo cáo và điều hành nguồn. Theo tài liệu và mã nguồn đang được đơn vị sử dụng, lớp này được xây dựng bằng Python để thực hiện đăng nhập có OTP, truy cập các biểu mẫu nghiệp vụ, tải các báo cáo chuyên đề như C1.1, C1.2, C1.3, C1.4, C1.5, SM1, SM2, SM4, thực tăng Fiber/MyTV, giao hoàn toàn trình, tiếp thị, suy hao cao, xác minh tạm dừng và các dữ liệu phục vụ KPI từ `baocao.hanoi.vnpt.vn`. Song song với đó, ở nhánh điều hành kỹ thuật dựa trên OneBSS, hệ thống tự động xử lý các báo cáo BRCD, PTTB và các dữ liệu tác nghiệp liên quan, thực hiện chuẩn hóa phiếu, theo dõi trạng thái xử lý, chụp snapshot vòng đời phiếu, tạo biểu đồ và sinh các cảnh báo điều hành cần thiết. Sau khi thu thập, các quy trình tiếp tục chuẩn hóa tên NVKT, đơn vị, kỳ báo cáo, xử lý loại trừ hoặc bóc tách bất thường và kết xuất thành các file Excel, bản ghi SQLite và dữ liệu trung gian phục vụ dashboard.

Lớp thứ hai là hệ thống dashboard nội bộ trên nền Flask để khai thác các đầu ra đã được chuẩn hóa. Dashboard hiện cung cấp các phân hệ C1 - Chất lượng phục vụ khách hàng, I1.5, I1.5 K2, kết quả xử lý SHC, BRCD, PTTB, KPI NVKT, quang chủ động, vật tư, tiếp thị và các phân hệ quản trị liên quan. Người dùng đăng nhập một lần, chọn đúng phân hệ và bộ lọc cần theo dõi là có thể xem tổng hợp, chi tiết theo tổ, chi tiết theo NVKT, theo ngày/tháng, trạng thái cảnh báo, tình hình phiếu nhận và phiếu xử lý, đồng thời tải lại file nguồn khi cần kiểm tra.

Điểm cốt lõi của sáng kiến nằm ở chỗ hai lớp này không vận hành rời rạc. Lớp tự động hóa báo cáo tạo ra nguồn dữ liệu đầu vào ổn định và chuẩn hóa; lớp dashboard tiếp nhận chính đầu ra đó để biến dữ liệu thành thông tin điều hành trực quan, có đối chiếu, có drill-down và có hỗ trợ quản trị. Nhờ vậy, sáng kiến không chỉ giải quyết bài toán “làm báo cáo” và cũng không chỉ giải quyết bài toán “xem báo cáo”, mà tạo ra một chuỗi điều hành số khép kín từ nguồn đến quyết định.

### 4. Kiến trúc phần mềm và nguyên lý hoạt động

Kiến trúc phần mềm của sáng kiến được tổ chức theo mô hình nhiều lớp liên thông.

Lớp thu thập và tạo nguồn dữ liệu đầu vào được xây dựng trên Python, theo các tài liệu sáng kiến nguồn có sử dụng `Playwright`, `pandas`, `openpyxl` và SQLite. Lớp này chịu trách nhiệm đăng nhập vào `baocao.hanoi.vnpt.vn`, lấy OTP từ cấu hình ngoài, truy cập các biểu mẫu nguồn, chọn điều kiện lọc, tải file Excel, sau đó làm sạch dữ liệu, chuẩn hóa tên NVKT và đơn vị, xử lý giảm trừ, tính toán một số chỉ tiêu trung gian, lưu lịch sử và kết xuất các file đầu ra phục vụ điều hành. Ở nhánh điều hành kỹ thuật trên OneBSS, lớp này còn thực hiện tải báo cáo BRCD/PTTB theo chu kỳ, gộp và bóc tách dữ liệu phiếu, theo dõi trạng thái xử lý, sinh biểu đồ tác nghiệp và chuẩn bị dữ liệu cho các cảnh báo theo đội hoặc theo loại bất thường. Đây là lớp thay thế phần lớn thao tác thủ công trước đây trên giao diện web và Excel.

Lớp ứng dụng điều hành được hiện thực bởi `dashboard.py` cùng các `blueprint` như `auth`, `operations`, `quality`, `statistics`, `growth`, `inventory`, `retention`, `quangchudong`. Hệ thống dùng Flask để tổ chức route, Jinja2 để dựng giao diện, session filesystem để quản lý đăng nhập và token CSRF để bảo vệ biểu mẫu. Giao diện dùng `templates/base.html` làm khung chung cho menu điều hướng và các trang nghiệp vụ; phía frontend dùng JavaScript để gọi API, nhận JSON và hiển thị bảng, tab, KPI, trạng thái bộ lọc và cảnh báo.

Ở lớp dữ liệu khai thác, dashboard đọc trực tiếp các file Excel và SQLite đã được tạo hoặc cập nhật từ các lớp nguồn. Các báo cáo C1.x, I1.5, các file chi tiết SHC, một phần KPI và các báo cáo chất lượng khác được lấy từ thư mục tải báo cáo Hà Nội. Dữ liệu điều hành BRCD/PTTB và lịch sử vòng đời phiếu được lưu tại `brcd.db`, cho phép khai thác thêm góc nhìn về phiếu mới nhận, phiếu xử lý xong, phiếu tồn đầu ngày và cuối ngày. Dữ liệu xử lý SHC được ghép từ `nvkt_results.db`, `suy_hao_history.db` và `suy_hao_history_k2.db`. Để tăng hiệu quả truy xuất, hệ thống sử dụng cache khi đọc Excel thông qua `read_excel_sheet_cached`.

Ở lớp xử lý và tổng hợp, dashboard dùng `pandas` để đọc sheet, nhóm dữ liệu theo tổ hoặc NVKT, sắp xếp theo tỷ lệ đạt hoặc mức cảnh báo và tạo payload JSON cho giao diện. Module `ticket_tracking.py` thực hiện chụp snapshot, so sánh giữa các thời điểm để xác định phiếu mới nhận, phiếu xử lý xong và sinh thống kê theo ngày, qua đó đưa logic theo dõi vòng đời phiếu vào dashboard điều hành. Dịch vụ `services/shc_processing_report.py` ghép dữ liệu “đã xử lý” và “cần xử lý”, tính tổng ngày, tổng tháng, tổng theo tổ và tổng theo cá nhân, đồng thời gắn cờ cảnh báo cho các trường hợp có nguy cơ chậm tiến độ. Tư duy cảnh báo tác nghiệp thể hiện trong sáng kiến điều hành OneBSS cũng được phản ánh ở đây thông qua các bộ lọc, trạng thái cảnh báo, so sánh khối lượng cần xử lý và khối lượng đã xử lý.

Nguyên lý hoạt động của toàn chuỗi giải pháp diễn ra như sau. Thứ nhất, các hệ thống tự động hóa đăng nhập vào `baocao.hanoi.vnpt.vn` và các quy trình điều hành kỹ thuật dựa trên OneBSS để tải báo cáo nguồn, chuẩn hóa dữ liệu, theo dõi vòng đời phiếu và kết xuất đầu ra theo cấu trúc thống nhất. Thứ hai, các file Excel và bản ghi SQLite đó trở thành nguồn đầu vào cho dashboard nội bộ. Thứ ba, khi người dùng truy cập một trang chất lượng hoặc điều hành, frontend gửi yêu cầu tới API tương ứng. Thứ tư, backend đọc dữ liệu đầu vào đã được chuẩn hóa, xử lý nhóm, tổng hợp, so sánh và trả dữ liệu JSON. Thứ năm, giao diện hiển thị dữ liệu dưới dạng bảng tổng hợp, bảng chi tiết theo tab, thẻ KPI, cảnh báo tiến độ, thống kê vòng đời phiếu và các lựa chọn tải file nguồn. Với nguyên lý này, chuỗi xử lý dữ liệu được khép kín từ nguồn phát sinh đến lớp điều hành cuối cùng.

### 5. Giá trị mới của sáng kiến

Giá trị mới thứ nhất là chuyển đổi từ cải tiến cục bộ sang cải tiến toàn chuỗi. Thay vì chỉ tự động tải báo cáo hoặc chỉ dựng dashboard hiển thị, sáng kiến kết nối cả hai thành một dây chuyền dữ liệu thống nhất, làm tăng giá trị sử dụng thực tế của từng thành phần.

Giá trị mới thứ hai là hình thành lớp chuẩn hóa dữ liệu đầu vào trước khi điều hành. Đây là điểm rất quan trọng vì dữ liệu từ nhiều biểu mẫu nguồn thường khác nhau về tên cột, cách ghi NVKT, đơn vị, tỷ lệ phần trăm và cấu trúc báo cáo. Khi đã có lớp chuẩn hóa này, dashboard mới có thể hoạt động ổn định, cho ra cùng một cách nhìn ở mọi kỳ báo cáo.

Giá trị mới thứ ba là dashboard không chỉ hiển thị dữ liệu hiện trạng mà còn hỗ trợ đối chiếu, drill-down và giám sát thực thi. Điều này thể hiện rõ ở các phân hệ C1, I1.5, kết quả xử lý SHC, BRCD và PTTB, nơi dữ liệu được theo dõi theo tổ, theo NVKT, theo ngày/tháng, theo vòng đời phiếu và có cảnh báo đối với các điểm cần đôn đốc.

Giá trị mới thứ tư là sáng kiến tạo được khả năng kiểm tra vết trên toàn chuỗi. Từ lớp tải báo cáo nguồn, xử lý, giảm trừ, chụp snapshot phiếu, lưu lịch sử cho đến lớp dashboard khai thác, dữ liệu đều có thể truy nguyên lại theo file, theo kỳ báo cáo và theo nguồn, giúp giảm rủi ro sai lệch và tăng tính minh bạch trong điều hành.

Giá trị mới thứ năm là giải pháp được xây dựng trên nền các công nghệ phổ biến, gọn nhẹ, dễ vận hành và dễ mở rộng theo nhu cầu thực tế của đơn vị. Điều này làm cho sáng kiến có tính khả thi cao, chi phí triển khai thấp và thuận lợi trong nhân rộng.

### 6. Đánh giá lợi ích

Về lợi ích quản trị, sáng kiến giúp lãnh đạo và bộ phận tổng hợp chuyển từ mô hình nhận file rời rạc sang mô hình điều hành trên một chuỗi dữ liệu thống nhất. Nhờ đó, việc theo dõi chỉ tiêu, phát hiện điểm nghẽn, giao việc, kiểm tra tiến độ và nhận diện sớm các trường hợp bất thường trong xử lý phiếu được thực hiện nhanh hơn, nhất quán hơn và ít phụ thuộc vào thao tác trung gian của cá nhân.

Về lợi ích nghiệp vụ, sáng kiến giảm đáng kể khối lượng công việc thủ công ở cả hai đầu của quy trình. Ở đầu vào, hệ thống giảm các công đoạn đăng nhập, nhập OTP, chọn biểu mẫu, tải file, đổi tên, làm sạch dữ liệu, ghép nối Excel và theo dõi thủ công các trạng thái phiếu kỹ thuật. Ở đầu ra, dashboard giảm thời gian tổng hợp, đối chiếu, tìm kiếm dữ liệu chi tiết, chuẩn bị báo cáo điều hành và hỗ trợ bóc tách nhanh các trường hợp cần ưu tiên xử lý theo tổ hoặc theo cá nhân.

Về lợi ích chất lượng dữ liệu, sáng kiến nâng cao độ tin cậy của số liệu do dữ liệu đầu vào được chuẩn hóa trước khi đưa lên dashboard, đồng thời có thể kiểm tra lại theo file nguồn, theo kỳ báo cáo và theo từng lớp xử lý. Điều này đặc biệt quan trọng trong môi trường có nhiều chỉ tiêu liên quan chặt chẽ đến KPI, đánh giá thực hiện và công tác quản trị.

Về lợi ích định lượng, đề nghị bổ sung số liệu thực tế sau quá trình áp dụng như: số giờ công tiết kiệm trong mỗi chu kỳ báo cáo, số lượng báo cáo được tự động hóa, tỷ lệ giảm thao tác thủ công, thời gian rút ngắn từ khi tải báo cáo đến khi có dashboard khai thác, mức giảm sai lệch tổng hợp và giá trị làm lợi quy đổi.  
**[Điền số liệu lợi ích định lượng]**

### 7. Khả năng áp dụng

Giải pháp có khả năng áp dụng ngay tại Trung tâm Viễn thông Sơn Tây vì được hình thành trực tiếp từ nhu cầu, dữ liệu và mô hình tổ chức thực tế của đơn vị. Cả hai lớp của sáng kiến đều đang bám sát các biểu mẫu nguồn và các nhóm chỉ tiêu đang dùng trong thực tế, nên mức độ phù hợp vận hành cao.

Ngoài phạm vi đơn vị, sáng kiến cũng có thể mở rộng cho các trung tâm viễn thông khác thuộc VNPT Hà Nội khi điều chỉnh cấu hình truy cập báo cáo nguồn, danh mục nhân sự, tham số đường dẫn dữ liệu và các quy tắc tổ chức theo địa bàn. Do thiết kế theo module, hệ thống có thể bổ sung thêm báo cáo mới hoặc chỉ tiêu mới mà không phải viết lại toàn bộ.

### 8. Hướng mở rộng

Trong giai đoạn tiếp theo, sáng kiến có thể mở rộng theo hướng đồng bộ sâu hơn giữa lớp tự động hóa và lớp dashboard, ví dụ như lập lịch chạy tự động theo ngày/tháng, cập nhật dữ liệu gần thời gian thực hơn, bổ sung cảnh báo chủ động qua các kênh điều hành như webhook, Zalo hoặc Telegram, mở rộng kho dữ liệu lịch sử tập trung và tăng cường các biểu đồ xu hướng phục vụ quản trị.

Ngoài ra, khi có thêm các bài toán điều hành mới tại đơn vị, chuỗi giải pháp này có thể được mở rộng thành một nền tảng bảng điều hành số tổng hợp, trong đó lớp báo cáo nguồn tiếp tục đóng vai trò tạo dữ liệu đầu vào, còn dashboard giữ vai trò trung tâm khai thác và ra quyết định.

## Cam đoan

Chúng tôi cam đoan những nội dung nêu trên là đúng sự thật; các mô tả kỹ thuật trong hồ sơ được xây dựng dựa trên chức năng hiện có của các phần mềm đang sử dụng tại đơn vị, các thông tin chưa đủ căn cứ xác định được giữ ở dạng chỗ trống để hoàn thiện sau.

## Ký xác nhận

**Xác nhận của đơn vị**  
[Ký, ghi rõ họ tên, đóng dấu]

**Sơn Tây, ngày ..... tháng ..... năm [Điền năm nộp hồ sơ]**  
**Tác giả sáng kiến**  
[Ký và ghi rõ họ tên]
