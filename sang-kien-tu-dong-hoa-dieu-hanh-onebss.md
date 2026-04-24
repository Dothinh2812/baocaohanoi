# YÊU CẦU CÔNG NHẬN SÁNG KIẾN CẤP VNPT HÀ NỘI NĂM [Điền năm đề nghị công nhận]

Kính gửi: Hội đồng sáng kiến Trung tâm Viễn thông Sơn Tây

Chúng tôi ghi tên dưới đây:

| TT | Họ tên tác giả | Nam/Nữ | Trình độ chuyên môn | Chức vụ, Đơn vị công tác | Chủ trì SK | Tỉ lệ đóng góp | Ký tên |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | [Họ tên tác giả 1] | [ ] | [Trình độ] | [Chức vụ, đơn vị] | X | [  ] % |  |
| 2 | [Họ tên tác giả 2] | [ ] | [Trình độ] | [Chức vụ, đơn vị] |  | [  ] % |  |
| 3 | [Họ tên tác giả 3] | [ ] | [Trình độ] | [Chức vụ, đơn vị] |  | [  ] % |  |

Điện thoại: [Điền số điện thoại]

Email: [Điền email]

Địa chỉ bưu điện: [Điền địa chỉ]

Căn cứ Quy định hiện hành của Tập đoàn về hoạt động sáng kiến, sáng tạo;

Yêu cầu xét công nhận sáng kiến cấp cơ sở:

**“Giải pháp tự động hóa điều hành kỹ thuật trên nền OneBSS phục vụ giám sát báo hỏng, phiếu tồn phát triển thuê bao và cảnh báo tác nghiệp tại Trung tâm Viễn thông Sơn Tây”**

Sáng kiến này đã được bắt đầu áp dụng từ: [Điền thời điểm áp dụng]

Địa điểm áp dụng: Trung tâm Viễn thông Sơn Tây - VNPT Hà Nội

## Mô tả sáng kiến

### 1. Mục tiêu

Mục tiêu của sáng kiến là xây dựng một chương trình điều phối tự động, chạy theo chu kỳ, nhằm thay thế phần lớn các thao tác thủ công đang phát sinh trong quá trình khai thác hệ thống OneBSS phục vụ công tác kỹ thuật. Chương trình phải tự đăng nhập, tự tải báo cáo, tự xử lý dữ liệu, tự đo kiểm, tự phát hiện bất thường, tự tạo biểu đồ và tự gửi cảnh báo đến đúng nhóm tác nghiệp, qua đó giúp giảm độ trễ điều hành và nâng cao khả năng phản ứng của các tổ kỹ thuật địa bàn.

Bên cạnh mục tiêu tự động hóa nghiệp vụ báo hỏng BRCD, sáng kiến còn hướng tới xử lý đồng thời các nhóm nghiệp vụ liên quan khác đang xuất hiện trong thực tế vận hành như tổng hợp phiếu tồn phát triển thuê bao, theo dõi vòng đời phiếu, cảnh báo khách hàng doanh nghiệp ưu tiên, cảnh báo phiếu quá hạn hoặc sắp quá hạn, và phát hiện sớm nguy cơ sự cố diện rộng trên cùng hướng mạng. Việc hợp nhất các khâu này vào một tiến trình chung giúp hình thành đầu mối điều hành kỹ thuật tập trung, thống nhất và có khả năng mở rộng.

### 2. Lý do đề xuất sáng kiến

Trong thực tế sản xuất, nhân viên kỹ thuật và tổ trưởng thường phải thao tác lặp lại nhiều bước trên OneBSS và các file Excel trung gian: đăng nhập hệ thống, chờ OTP, tải nhiều loại báo cáo, gộp file, lọc dữ liệu, đo kiểm theo file, phân loại phiếu theo đội, thống kê số lượng tồn, lập danh sách phiếu bất thường, sau đó mới tiếp tục gửi thông tin sang Zalo, Telegram hoặc các đầu mối CSKH. Khối lượng việc này diễn ra nhiều lần trong ngày, dễ phát sinh chậm trễ, sai sót và phụ thuộc lớn vào kinh nghiệm từng cá nhân.

Đối với báo hỏng BRCD, việc nhận biết phiếu mới, tình trạng cổng ON/OFF, thuê bao nào cần ưu tiên xử lý hay khả năng nhiều thuê bao cùng mất liên lạc trên một SA thường chưa được cảnh báo sớm nếu chỉ dựa vào cách theo dõi thủ công. Đối với PTTB, các chỉ tiêu quá giờ, sắp quá giờ hoặc phiếu chưa có lý do tồn cũng cần được bóc tách nhanh để tổ kỹ thuật kịp thời xử lý. Nếu mỗi phần việc được làm rời rạc bằng tay thì thông tin điều hành bị phân tán, không tạo được một chu trình liên tục từ dữ liệu gốc đến hành động xử lý.

Từ yêu cầu thực tiễn đó, sáng kiến được đề xuất nhằm tự động hóa toàn bộ chu trình điều hành kỹ thuật trên nền OneBSS, tận dụng dữ liệu có sẵn của hệ thống, kết hợp xử lý dữ liệu bằng Python, lưu vết bằng SQLite và phân phối cảnh báo qua webhook, Zalo, Telegram. Điểm cốt lõi là biến các báo cáo vốn chỉ mang tính tra cứu thành một luồng tác nghiệp tự động, có thể chủ động cảnh báo và hỗ trợ ra quyết định gần thời gian thực.

### 3. Nội dung sáng kiến

Sáng kiến hiện thực hóa một ứng dụng nền viết bằng Python, trong đó `app.py` giữ vai trò tiến trình điều phối trung tâm. Ngay khi khởi động, chương trình tự mở trình duyệt Playwright, đăng nhập OneBSS bằng tài khoản cấu hình sẵn, tiếp nhận mã OTP từ file hoặc cho phép nhập bổ sung khi cần. Sau khi vào hệ thống, ứng dụng duy trì cơ chế tự kiểm tra phiên làm việc, nếu hết hạn sẽ tự đăng nhập lại hoặc khởi tạo lại toàn bộ browser để bảo đảm các đợt chạy sau không bị đứt quãng.

Trên mỗi chu kỳ xử lý, chương trình thực hiện tải song song nhiều nhóm báo cáo của OneBSS như BRCD, Metronet, Megawan, TSL, IMS, cố định, MyTV và báo cáo PTTB. Sau khi tải xong, hệ thống tự gộp báo cáo, tạo file đo kiểm đầu vào, chạy chức năng đo theo file, đọc kết quả trả về và bóc tách các thông số quan trọng như trạng thái cổng ONU, suy hao Rx/Tx, thời gian tồn thực, giờ còn lại thực, đội phụ trách và dữ liệu SA. Từ đây, chương trình tự tạo các file tổng hợp chuẩn dùng chung cho các bước điều hành tiếp theo.

Đối với BRCD và DHSC-005, hệ thống chuẩn hóa dữ liệu báo hỏng, loại bỏ các bản ghi không hợp lệ, kết hợp kết quả đo kiểm và bổ sung thông tin thuê bao từ cơ sở dữ liệu danh bạ. Kết quả đầu ra không chỉ dừng ở file tổng hợp mà còn được chia theo đội, theo địa bàn, theo nhân viên, phục vụ cả thống kê quản trị lẫn tác nghiệp hiện trường. Ngoài ra, chương trình còn tự gửi các bản ghi báo hỏng mới đến webhook CSKH theo quy tắc lọc cấu hình sẵn để hỗ trợ phối hợp chăm sóc khách hàng.

Đối với PTTB, hệ thống tự đọc báo cáo phiếu tồn, tính lại chỉ tiêu thời gian, phân loại trạng thái bình thường, sắp quá giờ, quá giờ, thống kê theo đội, theo khu vực, theo loại dịch vụ và tạo file tách riêng cho từng đội. Nhờ đó, công tác theo dõi tiến độ phát triển thuê bao không còn phụ thuộc vào thao tác thủ công trên từng file nguồn, đồng thời tạo được cơ sở dữ liệu đầu vào cho biểu đồ và cảnh báo điều hành.

Trên nền dữ liệu đã được chuẩn hóa, chương trình tiếp tục sinh biểu đồ phục vụ điều hành cho cả BRCD và PTTB, đồng thời gọi các mô-đun cảnh báo tự động. Các cảnh báo đang được kích hoạt trực tiếp trong `app.py` gồm cảnh báo khách hàng doanh nghiệp ưu tiên, cảnh báo phiếu OB KHL, cảnh báo phiếu HLL kéo dài 7 ngày, cảnh báo phiếu tồn sắp quá giờ, cùng cơ chế theo dõi snapshot để nhận biết phiếu mới nhận, phiếu đã xử lý xong và cập nhật thống kê ngày. Ở cuối chu kỳ, hệ thống còn thực hiện một bước kiểm tra chuyên sâu để phát hiện nguy cơ mất hướng SA bằng cách lấy toàn bộ thuê bao trên cùng SA ra đo lại hàng loạt; nếu số lượng thuê bao OFF vượt ngưỡng cấu hình thì gửi cảnh báo ngay cho đội phụ trách và ghi log đầy đủ phục vụ tra cứu.

### 4. Kiến trúc phần mềm và nguyên lý hoạt động

Giải pháp được xây dựng theo mô hình ứng dụng điều phối nền, không phụ thuộc giao diện người dùng riêng, mà khai thác trực tiếp giao diện OneBSS bằng tự động hóa trình duyệt. Kiến trúc phần mềm có thể chia thành năm lớp chính: lớp truy cập hệ thống nguồn, lớp điều phối tác vụ, lớp xử lý và làm giàu dữ liệu, lớp lưu trữ và theo dõi lịch sử, và lớp cảnh báo - báo cáo đầu ra.

Lớp truy cập hệ thống nguồn được thực hiện bởi Playwright kết hợp các tham số cấu hình trong `config.py`. Thành phần này đảm nhiệm các công việc đăng nhập, xử lý OTP, theo dõi tình trạng phiên làm việc, mở thêm các trang con trong cùng browser context và tải song song nhiều báo cáo khác nhau từ OneBSS. Việc giới hạn số tác vụ đồng thời bằng semaphore giúp tăng tốc độ tải trong khi vẫn kiểm soát được mức độ ổn định của phiên đăng nhập.

Lớp điều phối tác vụ được hiện thực tập trung trong `app.py`. Tệp này đóng vai trò bộ lập lịch chu kỳ, tổ chức thứ tự các bước xử lý và cơ chế phục hồi lỗi. Mỗi chu kỳ bắt đầu từ việc xác nhận trạng thái đăng nhập, tiếp đó là tải báo cáo, gộp dữ liệu, chạy đo kiểm, xử lý BRCD và PTTB, đồng bộ dữ liệu vào cơ sở dữ liệu, cập nhật lịch sử phiếu, sinh biểu đồ và phát cảnh báo. Nếu gặp lỗi đăng nhập kéo dài, ứng dụng tự đóng browser, chờ và khởi tạo lại toàn bộ tiến trình để bảo đảm hệ thống có khả năng tự phục hồi mà không cần can thiệp thủ công thường xuyên.

Lớp xử lý và làm giàu dữ liệu sử dụng các mô-đun như `downloadBaoCao.py`, `downloadBaocaoPTTB.py`, `xly_dhsc.py`, `xly_dhsc_005.py`, `do_theo_file.py`, `xly_pttb.py` và `enrich_subscriber_data.py`. Dữ liệu đầu vào từ OneBSS được chuẩn hóa thành các file Excel làm việc; sau đó hệ thống trích xuất tình trạng cổng, suy hao, giờ tồn, nhân viên thực hiện, đội kỹ thuật, SA và các thuộc tính thuê bao. Đối với báo hỏng, chương trình tạo các sheet tổng hợp theo đội và theo địa bàn, đồng thời sinh file đo kiểm đầu vào để gọi lại chức năng đo theo file của OneBSS. Đối với PTTB, chương trình tính toán lại trạng thái theo ngưỡng thời gian và tổ chức dữ liệu theo đội, khu vực và loại hình dịch vụ.

Lớp lưu trữ và theo dõi lịch sử sử dụng kết hợp Excel, SQLite, JSON và CSV. Cơ sở dữ liệu `database/brcd.db` lưu snapshot từng lần tải báo cáo, lịch sử vòng đời phiếu và thống kê ngày; `danhba.db` và các bảng tra cứu thuê bao phục vụ đối soát SA, địa bàn và thông tin khách hàng; các file log CSV và text giúp ghi nhận bản ghi đã gửi webhook, lịch sử cảnh báo và nhật ký sự cố. Cách thiết kế này giúp hệ thống vừa đáp ứng nhu cầu vận hành nhanh bằng file quen thuộc, vừa có khả năng lưu vết lịch sử để kiểm chứng và đánh giá hiệu quả sau này.

Lớp cảnh báo và báo cáo đầu ra được thực hiện qua `make_chart.py`, `make_chart_pttb.py`, `send_zalo_via_n8n_webhook.py`, `send_new_brcd_to_webhook.py`, `ticket_tracking.py` và các mô-đun gửi Telegram. Hệ thống tạo biểu đồ điều hành cho BRCD và PTTB, gửi bản tin cảnh báo theo từng nhóm nghiệp vụ, định tuyến đúng đội kỹ thuật dựa trên `team_config.py`, và truyền dữ liệu ra các webhook liên thông phục vụ chỉ đạo nội bộ hoặc phối hợp CSKH. Một số hàm gửi ảnh Telegram hiện đã có sẵn trong mã nguồn và có thể bật lại khi đơn vị cần mở rộng kênh phân phối trực quan.

Nguyên lý hoạt động của toàn bộ giải pháp được mô tả như sau. Trước hết, chương trình tự động đăng nhập OneBSS và xác nhận phiên làm việc còn hiệu lực. Tiếp đó, hệ thống tải đồng thời các báo cáo kỹ thuật cần thiết, gộp và chuẩn hóa dữ liệu thành các file xử lý trung gian. Sau bước này, ứng dụng tạo danh sách thuê bao cần đo, gọi lại chức năng đo theo file của OneBSS, phân tích kết quả đo và kết hợp với dữ liệu danh bạ để làm giàu thông tin cho từng phiếu. Từ tập dữ liệu đã hoàn chỉnh, hệ thống sinh các bảng tổng hợp, biểu đồ và các danh sách ưu tiên xử lý. Cùng lúc, các mô-đun cảnh báo tự động rà soát những điều kiện bất thường đã cấu hình sẵn như khách hàng ưu tiên, phiếu kéo dài, phiếu sắp quá hạn hoặc dấu hiệu mất hướng SA để gửi đến đúng đầu mối phụ trách. Cuối cùng, trạng thái phiếu ở thời điểm hiện tại được chụp lại vào cơ sở dữ liệu để so sánh với các lần chạy sau, qua đó hình thành một chu trình khép kín từ thu thập dữ liệu, nhận diện bất thường, truyền thông cảnh báo đến lưu vết lịch sử và đánh giá xử lý.

### 5. Giá trị mới của sáng kiến

Giá trị mới thứ nhất của sáng kiến là chuyển từ mô hình làm việc rời rạc trên nhiều công cụ sang một tiến trình điều hành kỹ thuật tự động, có khả năng nối liền các bước từ tải báo cáo, xử lý dữ liệu, đo kiểm, phát hiện bất thường đến gửi cảnh báo. Trong phạm vi đơn vị, đây không chỉ là tự động hóa một thao tác đơn lẻ mà là tự động hóa chu trình điều hành của nhiều nghiệp vụ kỹ thuật đang chạy song song trên OneBSS.

Giá trị mới thứ hai là sáng kiến tận dụng trực tiếp dữ liệu và chức năng sẵn có của OneBSS nhưng nâng lên thành các quy tắc tác nghiệp chủ động. Ví dụ, chức năng đo theo file không còn là công cụ kiểm tra thủ công mà trở thành mắt xích để hệ thống tự xác minh tình trạng cổng, suy hao và phát hiện sự cố diện rộng theo SA. Tương tự, các báo cáo PTTB và BRCD không chỉ để xem số liệu mà được chuyển thành nguồn đầu vào cho cảnh báo ưu tiên, thống kê ngày và định tuyến xử lý theo đội.

Giá trị mới thứ ba là giải pháp có cơ chế lưu vết và chống lặp tương đối đầy đủ, bao gồm chụp snapshot phiếu, ghi lịch sử gửi webhook, log cảnh báo sự cố SA và cơ chế tái sử dụng cấu hình đội nhóm trong một đầu mối chung. Nhờ đó, hệ thống có tính thực dụng cao, có thể vận hành liên tục trong môi trường sản xuất và dễ mở rộng thêm các quy tắc cảnh báo khác mà không phải thiết kế lại từ đầu.

### 6. Đánh giá lợi ích

Việc áp dụng sáng kiến giúp giảm đáng kể thời gian và thao tác thủ công trong khâu điều hành kỹ thuật. Thay vì phải đăng nhập, tải báo cáo, gộp file, đo theo file, tổng hợp số liệu và nhắn tin cảnh báo bằng tay, cán bộ kỹ thuật có thể nhận ngay dữ liệu đã được chuẩn hóa và các cảnh báo ưu tiên sau mỗi chu kỳ chạy. Điều này giúp tổ trưởng tập trung hơn vào quyết định điều hành và xử lý thực địa, thay vì mất thời gian cho các thao tác lặp lại mang tính hành chính kỹ thuật.

Sáng kiến cũng giúp tăng tốc độ phát hiện và phản ứng trước các trường hợp bất thường như phiếu kéo dài, phiếu sắp quá hạn, khách hàng ưu tiên hoặc dấu hiệu nhiều thuê bao cùng OFF trên một SA. Khi thông tin được đưa đúng lúc, đúng đội phụ trách và có lưu vết lịch sử, việc phối hợp giữa kỹ thuật, điều hành và CSKH trở nên chủ động hơn, giảm nguy cơ bỏ sót thông tin hoặc xử lý chậm.

Ngoài lợi ích tức thời trong vận hành, giải pháp còn tạo nền dữ liệu có cấu trúc cho công tác thống kê, đánh giá chất lượng xử lý phiếu và mở rộng phân tích sau này. Dữ liệu snapshot, lịch sử phiếu và log cảnh báo là cơ sở để đơn vị xây dựng các chỉ tiêu quản trị theo ngày, theo đội hoặc theo nhân viên, thay vì chỉ nhìn vào số liệu thời điểm.

Phần lợi ích định lượng: [Điền số liệu về số giờ công giảm được mỗi ngày hoặc mỗi tháng, số lượt thao tác thủ công được cắt giảm, thời gian trung bình phát hiện bất thường trước và sau áp dụng, số vụ mất hướng SA được phát hiện sớm, tỷ lệ giảm chậm xử lý phiếu, giá trị làm lợi quy đổi hoặc các chỉ tiêu định lượng khác nếu có].

### 7. Khả năng áp dụng

Giải pháp có khả năng áp dụng trực tiếp tại Trung tâm Viễn thông Sơn Tây do đã bám sát cấu trúc đội kỹ thuật, quy trình khai thác OneBSS, dữ liệu báo hỏng BRCD, dữ liệu PTTB và các kênh cảnh báo nội bộ đang sử dụng thực tế. Vì được xây dựng trên các thành phần phổ biến như Python, Playwright, pandas, SQLite, Excel và webhook, hệ thống không đòi hỏi đầu tư hạ tầng phức tạp, chủ yếu cần chuẩn hóa cấu hình, tài khoản truy cập và dữ liệu tham chiếu.

Sáng kiến cũng có thể nhân rộng cho các trung tâm viễn thông khác của VNPT Hà Nội hoặc các đơn vị có mô hình nghiệp vụ tương tự. Khi triển khai sang đơn vị mới, chỉ cần cập nhật các tham số cấu hình như tài khoản OneBSS, đường dẫn OTP, danh sách đội kỹ thuật, webhook nhận cảnh báo, dữ liệu danh bạ thuê bao và ngưỡng cảnh báo. Kiến trúc điều phối theo mô-đun cho phép bổ sung hoặc tắt bớt từng nghiệp vụ mà không phải thay đổi toàn bộ chương trình.

### 8. Hướng mở rộng

Trong giai đoạn tiếp theo, sáng kiến có thể mở rộng theo các hướng sau: bổ sung dashboard tập trung để hiển thị trạng thái các chu kỳ chạy, số lượng phiếu bất thường và lịch sử cảnh báo theo thời gian thực; liên thông sâu hơn với các hệ thống CSKH hoặc KPI để tự động hóa tiếp phần đánh giá kết quả xử lý; mở rộng các quy tắc phát hiện sự cố từ cấp SA sang các mức chi tiết hơn như PON, OLT hoặc khu vực mạng ngoại vi; và tăng cường tự động gửi biểu đồ, ảnh báo cáo cho từng nhóm kỹ thuật theo lịch.

Một hướng mở rộng quan trọng khác là nâng cấp lớp dữ liệu lịch sử để phục vụ phân tích xu hướng, đánh giá hiệu suất đội kỹ thuật và gợi ý ưu tiên xử lý bằng dữ liệu nhiều ngày. Khi đó, chương trình không chỉ dừng ở vai trò trợ lý tự động hóa tác nghiệp mà có thể trở thành nền tảng điều hành kỹ thuật số cho đơn vị.

Chúng tôi cam đoan những nội dung trên đây là đúng sự thật.

| Xác nhận của đơn vị | Sơn Tây, ngày ..... tháng ..... năm [Điền năm ký] |
| --- | --- |
| [Ký, ghi rõ họ tên, đóng dấu] | Tác giả sáng kiến  |
|  | [Ký và ghi rõ họ tên] |
