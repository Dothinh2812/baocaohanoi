# YÊU CẦU CÔNG NHẬN SÁNG KIẾN CẤP [Điền cấp công nhận]

**Kính gửi:** [Điền Hội đồng/X đơn vị xét công nhận sáng kiến]

## Bảng tác giả

| STT | Họ và tên | Chức vụ/Đơn vị công tác | Tỷ lệ đóng góp |
|---|---|---|---|
| 1 | [Điền họ tên tác giả 1] | [Điền chức vụ/đơn vị] | [Điền %] |
| 2 | [Điền họ tên tác giả 2 nếu có] | [Điền chức vụ/đơn vị] | [Điền %] |

## Thông tin liên hệ

- Điện thoại: [Điền số điện thoại]
- Email: [Điền email]

## Tên sáng kiến

**Giải pháp tự động hóa thu thập, chuẩn hóa, tổng hợp, tính điểm KPI và lập báo cáo kỹ thuật từ hệ thống baocao.hanoi.vnpt.vn phục vụ điều hành tại TTVT Sơn Tây**

## Thời điểm bắt đầu áp dụng

[Điền thời điểm áp dụng]

## Địa điểm áp dụng

TTVT Sơn Tây thuộc VNPT Hà Nội [nếu phạm vi đăng ký khác, cập nhật lại cho phù hợp].

## Mô tả sáng kiến

Sáng kiến là một giải pháp phần mềm viết bằng Python nhằm tự động hóa gần như toàn bộ chu trình nghiệp vụ báo cáo kỹ thuật đang thực hiện trên hệ thống `baocao.hanoi.vnpt.vn`, từ khâu đăng nhập có OTP, tải nhiều nhóm báo cáo chuyên đề, chuẩn hóa dữ liệu nhân viên kỹ thuật và đơn vị, xử lý loại trừ các phiếu không tính KPI, tính điểm BSC/KPI, tổng hợp đa nguồn theo từng NVKT, lưu trữ lịch sử vào cơ sở dữ liệu SQLite và phát sinh các báo cáo Excel/Word phục vụ công tác chỉ đạo điều hành.

Giải pháp được hình thành trực tiếp từ nhu cầu thực tế thể hiện trong codebase hiện có: phải khai thác đồng thời nhiều loại báo cáo C1.1, C1.2, C1.3, C1.4, C1.5, các báo cáo thực tăng Fiber/MyTV, giao hoàn toàn trình, tiếp thị, suy hao cao, xác minh tạm dừng; đồng thời phải hợp nhất dữ liệu từ nhiều tệp Excel với tên cột, định dạng tỷ lệ và cách ghi tên NVKT không đồng nhất. Hệ thống đã tự động hóa các bước này thành một quy trình thống nhất, có log thực thi, có lưu trữ lịch sử ngày/tháng và có đầu ra phục vụ trực tiếp cho quản trị điều hành.

## 1. Mục tiêu

Mục tiêu của sáng kiến là xây dựng một công cụ phần mềm giúp giảm mạnh khối lượng thao tác thủ công trong công tác báo cáo kỹ thuật, chuẩn hóa dữ liệu đầu vào phục vụ đánh giá KPI/NVKT, tăng tính kịp thời của số liệu điều hành, tạo nguồn dữ liệu lịch sử để theo dõi xu hướng theo ngày và theo tháng, đồng thời hạn chế sai lệch do nhập liệu hoặc tổng hợp thủ công trên nhiều file Excel khác nhau.

Giải pháp hướng tới bốn kết quả chính: tự động thu thập dữ liệu từ hệ thống báo cáo nguồn; tự động làm sạch và đồng bộ khóa dữ liệu theo NVKT, đơn vị, kỳ báo cáo; tự động tính toán chỉ tiêu và tạo biểu mẫu tổng hợp; tự động lưu vết, tra cứu, so sánh trước/sau giảm trừ cũng như so sánh xu hướng qua nhiều kỳ.

## 2. Lý do đề xuất sáng kiến

Trước khi có giải pháp, quy trình tổng hợp báo cáo kỹ thuật phát sinh nhiều điểm nghẽn. Cán bộ phụ trách phải đăng nhập thủ công vào hệ thống báo cáo, chờ OTP, truy cập từng biểu mẫu, chọn đơn vị, chọn tháng hoặc khoảng ngày, xuất từng file Excel, đổi tên tệp, xử lý từng sheet rồi tiếp tục dùng Excel để ghép nối dữ liệu theo NVKT. Trong thực tế, chỉ riêng việc chuẩn hóa tên NVKT đã có thể phát sinh sai lệch do dữ liệu đầu vào tồn tại nhiều kiểu ghi khác nhau như tên có kèm địa bàn, tiền tố mã nhân viên hoặc phần chú thích trong ngoặc.

Ngoài ra, bài toán nghiệp vụ không dừng ở việc tải báo cáo gốc. Sau khi lấy dữ liệu còn phải xử lý các trường hợp giảm trừ phiếu báo hỏng, tính lại chỉ tiêu C1.1 và C1.2, đối chiếu dữ liệu trước và sau giảm trừ, tổng hợp nhiều nhóm KPI vào một biểu mẫu chung cho NVKT, đồng thời lưu lại lịch sử để theo dõi xu hướng tháng. Nếu tiếp tục làm theo phương thức rời rạc, phụ thuộc vào cá nhân và file Excel thủ công thì rất khó bảo đảm tính nhất quán, tính lặp lại và khả năng kiểm tra vết xử lý.

Từ yêu cầu thực tiễn đó, việc xây dựng một giải pháp phần mềm thống nhất, bám đúng các biểu mẫu đang dùng tại `baocao.hanoi.vnpt.vn`, là cần thiết và có giá trị áp dụng trực tiếp trong môi trường vận hành.

## 3. Nội dung sáng kiến

Nội dung cốt lõi của sáng kiến là thiết lập một chuỗi xử lý tự động, trong đó mỗi nghiệp vụ được đóng gói thành các module chuyên biệt nhưng liên kết thành một quy trình khép kín.

Thứ nhất, hệ thống tự động đăng nhập vào cổng báo cáo bằng tài khoản cấu hình trong môi trường và đọc OTP từ tệp cấu hình ngoài. Cách làm này giúp giảm thao tác nhập tay, đồng thời cho phép kiểm soát thời hạn hiệu lực của OTP và xóa mã sau khi dùng để tránh tái sử dụng sai.

Thứ hai, hệ thống tự động truy cập và tải các báo cáo nghiệp vụ đang được sử dụng thực tế, gồm nhóm báo cáo chỉ tiêu C1.x, báo cáo chi tiết SM1/SM2/SM4 phục vụ KPI, báo cáo thực tăng Fiber/MyTV, báo cáo giao hoàn toàn trình, kết quả tiếp thị, báo cáo suy hao cao, báo cáo xác minh tạm dừng và các báo cáo chuyên đề khác. Quá trình tải không chỉ dừng ở việc gọi URL mà còn tự động chọn đơn vị, kỳ báo cáo, khoảng ngày, nhóm dịch vụ và lưu tệp đầu ra theo tên chuẩn trong các thư mục nghiệp vụ.

Thứ ba, sau khi tải về, hệ thống tự động xử lý dữ liệu Excel bằng các quy tắc nghiệp vụ đã mã hóa sẵn. Những quy tắc này bao gồm tách tên NVKT từ các trường như `Nhóm địa bàn`, `TEN_KV`, `Nhân viên KT`; loại bỏ phần mã hoặc chú thích không cần thiết; đối chiếu với danh sách nhân sự tham chiếu; bổ sung đơn vị công tác; tính tổng theo tổ và theo cá nhân; hợp nhất dữ liệu từ nhiều nguồn thành một mô hình thống nhất.

Thứ tư, hệ thống tự động xử lý bài toán giảm trừ. Từ danh sách `BAOHONG_ID` loại trừ, chương trình đọc dữ liệu các chỉ tiêu liên quan, tính lại kết quả trước và sau giảm trừ, xuất các báo cáo so sánh chi tiết, đồng thời tạo bộ dữ liệu đầu vào cho khâu tính KPI. Đây là điểm có ý nghĩa thực tiễn lớn vì trước đây công việc này dễ sai nếu thao tác thủ công trên nhiều tệp.

Thứ năm, hệ thống tự động tính điểm KPI/BSC cho NVKT theo các công thức đã được mã hóa rõ ràng cho C1.1, C1.2, C1.4, C1.5; sau đó tiếp tục hợp nhất thêm các nguồn như tiếp thị, giao hoàn toàn trình, thực tăng Fiber/MyTV, suy hao cao để tạo một file tổng hợp KPI đa nguồn cho từng nhân viên kỹ thuật.

Thứ sáu, dữ liệu sau xử lý không chỉ dùng cho báo cáo tại thời điểm hiện tại mà còn được import vào SQLite để hình thành kho lịch sử ngày và tháng. Trên cơ sở đó, hệ thống tạo được các báo cáo xu hướng nhiều kỳ và biểu đồ phục vụ theo dõi biến động chỉ tiêu theo đơn vị và theo NVKT.

Thứ bảy, hệ thống còn tự động tạo các đầu ra phục vụ điều hành như báo cáo Word KPI, báo cáo tổng hợp Excel, log thực thi từng bước và một số cảnh báo tác nghiệp như cảnh báo xác minh tạm dừng khi có trường hợp cần theo dõi.

## 4. Kiến trúc phần mềm và nguyên lý hoạt động

Về mặt kiến trúc, giải pháp được tổ chức theo mô hình nhiều lớp, trong đó mỗi lớp đảm nhận một vai trò rõ ràng nhưng dữ liệu được truyền nối tiếp thành chu trình khép kín.

Lớp truy cập và thu thập dữ liệu được xây dựng trên `Playwright`, kết hợp cấu hình môi trường qua `.env`. Lớp này thực hiện đăng nhập vào hệ thống báo cáo, lấy OTP từ tệp cấu hình ngoài, mở các màn hình báo cáo, chọn tiêu chí lọc và tải file Excel. Các module tiêu biểu gồm `login.py`, `c1_report_download.py`, `thuc_tang_download.py`, `KR_download.py`, `kq_tiep_thi_download.py`, `xac_minh_tam_dung_download.py`. Đây là lớp thay thế hoàn toàn thao tác thủ công trên giao diện web.

Lớp xử lý nghiệp vụ được xây dựng chủ yếu trên `pandas` và `openpyxl`. Lớp này đảm nhiệm làm sạch dữ liệu, chuẩn hóa tên NVKT, bổ sung cột đơn vị, tính tổng hợp theo tổ và theo cá nhân, tạo báo cáo thực tăng, tổng hợp giao việc, xử lý danh sách xác minh tạm dừng, tính toán giảm trừ và chuẩn hóa cấu trúc file đầu ra. Các module tiêu biểu gồm `c1_process.py`, `thuc_tang_process.py`, `KR_process.py`, `kq_tiep_thi_process.py`, `exclusion_process.py`, `xac_minh_tam_dung_download.py`.

Lớp tính toán KPI và tổng hợp điều hành được thực hiện tại `kpi_calculator.py`, `kpi_tonghop_nvkt.py`, `report_generator.py`, `xuat_baocao_xuhung.py`, `make_chart_pttb.py`. Lớp này đọc dữ liệu đã chuẩn hóa, áp dụng công thức BSC/KPI cho từng chỉ tiêu, hợp nhất chín nhóm dữ liệu KPI về một biểu mẫu chung, tạo báo cáo Word và biểu đồ, từ đó phục vụ trực tiếp công tác đánh giá và điều hành.

Lớp lưu trữ và lịch sử số liệu sử dụng `SQLite` với các bảng cho đơn vị, NVKT và từng nhóm báo cáo ngày/tháng. Các module `import_baocao.py`, `import_baocao_thang.py` và phần schema hiện có cho thấy hệ thống đã thiết kế lưu trữ theo khóa duy nhất của kỳ báo cáo và đối tượng, cho phép import lặp lại theo cơ chế cập nhật, tạo lịch sử và truy vấn xu hướng nhiều kỳ.

Nguyên lý hoạt động của toàn hệ thống có thể mô tả theo chuỗi bước sau:

1. Người vận hành cấu hình tài khoản, mật khẩu, đường dẫn OTP, kỳ báo cáo và khoảng ngày.
2. Chương trình tự động đăng nhập vào `baocao.hanoi.vnpt.vn`, nhận OTP và xác lập phiên làm việc.
3. Hệ thống lần lượt truy cập các biểu mẫu nguồn, chọn đơn vị TTVT Sơn Tây hoặc các phạm vi liên quan, sau đó tải các file Excel về thư mục chuẩn.
4. Các module xử lý đọc file tải về, chuẩn hóa trường dữ liệu, tách tên NVKT, ghép với danh sách nhân sự và tạo các sheet tổng hợp theo tổ/NVKT.
5. Nếu bật chế độ giảm trừ, hệ thống đọc danh sách phiếu loại trừ, tính lại số liệu C1.1 và C1.2, xuất báo cáo so sánh trước/sau giảm trừ.
6. Hệ thống tính điểm KPI/BSC cho từng nhân viên kỹ thuật, đồng thời tổng hợp thêm các nguồn thực tăng, tiếp thị, giao hoàn toàn trình và suy hao cao vào một file KPI tổng hợp.
7. Dữ liệu đã xử lý được ghi vào cơ sở dữ liệu SQLite để lưu lịch sử ngày/tháng và phục vụ xuất báo cáo xu hướng.
8. Hệ thống phát sinh đầu ra cuối cùng dưới dạng Excel, Word, biểu đồ và log kiểm tra vết thực thi.

Điểm đáng chú ý là kiến trúc này không chỉ tự động hóa một báo cáo đơn lẻ mà hình thành một dây chuyền dữ liệu hoàn chỉnh, từ nguồn phát sinh đến sản phẩm quản trị cuối cùng.

## 5. Giá trị mới của sáng kiến

Giá trị mới thứ nhất là chuyển đổi từ cách làm rời rạc, phụ thuộc vào thao tác thủ công trên từng báo cáo sang một quy trình phần mềm hóa end-to-end có thể lặp lại. Trong codebase hiện tại, việc tải dữ liệu, xử lý chuẩn hóa, tính KPI, tạo báo cáo và lưu lịch sử đã được liên kết thành cùng một workflow tại `baocaohanoi.py`.

Giá trị mới thứ hai là chuẩn hóa khóa dữ liệu NVKT giữa nhiều nguồn có cấu trúc không đồng nhất. Hệ thống không chỉ đọc file mà còn làm sạch tên, loại bỏ tiền tố mã, bỏ nội dung trong ngoặc, quy đổi về định dạng thống nhất rồi mới thực hiện tổng hợp. Đây là phần có ý nghĩa quyết định đối với độ tin cậy của báo cáo tổng hợp.

Giá trị mới thứ ba là đưa nghiệp vụ giảm trừ vào phần mềm, thay vì để người dùng tính thủ công ngoài hệ thống. Các báo cáo so sánh trước/sau giảm trừ được sinh tự động, nhờ đó việc kiểm tra tác động của từng phiếu loại trừ tới các chỉ tiêu KPI trở nên minh bạch và có thể kiểm tra lại.

Giá trị mới thứ tư là hình thành kho dữ liệu lịch sử ngày và tháng trên SQLite, tạo nền tảng để theo dõi xu hướng và phục vụ điều hành theo thời gian. Đây là bước nâng cấp từ mô hình “xử lý file một lần” sang mô hình “dữ liệu có tích lũy và có thể khai thác lại”.

Giá trị mới thứ năm là tích hợp đầu ra đa dạng, không chỉ có Excel nghiệp vụ mà còn có Word, biểu đồ và log thực thi. Điều này giúp hệ thống vừa phục vụ xử lý nội bộ, vừa phục vụ báo cáo quản trị, vừa hỗ trợ kiểm soát quá trình vận hành.

## 6. Đánh giá lợi ích

Về lợi ích định tính, sáng kiến giúp rút ngắn đáng kể thời gian lập báo cáo định kỳ do thay thế nhiều công đoạn thao tác tay bằng quy trình tự động. Khối lượng công việc thủ công giảm ở tất cả các khâu: đăng nhập và tải báo cáo, đổi tên và phân loại tệp, chuẩn hóa tên nhân sự, ghép dữ liệu nhiều nguồn, tính KPI, lập báo cáo tổng hợp và lưu lịch sử.

Về độ chính xác, giải pháp giảm rủi ro sai lệch do sao chép công thức Excel, bỏ sót dòng dữ liệu, ghép sai NVKT hoặc quên cập nhật dữ liệu sau giảm trừ. Toàn bộ logic tính toán được mã hóa thành các hàm xử lý và có thể chạy lặp lại theo cùng một chuẩn.

Về khả năng kiểm tra vết, hệ thống có log thực thi từng bước và có đầu ra dữ liệu trung gian, giúp việc rà soát, đối chiếu và truy nguyên nguyên nhân khi có chênh lệch thuận lợi hơn so với cách làm thủ công.

Về quản trị dữ liệu, sáng kiến tạo ra nền tảng lưu trữ lịch sử, cho phép theo dõi xu hướng theo ngày, theo tháng, theo đơn vị và theo NVKT; từ đó hỗ trợ lãnh đạo và bộ phận chuyên môn đưa ra quyết định điều hành nhanh hơn, có căn cứ hơn.

Về lợi ích định lượng, đề nghị bổ sung số liệu thực tế sau quá trình áp dụng như: số giờ công tiết kiệm mỗi kỳ báo cáo, số lượng báo cáo/tệp được tự động hóa, tỷ lệ giảm lỗi tổng hợp, tỷ lệ rút ngắn thời gian hoàn thành báo cáo, hiệu quả kinh tế quy đổi và giá trị làm lợi. Các số liệu này hiện chưa thể suy ra an toàn từ codebase nên cần điền theo số liệu triển khai thực tế: `[Điền số liệu lợi ích định lượng]`.

## 7. Khả năng áp dụng

Sáng kiến có khả năng áp dụng ngay tại đơn vị đang vận hành nghiệp vụ tương tự vì giải pháp đã bám sát cấu trúc báo cáo hiện hành trên `baocao.hanoi.vnpt.vn`, sử dụng các thư viện phổ biến, triển khai trên môi trường Python thông dụng và lưu dữ liệu bằng SQLite gọn nhẹ.

Ngoài phạm vi TTVT Sơn Tây, giải pháp có thể mở rộng cho các đơn vị khác trong VNPT Hà Nội khi điều chỉnh cấu hình đơn vị, danh mục nhân sự và tham số truy cập báo cáo. Thiết kế module hóa của hệ thống cũng cho phép thêm mới các báo cáo hoặc thay đổi công thức tính mà không phải viết lại toàn bộ chương trình.

Đối với những nghiệp vụ cần kiểm tra định kỳ, cần đối chiếu trước/sau giảm trừ hoặc cần tổng hợp nhiều nguồn về cấp NVKT, giải pháp có thể nhân rộng với chi phí triển khai thấp hơn nhiều so với phát triển một hệ thống mới từ đầu.

## 8. Hướng mở rộng

Trong giai đoạn tiếp theo, có thể mở rộng sáng kiến theo các hướng sau: bổ sung giao diện cấu hình thân thiện để người dùng không cần sửa trực tiếp file nguồn; bổ sung cơ chế lập lịch chạy tự động theo ngày/tháng; mở rộng sang lưu trữ tập trung trên cơ sở dữ liệu dùng chung; kết nối cảnh báo qua Zalo/Telegram hoặc các kênh điều hành khác; bổ sung dashboard trực quan để khai thác dữ liệu lịch sử theo thời gian thực.

Ngoài ra, khi hệ thống báo cáo nguồn thay đổi cấu trúc giao diện hoặc phát sinh thêm chỉ tiêu mới, bộ giải pháp hiện tại vẫn có thể thích nghi tương đối nhanh do đã tách riêng lớp tải báo cáo, lớp xử lý dữ liệu và lớp báo cáo đầu ra.

## Cam đoan

Tôi/chúng tôi cam đoan nội dung trình bày trong hồ sơ này là kết quả nghiên cứu, xây dựng và áp dụng từ thực tiễn công việc tại đơn vị; các thông tin còn để trong ngoặc vuông là các nội dung cần hoàn thiện thêm theo hồ sơ nhân sự, quyết định áp dụng và số liệu thực tế trước khi trình xét công nhận.

## Ký xác nhận

- Đại diện tác giả/nhóm tác giả: [Điền họ tên và ký xác nhận]
- Xác nhận của đơn vị: [Điền nội dung xác nhận]
