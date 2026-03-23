# -*- coding: utf-8 -*-
"""
CTS Login Module - Đăng nhập tự động vào hệ thống CTS VNPT
Sử dụng thông tin đăng nhập từ file .env
"""

import os
import time
import re
from playwright.sync_api import sync_playwright
from config import Config


def read_otp_from_file():
    """
    Đọc mã OTP từ file (đường dẫn cấu hình trong .env)
    Chỉ chấp nhận file được tạo trong vòng OTP_MAX_AGE_SECONDS gần đây
    Xóa mã OTP sau khi đọc thành công

    Returns:
        str: Mã OTP 6 chữ số hoặc None nếu không tìm thấy
    """
    file_path = Config.OTP_FILE_PATH
    max_age_seconds = Config.OTP_MAX_AGE_SECONDS
    max_retries = 10
    retry_count = 0
    otp_code = None

    print(f"Đang đọc mã OTP từ file: {file_path}")
    print(f"OTP max age: {max_age_seconds} seconds")

    while retry_count < max_retries:
        if os.path.exists(file_path):
            file_time = os.path.getmtime(file_path)
            file_time_formatted = time.strftime('%H:%M:%S', time.localtime(file_time))
            print(f"File time: {file_time_formatted}")
            current_time = time.time()
            current_time_formatted = time.strftime('%H:%M:%S', time.localtime(current_time))
            print(f"Current time: {current_time_formatted}")
            time_diff = current_time - file_time
            print(f"Time difference: {time_diff:.2f} seconds")

            if time_diff <= max_age_seconds:  # File is recent enough
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    # Find 6-digit number using regex
                    otp_match = re.search(r'\b\d{6}\b', content)
                    if otp_match:
                        otp_code = otp_match.group(0)
                        print(f"✅ Found OTP code in file: {otp_code}")

                        # Xóa mã OTP bằng cách ghi đè file trống
                        try:
                            with open(file_path, 'w', encoding='utf-8') as f:
                                f.write('')
                            print("✅ Đã xóa mã OTP khỏi file")
                        except Exception as e:
                            print(f"⚠️ Không thể xóa mã OTP: {e}")

                        return otp_code
            else:
                print(f"⚠️ File quá cũ (hơn {max_age_seconds} giây), chờ file mới...")
        else:
            print(f"⚠️ Không tìm thấy file OTP tại {file_path}")

        retry_count += 1
        if retry_count < max_retries:
            print(f"Waiting for OTP... Attempt {retry_count}/{max_retries}")
            time.sleep(2)  # Wait 2 seconds before next attempt

    print("❌ Không tìm thấy OTP hợp lệ sau nhiều lần thử")
    return None


def login_cts():
    """
    Đăng nhập vào hệ thống CTS VNPT (https://cts.vnpt.vn)

    Returns:
        tuple: (page_cts, browser_cts, playwright_cts) - Trả về đối tượng page đã đăng nhập
               Trả về (None, None, None) nếu đăng nhập thất bại
    """
    print("=" * 60)
    print("🔐 Bắt đầu đăng nhập vào CTS (https://cts.vnpt.vn)")
    print("=" * 60)

    # Khởi tạo Playwright và Browser cho CTS
    playwright_cts = sync_playwright().start()
    browser_cts = playwright_cts.chromium.launch(headless=Config.BROWSER_HEADLESS)
    context_cts = browser_cts.new_context(accept_downloads=Config.ACCEPT_DOWNLOADS)
    page_cts = context_cts.new_page()

    try:
        # Bước 1: Truy cập trang đăng nhập CTS
        print("Đang truy cập trang đăng nhập CTS...")
        page_cts.goto('https://cts.vnpt.vn', timeout=Config.PAGE_LOAD_TIMEOUT)
        page_cts.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)

        # Bước 2: Điền username
        print(f"Đang điền username: {Config.BAOCAO_USERNAME}")
        username_field = page_cts.locator('//*[@id="username"]')
        username_field.wait_for(state="visible", timeout=30000)
        username_field.fill(Config.BAOCAO_USERNAME)
        time.sleep(1)

        # Bước 3: Điền password
        print("Đang điền password...")
        password_field = page_cts.locator('//*[@id="password"]')
        password_field.wait_for(state="visible", timeout=30000)
        password_field.fill(Config.BAOCAO_PASSWORD)
        time.sleep(1)

        # Bước 4: Click button Đăng nhập
        print("Đang click button Đăng nhập...")
        login_button = page_cts.locator('//*[@id="fm1"]/section/button')
        login_button.wait_for(state="visible", timeout=30000)
        login_button.click()
        time.sleep(3)

        # Bước 5: Đợi trường input OTP xuất hiện
        print("Đang đợi trường nhập OTP...")
        otp_field = page_cts.locator('//*[@id="passOTP"]')
        otp_field.wait_for(state="visible", timeout=30000)

        # Bước 6: Đọc OTP từ file
        otp_code = read_otp_from_file()

        if otp_code is None:
            print("❌ Không thể đọc OTP từ file.")
            print("⏸️  Vui lòng nhập OTP thủ công vào trường trên trang web và click xác nhận.")
            print("⏸️  Script sẽ chờ 60 giây để bạn hoàn tất đăng nhập...")
            time.sleep(60)
            print("✅ Tiếp tục sau khi chờ...")
        else:
            # Bước 7: Điền OTP
            print(f"Đang điền OTP: {otp_code}")
            otp_field.fill(otp_code)
            time.sleep(1)

            # Bước 8: Click button xác nhận OTP
            print("Đang click button xác nhận OTP...")
            otp_confirm_button = page_cts.locator('//*[@id="loginForm"]/div[1]/button')
            otp_confirm_button.wait_for(state="visible", timeout=30000)
            otp_confirm_button.click()
            time.sleep(5)

        # Bước 9: Kiểm tra đăng nhập thành công (URL không còn chứa "cas/login")
        page_cts.wait_for_url(lambda url: "cas/login" not in url, timeout=60000)
        page_cts.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)
        print("✅ Đăng nhập CTS thành công!")

        # Kiểm tra cookies và session
        cookies = context_cts.cookies()
        print(f"📝 Đã lưu {len(cookies)} cookies")

        return page_cts, browser_cts, playwright_cts

    except Exception as e:
        print(f"❌ Lỗi đăng nhập CTS: {str(e)}")
        # Đóng browser nếu lỗi
        browser_cts.close()
        playwright_cts.stop()
        return None, None, None


def login_cts_and_goto_report(report_url="https://cts.vnpt.vn/Linetest/Report/GponQualityByUnitvb8362"):
    """
    Đăng nhập CTS và truy cập vào trang báo cáo GPON Quality.

    Args:
        report_url: URL báo cáo cần truy cập. Mặc định là GPON Quality By Unit.

    Returns:
        tuple: (page_cts, browser_cts, playwright_cts) nếu thành công
               (None, None, None) nếu thất bại
    """
    # Bước 1: Đăng nhập CTS
    page_cts, browser_cts, playwright_cts = login_cts()

    if page_cts is None:
        print("❌ Đăng nhập CTS thất bại!")
        return None, None, None

    # Bước 2: Truy cập vào trang báo cáo
    try:
        print(f"\n📊 Đang truy cập báo cáo: {report_url}")
        page_cts.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)
        page_cts.wait_for_load_state("networkidle", timeout=Config.PAGE_LOAD_TIMEOUT)
        print("✅ Đã truy cập thành công trang báo cáo!")
        return page_cts, browser_cts, playwright_cts
    except Exception as e:
        print(f"❌ Lỗi khi truy cập báo cáo: {str(e)}")
        browser_cts.close()
        playwright_cts.stop()
        return None, None, None


def download_gpon_report(page_cts, report_date=None, output_dir=None):
    """
    Điền form và tải báo cáo GPON Quality (suy hao) từ CTS.

    Args:
        page_cts: Đối tượng Playwright page đã đăng nhập và ở trang báo cáo
        report_date: Ngày báo cáo theo format 'MM/DD/YYYY'. Mặc định là ngày hiện tại.
        output_dir: Thư mục lưu file. Mặc định là thư mục baocaohanoi hiện tại.

    Returns:
        str: Đường dẫn file đã tải về, hoặc None nếu thất bại
    """
    from datetime import datetime

    # Xác định ngày báo cáo
    if report_date is None:
        # Mặc định là ngày hiện tại
        report_date = datetime.now().strftime('%m/%d/%Y')

    # Xác định thư mục lưu file
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(__file__))

    output_filename = "suy_hao_cts.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    print("\n" + "=" * 60)
    print("📋 Bắt đầu điền form và tải báo cáo GPON Quality")
    print(f"   Ngày báo cáo: {report_date}")
    print(f"   Lưu tại: {output_path}")
    print("=" * 60)

    try:
        # Bước 1: Chọn "Trung tâm viễn thông cũ" từ dropdown loại đơn vị
        print("\n📌 Bước 1: Chọn loại đơn vị -> Trung tâm viễn thông cũ")
        unit_type_dropdown = page_cts.locator('//*[@id="app"]/form/div[1]/div/div/select')
        unit_type_dropdown.wait_for(state="visible", timeout=30000)
        unit_type_dropdown.select_option(value="0")  # value="0" = Trung tâm viễn thông cũ
        time.sleep(2)
        print("   ✅ Đã chọn 'Trung tâm viễn thông cũ'")

        # Bước 2: Click vào dropdown đơn vị và chọn "Hà Nội"
        print("\n📌 Bước 2: Chọn đơn vị -> Hà Nội")
        # Click vào input để mở dropdown (input là readonly nên không dùng fill)
        unit_input = page_cts.locator('//*[@id="unit-list"]/div/div[1]/div[1]/div[1]/input')
        unit_input.wait_for(state="visible", timeout=30000)
        unit_input.click()
        time.sleep(2)

        # Tìm và click vào option "Hà Nội" trong dropdown list
        # Thử nhiều cách để tìm option Hà Nội
        hanoi_option = None
        
        # Cách 1: Tìm trong danh sách ul/li
        try:
            hanoi_option = page_cts.locator('li:has-text("Hà Nội")').first
            if hanoi_option.is_visible():
                hanoi_option.click()
                time.sleep(2)
                print("   ✅ Đã chọn 'Hà Nội' (cách 1)")
            else:
                raise Exception("Option không visible")
        except:
            # Cách 2: Tìm span có text Hà Nội
            try:
                hanoi_option = page_cts.locator('span:has-text("Hà Nội")').first
                hanoi_option.click()
                time.sleep(2)
                print("   ✅ Đã chọn 'Hà Nội' (cách 2)")
            except:
                # Cách 3: Tìm div có text Hà Nội
                try:
                    hanoi_option = page_cts.locator('div:has-text("Hà Nội")').first
                    hanoi_option.click()
                    time.sleep(2)
                    print("   ✅ Đã chọn 'Hà Nội' (cách 3)")
                except:
                    # Cách 4: Dùng text selector generic
                    hanoi_option = page_cts.get_by_text("Hà Nội", exact=False).first
                    hanoi_option.click()
                    time.sleep(2)
                    print("   ✅ Đã chọn 'Hà Nội' (cách 4)")

        # Bước 3: Click vào trường chọn thời gian
        print("\n📌 Bước 3: Thiết lập khoảng thời gian")
        date_range_field = page_cts.locator('//*[@id="reservationtime"]')
        date_range_field.wait_for(state="visible", timeout=30000)
        date_range_field.click()
        time.sleep(2)  # Đợi date picker mở

        # Bước 4: Điền ngày bắt đầu (00:00)
        # Sử dụng xpath= prefix cho full path XPath
        start_date_input = page_cts.locator('xpath=/html/body/div[2]/div[1]/div[1]/input')
        start_date_input.wait_for(state="visible", timeout=30000)
        # Triple click để select all, sau đó gõ giá trị mới
        start_date_input.click(click_count=3)
        time.sleep(0.5)
        start_date_input.fill(f"{report_date} 00:00")
        print(f"   ✅ Đã điền ngày bắt đầu: {report_date} 00:00")

        # Bước 5: Điền ngày kết thúc (23:00)
        end_date_input = page_cts.locator('xpath=/html/body/div[2]/div[2]/div[1]/input')
        end_date_input.wait_for(state="visible", timeout=30000)
        # Triple click để select all, sau đó gõ giá trị mới
        end_date_input.click(click_count=3)
        time.sleep(0.5)
        end_date_input.fill(f"{report_date} 23:00")
        print(f"   ✅ Đã điền ngày kết thúc: {report_date} 23:00")

        # Bước 6: Click nút Apply để xác nhận thời gian
        print("\n📌 Bước 4: Apply thời gian")
        apply_button = page_cts.locator('xpath=/html/body/div[2]/div[3]/div/button[1]')
        apply_button.wait_for(state="visible", timeout=30000)
        apply_button.click()
        time.sleep(2)
        print("   ✅ Đã apply thời gian")

        # Bước 7: Click nút Tìm kiếm
        print("\n📌 Bước 5: Click nút Tìm kiếm")
        search_button = page_cts.locator('//*[@id="btnSearch"]')
        search_button.wait_for(state="visible", timeout=30000)
        search_button.click()
        print("   ✅ Đã click Tìm kiếm")

        # Bước 8: Đợi bảng dữ liệu load xong (có dữ liệu trong table)
        # Sau khi search, spinner xuất hiện, khi xong sẽ có dữ liệu trong bảng
        print("\n⏳ Đang đợi báo cáo load (đợi dữ liệu xuất hiện trong bảng)...")
        
        # Đợi 3 giây để bắt đầu loading
        time.sleep(3)
        
        # Đợi bảng dữ liệu có ít nhất 1 dòng (tối đa 5 phút)
        # Thử nhiều selector cho table row
        table_row_selectors = [
            'table tbody tr',
            '.table tbody tr',
            'table tr:not(:first-child)',
            '[class*="table"] tbody tr',
        ]
        
        max_wait_seconds = 300  # 5 phút
        check_interval = 5
        elapsed = 0
        data_found = False
        
        while elapsed < max_wait_seconds:
            for selector in table_row_selectors:
                try:
                    rows = page_cts.locator(selector)
                    count = rows.count()
                    if count > 0:
                        print(f"   ✅ Đã tìm thấy {count} dòng dữ liệu trong bảng sau {elapsed}s!")
                        data_found = True
                        break
                except:
                    continue
            
            if data_found:
                break
                
            print(f"   ⏳ Đang đợi dữ liệu... ({elapsed}s / {max_wait_seconds}s)")
            time.sleep(check_interval)
            elapsed += check_interval
        
        if not data_found:
            print(f"   ⚠️ Không tìm thấy dữ liệu sau {max_wait_seconds}s, thử download anyway...")

        # Đợi thêm 5 giây để đảm bảo dữ liệu đã hoàn toàn sẵn sàng
        time.sleep(5)

        # Bước 9: Click nút Download lần 1 (để trigger chuẩn bị file)
        print("\n📥 Bước 6: Tải báo cáo về")
        download_button = page_cts.locator('//*[@id="btnDetailKem"]')
        download_button.wait_for(state="visible", timeout=30000)
        download_button.click()
        print("   ✅ Đã click nút download (hệ thống đang chuẩn bị file)")

        # Bước 10: Đợi cố định 70 giây (hệ thống yêu cầu đợi tối thiểu 60 giây)
        print("   ⏳ Đang đợi hệ thống chuẩn bị file (tối thiểu 60 giây theo yêu cầu)...")
        wait_seconds = 70
        for i in range(wait_seconds):
            if (i + 1) % 10 == 0:
                print(f"   ⏳ Đã đợi {i + 1}s / {wait_seconds}s...")
            time.sleep(1)
        print(f"   ✅ Đã đợi {wait_seconds} giây!")

        # Bước 11: Click lại để download file
        print("   ⏳ Đang tải file...")
        with page_cts.expect_download(timeout=Config.DOWNLOAD_TIMEOUT) as download_info:
            download_button.click()

        download = download_info.value

        # Lưu file về thư mục đích với tên mong muốn
        download.save_as(output_path)
        print(f"   ✅ Đã lưu file: {output_path}")

        return output_path

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo: {str(e)}")
        return None


def download_cts_gpon_report(report_date=None, output_dir=None):
    """
    Hàm chính: Đăng nhập CTS, truy cập báo cáo, điền form và tải về.

    Args:
        report_date: Ngày báo cáo theo format 'MM/DD/YYYY'. Mặc định là ngày hiện tại.
        output_dir: Thư mục lưu file. Mặc định là thư mục baocaohanoi hiện tại.

    Returns:
        str: Đường dẫn file đã tải về, hoặc None nếu thất bại
    """
    # Bước 1: Login và truy cập trang báo cáo
    page_cts, browser_cts, playwright_cts = login_cts_and_goto_report()

    if page_cts is None:
        return None

    try:
        # Bước 2: Điền form và tải báo cáo
        result = download_gpon_report(page_cts, report_date, output_dir)
        return result

    finally:
        # Luôn đóng browser sau khi hoàn thành
        print("\n🔒 Đang đóng browser...")
        browser_cts.close()
        playwright_cts.stop()
        print("✅ Đã đóng browser")


# ============================================================================
# MAIN - Chạy standalone để test
# ============================================================================
if __name__ == "__main__":
    print("=" * 60)
    print("CTS GPON Report Downloader - Test")
    print("=" * 60)

    # Test tải báo cáo GPON với ngày cụ thể
    # Format ngày: MM/DD/YYYY
    result = download_cts_gpon_report(report_date="01/24/2026")

    if result:
        print("\n" + "=" * 60)
        print("✅ HOÀN THÀNH!")
        print(f"   File đã tải về: {result}")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("❌ Có lỗi xảy ra trong quá trình thực hiện")
        print("=" * 60)
