# -*- coding: utf-8 -*-
import time
import os
from config import Config
from login import login_baocao_hanoi, read_otp_from_file

def download_report_c11(page_baocao, report_month):
    """
    Tải báo cáo C1.1 CHỈ TIÊU 1.1 (Q4/2025) từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập
        report_month: Tháng báo cáo (ví dụ: "Tháng 12/2025", "Tháng 01/2026")

    Các bước:
    1. Truy cập URL báo cáo C1.1
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Chọn tháng báo cáo
    6. Click button "Báo cáo"
    7. Click button "Xuất Excel"
    8. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO CHỈ TIÊU 1.1 - {report_month}")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=522457&menu_id=522561"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn tháng báo cáo
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            month_select = page_baocao.locator("div.search-criteria > div:nth-child(2) select")
            month_select.wait_for(state="visible", timeout=10000)
            month_select.select_option(label=report_month)
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 6: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=10000)
        baocao_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 6: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 7: Click icon download thứ 2 để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c1.1 report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.1 CHỈ TIÊU 1.1 (Q4/2025)")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.1 CHỈ TIÊU 1.1 (Q4/2025): {e}")
        import traceback
        traceback.print_exc()
        return False

def download_report_c11_chitiet(page_baocao, start_date="01/01/2026", end_date="31/01/2026"):
    """
    Tải báo cáo SM4-C11 3.5.1 Sửa chữa toàn trình chi tiết

    Args:
        page_baocao: Đối tượng page đã đăng nhập
        start_date: Ngày bắt đầu (định dạng "dd/mm/yyyy")
        end_date: Ngày kết thúc (định dạng "dd/mm/yyyy")
    """
    print(f"\n=== Bắt đầu tải báo cáo SM4-C11 3.5.1 Sửa chữa toàn trình chi tiết ({start_date} - {end_date}) ===")

    # Truy cập trang báo cáo
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194'
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=Config.DOWNLOAD_TIMEOUT)
    time.sleep(3)

    # Bước 1: Click vào dropdown chọn đơn vị
    print("Đang chọn đơn vị TTVT Sơn Tây...")
    try:
        # Click vào button dropdown
        dropdown_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > button"
        page_baocao.click(dropdown_selector)
        time.sleep(2)
        print("✅ Đã click dropdown đơn vị")

        # Tìm input search và điền "ttvt sơn tây"
        search_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div.treeview-header > div.row.row-filter.ng-star-inserted > div > input"
        page_baocao.fill(search_selector, "ttvt sơn tây")
        time.sleep(2)
        print("✅ Đã điền từ khóa tìm kiếm")

        # Click vào TTVT Sơn Tây
        ttvt_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-valid.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div:nth-child(2) > div > ngx-treeview-item > div > div:nth-child(2) > ngx-treeview-item > div > div > span"
        page_baocao.click(ttvt_selector)
        time.sleep(2)
        print("✅ Đã chọn TTVT Sơn Tây")
    except Exception as e:
        print(f"⚠️ Lỗi khi chọn đơn vị: {e}")

    # Bước 2: Điền Từ ngày
    print(f"\n✓ Đang điền Từ ngày: {start_date}...")
    try:
        from_date_field = page_baocao.locator('#mat-input-0')
        from_date_field.wait_for(state="visible", timeout=30000)
        # Xóa text hiện tại
        from_date_field.click()
        time.sleep(0.5)
        from_date_field.press("Control+A")
        time.sleep(0.5)
        # Điền ngày mới
        from_date_field.fill(start_date)
        time.sleep(1)
        print(f"✅ Đã điền Từ ngày: {start_date}")

        # Đóng datepicker
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
    except Exception as e:
        print(f"⚠️ Lỗi khi điền Từ ngày: {e}")

    # Bước 2b: Điền Đến ngày
    print(f"\n✓ Đang điền Đến ngày: {end_date}...")
    try:
        to_date_field = page_baocao.locator('#mat-input-1')
        to_date_field.wait_for(state="visible", timeout=30000)
        # Xóa text hiện tại
        to_date_field.click()
        time.sleep(0.5)
        to_date_field.press("Control+A")
        time.sleep(0.5)
        # Điền ngày mới
        to_date_field.fill(end_date)
        time.sleep(1)
        print(f"✅ Đã điền Đến ngày: {end_date}")

        # Đóng datepicker
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

        # Đợi backdrop biến mất
        try:
            page_baocao.wait_for_selector(".cdk-overlay-backdrop", state="hidden", timeout=5000)
            print("✅ Datepicker đã đóng")
        except:
            print("✅ Không có backdrop hoặc đã đóng")

        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi khi điền Đến ngày: {e}")

    # Bước 3: Click dropdown "Dịch vụ" (div:nth-child(4))
    print("\nBước 3: Đang click dropdown 'Dịch vụ'...")
    try:
        # Đợi UI ổn định sau khi đóng datepicker
        time.sleep(2)

        # Tìm button trong div:nth-child(4) - đây là dropdown "Dịch vụ"
        div4_button = page_baocao.locator("div.search-criteria > div:nth-child(4) button")
        div4_button.wait_for(state="visible", timeout=30000)

        # Kiểm tra backdrop đã biến mất chưa
        backdrop = page_baocao.locator(".cdk-overlay-backdrop")
        if backdrop.count() > 0 and backdrop.is_visible():
            print("⚠️ Vẫn còn backdrop, đang đóng...")
            page_baocao.keyboard.press("Escape")
            time.sleep(1)

        # Click vào dropdown
        div4_button.click()
        time.sleep(2)
        print("✅ Đã click dropdown 'Dịch vụ'")

        # Tìm và click checkbox "Tất cả"
        print("Đang tìm checkbox 'Tất cả'...")

        # Thử cách 1: Click vào label thay vì checkbox
        try:
            label_all = page_baocao.locator("div.treeview-header div.row-all label.form-check-label").first
            label_all.wait_for(state="visible", timeout=5000)
            label_all.click()
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (click label)")
        except:
            # Cách 2: Force click checkbox
            print("Đang thử force click checkbox...")
            checkbox_all = page_baocao.locator("div.treeview-header div.row-all input[type='checkbox']").first
            checkbox_all.click(force=True)
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (force click)")

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)            

        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 3: {e}")
        import traceback
        traceback.print_exc()

    # Bước 4: Chọn "SM4 C11" trong dropdown "Loại phiếu" (sử dụng select element)
    print("\nBước 4: Đang chọn 'SM4 C11' trong dropdown 'Loại phiếu'...")
    try:
        # Đợi UI ổn định
        time.sleep(1)

        # Sử dụng CSS selector cho thẻ select trong div:nth-child(7)
        select_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(7) > div > div > div > div > select"

        select_element = page_baocao.locator(select_selector)
        select_element.wait_for(state="visible", timeout=30000)
        print("✅ Đã tìm thấy select dropdown 'Loại phiếu'")

        # Chọn option "SM4 C11" trong select dropdown
        # Có thể chọn bằng text hoặc value
        select_element.select_option(label="SM4 C11")
        time.sleep(2)
        print("✅ Đã chọn 'SM4 C11' cho Loại phiếu")

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 4: {e}")
        import traceback
        traceback.print_exc()

        # Thử cách khác nếu select_option không hoạt động
        print("Đang thử cách khác...")
        try:
            select_selector = "div.search-criteria > div:nth-child(7) select"
            select_element = page_baocao.locator(select_selector)
            select_element.wait_for(state="visible", timeout=30000)
            select_element.select_option(label="SM4 C11")
            time.sleep(2)
            print("✅ Đã chọn 'SM4 C11' (cách 2)")
        except Exception as e2:
            print(f"⚠️ Cách 2 cũng thất bại: {e2}")

    # Bước 5: Click button "Báo cáo"
    print("\nBước 5: Đang click button 'Báo cáo'...")
    try:
        time.sleep(2)

        # Tìm button "Báo cáo" bằng text
        baocao_btn = page_baocao.get_by_role("button", name="Báo cáo")
        baocao_btn.wait_for(state="visible", timeout=30000)
        baocao_btn.click()
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(2)

        # Đợi ngx-loading overlay biến mất (quan trọng cho Linux)
        print("Đang đợi loading overlay biến mất...")
        try:
            # Đợi ngx-loading element hidden
            page_baocao.wait_for_selector("ngx-loading .backdrop", state="hidden", timeout=60000)
            print("✅ Loading overlay đã biến mất")
        except:
            # Nếu không tìm thấy hoặc đã biến mất
            try:
                # Thử selector khác
                page_baocao.wait_for_selector(".backdrop.full-screen", state="hidden", timeout=10000)
                print("✅ Loading overlay đã biến mất")
            except:
                print("✅ Không có loading overlay hoặc đã biến mất")

        time.sleep(1)  # Buffer an toàn
        print("✅ Dữ liệu đã load xong")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 5: {e}")
        import traceback
        traceback.print_exc()

    # Bước 6: Click button "Xuất Excel" để mở dropdown
    print("\nBước 6: Đang click button 'Xuất Excel'...")
    try:
        time.sleep(2)

        # Tìm button "Xuất Excel" - đây là dropdown button
        xuatexcel_btn = page_baocao.get_by_role("button", name="Xuất Excel")
        xuatexcel_btn.wait_for(state="visible", timeout=30000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 6: {e}")
        import traceback
        traceback.print_exc()

    # Bước 7: Click vào option "2.Tất cả dữ liệu" trong dropdown
    print("\nBước 7: Đang click '2.Tất cả dữ liệu' để tải file...")
    try:
        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        time.sleep(2)

        # Sử dụng CSS selector chính xác để click vào option "2.Tất cả dữ liệu"
        download_option_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_option = page_baocao.locator(download_option_selector)
        download_option.wait_for(state="visible", timeout=30000)
        print("✅ Đã tìm thấy option '2.Tất cả dữ liệu'")

        # Bắt đầu theo dõi download
        with page_baocao.expect_download(timeout=300000) as download_info:
            download_option.click()
            print("✅ Đã click vào '2.Tất cả dữ liệu'")

        # Lưu file với tên mới
        download = download_info.value
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lấy extension từ file gốc
        file_extension = os.path.splitext(original_filename)[1]
        new_filename = f"SM4-C11{file_extension}"

        save_path = os.path.join(download_dir, new_filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 7: {e}")
        import traceback
        traceback.print_exc()

        # Thử cách khác nếu selector đầy đủ không hoạt động
        print("Đang thử cách khác...")
        try:
            # Thử selector ngắn hơn
            download_option = page_baocao.locator("div.button-action i:nth-child(2)")
            download_option.wait_for(state="visible", timeout=30000)

            with page_baocao.expect_download(timeout=300000) as download_info:
                download_option.click()
                print("✅ Đã click vào option download (cách 2)")

            download = download_info.value
            original_filename = download.suggested_filename
            file_extension = os.path.splitext(original_filename)[1]
            new_filename = f"SM4 C11{file_extension}"
            save_path = os.path.join(download_dir, new_filename)
            download.save_as(save_path)
            print(f"✅ Đã tải file về: {save_path}")
        except Exception as e2:
            print(f"⚠️ Cách 2 cũng thất bại: {e2}")


def download_report_c11_chitiet_SM2(page_baocao, start_date="01/01/2026", end_date="31/01/2026"):
    """
    Tải báo cáo xuất chủ động suy hao cao SM2 C11 chi tiết

    Args:
        page_baocao: Đối tượng page đã đăng nhập
        start_date: Ngày bắt đầu (định dạng "dd/mm/yyyy")
        end_date: Ngày kết thúc (định dạng "dd/mm/yyyy")
    """
    print(f"\n=== Bắt đầu tải báo cáo SM2-C11 chi tiết 3.5.1 Sửa chữa toàn trình chi tiết ({start_date} - {end_date}) ===")

    # Truy cập trang báo cáo
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194'
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=Config.DOWNLOAD_TIMEOUT)
    time.sleep(3)

    # Bước 1: Click vào dropdown chọn đơn vị
    print("Đang chọn đơn vị TTVT Sơn Tây...")
    try:
        # Click vào button dropdown
        dropdown_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > button"
        page_baocao.click(dropdown_selector)
        time.sleep(2)
        print("✅ Đã click dropdown đơn vị")

        # Tìm input search và điền "ttvt sơn tây"
        search_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div.treeview-header > div.row.row-filter.ng-star-inserted > div > input"
        page_baocao.fill(search_selector, "ttvt sơn tây")
        time.sleep(2)
        print("✅ Đã điền từ khóa tìm kiếm")

        # Click vào TTVT Sơn Tây
        ttvt_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-valid.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div:nth-child(2) > div > ngx-treeview-item > div > div:nth-child(2) > ngx-treeview-item > div > div > span"
        page_baocao.click(ttvt_selector)
        time.sleep(2)
        print("✅ Đã chọn TTVT Sơn Tây")
    except Exception as e:
        print(f"⚠️ Lỗi khi chọn đơn vị: {e}")

    # Bước 2: Điền Từ ngày
    print(f"\n✓ Đang điền Từ ngày: {start_date}...")
    try:
        from_date_field = page_baocao.locator('#mat-input-0')
        from_date_field.wait_for(state="visible", timeout=30000)
        # Xóa text hiện tại
        from_date_field.click()
        time.sleep(0.5)
        from_date_field.press("Control+A")
        time.sleep(0.5)
        # Điền ngày mới
        from_date_field.fill(start_date)
        time.sleep(1)
        print(f"✅ Đã điền Từ ngày: {start_date}")

        # Đóng datepicker
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
    except Exception as e:
        print(f"⚠️ Lỗi khi điền Từ ngày: {e}")

    # Bước 2b: Điền Đến ngày
    print(f"\n✓ Đang điền Đến ngày: {end_date}...")
    try:
        to_date_field = page_baocao.locator('#mat-input-1')
        to_date_field.wait_for(state="visible", timeout=30000)
        # Xóa text hiện tại
        to_date_field.click()
        time.sleep(0.5)
        to_date_field.press("Control+A")
        time.sleep(0.5)
        # Điền ngày mới
        to_date_field.fill(end_date)
        time.sleep(1)
        print(f"✅ Đã điền Đến ngày: {end_date}")

        # Đóng datepicker
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

        # Đợi backdrop biến mất
        try:
            page_baocao.wait_for_selector(".cdk-overlay-backdrop", state="hidden", timeout=5000)
            print("✅ Datepicker đã đóng")
        except:
            print("✅ Không có backdrop hoặc đã đóng")

        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi khi điền Đến ngày: {e}")

    # Bước 3: Click dropdown "Dịch vụ" (div:nth-child(4))
    print("\nBước 3: Đang click dropdown 'Dịch vụ'...")
    try:
        # Đợi UI ổn định sau khi đóng datepicker
        time.sleep(2)

        # Tìm button trong div:nth-child(4) - đây là dropdown "Dịch vụ"
        div4_button = page_baocao.locator("div.search-criteria > div:nth-child(4) button")
        div4_button.wait_for(state="visible", timeout=30000)

        # Kiểm tra backdrop đã biến mất chưa
        backdrop = page_baocao.locator(".cdk-overlay-backdrop")
        if backdrop.count() > 0 and backdrop.is_visible():
            print("⚠️ Vẫn còn backdrop, đang đóng...")
            page_baocao.keyboard.press("Escape")
            time.sleep(1)

        # Click vào dropdown
        div4_button.click()
        time.sleep(2)
        print("✅ Đã click dropdown 'Dịch vụ'")

        # Tìm và click checkbox "Tất cả"
        print("Đang tìm checkbox 'Tất cả'...")

        # Thử cách 1: Click vào label thay vì checkbox
        try:
            label_all = page_baocao.locator("div.treeview-header div.row-all label.form-check-label").first
            label_all.wait_for(state="visible", timeout=5000)
            label_all.click()
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (click label)")
        except:
            # Cách 2: Force click checkbox
            print("Đang thử force click checkbox...")
            checkbox_all = page_baocao.locator("div.treeview-header div.row-all input[type='checkbox']").first
            checkbox_all.click(force=True)
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (force click)")

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)            

        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 3: {e}")
        import traceback
        traceback.print_exc()

    # Bước 4: Chọn "SM2 C11" trong dropdown "Loại phiếu" (sử dụng select element)
    print("\nBước 4: Đang chọn 'SM2 C11' trong dropdown 'Loại phiếu'...")
    try:
        # Đợi UI ổn định
        time.sleep(1)

        # Sử dụng CSS selector cho thẻ select trong div:nth-child(7)
        select_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(7) > div > div > div > div > select"

        select_element = page_baocao.locator(select_selector)
        select_element.wait_for(state="visible", timeout=30000)
        print("✅ Đã tìm thấy select dropdown 'Loại phiếu'")

        # Chọn option "SM2 C11" trong select dropdown
        # Có thể chọn bằng text hoặc value
        select_element.select_option(label="SM2 C11")
        time.sleep(2)
        print("✅ Đã chọn 'SM2 C11' cho Loại phiếu")

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 4: {e}")
        import traceback
        traceback.print_exc()

        # Thử cách khác nếu select_option không hoạt động
        print("Đang thử cách khác...")
        try:
            select_selector = "div.search-criteria > div:nth-child(7) select"
            select_element = page_baocao.locator(select_selector)
            select_element.wait_for(state="visible", timeout=30000)
            select_element.select_option(label="SM2 C11")
            time.sleep(2)
            print("✅ Đã chọn 'SM2 C11' (cách 2)")
        except Exception as e2:
            print(f"⚠️ Cách 2 cũng thất bại: {e2}")

    # Bước 5: Click button "Báo cáo"
    print("\nBước 5: Đang click button 'Báo cáo'...")
    try:
        time.sleep(2)

        # Tìm button "Báo cáo" bằng text
        baocao_btn = page_baocao.get_by_role("button", name="Báo cáo")
        baocao_btn.wait_for(state="visible", timeout=30000)
        baocao_btn.click()
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(2)

        # Đợi ngx-loading overlay biến mất (quan trọng cho Linux)
        print("Đang đợi loading overlay biến mất...")
        try:
            # Đợi ngx-loading element hidden
            page_baocao.wait_for_selector("ngx-loading .backdrop", state="hidden", timeout=60000)
            print("✅ Loading overlay đã biến mất")
        except:
            # Nếu không tìm thấy hoặc đã biến mất
            try:
                # Thử selector khác
                page_baocao.wait_for_selector(".backdrop.full-screen", state="hidden", timeout=10000)
                print("✅ Loading overlay đã biến mất")
            except:
                print("✅ Không có loading overlay hoặc đã biến mất")

        time.sleep(1)  # Buffer an toàn
        print("✅ Dữ liệu đã load xong")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 5: {e}")
        import traceback
        traceback.print_exc()

    # Bước 6: Click button "Xuất Excel" để mở dropdown
    print("\nBước 6: Đang click button 'Xuất Excel'...")
    try:
        time.sleep(2)

        # Tìm button "Xuất Excel" - đây là dropdown button
        xuatexcel_btn = page_baocao.get_by_role("button", name="Xuất Excel")
        xuatexcel_btn.wait_for(state="visible", timeout=30000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 6: {e}")
        import traceback
        traceback.print_exc()

    # Bước 7: Click vào option "2.Tất cả dữ liệu" trong dropdown
    print("\nBước 7: Đang click '2.Tất cả dữ liệu' để tải file...")
    try:
        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        time.sleep(2)

        # Sử dụng CSS selector chính xác để click vào option "2.Tất cả dữ liệu"
        download_option_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_option = page_baocao.locator(download_option_selector)
        download_option.wait_for(state="visible", timeout=30000)
        print("✅ Đã tìm thấy option '2.Tất cả dữ liệu'")

        # Bắt đầu theo dõi download
        with page_baocao.expect_download(timeout=300000) as download_info:
            download_option.click()
            print("✅ Đã click vào '2.Tất cả dữ liệu'")

        # Lưu file với tên mới
        download = download_info.value
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lấy extension từ file gốc
        file_extension = os.path.splitext(original_filename)[1]
        new_filename = f"SM2-C11{file_extension}"

        save_path = os.path.join(download_dir, new_filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 7: {e}")
        import traceback
        traceback.print_exc()

        # Thử cách khác nếu selector đầy đủ không hoạt động
        print("Đang thử cách khác...")
        try:
            # Thử selector ngắn hơn
            download_option = page_baocao.locator("div.button-action i:nth-child(2)")
            download_option.wait_for(state="visible", timeout=30000)

            with page_baocao.expect_download(timeout=300000) as download_info:
                download_option.click()
                print("✅ Đã click vào option download (cách 2)")

            download = download_info.value
            original_filename = download.suggested_filename
            file_extension = os.path.splitext(original_filename)[1]
            new_filename = f"SM2 C11{file_extension}"
            save_path = os.path.join(download_dir, new_filename)
            download.save_as(save_path)
            print(f"✅ Đã tải file về: {save_path}")
        except Exception as e2:
            print(f"⚠️ Cách 2 cũng thất bại: {e2}")

def download_report_c12_chitiet_SM1(page_baocao, start_date="01/01/2026", end_date="31/01/2026"):
    """
    Tải báo cáo xuất chủ động suy hao cao SM1 C12 chi tiết

    Args:
        page_baocao: Đối tượng page đã đăng nhập
        start_date: Ngày bắt đầu (định dạng "dd/mm/yyyy")
        end_date: Ngày kết thúc (định dạng "dd/mm/yyyy")
    """
    print(f"\n=== Bắt đầu tải báo cáo SM1 C12 chi tiết 3.5.1 Sửa chữa toàn trình chi tiết ({start_date} - {end_date}) ===")

    # Truy cập trang báo cáo
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194'
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=Config.DOWNLOAD_TIMEOUT)
    time.sleep(3)

    # Bước 1: Click vào dropdown chọn đơn vị
    print("Đang chọn đơn vị TTVT Sơn Tây...")
    try:
        # Click vào button dropdown
        dropdown_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > button"
        page_baocao.click(dropdown_selector)
        time.sleep(2)
        print("✅ Đã click dropdown đơn vị")

        # Tìm input search và điền "ttvt sơn tây"
        search_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div.treeview-header > div.row.row-filter.ng-star-inserted > div > input"
        page_baocao.fill(search_selector, "ttvt sơn tây")
        time.sleep(2)
        print("✅ Đã điền từ khóa tìm kiếm")

        # Click vào TTVT Sơn Tây
        ttvt_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-valid.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div:nth-child(2) > div > ngx-treeview-item > div > div:nth-child(2) > ngx-treeview-item > div > div > span"
        page_baocao.click(ttvt_selector)
        time.sleep(2)
        print("✅ Đã chọn TTVT Sơn Tây")
    except Exception as e:
        print(f"⚠️ Lỗi khi chọn đơn vị: {e}")

    # Bước 2: Điền Từ ngày
    print(f"\n✓ Đang điền Từ ngày: {start_date}...")
    try:
        from_date_field = page_baocao.locator('#mat-input-0')
        from_date_field.wait_for(state="visible", timeout=30000)
        from_date_field.click()
        time.sleep(0.5)
        from_date_field.press("Control+A")
        time.sleep(0.5)
        from_date_field.fill(start_date)
        time.sleep(1)
        print(f"✅ Đã điền Từ ngày: {start_date}")
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
    except Exception as e:
        print(f"⚠️ Lỗi khi điền Từ ngày: {e}")

    # Bước 2b: Điền Đến ngày
    print(f"\n✓ Đang điền Đến ngày: {end_date}...")
    try:
        to_date_field = page_baocao.locator('#mat-input-1')
        to_date_field.wait_for(state="visible", timeout=30000)
        to_date_field.click()
        time.sleep(0.5)
        to_date_field.press("Control+A")
        time.sleep(0.5)
        to_date_field.fill(end_date)
        time.sleep(1)
        print(f"✅ Đã điền Đến ngày: {end_date}")
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

        try:
            page_baocao.wait_for_selector(".cdk-overlay-backdrop", state="hidden", timeout=5000)
            print("✅ Datepicker đã đóng")
        except:
            print("✅ Không có backdrop hoặc đã đóng")

        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi khi điền Đến ngày: {e}")

    # Bước 3: Click dropdown "Dịch vụ" (div:nth-child(4))
    print("\nBước 3: Đang click dropdown 'Dịch vụ'...")
    try:
        # Đợi UI ổn định sau khi đóng datepicker
        time.sleep(2)

        # Tìm button trong div:nth-child(4) - đây là dropdown "Dịch vụ"
        div4_button = page_baocao.locator("div.search-criteria > div:nth-child(4) button")
        div4_button.wait_for(state="visible", timeout=30000)

        # Kiểm tra backdrop đã biến mất chưa
        backdrop = page_baocao.locator(".cdk-overlay-backdrop")
        if backdrop.count() > 0 and backdrop.is_visible():
            print("⚠️ Vẫn còn backdrop, đang đóng...")
            page_baocao.keyboard.press("Escape")
            time.sleep(1)

        # Click vào dropdown
        div4_button.click()
        time.sleep(2)
        print("✅ Đã click dropdown 'Dịch vụ'")

        # Tìm và click checkbox "Tất cả"
        print("Đang tìm checkbox 'Tất cả'...")

        # Thử cách 1: Click vào label thay vì checkbox
        try:
            label_all = page_baocao.locator("div.treeview-header div.row-all label.form-check-label").first
            label_all.wait_for(state="visible", timeout=5000)
            label_all.click()
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (click label)")
        except:
            # Cách 2: Force click checkbox
            print("Đang thử force click checkbox...")
            checkbox_all = page_baocao.locator("div.treeview-header div.row-all input[type='checkbox']").first
            checkbox_all.click(force=True)
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (force click)")

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)            

        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 3: {e}")
        import traceback
        traceback.print_exc()

    # Bước 4: Chọn "SM2 C11" trong dropdown "Loại phiếu" (sử dụng select element)
    print("\nBước 4: Đang chọn 'SM1 C12' trong dropdown 'Loại phiếu'...")
    try:
        # Đợi UI ổn định
        time.sleep(1)

        # Sử dụng CSS selector cho thẻ select trong div:nth-child(7)
        select_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(7) > div > div > div > div > select"

        select_element = page_baocao.locator(select_selector)
        select_element.wait_for(state="visible", timeout=30000)
        print("✅ Đã tìm thấy select dropdown 'Loại phiếu'")

        # Chọn option "SM2 C11" trong select dropdown
        # Có thể chọn bằng text hoặc value
        select_element.select_option(label="SM1 C12")
        time.sleep(2)
        print("✅ Đã chọn 'SM1 C12' cho Loại phiếu")

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 4: {e}")
        import traceback
        traceback.print_exc()

        # Thử cách khác nếu select_option không hoạt động
        print("Đang thử cách khác...")
        try:
            select_selector = "div.search-criteria > div:nth-child(7) select"
            select_element = page_baocao.locator(select_selector)
            select_element.wait_for(state="visible", timeout=30000)
            select_element.select_option(label="SM1 C12")
            time.sleep(2)
            print("✅ Đã chọn 'SM1 C12' (cách 2)")
        except Exception as e2:
            print(f"⚠️ Cách 2 cũng thất bại: {e2}")

    # Bước 5: Click button "Báo cáo"
    print("\nBước 5: Đang click button 'Báo cáo'...")
    try:
        time.sleep(2)

        # Tìm button "Báo cáo" bằng text
        baocao_btn = page_baocao.get_by_role("button", name="Báo cáo")
        baocao_btn.wait_for(state="visible", timeout=30000)
        baocao_btn.click()
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(2)

        # Đợi ngx-loading overlay biến mất (quan trọng cho Linux)
        print("Đang đợi loading overlay biến mất...")
        try:
            # Đợi ngx-loading element hidden
            page_baocao.wait_for_selector("ngx-loading .backdrop", state="hidden", timeout=60000)
            print("✅ Loading overlay đã biến mất")
        except:
            # Nếu không tìm thấy hoặc đã biến mất
            try:
                # Thử selector khác
                page_baocao.wait_for_selector(".backdrop.full-screen", state="hidden", timeout=10000)
                print("✅ Loading overlay đã biến mất")
            except:
                print("✅ Không có loading overlay hoặc đã biến mất")

        time.sleep(1)  # Buffer an toàn
        print("✅ Dữ liệu đã load xong")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 5: {e}")
        import traceback
        traceback.print_exc()

    # Bước 6: Click button "Xuất Excel" để mở dropdown
    print("\nBước 6: Đang click button 'Xuất Excel'...")
    try:
        time.sleep(2)

        # Tìm button "Xuất Excel" - đây là dropdown button
        xuatexcel_btn = page_baocao.get_by_role("button", name="Xuất Excel")
        xuatexcel_btn.wait_for(state="visible", timeout=30000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 6: {e}")
        import traceback
        traceback.print_exc()

    # Bước 7: Click vào option "2.Tất cả dữ liệu" trong dropdown
    print("\nBước 7: Đang click '2.Tất cả dữ liệu' để tải file...")
    try:
        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        time.sleep(2)

        # Sử dụng CSS selector chính xác để click vào option "2.Tất cả dữ liệu"
        download_option_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_option = page_baocao.locator(download_option_selector)
        download_option.wait_for(state="visible", timeout=30000)
        print("✅ Đã tìm thấy option '2.Tất cả dữ liệu'")

        # Bắt đầu theo dõi download
        with page_baocao.expect_download(timeout=300000) as download_info:
            download_option.click()
            print("✅ Đã click vào '2.Tất cả dữ liệu'")

        # Lưu file với tên mới
        download = download_info.value
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lấy extension từ file gốc
        file_extension = os.path.splitext(original_filename)[1]
        new_filename = f"SM1-C12{file_extension}"

        save_path = os.path.join(download_dir, new_filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 7: {e}")
        import traceback
        traceback.print_exc()

        # Thử cách khác nếu selector đầy đủ không hoạt động
        print("Đang thử cách khác...")
        try:
            # Thử selector ngắn hơn
            download_option = page_baocao.locator("div.button-action i:nth-child(2)")
            download_option.wait_for(state="visible", timeout=30000)

            with page_baocao.expect_download(timeout=300000) as download_info:
                download_option.click()
                print("✅ Đã click vào option download (cách 2)")

            download = download_info.value
            original_filename = download.suggested_filename
            file_extension = os.path.splitext(original_filename)[1]
            new_filename = f"SM2 C11{file_extension}"
            save_path = os.path.join(download_dir, new_filename)
            download.save_as(save_path)
            print(f"✅ Đã tải file về: {save_path}")
        except Exception as e2:
            print(f"⚠️ Cách 2 cũng thất bại: {e2}")


def download_report_c12_chitiet_SM2(page_baocao, start_date="01/01/2026", end_date="31/01/2026"):
    """
    Tải báo cáo xuất chủ động suy hao cao SM2 C12 chi tiết

    Args:
        page_baocao: Đối tượng page đã đăng nhập
        start_date: Ngày bắt đầu (định dạng "dd/mm/yyyy")
        end_date: Ngày kết thúc (định dạng "dd/mm/yyyy")
    """
    print(f"\n=== Bắt đầu tải báo cáo SM2 C12 chi tiết ({start_date} - {end_date}) ===")

    # Truy cập trang báo cáo
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194'
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=Config.DOWNLOAD_TIMEOUT)
    time.sleep(3)

    # Bước 1: Click vào dropdown chọn đơn vị
    print("Đang chọn đơn vị TTVT Sơn Tây...")
    try:
        # Click vào button dropdown
        dropdown_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > button"
        page_baocao.click(dropdown_selector)
        time.sleep(2)
        print("✅ Đã click dropdown đơn vị")

        # Tìm input search và điền "ttvt sơn tây"
        search_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-pristine.ng-valid > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div.treeview-header > div.row.row-filter.ng-star-inserted > div > input"
        page_baocao.fill(search_selector, "ttvt sơn tây")
        time.sleep(2)
        print("✅ Đã điền từ khóa tìm kiếm")

        # Click vào TTVT Sơn Tây
        ttvt_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div.filter > div.ng-untouched.ng-valid.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(1) > div > div > div > div > div > div > ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > div > div > ngx-treeview > div:nth-child(2) > div > ngx-treeview-item > div > div:nth-child(2) > ngx-treeview-item > div > div > span"
        page_baocao.click(ttvt_selector)
        time.sleep(2)
        print("✅ Đã chọn TTVT Sơn Tây")
    except Exception as e:
        print(f"⚠️ Lỗi khi chọn đơn vị: {e}")

    # Bước 2: Điền Từ ngày
    print(f"\n✓ Đang điền Từ ngày: {start_date}...")
    try:
        from_date_field = page_baocao.locator('#mat-input-0')
        from_date_field.wait_for(state="visible", timeout=30000)
        from_date_field.click()
        time.sleep(0.5)
        from_date_field.press("Control+A")
        time.sleep(0.5)
        from_date_field.fill(start_date)
        time.sleep(1)
        print(f"✅ Đã điền Từ ngày: {start_date}")
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
    except Exception as e:
        print(f"⚠️ Lỗi khi điền Từ ngày: {e}")

    # Bước 2b: Điền Đến ngày
    print(f"\n✓ Đang điền Đến ngày: {end_date}...")
    try:
        to_date_field = page_baocao.locator('#mat-input-1')
        to_date_field.wait_for(state="visible", timeout=30000)
        to_date_field.click()
        time.sleep(0.5)
        to_date_field.press("Control+A")
        time.sleep(0.5)
        to_date_field.fill(end_date)
        time.sleep(1)
        print(f"✅ Đã điền Đến ngày: {end_date}")
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

        try:
            page_baocao.wait_for_selector(".cdk-overlay-backdrop", state="hidden", timeout=5000)
            print("✅ Datepicker đã đóng")
        except:
            print("✅ Không có backdrop hoặc đã đóng")

        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi khi điền Đến ngày: {e}")

    # Bước 3: Click dropdown "Dịch vụ" (div:nth-child(4))
    print("\nBước 3: Đang click dropdown 'Dịch vụ'...")
    try:
        # Đợi UI ổn định sau khi đóng datepicker
        time.sleep(2)

        # Tìm button trong div:nth-child(4) - đây là dropdown "Dịch vụ"
        div4_button = page_baocao.locator("div.search-criteria > div:nth-child(4) button")
        div4_button.wait_for(state="visible", timeout=30000)

        # Kiểm tra backdrop đã biến mất chưa
        backdrop = page_baocao.locator(".cdk-overlay-backdrop")
        if backdrop.count() > 0 and backdrop.is_visible():
            print("⚠️ Vẫn còn backdrop, đang đóng...")
            page_baocao.keyboard.press("Escape")
            time.sleep(1)

        # Click vào dropdown
        div4_button.click()
        time.sleep(2)
        print("✅ Đã click dropdown 'Dịch vụ'")

        # Tìm và click checkbox "Tất cả"
        print("Đang tìm checkbox 'Tất cả'...")

        # Thử cách 1: Click vào label thay vì checkbox
        try:
            label_all = page_baocao.locator("div.treeview-header div.row-all label.form-check-label").first
            label_all.wait_for(state="visible", timeout=5000)
            label_all.click()
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (click label)")
        except:
            # Cách 2: Force click checkbox
            print("Đang thử force click checkbox...")
            checkbox_all = page_baocao.locator("div.treeview-header div.row-all input[type='checkbox']").first
            checkbox_all.click(force=True)
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Dịch vụ (force click)")

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)            

        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)
        # Đóng dropdown
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 3: {e}")
        import traceback
        traceback.print_exc()
    print("Bước 4: không chọn gì, tự động chọn mục HTTT trong Loại phiếu mặc định.")

    # # Bước 4: Chọn "SM2 C11" trong dropdown "Loại phiếu" (sử dụng select element)
    # print("\nBước 4: Đang chọn 'SM1 C12' trong dropdown 'Loại phiếu'...")
    # try:
    #     # Đợi UI ổn định
    #     time.sleep(1)

    #     # Sử dụng CSS selector cho thẻ select trong div:nth-child(7)
    #     select_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.search-criteria > div:nth-child(7) > div > div > div > div > select"

    #     select_element = page_baocao.locator(select_selector)
    #     select_element.wait_for(state="visible", timeout=30000)
    #     print("✅ Đã tìm thấy select dropdown 'Loại phiếu'")

    #     # Chọn option "SM2 C11" trong select dropdown
    #     # Có thể chọn bằng text hoặc value
    #     select_element.select_option(label="SM1 C12")
    #     time.sleep(2)
    #     print("✅ Đã chọn 'SM1 C12' cho Loại phiếu")

    #     # Click vào trang để kích hoạt (activate) page
    #     print("Đang kích hoạt page...")
    #     page_baocao.click('body')
    #     time.sleep(1)

    # except Exception as e:
    #     print(f"⚠️ Lỗi ở bước 4: {e}")
    #     import traceback
    #     traceback.print_exc()

    #     # Thử cách khác nếu select_option không hoạt động
    #     print("Đang thử cách khác...")
    #     try:
    #         select_selector = "div.search-criteria > div:nth-child(7) select"
    #         select_element = page_baocao.locator(select_selector)
    #         select_element.wait_for(state="visible", timeout=30000)
    #         select_element.select_option(label="SM1 C12")
    #         time.sleep(2)
    #         print("✅ Đã chọn 'SM1 C12' (cách 2)")
    #     except Exception as e2:
    #         print(f"⚠️ Cách 2 cũng thất bại: {e2}")

    # Bước 5: Click button "Báo cáo"
    print("\nBước 5: Đang click button 'Báo cáo'...")
    try:
        time.sleep(2)

        # Tìm button "Báo cáo" bằng text
        baocao_btn = page_baocao.get_by_role("button", name="Báo cáo")
        baocao_btn.wait_for(state="visible", timeout=30000)
        baocao_btn.click()
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(2)

        # Đợi ngx-loading overlay biến mất (quan trọng cho Linux)
        print("Đang đợi loading overlay biến mất...")
        try:
            # Đợi ngx-loading element hidden
            page_baocao.wait_for_selector("ngx-loading .backdrop", state="hidden", timeout=60000)
            print("✅ Loading overlay đã biến mất")
        except:
            # Nếu không tìm thấy hoặc đã biến mất
            try:
                # Thử selector khác
                page_baocao.wait_for_selector(".backdrop.full-screen", state="hidden", timeout=10000)
                print("✅ Loading overlay đã biến mất")
            except:
                print("✅ Không có loading overlay hoặc đã biến mất")

        time.sleep(1)  # Buffer an toàn
        print("✅ Dữ liệu đã load xong")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 5: {e}")
        import traceback
        traceback.print_exc()

    # Bước 6: Click button "Xuất Excel" để mở dropdown
    print("\nBước 6: Đang click button 'Xuất Excel'...")
    try:
        time.sleep(2)

        # Tìm button "Xuất Excel" - đây là dropdown button
        xuatexcel_btn = page_baocao.get_by_role("button", name="Xuất Excel")
        xuatexcel_btn.wait_for(state="visible", timeout=30000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 6: {e}")
        import traceback
        traceback.print_exc()

    # Bước 7: Click vào option "2.Tất cả dữ liệu" trong dropdown
    print("\nBước 7: Đang click '2.Tất cả dữ liệu' để tải file...")
    try:
        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        time.sleep(2)

        # Sử dụng CSS selector chính xác để click vào option "2.Tất cả dữ liệu"
        download_option_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_option = page_baocao.locator(download_option_selector)
        download_option.wait_for(state="visible", timeout=30000)
        print("✅ Đã tìm thấy option '2.Tất cả dữ liệu'")

        # Bắt đầu theo dõi download
        with page_baocao.expect_download(timeout=300000) as download_info:
            download_option.click()
            print("✅ Đã click vào '2.Tất cả dữ liệu'")

        # Lưu file với tên mới
        download = download_info.value
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lấy extension từ file gốc
        file_extension = os.path.splitext(original_filename)[1]
        new_filename = f"SM2-C12{file_extension}"

        save_path = os.path.join(download_dir, new_filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 7: {e}")
        import traceback
        traceback.print_exc()

        # Thử cách khác nếu selector đầy đủ không hoạt động
        print("Đang thử cách khác...")
        try:
            # Thử selector ngắn hơn
            download_option = page_baocao.locator("div.button-action i:nth-child(2)")
            download_option.wait_for(state="visible", timeout=30000)

            with page_baocao.expect_download(timeout=300000) as download_info:
                download_option.click()
                print("✅ Đã click vào option download (cách 2)")

            download = download_info.value
            original_filename = download.suggested_filename
            file_extension = os.path.splitext(original_filename)[1]
            new_filename = f"SM2 C11{file_extension}"
            save_path = os.path.join(download_dir, new_filename)
            download.save_as(save_path)
            print(f"✅ Đã tải file về: {save_path}")
        except Exception as e2:
            print(f"⚠️ Cách 2 cũng thất bại: {e2}")

def download_report_c12(page_baocao, report_month="Tháng 01/2026"):
    """
    Tải báo cáo C1.2 từ hệ thống báo cáo VNPT Hà Nội - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại

    Args:
        page_baocao: Playwright page object đã đăng nhập
        report_month: Tháng báo cáo (ví dụ: "Tháng 12/2025", "Tháng 01/2026")

    Các bước:
    1. Truy cập URL báo cáo C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Chọn tháng báo cáo
    6. Click button "Báo cáo"
    7. Click button "Xuất Excel"
    8. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại - {report_month}")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=522562"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn tháng báo cáo
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            month_select = page_baocao.locator("div.search-criteria > div:nth-child(2) select")
            month_select.wait_for(state="visible", timeout=10000)
            month_select.select_option(label=report_month)
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 6: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=10000)
        baocao_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 6: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 7: Click icon download thứ 2 để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c1.2 report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại: {e}")
        import traceback
        traceback.print_exc()
        return False


def download_report_c13(page_baocao, report_month="Tháng 01/2026"):
    """
    Tải báo cáo C1.3 từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập
        report_month: Tháng báo cáo (ví dụ: "Tháng 12/2025", "Tháng 01/2026")

    Các bước:
    1. Truy cập URL báo cáo C1.3
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Chọn tháng báo cáo
    6. Click button "Báo cáo"
    7. Click button "Xuất Excel"
    8. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO C1.3 - Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định - {report_month}")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=522600&menu_id=522640"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn tháng báo cáo
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            month_select = page_baocao.locator("div.search-criteria > div:nth-child(2) select")
            month_select.wait_for(state="visible", timeout=10000)
            month_select.select_option(label=report_month)
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 6: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=10000)
        baocao_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 6: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 7: Click icon download thứ 2 để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c1.3 report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.3 -  Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.3 -  Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định: {e}")
        import traceback
        traceback.print_exc()
        return False


def download_report_c14(page_baocao, report_month="Tháng 01/2026"):
    """
    Tải báo cáo C1.4 từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập
        report_month: Tháng báo cáo (ví dụ: "Tháng 12/2025", "Tháng 01/2026")

    Các bước:
    1. Truy cập URL báo cáo C1.4
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Chọn tháng báo cáo
    6. Click vào select box và chọn option (Arrow Down 2 lần + Enter)
    7. Click button "Báo cáo"
    8. Click button "Xuất Excel"
    9. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO C1.4 - 6.2.2.Báo cáo CSKH tổng hợp - {report_month}")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=264107&menu_id=275688"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn tháng báo cáo (C1.4 dùng button dropdown, không phải select)
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            # Click vào button tháng trong div[2]
            month_btn = page_baocao.locator("div.search-criteria > div:nth-child(2) button")
            month_btn.wait_for(state="visible", timeout=10000)
            month_btn.click()
            time.sleep(1)
            print("✅ Đã mở dropdown tháng")
            
            # Tìm và click vào option tháng trong dropdown (dùng span để tránh trùng với button)
            month_option = page_baocao.locator(f"span.ng-option-label:has-text('{report_month}')")
            month_option.wait_for(state="visible", timeout=5000)
            month_option.click()
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 6: Click vào select box và chọn option
        print("\n✓ Đang click vào select box...")
        select_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/select"
        select_element = page_baocao.locator(f"xpath={select_xpath}")
        select_element.wait_for(state="visible", timeout=10000)
        select_element.click()
        time.sleep(1)
        print("✅ Đã click vào select box")

        print("\n✓ Đang nhấn Arrow Down 2 lần và Enter...")
        select_element.press("ArrowDown")
        time.sleep(0.5)
        select_element.press("ArrowDown")
        time.sleep(0.5)
        select_element.press("Enter")
        time.sleep(1)
        print("✅ Đã chọn option")

        # Bước 7: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=10000)
        baocao_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 7: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 8: Click icon download thứ 2 để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=120000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c1.4 report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.4 - 6.2.2.Báo cáo CSKH tổng hợp")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.4 - 6.2.2.Báo cáo CSKH tổng hợp: {e}")
        import traceback
        traceback.print_exc()
        return False

def download_report_c14_chitiet(page_baocao, report_month="Tháng 01/2026"):
    """
    Tải báo cáo C1.4 chi tiết từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập
        report_month: Tháng báo cáo (ví dụ: "Tháng 12/2025", "Tháng 01/2026")

    Các bước:
    1. Truy cập URL báo cáo C1.4 chi tiết
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Chọn tháng báo cáo
    6. Click button "Báo cáo"
    7. Click button "Xuất Excel"
    8. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO C1.4 chi tiết - 6.2.1.Báo cáo CSKH chi tiết - {report_month}")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=240277&menu_id=275687"
        print(f"\n✓ Đang truy cập - 6.2.1.Báo cáo CSKH chi tiết: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn tháng báo cáo (C1.4 chi tiết dùng select element)
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            month_select_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[2]/div/div/div/div/select"
            month_select = page_baocao.locator(f"xpath={month_select_xpath}")
            month_select.wait_for(state="visible", timeout=10000)
            month_select.select_option(label=report_month)
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 6: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=80000)
        baocao_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 7: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        
        # Đợi loading overlay biến mất
        try:
            loading_overlay = page_baocao.locator("ngx-loading .backdrop")
            loading_overlay.wait_for(state="hidden", timeout=60000)
            print("✓ Loading overlay đã biến mất")
        except:
            print("⚠️ Không detect được loading overlay, tiếp tục...")
        
        time.sleep(2)
        
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        
        # Thử click với force=True để bypass overlay nếu còn
        try:
            xuatexcel_btn.click(force=True, timeout=5000)
        except:
            print("⚠️ Click bình thường thất bại, dùng JavaScript...")
            page_baocao.evaluate("""
                const btn = document.querySelector('button.dropdown-toggle');
                if (btn) btn.click();
            """)
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 8: Click icon download thứ 2 để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=120000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c1.4_chitiet_report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.4 - 6.2.1.Báo cáo CSKH chi tiết")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.4 - 831018: {e}")
        import traceback
        traceback.print_exc()
        return False


def download_report_c15(page_baocao, report_month="Tháng 01/2026"):
    """
    Tải báo cáo C1.5 từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập
        report_month: Tháng báo cáo (ví dụ: "Tháng 12/2025", "Tháng 01/2026")

    Các bước:
    1. Truy cập URL báo cáo C1.5
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Chọn tháng báo cáo
    6. Click button "Báo cáo"
    7. Click button "Xuất Excel"
    8. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO C1.5 - Tỷ lệ thiết lập dịch vụ BRCĐ đạt thời gian quy định - {report_month}")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=258310&menu_id=305918"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn tháng báo cáo (C1.5 dùng button dropdown)
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            # Click vào button tháng trong div[2]
            month_btn = page_baocao.locator("div.search-criteria > div:nth-child(2) button")
            month_btn.wait_for(state="visible", timeout=10000)
            month_btn.click()
            time.sleep(1)
            print("✅ Đã mở dropdown tháng")
            
            # Tìm và click vào option tháng trong dropdown (dùng span để tránh trùng với button)
            month_option = page_baocao.locator(f"span.ng-option-label:has-text('{report_month}')")
            month_option.wait_for(state="visible", timeout=5000)
            month_option.click()
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
            
            # Đóng dropdown nếu còn mở
            page_baocao.keyboard.press("Escape")
            time.sleep(0.5)
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 6: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=10000)
        baocao_btn.click(force=True)
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 7: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        
        # Scroll xuống để button không bị navbar che
        page_baocao.evaluate("window.scrollBy(0, 300)")
        time.sleep(1)
        
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        
        # Thử click với force=True để bypass overlay
        try:
            xuatexcel_btn.click(force=True, timeout=5000)
            print("✅ Đã click button 'Xuất Excel'")
        except:
            print("⚠️ Click bình thường thất bại, dùng JavaScript...")
            page_baocao.evaluate("""
                const btn = document.querySelector('button.dropdown-toggle');
                if (btn) btn.click();
            """)
            print("✅ Đã click button 'Xuất Excel' bằng JavaScript")
        
        time.sleep(2)
        print("✅ Dropdown đã mở")

        # Bước 8: Click option "2.Tất cả dữ liệu" để tải file
        print("\n✓ Đang click vào option tải file...")
        
        # Bắt đầu tải file - dùng JavaScript để click trực tiếp
        with page_baocao.expect_download(timeout=60000) as download_info:
            # Click lại button Xuất Excel để đảm bảo dropdown mở
            xuatexcel_btn.click(force=True)
            time.sleep(0.5)
            
            # Dùng JavaScript để dispatch click event
            page_baocao.evaluate("""
                (function() {
                    const items = document.querySelectorAll('i.dropdown-item');
                    if (items.length >= 2) {
                        items[1].dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true}));
                    }
                })();
            """)
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c1.5 report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.5")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.4: {e}")
        import traceback
        traceback.print_exc()
        return False

def download_report_c15_chitiet(page_baocao):
    """
    Tải báo cáo C1.5 chi tiết từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập

    Các bước:
    1. Truy cập URL báo cáo C1.5 - Tỷ lệ thiết lập dịch vụ BRCD đạt thời gian quy định
    link chi tiết gốc: https://baocao.hanoi.vnpt.vn/report/report-info?id=258310&menu_id=279696
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Click vào select box và chọn option (Arrow Down 2 lần + Enter)
    6. Click button "Báo cáo"
    7. Click button "Xuất Excel"
    8. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TẢI BÁO CÁO 6.1.5 Chỉ tiêu C1.5 chi tiết")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info-data?id=522920&ploaibc=1&pdonvi_id=284656&pthoigianid=98944771"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=500000)
        time.sleep(20)
        print("✅ Đã tải trang thành công")
        
        # Đợi dữ liệu load hoàn toàn
        print("\\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=180000)
        time.sleep(20)
        print("✅ Dữ liệu đã load xong")

        # # Bước 2: Click vào dropdown button để mở treeview
        # print("\n✓ Đang click vào dropdown button...")
        # dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        # dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        # dropdown_btn.wait_for(state="visible", timeout=10000)
        # dropdown_btn.click()
        # time.sleep(1)
        # print("✅ Đã mở dropdown treeview")

        # # Bước 3: Tìm input search và nhập "sơn tây"
        # print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        # search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        # search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        # search_input.wait_for(state="visible", timeout=10000)
        # search_input.fill("sơn tây")
        # time.sleep(2)
        # print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        # print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        # sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        # sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        # sontay_option.wait_for(state="visible", timeout=10000)
        # sontay_option.click()
        # time.sleep(1)
        # print("✅ Đã chọn 'TTVT Sơn Tây'")



        # # Bước 6: Click button "Báo cáo"
        # print("\n✓ Đang click button 'Báo cáo'...")
        # baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        # baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        # baocao_btn.wait_for(state="visible", timeout=10000)
        # baocao_btn.click()
        # time.sleep(2)
        # print("✅ Đã click button 'Báo cáo'")

        # # Đợi dữ liệu load
        # print("\n✓ Đang đợi dữ liệu load...")
        # page_baocao.wait_for_load_state("networkidle", timeout=120000)
        # time.sleep(3)
        # print("✅ Dữ liệu đã load xong")

        # Bước 7: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-data > div > div:nth-child(1) > div > div > div > div > div.button-action > div > button"
        xuatexcel_btn = page_baocao.locator(xuatexcel_btn_selector)
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 8: Click icon download để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-data/div/div[1]/div/div/div/div/div[2]/div/div/i[3]"
        download_option = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_option.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=180000) as download_info:
            download_option.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c1.5_chitiet_report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO 6.1.5 Chỉ tiêu C1.5 chi tiết")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo 6.1.5 Chỉ tiêu C1.5: {e}")
        import traceback
        traceback.print_exc()
        return False

def download_report_I15(page_baocao):
    """
    Tải báo cáo I1.5 từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập

    Các bước:
    1. Truy cập URL báo cáo I1.5 3.3.4 Chi tiết sửa chữa SH chủ động theo Tập đoàn VNPT
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Click button "Báo cáo"
    6. Click button "Xuất Excel"
    7. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TẢI BÁO CÁO I1.5 - 3.3.4 Chi tiết sửa chữa SH chủ động theo Tập đoàn VNPT")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=283632&menu_id=283669"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")



        # Bước 6: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=10000)
        baocao_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 7: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 8: Click icon download thứ 2 để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "I1.5 report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO I1.5 - 3.3.4 Chi tiết sửa chữa SH chủ động theo Tập đoàn VNPT")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo I1.5 - 3.3.4 Chi tiết sửa chữa SH chủ động theo Tập đoàn VNPT: {e}")
        import traceback
        traceback.print_exc()
        return False


def download_report_I15_k2(page_baocao):
    """
    Tải báo cáo I1.5 K2 từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập

    Các bước:
    1. Truy cập URL báo cáo I1.5 K2
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Click button "Báo cáo"
    6. Click button "Xuất Excel"
    7. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TẢI BÁO CÁO I1.5 K2")
        print("="*80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=290125&menu_id=290161"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")



        # Bước 5: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        baocao_btn = page_baocao.locator(f"xpath={baocao_btn_xpath}")
        baocao_btn.wait_for(state="visible", timeout=10000)
        baocao_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Báo cáo'")

        # Đợi dữ liệu load
        print("\n✓ Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(3)
        print("✅ Dữ liệu đã load xong")

        # Bước 6: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 7: Click icon download thứ 2 để tải file
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "I1.5_k2 report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO I1.5 K2")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo I1.5 K2: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # Test hàm tải báo cáo
    result = login_baocao_hanoi()

    if result:
        page_baocao, browser_baocao, playwright_baocao = result

        # download_report_c11(page_baocao)
        # download_report_c12(page_baocao)
        # download_report_c13(page_baocao)
        # download_report_c11_chitiet_SM2(page_baocao)
        # download_report_c12_chitiet_SM1(page_baocao)
        # download_report_c12_chitiet_SM2(page_baocao)
        # download_report_c14_chitiet(page_baocao)
        #download_report_c15_chitiet(page_baocao)
        download_report_I15(page_baocao)
        download_report_I15_k2(page_baocao)

        # Đóng browser sau khi hoàn thành
        print("\n=== Đóng browser ===")
        browser_baocao.close()
        playwright_baocao.stop()
