# -*- coding: utf-8 -*-
"""
Module chứa các hàm download báo cáo KR6 và KR7
Có thể chạy standalone để test: python KR_download.py
"""
import time
import os
from datetime import datetime


def download_KR6_report_NVKT(page_baocao):
    """
    Tải báo cáo KR6 từ trang baocao.hanoi.vnpt.vn
    10.2.11.5.KR6.Tỷ lệ thuê bao hoàn thành gia hạn TTTC trong tháng T đạt 80% (ĐB mới))
    https://baocao.hanoi.vnpt.vn/report/report-info?id=521560&menu_id=521600

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo KR6 chi tiết NVKT===")

    # In ngày tra cứu để debug
    current_date = datetime.now().strftime("%d/%m/%Y")
    print(f"📅 Ngày tra cứu báo cáo: {current_date}")

    # Truy cập trang báo cáo
    #report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521560&menu_id=521600'
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521560' #mới
    print(f"🔗 URL báo cáo KR6 NVKT: {report_url}")
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=60000)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=120000)
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

        #chọn menu Loại BC
        # Click vào dropdown menu 'Loại BC'
        print("Đang chọn Loại BC...")
        page_baocao.locator('div.search-criteria > div').filter(has_text='Loại BC').locator('select, button').first.click()
        time.sleep(2)
        print("✅ Đã click dropdown Loại BC")
        # Nhấn phím arrow down 2 lần
        for _ in range(2):
            page_baocao.keyboard.press("ArrowDown")
            time.sleep(1)  # Đợi 1 giây giữa các lần nhấn

        # Nhấn Enter
        page_baocao.keyboard.press("Enter")
        time.sleep(2)  # Đợi 2 giây sau khi nhấn Enter
        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1) 

        # # Tìm và click menu "Loại dữ liệu"
        # print("Đang chọn Loại dữ liệu...")
        # page_baocao.locator('div.search-criteria > div').filter(has_text='Loại dữ liệu').locator('select, button').first.click()
        # time.sleep(2)
        # print("✅ Đã click dropdown Loại dữ liệu")

        # #Nhấn phím UP 1 lần sau đó Enter
        # page_baocao.keyboard.press("ArrowUp")
        # time.sleep(1)  # Đợi 1 giây sau khi nhấn
        # page_baocao.keyboard.press("Enter")
        # time.sleep(2)  # Đợi 2 giây sau khi nhấn Enter

        # 1. Click "Báo cáo" button
        print("\nĐang click button 'Báo cáo'...")
        button1_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > button"
        page_baocao.locator(button1_selector).click()
        print("✅ Đã click button 'Báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(5)

        # 2. Click "Xuất Excel" button
        print("\nĐang click button 'Xuất Excel'...")
        button2_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > button"
        page_baocao.locator(button2_selector).click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel'.")

        # 3. Click download icon and save file
        print("\nĐang click icon download...")
        download_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(download_selector).click()
            print("✅ Đã click icon download.")

        download = download_info.value

        # Lưu file với tên cố định = tên hàm + .xlsx (ghi đè file cũ)
        filename = "download_KR6_report_NVKT.xlsx"
        save_path = os.path.join(download_dir, filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo KR6 chi tiết NVKT: {e}")
        import traceback
        traceback.print_exc()


def download_KR6_report_tong_hop(page_baocao):
    """
    Tải báo cáo KR6 từ trang baocao.hanoi.vnpt.vn
    https://baocao.hanoi.vnpt.vn/report/report-info?id=521560&menu_id=521600

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo KR6 TỔNG HỢP===")

    # In ngày tra cứu để debug
    current_date = datetime.now().strftime("%d/%m/%Y")
    print(f"📅 Ngày tra cứu báo cáo: {current_date}")

    # Truy cập trang báo cáo
    #report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521560&menu_id=521600'
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521560' #mới
    print(f"🔗 URL báo cáo KR6 Tổng hợp: {report_url}")
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=60000)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=120000)
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

        # #chọn menu Loại BC
        # # Click vào dropdown menu 'Loại BC'
        # print("Đang chọn Loại BC...")
        # page_baocao.locator('div.search-criteria > div').filter(has_text='Loại BC').locator('select, button').first.click()
        # time.sleep(2)
        # print("✅ Đã click dropdown Loại BC")
        # # Nhấn phím arrow down 2 lần
        # for _ in range(2):
        #     page_baocao.keyboard.press("ArrowDown")
        #     time.sleep(1)  # Đợi 1 giây giữa các lần nhấn

        # # Nhấn Enter
        # page_baocao.keyboard.press("Enter")
        # time.sleep(2)  # Đợi 2 giây sau khi nhấn Enter

        # # Tìm và click menu "Loại dữ liệu"
        # print("Đang chọn Loại dữ liệu...")
        # page_baocao.locator('div.search-criteria > div').filter(has_text='Loại dữ liệu').locator('select, button').first.click()
        # time.sleep(2)
        # print("✅ Đã click dropdown Loại dữ liệu")

        # #Nhấn phím UP 1 lần sau đó Enter
        # page_baocao.keyboard.press("ArrowUp")
        # time.sleep(1)  # Đợi 1 giây sau khi nhấn
        # page_baocao.keyboard.press("Enter")
        # time.sleep(1)  # Đợi 1 giây sau khi nhấn Enter

        # Tìm và click menu "Loại dữ liệu"
        print("Đang chọn Loại dữ liệu...")
        page_baocao.locator('div.search-criteria > div').filter(has_text='Loại dữ liệu').locator('select, button').first.click()
        time.sleep(2)
        print("✅ Đã click dropdown Loại dữ liệu")

        #Nhấn phím UP 1 lần sau đó Enter
        page_baocao.keyboard.press("ArrowUp")
        time.sleep(1)  # Đợi 1 giây sau khi nhấn
        page_baocao.keyboard.press("Enter")
        time.sleep(1)  # Đợi 1 giây sau khi nhấn Enter

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)

        # 1. Click "Báo cáo" button
        print("\nĐang click button 'Báo cáo'...")
        button1_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > button"
        page_baocao.locator(button1_selector).click()
        print("✅ Đã click button 'Báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(1)

        # 2. Click "Xuất Excel" button
        print("\nĐang click button 'Xuất Excel'...")
        button2_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > button"
        page_baocao.locator(button2_selector).click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel'.")

        # 3. Click download icon and save file
        print("\nĐang click icon download...")
        download_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(download_selector).click()
            print("✅ Đã click icon download.")

        download = download_info.value

        # Lưu file với tên cố định = tên hàm + .xlsx (ghi đè file cũ)
        filename = "download_KR6_report_tong_hop.xlsx"
        save_path = os.path.join(download_dir, filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo KR6 Tổng hợp: {e}")
        import traceback
        traceback.print_exc()


def download_KR7_report_NVKT(page_baocao):
    """
    Tải báo cáo KR7 từ trang baocao.hanoi.vnpt.vn

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo KR7 chi tiết NVKT ===")

    # In ngày tra cứu để debug
    current_date = datetime.now().strftime("%d/%m/%Y")
    print(f"📅 Ngày tra cứu báo cáo: {current_date}")

    # Truy cập trang báo cáo
    #report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521580&menu_id=521601'
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521580' #mới
    print(f"🔗 URL báo cáo KR7 NVKT: {report_url}")
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=60000)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=120000)
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

        #chọn menu Loại BC
        # Click vào dropdown menu 'Loại BC'
        print("Đang chọn Loại BC...")
        page_baocao.locator('div.search-criteria > div').filter(has_text='Loại BC').locator('select, button').first.click()
        time.sleep(2)
        print("✅ Đã click dropdown Loại BC")
        # Nhấn phím arrow down 2 lần
        for _ in range(2):
            page_baocao.keyboard.press("ArrowDown")
            time.sleep(1)  # Đợi 1 giây giữa các lần nhấn

        # Nhấn Enter
        page_baocao.keyboard.press("Enter")
        time.sleep(2)  # Đợi 2 giây sau khi nhấn Enter
        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)

        # # Tìm và click menu "Loại dữ liệu"
        # print("Đang chọn Loại dữ liệu...")
        # page_baocao.locator('div.search-criteria > div').filter(has_text='Loại dữ liệu').locator('select, button').first.click()
        # time.sleep(2)
        # print("✅ Đã click dropdown Loại dữ liệu")

        # #Nhấn phím UP 1 lần sau đó Enter
        # page_baocao.keyboard.press("ArrowUp")
        # time.sleep(1)  # Đợi 1 giây sau khi nhấn
        # page_baocao.keyboard.press("Enter")
        # time.sleep(2)  # Đợi 2 giây sau khi nhấn Enter

        # 1. Click "Báo cáo" button
        print("\nĐang click button 'Báo cáo'...")
        button1_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > button"
        page_baocao.locator(button1_selector).click()
        print("✅ Đã click button 'Báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(5)

        # 2. Click "Xuất Excel" button
        print("\nĐang click button 'Xuất Excel'...")
        button2_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > button"
        page_baocao.locator(button2_selector).click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel'.")

        # 3. Click download icon and save file
        print("\nĐang click icon download...")
        download_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(download_selector).click()
            print("✅ Đã click icon download.")

        download = download_info.value

        # Lưu file với tên cố định = tên hàm + .xlsx (ghi đè file cũ)
        filename = "download_KR7_report_NVKT.xlsx"
        save_path = os.path.join(download_dir, filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo KR7 chi tiết NVKT: {e}")
        import traceback
        traceback.print_exc()


def download_KR7_report_tong_hop(page_baocao):
    """
    Tải báo cáo KR7 từ trang baocao.hanoi.vnpt.vn

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo KR7 TỔNG HỢP ===")

    # In ngày tra cứu để debug
    current_date = datetime.now().strftime("%d/%m/%Y")
    print(f"📅 Ngày tra cứu báo cáo: {current_date}")

    # Truy cập trang báo cáo
    #report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521580&menu_id=521601'
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=521580' #mới
    print(f"🔗 URL báo cáo KR7 Tổng hợp: {report_url}")
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=60000)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=120000)
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

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)

        # #chọn menu Loại BC
        # # Click vào dropdown menu 'Loại BC'
        # print("Đang chọn Loại BC...")
        # page_baocao.locator('div.search-criteria > div').filter(has_text='Loại BC').locator('select, button').first.click()
        # time.sleep(2)
        # print("✅ Đã click dropdown Loại BC")
        # # Nhấn phím arrow down 2 lần
        # for _ in range(2):
        #     page_baocao.keyboard.press("ArrowDown")
        #     time.sleep(1)  # Đợi 1 giây giữa các lần nhấn

        # # Nhấn Enter
        # page_baocao.keyboard.press("Enter")
        # time.sleep(2)  # Đợi 2 giây sau khi nhấn Enter

        # # Tìm và click menu "Loại dữ liệu"
        # print("Đang chọn Loại dữ liệu...")
        # page_baocao.locator('div.search-criteria > div').filter(has_text='Loại dữ liệu').locator('select, button').first.click()
        # time.sleep(2)
        # print("✅ Đã click dropdown Loại dữ liệu")

        # #Nhấn phím UP 1 lần sau đó Enter
        # page_baocao.keyboard.press("ArrowUp")
        # time.sleep(1)  # Đợi 1 giây sau khi nhấn
        # page_baocao.keyboard.press("Enter")
        # time.sleep(2)  # Đợi 2 giây sau khi nhấn Enter

        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1)

        # 1. Click "Báo cáo" button
        print("\nĐang click button 'Báo cáo'...")
        button1_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-untouched.ng-valid.ng-dirty > div > div > div.box-detail > div.button-action > button"
        page_baocao.locator(button1_selector).click()
        print("✅ Đã click button 'Báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(5)

        # 2. Click "Xuất Excel" button
        print("\nĐang click button 'Xuất Excel'...")
        button2_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-untouched.ng-valid.ng-dirty > div > div > div.box-detail > div.button-action > div > button"
        page_baocao.locator(button2_selector).click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel'.")

        # 3. Click download icon and save file
        print("\nĐang click icon download...")
        download_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"

        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(download_selector).click()
            print("✅ Đã click icon download.")

        download = download_info.value

        # Lưu file với tên cố định = tên hàm + .xlsx (ghi đè file cũ)
        filename = "download_KR7_report_tong_hop.xlsx"
        save_path = os.path.join(download_dir, filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo KR7 Tổng hợp: {e}")
        import traceback
        traceback.print_exc()


def main():
    """
    Hàm main để test standalone - tải tất cả báo cáo KR
    """
    try:
        # Import login function
        from login import login_baocao_hanoi

        print("=== Bắt đầu test module KR_download ===")

        # Đăng nhập
        print("\n1. Đăng nhập vào hệ thống...")
        page_baocao, browser_baocao, playwright_baocao = login_baocao_hanoi()
        print("✅ Đăng nhập thành công!")

        # Tải các báo cáo KR6
        print("\n2. Tải báo cáo KR6...")
        download_KR6_report_NVKT(page_baocao)
        download_KR6_report_tong_hop(page_baocao)

        # Tải các báo cáo KR7
        print("\n3. Tải báo cáo KR7...")
        download_KR7_report_NVKT(page_baocao)
        download_KR7_report_tong_hop(page_baocao)

        print("\n✅ Hoàn thành tải tất cả báo cáo KR!")
        print("Các file đã được lưu vào thư mục: downloads/baocao_hanoi/")

        # Đóng browser
        print("\nĐang đóng trình duyệt...")
        browser_baocao.close()
        playwright_baocao.stop()
        print("✅ Đã đóng trình duyệt!")

    except Exception as e:
        print(f"\n❌ Có lỗi xảy ra: {str(e)}")
        import traceback
        traceback.print_exc()

    finally:
        # Đảm bảo đóng browser
        try:
            browser_baocao.close()
            playwright_baocao.stop()
        except:
            pass


if __name__ == "__main__":
    main()
