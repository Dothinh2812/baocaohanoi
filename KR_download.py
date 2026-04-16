# -*- coding: utf-8 -*-
"""
Module chứa các hàm download báo cáo KR6 và KR7
Có thể chạy standalone để test: python KR_download.py
"""
import time
import os
from datetime import datetime


def download_GHTT_report_HNI(page_baocao):
    """
    Tải báo cáo KR6 từ trang baocao.hanoi.vnpt.vn
    > 10.BSC/OKR > 10.1 Báo cáo BSC > 10.1.7 Chỉ tiêu công tác gia hạn TTTC > 10.1.7.4 Duy trì thuê bao BRCĐ hiện hữu gia hạn TTTC
    https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo KR6 chi tiết NVKT===")

    # In ngày tra cứu để debug
    current_date = datetime.now().strftime("%d/%m/%Y")
    print(f"📅 Ngày tra cứu báo cáo: {current_date}")

    # Truy cập trang báo cáo
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238'
    print(f"🔗 URL báo cáo : {report_url}")
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=60000)

    # Đợi trang load xong
    print("Đang đợi trang load...")
    page_baocao.wait_for_load_state("networkidle", timeout=120000)
    time.sleep(3)

    try:
        # 1. Click "Báo cáo" button
        print("\nĐang click button 'Báo cáo'...")
        baocao_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        page_baocao.locator(f"xpath={baocao_xpath}").click()
        print("✅ Đã click button 'Báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(5)

        # 2. Click "Xuất Excel" button
        print("\nĐang click button 'Xuất Excel'...")
        xuat_excel_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        page_baocao.locator(f"xpath={xuat_excel_xpath}").click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel'.")

        # 3. Click download icon and save file
        print("\nĐang click icon download...")
        download_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"

        download_dir = "GHTT"
        os.makedirs(download_dir, exist_ok=True)

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(f"xpath={download_xpath}").click()
            print("✅ Đã click icon download.")

        download = download_info.value

        # Lưu file với tên cố định = tên hàm + .xlsx (ghi đè file cũ)
        filename = "tong_hop_ghtt_hni.xlsx"
        save_path = os.path.join(download_dir, filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo KR6 chi tiết NVKT: {e}")
        import traceback
        traceback.print_exc()


def download_GHTT_report_Son_Tay(page_baocao):
    """
    Tải báo cáo KR6 từ trang baocao.hanoi.vnpt.vn
    https://baocao.hanoi.vnpt.vn/report/report-info?id=523160

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo KR6 TỔNG HỢP===")

    # In ngày tra cứu để debug
    current_date = datetime.now().strftime("%d/%m/%Y")
    print(f"📅 Ngày tra cứu báo cáo: {current_date}")

    # Truy cập trang báo cáo
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238'
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

        # 1. Click "Báo cáo" button
        print("\nĐang click button 'Báo cáo'...")
        baocao_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        page_baocao.locator(f"xpath={baocao_xpath}").click()
        print("✅ Đã click button 'Báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(5)

        # 2. Click "Xuất Excel" button
        print("\nĐang click button 'Xuất Excel'...")
        xuat_excel_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        page_baocao.locator(f"xpath={xuat_excel_xpath}").click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel'.")

        # 3. Click download icon and save file
        print("\nĐang click icon download...")
        download_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"

        download_dir = "GHTT"
        os.makedirs(download_dir, exist_ok=True)

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(f"xpath={download_xpath}").click()
            print("✅ Đã click icon download.")

        download = download_info.value

        # Lưu file với tên cố định = tên hàm + .xlsx (ghi đè file cũ)
        filename = "tong_hop_ghtt_sontay.xlsx"
        save_path = os.path.join(download_dir, filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo KR6 Tổng hợp: {e}")
        import traceback
        traceback.print_exc()


def download_GHTT_report_nvktdb(page_baocao):
    """
    Tải báo cáo GHTT chi tiết NVKT DB từ trang baocao.hanoi.vnpt.vn
    https://baocao.hanoi.vnpt.vn/report/report-info?id=523160

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo GHTT chi tiết NVKT DB===")

    # In ngày tra cứu để debug
    current_date = datetime.now().strftime("%d/%m/%Y")
    print(f"📅 Ngày tra cứu báo cáo: {current_date}")

    # Truy cập trang báo cáo
    report_url = 'https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238'
    print(f"🔗 URL báo cáo GHTT NVKT DB: {report_url}")
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

        # Bước 2: Click vào trường "Loại", sau đó dùng phím mũi tên xuống và Enter
        print("Đang chọn loại báo cáo NVKT DB bằng bàn phím...")
        select_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/select"
        page_baocao.locator(f"xpath={select_xpath}").click()
        time.sleep(1)
        page_baocao.keyboard.press("ArrowDown")
        time.sleep(1)
        page_baocao.keyboard.press("Enter")
        time.sleep(2)
        print("✅ Đã chọn loại báo cáo NVKT DB")

        # 1. Click "Báo cáo" button
        print("\nĐang click button 'Báo cáo'...")
        baocao_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        page_baocao.locator(f"xpath={baocao_xpath}").click()
        print("✅ Đã click button 'Báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(5)

        # 2. Click "Xuất Excel" button
        print("\nĐang click button 'Xuất Excel'...")
        xuat_excel_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        page_baocao.locator(f"xpath={xuat_excel_xpath}").click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel'.")

        # 3. Click download icon and save file
        print("\nĐang click icon download...")
        download_xpath = "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"

        download_dir = "GHTT"
        os.makedirs(download_dir, exist_ok=True)

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(f"xpath={download_xpath}").click()
            print("✅ Đã click icon download.")

        download = download_info.value

        # Lưu file với tên cố định = tên hàm + .xlsx (ghi đè file cũ)
        filename = "tong_hop_ghtt_nvktdb.xlsx"
        save_path = os.path.join(download_dir, filename)
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo GHTT NVKT DB: {e}")
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

        # 3. Click vào option "2.Tất cả dữ liệu" trong dropdown để tải file
        print("\nĐang click '2.Tất cả dữ liệu' để tải file...")

        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        time.sleep(1)

        # Tìm element <i> có class "dropdown-item" và text "2.Tất cả dữ liệu"
        try:
            # Cách 1: Tìm theo text
            download_option = page_baocao.locator("i.dropdown-item").filter(has_text="2.Tất cả dữ liệu").first
            download_option.wait_for(state="visible", timeout=30000)
            print("✅ Đã tìm thấy option '2.Tất cả dữ liệu'")
        except:
            # Cách 2: Sử dụng selector CSS cụ thể
            print("Đang thử selector khác...")
            download_selector = "body > app-root > app-layout > app-vertical > div.body > div.main-content > div > app-report-info-list > div > div:nth-child(1) > div.ng-valid.ng-touched.ng-dirty > div > div > div.box-detail > div.button-action > div > div > i:nth-child(2)"
            download_option = page_baocao.locator(download_selector)
            download_option.wait_for(state="visible", timeout=30000)
            print("✅ Đã tìm thấy option download (selector CSS)")

        print("Đang chờ và tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            download_option.click()
            print("✅ Đã click vào '2.Tất cả dữ liệu'")

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
        # download_GHTT_report_HNI(page_baocao)
        # download_GHTT_report_Son_Tay(page_baocao)
        download_GHTT_report_nvktdb(page_baocao)

        # # Tải các báo cáo KR7
        # print("\n3. Tải báo cáo KR7...")
        # download_KR7_report_NVKT(page_baocao)
        # download_KR7_report_tong_hop(page_baocao)

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
