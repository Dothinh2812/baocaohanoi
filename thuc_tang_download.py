# -*- coding: utf-8 -*-
"""
Module chứa các hàm download báo cáo thực tăng PTTB và MyTV
"""
import time
import os
from datetime import datetime, timedelta
from urllib.parse import quote
from config import Config


def download_report_pttb_ngung_psc(page_baocao):
    """
    Tải báo cáo PTTB Ngưng PSC từ trang đã chỉ định

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo PTTB Ngưng PSC ===")

    # Lấy ngày T-1 (hôm qua) vì số liệu chỉ có đến ngày T-1
    yesterday = datetime.now() - timedelta(days=1)
    current_date = yesterday.strftime("%d/%m/%Y")
    # Encode ngày cho URL (/ -> %2F)
    encoded_date = quote(current_date, safe='')

    print(f"Ngày báo cáo: {current_date}")

    # Truy cập trang báo cáo với ngày hiện tại
    report_url = Config.get_report_data_url('kr6_nvkt', encoded_date)
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.NETWORK_IDLE_TIMEOUT)

    # Đợi trang load xong dữ liệu
    print("Đang đợi dữ liệu load...")
    page_baocao.wait_for_load_state("networkidle", timeout=500000)
    time.sleep(5)  # Đợi thêm để đảm bảo dữ liệu đã load đầy đủ

    # Tìm và click button "Xuất Excel"
    print("Đang tìm button 'Xuất Excel'...")

    # Thử tìm button có text "Xuất Excel" (có thể có dấu ... hoặc không)
    try:
        # Thử tìm với text chính xác
        export_button = page_baocao.get_by_text("Xuất Excel", exact=False)
        export_button.wait_for(state="visible", timeout=500000)
        print("Đã tìm thấy button 'Xuất Excel', đang click...")
        export_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"⚠️ Không tìm thấy button với text 'Xuất Excel': {e}")
        print("Đang thử tìm bằng selector khác...")
        # Có thể thêm selector khác nếu cần

    # Tìm và click button "3.Database"
    print("Đang tìm và click '3.Database'...")
    try:
        all_data_button = page_baocao.get_by_text("3.Database", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = "PTTB-PSC"
        os.makedirs(download_dir, exist_ok=True)

        # Tạo tên file theo format: ngung_psc_DDMMYYYY.xlsx
        date_str = datetime.now().strftime("%d%m%Y")
        custom_filename = f"ngung_psc_{date_str}.xlsx"
        save_path = os.path.join(download_dir, custom_filename)

        print("Đang tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            all_data_button.click()

        download = download_info.value
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải file: {e}")


def download_report_pttb_hoan_cong(page_baocao):
    """
    Tải báo cáo PTTB Hoàn công từ trang đã chỉ định

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo PTTB Hoàn công ===")

    # Lấy ngày T-1 (hôm qua) vì số liệu chỉ có đến ngày T-1
    yesterday = datetime.now() - timedelta(days=1)
    current_date = yesterday.strftime("%d/%m/%Y")
    # Encode ngày cho URL (/ -> %2F)
    encoded_date = quote(current_date, safe='')

    print(f"Ngày báo cáo: {current_date}")

    # Truy cập trang báo cáo với ngày hiện tại
    report_url = Config.get_report_data_url('kr6_tonghop', encoded_date)
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.NETWORK_IDLE_TIMEOUT)

    # Đợi trang load xong dữ liệu
    print("Đang đợi dữ liệu load...")
    page_baocao.wait_for_load_state("networkidle", timeout=Config.NETWORK_IDLE_TIMEOUT)
    time.sleep(5)  # Đợi thêm để đảm bảo dữ liệu đã load đầy đủ

    # Tìm và click button "Xuất Excel"
    print("Đang tìm button 'Xuất Excel'...")

    # Thử tìm button có text "Xuất Excel" (có thể có dấu ... hoặc không)
    try:
        # Thử tìm với text chính xác
        export_button = page_baocao.get_by_text("Xuất Excel", exact=False)
        export_button.wait_for(state="visible", timeout=Config.NETWORK_IDLE_TIMEOUT)
        print("Đã tìm thấy button 'Xuất Excel', đang click...")
        export_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"⚠️ Không tìm thấy button với text 'Xuất Excel': {e}")
        print("Đang thử tìm bằng selector khác...")
        # Có thể thêm selector khác nếu cần

    # Tìm và click button "3.Database"
    print("Đang tìm và click '3.Database'...")
    try:
        all_data_button = page_baocao.get_by_text("3.Database", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = "PTTB-PSC"
        os.makedirs(download_dir, exist_ok=True)

        # Tạo tên file theo format: hoan_cong_DDMMYYYY.xlsx
        date_str = datetime.now().strftime("%d%m%Y")
        custom_filename = f"hoan_cong_{date_str}.xlsx"
        save_path = os.path.join(download_dir, custom_filename)

        print("Đang tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            all_data_button.click()

        download = download_info.value
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải file: {e}")


def download_report_mytv_hoan_cong(page_baocao):
    """
    Tải báo cáo MyTV Hoàn công từ trang đã chỉ định

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo MyTV Hoàn công ===")

    # Lấy ngày T-1 (hôm qua) vì số liệu chỉ có đến ngày T-1
    yesterday = datetime.now() - timedelta(days=1)
    current_date = yesterday.strftime("%d/%m/%Y")
    # Encode ngày cho URL (/ -> %2F)
    encoded_date = quote(current_date, safe='')

    print(f"Ngày báo cáo: {current_date}")

    # Truy cập trang báo cáo với ngày hiện tại
    # vdvvt_id=8 là MyTV
    report_url = Config.get_report_data_url('kr7_tonghop', encoded_date)
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.NETWORK_IDLE_TIMEOUT)

    # Đợi trang load xong dữ liệu
    print("Đang đợi dữ liệu load...")
    page_baocao.wait_for_load_state("networkidle", timeout=500000)
    time.sleep(5)  # Đợi thêm để đảm bảo dữ liệu đã load đầy đủ

    # Tìm và click button "Xuất Excel"
    print("Đang tìm button 'Xuất Excel'...")

    # Thử tìm button có text "Xuất Excel" (có thể có dấu ... hoặc không)
    try:
        # Thử tìm với text chính xác
        export_button = page_baocao.get_by_text("Xuất Excel", exact=False)
        export_button.wait_for(state="visible", timeout=500000)
        print("Đã tìm thấy button 'Xuất Excel', đang click...")
        export_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"⚠️ Không tìm thấy button với text 'Xuất Excel': {e}")
        print("Đang thử tìm bằng selector khác...")
        # Có thể thêm selector khác nếu cần

    # Tìm và click button "3.Database"
    print("Đang tìm và click '3.Database'...")
    try:
        all_data_button = page_baocao.get_by_text("3.Database", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = "PTTB-PSC"
        os.makedirs(download_dir, exist_ok=True)

        # Tạo tên file theo format: mytv_hoan_cong_DDMMYYYY.xlsx
        date_str = datetime.now().strftime("%d%m%Y")
        custom_filename = f"mytv_hoan_cong_{date_str}.xlsx"
        save_path = os.path.join(download_dir, custom_filename)

        print("Đang tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            all_data_button.click()

        download = download_info.value
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải file: {e}")


def download_report_mytv_ngung_psc(page_baocao):
    """
    Tải báo cáo MyTV Ngưng PSC từ trang đã chỉ định

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo MyTV Ngưng PSC ===")

    # Lấy ngày T-1 (hôm qua) vì số liệu chỉ có đến ngày T-1
    yesterday = datetime.now() - timedelta(days=1)
    current_date = yesterday.strftime("%d/%m/%Y")
    # Encode ngày cho URL (/ -> %2F)
    encoded_date = quote(current_date, safe='')

    print(f"Ngày báo cáo: {current_date}")

    # Truy cập trang báo cáo với ngày hiện tại
    # vdvvt_id=8 là MyTV
    report_url = Config.get_report_data_url('kr7_nvkt', encoded_date)
    print(f"Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.NETWORK_IDLE_TIMEOUT)

    # Đợi trang load xong dữ liệu
    print("Đang đợi dữ liệu load...")
    page_baocao.wait_for_load_state("networkidle", timeout=Config.NETWORK_IDLE_TIMEOUT)
    time.sleep(5)  # Đợi thêm để đảm bảo dữ liệu đã load đầy đủ

    # Tìm và click button "Xuất Excel"
    print("Đang tìm button 'Xuất Excel'...")

    # Thử tìm button có text "Xuất Excel" (có thể có dấu ... hoặc không)
    try:
        # Thử tìm với text chính xác
        export_button = page_baocao.get_by_text("Xuất Excel", exact=False)
        export_button.wait_for(state="visible", timeout=Config.NETWORK_IDLE_TIMEOUT)
        print("Đã tìm thấy button 'Xuất Excel', đang click...")
        export_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"⚠️ Không tìm thấy button với text 'Xuất Excel': {e}")
        print("Đang thử tìm bằng selector khác...")
        # Có thể thêm selector khác nếu cần

    # Tìm và click button "3.Database"
    print("Đang tìm và click '3.Database'...")
    try:
        all_data_button = page_baocao.get_by_text("3.Database", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = "PTTB-PSC"
        os.makedirs(download_dir, exist_ok=True)

        # Tạo tên file theo format: mytv_ngung_psc_DDMMYYYY.xlsx
        date_str = datetime.now().strftime("%d%m%Y")
        custom_filename = f"mytv_ngung_psc_{date_str}.xlsx"
        save_path = os.path.join(download_dir, custom_filename)

        print("Đang tải file...")
        with page_baocao.expect_download(timeout=300000) as download_info:
            all_data_button.click()

        download = download_info.value
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải file: {e}")


def download_report_ngung_psc_fiber_thang_t_1_son_tay(page_baocao):
    """
    Tải báo cáo Sơn Tây từ trang đã chỉ định

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo Sơn Tây ===")
    url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=263889&menu_id=276187"
    print(f"Đang truy cập: {url}")
    page_baocao.goto(url, timeout=Config.NETWORK_IDLE_TIMEOUT)

    # Đợi trang load
    page_baocao.wait_for_load_state("networkidle", timeout=Config.NETWORK_IDLE_TIMEOUT)
    time.sleep(5)

    try:
        # Tìm và click dropdown treeview
        print("Đang mở dropdown...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button").click()
        time.sleep(2)

        # Type "ttvt sơn tây" vào input
        print("Đang nhập 'ttvt sơn tây'...")
        input_selector = "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        page_baocao.locator(input_selector).fill("ttvt sơn tây")
        time.sleep(2)

        # Chọn "ttvt sơn tây"
        print("Đang chọn 'ttvt sơn tây'...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span").click()
        time.sleep(2)

        # Click "Xem báo cáo"
        print("Đang click 'Xem báo cáo'...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button").click()
        time.sleep(5)

        # Click dropdown để xuất file
        print("Đang click dropdown xuất file...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button").click()
        time.sleep(2)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = "PTTB-PSC"
        os.makedirs(download_dir, exist_ok=True)

        # Tải file
        print("Đang tải file...")
        date_str = datetime.now().strftime("%d%m%Y")
        save_path = os.path.join(download_dir, f"ngung_psc_fiber_thang_t-1_sontay_{date_str}.xlsx")

        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]").click()

        download = download_info.value
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo Sơn Tây: {e}")


def download_report_ngung_psc_mytv_thang_t_1_son_tay(page_baocao):
    """
    Tải báo cáo MyTV Sơn Tây (Ngưng PSC Tháng T-1) từ trang đã chỉ định

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo MyTV Sơn Tây ===")
    url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=263889&menu_id=276187"
    print(f"Đang truy cập: {url}")
    page_baocao.goto(url, timeout=Config.NETWORK_IDLE_TIMEOUT)

    # Đợi trang load
    page_baocao.wait_for_load_state("networkidle", timeout=Config.NETWORK_IDLE_TIMEOUT)
    time.sleep(5)

    try:
        # Tìm và click dropdown treeview (đơn vị)
        print("Đang mở dropdown đơn vị...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button").click()
        time.sleep(2)

        # Type "ttvt sơn tây" vào input
        print("Đang nhập 'ttvt sơn tây'...")
        input_selector = "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        page_baocao.locator(input_selector).fill("ttvt sơn tây")
        time.sleep(2)

        # Chọn "ttvt sơn tây"
        print("Đang chọn 'ttvt sơn tây'...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span").click()
        time.sleep(2)

        # Chọn loại dịch vụ MyTV
        print("Đang mở dropdown loại dịch vụ...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button").click()
        time.sleep(2)

        print("Đang chọn MyTV...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item[2]/div/div[1]/span").click()
        time.sleep(2)

        # Click "Xem báo cáo"
        print("Đang click 'Xem báo cáo'...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button").click()
        time.sleep(5)

        # Click dropdown để xuất file
        print("Đang click dropdown xuất file...")
        page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button").click()
        time.sleep(2)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = "PTTB-PSC"
        os.makedirs(download_dir, exist_ok=True)

        # Tải file
        print("Đang tải file...")
        date_str = datetime.now().strftime("%d%m%Y")
        save_path = os.path.join(download_dir, f"ngung_psc_mytv_thang_t-1_sontay_{date_str}.xlsx")

        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator("xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]").click()

        download = download_info.value
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo MyTV Sơn Tây: {e}")


def main():
    """
    Hàm main để chạy standalone - tải tất cả báo cáo thực tăng
    """
    try:
        # Import login function
        from login import login_baocao_hanoi

        print("=== Bắt đầu tải báo cáo Thực tăng ===")

        # Đăng nhập
        print("\n1. Đăng nhập vào hệ thống...")
        page_baocao, browser_baocao, playwright_baocao = login_baocao_hanoi()
        print("✅ Đăng nhập thành công!")

        # Tải các báo cáo
        print("\n2. Tải báo cáo PTTB Ngưng PSC...")
        download_report_pttb_ngung_psc(page_baocao)

        print("\n3. Tải báo cáo PTTB Hoàn công...")
        download_report_pttb_hoan_cong(page_baocao)

        print("\n4. Tải báo cáo MyTV Hoàn công...")
        download_report_mytv_hoan_cong(page_baocao)

        print("\n5. Tải báo cáo MyTV Ngưng PSC...")
        download_report_mytv_ngung_psc(page_baocao)

        print("\n6. Tải báo cáo Fiber Sơn Tây...")
        download_report_ngung_psc_fiber_thang_t_1_son_tay(page_baocao)

        print("\n7. Tải báo cáo MyTV Sơn Tây...")
        download_report_ngung_psc_mytv_thang_t_1_son_tay(page_baocao)

        print("\n✅ Hoàn thành tải tất cả báo cáo Thực tăng!")
        print("Các file đã được lưu vào thư mục: PTTB-PSC/")

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
