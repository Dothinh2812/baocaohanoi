# -*- coding: utf-8 -*-
"""
Module chứa các hàm download báo cáo thực tăng PTTB và MyTV
"""
import time
import os
from datetime import datetime
from urllib.parse import quote
from config import Config


def download_report_pttb_ngung_psc(page_baocao):
    """
    Tải báo cáo PTTB Ngưng PSC từ trang đã chỉ định

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo PTTB Ngưng PSC ===")

    # Lấy ngày hiện tại và format theo định dạng dd/mm/yyyy
    current_date = datetime.now().strftime("%d/%m/%Y")
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

    # Tìm và click button "2.Tất cả dữ liệu"
    print("Đang tìm và click '2.Tất cả dữ liệu'...")
    try:
        all_data_button = page_baocao.get_by_text("2.Tất cả dữ liệu", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
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

    # Lấy ngày hiện tại và format theo định dạng dd/mm/yyyy
    current_date = datetime.now().strftime("%d/%m/%Y")
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

    # Tìm và click button "2.Tất cả dữ liệu"
    print("Đang tìm và click '2.Tất cả dữ liệu'...")
    try:
        all_data_button = page_baocao.get_by_text("2.Tất cả dữ liệu", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
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

    # Lấy ngày hiện tại và format theo định dạng dd/mm/yyyy
    current_date = datetime.now().strftime("%d/%m/%Y")
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

    # Tìm và click button "2.Tất cả dữ liệu"
    print("Đang tìm và click '2.Tất cả dữ liệu'...")
    try:
        all_data_button = page_baocao.get_by_text("2.Tất cả dữ liệu", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
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

    # Lấy ngày hiện tại và format theo định dạng dd/mm/yyyy
    current_date = datetime.now().strftime("%d/%m/%Y")
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

    # Tìm và click button "2.Tất cả dữ liệu"
    print("Đang tìm và click '2.Tất cả dữ liệu'...")
    try:
        all_data_button = page_baocao.get_by_text("2.Tất cả dữ liệu", exact=False)
        all_data_button.wait_for(state="visible", timeout=15000)

        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
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
