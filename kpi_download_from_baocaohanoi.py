# -*- coding: utf-8 -*-
"""
Module chứa các hàm download báo cáo chỉ tiêu từ hệ thống báo cáo VNPT Hà Nội.
"""

import os
import time
from config import Config
from login import login_baocao_hanoi


def c11_download_report_nvkt(page_baocao, report_month):
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
    5. Chọn option[3] trong dropdown loại báo cáo
    6. Chọn tháng báo cáo
    7. Click button "Báo cáo"
    8. Click button "Xuất Excel"
    9. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "=" * 80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO CHỈ TIÊU 1.1 TỪ BAOCAOHANOI - {report_month}")
        print("=" * 80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=522457&menu_id=522561"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/button"
        )
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        )
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div"
            "/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        )
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn option[3] trong dropdown loại báo cáo
        print("\n✓ Đang chọn option[3] trong dropdown loại báo cáo...")
        loai_baocao_select_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[4]"
            "/div/div/div/div/select"
        )
        loai_baocao_option3_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[4]"
            "/div/div/div/div/select/option[3]"
        )
        loai_baocao_select = page_baocao.locator(f"xpath={loai_baocao_select_xpath}")
        loai_baocao_select.wait_for(state="visible", timeout=10000)
        # Lấy giá trị của option[3] rồi select
        option3 = page_baocao.locator(f"xpath={loai_baocao_option3_xpath}")
        option3_value = option3.get_attribute("value")
        if option3_value:
            loai_baocao_select.select_option(value=option3_value)
        else:
            # Fallback: click trực tiếp vào option[3]
            option3.click()
        time.sleep(1)
        print("✅ Đã chọn option[3] trong dropdown loại báo cáo")

        # Bước 6: Chọn tháng báo cáo
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            month_select = page_baocao.locator("div.search-criteria > div:nth-child(2) select")
            month_select.wait_for(state="visible", timeout=10000)
            month_select.select_option(label=report_month)
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 7: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        )
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

        # Bước 8: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        )
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 9: Click icon download thứ 2 để tải file (2. Tất cả dữ liệu)
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        )
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục downloads
        download_dir = "KPI-DOWNLOAD"
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c11-nvktdb report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "=" * 80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.1 CHỈ TIÊU 1.1 (Q4/2025)")
        print("=" * 80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.1 CHỈ TIÊU 1.1 (Q4/2025): {e}")
        import traceback
        traceback.print_exc()
        return False


def c12_download_report_nvkt(page_baocao, report_month="Tháng 01/2026"):
    """
    Tải báo cáo C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại từ hệ thống báo cáo VNPT Hà Nội

    Args:
        page_baocao: Playwright page object đã đăng nhập
        report_month: Tháng báo cáo (ví dụ: "Tháng 12/2025", "Tháng 01/2026")

    Các bước:
    1. Truy cập URL báo cáo C1.2
    2. Click vào dropdown button để mở treeview
    3. Tìm input search và nhập "sơn tây"
    4. Chọn đơn vị TTVT Sơn Tây từ kết quả
    5. Chọn option[3] (NVKT quản lý địa bàn) trong dropdown loại báo cáo
    6. Chọn tháng báo cáo
    7. Click button "Báo cáo"
    8. Click button "Xuất Excel"
    9. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "=" * 80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại TỪ BAOCAOHANOI- {report_month}")
        print("=" * 80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=522562"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/button"
        )
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        )
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div"
            "/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        )
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn option[3] (NVKT quản lý địa bàn) trong dropdown loại báo cáo
        print("\n✓ Đang chọn option[3] (NVKT quản lý địa bàn) trong dropdown loại báo cáo...")
        loai_baocao_select_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[4]"
            "/div/div/div/div/select"
        )
        loai_baocao_option3_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[4]"
            "/div/div/div/div/select/option[3]"
        )
        loai_baocao_select = page_baocao.locator(f"xpath={loai_baocao_select_xpath}")
        loai_baocao_select.wait_for(state="visible", timeout=10000)
        option3 = page_baocao.locator(f"xpath={loai_baocao_option3_xpath}")
        option3_value = option3.get_attribute("value")
        if option3_value:
            loai_baocao_select.select_option(value=option3_value)
        else:
            option3.click()
        time.sleep(1)
        print("✅ Đã chọn option[3] trong dropdown loại báo cáo")

        # Bước 6: Chọn tháng báo cáo
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            month_select = page_baocao.locator("div.search-criteria > div:nth-child(2) select")
            month_select.wait_for(state="visible", timeout=10000)
            month_select.select_option(label=report_month)
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 7: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        )
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

        # Bước 8: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        )
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 9: Click icon download thứ 2 để tải file (2. Tất cả dữ liệu)
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        )
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục KPI-DOWNLOAD
        download_dir = "KPI-DOWNLOAD"
        os.makedirs(download_dir, exist_ok=True)

        # Lấy tên file gốc
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        # Lưu file với tên mới
        output_filename = "c12-nvktdb report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "=" * 80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại")
        print("=" * 80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.2 - Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại: {e}")
        import traceback
        traceback.print_exc()
        return False


def c13_download_report_nvkt(page_baocao, report_month="Tháng 01/2026"):
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
    5. Chọn option[3] (NVKT quản lý địa bàn) trong dropdown loại báo cáo
    6. Chọn tháng báo cáo
    7. Click button "Báo cáo"
    8. Click button "Xuất Excel"
    9. Click "2.Tất cả dữ liệu" để tải file
    """

    try:
        print("\n" + "=" * 80)
        print(f"BẮT ĐẦU TẢI BÁO CÁO C1.3 TỪ BAOCAOHANOI - {report_month}")
        print("=" * 80)

        # Bước 1: Truy cập URL
        url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=522600&menu_id=522640"
        print(f"\n✓ Đang truy cập: {url}")
        page_baocao.goto(url, wait_until="networkidle", timeout=60000)
        time.sleep(3)
        print("✅ Đã tải trang thành công")

        # Bước 2: Click vào dropdown button để mở treeview
        print("\n✓ Đang click vào dropdown button...")
        dropdown_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/button"
        )
        dropdown_btn = page_baocao.locator(f"xpath={dropdown_btn_xpath}")
        dropdown_btn.wait_for(state="visible", timeout=10000)
        dropdown_btn.click()
        time.sleep(1)
        print("✅ Đã mở dropdown treeview")

        # Bước 3: Tìm input search và nhập "sơn tây"
        print("\n✓ Đang tìm input search và nhập 'sơn tây'...")
        search_input_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        )
        search_input = page_baocao.locator(f"xpath={search_input_xpath}")
        search_input.wait_for(state="visible", timeout=10000)
        search_input.fill("sơn tây")
        time.sleep(2)
        print("✅ Đã nhập 'sơn tây' vào ô tìm kiếm")

        # Bước 4: Chọn đơn vị TTVT Sơn Tây từ kết quả
        print("\n✓ Đang chọn 'TTVT Sơn Tây'...")
        sontay_option_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]"
            "/div/div/div/div/div/div/ngx-dropdown-treeview-select"
            "/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div"
            "/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        )
        sontay_option = page_baocao.locator(f"xpath={sontay_option_xpath}")
        sontay_option.wait_for(state="visible", timeout=10000)
        sontay_option.click()
        time.sleep(1)
        print("✅ Đã chọn 'TTVT Sơn Tây'")

        # Bước 5: Chọn option[3] (NVKT quản lý địa bàn) trong dropdown loại báo cáo
        print("\n✓ Đang chọn option[3] (NVKT quản lý địa bàn) trong dropdown loại báo cáo...")
        loai_baocao_select_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[4]"
            "/div/div/div/div/select"
        )
        loai_baocao_option3_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[4]"
            "/div/div/div/div/select/option[3]"
        )
        loai_baocao_select = page_baocao.locator(f"xpath={loai_baocao_select_xpath}")
        loai_baocao_select.wait_for(state="visible", timeout=10000)
        option3 = page_baocao.locator(f"xpath={loai_baocao_option3_xpath}")
        option3_value = option3.get_attribute("value")
        if option3_value:
            loai_baocao_select.select_option(value=option3_value)
        else:
            option3.click()
        time.sleep(1)
        print("✅ Đã chọn option[3] trong dropdown loại báo cáo")

        # Bước 6: Chọn tháng báo cáo
        print(f"\n✓ Đang chọn tháng báo cáo: {report_month}...")
        try:
            month_select = page_baocao.locator("div.search-criteria > div:nth-child(2) select")
            month_select.wait_for(state="visible", timeout=10000)
            month_select.select_option(label=report_month)
            time.sleep(1)
            print(f"✅ Đã chọn tháng: {report_month}")
        except Exception as e:
            print(f"⚠️ Lỗi khi chọn tháng: {e}")

        # Bước 7: Click button "Báo cáo"
        print("\n✓ Đang click button 'Báo cáo'...")
        baocao_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        )
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

        # Bước 8: Click button "Xuất Excel"
        print("\n✓ Đang click button 'Xuất Excel'...")
        xuatexcel_btn_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        )
        xuatexcel_btn = page_baocao.locator(f"xpath={xuatexcel_btn_xpath}")
        xuatexcel_btn.wait_for(state="visible", timeout=10000)
        xuatexcel_btn.click()
        time.sleep(2)
        print("✅ Đã click button 'Xuất Excel', dropdown đã mở")

        # Bước 9: Click icon download thứ 2 để tải file (2. Tất cả dữ liệu)
        print("\n✓ Đang click icon download để tải file...")
        download_icon_xpath = (
            "/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div"
            "/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"
        )
        download_icon = page_baocao.locator(f"xpath={download_icon_xpath}")
        download_icon.wait_for(state="visible", timeout=10000)

        # Bắt đầu tải file
        with page_baocao.expect_download(timeout=60000) as download_info:
            download_icon.click()
            time.sleep(2)

        download = download_info.value

        # Lưu file vào thư mục KPI-DOWNLOAD
        download_dir = "KPI-DOWNLOAD"
        os.makedirs(download_dir, exist_ok=True)

        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        output_filename = "c13-nvktdb report.xlsx"
        output_path = os.path.join(download_dir, output_filename)
        download.save_as(output_path)

        print(f"✅ Đã tải file về: {output_path}")

        print("\n" + "=" * 80)
        print("✅ HOÀN THÀNH TẢI BÁO CÁO C1.3")
        print("=" * 80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi tải báo cáo C1.3: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """
    Hàm main để chạy standalone (test độc lập).
    Đăng nhập vào hệ thống báo cáo rồi gọi c11, c12, c13_download_report_nvkt.
    """
    page_baocao = None
    browser_baocao = None
    playwright_baocao = None
    report_month = "Tháng 03/2026"  # Thay đổi tháng tại đây nếu cần
    try:
        # Bước 1: Đăng nhập
        print("=== KHỞI ĐỘNG MODULE KPI-DOWNLOAD ===")
        page_baocao, browser_baocao, playwright_baocao = login_baocao_hanoi()

        # Bước 2: Tải báo cáo C1.1
        #report_month = "Tháng 03/2026"  # Thay đổi tháng tại đây nếu cần
        result_c11 = c11_download_report_nvkt(page_baocao, report_month)
        if result_c11:
            print("\n✅ C1.1 - Tải thành công!")
        else:
            print("\n❌ C1.1 - Có lỗi xảy ra.")

        # Bước 3: Tải báo cáo C1.2
        result_c12 = c12_download_report_nvkt(page_baocao, report_month)
        if result_c12:
            print("\n✅ C1.2 - Tải thành công!")
        else:
            print("\n❌ C1.2 - Có lỗi xảy ra.")

        # Bước 4: Tải báo cáo C1.3
        result_c13 = c13_download_report_nvkt(page_baocao, report_month)
        if result_c13:
            print("\n✅ C1.3 - Tải thành công!")
        else:
            print("\n❌ C1.3 - Có lỗi xảy ra.")

        if result_c11 and result_c12 and result_c13:
            print("\n✅ Tất cả báo cáo đã được tải thành công!")
        else:
            print("\n⚠️ Một số báo cáo có lỗi, vui lòng kiểm tra lại.")

    except Exception as e:
        print(f"\n❌ Lỗi nghiêm trọng: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Đóng trình duyệt
        print("\n=== Đang đóng trình duyệt... ===")
        try:
            if browser_baocao:
                browser_baocao.close()
            if playwright_baocao:
                playwright_baocao.stop()
            print("✅ Đã đóng trình duyệt.")
        except Exception as e:
            print(f"⚠️ Lỗi khi đóng trình duyệt: {e}")


if __name__ == "__main__":
    main()
