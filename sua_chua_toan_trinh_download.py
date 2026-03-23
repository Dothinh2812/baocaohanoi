# -*- coding: utf-8 -*-
"""
Module chứa hàm download báo cáo sửa chữa toàn trình chi tiết
Sử dụng CSS selector chính xác từ DOM
"""
import time
import os
from datetime import datetime, timedelta
from config import Config

# Prefix chung cho các selector trong form báo cáo
_SEARCH = "div.box-detail > div.search-criteria"
_ACTION = "div.box-detail > div.button-action"


def download_report_sua_chua_toan_trinh(page_baocao, report_date=None):
    """
    Tải báo cáo sửa chữa toàn trình chi tiết

    Args:
        page_baocao: Đối tượng page đã đăng nhập
        report_date: Ngày báo cáo (datetime object). Mặc định là ngày hôm qua.
    """
    print("\n" + "=" * 80)
    print("=== Bắt đầu tải báo cáo Sửa chữa toàn trình chi tiết ===")
    print("=" * 80)

    if report_date is None:
        report_date = datetime.now() - timedelta(days=1)

    date_str = report_date.strftime("%d%m%Y")
    date_input = report_date.strftime("%d/%m/%Y")
    print(f"Ngày báo cáo: {date_input}")

    download_dir = os.path.join("downloads", "baocao_hanoi")
    os.makedirs(download_dir, exist_ok=True)

    # ======================================================================
    # Bước 1: Truy cập trang báo cáo
    # ======================================================================
    report_url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194"
    print(f"\nBước 1: Đang truy cập: {report_url}")
    page_baocao.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)
    page_baocao.wait_for_load_state("networkidle", timeout=Config.DOWNLOAD_TIMEOUT)
    time.sleep(3)
    print("Đã load trang báo cáo")

    # ======================================================================
    # Bước 2: Click dropdown đơn vị
    # ======================================================================
    print("\nBước 2: Đang mở dropdown đơn vị...")
    try:
        dropdown_donvi = page_baocao.locator(
            f"{_SEARCH} > div:nth-child(1) ngx-dropdown-treeview-select > ngx-dropdown-treeview > div > button"
        )
        dropdown_donvi.wait_for(state="visible", timeout=30000)
        dropdown_donvi.click()
        time.sleep(2)
        print("Đã mở dropdown đơn vị")
    except Exception as e:
        print(f"Lỗi mở dropdown đơn vị: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 3: Điền "sơn tây" vào ô tìm kiếm
    # ======================================================================
    print("\nBước 3: Đang tìm kiếm 'sơn tây'...")
    try:
        search_box = page_baocao.locator(
            f"{_SEARCH} > div:nth-child(1) ngx-dropdown-treeview-select input[type='text']"
        )
        search_box.wait_for(state="visible", timeout=10000)
        search_box.fill("sơn tây")
        time.sleep(2)
        print("Đã điền 'sơn tây'")
    except Exception as e:
        print(f"Lỗi tìm kiếm: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 4: Click chọn TTVT Sơn Tây
    # ======================================================================
    print("\nBước 4: Đang chọn TTVT Sơn Tây...")
    try:
        son_tay_item = page_baocao.locator(
            f"{_SEARCH} > div:nth-child(1) ngx-dropdown-treeview-select ngx-treeview > div:nth-child(2) > div > ngx-treeview-item > div > div:nth-child(2) > ngx-treeview-item > div > div > span"
        )
        son_tay_item.wait_for(state="visible", timeout=10000)
        son_tay_item.click()
        time.sleep(2)
        print("Đã chọn TTVT Sơn Tây")
    except Exception as e:
        print(f"Lỗi chọn TTVT Sơn Tây: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 5: Điền ngày bắt đầu (mat-input-0)
    # ======================================================================
    print(f"\nBước 5: Đang điền ngày bắt đầu: {date_input}...")
    try:
        input_from = page_baocao.locator("#mat-input-0")
        input_from.wait_for(state="visible", timeout=10000)
        input_from.click()
        input_from.fill("")
        time.sleep(0.5)
        input_from.fill(date_input)
        time.sleep(1)
        print(f"Đã điền ngày bắt đầu: {date_input}")
    except Exception as e:
        print(f"Lỗi điền ngày bắt đầu: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 6: Điền ngày kết thúc (mat-input-1)
    # ======================================================================
    print(f"\nBước 6: Đang điền ngày kết thúc: {date_input}...")
    try:
        input_to = page_baocao.locator("#mat-input-1")
        input_to.wait_for(state="visible", timeout=10000)
        input_to.click()
        input_to.fill("")
        time.sleep(0.5)
        input_to.fill(date_input)
        time.sleep(1)
        print(f"Đã điền ngày kết thúc: {date_input}")
    except Exception as e:
        print(f"Lỗi điền ngày kết thúc: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 7: Click dropdown dịch vụ (div:nth-child(4))
    # ======================================================================
    print("\nBước 7: Đang mở dropdown dịch vụ...")
    try:
        dropdown_dichvu = page_baocao.locator(
            f"{_SEARCH} > div:nth-child(4) ngx-dropdown-treeview > div > button"
        )
        dropdown_dichvu.wait_for(state="visible", timeout=10000)
        dropdown_dichvu.click()
        time.sleep(1)
        print("Đã mở dropdown dịch vụ")
    except Exception as e:
        print(f"Lỗi mở dropdown dịch vụ: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 8: Tick chọn "Tất cả"
    # ======================================================================
    print("\nBước 8: Đang chọn 'Tất cả'...")
    try:
        tat_ca_checkbox = page_baocao.locator(
            f"{_SEARCH} > div:nth-child(4) ngx-dropdown-treeview ngx-treeview > div.treeview-header > div:nth-child(2) > div.row.row-all.ng-star-inserted > div > div > input"
        )
        tat_ca_checkbox.wait_for(state="visible", timeout=10000)
        tat_ca_checkbox.click()
        time.sleep(1)
        print("Đã chọn 'Tất cả'")
    except Exception as e:
        print(f"Lỗi chọn 'Tất cả': {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 9: Chọn option thứ 2 trong select (div:nth-child(6))
    # ======================================================================
    print("\nBước 9: Đang chọn giá trị select...")
    try:
        select_el = page_baocao.locator(
            f"{_SEARCH} > div:nth-child(6) select"
        )
        select_el.wait_for(state="visible", timeout=10000)
        # Chọn option:nth-child(2) - tức option thứ 2
        select_el.select_option(index=1)
        time.sleep(1)
        print("Đã chọn option thứ 2")
    except Exception as e:
        print(f"Lỗi chọn select: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 10: Click button "Báo cáo"
    # ======================================================================
    print("\nBước 10: Đang click 'Báo cáo'...")
    try:
        btn_baocao = page_baocao.locator(f"{_ACTION} > button")
        btn_baocao.wait_for(state="visible", timeout=30000)
        btn_baocao.click()
        print("Đã click 'Báo cáo'")

        page_baocao.wait_for_load_state("networkidle", timeout=Config.NETWORK_IDLE_TIMEOUT)
        time.sleep(2)

        # Đợi loading overlay biến mất
        try:
            page_baocao.wait_for_selector("ngx-loading .backdrop", state="hidden", timeout=120000)
        except Exception:
            try:
                page_baocao.wait_for_selector(".backdrop.full-screen", state="hidden", timeout=10000)
            except Exception:
                pass

        time.sleep(2)
        print("Dữ liệu đã load xong")
    except Exception as e:
        print(f"Lỗi khi load dữ liệu: {e}")
        import traceback
        traceback.print_exc()

    # ======================================================================
    # Bước 11: Click button "Xuất Excel"
    # ======================================================================
    print("\nBước 11: Đang click 'Xuất Excel'...")
    save_path = None
    try:
        time.sleep(2)
        btn_xuat_excel = page_baocao.locator(f"{_ACTION} > div > button")
        btn_xuat_excel.wait_for(state="visible", timeout=30000)
        btn_xuat_excel.click()
        time.sleep(2)
        print("Đã mở dropdown 'Xuất Excel'")

        # ==================================================================
        # Bước 12: Click icon tải tất cả dữ liệu (i:nth-child(2))
        # ==================================================================
        print("\nBước 12: Đang click icon tải tất cả dữ liệu...")
        download_icon = page_baocao.locator(f"{_ACTION} > div > div > i:nth-child(2)")
        download_icon.wait_for(state="visible", timeout=30000)

        with page_baocao.expect_download(timeout=300000) as download_info:
            download_icon.click()

        download = download_info.value
        original_filename = download.suggested_filename
        print(f"Tên file gốc: {original_filename}")

        file_extension = os.path.splitext(original_filename)[1] or ".xlsx"
        custom_filename = f"sua_chua_toan_trinh_{date_str}{file_extension}"
        save_path = os.path.join(download_dir, custom_filename)

        download.save_as(save_path)
        print(f"Đã tải file về: {save_path}")

    except Exception as e:
        print(f"Lỗi khi tải file: {e}")
        import traceback
        traceback.print_exc()
        save_path = None

    print("\n" + "=" * 80)
    print("=== Kết thúc tải báo cáo Sửa chữa toàn trình chi tiết ===")
    print("=" * 80)

    return save_path


if __name__ == "__main__":
    from login import login_baocao_hanoi

    print("=" * 80)
    print("CHẠY STANDALONE: Download báo cáo sửa chữa toàn trình chi tiết")
    print("=" * 80)

    page_baocao = None
    browser_baocao = None
    playwright_baocao = None

    try:
        page_baocao, browser_baocao, playwright_baocao = login_baocao_hanoi()
        result = download_report_sua_chua_toan_trinh(page_baocao)

        if result:
            print(f"\nHOÀN THÀNH! File đã lưu tại: {result}")
        else:
            print("\nKhông tải được file")

    except Exception as e:
        print(f"\nLỗi: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if browser_baocao:
            print("\nĐang đóng browser...")
            browser_baocao.close()
        if playwright_baocao:
            playwright_baocao.stop()
        print("Đã đóng browser")
