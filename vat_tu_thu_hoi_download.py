# -*- coding: utf-8 -*-
"""
Module chứa hàm download báo cáo vật tư thu hồi
"""
import time
import os
from config import Config

def download_report_vattu_thuhoi(page_baocao):
    """
    Tải báo cáo Vật tư thu hồi

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo Vật tư thu hồi ===")

    # Truy cập trang báo cáo
    report_url = Config.get_report_url('tbm')
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

    # Bước 2: Điền ngày tháng
    print("Đang điền ngày tháng...")
    try:
        date_field = page_baocao.locator('#mat-input-0')
        date_field.wait_for(state="visible", timeout=30000)
        # Xóa text hiện tại
        date_field.click()
        time.sleep(0.5)
        date_field.press("Control+A")
        time.sleep(0.5)
        # Điền ngày mới
        date_field.fill("24/11/2025")
        time.sleep(1)
        print("✅ Đã điền ngày 24/11/2025")

        # QUAN TRỌNG: Đóng datepicker bằng cách nhấn Escape
        print("Đang đóng datepicker...")
        page_baocao.keyboard.press("Escape")
        time.sleep(1)

        # Đợi backdrop biến mất
        try:
            page_baocao.wait_for_selector(".cdk-overlay-backdrop", state="hidden", timeout=5000)
            print("✅ Datepicker đã đóng")
        except:
            # Nếu không có backdrop hoặc đã biến mất rồi thì OK
            print("✅ Không có backdrop hoặc đã đóng")

        time.sleep(1)

    except Exception as e:
        print(f"⚠️ Lỗi khi điền ngày: {e}")

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

    # Bước 4: Click dropdown "Loại vật tư" (div:nth-child(7)) và chọn "Tất cả"
    print("\nBước 4: Đang click dropdown 'Loại vật tư'...")
    try:
        # Đợi UI ổn định
        time.sleep(1)

        # Sử dụng CSS selector cho button dropdown div[7] - "Loại vật tư"
        button_selector = "div.search-criteria > div:nth-child(7) ngx-dropdown-treeview button"
        dropdown_button = page_baocao.locator(button_selector)
        dropdown_button.wait_for(state="visible", timeout=30000)

        # Click vào dropdown
        dropdown_button.click()
        time.sleep(2)
        print("✅ Đã click dropdown 'Loại vật tư'")

        # Tìm checkbox "Tất cả" trong dropdown menu vừa mở
        # Find and click the "Tất cả" checkbox for "Loại vật tư"
        try:
            # Giới hạn tìm kiếm trong div:nth-child(7) để tránh nhầm lẫn với dropdown khác
            checkbox_selector = "div.search-criteria > div:nth-child(7) ngx-dropdown-treeview label.form-check-label"
            checkbox = page_baocao.locator(checkbox_selector).filter(has_text="Tất cả").first
            checkbox.wait_for(state="visible", timeout=30000)
            checkbox.click()
            time.sleep(2)
            print("✅ Đã chọn 'Tất cả' cho Loại vật tư")

            # Đóng dropdown bằng Escape
            page_baocao.keyboard.press("Escape")
            time.sleep(1)

        except Exception as e:
            print(f"⚠️ Lỗi khi chọn checkbox 'Tất cả': {e}")
            # Thử cách khác nếu cần


        # Click vào trang để kích hoạt (activate) page
        print("Đang kích hoạt page...")
        page_baocao.click('body')
        time.sleep(1) 

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 4: {e}")
        import traceback
        traceback.print_exc()

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

    # Bước 7: Click vào option 2.Tất cả dữ liệu trong dropdown
    print("\nBước 7: Đang click 2.Tất cả dữ liệu để tải file...")
    try:
        # Đảm bảo thư mục downloads tồn tại
        download_dir = os.path.join("downloads", "baocao_hanoi")
        os.makedirs(download_dir, exist_ok=True)

        time.sleep(2)

        # Tìm tất cả icon download (i tag) trong button-action
        download_icons = page_baocao.locator("div.button-action i")
        print(f"Tìm thấy {download_icons.count()} icon")

        if download_icons.count() >= 2:
            # Click vào icon thứ 2 (nth-child(2))
            target_icon = download_icons.nth(1)  # nth-child(2) tương ứng với nth(1) vì 0-indexed
            target_icon.wait_for(state="visible", timeout=30000)

            # Bắt đầu theo dõi download
            with page_baocao.expect_download(timeout=300000) as download_info:
                target_icon.click()
                print("✅ Đã click vào icon download thứ 2")

            # Lưu file với tên mới
            download = download_info.value
            original_filename = download.suggested_filename
            print(f"Tên file gốc: {original_filename}")

            # Lấy extension từ file gốc
            file_extension = os.path.splitext(original_filename)[1]
            new_filename = f"bc_thu_hoi_vat_tu{file_extension}"

            save_path = os.path.join(download_dir, new_filename)
            download.save_as(save_path)
            print(f"✅ Đã tải file về: {save_path}")
        else:
            print(f"❌ Không tìm thấy đủ icon download. Chỉ tìm thấy {download_icons.count()}")

    except Exception as e:
        print(f"⚠️ Lỗi ở bước 7: {e}")
        import traceback
        traceback.print_exc()
