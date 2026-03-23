# -*- coding: utf-8 -*-
"""
Module chứa hàm download báo cáo Kết quả Tiếp thị
"""
import time
import os
from datetime import datetime
from config import Config


def kq_tiep_thi_download(page_baocao):
    """
    Tải báo cáo Kết quả Tiếp thị từ trang baocao.hanoi.vnpt.vn

    Các bước:
    1. Vào URL báo cáo
    2. Mở dropdown đơn vị, tìm và chọn "ttvt sơn tây"
    3. Click "Xem báo cáo"
    4. Click "Xuất Excel"
    5. Tải file và lưu với tên kq_tiep_thi

    Args:
        page_baocao: Đối tượng page đã đăng nhập
    """
    print("\n=== Bắt đầu tải báo cáo Kết quả Tiếp thị ===")

    url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=257495&menu_id=276101"
    print(f"Đang truy cập: {url}")
    page_baocao.goto(url, timeout=Config.NETWORK_IDLE_TIMEOUT)

    # Đợi trang load
    page_baocao.wait_for_load_state("networkidle", timeout=Config.NETWORK_IDLE_TIMEOUT)
    time.sleep(5)

    try:
        # Bước 1: Click dropdown đơn vị
        print("Đang mở dropdown đơn vị...")
        page_baocao.locator(
            "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/button"
        ).click()
        time.sleep(2)

        # Bước 2: Nhập "ttvt sơn tây" vào ô tìm kiếm
        print("Đang nhập 'ttvt sơn tây'...")
        input_xpath = "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[1]/div[1]/div/input"
        page_baocao.locator(input_xpath).fill("ttvt sơn tây")
        time.sleep(2)

        # Bước 3: Chọn "ttvt sơn tây" trong danh sách
        print("Đang chọn 'ttvt sơn tây'...")
        page_baocao.locator(
            "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[1]/div[3]/div/div/div/div/div/div/ngx-dropdown-treeview-select/ngx-dropdown-treeview/div/div/div/ngx-treeview/div[2]/div/ngx-treeview-item/div/div[2]/ngx-treeview-item/div/div/span"
        ).click()
        time.sleep(2)

        # Bước 4: Click "Xem báo cáo"
        print("Đang click 'Xem báo cáo'...")
        page_baocao.locator(
            "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/button"
        ).click()
        print("✅ Đã click 'Xem báo cáo'. Đang đợi dữ liệu load...")
        page_baocao.wait_for_load_state("networkidle", timeout=Config.NETWORK_IDLE_TIMEOUT)
        time.sleep(5)

        # Bước 5: Click "Xuất Excel"
        print("Đang click 'Xuất Excel'...")
        page_baocao.locator(
            "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/button"
        ).click()
        time.sleep(2)
        print("✅ Đã click 'Xuất Excel'.")

        # Bước 6: Tải file báo cáo
        print("Đang tải file báo cáo...")
        download_xpath = "xpath=/html/body/app-root/app-layout/app-vertical/div[2]/div[2]/div/app-report-info-list/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/i[2]"

        # Đảm bảo thư mục lưu trữ tồn tại
        download_dir = "KQ-TIEP-THI"
        os.makedirs(download_dir, exist_ok=True)

        # Tạo tên file có ngày để dễ quản lý (optional but good practice based on other scripts)
        date_str = datetime.now().strftime("%d%m%Y")
        filename = f"kq_tiep_thi_{date_str}.xlsx"
        save_path = os.path.join(download_dir, filename)

        with page_baocao.expect_download(timeout=300000) as download_info:
            page_baocao.locator(download_xpath).click()

        download = download_info.value
        download.save_as(save_path)
        print(f"✅ Đã tải file về: {save_path}")

    except Exception as e:
        print(f"❌ Lỗi khi tải báo cáo Kết quả Tiếp thị: {e}")
        import traceback
        traceback.print_exc()


def main():
    """
    Hàm main để chạy standalone
    """
    try:
        from login import login_baocao_hanoi

        print("=== Bắt đầu tải báo cáo Kết quả Tiếp thị ===")

        # Đăng nhập
        print("\n1. Đăng nhập vào hệ thống...")
        page_baocao, browser_baocao, playwright_baocao = login_baocao_hanoi()
        print("✅ Đăng nhập thành công!")

        # Tải báo cáo
        print("\n2. Tải báo cáo Kết quả Tiếp thị...")
        kq_tiep_thi_download(page_baocao)

        print("\n✅ Hoàn thành!")

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
        try:
            browser_baocao.close()
            playwright_baocao.stop()
        except:
            pass


if __name__ == "__main__":
    main()
