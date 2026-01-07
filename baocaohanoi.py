# -*- coding: utf-8 -*-
from playwright.sync_api import sync_playwright
import time
import os
import re
from datetime import datetime
from urllib.parse import quote
import pandas as pd
from config import Config
from make_chart_pttb import make_chart_pttb_thuc_tang_fiber, make_chart_pttb_mytv_thuc_tang, make_chart_pttb_thuc_tang_nvkt, make_chart_mytv_thuc_tang_nvkt
from login import login_baocao_hanoi
from KR_process import process_KR6_report_NVKT, process_KR6_report_tong_hop, process_KR7_report_NVKT, process_KR7_report_tong_hop
from KR_download import download_KR6_report_NVKT, download_KR6_report_tong_hop, download_KR7_report_NVKT, download_KR7_report_tong_hop
from thuc_tang_download import download_report_pttb_ngung_psc, download_report_pttb_hoan_cong, download_report_mytv_hoan_cong, download_report_mytv_ngung_psc
from thuc_tang_process import process_ngung_psc_report, process_hoan_cong_report, create_thuc_tang_report, process_mytv_ngung_psc_report, process_mytv_hoan_cong_report, create_mytv_thuc_tang_report
from vat_tu_thu_hoi_download import download_report_vattu_thuhoi
from vat_tu_thu_hoi_process import vat_tu_thu_hoi_process
from c1_report_download import download_report_c11, download_report_c12,download_report_c12_chitiet_SM2, download_report_c13, download_report_c14,download_report_c14_chitiet, download_report_c15, download_report_I15,download_report_I15_k2, download_report_c11_chitiet, download_report_c11_chitiet_SM2, download_report_c12_chitiet_SM1, download_report_c15_chitiet
from c1_process import process_c11_report, process_c11_chitiet_report, process_c12_report, process_c13_report, process_c14_report,process_c14_chitiet_report, process_c15_report,process_c15_chitiet_report, process_I15_report, process_I15_k2_report, process_c11_chitiet_report_SM2, process_c12_chitiet_report_SM1SM2
from suy_hao_reports import generate_daily_comparison_report, generate_daily_comparison_report_k2
from exclusion_process import process_exclusion_reports
from kpi_calculator import tao_bao_cao_kpi
from import_baocao import main as import_baocao_main

# =============================================================================
# CẤU HÌNH NGÀY BÁO CÁO CHI TIẾT (SM4-C11, SM2-C11, SM1-C12, SM2-C12)
# Định dạng: "dd/mm/yyyy" hoặc None (mặc định: ngày đầu/cuối tháng hiện tại)
# =============================================================================
start_date = "01/01/2026"  # Từ ngày (None = ngày đầu tháng hiện tại)
end_date = "31/01/2026"     # Đến ngày (None = ngày cuối tháng hiện tại)

# =============================================================================
# CẤU HÌNH THÁNG BÁO CÁO (C1.1, C1.2, C1.3, C1.4, C1.5, C1.4 chi tiết, C1.5 chi tiết)
# Định dạng: "Tháng MM/YYYY" hoặc None (mặc định: tháng hiện tại)
# Ví dụ: "Tháng 12/2025", "Tháng 01/2026"
# =============================================================================
report_month = "Tháng 01/2026"  # Tháng báo cáo (None = tháng hiện tại)

# =============================================================================
# CẤU HÌNH GIẢM TRỪ PHIẾU BÁO HỎNG
# Khi ENABLE_EXCLUSION = True: Tạo báo cáo so sánh trước/sau giảm trừ
# Khi ENABLE_EXCLUSION = False: Không tạo báo cáo giảm trừ
# File danh sách loại trừ: du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx (cột BAOHONG_ID)
# Kết quả xuất ra: downloads/kq_sau_giam_tru/
# =============================================================================
ENABLE_EXCLUSION = False  # Bật/tắt tính năng giảm trừ


def main():
    """
    Hàm main chạy toàn bộ workflow
    """
    try:
        #Đăng nhập
        page_baocao, browser_baocao, playwright_baocao = login_baocao_hanoi()

        # # #Tải báo cáo PTTB
        # download_report_pttb_ngung_psc(page_baocao)
        # download_report_pttb_hoan_cong(page_baocao)

        # #Tải báo cáo MyTV
        # download_report_mytv_ngung_psc(page_baocao)
        # download_report_mytv_hoan_cong(page_baocao)
        # #Tải báo cáo thu hồi vật tư
        # print("\n=== Bắt đầu tải báo cáo vật tư thu hồi===")
        # download_report_vattu_thuhoi(page_baocao)
        # #xử lý báo cáo thu hồi vật tư
        # print("\n=== Bắt đầu xử lý báo cáo vật tư thu hồi ===")
        # vat_tu_thu_hoi_process()

        # #Tải báo cáo KR6, KR7
        # print("\n=== Bắt đầu tải các báo cáo KR ===")
        # download_KR6_report_NVKT(page_baocao)
        # process_KR6_report_NVKT()
        # download_KR6_report_tong_hop(page_baocao)
        # process_KR6_report_tong_hop()
        # download_KR7_report_NVKT(page_baocao)
        # process_KR7_report_NVKT()
        # download_KR7_report_tong_hop(page_baocao)
        # process_KR7_report_tong_hop()

        # Tải và xử lý báo cáo C1.x
        print("\n=== Bắt đầu tải các báo cáo C1.x ===")
        download_report_c11(page_baocao, report_month)
        process_c11_report()
        # C11 chi tiết (sử dụng start_date và end_date)
        download_report_c11_chitiet(page_baocao, start_date, end_date)
        process_c11_chitiet_report()
        download_report_c11_chitiet_SM2(page_baocao, start_date, end_date)
        process_c11_chitiet_report_SM2()
        # C12 chi tiết SM1, SM2 (sử dụng start_date và end_date)
        download_report_c12_chitiet_SM1(page_baocao, start_date, end_date)
        download_report_c12_chitiet_SM2(page_baocao, start_date, end_date)
        process_c12_chitiet_report_SM1SM2()
        download_report_c12(page_baocao, report_month)
        process_c12_report()
        download_report_c13(page_baocao, report_month)
        process_c13_report()
        #C14 và chi tiết
        download_report_c14(page_baocao, report_month)
        process_c14_report()
        download_report_c14_chitiet(page_baocao, report_month)
        process_c14_chitiet_report()

        # download_report_c15(page_baocao, report_month)
        # process_c15_report()
         # C15 chi tiết
        download_report_c15_chitiet(page_baocao)
        process_c15_chitiet_report()
        download_report_I15(page_baocao)
        process_I15_report()
        download_report_I15_k2(page_baocao)
        process_I15_k2_report()
        
        # Tạo báo cáo so sánh SHC ngày (T so với T-1)
        generate_daily_comparison_report()
        
        # Tạo báo cáo so sánh SHC K2 ngày (T so với T-1)
        generate_daily_comparison_report_k2()

        #     print("\n✅ Hoàn thành tải báo cáo!")
        #     print("Trình duyệt sẽ giữ mở trong 10 giây để bạn kiểm tra.")
        #     print("Bạn có thể đóng trình duyệt thủ công hoặc đợi tự động đóng...")

        #     # Giữ trình duyệt mở 10 giây
        time.sleep(1)

        # Đóng browser và playwright trước khi xử lý file
        print("\nĐang đóng trình duyệt...")
        browser_baocao.close()
        playwright_baocao.stop()

        # # Xử lý báo cáo PTTB Ngưng PSC
        # process_ngung_psc_report()
        # # Xử lý báo cáo PTTB Hoàn công
        # process_hoan_cong_report()
        # # Tạo báo cáo PTTB Thực tăng
        # create_thuc_tang_report()


        # # Xử lý báo cáo MyTV Ngưng PSC
        # process_mytv_ngung_psc_report()
        # # Xử lý báo cáo MyTV Hoàn công
        # process_mytv_hoan_cong_report()
        # # Tạo báo cáo MyTV Thực tăng
        # create_mytv_thuc_tang_report()

        # #xử lý các báo cáo KR 
        # print("\n=== Bắt đầu xử lý các báo cáo KR ===")
        # process_KR6_report_NVKT()
        # process_KR6_report_tong_hop()
        # process_KR7_report_NVKT()
        # process_KR7_report_tong_hop()

        # #xử lý báo cáo thu hồi vật tư
        # print("\n=== Bắt đầu xử lý báo cáo vật tư thu hồi ===")
        # vat_tu_thu_hoi_process()

        # # Tạo biểu đồ Thực tăng PTTB
        # make_chart_pttb_thuc_tang_fiber()
        # # Tạo biểu đồ Thực tăng MyTV
        # make_chart_pttb_mytv_thuc_tang()

        # # Biểu đồ theo NVKT
        # print("\n=== Tạo biểu đồ theo NVKT ===")
        # make_chart_pttb_thuc_tang_nvkt()
        # make_chart_mytv_thuc_tang_nvkt()

        # Tạo báo cáo so sánh giảm trừ (nếu được bật)
        if ENABLE_EXCLUSION:
            print("\n=== Bắt đầu tạo báo cáo so sánh giảm trừ ===")
            process_exclusion_reports()

        # Tính điểm KPI cho NVKT
        print("\n=== Tính điểm KPI NVKT ===")
        tao_bao_cao_kpi("downloads/baocao_hanoi", "downloads/KPI", "TRƯỚC GIẢM TRỪ")

        # Import dữ liệu vào SQLite database
        print("\n=== Import dữ liệu vào database ===")
        import_baocao_main()

        print("\n✅ Hoàn thành toàn bộ quá trình!")


        # #xử lý các báo cáo C1.x
        # print("\n=== Bắt đầu xử lý các báo cáo C1.x ===")
        # process_c11_report()
        # process_c11_chitiet_report()
        # process_c11_chitiet_report_SM2()
        # process_c12_chitiet_report_SM1SM2()
        # process_c12_report()
        # process_c13_report()
        # process_c14_report()
        # process_c14_chitiet_report()
        # process_c15_report()
        # process_c15_chitiet_report()
        # process_I15_report()

    except Exception as e:
        print(f"❌ Có lỗi xảy ra: {str(e)}")
        import traceback
        traceback.print_exc()

    finally:
        # Đóng browser và playwright (nếu chưa đóng)
        try:
            browser_baocao.close()
            playwright_baocao.stop()
        except:
            pass  # Đã đóng rồi hoặc có lỗi


if __name__ == "__main__":
    main()
