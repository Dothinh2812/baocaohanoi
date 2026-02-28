# -*- coding: utf-8 -*-
from playwright.sync_api import sync_playwright
import time
import os
import re
import traceback
from datetime import datetime
from urllib.parse import quote
import pandas as pd
from config import Config
from make_chart_pttb import make_chart_pttb_thuc_tang_fiber, make_chart_pttb_mytv_thuc_tang, make_chart_pttb_thuc_tang_nvkt, make_chart_mytv_thuc_tang_nvkt
from login import login_baocao_hanoi
from KR_process import process_GHTT_report_NVKT
from KR_download import download_GHTT_report_HNI, download_GHTT_report_Son_Tay, download_GHTT_report_nvktdb
from thuc_tang_download import download_report_pttb_ngung_psc, download_report_pttb_hoan_cong, download_report_mytv_hoan_cong, download_report_mytv_ngung_psc
from thuc_tang_process import process_ngung_psc_report, process_hoan_cong_report, create_thuc_tang_report, process_mytv_ngung_psc_report, process_mytv_hoan_cong_report, create_mytv_thuc_tang_report
from vat_tu_thu_hoi_download import download_report_vattu_thuhoi
from vat_tu_thu_hoi_process import vat_tu_thu_hoi_process
from c1_report_download import download_report_c11, download_report_c12,download_report_c12_chitiet_SM2, download_report_c13, download_report_c14,download_report_c14_chitiet, download_report_c15, download_report_I15,download_report_I15_k2, download_report_c11_chitiet, download_report_c11_chitiet_SM2, download_report_c12_chitiet_SM1, download_report_c15_chitiet
from c1_process import process_c11_report, process_c11_chitiet_report, process_c12_report, process_c13_report, process_c14_report,process_c14_chitiet_report, process_c15_report,process_c15_chitiet_report, process_c11_chitiet_report_SM2, process_c12_chitiet_report_SM1SM2
from i15_process import process_I15_report_with_tracking, process_I15_k2_report_with_tracking
from i15_cts_converter import convert_cts_to_i15_report
from suy_hao_reports import generate_daily_comparison_report, generate_daily_comparison_report_k2
from exclusion_process import process_exclusion_reports
from kpi_calculator import tao_bao_cao_kpi
from import_baocao import main as import_baocao_main
from report_generator import generate_kpi_report
from kq_tiep_thi_download import kq_tiep_thi_download
from kq_tiep_thi_process import process_kq_tiep_thi_report

# =============================================================================
# CẤU HÌNH NGÀY BÁO CÁO CHI TIẾT (SM4-C11, SM2-C11, SM1-C12, SM2-C12)
# Định dạng: "dd/mm/yyyy" hoặc None (mặc định: ngày đầu/cuối tháng hiện tại)
# =============================================================================
start_date = "26/01/2026"  # Từ ngày (None = ngày đầu tháng hiện tại)
end_date = "25/02/2026"     # Đến ngày (None = ngày cuối tháng hiện tại)

# =============================================================================
# CẤU HÌNH THÁNG BÁO CÁO (C1.1, C1.2, C1.3, C1.4, C1.5, C1.4 chi tiết, C1.5 chi tiết)
# Định dạng: "Tháng MM/YYYY" hoặc None (mặc định: tháng hiện tại)
# Ví dụ: "Tháng 12/2025", "Tháng 01/2026"
# =============================================================================
report_month = "Tháng 02/2026"  # Tháng báo cáo (None = tháng hiện tại)

# =============================================================================
# CẤU HÌNH GIẢM TRỪ PHIẾU BÁO HỎNG
# Khi ENABLE_EXCLUSION = True: Tạo báo cáo so sánh trước/sau giảm trừ
# Khi ENABLE_EXCLUSION = False: Không tạo báo cáo giảm trừ
# File danh sách loại trừ: du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx (cột BAOHONG_ID)
# Kết quả xuất ra: downloads/kq_sau_giam_tru/
# =============================================================================
ENABLE_EXCLUSION = True # Bật/tắt tính năng giảm trừ


# =============================================================================
# HỆ THỐNG GHI LOG
# =============================================================================
class TaskLogger:
    """Ghi log chi tiết quá trình thực thi"""
    
    def __init__(self):
        self.start_time = datetime.now()
        self.tasks = []  # Danh sách các bước đã thực hiện
        self._current_step = 0
    
    def run_task(self, task_name, func, output_files=None, *args, **kwargs):
        """
        Chạy một tác vụ và ghi log kết quả.
        
        Args:
            task_name: Tên mô tả công việc
            func: Hàm cần gọi
            output_files: list đường dẫn file đầu ra mong đợi
            *args, **kwargs: Tham số truyền vào func
        
        Returns:
            True nếu thành công, False nếu thất bại
        """
        self._current_step += 1
        step_start = time.time()
        task_entry = {
            'step': self._current_step,
            'name': task_name,
            'func_name': func.__name__ if hasattr(func, '__name__') else str(func),
            'status': None,
            'duration': 0,
            'output_files': output_files or [],
            'files_exist': [],
            'error': None
        }
        
        try:
            func(*args, **kwargs)
            task_entry['status'] = '✅ Thành công'
            # Kiểm tra file đầu ra
            if output_files:
                for f in output_files:
                    task_entry['files_exist'].append({
                        'path': f,
                        'exists': os.path.exists(f),
                        'size': os.path.getsize(f) if os.path.exists(f) else 0
                    })
            task_entry['duration'] = round(time.time() - step_start, 1)
            self.tasks.append(task_entry)
            return True
        except Exception as e:
            task_entry['status'] = '❌ Thất bại'
            task_entry['error'] = str(e)
            task_entry['duration'] = round(time.time() - step_start, 1)
            self.tasks.append(task_entry)
            print(f"⚠️ Lỗi khi {task_name}: {e}")
            return False
    
    def add_note(self, note):
        """Thêm ghi chú vào log (không phải task)"""
        self.tasks.append({
            'step': '-',
            'name': f'📝 {note}',
            'func_name': '',
            'status': 'ℹ️ Ghi chú',
            'duration': 0,
            'output_files': [],
            'files_exist': [],
            'error': None
        })
    
    def write_log(self):
        """Ghi toàn bộ log ra file trong thư mục logs/"""
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
        
        timestamp = self.start_time.strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"log_{timestamp}.txt")
        
        total_duration = (datetime.now() - self.start_time).total_seconds()
        minutes = int(total_duration // 60)
        seconds = int(total_duration % 60)
        
        success_tasks = [t for t in self.tasks if t['status'] == '✅ Thành công']
        failed_tasks = [t for t in self.tasks if t['status'] == '❌ Thất bại']
        note_tasks = [t for t in self.tasks if t['status'] == 'ℹ️ Ghi chú']
        
        lines = []
        lines.append("=" * 70)
        lines.append(f"BÁO CÁO THỰC THI - {self.start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("=" * 70)
        lines.append("")
        lines.append("CẤU HÌNH:")
        lines.append(f"  Kỳ báo cáo     : {report_month}")
        lines.append(f"  Ngày bắt đầu   : {start_date}")
        lines.append(f"  Ngày kết thúc   : {end_date}")
        lines.append(f"  Giảm trừ        : {'Bật' if ENABLE_EXCLUSION else 'Tắt'}")
        lines.append("")
        
        # Chi tiết từng bước
        lines.append("=" * 70)
        lines.append("CHI TIẾT TỪNG BƯỚC")
        lines.append("=" * 70)
        
        for task in self.tasks:
            lines.append("")
            lines.append(f"[{task['step']}] {task['name']}")
            if task['func_name']:
                lines.append(f"    Hàm    : {task['func_name']}")
            lines.append(f"    Kết quả: {task['status']}")
            if task['duration'] > 0:
                lines.append(f"    Thời gian: {task['duration']}s")
            if task['error']:
                lines.append(f"    Lỗi    : {task['error']}")
            if task['files_exist']:
                for fi in task['files_exist']:
                    if fi['exists']:
                        size_kb = round(fi['size'] / 1024, 1)
                        lines.append(f"    File   : ✅ {fi['path']} ({size_kb} KB)")
                    else:
                        lines.append(f"    File   : ❌ {fi['path']} (KHÔNG TỒN TẠI)")
        
        # Tổng kết
        lines.append("")
        lines.append("=" * 70)
        lines.append("TỔNG KẾT")
        lines.append("=" * 70)
        lines.append(f"  Tổng số bước : {len(success_tasks) + len(failed_tasks)}")
        lines.append(f"  Thành công   : {len(success_tasks)}")
        lines.append(f"  Thất bại     : {len(failed_tasks)}")
        lines.append(f"  Ghi chú      : {len(note_tasks)}")
        lines.append(f"  Tổng thời gian: {minutes} phút {seconds} giây")
        
        if failed_tasks:
            lines.append("")
            lines.append("DANH SÁCH LỖI:")
            for task in failed_tasks:
                lines.append(f"  [{task['step']}] {task['name']}: {task['error']}")
        
        # Kiểm tra file đầu ra quan trọng
        lines.append("")
        lines.append("=" * 70)
        lines.append("FILE ĐẦU RA CHÍNH")
        lines.append("=" * 70)
        important_files = [
            ("Báo cáo C1.1 tổng hợp", "downloads/baocao_hanoi/c1.1 report.xlsx"),
            ("Báo cáo C1.1 SM4 chi tiết", "downloads/baocao_hanoi/SM4-C11.xlsx"),
            ("Báo cáo C1.1 SM2 chi tiết", "downloads/baocao_hanoi/SM2-C11.xlsx"),
            ("Báo cáo C1.2 tổng hợp", "downloads/baocao_hanoi/c1.2 report.xlsx"),
            ("Báo cáo C1.2 SM1 chi tiết", "downloads/baocao_hanoi/SM1-C12.xlsx"),
            ("Báo cáo C1.2 SM2 chi tiết", "downloads/baocao_hanoi/SM2-C12.xlsx"),
            ("Báo cáo C1.3 tổng hợp", "downloads/baocao_hanoi/c1.3 report.xlsx"),
            ("Báo cáo C1.4 tổng hợp", "downloads/baocao_hanoi/c1.4 report.xlsx"),
            ("Báo cáo C1.4 chi tiết", "downloads/baocao_hanoi/c1.4_chitiet_report.xlsx"),
            ("Báo cáo C1.5 chi tiết", "downloads/baocao_hanoi/c1.5_chitiet_report.xlsx"),
            ("So sánh C1.1 SM4", "downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx"),
            ("So sánh C1.1 SM2", "downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx"),
            ("So sánh C1.2 SM1", "downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx"),
            ("So sánh C1.4", "downloads/kq_sau_giam_tru/So_sanh_C14.xlsx"),
            ("So sánh C1.5", "downloads/kq_sau_giam_tru/So_sanh_C15.xlsx"),
            ("Tổng hợp BSC đơn vị", "downloads/kq_sau_giam_tru/Tong_hop_Diem_BSC_Don_Vi.xlsx"),
            ("KPI NVKT tóm tắt", "downloads/KPI/KPI_NVKT_TomTat.xlsx"),
            ("KPI NVKT chi tiết", "downloads/KPI/KPI_NVKT_ChiTiet.xlsx"),
        ]
        # Tìm file báo cáo Word động
        reports_dir = "downloads/reports"
        if os.path.exists(reports_dir):
            for f in os.listdir(reports_dir):
                if f.startswith("Bao_cao_KPI_NVKT_") and f.endswith(".docx"):
                    important_files.append(("Báo cáo Word KPI", os.path.join(reports_dir, f)))
        
        for label, fpath in important_files:
            if os.path.exists(fpath):
                size_kb = round(os.path.getsize(fpath) / 1024, 1)
                mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%H:%M:%S")
                lines.append(f"  ✅ {label}")
                lines.append(f"     {fpath} ({size_kb} KB, cập nhật: {mtime})")
            else:
                lines.append(f"  ❌ {label}")
                lines.append(f"     {fpath} (KHÔNG TỒN TẠI)")
        
        lines.append("")
        lines.append("=" * 70)
        lines.append(f"Kết thúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("=" * 70)
        
        # Ghi file
        log_content = "\n".join(lines)
        with open(log_file, "w", encoding="utf-8") as f:
            f.write(log_content)
        
        print(f"\n📋 Đã ghi log: {log_file}")
        return log_file


def main():
    """
    Hàm main chạy toàn bộ workflow
    """
    logger = TaskLogger()
    
    try:
        #Đăng nhập
        page_baocao, browser_baocao, playwright_baocao = login_baocao_hanoi()
        logger.add_note("Đăng nhập thành công")

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

        # #Tải báo cáo GHTT trước cước
        # print("\n=== Bắt đầu tải báo cáo GHTT trước cước ===")
        download_GHTT_report_HNI(page_baocao)
        download_GHTT_report_Son_Tay(page_baocao)
        download_GHTT_report_nvktdb(page_baocao)
        #print("\n=== Bắt đầu tải báo cáo GHTT trước cước ===")
        process_GHTT_report_NVKT()
        # #Tải báo cáo tiếp thị fiber, mytv
        # print("\n=== Bắt đầu tải báo cáo Tiếp thị Fiber, MyTV ===")
        kq_tiep_thi_download(page_baocao)
        # #Xử lý báo cáo tiếp thị fiber, mytv
        # print("\n=== Bắt đầu xử lý báo cáo Tiếp thị Fiber, MyTV ===")
        process_kq_tiep_thi_report()
    

        # Tải và xử lý báo cáo C1.x
        print("\n=== Bắt đầu tải các báo cáo C1.x ===")
        
        # Danh sách lưu các báo cáo bị lỗi
        failed_reports = []
        
        # C1.1
        logger.run_task(
            "Tải báo cáo C1.1 tổng hợp",
            lambda: (download_report_c11(page_baocao, report_month), process_c11_report()),
            ["downloads/baocao_hanoi/c1.1 report.xlsx"]
        ) or failed_reports.append("C1.1")
        
        # C11 chi tiết SM4 (sử dụng start_date và end_date)
        logger.run_task(
            "Tải + xử lý C1.1 chi tiết SM4",
            lambda: (download_report_c11_chitiet(page_baocao, start_date, end_date), process_c11_chitiet_report()),
            ["downloads/baocao_hanoi/SM4-C11.xlsx"]
        ) or failed_reports.append("C11 chi tiết SM4")
        
        # C11 chi tiết SM2
        def _task_c11_sm2():
            download_report_c11_chitiet_SM2(page_baocao, start_date, end_date)
            process_c11_chitiet_report_SM2()
            from c1_process import process_c11_ct_tp1_72h
            process_c11_ct_tp1_72h()
        
        logger.run_task(
            "Tải + xử lý C1.1 chi tiết SM2",
            _task_c11_sm2,
            ["downloads/baocao_hanoi/SM2-C11.xlsx"]
        ) or failed_reports.append("C11 chi tiết SM2")
        
        # C12 chi tiết SM1 (sử dụng start_date và end_date)
        logger.run_task(
            "Tải C1.2 chi tiết SM1",
            lambda: download_report_c12_chitiet_SM1(page_baocao, start_date, end_date),
            ["downloads/baocao_hanoi/SM1-C12.xlsx"]
        ) or failed_reports.append("C12 chi tiết SM1")
        
        # C12 chi tiết SM2
        logger.run_task(
            "Tải C1.2 chi tiết SM2",
            lambda: download_report_c12_chitiet_SM2(page_baocao, start_date, end_date),
            ["downloads/baocao_hanoi/SM2-C12.xlsx"]
        ) or failed_reports.append("C12 chi tiết SM2")
        
        # Xử lý C12 SM1/SM2
        logger.run_task(
            "Xử lý C1.2 chi tiết SM1/SM2",
            process_c12_chitiet_report_SM1SM2,
            ["downloads/baocao_hanoi/SM1-C12.xlsx", "downloads/baocao_hanoi/SM2-C12.xlsx"]
        ) or failed_reports.append("Xử lý C12 chi tiết SM1/SM2")
        
        # C1.2
        logger.run_task(
            "Tải + xử lý C1.2 tổng hợp",
            lambda: (download_report_c12(page_baocao, report_month), process_c12_report()),
            ["downloads/baocao_hanoi/c1.2 report.xlsx"]
        ) or failed_reports.append("C1.2")
        
        # C1.3
        logger.run_task(
            "Tải + xử lý C1.3 tổng hợp",
            lambda: (download_report_c13(page_baocao, report_month), process_c13_report()),
            ["downloads/baocao_hanoi/c1.3 report.xlsx"]
        ) or failed_reports.append("C1.3")
        
        # C1.4 và chi tiết
        logger.run_task(
            "Tải + xử lý C1.4 tổng hợp",
            lambda: (download_report_c14(page_baocao, report_month), process_c14_report()),
            ["downloads/baocao_hanoi/c1.4 report.xlsx"]
        ) or failed_reports.append("C1.4")
        
        logger.run_task(
            "Tải + xử lý C1.4 chi tiết",
            lambda: (download_report_c14_chitiet(page_baocao, report_month), process_c14_chitiet_report()),
            ["downloads/baocao_hanoi/c1.4_chitiet_report.xlsx"]
        ) or failed_reports.append("C1.4 chi tiết")

        # C1.5 chi tiết
        logger.run_task(
            "Tải + xử lý C1.5 chi tiết",
            lambda: (download_report_c15_chitiet(page_baocao), process_c15_chitiet_report()),
            ["downloads/baocao_hanoi/c1.5_chitiet_report.xlsx"]
        ) or failed_reports.append("C1.5 chi tiết")
        
        # I1.5 (đang tạm tắt)
        logger.add_note("I1.5: Đang tạm tắt (comment out)")
        
        # I1.5 K2
        logger.run_task(
            "Tải + xử lý I1.5 K2",
            lambda: (download_report_I15_k2(page_baocao), process_I15_k2_report_with_tracking()),
        ) or failed_reports.append("I1.5 K2")
        
        # Tạo báo cáo so sánh SHC ngày (T so với T-1)
        logger.run_task(
            "Tạo báo cáo so sánh SHC ngày",
            generate_daily_comparison_report,
        ) or failed_reports.append("So sánh SHC ngày")
        
        # Tạo báo cáo so sánh SHC K2 ngày (T so với T-1)
        logger.run_task(
            "Tạo báo cáo so sánh SHC K2 ngày",
            generate_daily_comparison_report_k2,
        ) or failed_reports.append("So sánh SHC K2 ngày")
        
        # In tổng kết các báo cáo bị lỗi
        if failed_reports:
            print("\n" + "="*60)
            print(f"⚠️ CÓ {len(failed_reports)} BÁO CÁO BỊ LỖI:")
            for report_name in failed_reports:
                print(f"   - {report_name}")
            print("="*60)

        time.sleep(1)

        # Đóng browser và playwright trước khi xử lý file
        print("\nĐang đóng trình duyệt...")
        browser_baocao.close()
        playwright_baocao.stop()
        logger.add_note("Đã đóng trình duyệt")

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
            logger.run_task(
                "Xử lý giảm trừ (exclusion_process)",
                process_exclusion_reports,
                [
                    "downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx",
                    "downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx",
                    "downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx",
                    "downloads/kq_sau_giam_tru/So_sanh_C14.xlsx",
                    "downloads/kq_sau_giam_tru/So_sanh_C15.xlsx",
                    "downloads/kq_sau_giam_tru/Tong_hop_Diem_BSC_Don_Vi.xlsx",
                    "downloads/kq_sau_giam_tru/Tong_hop_giam_tru.xlsx",
                ]
            )
        else:
            logger.add_note("Giảm trừ: TẮT (bỏ qua)")

        # Tính điểm KPI cho NVKT
        print("\n=== Tính điểm KPI NVKT ===")
        logger.run_task(
            "Tính điểm KPI NVKT (trước giảm trừ)",
            lambda: tao_bao_cao_kpi("downloads/baocao_hanoi", "downloads/KPI", "TRƯỚC GIẢM TRỪ"),
            ["downloads/KPI/KPI_NVKT_TomTat.xlsx", "downloads/KPI/KPI_NVKT_ChiTiet.xlsx"]
        )

        # Tạo báo cáo Word KPI
        print("\n=== Tạo báo cáo Word KPI ===")
        _rm = report_month.replace("Tháng ", "") if report_month else None
        logger.run_task(
            "Tạo báo cáo Word KPI",
            lambda: generate_kpi_report(
                kpi_folder="downloads/KPI",
                output_folder="downloads/reports",
                report_month=_rm
            ),
            [f"downloads/reports/Bao_cao_KPI_NVKT_{_rm.replace('/', '_')}.docx"] if _rm else []
        )

        # Import dữ liệu vào SQLite database
        print("\n=== Import dữ liệu vào database ===")
        logger.run_task(
            "Import dữ liệu vào database",
            import_baocao_main,
            ["database.db"]
        )

        print("\n✅ Hoàn thành toàn bộ quá trình!")

    except Exception as e:
        print(f"❌ Có lỗi xảy ra: {str(e)}")
        traceback.print_exc()
        logger.add_note(f"LỖI NGHIÊM TRỌNG: {str(e)}")

    finally:
        # Đóng browser và playwright (nếu chưa đóng)
        try:
            browser_baocao.close()
            playwright_baocao.stop()
        except:
            pass  # Đã đóng rồi hoặc có lỗi
        
        # Ghi log ra file
        try:
            log_file = logger.write_log()
            print(f"📋 Log đã được ghi vào: {log_file}")
        except Exception as e:
            print(f"⚠️ Không thể ghi log: {e}")


if __name__ == "__main__":
    main()
