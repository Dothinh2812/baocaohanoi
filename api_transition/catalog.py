# -*- coding: utf-8 -*-
"""Danh mục các hàm download cũ và trạng thái chuyển đổi."""

from dataclasses import dataclass


@dataclass(frozen=True)
class MigrationItem:
    old_function: str
    report_page_url: str
    status: str
    recipe_name: str = ""
    note: str = ""


MIGRATION_CATALOG = [
    MigrationItem(
        old_function="download_report_c11",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=534964&menu_id=535020",
        status="implemented",
        recipe_name="c11_q2_2026",
        note="Đã export thành công qua API get-data-export.",
    ),
    MigrationItem(
        old_function="download_report_c12",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=522513&menu_id=535021",
        status="implemented",
        recipe_name="c12_q2_2026",
        note="Đã capture thành công và có downloader API riêng.",
    ),
    MigrationItem(
        old_function="download_report_c13",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=522600&menu_id=535022",
        status="implemented",
        recipe_name="c13_q2_2026",
        note="Đã capture thành công và có downloader API riêng.",
    ),
    MigrationItem(
        old_function="download_report_c14",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=264107&menu_id=275688",
        status="implemented",
        recipe_name="c14_q2_2026",
        note="Đã capture thành công và có downloader API riêng.",
    ),
    MigrationItem(
        old_function="download_report_c14_chitiet",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=240277&menu_id=275687",
        status="implemented",
        recipe_name="c14_chitiet_q2_2026",
        note="Đã capture thành công và có downloader API riêng.",
    ),
    MigrationItem(
        old_function="download_report_c15",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=258310&menu_id=305918",
        status="capture_required",
        recipe_name="c15_q2_2026",
        note="Đã có recipe bước đầu nhưng luồng export hiện đang lỗi, tạm hoãn xử lý tiếp.",
    ),
    MigrationItem(
        old_function="download_report_c15_chitiet",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info-data?id=522920&ploaibc=1&pdonvi_id=284656&pthoigianid=98944630",
        status="capture_optional",
        note="Đang có luồng gần API sẵn nhưng hiện chạy lỗi, tạm hoãn cùng C1.5 tổng hợp.",
    ),
    MigrationItem(
        old_function="download_report_c11_chitiet",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194",
        status="implemented",
        recipe_name="c11_chitiet_q2_2026",
        note="Đã capture thành công và có downloader API riêng.",
    ),
    MigrationItem(
        old_function="download_report_c11_chitiet_SM2",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194",
        status="capture_required",
        note="Luồng cũ đang lỗi khi tải báo cáo, chưa capture được recipe ổn định.",
    ),
    MigrationItem(
        old_function="download_report_c12_chitiet_SM1",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194",
        status="implemented",
        recipe_name="c12_chitiet_sm1_q2_2026",
        note="Đã capture thành công và có downloader API riêng.",
    ),
    MigrationItem(
        old_function="download_report_c12_chitiet_SM2",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=267215&menu_id=276194",
        status="implemented",
        recipe_name="c12_chitiet_sm2_q2_2026",
        note="Đã capture thành công và có downloader API riêng.",
    ),
    MigrationItem(
        old_function="download_report_I15",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=283632&menu_id=283669",
        status="implemented",
        recipe_name="i15_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
    MigrationItem(
        old_function="download_report_I15_k2",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=290125&menu_id=290161",
        status="implemented",
        recipe_name="i15_k2_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
    MigrationItem(
        old_function="download_GHTT_report_HNI",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238",
        status="implemented",
        recipe_name="ghtt_hni_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
    MigrationItem(
        old_function="download_GHTT_report_Son_Tay",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238",
        status="implemented",
        recipe_name="ghtt_sontay_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
    MigrationItem(
        old_function="download_GHTT_report_nvktdb",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=534220&menu_id=534238",
        status="implemented",
        recipe_name="ghtt_nvktdb_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
    MigrationItem(
        old_function="download_KR7_report_NVKT",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=521580",
        status="capture_required",
    ),
    MigrationItem(
        old_function="download_KR7_report_tong_hop",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=521580",
        status="capture_required",
    ),
    MigrationItem(
        old_function="download_report_vattu_thuhoi",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=270922&menu_id=276242",
        status="implemented",
        recipe_name="vattu_thuhoi_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
    MigrationItem(
        old_function="xac_minh_tam_dung_download",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=267844&menu_id=276199",
        status="implemented",
        recipe_name="xac_minh_tam_dung_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
    MigrationItem(
        old_function="kq_tiep_thi_download",
        report_page_url="https://baocao.hanoi.vnpt.vn/report/report-info?id=257495&menu_id=276101",
        status="implemented",
        recipe_name="kq_tiep_thi_q2_2026",
        note="Đã capture thành công, có downloader API riêng và đã tải file thành công trên môi trường thực.",
    ),
]
