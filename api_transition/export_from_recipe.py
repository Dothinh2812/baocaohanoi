# -*- coding: utf-8 -*-
"""CLI generic để export bằng recipe đã capture."""

import argparse
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from api_transition.downloaders import (
    download_cau_hinh_tu_dong_api,
    download_cau_hinh_tu_dong_chi_tiet_api,
    download_cau_hinh_tu_dong_ptm_api,
    download_cau_hinh_tu_dong_thay_the_api,
    download_report_c11_api,
    download_report_c11_chitiet_api,
    download_report_c12_api,
    download_report_c12_chitiet_sm1_api,
    download_report_c12_chitiet_sm2_api,
    download_report_c13_api,
    download_report_c14_api,
    download_report_c14_chitiet_api,
    download_ghtt_report_hni_api,
    download_ghtt_report_nvktdb_api,
    download_ghtt_report_sontay_api,
    download_report_i15_api,
    download_report_i15_k2_api,
    download_kq_tiep_thi_api,
    download_report_vattu_thuhoi_api,
    download_xac_minh_tam_dung_api,
    download_with_recipe,
)


def parse_key_value(pairs):
    overrides = {}
    for item in pairs or []:
        if "=" not in item:
            raise ValueError(f"Override không hợp lệ: {item}. Dùng dạng key=value")
        key, value = item.split("=", 1)
        overrides[key.strip()] = value.strip()
    return overrides


def parse_args():
    parser = argparse.ArgumentParser(description="Export báo cáo bằng recipe API")
    parser.add_argument("--recipe", default="", help="Tên recipe, ví dụ c11_q2_2026")
    parser.add_argument("--month-id", default="")
    parser.add_argument("--month-label", default="")
    parser.add_argument("--set", dest="overrides", action="append", default=[])
    parser.add_argument("--start-date", default="")
    parser.add_argument("--end-date", default="")
    parser.add_argument("--output-dir", default="", help="Thư mục đầu ra. Bỏ trống để dùng thư mục nhóm mặc định trong api_transition/downloads/")
    parser.add_argument("--output-name", default="")
    parser.add_argument("--headed", action="store_true")
    parser.add_argument("--c11", action="store_true", help="Shortcut chạy downloader C1.1 đã xác nhận")
    parser.add_argument("--c11-chi-tiet", action="store_true", help="Shortcut chạy downloader C1.1 chi tiết đã xác nhận")
    parser.add_argument("--c12", action="store_true", help="Shortcut chạy downloader C1.2 đã xác nhận")
    parser.add_argument("--c12-chi-tiet-sm1", action="store_true", help="Shortcut chạy downloader C1.2 chi tiết SM1 đã xác nhận")
    parser.add_argument("--c12-chi-tiet-sm2", action="store_true", help="Shortcut chạy downloader C1.2 chi tiết SM2 đã xác nhận")
    parser.add_argument("--c13", action="store_true", help="Shortcut chạy downloader C1.3 đã xác nhận")
    parser.add_argument("--c14", action="store_true", help="Shortcut chạy downloader C1.4 tổng hợp đã xác nhận")
    parser.add_argument("--c14-chi-tiet", action="store_true", help="Shortcut chạy downloader C1.4 chi tiết đã xác nhận")
    parser.add_argument("--i15", action="store_true", help="Shortcut chạy downloader I1.5 đã xác nhận")
    parser.add_argument("--i15-k2", action="store_true", help="Shortcut chạy downloader I1.5 K2 đã xác nhận")
    parser.add_argument("--ghtt-hni", action="store_true", help="Shortcut chạy downloader GHTT HNI đã xác nhận")
    parser.add_argument("--ghtt-sontay", action="store_true", help="Shortcut chạy downloader GHTT Sơn Tây đã xác nhận")
    parser.add_argument("--ghtt-nvktdb", action="store_true", help="Shortcut chạy downloader GHTT NVKT DB đã xác nhận")
    parser.add_argument("--xac-minh-tam-dung", action="store_true", help="Shortcut chạy downloader xác minh tạm dừng")
    parser.add_argument("--kq-tiep-thi", action="store_true", help="Shortcut chạy downloader kết quả tiếp thị")
    parser.add_argument("--vattu-thuhoi", action="store_true", help="Shortcut chạy downloader vật tư thu hồi")
    parser.add_argument("--cau-hinh-tu-dong", action="store_true", help="Shortcut generic chạy downloader báo cáo cấu hình tự động")
    parser.add_argument("--cau-hinh-tu-dong-ptm", action="store_true", help="Shortcut chạy downloader cấu hình tự động loại PTM")
    parser.add_argument("--cau-hinh-tu-dong-thay-the", action="store_true", help="Shortcut chạy downloader cấu hình tự động loại Thay Thế")
    parser.add_argument("--cau-hinh-tu-dong-chi-tiet", action="store_true", help="Shortcut chạy downloader cấu hình tự động chi tiết")
    return parser.parse_args()


def main():
    args = parse_args()
    overrides = parse_key_value(args.overrides)

    if args.c11:
        output_path = download_report_c11_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("ptrungtamid", "14324"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.c11_chi_tiet:
        output_path = download_report_c11_chitiet_api(
            start_date=args.start_date or "26/03/2026",
            end_date=args.end_date or "25/04/2026",
            unit_id=overrides.get("pdonvi_id", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.c12:
        output_path = download_report_c12_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("ptrungtamid", "14324"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.c12_chi_tiet_sm1:
        output_path = download_report_c12_chitiet_sm1_api(
            start_date=args.start_date or "26/03/2026",
            end_date=args.end_date or "25/04/2026",
            unit_id=overrides.get("pdonvi_id", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.c12_chi_tiet_sm2:
        output_path = download_report_c12_chitiet_sm2_api(
            start_date=args.start_date or "26/03/2026",
            end_date=args.end_date or "25/04/2026",
            unit_id=overrides.get("pdonvi_id", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.c13:
        output_path = download_report_c13_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("ptrungtamid", "14324"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.c14:
        output_path = download_report_c14_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("vdonvi", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.c14_chi_tiet:
        output_path = download_report_c14_chitiet_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("vdvvt", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.i15:
        output_path = download_report_i15_api(
            start_date=args.start_date or "14/04/2026",
            end_date=args.end_date or "14/04/2026",
            unit_id=overrides.get("vdv", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.i15_k2:
        output_path = download_report_i15_k2_api(
            start_date=args.start_date or "14/04/2026",
            end_date=args.end_date or "14/04/2026",
            unit_id=overrides.get("vdv", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.ghtt_hni:
        output_path = download_ghtt_report_hni_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("vdonvi", "284412"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.ghtt_sontay:
        output_path = download_ghtt_report_sontay_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("vdonvi", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.ghtt_nvktdb:
        output_path = download_ghtt_report_nvktdb_api(
            month_id=args.month_id,
            month_label=args.month_label,
            unit_id=overrides.get("vdonvi", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.xac_minh_tam_dung:
        output_path = download_xac_minh_tam_dung_api(
            start_date=args.start_date or "01/04/2026",
            end_date=args.end_date or "16/04/2026",
            unit_id=overrides.get("pdonvi_id", "284656"),
            service_ids=overrides.get("vloaidv", "8,9"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.kq_tiep_thi:
        output_path = download_kq_tiep_thi_api(
            start_date=args.start_date or "16/04/2026",
            end_date=args.end_date or "16/04/2026",
            unit_id=overrides.get("vdonvi_id", "284656"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.vattu_thuhoi:
        output_path = download_report_vattu_thuhoi_api(
            start_date=args.start_date or "24/11/2025",
            end_date=args.end_date or "16/04/2026",
            unit_id=overrides.get("vttvt", "284656"),
            service_ids=overrides.get("vdichvuvt_erp", "1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16"),
            vat_tu_ids=overrides.get("vvattu", "1,2,3,4,8,6,5"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.cau_hinh_tu_dong:
        output_path = download_cau_hinh_tu_dong_api(
            month_id=args.month_id,
            month_label=args.month_label,
            contract_type=overrides.get("pdv", "1"),
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.cau_hinh_tu_dong_ptm:
        output_path = download_cau_hinh_tu_dong_ptm_api(
            month_id=args.month_id,
            month_label=args.month_label,
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.cau_hinh_tu_dong_thay_the:
        output_path = download_cau_hinh_tu_dong_thay_the_api(
            month_id=args.month_id,
            month_label=args.month_label,
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if args.cau_hinh_tu_dong_chi_tiet:
        output_path = download_cau_hinh_tu_dong_chi_tiet_api(
            month_id=args.month_id,
            month_label=args.month_label,
            headed=args.headed,
            output_dir=args.output_dir,
        )
        print(f"✅ Đã lưu file: {output_path}")
        return

    if not args.recipe:
        raise ValueError(
            "Cần truyền --recipe hoặc dùng --c11/--c11-chi-tiet/--c12/--c12-chi-tiet-sm1/--c12-chi-tiet-sm2/--c13/--c14/--c14-chi-tiet/--i15/--i15-k2/--ghtt-hni/--ghtt-sontay/--ghtt-nvktdb/--xac-minh-tam-dung/--kq-tiep-thi/--vattu-thuhoi/--cau-hinh-tu-dong/--cau-hinh-tu-dong-ptm/--cau-hinh-tu-dong-thay-the/--cau-hinh-tu-dong-chi-tiet"
        )

    output_path = download_with_recipe(
        args.recipe,
        headed=args.headed,
        output_dir=args.output_dir,
        output_name=args.output_name,
        overrides=overrides,
        month_id=args.month_id,
        month_label=args.month_label,
    )
    print(f"✅ Đã lưu file: {output_path}")


if __name__ == "__main__":
    main()
