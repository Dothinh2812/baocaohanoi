# -*- coding: utf-8 -*-
"""Các hàm download API mới trong thư mục chuyển đổi."""

import copy
import json
from datetime import datetime, timedelta
from pathlib import Path

from api_transition.auth import capture_authorization, login
from api_transition.report_api_client import (
    export_report,
    find_value_by_label,
    get_info_report,
    make_common_headers,
    print_candidate_pairs,
    save_export_file,
)

API_TRANSITION_DIR = Path(__file__).resolve().parent
RECIPES_DIR = API_TRANSITION_DIR / "recipes"
DOWNLOADS_DIR = API_TRANSITION_DIR / "downloads"


def group_output_dir(group_name):
    return str(DOWNLOADS_DIR / group_name)


def load_recipe(recipe_name):
    recipe_path = RECIPES_DIR / f"{recipe_name}.json"
    if not recipe_path.exists():
        raise FileNotFoundError(f"Không tìm thấy recipe: {recipe_path}")
    return json.loads(recipe_path.read_text(encoding="utf-8"))


def update_payload_input_values(payload, overrides):
    payload_copy = copy.deepcopy(payload)
    normalized_overrides = {str(k): str(v) for k, v in (overrides or {}).items()}

    for input_param in payload_copy.get("lstInputParams", []):
        param_name = input_param.get("name")
        if param_name in normalized_overrides:
            input_param["value"] = normalized_overrides[param_name]
    return payload_copy


def resolve_month_override(recipe, headers, month_id="", month_label=""):
    if month_id:
        return str(month_id)
    if not month_label:
        return ""

    info_payload = get_info_report(recipe["report_id"], recipe["menu_id"], headers)
    try:
        return str(find_value_by_label(info_payload, month_label))
    except RuntimeError:
        print_candidate_pairs(info_payload, month_label)
        raise


def download_with_recipe(
    recipe_name,
    headed=False,
    output_dir="",
    output_name="",
    overrides=None,
    month_id="",
    month_label="",
    month_override_key="pthang",
    session=None,
):
    recipe = load_recipe(recipe_name)
    own_session = session is None
    playwright = browser = context = page = None

    try:
        if own_session:
            playwright, browser, context, page = login(headless=not headed)
            auth_state = capture_authorization(page, recipe["report_page_url"])
            headers = make_common_headers(auth_state, context.cookies())
        else:
            headers = session["headers"]

        effective_overrides = dict(overrides or {})
        resolved_month = resolve_month_override(recipe, headers, month_id=month_id, month_label=month_label)
        if resolved_month:
            effective_overrides[month_override_key] = resolved_month

        payload = update_payload_input_values(recipe["export_payload"], effective_overrides)
        api_timeout = session.get("api_timeout", 120) if session else 120
        export_response = export_report(headers, payload, timeout=api_timeout)

        target_dir = output_dir or group_output_dir("misc")
        target_name = output_name or recipe.get("default_output_name") or ""
        return save_export_file(export_response, target_dir, target_name)

    finally:
        if own_session:
            if browser is not None:
                browser.close()
            if playwright is not None:
                playwright.stop()


def download_report_c11_api(month_id="", month_label="", unit_id="14324", headed=False, output_dir=group_output_dir("chi_tieu_c"), session=None):
    overrides = {"ptrungtamid": str(unit_id)}
    return download_with_recipe(
        "c11_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.1 report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_report_c12_api(month_id="", month_label="", unit_id="14324", headed=False, output_dir=group_output_dir("chi_tieu_c"), session=None):
    overrides = {"ptrungtamid": str(unit_id)}
    return download_with_recipe(
        "c12_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.2 report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_report_c13_api(month_id="", month_label="", unit_id="14324", headed=False, output_dir=group_output_dir("chi_tieu_c"), session=None):
    overrides = {"ptrungtamid": str(unit_id)}
    return download_with_recipe(
        "c13_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.3 report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_kpi_nvkt_c11_api(month_id="", month_label="", unit_id="14324", headed=False, output_dir="", session=None):
    overrides = {"ptrungtamid": str(unit_id)}
    return download_with_recipe(
        "kpi_nvkt_c11_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c11-nvktdb report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_kpi_nvkt_c12_api(month_id="", month_label="", unit_id="14324", headed=False, output_dir="", session=None):
    overrides = {"ptrungtamid": str(unit_id)}
    return download_with_recipe(
        "kpi_nvkt_c12_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c12-nvktdb report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_kpi_nvkt_c13_api(month_id="", month_label="", unit_id="14324", headed=False, output_dir="", session=None):
    overrides = {"ptrungtamid": str(unit_id)}
    return download_with_recipe(
        "kpi_nvkt_c13_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c13-nvktdb report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_report_c14_api(month_id="", month_label="", unit_id="284656", headed=False, output_dir=group_output_dir("chi_tieu_c"), session=None):
    overrides = {"vdonvi": str(unit_id)}
    return download_with_recipe(
        "c14_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.4 report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        month_override_key="vthoigian",
        session=session,
    )


def download_report_c14_chitiet_api(month_id="", month_label="", unit_id="284656", headed=False, output_dir=group_output_dir("chi_tieu_c"), session=None):
    overrides = {"vdvvt": str(unit_id)}
    return download_with_recipe(
        "c14_chitiet_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.4_chitiet_report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        month_override_key="vthoigian",
        session=session,
    )


def download_report_c11_chitiet_api(
    start_date="26/03/2026",
    end_date="25/04/2026",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("chi_tieu_c"),
    session=None,
):
    overrides = {
        "pdonvi_id": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
    }
    return download_with_recipe(
        "c11_chitiet_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.1_chitiet_report.xlsx",
        overrides=overrides,
        session=session,
    )


def download_report_c12_chitiet_sm1_api(
    start_date="26/03/2026",
    end_date="25/04/2026",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("chi_tieu_c"),
    session=None,
):
    overrides = {
        "pdonvi_id": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
    }
    return download_with_recipe(
        "c12_chitiet_sm1_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.2_chitiet_sm1_report.xlsx",
        overrides=overrides,
        session=session,
    )


def download_report_c12_chitiet_sm2_api(
    start_date="26/03/2026",
    end_date="25/04/2026",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("chi_tieu_c"),
    session=None,
):
    overrides = {
        "pdonvi_id": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
    }
    return download_with_recipe(
        "c12_chitiet_sm2_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="c1.2_chitiet_sm2_report.xlsx",
        overrides=overrides,
        session=session,
    )


def download_report_i15_api(
    start_date="14/04/2026",
    end_date="14/04/2026",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("chi_tieu_i"),
    session=None,
):
    overrides = {
        "vdv": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
        "vdk": "0",
    }
    return download_with_recipe(
        "i15_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="i1.5 report.xlsx",
        overrides=overrides,
        session=session,
    )


def download_report_i15_k2_api(
    start_date="14/04/2026",
    end_date="14/04/2026",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("chi_tieu_i"),
    session=None,
):
    overrides = {
        "vdv": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
    }
    return download_with_recipe(
        "i15_k2_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="i1.5_k2 report.xlsx",
        overrides=overrides,
        session=session,
    )


def download_ghtt_report_hni_api(
    month_id="",
    month_label="",
    unit_id="284412",
    headed=False,
    output_dir=group_output_dir("ghtt"),
    session=None,
):
    overrides = {
        "vdonvi": str(unit_id),
        "vloai": "1",
    }
    return download_with_recipe(
        "ghtt_hni_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="ghtt_hni report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        month_override_key="vthoigian",
        session=session,
    )


def download_ghtt_report_sontay_api(
    month_id="",
    month_label="",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("ghtt"),
    session=None,
):
    overrides = {
        "vdonvi": str(unit_id),
        "vloai": "1",
    }
    return download_with_recipe(
        "ghtt_sontay_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="ghtt_sontay report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        month_override_key="vthoigian",
        session=session,
    )


def download_ghtt_report_nvktdb_api(
    month_id="",
    month_label="",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("ghtt"),
    session=None,
):
    overrides = {
        "vdonvi": str(unit_id),
        "vloai": "2",
    }
    return download_with_recipe(
        "ghtt_nvktdb_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="ghtt_nvktdb report.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        month_override_key="vthoigian",
        session=session,
    )


def download_xac_minh_tam_dung_api(
    start_date="01/04/2026",
    end_date="16/04/2026",
    unit_id="284656",
    service_ids="8,9",
    headed=False,
    output_dir=group_output_dir("xac_minh_tam_dung"),
    session=None,
):
    overrides = {
        "pdonvi_id": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
        "vloaidv": str(service_ids),
        "vloaingay": "0",
        "vloaibc": "0",
    }
    return download_with_recipe(
        "xac_minh_tam_dung_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="xac_minh_tam_dung report.xlsx",
        overrides=overrides,
        session=session,
    )


def download_phieu_hoan_cong_dich_vu_chi_tiet_api(
    start_date="01/04/2026",
    end_date="16/04/2026",
    unit_id="284656",
    service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16",
    customer_type="0",
    contract_type="0",
    ticket_type="0",
    headed=False,
    output_dir=group_output_dir("phieu_hoan_cong_dich_vu"),
    session=None,
):
    overrides = {
        "vdv": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
        "vloaidv": str(service_ids),
        "vloaikh": str(customer_type),
        "vloaihd": str(contract_type),
        "vphieu": str(ticket_type),
    }
    return download_with_recipe(
        "phieu_hoan_cong_dich_vu_chi_tiet_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="phieu_hoan_cong_dich_vu_chi_tiet.xlsx",
        overrides=overrides,
        session=session,
    )


def download_tam_dung_khoi_phuc_dich_vu_chi_tiet_api(
    start_date="01/04/2026",
    end_date="16/04/2026",
    unit_id="284656",
    service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16",
    date_type="1",
    report_type="0",
    headed=False,
    output_dir=group_output_dir("tam_dung_khoi_phuc_dich_vu"),
    session=None,
):
    overrides = {
        "pdonvi_id": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
        "vloaidv": str(service_ids),
        "vloaingay": str(date_type),
        "vloaibc": str(report_type),
    }
    return download_with_recipe(
        "tam_dung_khoi_phuc_dich_vu_chi_tiet_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="tam_dung_khoi_phuc_dich_vu_chi_tiet.xlsx",
        overrides=overrides,
        session=session,
    )


def download_tam_dung_khoi_phuc_dich_vu_tong_hop_api(
    start_date="01/04/2026",
    end_date="16/04/2026",
    unit_id="284656",
    service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16",
    report_type="0",
    headed=False,
    output_dir=group_output_dir("tam_dung_khoi_phuc_dich_vu"),
    session=None,
):
    overrides = {
        "vdv": str(unit_id),
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
        "vloaidv": str(service_ids),
        "vloaibc": str(report_type),
    }
    return download_with_recipe(
        "tam_dung_khoi_phuc_dich_vu_tong_hop_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="tam_dung_khoi_phuc_dich_vu_tong_hop.xlsx",
        overrides=overrides,
        session=session,
    )


def download_ngung_psc_mytv_thang_t_1_cap_ttvt_api(
    report_date="",
    unit_id="14316",
    service_id="8",
    report_type="1",
    headed=False,
    output_dir=group_output_dir("mytv_dich_vu"),
    session=None,
):
    if not report_date:
        report_date = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

    overrides = {
        "vdvvt_id": str(service_id),
        "vdenngay": str(report_date),
        "vdonvi_id": str(unit_id),
        "vloai": str(report_type),
    }
    return download_with_recipe(
        "ngung_psc_mytv_thang_t_1_cap_ttvt_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="ngung_psc_mytv_thang_t-1_cap_ttvt.xlsx",
        overrides=overrides,
        session=session,
    )


def download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api(
    month_id="",
    month_label="",
    unit_id="284656",
    report_scope="-1",
    verification_type="-1",
    headed=False,
    output_dir=group_output_dir("ty_le_xac_minh"),
    session=None,
):
    overrides = {
        "vdv": str(unit_id),
        "vloai": str(report_scope),
        "vloaixacminh": str(verification_type),
    }
    return download_with_recipe(
        "ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        month_override_key="vthoigian",
        session=session,
    )


def download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api(
    month_id="",
    month_label="",
    unit_id="284656",
    report_type="-1",
    contract_type="-1",
    service_type="-1",
    headed=False,
    output_dir=group_output_dir("ty_le_xac_minh"),
    session=None,
):
    overrides = {
        "ploaibc": str(report_type),
        "pdonvi_id": str(unit_id),
        "ploaihd": str(contract_type),
        "ploaidv": str(service_type),
    }
    return download_with_recipe(
        "ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet.xlsx",
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        month_override_key="pthoigianid",
        session=session,
    )


def download_kq_tiep_thi_api(
    start_date="16/04/2026",
    end_date="16/04/2026",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("kq_tiep_thi"),
    session=None,
):
    overrides = {
        "vngay_bd": str(start_date),
        "vngay_kt": str(end_date),
        "vdonvi_id": str(unit_id),
    }
    return download_with_recipe(
        "kq_tiep_thi_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="kq_tiep_thi report.xlsx",
        overrides=overrides,
        session=session,
    )


def download_report_vattu_thuhoi_api(
    start_date="24/11/2025",
    end_date="16/04/2026",
    unit_id="284656",
    service_ids="1,4,6,7,8,9,10,13,27,98,25,26,15,22,2,12,14,16",
    vat_tu_ids="1,2,3,4,8,6,5",
    headed=False,
    output_dir=group_output_dir("vat_tu_thu_hoi"),
    session=None,
):
    overrides = {
        "vttvt": str(unit_id),
        "vtungay": str(start_date),
        "vdenngay": str(end_date),
        "vdichvuvt_erp": str(service_ids),
        "vloaithu": "0",
        "vloaibatbuoc": "0",
        "vvattu": str(vat_tu_ids),
        "vloaingay": "1",
        "vtrangthai": "0",
    }
    return download_with_recipe(
        "vattu_thuhoi_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="bc_thu_hoi_vat_tu.xlsx",
        overrides=overrides,
        session=session,
    )


def download_cau_hinh_tu_dong_api(
    month_id="",
    month_label="",
    contract_type="1",
    output_name="cau_hinh_tu_dong report.xlsx",
    headed=False,
    output_dir=group_output_dir("cau_hinh_tu_dong"),
    session=None,
):
    overrides = {
        "pdv": str(contract_type),
    }
    return download_with_recipe(
        "cau_hinh_tu_dong_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name=output_name,
        overrides=overrides,
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_cau_hinh_tu_dong_ptm_api(
    month_id="",
    month_label="",
    headed=False,
    output_dir=group_output_dir("cau_hinh_tu_dong"),
    session=None,
):
    return download_cau_hinh_tu_dong_api(
        month_id=month_id,
        month_label=month_label,
        contract_type="1",
        output_name="cau_hinh_tu_dong_ptm.xlsx",
        headed=headed,
        output_dir=output_dir,
        session=session,
    )


def download_cau_hinh_tu_dong_thay_the_api(
    month_id="",
    month_label="",
    headed=False,
    output_dir=group_output_dir("cau_hinh_tu_dong"),
    session=None,
):
    return download_cau_hinh_tu_dong_api(
        month_id=month_id,
        month_label=month_label,
        contract_type="13",
        output_name="cau_hinh_tu_dong_thay_the.xlsx",
        headed=headed,
        output_dir=output_dir,
        session=session,
    )


def download_cau_hinh_tu_dong_chi_tiet_api(
    month_id="",
    month_label="",
    headed=False,
    output_dir=group_output_dir("cau_hinh_tu_dong"),
    session=None,
):
    return download_with_recipe(
        "cau_hinh_tu_dong_chi_tiet_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="cau_hinh_tu_dong_chi_tiet.xlsx",
        month_id=month_id,
        month_label=month_label,
        session=session,
    )


def download_quyet_toan_vattu_api(
    start_date="01/04/2026",
    end_date="16/04/2026",
    unit_id="284656",
    headed=False,
    output_dir=group_output_dir("vat_tu_thu_hoi"),
    session=None,
):
    overrides = {
        "vttvt": str(unit_id),
        "vtungay": str(start_date),
        "vdenngay": str(end_date),
    }
    return download_with_recipe(
        "quyet_toan_vattu_q2_2026",
        headed=headed,
        output_dir=output_dir,
        output_name="quyet_toan_vat_tu.xlsx",
        overrides=overrides,
        session=session,
    )
