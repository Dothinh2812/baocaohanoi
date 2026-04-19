# -*- coding: utf-8 -*-
"""Downloader generic cho BI report của OneBSS."""

import argparse
import inspect
import json
import sys
from pathlib import Path
from urllib.parse import urlencode

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from api_transition.onebss_auth import close_session, create_session
from api_transition.onebss_report_client import (
    build_run_v3_payload,
    fill_parameter_values,
    get_report_parameters,
    prepare_parameter_item,
    refresh_report_parameters,
    run_report_export,
    save_binary_export_file,
)


API_TRANSITION_DIR = Path(__file__).resolve().parent
DOWNLOADS_DIR = API_TRANSITION_DIR / "downloads"


def group_output_dir(group_name):
    return str(DOWNLOADS_DIR / group_name)


def extract_parameter_items(parameters_response):
    items = parameters_response.get("listOfParamNameValues", {}).get("item", [])
    if isinstance(items, dict):
        return [items]
    return list(items or [])


def build_refresh_items(parameter_items, refresh_values):
    refresh_map = refresh_values or {}
    refresh_items = []
    for item in parameter_items:
        name = item.get("name")
        if name not in refresh_map:
            continue
        refresh_items.append(prepare_parameter_item(item, refresh_map[name]))
    return refresh_items


def _save_playwright_download(download, output_dir, output_name):
    target_dir = Path(output_dir or group_output_dir("onebss"))
    target_dir.mkdir(parents=True, exist_ok=True)

    suggested_name = download.suggested_filename or "report.xlsx"
    target_name = output_name or suggested_name
    suggested_suffix = Path(suggested_name).suffix.lower()
    target_suffix = Path(target_name).suffix.lower()
    if suggested_suffix and target_suffix and suggested_suffix != target_suffix:
        target_name = str(Path(target_name).with_suffix(suggested_suffix))
    output_path = target_dir / target_name
    download.save_as(str(output_path))
    return output_path


def download_onebss_reportviewer_report(
    report_id,
    report_query_params,
    headed=False,
    output_dir="",
    output_name="",
    session=None,
    api_timeout=120,
):
    own_session = session is None

    try:
        if own_session:
            session = create_session(headed=headed)

        page = session.get("page")
        if page is None:
            raise RuntimeError("Session OneBSS không có page để thao tác ReportViewer.")

        token = session.get("token", "")
        if not token:
            raise RuntimeError("Không tìm thấy token OneBSS để mở ReportViewer.")

        bootstrap_url = f"https://report-onebss.vnpt.vn/?{urlencode({'baocao_id': str(report_id), 'token': token})}"
        print(f"Đang mở bootstrap ReportViewer: {bootstrap_url}")
        page.goto(bootstrap_url, wait_until="networkidle", timeout=api_timeout * 1000)

        report_url = f"https://report-onebss.vnpt.vn/ReportViewer.aspx?{urlencode(report_query_params)}"
        print(f"Đang render ReportViewer: {report_url}")
        page.goto(report_url, wait_until="networkidle", timeout=api_timeout * 1000)

        guid_input = page.locator("#ssrBaoCaoCachedReportNameInputId")
        guid_input.wait_for(state="attached", timeout=api_timeout * 1000)
        report_guid = guid_input.input_value().strip()
        if not report_guid:
            raise RuntimeError("Không lấy được guid của ReportViewer.")

        export_url = (
            "https://report-onebss.vnpt.vn/ReportViewer.aspx?"
            + urlencode(
                {
                    "guid": report_guid,
                    "action": "export",
                    "format": "excel",
                }
            )
        )
        print(f"Đang export ReportViewer: {export_url}")
        with page.expect_download(timeout=api_timeout * 1000) as download_info:
            page.evaluate("url => { window.location.href = url; }", export_url)
        download = download_info.value
        return _save_playwright_download(download, output_dir, output_name)
    finally:
        if own_session:
            close_session(session)


def download_onebss_report(
    report_path,
    baocao_id,
    overrides=None,
    refresh_values=None,
    headed=False,
    output_dir="",
    output_name="",
    session=None,
    api_timeout=120,
):
    own_session = session is None

    try:
        if own_session:
            session = create_session(headed=headed)

        headers = session["headers"]
        api_base_url = session.get("api_base_url", "")

        parameters_response = get_report_parameters(
            report_path,
            headers,
            api_base_url=api_base_url,
            timeout=api_timeout,
        )
        parameter_items = extract_parameter_items(parameters_response)

        refresh_items = build_refresh_items(parameter_items, refresh_values)
        if refresh_items:
            refreshed_response = refresh_report_parameters(
                report_path,
                refresh_items,
                headers,
                api_base_url=api_base_url,
                timeout=api_timeout,
            )
            parameter_items = extract_parameter_items(refreshed_response)

        filled_items = fill_parameter_values(parameter_items, overrides)
        payload = build_run_v3_payload(
            baocao_id=baocao_id,
            report_path=report_path,
            items=filled_items,
            file_name=output_name,
        )
        export_response = run_report_export(
            payload,
            headers,
            api_base_url=api_base_url,
            timeout=api_timeout,
        )

        target_dir = output_dir or group_output_dir("onebss")
        target_name = output_name or payload["file_name"]
        return save_binary_export_file(export_response, target_dir, target_name)

    finally:
        if own_session:
            close_session(session)


def download_hni_pttb_001(
    unit_id="14324",
    team_id="0",
    service_ids=None,
    congnghe_id="0",
    headed=False,
    output_dir="",
    output_name="",
    session=None,
):
    if service_ids is None:
        service_ids = ["1", "4", "7", "8", "11", "12"]

    return download_onebss_report(
        report_path="TINH/HANOI/HNI_PTTB_001/RP_HNI_PTTB_001",
        baocao_id=40618,
        refresh_values={
            "TT_ID": str(unit_id),
            "DOI_ID": str(team_id),
        },
        overrides={
            "TT_ID": str(unit_id),
            "DOI_ID": str(team_id),
            "DICHVUVT_ID": [str(value) for value in service_ids],
            "CONGNGHE_ID": str(congnghe_id),
        },
        headed=headed,
        output_dir=output_dir or group_output_dir("onebss"),
        output_name=output_name or "HNI_PTTB_001.xlsx",
        session=session,
    )


def download_bc_phieu_ton_dv_chi_tiet_hni(
    unit_id="14324",
    team_id="0",
    service_ids=None,
    congnghe_id="0",
    headed=False,
    output_dir="",
    output_name="",
    session=None,
):
    """Tải báo cáo phiếu tồn dịch vụ chi tiết HNI trên OneBSS."""
    if service_ids is None:
        service_ids = ["1", "4", "7", "8", "11", "12"]

    return download_onebss_report(
        report_path="TINH/HANOI/HNI_PTTB_001/RP_HNI_PTTB_001",
        baocao_id=40618,
        refresh_values={
            "TT_ID": str(unit_id),
            "DOI_ID": str(team_id),
        },
        overrides={
            "TT_ID": str(unit_id),
            "DOI_ID": str(team_id),
            "DICHVUVT_ID": [str(value) for value in service_ids],
            "CONGNGHE_ID": str(congnghe_id),
        },
        headed=headed,
        output_dir=output_dir or group_output_dir("onebss"),
        output_name=output_name or "bc_phieu_ton_dv_chi_tiet_hni.xlsx",
        session=session,
    )


def download_bc_ton_sua_chua_sontay_2026(
    unit_id="14324",
    team_id="0",
    service_id="0",
    headed=False,
    output_dir="",
    output_name="",
    session=None,
):
    """Tải báo cáo tồn sửa chữa Sơn Tây trên OneBSS."""
    return download_onebss_report(
        report_path="TINH/HANOI/HNI_BHSC_005/RP_HNI_BHSC_005",
        baocao_id=40622,
        refresh_values={
            "TT_ID": str(unit_id),
            "DOI_ID": str(team_id),
        },
        overrides={
            "TT_ID": str(unit_id),
            "DOI_ID": str(team_id),
            "DICHVUVT_ID": str(service_id),
        },
        headed=headed,
        output_dir=output_dir or group_output_dir("onebss"),
        output_name=output_name or "bc_ton_sua_chua_sontay_2026.xlsx",
        session=session,
    )


def download_bc_chi_tiet_ket_qua_cskh_uc3_sontay(
    customer_batch_code="UC3_CSKH_042026",
    start_date="01/04/2026",
    end_date="17/04/2026",
    unit_id="284656",
    employee_id="0",
    region_id="21",
    region_text="Thành phố Hà Nội",
    unit_text="Trung tâm Viễn thông Sơn Tây",
    employee_text="Tất cả",
    headed=False,
    output_dir="",
    output_name="",
    session=None,
):
    """Tải báo cáo chi tiết kết quả CSKH UC3 Sơn Tây qua ReportViewer."""
    normalized_batch_code = str(customer_batch_code)
    if not (normalized_batch_code.startswith("'") and normalized_batch_code.endswith("'")):
        normalized_batch_code = f"'{normalized_batch_code}'"

    return download_onebss_reportviewer_report(
        report_id="49544",
        report_query_params={
            "baocao_id": "49544",
            "vphanvung_id": str(region_id),
            "vphanvung_id_text": str(region_text),
            "vma_tap": normalized_batch_code,
            "vtungay": str(start_date),
            "vtungay_text": str(start_date),
            "vdenngay": str(end_date),
            "vdenngay_text": str(end_date),
            "vdonvi_id": str(unit_id),
            "vdonvi_id_text": str(unit_text),
            "vnhanvien_id": str(employee_id),
            "vnhanvien_id_text": str(employee_text),
        },
        headed=headed,
        output_dir=output_dir or group_output_dir("onebss"),
        output_name=output_name or "bc_chi_tiet_ket_qua_cskh_uc3_sontay.xls",
        session=session,
    )


def dump_payload_preview(
    report_path,
    baocao_id,
    overrides=None,
    refresh_values=None,
    headed=False,
    session=None,
    api_timeout=120,
):
    own_session = session is None

    try:
        if own_session:
            session = create_session(headed=headed)

        headers = session["headers"]
        api_base_url = session.get("api_base_url", "")
        parameters_response = get_report_parameters(
            report_path,
            headers,
            api_base_url=api_base_url,
            timeout=api_timeout,
        )
        parameter_items = extract_parameter_items(parameters_response)

        refresh_items = build_refresh_items(parameter_items, refresh_values)
        if refresh_items:
            parameters_response = refresh_report_parameters(
                report_path,
                refresh_items,
                headers,
                api_base_url=api_base_url,
                timeout=api_timeout,
            )
            parameter_items = extract_parameter_items(parameters_response)

        filled_items = fill_parameter_values(parameter_items, overrides)
        return build_run_v3_payload(
            baocao_id=baocao_id,
            report_path=report_path,
            items=filled_items,
        )
    finally:
        if own_session:
            close_session(session)


def payload_to_pretty_json(payload):
    return json.dumps(payload, ensure_ascii=False, indent=2)


CLI_DOWNLOADERS = dict(
    sorted(
        (
            name,
            obj,
        )
        for name, obj in globals().items()
        if callable(obj)
        and name.startswith("download_")
        and name not in {"download_onebss_report", "download_onebss_reportviewer_report"}
    )
)


def parse_args():
    parser = argparse.ArgumentParser(description="Chạy downloader OneBSS từ command line")
    parser.add_argument("report", nargs="?", help="Tên hàm downloader cần chạy")
    parser.add_argument(
        "--params",
        default="{}",
        help='JSON kwargs truyền vào downloader, ví dụ: \'{"unit_id":"14324","headed":true}\'',
    )
    parser.add_argument("--list", action="store_true", help="Liệt kê các downloader khả dụng")
    return parser.parse_args()


def load_cli_params(raw_params):
    try:
        params = json.loads(raw_params)
    except json.JSONDecodeError as exc:
        raise SystemExit(f"JSON không hợp lệ ở --params: {exc}") from exc

    if not isinstance(params, dict):
        raise SystemExit("--params phải là JSON object, ví dụ: '{\"unit_id\":\"14324\"}'")

    return params


def print_available_downloaders():
    print("Các downloader khả dụng:")
    for name, func in CLI_DOWNLOADERS.items():
        print(f"- {name}{inspect.signature(func)}")


def build_cli_kwargs(downloader, params, session=None):
    signature = inspect.signature(downloader)
    accepted_params = {
        name: value for name, value in params.items() if name in signature.parameters and name != "session"
    }
    if session is not None and "session" in signature.parameters:
        accepted_params["session"] = session
    return accepted_params


def run_named_downloader(name, params, session=None):
    downloader = CLI_DOWNLOADERS[name]
    kwargs = build_cli_kwargs(downloader, params, session=session)
    print(f"=== Chạy {name} ===")
    try:
        output_path = downloader(**kwargs)
    except TypeError as exc:
        signature = inspect.signature(downloader)
        raise SystemExit(f"Tham số không hợp lệ cho {name}{signature}: {exc}") from exc

    if output_path:
        print(f"✅ {name}: {output_path}")
    else:
        print(f"✅ {name}: hoàn tất")
    return output_path


def main():
    args = parse_args()

    if args.list:
        print_available_downloaders()
        return

    params = load_cli_params(args.params)

    if args.report and args.report not in CLI_DOWNLOADERS:
        print_available_downloaders()
        raise SystemExit(f"Không tìm thấy downloader: {args.report}")

    target_names = [args.report] if args.report else list(CLI_DOWNLOADERS)

    if len(target_names) == 1:
        run_named_downloader(target_names[0], params)
        return

    shared_session = None
    failures = []
    headed = bool(params.get("headed", False))

    try:
        print(f"Không truyền report, sẽ tải toàn bộ {len(target_names)} downloader.")
        shared_session = create_session(headed=headed)
        for name in target_names:
            try:
                run_named_downloader(name, params, session=shared_session)
            except Exception as exc:
                failures.append((name, str(exc)))
                print(f"❌ {name}: {exc}")
    finally:
        if shared_session is not None:
            close_session(shared_session)

    if failures:
        raise SystemExit(
            "Có downloader lỗi: "
            + "; ".join(f"{name} -> {message}" for name, message in failures)
        )


if __name__ == "__main__":
    main()
