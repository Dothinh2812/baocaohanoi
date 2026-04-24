# -*- coding: utf-8 -*-
"""Module batch download: login 1 lần, tải tuần tự tất cả báo cáo đã implemented.

Có thể chạy standalone:
    python3 api_transition/batch_download.py
    python3 api_transition/batch_download.py --month 5 --year 2026 --month-id 99001234

Hoặc import vào module main:
    from api_transition.batch_download import run_batch_download
    results = run_batch_download(report_month=4, report_year=2026, month_id="98944548")
"""

import argparse
import inspect
import sys
import time
import traceback
from calendar import monthrange
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from api_transition.auth import capture_authorization, login
from api_transition.runtime_config import RuntimeContext, load_runtime_context
from api_transition.report_api_client import build_report_page_url, make_common_headers
from api_transition.settings import Settings
from api_transition.cts_api import download_cts_gpon_quality_detail_api
from api_transition.downloaders import (
    download_report_c11_api,
    download_report_c12_api,
    download_report_c13_api,
    download_kpi_nvkt_c11_api,
    download_kpi_nvkt_c12_api,
    download_kpi_nvkt_c13_api,
    group_output_dir,
    download_report_c14_api,
    download_report_c14_chitiet_api,
    download_report_c15_api,
    download_report_c15_chitiet_api,
    download_report_c11_chitiet_api,
    download_report_c12_chitiet_sm1_api,
    download_report_c12_chitiet_sm2_api,
    download_report_i15_api,
    download_report_i15_k2_api,
    download_ghtt_report_hni_api,
    download_ghtt_report_sontay_api,
    download_ghtt_report_nvktdb_api,
    download_xac_minh_tam_dung_api,
    download_phieu_hoan_cong_dich_vu_chi_tiet_api,
    download_tam_dung_khoi_phuc_dich_vu_chi_tiet_api,
    download_tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc_api,
    download_tam_dung_khoi_phuc_dich_vu_tong_hop_api,
    download_ngung_psc_mytv_thang_t_1_cap_ttvt_api,
    download_ngung_psc_fiber_thang_t_1_cap_ttvt_api,
    download_ngung_psc_fiber_thang_t_1_cap_to_api,
    download_ngung_psc_mytv_thang_t_1_cap_to_api,
    download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api,
    download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api,
    download_kq_tiep_thi_api,
    download_report_vattu_thuhoi_api,
    download_cau_hinh_tu_dong_ptm_api,
    download_cau_hinh_tu_dong_thay_the_api,
    download_cau_hinh_tu_dong_chi_tiet_api,
)


# ===========================================================================
# CẤU HÌNH THAM SỐ ĐẦU VÀO — CHỈNH SỬA TẠI ĐÂY
# ===========================================================================
# Chỉ cần chỉnh REPORT_MONTH, REPORT_YEAR và MONTH_ID.
# Các ngày bắt đầu/kết thúc sẽ được tự động tính.
# ---------------------------------------------------------------------------

# --- (1) Tháng báo cáo — tâm của tất cả tính toán ---
REPORT_MONTH = 4                         # Tháng (1-12)
REPORT_YEAR = 2026                       # Năm

# --- (2) month_id — tra thủ công trên web, không tự tính được ---
MONTH_ID = "98944548"
MONTH_LABEL = ""                         # Fallback nếu không có month_id

# --- (3) Vật tư thu hồi — mốc bắt đầu cố định ---
VATTU_START_DATE = "01/04/2025"

# --- Chế độ hiển thị trình duyệt ---
HEADED = False                           # True = mở browser có giao diện

# --- Retry khi timeout ---
MAX_RETRIES = 3                          # Số lần thử tối đa (bao gồm lần đầu)
RETRY_TIMEOUTS = [180, 300, 500]         # Timeout (giây) cho lần 1, 2, 3
RETRY_DELAY = 3                          # Chờ (giây) giữa các lần retry

# ===========================================================================
# HẾT PHẦN CẤU HÌNH
# ===========================================================================


# ---------------------------------------------------------------------------
# Tự động tính ngày từ REPORT_MONTH / REPORT_YEAR
# ---------------------------------------------------------------------------

def _compute_dates(month: int, year: int):
    """Tính tất cả biến ngày từ tháng/năm báo cáo.

    Returns:
        tuple: (start_date, end_date, cal_start_date, cal_end_date, t_minus_1)
            - start_date     : kỳ BC lệch tháng (26 tháng trước → 25 tháng này)
            - end_date       : --
            - cal_start_date : tháng dương lịch (01 → cuối tháng)
            - cal_end_date   : --
            - t_minus_1      : ngày hôm qua (dd/mm/yyyy)
    """
    from datetime import date, timedelta

    # Kỳ báo cáo: ngày 26 tháng trước → ngày 25 tháng này
    prev_m = month - 1 if month > 1 else 12
    prev_y = year if month > 1 else year - 1
    start_date = f"26/{prev_m:02d}/{prev_y}"
    end_date   = f"25/{month:02d}/{year}"

    # Tháng dương lịch: ngày 01 → ngày cuối tháng
    last_day = monthrange(year, month)[1]
    cal_start_date = f"01/{month:02d}/{year}"
    cal_end_date   = f"{last_day:02d}/{month:02d}/{year}"

    # T-1: ngày hôm qua
    yesterday = date.today() - timedelta(days=1)
    t_minus_1 = yesterday.strftime("%d/%m/%Y")

    return start_date, end_date, cal_start_date, cal_end_date, t_minus_1


# Tính sẵn các biến ngày từ cấu hình
START_DATE, END_DATE, CAL_START_DATE, CAL_END_DATE, T_MINUS_1 = _compute_dates(REPORT_MONTH, REPORT_YEAR)


# ---------------------------------------------------------------------------
# Định nghĩa danh sách báo cáo
# ---------------------------------------------------------------------------

@dataclass
class ReportTask:
    """Mô tả 1 báo cáo cần tải."""
    name: str                      # Tên hiển thị
    report_key: str                # Key dùng trong config reports
    func: Callable                 # Hàm download
    params_type: str               # "month" | "date_range" | "calendar_month" | "t_minus_1" | "date_range_long"
    group: str                     # Nhóm nghiệp vụ
    id_family: Optional[str] = None
    extra_kwargs: Dict[str, Any] = field(default_factory=dict)
    use_shared_session: bool = True


REPORT_TASKS: List[ReportTask] = [
    # --- Nhóm Chỉ tiêu C (dùng month_id) ---
    ReportTask(
        name="C1.1",
        report_key="c11",
        func=download_report_c11_api,
        params_type="month",
        group="chi_tieu_c",
        id_family="center_id_14",
    ),
    ReportTask(
        name="C1.2",
        report_key="c12",
        func=download_report_c12_api,
        params_type="month",
        group="chi_tieu_c",
        id_family="center_id_14",
    ),
    ReportTask(
        name="C1.3",
        report_key="c13",
        func=download_report_c13_api,
        params_type="month",
        group="chi_tieu_c",
        id_family="center_id_14",
    ),
    # --- Nhóm KPI NVKT (dùng month_id) ---
    ReportTask(
        name="KPI NVKT C1.1",
        report_key="kpi_nvkt_c11",
        func=download_kpi_nvkt_c11_api,
        params_type="month",
        group="kpi_nvkt",
        id_family="center_id_14",
        extra_kwargs={"output_dir": group_output_dir("kpi_nvkt")},
    ),
    ReportTask(
        name="KPI NVKT C1.2",
        report_key="kpi_nvkt_c12",
        func=download_kpi_nvkt_c12_api,
        params_type="month",
        group="kpi_nvkt",
        id_family="center_id_14",
        extra_kwargs={"output_dir": group_output_dir("kpi_nvkt")},
    ),
    ReportTask(
        name="KPI NVKT C1.3",
        report_key="kpi_nvkt_c13",
        func=download_kpi_nvkt_c13_api,
        params_type="month",
        group="kpi_nvkt",
        id_family="center_id_14",
        extra_kwargs={"output_dir": group_output_dir("kpi_nvkt")},
    ),
    ReportTask(
        name="C1.4",
        report_key="c14",
        func=download_report_c14_api,
        params_type="month",
        group="chi_tieu_c",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="C1.4 Chi tiết",
        report_key="c14_chi_tiet",
        func=download_report_c14_chitiet_api,
        params_type="month",
        group="chi_tieu_c",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="C1.5",
        report_key="c15",
        func=download_report_c15_api,
        params_type="month",
        group="chi_tieu_c",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="C1.5 Chi tiết",
        report_key="c15_chitiet",
        func=download_report_c15_chitiet_api,
        params_type="month",
        group="chi_tieu_c",
        id_family="unit_id_28",
    ),
    # --- Nhóm Chỉ tiêu C chi tiết (dùng date_range) ---
    ReportTask(
        name="C1.1 Chi tiết",
        report_key="c11_chi_tiet",
        func=download_report_c11_chitiet_api,
        params_type="date_range",
        group="chi_tieu_c",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="C1.2 Chi tiết SM1",
        report_key="c12_chi_tiet_sm1",
        func=download_report_c12_chitiet_sm1_api,
        params_type="date_range",
        group="chi_tieu_c",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="C1.2 Chi tiết SM2",
        report_key="c12_chi_tiet_sm2",
        func=download_report_c12_chitiet_sm2_api,
        params_type="date_range",
        group="chi_tieu_c",
        id_family="unit_id_28",
    ),
    # --- Nhóm Chỉ tiêu I (dùng T-1 = ngày hôm qua) ---
    ReportTask(
        name="I1.5",
        report_key="i15",
        func=download_report_i15_api,
        params_type="t_minus_1",
        group="chi_tieu_i",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="I1.5 K2",
        report_key="i15_k2",
        func=download_report_i15_k2_api,
        params_type="t_minus_1",
        group="chi_tieu_i",
        id_family="unit_id_28",
    ),
    # --- Nhóm GHTT (dùng month_id) ---
    ReportTask(
        name="GHTT HNI",
        report_key="ghtt_hni",
        func=download_ghtt_report_hni_api,
        params_type="month",
        group="ghtt",
    ),
    ReportTask(
        name="GHTT Sơn Tây",
        report_key="ghtt_sontay",
        func=download_ghtt_report_sontay_api,
        params_type="month",
        group="ghtt",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="GHTT NVKT DB",
        report_key="ghtt_nvktdb",
        func=download_ghtt_report_nvktdb_api,
        params_type="month",
        group="ghtt",
        id_family="unit_id_28",
    ),
    # --- Nhóm khác (dùng tháng dương lịch) ---
    ReportTask(
        name="Xác minh tạm dừng",
        report_key="xac_minh_tam_dung",
        func=download_xac_minh_tam_dung_api,
        params_type="calendar_month",
        group="xac_minh_tam_dung",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="Phiếu hoàn công dịch vụ chi tiết",
        report_key="phieu_hoan_cong_dich_vu_chi_tiet",
        func=download_phieu_hoan_cong_dich_vu_chi_tiet_api,
        params_type="calendar_month",
        group="phieu_hoan_cong_dich_vu",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="Tạm dừng, khôi phục dịch vụ chi tiết",
        report_key="tam_dung_khoi_phuc_dich_vu_chi_tiet",
        func=download_tam_dung_khoi_phuc_dich_vu_chi_tiet_api,
        params_type="calendar_month",
        group="tam_dung_khoi_phuc_dich_vu",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="Tạm dừng, khôi phục dịch vụ chi tiết - khôi phục",
        report_key="tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc",
        func=download_tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc_api,
        params_type="calendar_month",
        group="tam_dung_khoi_phuc_dich_vu",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="Tạm dừng, khôi phục dịch vụ tổng hợp",
        report_key="tam_dung_khoi_phuc_dich_vu_tong_hop",
        func=download_tam_dung_khoi_phuc_dich_vu_tong_hop_api,
        params_type="calendar_month",
        group="tam_dung_khoi_phuc_dich_vu",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="Ngưng PSC MyTV tháng T-1 cấp TTVT",
        report_key="ngung_psc_mytv_thang_t_1_cap_ttvt",
        func=download_ngung_psc_mytv_thang_t_1_cap_ttvt_api,
        params_type="t_minus_1",
        group="tam_dung_khoi_phuc_dich_vu",
        id_family="service_scope_ttvt_id",
        extra_kwargs={"t_minus_1_as_report_date": True},
    ),
    ReportTask(
        name="Ngưng PSC Fiber tháng T-1 cấp TTVT",
        report_key="ngung_psc_fiber_thang_t_1_cap_ttvt",
        func=download_ngung_psc_fiber_thang_t_1_cap_ttvt_api,
        params_type="t_minus_1",
        group="tam_dung_khoi_phuc_dich_vu",
        id_family="service_scope_ttvt_id",
        extra_kwargs={"t_minus_1_as_report_date": True},
    ),
    ReportTask(
        name="Ngưng PSC Fiber tháng T-1 cấp Tổ",
        report_key="ngung_psc_fiber_thang_t_1_cap_to",
        func=download_ngung_psc_fiber_thang_t_1_cap_to_api,
        params_type="t_minus_1",
        group="tam_dung_khoi_phuc_dich_vu",
        id_family="service_scope_team_id",
        extra_kwargs={"t_minus_1_as_report_date": True},
    ),
    ReportTask(
        name="Ngưng PSC MyTV tháng T-1 cấp Tổ",
        report_key="ngung_psc_mytv_thang_t_1_cap_to",
        func=download_ngung_psc_mytv_thang_t_1_cap_to_api,
        params_type="t_minus_1",
        group="tam_dung_khoi_phuc_dich_vu",
        id_family="service_scope_team_id",
        extra_kwargs={"t_minus_1_as_report_date": True},
    ),
    ReportTask(
        name="Tỷ lệ xác minh đúng thời gian quy định - TTVTKV",
        report_key="ty_le_xac_minh_ttvtkv",
        func=download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api,
        params_type="month",
        group="ty_le_xac_minh",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="Tỷ lệ xác minh đúng thời gian quy định chi tiết",
        report_key="ty_le_xac_minh_chi_tiet",
        func=download_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api,
        params_type="month",
        group="ty_le_xac_minh",
        id_family="unit_id_28",
    ),
    ReportTask(
        name="Kết quả tiếp thị",
        report_key="kq_tiep_thi",
        func=download_kq_tiep_thi_api,
        params_type="calendar_month",
        group="kq_tiep_thi",
        id_family="unit_id_28",
    ),
    # ReportTask(
    #     name="CTS SHC ngày",
    #     func=download_cts_gpon_quality_detail_api,
    #     params_type="t_minus_1",
    #     group="cts",
    #     extra_kwargs={"t_minus_1_as_report_date": True},
    #     use_shared_session=False,
    # ),
    # --- Nhóm Vật tư thu hồi (date_range dài hạn) ---
    ReportTask(
        name="Vật tư thu hồi",
        report_key="vattu_thu_hoi",
        func=download_report_vattu_thuhoi_api,
        params_type="date_range_long",
        group="vat_tu_thu_hoi",
        id_family="unit_id_28",
    ),
    # --- Nhóm Cấu hình tự động (dùng month_id) ---
    ReportTask(
        name="Cấu hình tự động PTM",
        report_key="cau_hinh_tu_dong_ptm",
        func=download_cau_hinh_tu_dong_ptm_api,
        params_type="month",
        group="cau_hinh_tu_dong",
    ),
    ReportTask(
        name="Cấu hình tự động Thay thế",
        report_key="cau_hinh_tu_dong_thay_the",
        func=download_cau_hinh_tu_dong_thay_the_api,
        params_type="month",
        group="cau_hinh_tu_dong",
    ),
    ReportTask(
        name="Cấu hình tự động Chi tiết",
        report_key="cau_hinh_tu_dong_chi_tiet",
        func=download_cau_hinh_tu_dong_chi_tiet_api,
        params_type="month",
        group="cau_hinh_tu_dong",
    ),
]


# ---------------------------------------------------------------------------
# Session management
# ---------------------------------------------------------------------------

# URL bất kỳ của 1 report đã implemented để capture Authorization header
_DEFAULT_AUTH_REPORT_URL = "https://baocao.hanoi.vnpt.vn/report/report-info?id=534964&menu_id=535020"


def create_session(headed=False, auth_report_url=""):
    """Login và capture Authorization, trả về session dict.

    Returns:
        dict:
            - "headers": dict headers cho mọi request API
            - "playwright": playwright instance
            - "browser": browser instance
            - "context": browser context
            - "page": page instance
    """
    Settings.validate()
    report_url = auth_report_url or _DEFAULT_AUTH_REPORT_URL

    playwright, browser, context, page = login(headless=not headed)
    auth_state = capture_authorization(page, report_url)
    headers = make_common_headers(auth_state, context.cookies())

    return {
        "headers": headers,
        "playwright": playwright,
        "browser": browser,
        "context": context,
        "page": page,
    }


def close_session(session):
    """Đóng browser và playwright trong session."""
    if session is None:
        return
    browser = session.get("browser")
    playwright = session.get("playwright")
    if browser is not None:
        try:
            browser.close()
        except Exception:
            pass
    if playwright is not None:
        try:
            playwright.stop()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Batch download
# ---------------------------------------------------------------------------

@dataclass
class DownloadResult:
    """Kết quả tải 1 báo cáo."""
    name: str
    group: str
    status: str              # "success" | "failed" | "skipped"
    output_path: str = ""
    error: str = ""
    duration_seconds: float = 0.0


_REPORT_CONFIG_RESERVED_KEYS = {"enabled", "note"}


def _task_output_dir(task: ReportTask, runtime_context: Optional[RuntimeContext]) -> Optional[str]:
    if runtime_context is None:
        return None
    return str(runtime_context.download_group_dir(task.group))


def _task_unit_id(task: ReportTask, runtime_context: Optional[RuntimeContext]) -> Optional[str]:
    if runtime_context is None or not task.id_family:
        return None
    unit_id = runtime_context.unit.ids.get(task.id_family, "")
    if not unit_id:
        raise ValueError(
            f"Report '{task.report_key}' dang duoc bat nhung config don vi thieu id '{task.id_family}'"
        )
    return unit_id


def _report_override_kwargs(task: ReportTask, runtime_context: Optional[RuntimeContext]) -> Dict[str, Any]:
    if runtime_context is None:
        return {}

    report_config = runtime_context.get_report_config(task.report_key)
    if not report_config:
        return {}

    signature = inspect.signature(task.func)
    supported_params = set(signature.parameters)
    overrides: Dict[str, Any] = {}
    for key, value in report_config.items():
        if key in _REPORT_CONFIG_RESERVED_KEYS:
            continue
        if key in supported_params:
            overrides[key] = value
    return overrides


def _build_kwargs(task: ReportTask, params: dict, runtime_context: Optional[RuntimeContext] = None) -> dict:
    """Xây dựng kwargs phù hợp cho từng loại báo cáo."""
    kwargs = dict(task.extra_kwargs)

    if task.params_type == "month":
        if params.get("month_id"):
            kwargs["month_id"] = params["month_id"]
        if params.get("month_label"):
            kwargs["month_label"] = params["month_label"]

    elif task.params_type == "date_range":
        # Kỳ BC lệch tháng: 26 tháng trước → 25 tháng này
        if params.get("start_date"):
            kwargs["start_date"] = params["start_date"]
        if params.get("end_date"):
            kwargs["end_date"] = params["end_date"]

    elif task.params_type == "calendar_month":
        # Tháng dương lịch: 01 → cuối tháng
        if params.get("cal_start_date"):
            kwargs["start_date"] = params["cal_start_date"]
        if params.get("cal_end_date"):
            kwargs["end_date"] = params["cal_end_date"]

    elif task.params_type == "t_minus_1":
        # Ngày hôm qua (start = end = T-1)
        if params.get("t_minus_1"):
            if kwargs.pop("t_minus_1_as_report_date", False):
                kwargs["report_date"] = params["t_minus_1"]
            else:
                kwargs["start_date"] = params["t_minus_1"]
                kwargs["end_date"] = params["t_minus_1"]

    elif task.params_type == "date_range_long":
        # Vật tư thu hồi có ngày bắt đầu riêng
        if params.get("vattu_start_date"):
            kwargs["start_date"] = params["vattu_start_date"]
        elif params.get("start_date"):
            kwargs["start_date"] = params["start_date"]
        if params.get("end_date"):
            kwargs["end_date"] = params["end_date"]

    output_dir = _task_output_dir(task, runtime_context)
    if output_dir:
        kwargs["output_dir"] = output_dir

    unit_id = _task_unit_id(task, runtime_context)
    if unit_id:
        kwargs["unit_id"] = unit_id

    kwargs.update(_report_override_kwargs(task, runtime_context))
    return kwargs


def run_batch_download(
    report_month=None,
    report_year=None,
    month_id=None,
    month_label=None,
    vattu_start_date=None,
    headed=None,
    skip_reports: Optional[List[str]] = None,
    only_reports: Optional[List[str]] = None,
    session=None,
    config_path: Optional[str] = None,
    runtime_context: Optional[RuntimeContext] = None,
) -> Dict[str, Any]:
    """Login 1 lần, tải tuần tự tất cả báo cáo đã implemented.

    Nếu tham số là None, sẽ dùng giá trị trong phần CẤU HÌNH ở đầu file.

    Args:
        report_month: Tháng báo cáo (1-12).
        report_year: Năm báo cáo.
        month_id: ID kỳ báo cáo tháng (ví dụ "98944548").
        month_label: Nhãn kỳ báo cáo (ví dụ "Tháng 04/2026").
        vattu_start_date: Ngày bắt đầu riêng cho Vật tư thu hồi.
        headed: Mở trình duyệt có giao diện hay không.
        skip_reports: Danh sách tên report muốn bỏ qua.
        only_reports: Chỉ chạy các report có tên trong list này.
        session: Session dict đã tạo trước (nếu None sẽ tự login).

    Returns:
        dict với keys "success", "failed", "skipped" — mỗi key là list[DownloadResult].
    """
    if config_path and runtime_context is not None:
        raise ValueError("Chi duoc truyen mot trong hai tham so: config_path hoac runtime_context")

    if config_path:
        runtime_context = load_runtime_context(config_path)

    # Áp dụng cấu hình mặc định nếu tham số là None
    eff_month = report_month if report_month is not None else (
        runtime_context.period.report_month if runtime_context is not None else REPORT_MONTH
    )
    eff_year = report_year if report_year is not None else (
        runtime_context.period.report_year if runtime_context is not None else REPORT_YEAR
    )
    eff_month_id = month_id if month_id is not None else (
        runtime_context.period.month_id if runtime_context is not None else MONTH_ID
    )
    eff_month_label = month_label if month_label is not None else (
        runtime_context.period.month_label if runtime_context is not None else MONTH_LABEL
    )
    eff_vattu_start = vattu_start_date if vattu_start_date is not None else (
        runtime_context.period.vattu_start_date if runtime_context is not None else VATTU_START_DATE
    )
    eff_headed = headed if headed is not None else (
        runtime_context.download.headed if runtime_context is not None else HEADED
    )
    eff_max_retries = runtime_context.download.max_retries if runtime_context is not None else MAX_RETRIES
    eff_retry_timeouts = list(runtime_context.download.retry_timeouts) if runtime_context is not None else list(RETRY_TIMEOUTS)
    eff_retry_delay = runtime_context.download.retry_delay_seconds if runtime_context is not None else RETRY_DELAY

    # Tự tính các biến ngày từ tháng/năm
    start_date, end_date, cal_start_date, cal_end_date, t_minus_1 = _compute_dates(eff_month, eff_year)

    skip_set = set(skip_reports or [])
    only_set = set(only_reports or [])
    own_session = session is None

    results = {
        "success": [],
        "failed": [],
        "skipped": [],
    }

    params = {
        "month_id": eff_month_id,
        "month_label": eff_month_label,
        "start_date": start_date,
        "end_date": end_date,
        "cal_start_date": cal_start_date,
        "cal_end_date": cal_end_date,
        "t_minus_1": t_minus_1,
        "vattu_start_date": eff_vattu_start,
    }

    # In cấu hình đang dùng
    print("\n" + "=" * 70)
    print(f"CẤU HÌNH BATCH DOWNLOAD — Tháng {eff_month:02d}/{eff_year}")
    print("=" * 70)
    print(f"  MONTH_ID         : {eff_month_id or '(trống)'}")
    print(f"  Kỳ BC (26→25)    : {start_date} → {end_date}")
    print(f"  Tháng DL (01→cc) : {cal_start_date} → {cal_end_date}")
    print(f"  T-1 (hôm qua)    : {t_minus_1}")
    print(f"  VATTU_START_DATE : {eff_vattu_start or '(trống)'}")
    print(f"  HEADED           : {eff_headed}")
    print(f"  MAX_RETRIES      : {eff_max_retries}")
    print(f"  RETRY_TIMEOUTS   : {eff_retry_timeouts}")
    print(f"  RETRY_DELAY      : {eff_retry_delay}s")
    if runtime_context is not None:
        print(f"  UNIT_CODE        : {runtime_context.unit.code}")
        print(f"  UNIT_NAME        : {runtime_context.unit.name}")
        print(f"  INSTANCE_ROOT    : {runtime_context.paths.instance_root}")
    if only_set:
        print(f"  CHỈ CHẠY         : {', '.join(sorted(only_set))}")
    if skip_set:
        print(f"  BỎ QUA           : {', '.join(sorted(skip_set))}")
    print("=" * 70)

    try:
        def _task_should_run(task: ReportTask) -> bool:
            if only_set and task.name not in only_set:
                return False
            if task.name in skip_set:
                return False
            if runtime_context is not None and not runtime_context.is_report_enabled(task.report_key):
                return False
            return True

        # --- 1. Tạo session nếu chưa có ---
        needs_shared_session = any(
            task.use_shared_session and _task_should_run(task)
            for task in REPORT_TASKS
        )

        if own_session and needs_shared_session:
            print("\nĐang đăng nhập...")
            session = create_session(headed=eff_headed)
            print("✅ Đăng nhập và capture Authorization thành công.\n")

        # --- 2. Duyệt từng báo cáo ---
        total = len(REPORT_TASKS)
        for index, task in enumerate(REPORT_TASKS, start=1):

            # Kiểm tra skip / only
            if only_set and task.name not in only_set:
                results["skipped"].append(
                    DownloadResult(name=task.name, group=task.group, status="skipped")
                )
                continue
            if task.name in skip_set:
                results["skipped"].append(
                    DownloadResult(name=task.name, group=task.group, status="skipped")
                )
                print(f"[{index}/{total}] ⏭️  Bỏ qua: {task.name}")
                continue
            if runtime_context is not None and not runtime_context.is_report_enabled(task.report_key):
                results["skipped"].append(
                    DownloadResult(
                        name=task.name,
                        group=task.group,
                        status="skipped",
                        error="Disabled by config",
                    )
                )
                print(f"[{index}/{total}] ⏭️  Tắt theo config: {task.name}")
                continue

            print(f"[{index}/{total}] 📥 Đang tải: {task.name} (nhóm: {task.group})")
            started = time.time()
            last_error = None
            succeeded = False

            for attempt in range(eff_max_retries):
                attempt_timeout = eff_retry_timeouts[attempt] if attempt < len(eff_retry_timeouts) else eff_retry_timeouts[-1]

                if attempt > 0:
                    print(f"         🔄 Retry lần {attempt + 1}/{eff_max_retries} (timeout={attempt_timeout}s, chờ {eff_retry_delay}s...)")
                    time.sleep(eff_retry_delay)

                try:
                    kwargs = _build_kwargs(task, params, runtime_context=runtime_context)
                    if task.use_shared_session:
                        kwargs["session"] = session
                    kwargs["headed"] = eff_headed

                    # Đặt timeout cho lần thử này
                    if task.use_shared_session:
                        session["api_timeout"] = attempt_timeout

                    output_path = task.func(**kwargs)
                    elapsed = time.time() - started

                    results["success"].append(
                        DownloadResult(
                            name=task.name,
                            group=task.group,
                            status="success",
                            output_path=str(output_path),
                            duration_seconds=round(elapsed, 1),
                        )
                    )
                    print(f"         ✅ Thành công → {output_path} ({elapsed:.1f}s)")
                    succeeded = True
                    break

                except (TimeoutError, OSError) as exc:
                    last_error = exc
                    print(f"         ⏱️  Timeout lần {attempt + 1}: {type(exc).__name__}: {exc}")

                except Exception as exc:
                    # Lỗi không phải timeout → không retry
                    last_error = exc
                    print(f"         ❌ Lỗi: {type(exc).__name__}: {exc}")
                    traceback.print_exc()
                    break

            if not succeeded:
                elapsed = time.time() - started
                error_msg = f"{type(last_error).__name__}: {last_error}"
                results["failed"].append(
                    DownloadResult(
                        name=task.name,
                        group=task.group,
                        status="failed",
                        error=error_msg,
                        duration_seconds=round(elapsed, 1),
                    )
                )
                print(f"         ❌ Thất bại sau {eff_max_retries} lần thử: {error_msg}")

    finally:
        if own_session:
            close_session(session)

    # --- 3. In bảng tổng kết ---
    _print_summary(results)
    return results


def _print_summary(results: Dict[str, Any]):
    """In bảng tổng kết kết quả batch download."""
    success = results["success"]
    failed = results["failed"]
    skipped = results["skipped"]
    total = len(success) + len(failed) + len(skipped)

    print("\n" + "=" * 70)
    print("TỔNG KẾT BATCH DOWNLOAD")
    print("=" * 70)
    print(f"  Tổng số báo cáo : {total}")
    print(f"  ✅ Thành công    : {len(success)}")
    print(f"  ❌ Thất bại      : {len(failed)}")
    print(f"  ⏭️  Bỏ qua       : {len(skipped)}")

    if success:
        total_time = sum(r.duration_seconds for r in success)
        print(f"\n  Tổng thời gian tải: {total_time:.1f}s")
        print("\n  📁 Danh sách file đã tải:")
        for r in success:
            print(f"     • [{r.group}] {r.name} → {r.output_path}")

    if failed:
        print("\n  ⚠️  Danh sách báo cáo bị lỗi:")
        for r in failed:
            print(f"     • {r.name}: {r.error}")

    print("=" * 70 + "\n")


# ---------------------------------------------------------------------------
# Standalone CLI
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Batch download: login 1 lần, tải tuần tự tất cả báo cáo.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Ví dụ:
  # Tải tất cả (dùng cấu hình mặc định trong file)
  python3 api_transition/batch_download.py

  # Override tháng/năm từ CLI
  python3 api_transition/batch_download.py --month 5 --year 2026 --month-id 99001234

  # Chỉ tải 2 báo cáo cụ thể
  python3 api_transition/batch_download.py --only "C1.1" "C1.2"

  # Tải tất cả trừ vật tư thu hồi
  python3 api_transition/batch_download.py --skip "Vật tư thu hồi"

  # Liệt kê tên tất cả báo cáo có thể tải
  python3 api_transition/batch_download.py --list
""",
    )
    parser.add_argument("--month", type=int, default=None, help=f"Tháng báo cáo 1-12 (mặc định: {REPORT_MONTH})")
    parser.add_argument("--year", type=int, default=None, help=f"Năm báo cáo (mặc định: {REPORT_YEAR})")
    parser.add_argument("--month-id", default=None, help=f"ID kỳ tháng (mặc định: {MONTH_ID!r})")
    parser.add_argument("--month-label", default=None, help=f"Nhãn kỳ báo cáo (mặc định: {MONTH_LABEL!r})")
    parser.add_argument("--vattu-start-date", default=None, help=f"Mốc bắt đầu Vật tư thu hồi (mặc định: {VATTU_START_DATE!r})")
    parser.add_argument("--config", default=None, help="Đường dẫn file YAML config đơn vị")
    parser.add_argument("--headed", action="store_true", default=None, help="Mở trình duyệt có giao diện")
    parser.add_argument("--only", nargs="+", default=[], metavar="NAME", help="Chỉ chạy các report có tên này")
    parser.add_argument("--skip", nargs="+", default=[], metavar="NAME", help="Bỏ qua các report có tên này")
    parser.add_argument("--list", action="store_true", help="Liệt kê tên tất cả báo cáo rồi thoát")
    return parser.parse_args()


def main():
    args = parse_args()
    runtime_context = load_runtime_context(args.config) if args.config else None

    if args.list:
        eff_m = args.month if args.month is not None else (
            runtime_context.period.report_month if runtime_context is not None else REPORT_MONTH
        )
        eff_y = args.year if args.year is not None else (
            runtime_context.period.report_year if runtime_context is not None else REPORT_YEAR
        )
        sd, ed, csd, ced, tm1 = _compute_dates(eff_m, eff_y)

        print(f"\nDanh sách báo cáo — Tháng {eff_m:02d}/{eff_y}:\n")
        for i, task in enumerate(REPORT_TASKS, start=1):
            enabled_label = ""
            if runtime_context is not None:
                enabled_label = " enabled" if runtime_context.is_report_enabled(task.report_key) else " disabled"
            print(f"  {i:>2}. [{task.params_type:<16}] [{task.group:<20}] [{task.report_key:<38}] {task.name}{enabled_label}")
        print(f"\nTổng: {len(REPORT_TASKS)} báo cáo")
        print(f"\nCấu hình (tự tính từ tháng {eff_m:02d}/{eff_y}):")
        print(f"  MONTH_ID         = {(args.month_id if args.month_id is not None else (runtime_context.period.month_id if runtime_context is not None else MONTH_ID))!r}")
        print(f"  Kỳ BC (26→25)    = {sd!r} → {ed!r}")
        print(f"  Tháng DL (01→cc) = {csd!r} → {ced!r}")
        print(f"  T-1 (hôm qua)    = {tm1!r}")
        print(f"  VATTU_START_DATE = {(args.vattu_start_date if args.vattu_start_date is not None else (runtime_context.period.vattu_start_date if runtime_context is not None else VATTU_START_DATE))!r}")
        if runtime_context is not None:
            print(f"  CONFIG           = {runtime_context.config_path}")
            print(f"  UNIT_CODE        = {runtime_context.unit.code!r}")
            print(f"  INSTANCE_ROOT    = {str(runtime_context.paths.instance_root)!r}")
        print()
        return

    results = run_batch_download(
        report_month=args.month,
        report_year=args.year,
        month_id=args.month_id,
        month_label=args.month_label,
        vattu_start_date=args.vattu_start_date,
        headed=args.headed,
        skip_reports=args.skip,
        only_reports=args.only,
        runtime_context=runtime_context,
    )

    sys.exit(0 if not results["failed"] else 1)


if __name__ == "__main__":
    main()
