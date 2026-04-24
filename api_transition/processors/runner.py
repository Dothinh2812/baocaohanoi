"""Runner gom va dieu phoi cac processor trong api_transition.processors."""

from __future__ import annotations

import argparse
import inspect
import time
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from api_transition.runtime_config import RuntimeContext, load_runtime_context
from api_transition.processors.c_processors import (
    process_c11_chitiet_report_api_output,
    process_c11_report_api_output,
    process_c12_chitiet_reports_api_output,
    process_c12_report_api_output,
    process_c13_report_api_output,
    process_c14_chitiet_report_api_output,
    process_c14_report_api_output,
    process_c15_chitiet_report_api_output,
    process_c15_report_api_output,
)
from api_transition.processors.i15_processors import (
    process_i15_k2_report_api_output,
    process_i15_report_api_output,
)
from api_transition.processors.cau_hinh_tu_dong_processors import (
    process_cau_hinh_tu_dong_chi_tiet_api_output,
    process_cau_hinh_tu_dong_ptm_api_output,
    process_cau_hinh_tu_dong_thay_the_api_output,
)
from api_transition.processors.ghtt_processors import (
    process_ghtt_hni_api_output,
    process_ghtt_nvktdb_api_output,
    process_ghtt_sontay_api_output,
)
from api_transition.processors.kpi_processors import (
    process_kpi_nvkt_c11_api_output,
    process_kpi_nvkt_c12_api_output,
    process_kpi_nvkt_c13_api_output,
)
from api_transition.processors.kq_tiep_thi_processors import (
    process_kq_tiep_thi_api_output,
)
from api_transition.processors.service_flow_processors import (
    process_fiber_thuc_tang_api_output,
    process_mytv_hoan_cong_api_output,
    process_mytv_ngung_psc_api_output,
    process_mytv_ngung_psc_ttvt_api_output,
    process_mytv_thuc_tang_api_output,
    process_phieu_hoan_cong_dich_vu_chi_tiet_api_output,
    process_son_tay_fiber_ngung_psc_t_minus_1_api_output,
    process_son_tay_fiber_ngung_psc_t_minus_1_ttvt_api_output,
    process_son_tay_mytv_ngung_psc_t_minus_1_api_output,
    process_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_api_output,
    process_tam_dung_khoi_phuc_dich_vu_tong_hop_api_output,
)
from api_transition.processors.vattu_processors import (
    process_vat_tu_thu_hoi_api_output,
)
from api_transition.processors.verification_processors import (
    process_xac_minh_tam_dung_api_output,
    process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api_output,
    process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api_output,
)
from api_transition.processors.common import configure_runtime_roots, reset_runtime_roots


API_TRANSITION_DIR = Path(__file__).resolve().parent.parent
DEFAULT_DSNV_FILE = API_TRANSITION_DIR / "dsnv.xlsx"
DEFAULT_DANHBA_DB_FILE = API_TRANSITION_DIR.parent / "danhba.db"


@dataclass(frozen=True)
class ProcessorTask:
    """Mo ta 1 processor co the chay doc lap trong batch runner."""

    name: str
    group: str
    func: Callable[..., Any]
    source_report_keys: Tuple[str, ...] = ()


@dataclass
class ProcessorRunResult:
    """Ket qua chay 1 processor."""

    name: str
    group: str
    status: str
    duration_seconds: float = 0.0
    result: Any = None
    error: str = ""


PROCESSOR_TASKS: Tuple[ProcessorTask, ...] = (
    ProcessorTask("c11", "chi_tieu_c", process_c11_report_api_output, ("c11",)),
    ProcessorTask("c12", "chi_tieu_c", process_c12_report_api_output, ("c12",)),
    ProcessorTask("c13", "chi_tieu_c", process_c13_report_api_output, ("c13",)),
    ProcessorTask("c14", "chi_tieu_c", process_c14_report_api_output, ("c14",)),
    ProcessorTask("c15", "chi_tieu_c", process_c15_report_api_output, ("c15",)),
    ProcessorTask("c15_chitiet", "chi_tieu_c", process_c15_chitiet_report_api_output, ("c15_chitiet",)),
    ProcessorTask("i15", "chi_tieu_i", process_i15_report_api_output, ("i15",)),
    ProcessorTask("i15_k2", "chi_tieu_i", process_i15_k2_report_api_output, ("i15_k2",)),
    ProcessorTask("c14_chi_tiet", "chi_tieu_c", process_c14_chitiet_report_api_output, ("c14_chi_tiet",)),
    ProcessorTask("c11_chi_tiet", "chi_tieu_c", process_c11_chitiet_report_api_output, ("c11_chi_tiet",)),
    ProcessorTask("c12_chi_tiet", "chi_tieu_c", process_c12_chitiet_reports_api_output, ("c12_chi_tiet_sm1", "c12_chi_tiet_sm2")),
    ProcessorTask("kpi_nvkt_c11", "kpi_nvkt", process_kpi_nvkt_c11_api_output, ("kpi_nvkt_c11",)),
    ProcessorTask("kpi_nvkt_c12", "kpi_nvkt", process_kpi_nvkt_c12_api_output, ("kpi_nvkt_c12",)),
    ProcessorTask("kpi_nvkt_c13", "kpi_nvkt", process_kpi_nvkt_c13_api_output, ("kpi_nvkt_c13",)),
    ProcessorTask("kq_tiep_thi", "kq_tiep_thi", process_kq_tiep_thi_api_output, ("kq_tiep_thi",)),
    ProcessorTask("ghtt_hni", "ghtt", process_ghtt_hni_api_output, ("ghtt_hni",)),
    ProcessorTask("ghtt_son_tay", "ghtt", process_ghtt_sontay_api_output, ("ghtt_sontay",)),
    ProcessorTask("ghtt_nvktdb", "ghtt", process_ghtt_nvktdb_api_output, ("ghtt_nvktdb",)),
    ProcessorTask("xac_minh_tam_dung", "xac_minh_tam_dung", process_xac_minh_tam_dung_api_output, ("xac_minh_tam_dung",)),
    ProcessorTask("xac_minh_ttvtkv", "ty_le_xac_minh", process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api_output, ("ty_le_xac_minh_ttvtkv",)),
    ProcessorTask("xac_minh_chi_tiet", "ty_le_xac_minh", process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api_output, ("ty_le_xac_minh_chi_tiet",)),
    ProcessorTask("phieu_hoan_cong_dich_vu", "dich_vu", process_phieu_hoan_cong_dich_vu_chi_tiet_api_output, ("phieu_hoan_cong_dich_vu_chi_tiet",)),
    ProcessorTask("tam_dung_khoi_phuc_chi_tiet", "dich_vu", process_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_api_output, ("tam_dung_khoi_phuc_dich_vu_chi_tiet", "tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc")),
    ProcessorTask("tam_dung_khoi_phuc_tong_hop", "dich_vu", process_tam_dung_khoi_phuc_dich_vu_tong_hop_api_output, ("tam_dung_khoi_phuc_dich_vu_tong_hop",)),
    ProcessorTask("fiber_thuc_tang", "tam_dung_khoi_phuc_dich_vu", process_fiber_thuc_tang_api_output, ("tam_dung_khoi_phuc_dich_vu_chi_tiet", "phieu_hoan_cong_dich_vu_chi_tiet")),
    ProcessorTask("mytv_ngung_psc", "tam_dung_khoi_phuc_dich_vu", process_mytv_ngung_psc_api_output, ("ngung_psc_mytv_thang_t_1_cap_to", "ngung_psc_mytv_thang_t_1_cap_ttvt")),
    ProcessorTask("mytv_ngung_psc_ttvt", "tam_dung_khoi_phuc_dich_vu", process_mytv_ngung_psc_ttvt_api_output, ("ngung_psc_mytv_thang_t_1_cap_ttvt",)),
    ProcessorTask("mytv_hoan_cong", "phieu_hoan_cong_dich_vu", process_mytv_hoan_cong_api_output, ("phieu_hoan_cong_dich_vu_chi_tiet",)),
    ProcessorTask("mytv_thuc_tang", "tam_dung_khoi_phuc_dich_vu", process_mytv_thuc_tang_api_output, ("ngung_psc_mytv_thang_t_1_cap_to", "phieu_hoan_cong_dich_vu_chi_tiet")),
    ProcessorTask("son_tay_mytv_ngung_psc_t_minus_1", "tam_dung_khoi_phuc_dich_vu", process_son_tay_mytv_ngung_psc_t_minus_1_api_output, ("ngung_psc_mytv_thang_t_1_cap_to",)),
    ProcessorTask("son_tay_fiber_ngung_psc_t_minus_1", "tam_dung_khoi_phuc_dich_vu", process_son_tay_fiber_ngung_psc_t_minus_1_api_output, ("ngung_psc_fiber_thang_t_1_cap_to",)),
    ProcessorTask("son_tay_fiber_ngung_psc_t_minus_1_ttvt", "tam_dung_khoi_phuc_dich_vu", process_son_tay_fiber_ngung_psc_t_minus_1_ttvt_api_output, ("ngung_psc_fiber_thang_t_1_cap_ttvt",)),
    ProcessorTask("vat_tu_thu_hoi", "vat_tu_thu_hoi", process_vat_tu_thu_hoi_api_output, ("vattu_thu_hoi",)),
    ProcessorTask("cau_hinh_tu_dong_ptm", "cau_hinh_tu_dong", process_cau_hinh_tu_dong_ptm_api_output, ("cau_hinh_tu_dong_ptm",)),
    ProcessorTask("cau_hinh_tu_dong_thay_the", "cau_hinh_tu_dong", process_cau_hinh_tu_dong_thay_the_api_output, ("cau_hinh_tu_dong_thay_the",)),
    ProcessorTask("cau_hinh_tu_dong_chi_tiet", "cau_hinh_tu_dong", process_cau_hinh_tu_dong_chi_tiet_api_output, ("cau_hinh_tu_dong_chi_tiet",)),
)

PROCESSOR_TASK_BY_NAME = {task.name: task for task in PROCESSOR_TASKS}
PROCESSOR_GROUPS = tuple(dict.fromkeys(task.group for task in PROCESSOR_TASKS))

DOWNLOADED_REPORT_FILES: Dict[str, Tuple[str, str]] = {
    "c11": ("chi_tieu_c", "c1.1 report.xlsx"),
    "c12": ("chi_tieu_c", "c1.2 report.xlsx"),
    "c13": ("chi_tieu_c", "c1.3 report.xlsx"),
    "c14": ("chi_tieu_c", "c1.4 report.xlsx"),
    "c15": ("chi_tieu_c", "c1.5 report.xlsx"),
    "c15_chitiet": ("chi_tieu_c", "c1.5_chitiet_report.xlsx"),
    "i15": ("chi_tieu_i", "i1.5 report.xlsx"),
    "i15_k2": ("chi_tieu_i", "i1.5_k2 report.xlsx"),
    "c14_chi_tiet": ("chi_tieu_c", "c1.4_chitiet_report.xlsx"),
    "c11_chi_tiet": ("chi_tieu_c", "c1.1_chitiet_report.xlsx"),
    "c12_chi_tiet_sm1": ("chi_tieu_c", "c1.2_chitiet_sm1_report.xlsx"),
    "c12_chi_tiet_sm2": ("chi_tieu_c", "c1.2_chitiet_sm2_report.xlsx"),
    "kpi_nvkt_c11": ("kpi_nvkt", "c11-nvktdb report.xlsx"),
    "kpi_nvkt_c12": ("kpi_nvkt", "c12-nvktdb report.xlsx"),
    "kpi_nvkt_c13": ("kpi_nvkt", "c13-nvktdb report.xlsx"),
    "kq_tiep_thi": ("kq_tiep_thi", "kq_tiep_thi report.xlsx"),
    "ghtt_hni": ("ghtt", "ghtt_hni report.xlsx"),
    "ghtt_sontay": ("ghtt", "ghtt_sontay report.xlsx"),
    "ghtt_nvktdb": ("ghtt", "ghtt_nvktdb report.xlsx"),
    "xac_minh_tam_dung": ("xac_minh_tam_dung", "xac_minh_tam_dung report.xlsx"),
    "ty_le_xac_minh_ttvtkv": ("ty_le_xac_minh", "ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv.xlsx"),
    "ty_le_xac_minh_chi_tiet": ("ty_le_xac_minh", "ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet.xlsx"),
    "phieu_hoan_cong_dich_vu_chi_tiet": ("phieu_hoan_cong_dich_vu", "phieu_hoan_cong_dich_vu_chi_tiet.xlsx"),
    "tam_dung_khoi_phuc_dich_vu_chi_tiet": ("tam_dung_khoi_phuc_dich_vu", "tam_dung_khoi_phuc_dich_vu_chi_tiet.xlsx"),
    "tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc": ("tam_dung_khoi_phuc_dich_vu", "tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc.xlsx"),
    "tam_dung_khoi_phuc_dich_vu_tong_hop": ("tam_dung_khoi_phuc_dich_vu", "tam_dung_khoi_phuc_dich_vu_tong_hop.xlsx"),
    "ngung_psc_mytv_thang_t_1_cap_ttvt": ("tam_dung_khoi_phuc_dich_vu", "ngung_psc_mytv_thang_t-1_cap_ttvt.xlsx"),
    "ngung_psc_fiber_thang_t_1_cap_ttvt": ("tam_dung_khoi_phuc_dich_vu", "ngung_psc_fiber_thang_t-1_cap_ttvt.xlsx"),
    "ngung_psc_fiber_thang_t_1_cap_to": ("tam_dung_khoi_phuc_dich_vu", "ngung_psc_fiber_thang_t-1_cap_to.xlsx"),
    "ngung_psc_mytv_thang_t_1_cap_to": ("tam_dung_khoi_phuc_dich_vu", "ngung_psc_mytv_thang_t-1_cap_to.xlsx"),
    "vattu_thu_hoi": ("vat_tu_thu_hoi", "bc_thu_hoi_vat_tu.xlsx"),
    "cau_hinh_tu_dong_ptm": ("cau_hinh_tu_dong", "cau_hinh_tu_dong_ptm.xlsx"),
    "cau_hinh_tu_dong_thay_the": ("cau_hinh_tu_dong", "cau_hinh_tu_dong_thay_the.xlsx"),
    "cau_hinh_tu_dong_chi_tiet": ("cau_hinh_tu_dong", "cau_hinh_tu_dong_chi_tiet.xlsx"),
}

SINGLE_INPUT_PROCESSORS: Dict[str, str] = {
    "c11": "c11",
    "c12": "c12",
    "c13": "c13",
    "c14": "c14",
    "c15": "c15",
    "c15_chitiet": "c15_chitiet",
    "c14_chi_tiet": "c14_chi_tiet",
    "c11_chi_tiet": "c11_chi_tiet",
    "ghtt_hni": "ghtt_hni",
    "ghtt_son_tay": "ghtt_sontay",
    "ghtt_nvktdb": "ghtt_nvktdb",
    "xac_minh_tam_dung": "xac_minh_tam_dung",
    "xac_minh_ttvtkv": "ty_le_xac_minh_ttvtkv",
    "xac_minh_chi_tiet": "ty_le_xac_minh_chi_tiet",
    "phieu_hoan_cong_dich_vu": "phieu_hoan_cong_dich_vu_chi_tiet",
    "tam_dung_khoi_phuc_tong_hop": "tam_dung_khoi_phuc_dich_vu_tong_hop",
    "son_tay_mytv_ngung_psc_t_minus_1": "ngung_psc_mytv_thang_t_1_cap_to",
    "son_tay_fiber_ngung_psc_t_minus_1": "ngung_psc_fiber_thang_t_1_cap_to",
    "son_tay_fiber_ngung_psc_t_minus_1_ttvt": "ngung_psc_fiber_thang_t_1_cap_ttvt",
    "vat_tu_thu_hoi": "vattu_thu_hoi",
    "cau_hinh_tu_dong_ptm": "cau_hinh_tu_dong_ptm",
    "cau_hinh_tu_dong_thay_the": "cau_hinh_tu_dong_thay_the",
    "cau_hinh_tu_dong_chi_tiet": "cau_hinh_tu_dong_chi_tiet",
}


def list_processors() -> Tuple[ProcessorTask, ...]:
    """Tra ve danh sach processor chinh dang duoc runner quan ly."""

    return PROCESSOR_TASKS


def _normalise_name_set(names: Optional[Sequence[str]]) -> set[str]:
    return {name.strip() for name in (names or []) if name and name.strip()}


def _select_tasks(
    only: Optional[Sequence[str]] = None,
    skip: Optional[Sequence[str]] = None,
    groups: Optional[Sequence[str]] = None,
) -> Tuple[List[ProcessorTask], List[str]]:
    only_set = _normalise_name_set(only)
    skip_set = _normalise_name_set(skip)
    group_set = _normalise_name_set(groups)
    unknown = sorted((only_set | skip_set) - set(PROCESSOR_TASK_BY_NAME))
    unknown_groups = sorted(group_set - set(PROCESSOR_GROUPS))
    unknown.extend(f"group:{group}" for group in unknown_groups)

    selected: List[ProcessorTask] = []
    for task in PROCESSOR_TASKS:
        if only_set and task.name not in only_set:
            continue
        if task.name in skip_set:
            continue
        if group_set and task.group not in group_set:
            continue
        selected.append(task)
    return selected, unknown


def _downloaded_report_path(runtime_context: RuntimeContext, report_key: str) -> Path:
    try:
        group_name, filename = DOWNLOADED_REPORT_FILES[report_key]
    except KeyError as exc:
        raise KeyError(f"Chua khai bao mapping file raw cho report_key='{report_key}'") from exc
    return runtime_context.download_group_dir(group_name) / filename


def _runtime_task_kwargs(task: ProcessorTask, runtime_context: RuntimeContext) -> Dict[str, Any]:
    if task.name in {"i15", "i15_k2"}:
        return {
            "input_path": _downloaded_report_path(runtime_context, task.name),
            "history_db_path": runtime_context.paths.sqlite_db_path,
            "dsnv_db_path": DEFAULT_DANHBA_DB_FILE,
        }

    if task.name in SINGLE_INPUT_PROCESSORS:
        return {"input_path": _downloaded_report_path(runtime_context, SINGLE_INPUT_PROCESSORS[task.name])}

    if task.name == "c12_chi_tiet":
        return {
            "sm1_input_path": _downloaded_report_path(runtime_context, "c12_chi_tiet_sm1"),
            "sm2_input_path": _downloaded_report_path(runtime_context, "c12_chi_tiet_sm2"),
        }

    if task.name in {"kpi_nvkt_c11", "kpi_nvkt_c12", "kpi_nvkt_c13"}:
        return {
            "input_path": _downloaded_report_path(runtime_context, task.name),
            "dsnv_file": DEFAULT_DSNV_FILE,
        }

    if task.name == "kq_tiep_thi":
        return {
            "input_path": _downloaded_report_path(runtime_context, "kq_tiep_thi"),
            "dsnv_file": DEFAULT_DSNV_FILE,
        }

    if task.name == "ghtt_nvktdb":
        return {
            "input_path": _downloaded_report_path(runtime_context, "ghtt_nvktdb"),
            "dsnv_file": DEFAULT_DSNV_FILE,
        }

    if task.name == "tam_dung_khoi_phuc_chi_tiet":
        return {
            "tam_dung_input_path": _downloaded_report_path(runtime_context, "tam_dung_khoi_phuc_dich_vu_chi_tiet"),
            "khoi_phuc_input_path": _downloaded_report_path(runtime_context, "tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc"),
            "combined_output_path": runtime_context.processed_group_dir("tam_dung_khoi_phuc_dich_vu")
            / "tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_processed.xlsx",
        }

    if task.name == "fiber_thuc_tang":
        return {
            "ngung_psc_input": _downloaded_report_path(runtime_context, "tam_dung_khoi_phuc_dich_vu_chi_tiet"),
            "hoan_cong_input": _downloaded_report_path(runtime_context, "phieu_hoan_cong_dich_vu_chi_tiet"),
            "output_path": runtime_context.processed_group_dir("tam_dung_khoi_phuc_dich_vu") / "fiber_thuc_tang_processed.xlsx",
        }

    if task.name == "mytv_ngung_psc":
        return {
            "input_path": _downloaded_report_path(runtime_context, "ngung_psc_mytv_thang_t_1_cap_to"),
            "ttvt_input_path": _downloaded_report_path(runtime_context, "ngung_psc_mytv_thang_t_1_cap_ttvt"),
            "output_path": runtime_context.processed_group_dir("tam_dung_khoi_phuc_dich_vu")
            / "ngung_psc_mytv_thang_t-1_cap_to_processed.xlsx",
        }

    if task.name == "mytv_ngung_psc_ttvt":
        return {
            "input_path": _downloaded_report_path(runtime_context, "ngung_psc_mytv_thang_t_1_cap_ttvt"),
            "output_path": runtime_context.processed_group_dir("tam_dung_khoi_phuc_dich_vu")
            / "ngung_psc_mytv_thang_t-1_cap_ttvt_processed.xlsx",
        }

    if task.name == "mytv_hoan_cong":
        return {
            "input_path": _downloaded_report_path(runtime_context, "phieu_hoan_cong_dich_vu_chi_tiet"),
            "output_path": runtime_context.processed_group_dir("phieu_hoan_cong_dich_vu")
            / "phieu_hoan_cong_dich_vu_chi_tiet_processed.xlsx",
        }

    if task.name == "mytv_thuc_tang":
        return {
            "ngung_psc_input": _downloaded_report_path(runtime_context, "ngung_psc_mytv_thang_t_1_cap_to"),
            "hoan_cong_input": _downloaded_report_path(runtime_context, "phieu_hoan_cong_dich_vu_chi_tiet"),
            "output_path": runtime_context.processed_group_dir("tam_dung_khoi_phuc_dich_vu")
            / "mytv_thuc_tang_processed.xlsx",
        }

    return {}


def _is_task_enabled(task: ProcessorTask, runtime_context: Optional[RuntimeContext]) -> bool:
    if runtime_context is None or not task.source_report_keys:
        return True
    return all(runtime_context.is_report_enabled(report_key) for report_key in task.source_report_keys)


def _build_call_kwargs(
    task: ProcessorTask,
    overwrite_processed: bool,
    runtime_context: Optional[RuntimeContext] = None,
) -> Dict[str, Any]:
    signature = inspect.signature(task.func)
    kwargs: Dict[str, Any] = {}
    if "overwrite_processed" in signature.parameters:
        kwargs["overwrite_processed"] = overwrite_processed
    if runtime_context is None:
        return kwargs

    for key, value in _runtime_task_kwargs(task, runtime_context).items():
        if key in signature.parameters:
            kwargs[key] = value
    return kwargs


def run_all_processors(
    *,
    overwrite_processed: bool = False,
    only: Optional[Sequence[str]] = None,
    skip: Optional[Sequence[str]] = None,
    groups: Optional[Sequence[str]] = None,
    config_path: Optional[str] = None,
    runtime_context: Optional[RuntimeContext] = None,
    stop_on_error: bool = False,
    verbose: bool = True,
) -> Dict[str, List[ProcessorRunResult]]:
    """Chay tat ca processor dang co trong `api_transition/processors`.

    Runner nay chi gom cac processor da co implementation trong package
    `api_transition.processors`. Cac luong chua duoc port vao package nay
    se khong nam trong danh sach.
    """

    active_runtime_context = runtime_context
    if active_runtime_context is None and config_path:
        active_runtime_context = load_runtime_context(config_path)

    tasks, unknown = _select_tasks(only=only, skip=skip, groups=groups)
    results: Dict[str, List[ProcessorRunResult]] = {
        "success": [],
        "failed": [],
        "skipped": [],
    }

    for item in unknown:
        results["skipped"].append(
            ProcessorRunResult(
                name=item,
                group="unknown",
                status="skipped",
                error="Unknown processor or group.",
            )
        )

    if active_runtime_context is not None:
        configure_runtime_roots(
            downloads_root=active_runtime_context.paths.downloads_root,
            processed_root=active_runtime_context.paths.processed_root,
        )

    try:
        if verbose:
            print(f"[processors] Selected {len(tasks)} task(s). overwrite_processed={overwrite_processed}")
            if active_runtime_context is not None:
                print(
                    f"[processors] Unit={active_runtime_context.unit.code}"
                    f" downloads_root={active_runtime_context.paths.downloads_root}"
                    f" processed_root={active_runtime_context.paths.processed_root}"
                )

        for index, task in enumerate(tasks, start=1):
            if not _is_task_enabled(task, active_runtime_context):
                run_result = ProcessorRunResult(
                    name=task.name,
                    group=task.group,
                    status="skipped",
                    error="Disabled by runtime config.",
                )
                results["skipped"].append(run_result)
                if verbose:
                    print(f"[{index}/{len(tasks)}] Skipping {task.name} ({task.group}) - disabled by config")
                continue

            if verbose:
                print(f"[{index}/{len(tasks)}] Running {task.name} ({task.group})")

            started = time.perf_counter()
            try:
                result = task.func(
                    **_build_call_kwargs(task, overwrite_processed, active_runtime_context)
                )
                duration = time.perf_counter() - started
                run_result = ProcessorRunResult(
                    name=task.name,
                    group=task.group,
                    status="success",
                    duration_seconds=duration,
                    result=result,
                )
                results["success"].append(run_result)
                if verbose:
                    print(f"    OK in {duration:.2f}s")
            except Exception as exc:
                duration = time.perf_counter() - started
                run_result = ProcessorRunResult(
                    name=task.name,
                    group=task.group,
                    status="failed",
                    duration_seconds=duration,
                    error="".join(traceback.format_exception_only(type(exc), exc)).strip(),
                )
                results["failed"].append(run_result)
                if verbose:
                    print(f"    FAILED in {duration:.2f}s: {run_result.error}")
                if stop_on_error:
                    break
    finally:
        if active_runtime_context is not None:
            reset_runtime_roots()

    if verbose:
        print(
            "[processors] Completed:"
            f" success={len(results['success'])}"
            f" failed={len(results['failed'])}"
            f" skipped={len(results['skipped'])}"
        )

    return results


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run all processors in api_transition.processors.")
    parser.add_argument(
        "--only",
        action="append",
        default=[],
        help="Chi chay processor co ten nay. Co the lap lai nhieu lan.",
    )
    parser.add_argument(
        "--skip",
        action="append",
        default=[],
        help="Bo qua processor co ten nay. Co the lap lai nhieu lan.",
    )
    parser.add_argument(
        "--group",
        action="append",
        default=[],
        help="Chi chay nhom processor nay. Co the lap lai nhieu lan.",
    )
    parser.add_argument(
        "--overwrite-processed",
        action="store_true",
        help="Cho phep ghi de workbook processed neu processor ho tro.",
    )
    parser.add_argument(
        "--stop-on-error",
        action="store_true",
        help="Dung ngay khi gap processor loi.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Tat log tien trinh runner.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Chi in danh sach processor va thoat.",
    )
    parser.add_argument(
        "--config",
        help="Duong dan toi file config runtime theo don vi.",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    runtime_context = load_runtime_context(args.config) if args.config else None

    if args.list:
        for task in PROCESSOR_TASKS:
            status = "enabled"
            if runtime_context is not None and not _is_task_enabled(task, runtime_context):
                status = "disabled"
            print(f"{task.name}\t{task.group}\t{status}")
        return 0

    results = run_all_processors(
        overwrite_processed=args.overwrite_processed,
        only=args.only,
        skip=args.skip,
        groups=args.group,
        runtime_context=runtime_context,
        stop_on_error=args.stop_on_error,
        verbose=not args.quiet,
    )
    return 1 if results["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
