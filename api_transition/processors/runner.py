"""Runner gom va dieu phoi cac processor trong api_transition.processors."""

from __future__ import annotations

import argparse
import inspect
import time
import traceback
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from api_transition.processors.c_processors import (
    process_c11_chitiet_report_api_output,
    process_c11_report_api_output,
    process_c12_chitiet_reports_api_output,
    process_c12_report_api_output,
    process_c13_report_api_output,
    process_c14_chitiet_report_api_output,
    process_c14_report_api_output,
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
    process_mytv_thuc_tang_api_output,
    process_phieu_hoan_cong_dich_vu_chi_tiet_api_output,
    process_son_tay_fiber_ngung_psc_t_minus_1_api_output,
    process_son_tay_mytv_ngung_psc_t_minus_1_api_output,
    process_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_api_output,
    process_tam_dung_khoi_phuc_dich_vu_tong_hop_api_output,
)
from api_transition.processors.vattu_processors import (
    process_quyet_toan_vat_tu_api_output,
    process_vat_tu_thu_hoi_api_output,
)
from api_transition.processors.verification_processors import (
    process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api_output,
    process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api_output,
)


@dataclass(frozen=True)
class ProcessorTask:
    """Mo ta 1 processor co the chay doc lap trong batch runner."""

    name: str
    group: str
    func: Callable[..., Any]


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
    ProcessorTask("c11", "chi_tieu_c", process_c11_report_api_output),
    ProcessorTask("c12", "chi_tieu_c", process_c12_report_api_output),
    ProcessorTask("c13", "chi_tieu_c", process_c13_report_api_output),
    ProcessorTask("c14", "chi_tieu_c", process_c14_report_api_output),
    ProcessorTask("c14_chi_tiet", "chi_tieu_c", process_c14_chitiet_report_api_output),
    ProcessorTask("c11_chi_tiet", "chi_tieu_c", process_c11_chitiet_report_api_output),
    ProcessorTask("c12_chi_tiet", "chi_tieu_c", process_c12_chitiet_reports_api_output),
    ProcessorTask("kpi_nvkt_c11", "kpi_nvkt", process_kpi_nvkt_c11_api_output),
    ProcessorTask("kpi_nvkt_c12", "kpi_nvkt", process_kpi_nvkt_c12_api_output),
    ProcessorTask("kpi_nvkt_c13", "kpi_nvkt", process_kpi_nvkt_c13_api_output),
    ProcessorTask("kq_tiep_thi", "kq_tiep_thi", process_kq_tiep_thi_api_output),
    ProcessorTask("ghtt_hni", "ghtt", process_ghtt_hni_api_output),
    ProcessorTask("ghtt_son_tay", "ghtt", process_ghtt_sontay_api_output),
    ProcessorTask("ghtt_nvktdb", "ghtt", process_ghtt_nvktdb_api_output),
    ProcessorTask("xac_minh_ttvtkv", "ty_le_xac_minh", process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api_output),
    ProcessorTask("xac_minh_chi_tiet", "ty_le_xac_minh", process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api_output),
    ProcessorTask("phieu_hoan_cong_dich_vu", "dich_vu", process_phieu_hoan_cong_dich_vu_chi_tiet_api_output),
    ProcessorTask("tam_dung_khoi_phuc_chi_tiet", "dich_vu", process_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_api_output),
    ProcessorTask("tam_dung_khoi_phuc_tong_hop", "dich_vu", process_tam_dung_khoi_phuc_dich_vu_tong_hop_api_output),
    ProcessorTask("fiber_thuc_tang", "dich_vu", process_fiber_thuc_tang_api_output),
    ProcessorTask("mytv_ngung_psc", "mytv_dich_vu", process_mytv_ngung_psc_api_output),
    ProcessorTask("mytv_hoan_cong", "mytv_dich_vu", process_mytv_hoan_cong_api_output),
    ProcessorTask("mytv_thuc_tang", "mytv_dich_vu", process_mytv_thuc_tang_api_output),
    ProcessorTask("son_tay_mytv_ngung_psc_t_minus_1", "mytv_dich_vu", process_son_tay_mytv_ngung_psc_t_minus_1_api_output),
    ProcessorTask("son_tay_fiber_ngung_psc_t_minus_1", "thuc_tang_ngung_psc", process_son_tay_fiber_ngung_psc_t_minus_1_api_output),
    ProcessorTask("vat_tu_thu_hoi", "vat_tu_thu_hoi", process_vat_tu_thu_hoi_api_output),
    ProcessorTask("quyet_toan_vat_tu", "vat_tu_thu_hoi", process_quyet_toan_vat_tu_api_output),
    ProcessorTask("cau_hinh_tu_dong_ptm", "cau_hinh_tu_dong", process_cau_hinh_tu_dong_ptm_api_output),
    ProcessorTask("cau_hinh_tu_dong_thay_the", "cau_hinh_tu_dong", process_cau_hinh_tu_dong_thay_the_api_output),
    ProcessorTask("cau_hinh_tu_dong_chi_tiet", "cau_hinh_tu_dong", process_cau_hinh_tu_dong_chi_tiet_api_output),
)

PROCESSOR_TASK_BY_NAME = {task.name: task for task in PROCESSOR_TASKS}
PROCESSOR_GROUPS = tuple(dict.fromkeys(task.group for task in PROCESSOR_TASKS))


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


def _build_call_kwargs(task: ProcessorTask, overwrite_processed: bool) -> Dict[str, Any]:
    signature = inspect.signature(task.func)
    kwargs: Dict[str, Any] = {}
    if "overwrite_processed" in signature.parameters:
        kwargs["overwrite_processed"] = overwrite_processed
    return kwargs


def run_all_processors(
    *,
    overwrite_processed: bool = False,
    only: Optional[Sequence[str]] = None,
    skip: Optional[Sequence[str]] = None,
    groups: Optional[Sequence[str]] = None,
    stop_on_error: bool = False,
    verbose: bool = True,
) -> Dict[str, List[ProcessorRunResult]]:
    """Chay tat ca processor dang co trong `api_transition/processors`.

    Runner nay chi gom cac processor da co implementation trong package
    `api_transition.processors`. Cac luong chua duoc port vao package nay
    se khong nam trong danh sach.
    """

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

    if verbose:
        print(f"[processors] Selected {len(tasks)} task(s). overwrite_processed={overwrite_processed}")

    for index, task in enumerate(tasks, start=1):
        if verbose:
            print(f"[{index}/{len(tasks)}] Running {task.name} ({task.group})")

        started = time.perf_counter()
        try:
            result = task.func(**_build_call_kwargs(task, overwrite_processed))
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
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.list:
        for task in PROCESSOR_TASKS:
            print(f"{task.name}\t{task.group}")
        return 0

    results = run_all_processors(
        overwrite_processed=args.overwrite_processed,
        only=args.only,
        skip=args.skip,
        groups=args.group,
        stop_on_error=args.stop_on_error,
        verbose=not args.quiet,
    )
    return 1 if results["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
