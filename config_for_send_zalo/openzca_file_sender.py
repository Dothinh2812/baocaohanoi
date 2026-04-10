#!/usr/bin/env python3
"""
Helpers for sending local files to Zalo groups through the openzca CLI.

The group IDs below were looked up from the logged-in "Xuân Thịnh" account
via `openzca group list --json` on 2026-03-31.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


DEFAULT_OPENZCA_BIN = Path(
    os.getenv(
        "OPENZCA_BIN",
        "/home/vtst/.nvm/versions/node/v22.22.2/bin/openzca",
    )
)

DEFAULT_NVKT_BASE_DIR = Path("downloads/baocao_hanoi")
DEFAULT_NVKT_FOLDERS = (
    "shc_NVKT_danh_sach_chi_tiet_K1",
    "shc_NVKT_danh_sach_chi_tiet_K2",
)
DEFAULT_DELAY_SECONDS = float(os.getenv("OPENZCA_FILE_SEND_DELAY_SECONDS", "2"))
DEFAULT_UPLOAD_TIMEOUT_SECONDS = int(os.getenv("OPENZCA_FILE_SEND_TIMEOUT_SECONDS", "45"))
DEFAULT_CONNECT_TIMEOUT_MS = int(os.getenv("OPENZCA_UPLOAD_LISTENER_CONNECT_TIMEOUT_MS", "8000"))

NVKT_FOLDER_TO_GROUP = {
    "Tổ Kỹ thuật Địa bàn Phúc Thọ": {
        "group_name": "Tổ Phúc Thọ",
        "group_id": "8368428594496880128",
    },
    "Tổ Kỹ thuật Địa bàn Quảng Oai": {
        "group_name": "Tổ Quảng Oai",
        "group_id": "6938134403746251115",
    },
    "Tổ Kỹ thuật Địa bàn Sơn Tây": {
        "group_name": "Tổ Sơn Tây",
        "group_id": "9064103357609041896",
    },
    "Tổ Kỹ thuật Địa bàn Suối hai": {
        "group_name": "Tổ Suối Hai",
        "group_id": "9100414179385695593",
    },
}


@dataclass
class UploadResult:
    batch_folder: str
    team_folder: str
    group_name: str
    group_id: str
    file_path: str
    success: bool
    return_code: int
    stdout: str
    stderr: str


def _run_openzca(args: list[str], openzca_bin: Path = DEFAULT_OPENZCA_BIN) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [str(openzca_bin), *args],
        capture_output=True,
        text=True,
        check=False,
        stdin=subprocess.DEVNULL,
    )


def check_openzca_session(openzca_bin: Path = DEFAULT_OPENZCA_BIN) -> bool:
    process = _run_openzca(["auth", "status"], openzca_bin=openzca_bin)
    return process.returncode == 0


def send_file_to_group(
    group_id: str,
    file_path: Path | str,
    *,
    openzca_bin: Path = DEFAULT_OPENZCA_BIN,
    dry_run: bool = False,
    timeout_seconds: int = DEFAULT_UPLOAD_TIMEOUT_SECONDS,
    debug: bool = False,
) -> subprocess.CompletedProcess[str]:
    file_path = Path(file_path)
    if not file_path.is_file():
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    if dry_run:
        return subprocess.CompletedProcess(
            args=[str(openzca_bin), "msg", "upload", str(file_path), group_id, "--group"],
            returncode=0,
            stdout=f"DRY RUN: {file_path} -> {group_id}",
            stderr="",
        )

    env = os.environ.copy()
    env["OPENZCA_UPLOAD_TIMEOUT_MS"] = str(max(timeout_seconds, 1) * 1000)
    env["OPENZCA_UPLOAD_LISTENER_CONNECT_TIMEOUT_MS"] = str(DEFAULT_CONNECT_TIMEOUT_MS)
    env["OPENZCA_UPLOAD_IPC_TIMEOUT_MS"] = str(max(timeout_seconds, 1) * 1000 + 5000)
    env["OPENZCA_UPLOAD_IPC_HANDLER_TIMEOUT_MS"] = str(max(timeout_seconds, 1) * 1000)
    if debug:
        env["OPENZCA_DEBUG"] = "1"

    command = [str(openzca_bin), "msg", "upload", str(file_path), group_id, "--group"]
    try:
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            stdin=subprocess.DEVNULL,
            timeout=max(timeout_seconds, 1) + 10,
            env=env,
        )
    except subprocess.TimeoutExpired as exc:
        return subprocess.CompletedProcess(
            args=command,
            returncode=124,
            stdout=(exc.stdout or "").strip() if isinstance(exc.stdout, str) else "",
            stderr=f"Hết thời gian chờ sau {timeout_seconds}s khi gửi file: {file_path}",
        )


def send_text_to_group(
    group_id: str,
    message: str,
    *,
    openzca_bin: Path = DEFAULT_OPENZCA_BIN,
    dry_run: bool = False,
    timeout_seconds: int = DEFAULT_UPLOAD_TIMEOUT_SECONDS,
    debug: bool = False,
) -> subprocess.CompletedProcess[str]:
    command = [str(openzca_bin), "msg", "send", group_id, message, "--group"]

    if dry_run:
        return subprocess.CompletedProcess(
            args=command,
            returncode=0,
            stdout=f"DRY RUN TEXT: {group_id} <- {message}",
            stderr="",
        )

    env = os.environ.copy()
    if debug:
        env["OPENZCA_DEBUG"] = "1"

    try:
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            stdin=subprocess.DEVNULL,
            timeout=max(timeout_seconds, 1),
            env=env,
        )
    except subprocess.TimeoutExpired:
        return subprocess.CompletedProcess(
            args=command,
            returncode=124,
            stdout="",
            stderr=f"Hết thời gian chờ sau {timeout_seconds}s khi gửi tiêu đề tới nhóm {group_id}",
        )


def iter_team_files(team_dir: Path) -> Iterable[Path]:
    return sorted(
        path for path in team_dir.iterdir()
        if path.is_file() and path.suffix.lower() in {".xlsx", ".xls", ".csv"}
    )


def send_nvkt_files_in_folder(
    folder_path: Path | str,
    *,
    openzca_bin: Path = DEFAULT_OPENZCA_BIN,
    dry_run: bool = False,
    delay_seconds: float = DEFAULT_DELAY_SECONDS,
    verify_session: bool = True,
    timeout_seconds: int = DEFAULT_UPLOAD_TIMEOUT_SECONDS,
    debug: bool = False,
    message_text: str | None = None,
) -> list[UploadResult]:
    folder_path = Path(folder_path)
    if not folder_path.is_dir():
        raise FileNotFoundError(f"Không tìm thấy thư mục: {folder_path}")

    if verify_session and not dry_run and not check_openzca_session(openzca_bin=openzca_bin):
        raise RuntimeError("openzca chưa đăng nhập hoặc phiên làm việc không hợp lệ.")

    results: list[UploadResult] = []

    team_jobs: list[tuple[str, dict[str, str], Path]] = []
    for team_folder, group_info in NVKT_FOLDER_TO_GROUP.items():
        team_dir = folder_path / team_folder
        if not team_dir.is_dir():
            continue
        for file_path in iter_team_files(team_dir):
            team_jobs.append((team_folder, group_info, file_path))

    total_jobs = len(team_jobs)
    header_sent_group_ids: set[str] = set()

    for index, (team_folder, group_info, file_path) in enumerate(team_jobs, start=1):
        if message_text and group_info["group_id"] not in header_sent_group_ids:
            print(
                f"Gửi tiêu đề tới {group_info['group_name']}: {message_text}",
                flush=True,
            )
            header_process = send_text_to_group(
                group_info["group_id"],
                message_text,
                openzca_bin=openzca_bin,
                dry_run=dry_run,
                timeout_seconds=timeout_seconds,
                debug=debug,
            )
            if header_process.returncode == 0:
                print(f"  OK  tiêu đề {group_info['group_name']}", flush=True)
            else:
                print(f"  FAIL tiêu đề {group_info['group_name']}", flush=True)
                if header_process.stderr:
                    print(f"    stderr: {header_process.stderr}", flush=True)
            header_sent_group_ids.add(group_info["group_id"])

        print(
            f"[{index}/{total_jobs}] {folder_path.name} -> {group_info['group_name']} | {file_path.name}",
            flush=True,
        )
        process = send_file_to_group(
            group_info["group_id"],
            file_path,
            openzca_bin=openzca_bin,
            dry_run=dry_run,
            timeout_seconds=timeout_seconds,
            debug=debug,
        )
        result = UploadResult(
            batch_folder=folder_path.name,
            team_folder=team_folder,
            group_name=group_info["group_name"],
            group_id=group_info["group_id"],
            file_path=str(file_path),
            success=process.returncode == 0,
            return_code=process.returncode,
            stdout=(process.stdout or "").strip(),
            stderr=(process.stderr or "").strip(),
        )
        results.append(result)
        if result.success:
            print(f"  OK  {file_path.name}", flush=True)
        else:
            print(f"  FAIL {file_path.name}", flush=True)
            if result.stderr:
                print(f"    stderr: {result.stderr}", flush=True)
        if delay_seconds > 0 and not dry_run:
            time.sleep(delay_seconds)

    return results


def send_nvkt_files_k1_k2(
    *,
    base_dir: Path | str = DEFAULT_NVKT_BASE_DIR,
    folder_names: Iterable[str] = DEFAULT_NVKT_FOLDERS,
    openzca_bin: Path = DEFAULT_OPENZCA_BIN,
    dry_run: bool = False,
    delay_seconds: float = DEFAULT_DELAY_SECONDS,
    verify_session: bool = True,
    timeout_seconds: int = DEFAULT_UPLOAD_TIMEOUT_SECONDS,
    debug: bool = False,
    message_text: str | None = None,
) -> list[UploadResult]:
    base_dir = Path(base_dir)
    all_results: list[UploadResult] = []

    for folder_name in folder_names:
        folder_path = base_dir / folder_name
        all_results.extend(
            send_nvkt_files_in_folder(
                folder_path,
                openzca_bin=openzca_bin,
                dry_run=dry_run,
                delay_seconds=delay_seconds,
                verify_session=verify_session,
                timeout_seconds=timeout_seconds,
                debug=debug,
                message_text=message_text,
            )
        )

    return all_results


def build_summary(results: Iterable[UploadResult]) -> dict[str, int]:
    results = list(results)
    return {
        "total": len(results),
        "success": sum(1 for item in results if item.success),
        "failed": sum(1 for item in results if not item.success),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Send NVKT detail files to Zalo groups via openzca.")
    parser.add_argument(
        "--base-dir",
        default=str(DEFAULT_NVKT_BASE_DIR),
        help="Thư mục gốc chứa các thư mục shc_NVKT_danh_sach_chi_tiet_K1/K2.",
    )
    parser.add_argument(
        "--folder",
        action="append",
        dest="folders",
        help="Chỉ định thư mục con cần gửi. Có thể truyền nhiều lần.",
    )
    parser.add_argument(
        "--delay-seconds",
        type=float,
        default=DEFAULT_DELAY_SECONDS,
        help="Số giây nghỉ giữa các lần upload file.",
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Thực hiện gửi thật. Mặc định chỉ dry run.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=DEFAULT_UPLOAD_TIMEOUT_SECONDS,
        help="Timeout cho mỗi file khi gọi openzca.",
    )
    parser.add_argument(
        "--debug-openzca",
        action="store_true",
        help="Bật OPENZCA_DEBUG=1 cho từng lệnh upload.",
    )
    args = parser.parse_args()

    folder_names = tuple(args.folders) if args.folders else DEFAULT_NVKT_FOLDERS
    results = send_nvkt_files_k1_k2(
        base_dir=args.base_dir,
        folder_names=folder_names,
        dry_run=not args.send,
        delay_seconds=args.delay_seconds,
        timeout_seconds=args.timeout_seconds,
        debug=args.debug_openzca,
    )

    summary = build_summary(results)
    print(f"Tổng file: {summary['total']}")
    print(f"Thành công: {summary['success']}")
    print(f"Thất bại: {summary['failed']}")

    for item in results:
        status = "OK" if item.success else "FAIL"
        print(f"[{status}] {item.batch_folder} | {item.group_name} | {item.file_path}")
        if item.stderr:
            print(f"  stderr: {item.stderr}")

    return 0 if summary["failed"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
