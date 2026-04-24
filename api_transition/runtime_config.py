# -*- coding: utf-8 -*-
"""Load va validate runtime config cho mo hinh multi-instance theo don vi."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Mapping

import yaml


API_TRANSITION_DIR = Path(__file__).resolve().parent


@dataclass(frozen=True)
class RuntimePaths:
    instance_root: Path
    downloads_root: Path
    processed_root: Path
    archive_root: Path
    sqlite_root: Path
    sqlite_db_path: Path
    sqlite_import_logs_root: Path
    sqlite_exports_root: Path
    lock_file_path: Path


@dataclass(frozen=True)
class UnitProfile:
    code: str
    name: str
    ids: Mapping[str, str]
    team_ids: Mapping[str, Mapping[str, str]] = field(default_factory=dict)


@dataclass(frozen=True)
class PeriodConfig:
    report_month: int
    report_year: int
    month_id: str = ""
    month_label: str = ""
    vattu_start_date: str = ""


@dataclass(frozen=True)
class DownloadConfig:
    headed: bool = False
    max_retries: int = 3
    retry_timeouts: tuple[int, ...] = (180, 300, 500)
    retry_delay_seconds: int = 3


def _deep_merge_dict(base: Mapping[str, Any], override: Mapping[str, Any]) -> Dict[str, Any]:
    merged: Dict[str, Any] = dict(base)
    for key, value in override.items():
        if isinstance(value, Mapping) and isinstance(base.get(key), Mapping):
            merged[key] = _deep_merge_dict(base[key], value)
        else:
            merged[key] = value
    return merged


@dataclass(frozen=True)
class RuntimeContext:
    config_path: Path
    paths: RuntimePaths
    unit: UnitProfile
    period: PeriodConfig
    download: DownloadConfig
    report_configs: Mapping[str, Mapping[str, Any]] = field(default_factory=dict)
    raw_config: Mapping[str, Any] = field(default_factory=dict)

    def get_report_config(self, report_key: str) -> Mapping[str, Any]:
        defaults = self.report_configs.get("defaults", {})
        report_config = self.report_configs.get(report_key, {})
        return _deep_merge_dict(defaults, report_config)

    def is_report_enabled(self, report_key: str) -> bool:
        report_config = self.get_report_config(report_key)
        enabled = report_config.get("enabled", True)
        return bool(enabled)

    def download_group_dir(self, group_name: str) -> Path:
        return self.paths.downloads_root / group_name

    def processed_group_dir(self, group_name: str) -> Path:
        return self.paths.processed_root / group_name

    def archive_group_dir(self, group_name: str) -> Path:
        return self.paths.archive_root / group_name

    def sqlite_log_path(self, filename: str) -> Path:
        return self.paths.sqlite_import_logs_root / filename

    def lock_file_path(self) -> Path:
        return self.paths.lock_file_path


def _read_yaml(config_path: Path) -> Mapping[str, Any]:
    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except yaml.YAMLError as exc:
        raise ValueError(f"Config YAML khong hop le: {config_path}: {exc}") from exc

    if raw is None:
        raise ValueError(f"Config rong: {config_path}")
    if not isinstance(raw, Mapping):
        raise ValueError(f"Config phai la YAML object o cap goc: {config_path}")
    return raw


def _require_mapping(parent: Mapping[str, Any], key: str, *, context: str) -> Mapping[str, Any]:
    value = parent.get(key)
    if value is None:
        raise ValueError(f"Thieu block bat buoc '{key}' trong {context}")
    if not isinstance(value, Mapping):
        raise ValueError(f"Block '{key}' trong {context} phai la object")
    return value


def _require_non_empty_str(parent: Mapping[str, Any], key: str, *, context: str) -> str:
    value = parent.get(key)
    if value is None:
        raise ValueError(f"Thieu truong bat buoc '{key}' trong {context}")
    text = str(value).strip()
    if not text:
        raise ValueError(f"Truong '{key}' trong {context} khong duoc rong")
    return text


def _optional_str(parent: Mapping[str, Any], key: str, default: str = "") -> str:
    value = parent.get(key, default)
    if value is None:
        return default
    return str(value).strip()


def _require_int(parent: Mapping[str, Any], key: str, *, context: str) -> int:
    value = parent.get(key)
    if value is None:
        raise ValueError(f"Thieu truong bat buoc '{key}' trong {context}")
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Truong '{key}' trong {context} phai la so nguyen") from exc


def _coerce_bool(value: Any, *, field_name: str) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"true", "1", "yes", "y", "on"}:
            return True
        if lowered in {"false", "0", "no", "n", "off"}:
            return False
    if isinstance(value, int):
        return bool(value)
    raise ValueError(f"Truong '{field_name}' phai la bool")


def _coerce_retry_timeouts(value: Any) -> tuple[int, ...]:
    if value is None:
        return (180, 300, 500)
    if not isinstance(value, (list, tuple)):
        raise ValueError("Truong 'download.retry_timeouts' phai la list/tuple")
    result = []
    for item in value:
        try:
            result.append(int(item))
        except (TypeError, ValueError) as exc:
            raise ValueError("Moi gia tri trong 'download.retry_timeouts' phai la so nguyen") from exc
    if not result:
        raise ValueError("Truong 'download.retry_timeouts' khong duoc rong")
    return tuple(result)


def _normalize_string_mapping(
    value: Any,
    *,
    field_name: str,
    allow_empty: bool = True,
) -> Dict[str, str]:
    if value is None:
        return {}
    if not isinstance(value, Mapping):
        raise ValueError(f"Truong '{field_name}' phai la object")

    normalized: Dict[str, str] = {}
    for raw_key, raw_value in value.items():
        key = str(raw_key).strip()
        if not key:
            raise ValueError(f"Truong '{field_name}' khong duoc co key rong")
        text = "" if raw_value is None else str(raw_value).strip()
        if not allow_empty and not text:
            raise ValueError(f"Truong '{field_name}.{key}' khong duoc rong")
        normalized[key] = text
    return normalized


def _resolve_instance_root(instance_root: str) -> Path:
    raw_path = Path(instance_root).expanduser()
    if raw_path.is_absolute():
        return raw_path.resolve()
    return (API_TRANSITION_DIR / raw_path).resolve()


def build_runtime_paths(instance_root: str | Path) -> RuntimePaths:
    if isinstance(instance_root, Path):
        root = instance_root.expanduser()
        instance_root_path = root.resolve() if root.is_absolute() else (API_TRANSITION_DIR / root).resolve()
    else:
        instance_root_path = _resolve_instance_root(str(instance_root))

    downloads_root = instance_root_path / "downloads"
    processed_root = instance_root_path / "Processed"
    archive_root = instance_root_path / "ProcessedDaily"
    sqlite_root = instance_root_path / "sqlite_history"
    sqlite_db_path = sqlite_root / "report_history.db"
    sqlite_import_logs_root = sqlite_root / "import_logs"
    sqlite_exports_root = sqlite_root / "exports"
    lock_file_path = instance_root_path / ".pipeline.lock"

    return RuntimePaths(
        instance_root=instance_root_path,
        downloads_root=downloads_root,
        processed_root=processed_root,
        archive_root=archive_root,
        sqlite_root=sqlite_root,
        sqlite_db_path=sqlite_db_path,
        sqlite_import_logs_root=sqlite_import_logs_root,
        sqlite_exports_root=sqlite_exports_root,
        lock_file_path=lock_file_path,
    )


def ensure_runtime_dirs(paths: RuntimePaths) -> None:
    dirs = (
        paths.instance_root,
        paths.downloads_root,
        paths.processed_root,
        paths.archive_root,
        paths.sqlite_root,
        paths.sqlite_import_logs_root,
        paths.sqlite_exports_root,
    )
    for path in dirs:
        path.mkdir(parents=True, exist_ok=True)


def validate_runtime_config(raw: Mapping[str, Any]) -> None:
    version = raw.get("version")
    if version is None:
        raise ValueError("Thieu truong bat buoc 'version'")
    try:
        version_int = int(version)
    except (TypeError, ValueError) as exc:
        raise ValueError("Truong 'version' phai la so nguyen") from exc
    if version_int != 1:
        raise ValueError(f"Schema config version khong duoc ho tro: {version_int}")

    unit = _require_mapping(raw, "unit", context="config")
    _require_non_empty_str(unit, "code", context="unit")
    _require_non_empty_str(unit, "name", context="unit")

    runtime = _require_mapping(raw, "runtime", context="config")
    _require_non_empty_str(runtime, "instance_root", context="runtime")

    period = _require_mapping(raw, "period", context="config")
    report_month = _require_int(period, "report_month", context="period")
    report_year = _require_int(period, "report_year", context="period")
    if not 1 <= report_month <= 12:
        raise ValueError("Truong 'period.report_month' phai nam trong khoang 1..12")
    if report_year < 2000:
        raise ValueError("Truong 'period.report_year' co ve khong hop le")

    ids = _require_mapping(raw, "ids", context="config")
    _require_non_empty_str(ids, "center_id_14", context="ids")
    _require_non_empty_str(ids, "unit_id_28", context="ids")
    _require_non_empty_str(ids, "onebss_tt_id", context="ids")

    team_ids = raw.get("team_ids", {})
    if team_ids is not None and not isinstance(team_ids, Mapping):
        raise ValueError("Block 'team_ids' phai la object neu duoc khai bao")

    reports = raw.get("reports", {})
    if reports is not None and not isinstance(reports, Mapping):
        raise ValueError("Block 'reports' phai la object neu duoc khai bao")

    download = raw.get("download", {})
    if download is not None and not isinstance(download, Mapping):
        raise ValueError("Block 'download' phai la object neu duoc khai bao")


def load_runtime_context(config_path: str | Path) -> RuntimeContext:
    config_file = Path(config_path).expanduser().resolve()
    if not config_file.exists():
        raise FileNotFoundError(f"Khong tim thay config: {config_file}")

    raw = _read_yaml(config_file)
    validate_runtime_config(raw)

    unit_raw = _require_mapping(raw, "unit", context="config")
    runtime_raw = _require_mapping(raw, "runtime", context="config")
    period_raw = _require_mapping(raw, "period", context="config")
    ids_raw = _require_mapping(raw, "ids", context="config")
    download_raw = raw.get("download", {})
    reports_raw = raw.get("reports", {})
    team_ids_raw = raw.get("team_ids", {})

    paths = build_runtime_paths(_require_non_empty_str(runtime_raw, "instance_root", context="runtime"))
    create_dirs = _coerce_bool(runtime_raw.get("create_dirs", True), field_name="runtime.create_dirs")
    if create_dirs:
        ensure_runtime_dirs(paths)

    normalized_team_ids: Dict[str, Dict[str, str]] = {}
    if team_ids_raw:
        for family, value in team_ids_raw.items():
            family_name = str(family).strip()
            if not family_name:
                raise ValueError("team_ids khong duoc co key rong")
            normalized_team_ids[family_name] = _normalize_string_mapping(
                value,
                field_name=f"team_ids.{family_name}",
            )

    unit = UnitProfile(
        code=_require_non_empty_str(unit_raw, "code", context="unit"),
        name=_require_non_empty_str(unit_raw, "name", context="unit"),
        ids=_normalize_string_mapping(ids_raw, field_name="ids", allow_empty=False),
        team_ids=normalized_team_ids,
    )

    period = PeriodConfig(
        report_month=_require_int(period_raw, "report_month", context="period"),
        report_year=_require_int(period_raw, "report_year", context="period"),
        month_id=_optional_str(period_raw, "month_id"),
        month_label=_optional_str(period_raw, "month_label"),
        vattu_start_date=_optional_str(period_raw, "vattu_start_date"),
    )

    download = DownloadConfig(
        headed=_coerce_bool(download_raw.get("headed", False), field_name="download.headed")
        if download_raw
        else False,
        max_retries=int(download_raw.get("max_retries", 3)) if download_raw else 3,
        retry_timeouts=_coerce_retry_timeouts(download_raw.get("retry_timeouts")) if download_raw else (180, 300, 500),
        retry_delay_seconds=int(download_raw.get("retry_delay_seconds", 3)) if download_raw else 3,
    )

    report_configs = {
        str(key): value
        for key, value in reports_raw.items()
    }
    for report_key, config in report_configs.items():
        if not isinstance(config, Mapping):
            raise ValueError(f"reports.{report_key} phai la object")

    return RuntimeContext(
        config_path=config_file,
        paths=paths,
        unit=unit,
        period=period,
        download=download,
        report_configs=report_configs,
        raw_config=raw,
    )
