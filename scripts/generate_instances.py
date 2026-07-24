#!/usr/bin/env python3
import argparse
import os
import stat
from dataclasses import dataclass
from pathlib import Path

try:
    import yaml
except ImportError:  # pragma: no cover - PyYAML is available on the target host.
    yaml = None


APP_DIR = "/home/vtst/dashv4"
DEFAULT_ENV_DIR = "/etc/dashv4"
DEFAULT_TUNNEL_ID = "CHANGE-ME-TUNNEL-ID"
DEFAULT_CREDENTIALS_FILE = "/root/.cloudflared/CHANGE-ME-TUNNEL-ID.json"


class UnitConfigError(ValueError):
    pass


@dataclass(frozen=True)
class UnitConfig:
    code: str
    slug: str
    name: str
    port: int
    hostname: str
    db_path: str
    tiep_thi_db_path: str = ""
    disabled_endpoints: tuple[str, ...] = ()
    enabled_endpoints: tuple[str, ...] = ()


def _quote_env(value):
    text = str(value)
    if not text or any(char.isspace() for char in text) or '"' in text or "'" in text:
        return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return text


def _load_yaml(path):
    if yaml is None:
        raise UnitConfigError("PyYAML is required to read units.yaml")
    with open(path, "r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh)
    if not isinstance(data, list):
        raise UnitConfigError("units file must contain a YAML list")
    return data


def _require_text(raw, field, index):
    value = raw.get(field)
    if not isinstance(value, str) or not value.strip():
        raise UnitConfigError(f"unit #{index} missing required field {field}")
    return value.strip()


def _require_port(raw, index):
    value = raw.get("port")
    if not isinstance(value, int):
        raise UnitConfigError(f"unit #{index} missing required integer field port")
    if value < 1 or value > 65535:
        raise UnitConfigError(f"unit #{index} has invalid port {value}")
    return value


def _optional_endpoint_tuple(raw, field, index):
    value = raw.get(field, [])
    if value is None:
        return ()
    if not isinstance(value, list):
        raise UnitConfigError(f"unit #{index} field {field} must be a list")
    endpoints = []
    for endpoint_index, endpoint in enumerate(value, start=1):
        if not isinstance(endpoint, str) or not endpoint.strip():
            raise UnitConfigError(f"unit #{index} field {field} item #{endpoint_index} must be a non-empty string")
        endpoints.append(endpoint.strip())
    return tuple(endpoints)


def _validate_unique(units, attr, label):
    seen = {}
    for unit in units:
        value = getattr(unit, attr)
        if value in seen:
            raise UnitConfigError(f"duplicate {label}: {value}")
        seen[value] = unit.code


def load_units(path):
    raw_units = _load_yaml(path)
    units = []
    for index, raw in enumerate(raw_units, start=1):
        if not isinstance(raw, dict):
            raise UnitConfigError(f"unit #{index} must be a mapping")
        units.append(
            UnitConfig(
                code=_require_text(raw, "code", index),
                slug=_require_text(raw, "slug", index),
                name=_require_text(raw, "name", index),
                port=_require_port(raw, index),
                hostname=_require_text(raw, "hostname", index),
                db_path=_require_text(raw, "db_path", index),
                tiep_thi_db_path=str(raw.get("tiep_thi_db_path") or "").strip(),
                disabled_endpoints=_optional_endpoint_tuple(raw, "disabled_endpoints", index),
                enabled_endpoints=_optional_endpoint_tuple(raw, "enabled_endpoints", index),
            )
        )

    _validate_unique(units, "code", "code")
    _validate_unique(units, "port", "port")
    _validate_unique(units, "hostname", "hostname")
    return units


def render_env(unit, *, app_dir=APP_DIR, secret_key=None):
    runtime_dir = f"{app_dir}/runtime_app/{unit.code}"
    log_dir = f"{app_dir}/logs/{unit.code}"
    secret = secret_key or f"change-me-{unit.code}"
    values = {
        "DASHV4_UNIT_CODE": unit.code,
        "DASHV4_UNIT_NAME": unit.name,
        "DASHV4_APP_NAME": f"Dashboard {unit.name}",
        "DASHV4_HOST": "127.0.0.1",
        "DASHV4_PORT": unit.port,
        "DASHV4_DB_PATH": unit.db_path,
        "DASHV4_SECRET_KEY": secret,
        "DASHV4_TIMEZONE": "Asia/Ho_Chi_Minh",
        "DASHV4_RUNTIME_DIR": runtime_dir,
        "DASHV4_SESSION_FILE_DIR": f"{runtime_dir}/flask_session",
        "DASHV4_CACHE_DIR": f"{runtime_dir}/cache",
        "DASHV4_USER_FILE": f"{runtime_dir}/users.xlsx",
        "DASHV4_LOGIN_LOG_FILE": f"{log_dir}/login_history.csv",
        "DASHV4_LOG_DIR": log_dir,
        "DASHV4_PID_FILE": f"/tmp/dashv4-{unit.code}.pid",
        "DASHV4_SESSION_COOKIE_SECURE": "true",
        "DASHV4_TRAINING_DB_PATH": f"{runtime_dir}/training.db",
        "DASHV4_TRAINING_FILES_DIR": f"{runtime_dir}/training_files",
        "DASHV4_TRAINING_EXPORT_DIR": f"{runtime_dir}/training_export",
        "DASHV4_TRAINING_AI_ENABLED": "false",
    }
    if unit.tiep_thi_db_path:
        values["DASH_TIEP_THI_DB_PATH"] = unit.tiep_thi_db_path
    if unit.disabled_endpoints:
        values["DASHV4_DISABLED_ENDPOINTS"] = ",".join(unit.disabled_endpoints)
    if unit.enabled_endpoints:
        values["DASHV4_ENABLED_ENDPOINTS"] = ",".join(unit.enabled_endpoints)
    lines = [
        "# Generated by scripts/generate_instances.py",
        f"# Unit: {unit.name}",
    ]
    lines.extend(f"{key}={_quote_env(value)}" for key, value in values.items())
    return "\n".join(lines) + "\n"


def render_systemd_service(*, app_dir=APP_DIR, env_dir=DEFAULT_ENV_DIR, gunicorn_bin='/home/vtst/.local/bin/gunicorn'):
    return f"""[Unit]
Description=Dashv4 Dashboard Instance %i
After=network.target

[Service]
Type=simple
WorkingDirectory={app_dir}
EnvironmentFile={env_dir}/%i.env
ExecStart={gunicorn_bin} -c gunicorn_config.py dashboard:app
Restart=always
RestartSec=5
User=vtst

[Install]
WantedBy=multi-user.target
"""


def render_cloudflared_config(units, *, tunnel_id, credentials_file):
    lines = [
        f"tunnel: {tunnel_id}",
        f"credentials-file: {credentials_file}",
        "",
        "ingress:",
    ]
    for unit in units:
        lines.extend(
            [
                f"  - hostname: {unit.hostname}",
                f"    service: http://127.0.0.1:{unit.port}",
                "",
            ]
        )
    lines.append("  - service: http_status:404")
    return "\n".join(lines) + "\n"


def render_smoke_script(units):
    lines = [
        "#!/usr/bin/env bash",
        "set -euo pipefail",
        "",
    ]
    for unit in units:
        lines.extend(
            [
                f'echo "Checking {unit.code} on port {unit.port}"',
                f"curl -fsS http://127.0.0.1:{unit.port}/ >/dev/null",
                "",
            ]
        )
    return "\n".join(lines)


def generate_files(*, units_file, output_dir, tunnel_id=DEFAULT_TUNNEL_ID, credentials_file=DEFAULT_CREDENTIALS_FILE):
    units = load_units(units_file)
    output_path = Path(output_dir)
    env_path = output_path / "env"
    systemd_path = output_path / "systemd"
    cloudflared_path = output_path / "cloudflared"
    for directory in [env_path, systemd_path, cloudflared_path]:
        directory.mkdir(parents=True, exist_ok=True)

    for unit in units:
        (env_path / f"{unit.code}.env").write_text(render_env(unit), encoding="utf-8")

    (systemd_path / "dashv4@.service").write_text(render_systemd_service(), encoding="utf-8")
    (cloudflared_path / "config.yml").write_text(
        render_cloudflared_config(units, tunnel_id=tunnel_id, credentials_file=credentials_file),
        encoding="utf-8",
    )
    smoke_path = output_path / "smoke_test.sh"
    smoke_path.write_text(render_smoke_script(units), encoding="utf-8")
    smoke_path.chmod(smoke_path.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP)
    return output_path


def main(argv=None):
    parser = argparse.ArgumentParser(description="Generate dashv4 multi-instance deployment files.")
    parser.add_argument("--units-file", default="deploy/units.yaml")
    parser.add_argument("--output-dir", default="deploy/generated")
    parser.add_argument("--tunnel-id", default=os.getenv("CLOUDFLARED_TUNNEL_ID", DEFAULT_TUNNEL_ID))
    parser.add_argument(
        "--credentials-file",
        default=os.getenv("CLOUDFLARED_CREDENTIALS_FILE", DEFAULT_CREDENTIALS_FILE),
    )
    args = parser.parse_args(argv)
    output_path = generate_files(
        units_file=Path(args.units_file),
        output_dir=Path(args.output_dir),
        tunnel_id=args.tunnel_id,
        credentials_file=args.credentials_file,
    )
    print(f"Generated deployment files in {output_path}")


if __name__ == "__main__":
    main()
