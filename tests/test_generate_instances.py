from pathlib import Path
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts import generate_instances


SAMPLE_UNITS = """
- code: son_tay
  slug: son-tay
  name: "TTVT Sơn Tây"
  port: 5011
  hostname: son-tay.example.vn
  db_path: /runtime/son_tay/sqlite_history/report_history.db
- code: ba_dinh
  slug: ba-dinh
  name: "TTVT Ba Đình"
  port: 5012
  hostname: ba-dinh.example.vn
  db_path: /runtime/ba_dinh/sqlite_history/report_history.db
"""


def test_load_units_and_validate_required_fields(tmp_path):
    units_file = tmp_path / "units.yaml"
    units_file.write_text(SAMPLE_UNITS, encoding="utf-8")

    units = generate_instances.load_units(units_file)

    assert [unit.code for unit in units] == ["son_tay", "ba_dinh"]
    assert units[0].name == "TTVT Sơn Tây"
    assert units[1].port == 5012
    assert units[1].hostname == "ba-dinh.example.vn"


def test_load_units_rejects_duplicate_ports(tmp_path):
    units_file = tmp_path / "units.yaml"
    units_file.write_text(
        SAMPLE_UNITS.replace("port: 5012", "port: 5011"),
        encoding="utf-8",
    )

    with pytest.raises(generate_instances.UnitConfigError, match="duplicate port"):
        generate_instances.load_units(units_file)


def test_render_env_file_contains_instance_isolation_paths():
    unit = generate_instances.UnitConfig(
        code="ba_dinh",
        slug="ba-dinh",
        name="TTVT Ba Đình",
        port=5012,
        hostname="ba-dinh.example.vn",
        db_path="/runtime/ba_dinh/sqlite_history/report_history.db",
    )

    content = generate_instances.render_env(unit)

    assert "DASHV4_UNIT_CODE=ba_dinh" in content
    assert 'DASHV4_UNIT_NAME="TTVT Ba Đình"' in content
    assert "DASHV4_HOST=127.0.0.1" in content
    assert "DASHV4_PORT=5012" in content
    assert "DASHV4_DB_PATH=/runtime/ba_dinh/sqlite_history/report_history.db" in content
    assert "DASHV4_USER_FILE=/home/vtst/dashv4/runtime_app/ba_dinh/users.xlsx" in content
    assert "DASHV4_LOGIN_LOG_FILE=/home/vtst/dashv4/logs/ba_dinh/login_history.csv" in content
    assert "DASHV4_SESSION_COOKIE_SECURE=true" in content


def test_render_cloudflared_ingress_routes_hostnames_to_ports(tmp_path):
    units_file = tmp_path / "units.yaml"
    units_file.write_text(SAMPLE_UNITS, encoding="utf-8")
    units = generate_instances.load_units(units_file)

    content = generate_instances.render_cloudflared_config(
        units,
        tunnel_id="demo-tunnel",
        credentials_file="/root/.cloudflared/demo-tunnel.json",
    )

    assert "tunnel: demo-tunnel" in content
    assert "hostname: son-tay.example.vn" in content
    assert "service: http://127.0.0.1:5011" in content
    assert "hostname: ba-dinh.example.vn" in content
    assert "service: http://127.0.0.1:5012" in content
    assert content.rstrip().endswith("- service: http_status:404")


def test_generate_files_writes_env_systemd_cloudflared_and_smoke_script(tmp_path):
    units_file = tmp_path / "units.yaml"
    output_dir = tmp_path / "generated"
    units_file.write_text(SAMPLE_UNITS, encoding="utf-8")

    generate_instances.generate_files(
        units_file=units_file,
        output_dir=output_dir,
        tunnel_id="demo-tunnel",
        credentials_file="/root/.cloudflared/demo-tunnel.json",
    )

    assert (output_dir / "env" / "son_tay.env").exists()
    assert (output_dir / "env" / "ba_dinh.env").exists()
    assert (output_dir / "systemd" / "dashv4@.service").exists()
    assert (output_dir / "cloudflared" / "config.yml").exists()
    smoke_script = output_dir / "smoke_test.sh"
    assert smoke_script.exists()
    assert "curl -fsS http://127.0.0.1:5011/" in smoke_script.read_text(encoding="utf-8")


def test_load_units_accepts_optional_route_policy_fields(tmp_path):
    units_file = tmp_path / "units.yaml"
    units_file.write_text(
        """
- code: ba_dinh
  slug: ba-dinh
  name: "TTVT Ba Đình"
  port: 5012
  hostname: ba-dinh.example.vn
  db_path: /runtime/ba_dinh/sqlite_history/report_history.db
  disabled_endpoints:
    - quangchudong.page_quangchudong
    - sa_outage.page_su_co_sa
  enabled_endpoints:
    - quality.page_shc_processing
""",
        encoding="utf-8",
    )

    units = generate_instances.load_units(units_file)

    assert units[0].disabled_endpoints == (
        "quangchudong.page_quangchudong",
        "sa_outage.page_su_co_sa",
    )
    assert units[0].enabled_endpoints == ("quality.page_shc_processing",)


def test_render_env_file_contains_route_policy_when_configured():
    unit = generate_instances.UnitConfig(
        code="ba_dinh",
        slug="ba-dinh",
        name="TTVT Ba Đình",
        port=5012,
        hostname="ba-dinh.example.vn",
        db_path="/runtime/ba_dinh/sqlite_history/report_history.db",
        disabled_endpoints=(
            "quangchudong.page_quangchudong",
            "sa_outage.page_su_co_sa",
        ),
        enabled_endpoints=("quality.page_shc_processing",),
    )

    content = generate_instances.render_env(unit)

    assert "DASHV4_DISABLED_ENDPOINTS=quangchudong.page_quangchudong,sa_outage.page_su_co_sa" in content
    assert "DASHV4_ENABLED_ENDPOINTS=quality.page_shc_processing" in content


def test_load_units_accepts_optional_tiep_thi_db_path(tmp_path):
    units_file = tmp_path / "units.yaml"
    units_file.write_text(
        """
- code: son_tay
  slug: son-tay
  name: "TTVT Sơn Tây"
  port: 5011
  hostname: son-tay.example.vn
  db_path: /runtime/son_tay/sqlite_history/report_history.db
  tiep_thi_db_path: /home/vtst/onebss_bao_cao/data/bao_cao_tiep_thi.db
""",
        encoding="utf-8",
    )

    units = generate_instances.load_units(units_file)

    assert units[0].tiep_thi_db_path == "/home/vtst/onebss_bao_cao/data/bao_cao_tiep_thi.db"


def test_render_env_file_contains_tiep_thi_db_path_when_configured():
    unit = generate_instances.UnitConfig(
        code="son_tay",
        slug="son-tay",
        name="TTVT Sơn Tây",
        port=5011,
        hostname="son-tay.example.vn",
        db_path="/runtime/son_tay/sqlite_history/report_history.db",
        tiep_thi_db_path="/home/vtst/onebss_bao_cao/data/bao_cao_tiep_thi.db",
    )

    content = generate_instances.render_env(unit)

    assert "DASH_TIEP_THI_DB_PATH=/home/vtst/onebss_bao_cao/data/bao_cao_tiep_thi.db" in content


def test_render_env_file_omits_tiep_thi_db_path_when_unconfigured():
    unit = generate_instances.UnitConfig(
        code="ba_dinh",
        slug="ba-dinh",
        name="TTVT Ba Đình",
        port=5012,
        hostname="ba-dinh.example.vn",
        db_path="/runtime/ba_dinh/sqlite_history/report_history.db",
    )

    content = generate_instances.render_env(unit)

    assert "DASH_TIEP_THI_DB_PATH" not in content


def test_render_env_file_contains_training_paths_and_ai_disabled_default():
    unit = generate_instances.UnitConfig(
        code="son_tay",
        slug="son-tay",
        name="TTVT Sơn Tây",
        port=5011,
        hostname="son-tay.example.vn",
        db_path="/runtime/son_tay/sqlite_history/report_history.db",
    )

    content = generate_instances.render_env(unit)

    assert "DASHV4_TRAINING_DB_PATH=/home/vtst/dashv4/runtime_app/son_tay/training.db" in content
    assert "DASHV4_TRAINING_FILES_DIR=/home/vtst/dashv4/runtime_app/son_tay/training_files" in content
    assert "DASHV4_TRAINING_EXPORT_DIR=/home/vtst/dashv4/runtime_app/son_tay/training_export" in content
    assert "DASHV4_TRAINING_AI_ENABLED=false" in content


def test_render_env_file_does_not_contain_api_key():
    unit = generate_instances.UnitConfig(
        code="son_tay",
        slug="son-tay",
        name="TTVT Sơn Tây",
        port=5011,
        hostname="son-tay.example.vn",
        db_path="/runtime/son_tay/sqlite_history/report_history.db",
    )

    content = generate_instances.render_env(unit)

    assert "API_KEY" not in content.upper()
    assert "SECRET" not in content.replace("DASHV4_SECRET_KEY", "")
