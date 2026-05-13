import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from flask import Flask

import route_policy


def _app(disabled=None, enabled=None):
    app = Flask(__name__)
    app.config["INSTANCE_DISABLED_ENDPOINTS"] = set(disabled or [])
    app.config["INSTANCE_ENABLED_ENDPOINTS"] = set(enabled or [])
    return app


def test_global_disabled_page_is_blocked_by_default():
    app = _app()

    with app.app_context():
        feature = route_policy.get_disabled_feature("quality.page_shc_processing")
        enabled = route_policy.is_endpoint_enabled("quality.page_shc_processing")

    assert feature["kind"] == "page"
    assert feature["title"] == "SHC Processing"
    assert enabled is False


def test_instance_enabled_endpoint_unlocks_global_disabled_page():
    app = _app(enabled={"quality.page_shc_processing"})

    with app.app_context():
        feature = route_policy.get_disabled_feature("quality.page_shc_processing")
        enabled = route_policy.is_endpoint_enabled("quality.page_shc_processing")

    assert feature is None
    assert enabled is True


def test_instance_disabled_endpoint_blocks_page_with_generic_metadata():
    app = _app(disabled={"quangchudong.page_quangchudong"})

    with app.app_context():
        feature = route_policy.get_disabled_feature("quangchudong.page_quangchudong")
        enabled = route_policy.is_endpoint_enabled("quangchudong.page_quangchudong")

    assert feature == {
        "kind": "page",
        "title": "quangchudong.page_quangchudong",
        "active_page": "quangchudong",
        "reason": "Route này đang bị khóa trong cấu hình instance hiện tại.",
    }
    assert enabled is False


def test_instance_disabled_endpoint_blocks_api_with_generic_metadata():
    app = _app(disabled={"quangchudong.get_quangchudong_dashboard"})

    with app.app_context():
        feature = route_policy.get_disabled_feature("quangchudong.get_quangchudong_dashboard")
        enabled = route_policy.is_endpoint_enabled("quangchudong.get_quangchudong_dashboard")

    assert feature == {
        "kind": "api",
        "title": "quangchudong.get_quangchudong_dashboard",
        "reason": "Route này đang bị khóa trong cấu hình instance hiện tại.",
        "required_display_contract": {},
    }
    assert enabled is False
