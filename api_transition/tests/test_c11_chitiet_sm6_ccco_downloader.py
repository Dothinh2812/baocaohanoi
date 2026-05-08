import json
from pathlib import Path


def test_c11_chitiet_sm6_ccco_recipe_uses_pcot_6_contract():
    recipe_path = (
        Path(__file__).resolve().parents[1]
        / "recipes"
        / "c11_chitiet_sm6_ccco_q2_2026.json"
    )
    recipe = json.loads(recipe_path.read_text(encoding="utf-8"))

    assert recipe["name"] == "c11_chitiet_sm6_ccco_q2_2026"
    assert recipe["report_id"] == "535080"
    assert recipe["export_payload"]["reportId"] == "535080"

    input_params = {
        param["name"]: param["value"]
        for param in recipe["export_payload"]["lstInputParams"]
    }
    assert input_params == {
        "ptrungtamid": "14324",
        "pthang": "98944805",
        "ploaict": "1",
        "ploai": "1",
        "pcot": "6",
        "ppageindex": "1",
        "ppagesize": "1000000",
    }
    assert recipe["export_payload"]["lstOutputParams"][1]["name"] == "odata"


def test_c11_chitiet_sm6_ccco_downloader_passes_month_and_center_overrides(monkeypatch):
    from api_transition import downloaders

    captured = {}

    def fake_download_with_recipe(*args, **kwargs):
        captured["args"] = args
        captured["kwargs"] = kwargs
        return "/tmp/c1.1_chitiet_sm6_ccco_report.xlsx"

    monkeypatch.setattr(downloaders, "download_with_recipe", fake_download_with_recipe)

    result = downloaders.download_report_c11_chitiet_sm6_ccco_api(
        month_id="98944805",
        month_label="Tháng 05/2026",
        unit_id="14324",
        headed=True,
        output_dir="/tmp/raw",
        session={"headers": {}, "api_timeout": 1},
    )

    assert result == "/tmp/c1.1_chitiet_sm6_ccco_report.xlsx"
    assert captured["args"] == ("c11_chitiet_sm6_ccco_q2_2026",)
    assert captured["kwargs"]["overrides"] == {"ptrungtamid": "14324"}
    assert captured["kwargs"]["month_id"] == "98944805"
    assert captured["kwargs"]["month_label"] == "Tháng 05/2026"
    assert captured["kwargs"]["month_override_key"] == "pthang"
    assert captured["kwargs"]["output_name"] == "c1.1_chitiet_sm6_ccco_report.xlsx"


def test_c11_chitiet_sm6_ccco_batch_task_uses_month_center_id():
    from api_transition.batch_download import REPORT_TASKS

    task = next(
        task for task in REPORT_TASKS if task.report_key == "c11_chi_tiet_sm6_ccco"
    )

    assert task.name == "C1.1 Chi tiết SM6 CCCO"
    assert task.params_type == "month"
    assert task.group == "chi_tieu_c"
    assert task.id_family == "center_id_14"
