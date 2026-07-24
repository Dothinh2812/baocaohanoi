from pathlib import Path

from services.training_report_service import export_report_excel


def test_excel_export_has_required_sheets_and_escapes_formula(tmp_path):
    output = tmp_path / "report.xlsx"
    payload = {
        "summary": {"assigned": 1, "completed": 1},
        "individual": [{"assignment": {"username": "=danger", "display_name": "+name"}, "result": None}],
    }

    export_report_excel(payload, output)

    import openpyxl
    workbook = openpyxl.load_workbook(output, data_only=False)
    assert workbook.sheetnames == ["Summary", "Individual", "Topics", "Questions", "Retake"]
    assert workbook["Individual"][2][0].value == "'=danger"
    assert workbook["Individual"][2][1].value == "'+name"
