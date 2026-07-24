from pathlib import Path
import sys
import os
import time

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import quality_routes


def test_shc_cts_api_reads_summary_and_groups_detail_by_unit(tmp_path, monkeypatch):
    excel_path = tmp_path / "So_sanh_SHC_theo_ngay_T-1.xlsx"
    intraday_old_path = tmp_path / "Bao_cao_tien_trinh_20260514.xlsx"
    intraday_latest_path = tmp_path / "Bao_cao_tien_trinh_20260515.xlsx"
    intraday_temp_path = tmp_path / "~$Bao_cao_tien_trinh_20260516.xlsx"
    summary_df = pd.DataFrame(
        [
            {"Đơn vị": "Tổ A", "SL 13/05": 1, "SL 14/05": 2},
            {"Đơn vị": "Tổ B", "SL 13/05": 3, "SL 14/05": 4},
        ]
    )
    detail_df = pd.DataFrame(
        [
            {"Đơn vị": "Tổ A", "NVKT": "Nguyễn Văn A", "SL 14/05": 2},
            {"Đơn vị": "Tổ B", "NVKT": "Nguyễn Văn B", "SL 14/05": 4},
            {"Đơn vị": "Tổ C", "NVKT": "Nguyễn Văn C", "SL 14/05": 6},
            {"Đơn vị": "Tổ D", "NVKT": "Nguyễn Văn D", "SL 14/05": 8},
        ]
    )
    progress_df = pd.DataFrame(
        [
            {"Đơn vị": "Tổ A", "NVKT_DB": "Nguyễn Văn A", "Tổng số": 2, "Đã xử lý trong ngày": 1},
            {"Đơn vị": "Tổ B", "NVKT_DB": "Nguyễn Văn B", "Tổng số": 4, "Đã xử lý trong ngày": 3},
        ]
    )
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Theo_don_vi", index=False)
        detail_df.to_excel(writer, sheet_name="Chi_tiet_NVKT", index=False)
    with pd.ExcelWriter(intraday_old_path, engine="openpyxl") as writer:
        pd.DataFrame(
            [
                {"Đơn vị": "Tổ C", "NVKT_DB": "Nguyễn Văn C", "Tổng số": 9, "Đã xử lý trong ngày": 0},
            ]
        ).to_excel(writer, sheet_name="Theo NVKT", index=False)
    old_timestamp = time.mktime((2026, 5, 14, 19, 33, 0, 0, 0, -1))
    os.utime(intraday_old_path, (old_timestamp, old_timestamp))
    with pd.ExcelWriter(intraday_latest_path, engine="openpyxl") as writer:
        progress_df.to_excel(writer, sheet_name="Theo NVKT", index=False)
    latest_timestamp = time.mktime((2026, 5, 15, 8, 9, 0, 0, 0, -1))
    os.utime(intraday_latest_path, (latest_timestamp, latest_timestamp))
    intraday_temp_path.write_text("temporary lock file", encoding="utf-8")

    monkeypatch.setattr(quality_routes, "SHC_CTS_REPORT_PATH", str(excel_path))
    monkeypatch.setattr(quality_routes, "SHC_CTS_INTRADAY_REPORT_DIR", str(tmp_path))

    client = app.test_client()
    with client.session_transaction() as session:
        session["username"] = "test-user"

    response = client.get("/api/shc-cts-data")

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["tong_hop"]["columns"] == ["Đơn vị", "SL 13/05", "SL 14/05"]
    assert [row["Đơn vị"] for row in payload["tong_hop"]["data"]] == ["Tổ A", "Tổ B"]
    assert list(payload["don_vi"].keys()) == ["Tổ A", "Tổ B", "Tổ C", "Tổ D"]
    assert payload["don_vi"]["Tổ A"]["data"][0]["NVKT"] == "Nguyễn Văn A"
    assert payload["file_info"]["name"] == "So_sanh_SHC_theo_ngay_T-1.xlsx"
    assert payload["tien_do_xu_ly"]["columns"] == [
        "Đơn vị",
        "Timestamp",
        "NVKT_DB",
        "Tổng số",
        "Đã xử lý trong ngày",
    ]
    assert payload["tien_do_xu_ly"]["data"][0]["NVKT_DB"] == "Nguyễn Văn A"
    assert payload["tien_do_xu_ly"]["data"][0]["Timestamp"] == "15/05/2026 08:09:00"
    assert list(payload["tien_do_theo_don_vi"].keys()) == ["Tổ A", "Tổ B"]
    assert payload["tien_do_theo_don_vi"]["Tổ B"]["data"][0]["NVKT_DB"] == "Nguyễn Văn B"
    assert payload["tien_do_theo_don_vi"]["Tổ B"]["data"][0]["Timestamp"] == "15/05/2026 08:09:00"
    assert payload["tien_do_file_info"]["name"] == "Bao_cao_tien_trinh_20260515.xlsx"


def test_shc_cts_nvkt_detail_uses_latest_k1_directory_for_files_preview_and_download(tmp_path, monkeypatch):
    old_dir = tmp_path / "shc_NVKT_danh_sach_chi_tiet_K1-13-05-2026"
    latest_dir = tmp_path / "shc_NVKT_danh_sach_chi_tiet_K1-14-05-2026"
    old_team_dir = old_dir / "Tổ A"
    latest_team_dir = latest_dir / "Tổ A"
    old_team_dir.mkdir(parents=True)
    latest_team_dir.mkdir(parents=True)

    old_file = old_team_dir / "Nguyễn Văn Cũ.xlsx"
    latest_file = latest_team_dir / "Nguyễn Văn Mới.xlsx"
    pd.DataFrame([{"MA_TB": "OLD"}]).to_excel(old_file, sheet_name="Chi tiết SHC", index=False)
    pd.DataFrame([{"MA_TB": "NEW"}]).to_excel(latest_file, sheet_name="Chi tiết SHC", index=False)

    monkeypatch.setattr(quality_routes, "SHC_CTS_NVKT_DETAIL_ROOT", str(tmp_path))

    client = app.test_client()
    with client.session_transaction() as session:
        session["username"] = "test-user"

    options_response = client.get("/api/shc-cts-nvkt-detail/options")
    assert options_response.status_code == 200
    options_payload = options_response.get_json()
    assert options_payload["source_dir"].endswith("shc_NVKT_danh_sach_chi_tiet_K1-14-05-2026")
    assert options_payload["teams"] == [{"key": "Tổ A", "label": "Tổ A"}]

    files_response = client.get("/api/shc-cts-nvkt-detail/files?team=T%E1%BB%95+A")
    assert files_response.status_code == 200
    assert files_response.get_json()["files"][0]["name"] == "Nguyễn Văn Mới.xlsx"

    preview_response = client.get(
        "/api/shc-cts-nvkt-detail/preview?team=T%E1%BB%95+A&file_name=Nguy%E1%BB%85n+V%C4%83n+M%E1%BB%9Bi.xlsx"
    )
    assert preview_response.status_code == 200
    preview_payload = preview_response.get_json()
    assert preview_payload["sheets"]["Chi tiết SHC"]["data"][0]["MA_TB"] == "NEW"

    download_response = client.get(
        "/download/shc-cts-nvkt-detail/T%E1%BB%95%20A/Nguy%E1%BB%85n%20V%C4%83n%20M%E1%BB%9Bi.xlsx"
    )
    assert download_response.status_code == 200
    assert 'filename="Nguyen Van Moi.xlsx"' in download_response.headers["Content-Disposition"]
    assert "filename*=UTF-8''Nguy%E1%BB%85n%20V%C4%83n%20M%E1%BB%9Bi.xlsx" in download_response.headers["Content-Disposition"]
