"""Client/UI test (qua node) cho training-exams.js.

Kiểm tra hành vi thật: nạp file JS production trong DOM mô phỏng rồi quan sát
các nút thao tác vòng đời kỳ thi (draft/ready/open/closed/cancelled) được render
đúng theo transition hợp lệ ở backend. Bảo đảm kỳ thi đã chốt (finalized) hoặc
đã hủy không hiện nút hành động nào.

Bỏ qua (skip) nếu môi trường không có node.
"""
import shutil
import subprocess
from pathlib import Path

NODE = shutil.which("node")
_SCRIPT = Path(__file__).parent / "js" / "test_training_exams_lifecycle.mjs"


def test_exam_lifecycle_buttons_match_legal_transitions():
    if NODE is None:
        import pytest

        pytest.skip("node runtime không có sẵn; bỏ qua client/UI test.")
    result = subprocess.run(
        [NODE, str(_SCRIPT)], capture_output=True, text=True, check=False, timeout=30
    )
    assert result.returncode == 0, (
        "client/UI test thất bại:\n" + result.stdout + result.stderr
    )
