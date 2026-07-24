"""Client/UI test (qua node) cho training-question-bank.js.

Kiểm tra hành vi thật: nạp file JS production trong DOM mô phỏng rồi quan sát
các nút thao tác duyệt được render theo trạng thái review/publication. Bảo đảm
transition bất hợp lệ (vd. "Từ chối" câu đã bị rejected) không xuất hiện nút.

Bỏ qua (skip) nếu môi trường không có node.
"""
import shutil
import subprocess
from pathlib import Path

NODE = shutil.which("node")
_SCRIPT = Path(__file__).parent / "js" / "test_question_bank_review_actions.mjs"


def test_question_bank_review_actions_match_legal_transitions():
    if NODE is None:
        import pytest

        pytest.skip("node runtime không có sẵn; bỏ qua client/UI test.")
    result = subprocess.run(
        [NODE, str(_SCRIPT)], capture_output=True, text=True, check=False
    )
    assert result.returncode == 0, (
        "client/UI test thất bại:\n" + result.stdout + result.stderr
    )
