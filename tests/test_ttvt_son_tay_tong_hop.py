from pathlib import Path
import sys

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import operations_routes


def test_build_chat_luong_c_table_merges_selected_bsc_columns():
    date_context = {
        'selected_date': '2026-04-22',
        'latest_available_date': '2026-04-22',
        'available_dates': ['2026-04-22'],
        'date_has_data': True,
    }

    frames = {
        'c1_1': pd.DataFrame([
            {'Đơn vị': 'Tổ Kỹ thuật Địa bàn Sơn Tây', 'Chỉ tiêu BSC': 1.8, 'SM1': 11},
            {'Đơn vị': 'Tổng', 'Chỉ tiêu BSC': 2.3, 'SM1': 22},
        ]),
        'c1_2': pd.DataFrame([
            {'Đơn vị': 'Tổ Kỹ thuật Địa bàn Sơn Tây', 'Chỉ tiêu BSC': 1.0},
            {'Đơn vị': 'Tổng', 'Chỉ tiêu BSC': 1.69},
        ]),
        'c1_3': pd.DataFrame([
            {'Đơn vị': 'Tổ Kỹ thuật Địa bàn Sơn Tây', 'Chỉ tiêu BSC': 5.0},
            {'Đơn vị': 'Tổng', 'Chỉ tiêu BSC': 4.2},
        ]),
        'c1_4': pd.DataFrame([
            {'Đơn vị': 'Tổ Kỹ thuật Địa bàn Sơn Tây', 'Điểm BSC': 1.62},
            {'Đơn vị': 'Tổng', 'Điểm BSC': 1.46},
        ]),
        'c1_5': pd.DataFrame([
            {'Đơn vị': 'Tổ Kỹ thuật Địa bàn Sơn Tây', 'Tổng - Điểm BSC': 2.77},
            {'Đơn vị': 'Tổng', 'Tổng - Điểm BSC': 3.99},
        ]),
    }

    sheet = operations_routes._build_ttvt_son_tay_chat_luong_c_sheet(date_context, frames)

    assert sheet['columns'] == [
        'Đơn vị',
        'C1.1 - Chỉ tiêu BSC',
        'C1.2 - Chỉ tiêu BSC',
        'C1.3 - Chỉ tiêu BSC',
        'C1.4 - Điểm BSC',
        'C1.5 - Tổng - Điểm BSC',
    ]
    assert sheet['data'][0] == {
        'Đơn vị': 'Tổ Kỹ thuật Địa bàn Sơn Tây',
        'C1.1 - Chỉ tiêu BSC': 1.8,
        'C1.2 - Chỉ tiêu BSC': 1.0,
        'C1.3 - Chỉ tiêu BSC': 5.0,
        'C1.4 - Điểm BSC': 1.62,
        'C1.5 - Tổng - Điểm BSC': 2.77,
    }
    assert sheet['data'][1] == {
        'Đơn vị': 'Tổng',
        'C1.1 - Chỉ tiêu BSC': 2.3,
        'C1.2 - Chỉ tiêu BSC': 1.69,
        'C1.3 - Chỉ tiêu BSC': 4.2,
        'C1.4 - Điểm BSC': 1.46,
        'C1.5 - Tổng - Điểm BSC': 3.99,
    }


def test_build_ttvt_son_tay_payload_includes_chat_luong_date_state(monkeypatch):
    date_context = {
        'selected_date': '2026-04-22',
        'latest_available_date': '2026-04-22',
        'available_dates': ['2026-04-22', '2026-04-21'],
        'date_has_data': True,
    }

    monkeypatch.setattr(
        operations_routes,
        '_load_ttvt_son_tay_tong_hop_rows',
        lambda: [{
            'ngay_du_lieu': '2026-04-22',
            'ttvt': 'TTVT Sơn Tây',
            'nhom_du_lieu': 'chat_luong',
            'nhom_chi_tieu': 'c11',
            'don_vi': 'Tổ Kỹ thuật Địa bàn Sơn Tây',
            'ten_chi_so': 'ty_le_sua_chua_dung_han',
            'gia_tri_so': 99.5,
            'chi_tieu_bsc': 1.8,
            'nguon_view': 'v_demo',
        }],
    )
    monkeypatch.setattr(
        operations_routes,
        '_load_ttvt_son_tay_chat_luong_c_frames',
        lambda selected_date: {
            'c1_1': pd.DataFrame([{'Đơn vị': 'Tổng', 'Chỉ tiêu BSC': 2.3}]),
            'c1_2': pd.DataFrame([{'Đơn vị': 'Tổng', 'Chỉ tiêu BSC': 1.69}]),
            'c1_3': pd.DataFrame([{'Đơn vị': 'Tổng', 'Chỉ tiêu BSC': 4.2}]),
            'c1_4': pd.DataFrame([{'Đơn vị': 'Tổng', 'Điểm BSC': 1.46}]),
            'c1_5': pd.DataFrame([{'Đơn vị': 'Tổng', 'Tổng - Điểm BSC': 3.99}]),
        },
    )
    monkeypatch.setattr(
        operations_routes,
        'build_file_info',
        lambda *_args, **_kwargs: {'name': 'report_history.db', 'modified': '2026-04-23 08:00:00'},
    )

    payload = operations_routes._build_ttvt_son_tay_tong_hop_payload(date_context)

    assert payload['selected_date'] == '2026-04-22'
    assert payload['latest_available_date'] == '2026-04-22'
    assert payload['available_dates'] == ['2026-04-22', '2026-04-21']
    assert payload['date_has_data'] is True
    assert payload['chat_luong_c']['columns'][0] == 'Đơn vị'
    assert payload['chat_luong_c']['data'][0] == {
        'Đơn vị': 'Tổng',
        'C1.1 - Chỉ tiêu BSC': 2.3,
        'C1.2 - Chỉ tiêu BSC': 1.69,
        'C1.3 - Chỉ tiêu BSC': 4.2,
        'C1.4 - Điểm BSC': 1.46,
        'C1.5 - Tổng - Điểm BSC': 3.99,
    }
    assert sorted(payload.keys()) == [
        'available_dates',
        'chat_luong_c',
        'date_has_data',
        'file_info',
        'latest_available_date',
        'selected_date',
    ]


def test_ttvt_son_tay_page_keeps_only_chat_luong_c_section():
    with app.test_request_context('/ttvt-son-tay-tong-hop'):
        from flask import session
        session['username'] = 'test-user'
        html = operations_routes.page_ttvt_son_tay_tong_hop()

    assert 'Chỉ tiêu chất lượng C' in html
    assert 'Bộ lọc ngày dữ liệu' in html
    assert 'Chỉ số nổi bật' not in html
    assert 'Điểm BSC chất lượng' not in html
    assert 'Tóm tắt theo nhóm dữ liệu' not in html
    assert 'Chi tiết toàn bộ chỉ tiêu' not in html
    assert 'class="excel-table-card" id="ttvt-son-tay-chat-luong-c"' not in html
