import sqlite3
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app


def _logged_in_client():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'
    return client


def _prepare_tiep_thi_db(tmp_path):
    db_path = tmp_path / 'bao_cao_tiep_thi.db'
    conn = sqlite3.connect(db_path)
    conn.executescript(
        '''
        CREATE TABLE tiep_thi_daily (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_date TEXT NOT NULL,
            ma_tb TEXT NOT NULL,
            ma_gd TEXT,
            ten_phong_tthi TEXT,
            ten_to TEXT,
            ten_nguoi_tiepthi TEXT,
            loaihinh_tb TEXT,
            source_file TEXT,
            imported_at TEXT NOT NULL,
            UNIQUE(report_date, ma_tb)
        );
        CREATE VIEW v_tiep_thi_quarterly_summary AS
        SELECT
            strftime('%Y', report_date) AS year,
            ((CAST(strftime('%m', report_date) AS INTEGER) - 1) / 3 + 1) AS quarter,
            ten_phong_tthi,
            ten_to,
            ten_nguoi_tiepthi,
            COUNT(*) AS so_ma_tiep_thi
        FROM tiep_thi_daily
        GROUP BY year, quarter, ten_phong_tthi, ten_to, ten_nguoi_tiepthi;
        CREATE VIEW v_tiep_thi_yearly_summary AS
        SELECT
            strftime('%Y', report_date) AS year,
            ten_phong_tthi,
            ten_to,
            ten_nguoi_tiepthi,
            COUNT(*) AS so_ma_tiep_thi
        FROM tiep_thi_daily
        GROUP BY year, ten_phong_tthi, ten_to, ten_nguoi_tiepthi;
        '''
    )
    conn.executemany(
        '''
        INSERT INTO tiep_thi_daily (
            report_date, ma_tb, ma_gd, ten_phong_tthi, ten_to,
            ten_nguoi_tiepthi, loaihinh_tb, source_file, imported_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''',
        [
            (
                '2026-06-01',
                'fbr000',
                'GD-0',
                'Trung tâm Viễn thông Sơn Tây',
                'Tổ B',
                'Nguyễn Văn B',
                'MyTV',
                'source.xlsx',
                '2026-06-08T08:00:00',
            ),
            (
                '2026-06-07',
                'fbr001',
                'GD-1',
                'Trung tâm Viễn thông Sơn Tây',
                'Tổ A',
                'Nguyễn Văn A',
                'Fiber',
                'source.xlsx',
                '2026-06-08T08:00:00',
            ),
            (
                '2026-06-07',
                'fbr002',
                'GD-2',
                'Trung tâm Viễn thông Sơn Tây',
                'Tổ A',
                'Nguyễn Văn A',
                'MyTV',
                'source.xlsx',
                '2026-06-08T08:00:00',
            ),
            (
                '2026-06-06',
                'fbr003',
                'GD-3',
                'Trung tâm Viễn thông Sơn Tây',
                'Tổ B',
                'Nguyễn Văn B',
                'Fiber',
                'source.xlsx',
                '2026-06-08T08:00:00',
            ),
            (
                '2026-04-15',
                'fbr004',
                'GD-4',
                'Trung tâm Viễn thông Sơn Tây',
                'Tổ B',
                'Nguyễn Văn B',
                'Fiber',
                'source.xlsx',
                '2026-06-08T08:00:00',
            ),
            (
                '2026-01-15',
                'fbr005',
                'GD-5',
                'Trung tâm Viễn thông Sơn Tây',
                'Tổ C',
                'Nguyễn Văn C',
                'Wifi Mesh',
                'source.xlsx',
                '2026-06-08T08:00:00',
            ),
            (
                '2025-12-15',
                'fbr006',
                'GD-6',
                'Trung tâm Viễn thông Sơn Tây',
                'Tổ D',
                'Nguyễn Văn D',
                'Fiber',
                'source.xlsx',
                '2026-06-08T08:00:00',
            ),
        ],
    )
    conn.commit()
    conn.close()
    return db_path


def test_tiepthi_api_reads_configured_new_source_db(tmp_path):
    db_path = _prepare_tiep_thi_db(tmp_path)

    old_path = app.config.get('TIEP_THI_DB_PATH')
    app.config['TIEP_THI_DB_PATH'] = str(db_path)
    try:
        response = _logged_in_client().get('/api/tiepthi-data')
    finally:
        app.config['TIEP_THI_DB_PATH'] = old_path

    assert response.status_code == 200
    payload = response.get_json()
    assert payload['file_info']['name'] == 'bao_cao_tiep_thi.db'
    assert payload['selected_date'] == '2026-06-07'
    assert payload['latest_available_date'] == '2026-06-07'
    assert payload['available_dates'] == [
        '2026-06-07',
        '2026-06-06',
        '2026-06-01',
        '2026-04-15',
        '2026-01-15',
        '2025-12-15',
    ]
    assert payload['date_has_data'] is True
    assert payload['tong_hop']['columns'] == ['STT', 'Đơn vị', 'Người tiếp thị', 'Số mã tiếp thị']
    assert payload['tong_hop']['data'] == [
        {'STT': 1, 'Đơn vị': 'Tổ A', 'Người tiếp thị': 'Nguyễn Văn A', 'Số mã tiếp thị': 2},
        {'STT': 2, 'Đơn vị': 'Tổ B', 'Người tiếp thị': 'Nguyễn Văn B', 'Số mã tiếp thị': 2},
        {'STT': '', 'Đơn vị': 'TỔNG CỘNG', 'Người tiếp thị': '', 'Số mã tiếp thị': 4},
    ]
    assert payload['tong_hop_don_vi']['columns'] == ['STT', 'Đơn vị', 'Số mã tiếp thị']
    assert payload['tong_hop_don_vi']['data'] == [
        {'STT': 1, 'Đơn vị': 'Tổ A', 'Số mã tiếp thị': 2},
        {'STT': 2, 'Đơn vị': 'Tổ B', 'Số mã tiếp thị': 2},
        {'STT': '', 'Đơn vị': 'TỔNG CỘNG', 'Số mã tiếp thị': 4},
    ]
    assert payload['tong_hop_loai_dich_vu']['columns'] == ['STT', 'Loại dịch vụ', 'Số mã tiếp thị']
    assert payload['tong_hop_loai_dich_vu']['data'] == [
        {'STT': 1, 'Loại dịch vụ': 'Fiber', 'Số mã tiếp thị': 2},
        {'STT': 2, 'Loại dịch vụ': 'MyTV', 'Số mã tiếp thị': 2},
        {'STT': '', 'Loại dịch vụ': 'TỔNG CỘNG', 'Số mã tiếp thị': 4},
    ]
    assert payload['tong_hop_don_vi_loai_dich_vu']['columns'] == [
        'STT',
        'Đơn vị',
        'Loại dịch vụ',
        'Số mã tiếp thị',
    ]
    assert payload['tong_hop_don_vi_loai_dich_vu']['data'] == [
        {'STT': 1, 'Đơn vị': 'Tổ A', 'Loại dịch vụ': 'Fiber', 'Số mã tiếp thị': 1},
        {'STT': 2, 'Đơn vị': 'Tổ A', 'Loại dịch vụ': 'MyTV', 'Số mã tiếp thị': 1},
        {'STT': 3, 'Đơn vị': 'Tổ B', 'Loại dịch vụ': 'Fiber', 'Số mã tiếp thị': 1},
        {'STT': 4, 'Đơn vị': 'Tổ B', 'Loại dịch vụ': 'MyTV', 'Số mã tiếp thị': 1},
        {'STT': '', 'Đơn vị': 'TỔNG CỘNG', 'Loại dịch vụ': '', 'Số mã tiếp thị': 4},
    ]
    assert payload['tong_hop_nguoi_tiep_thi_loai_dich_vu']['columns'] == [
        'STT',
        'Đơn vị',
        'Người tiếp thị',
        'Loại dịch vụ',
        'Số mã tiếp thị',
    ]
    assert payload['tong_hop_nguoi_tiep_thi_loai_dich_vu']['data'] == [
        {'STT': 1, 'Đơn vị': 'Tổ A', 'Người tiếp thị': 'Nguyễn Văn A', 'Loại dịch vụ': 'Fiber', 'Số mã tiếp thị': 1},
        {'STT': 2, 'Đơn vị': 'Tổ A', 'Người tiếp thị': 'Nguyễn Văn A', 'Loại dịch vụ': 'MyTV', 'Số mã tiếp thị': 1},
        {'STT': 3, 'Đơn vị': 'Tổ B', 'Người tiếp thị': 'Nguyễn Văn B', 'Loại dịch vụ': 'Fiber', 'Số mã tiếp thị': 1},
        {'STT': 4, 'Đơn vị': 'Tổ B', 'Người tiếp thị': 'Nguyễn Văn B', 'Loại dịch vụ': 'MyTV', 'Số mã tiếp thị': 1},
        {'STT': '', 'Đơn vị': 'TỔNG CỘNG', 'Người tiếp thị': '', 'Loại dịch vụ': '', 'Số mã tiếp thị': 4},
    ]
    assert payload['tong_hop_quy']['columns'] == ['STT', 'Năm', 'Quý', 'Đơn vị', 'Người tiếp thị', 'Số mã tiếp thị']
    assert payload['tong_hop_quy']['data'] == [
        {'STT': 1, 'Năm': '2026', 'Quý': 2, 'Đơn vị': 'Tổ A', 'Người tiếp thị': 'Nguyễn Văn A', 'Số mã tiếp thị': 2},
        {'STT': 2, 'Năm': '2026', 'Quý': 2, 'Đơn vị': 'Tổ B', 'Người tiếp thị': 'Nguyễn Văn B', 'Số mã tiếp thị': 3},
        {'STT': '', 'Năm': '2026', 'Quý': 2, 'Đơn vị': 'TỔNG CỘNG', 'Người tiếp thị': '', 'Số mã tiếp thị': 5},
    ]
    assert payload['tong_hop_nam']['columns'] == ['STT', 'Năm', 'Đơn vị', 'Người tiếp thị', 'Số mã tiếp thị']
    assert payload['tong_hop_nam']['data'] == [
        {'STT': 1, 'Năm': '2026', 'Đơn vị': 'Tổ A', 'Người tiếp thị': 'Nguyễn Văn A', 'Số mã tiếp thị': 2},
        {'STT': 2, 'Năm': '2026', 'Đơn vị': 'Tổ B', 'Người tiếp thị': 'Nguyễn Văn B', 'Số mã tiếp thị': 3},
        {'STT': 3, 'Năm': '2026', 'Đơn vị': 'Tổ C', 'Người tiếp thị': 'Nguyễn Văn C', 'Số mã tiếp thị': 1},
        {'STT': '', 'Năm': '2026', 'Đơn vị': 'TỔNG CỘNG', 'Người tiếp thị': '', 'Số mã tiếp thị': 6},
    ]
    assert list(payload['sheets'].keys()) == ['Tất cả', 'Tổ A', 'Tổ B']
    assert payload['sheets']['Tất cả']['columns'] == [
        'STT',
        'Ngày',
        'Mã TB',
        'Mã GD',
        'Phòng tiếp thị',
        'Đơn vị',
        'Người tiếp thị',
        'Loại dịch vụ',
    ]
    assert payload['sheets']['Tất cả']['data'][0]['Mã TB'] == 'fbr001'
    assert len(payload['sheets']['Tổ A']['data']) == 2
    assert len(payload['sheets']['Tổ B']['data']) == 2


def test_tiepthi_api_reads_requested_date_from_new_source_db(tmp_path):
    db_path = _prepare_tiep_thi_db(tmp_path)

    old_path = app.config.get('TIEP_THI_DB_PATH')
    app.config['TIEP_THI_DB_PATH'] = str(db_path)
    try:
        response = _logged_in_client().get('/api/tiepthi-data?date=2026-06-06')
    finally:
        app.config['TIEP_THI_DB_PATH'] = old_path

    assert response.status_code == 200
    payload = response.get_json()
    assert payload['selected_date'] == '2026-06-06'
    assert payload['date_has_data'] is True
    assert payload['tong_hop']['data'] == [
        {'STT': 1, 'Đơn vị': 'Tổ B', 'Người tiếp thị': 'Nguyễn Văn B', 'Số mã tiếp thị': 2},
        {'STT': '', 'Đơn vị': 'TỔNG CỘNG', 'Người tiếp thị': '', 'Số mã tiếp thị': 2},
    ]
    assert payload['tong_hop_don_vi']['data'] == [
        {'STT': 1, 'Đơn vị': 'Tổ B', 'Số mã tiếp thị': 2},
        {'STT': '', 'Đơn vị': 'TỔNG CỘNG', 'Số mã tiếp thị': 2},
    ]


def test_tiepthi_api_returns_empty_payload_when_configured_date_has_no_new_source_data(tmp_path):
    db_path = _prepare_tiep_thi_db(tmp_path)

    old_path = app.config.get('TIEP_THI_DB_PATH')
    app.config['TIEP_THI_DB_PATH'] = str(db_path)
    try:
        response = _logged_in_client().get('/api/tiepthi-data?date=2026-06-05')
    finally:
        app.config['TIEP_THI_DB_PATH'] = old_path

    assert response.status_code == 200
    payload = response.get_json()
    assert payload['selected_date'] == '2026-06-05'
    assert payload['date_has_data'] is False
    assert payload['latest_available_date'] == '2026-06-07'
    assert payload['tong_hop']['data'] == []
    assert payload['tong_hop_don_vi']['data'] == []
    assert payload['tong_hop_quy']['data'] == []
    assert payload['tong_hop_nam']['data'] == []
    assert payload['sheets']['Tất cả']['data'] == []
