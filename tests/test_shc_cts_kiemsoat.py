import sqlite3
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import quality_routes


_PROGRESS_COLUMNS = [
    'Đơn vị', 'NVKT_DB', 'Tổng số', 'Đạt baseline',
    'Đã xử lý trong ngày', 'Tổng đã đạt', 'Chưa đạt', 'OFF/Lỗi', '% đạt',
]


def _prepare_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'shc_cts.db'
    monkeypatch.setattr(quality_routes, 'SHC_CTS_HISTORY_DB_PATH', str(db_path))
    quality_routes._shc_cts_schema_ready_path = None
    quality_routes._ensure_shc_cts_schema()
    return db_path


def _logged_in_client():
    client = app.test_client()
    with client.session_transaction() as sess:
        sess['username'] = 'test-user'
    return client


def _write_intraday(tmp_path, monkeypatch, rows, filename='Bao_cao_tien_trinh_20260625.xlsx'):
    intraday_path = tmp_path / filename
    df = pd.DataFrame(rows, columns=_PROGRESS_COLUMNS)
    with pd.ExcelWriter(intraday_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Theo NVKT', index=False)
    monkeypatch.setattr(quality_routes, 'SHC_CTS_INTRADAY_REPORT_DIR', str(tmp_path))
    return intraday_path


def _progress_rows():
    return [
        {
            'Đơn vị': 'Tổ Kỹ thuật Địa bàn Phúc Thọ', 'NVKT_DB': 'Nguyễn Văn A',
            'Tổng số': 10, 'Đạt baseline': 5, 'Đã xử lý trong ngày': 3,
            'Tổng đã đạt': 6, 'Chưa đạt': 4, 'OFF/Lỗi': 0, '% đạt': 60,
        },
        {
            'Đơn vị': 'Tổ Kỹ thuật Địa bàn Phúc Thọ', 'NVKT_DB': 'Trần Văn B',
            'Tổng số': 8, 'Đạt baseline': 4, 'Đã xử lý trong ngày': 2,
            'Tổng đã đạt': 5, 'Chưa đạt': 3, 'OFF/Lỗi': 1, '% đạt': 62.5,
        },
    ]


# --- Parse ngày ---

def test_parse_intraday_date_ok():
    assert quality_routes._parse_shc_cts_intraday_date(
        'Bao_cao_tien_trinh_20260625.xlsx') == '2026-06-25'


def test_parse_intraday_date_invalid_returns_none():
    assert quality_routes._parse_shc_cts_intraday_date('Bao_cao_chot_ngay_20260625.xlsx') is None
    assert quality_routes._parse_shc_cts_intraday_date('Bao_cao_tien_trinh_abc.xlsx') is None


# --- Schema ---

def test_ensure_schema_creates_shc_cts_tables(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(db_path)
    cols_td = {r[1] for r in conn.execute('PRAGMA table_info(shc_cts_tien_do)')}
    cols_ks = {r[1] for r in conn.execute('PRAGMA table_info(shc_cts_kiemsoat)')}
    conn.close()
    assert {'ngay_xu_ly', 'don_vi', 'nvkt_db', 'tong_so', 'dat_baseline',
            'da_xu_ly_ngay', 'tong_dat', 'chua_dat', 'off_loi', 'ty_le_dat',
            'captured_at'} <= cols_td
    assert {'ngay_xu_ly', 'nvkt_db', 'don_vi', 'noi_dung_kiem_soat',
            'nguoi_nhap', 'thoi_diem_nhap', 'thoi_diem_cap_nhat'} <= cols_ks


# --- Sync ---

def test_sync_first_time(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    result = quality_routes._sync_shc_cts_tien_do_to_db()
    assert result['synced'] == 2
    assert result['new'] == 2
    assert result['updated'] == 0
    assert result['skipped'] == 0
    assert result['ngay_xu_ly'] == '2026-06-25'


def test_sync_upsert_keeps_latest_values(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    rows = _progress_rows()
    _write_intraday(tmp_path, monkeypatch, rows, 'Bao_cao_tien_trinh_20260625.xlsx')
    quality_routes._sync_shc_cts_tien_do_to_db()

    rows[0] = {**rows[0], 'Tổng đã đạt': 9, 'Chưa đạt': 1}
    _write_intraday(tmp_path, monkeypatch, rows, 'Bao_cao_tien_trinh_20260625.xlsx')
    result = quality_routes._sync_shc_cts_tien_do_to_db()

    assert result['synced'] == 2
    assert result['new'] == 0
    assert result['updated'] == 2
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        "SELECT tong_dat, chua_dat FROM shc_cts_tien_do WHERE nvkt_db='Nguyễn Văn A'"
    ).fetchone()
    conn.close()
    assert row[0] == 9
    assert row[1] == 1


def test_sync_skips_when_intraday_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(quality_routes, 'SHC_CTS_INTRADAY_REPORT_DIR', str(tmp_path))
    result = quality_routes._sync_shc_cts_tien_do_to_db()
    assert result['skipped'] == 1
    assert result['reason'] == 'intraday_missing'


# --- Luu kiem soat ---

def test_kiemsoat_luu_creates(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    response = _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'Nguyễn Văn A',
              'don_vi': 'Tổ Phúc Thọ', 'noi_dung': 'Đã kiểm tra'},
    )
    assert response.status_code == 200
    body = response.get_json()
    assert body['ok'] is True
    assert body['nguoi_nhap'] == 'test-user'
    assert body['noi_dung'] == 'Đã kiểm tra'


def test_kiemsoat_luu_update_keeps_thoi_diem_nhap(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO shc_cts_kiemsoat (ngay_xu_ly, nvkt_db, noi_dung_kiem_soat, "
            "nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat) "
            "VALUES ('2026-06-25','NV1','cũ','u','2020-01-01 00:00:00','2020-01-01 00:00:00')"
        )
        conn.commit()
    _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'NV1', 'noi_dung': 'mới'},
    )
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        "SELECT thoi_diem_nhap, noi_dung_kiem_soat FROM shc_cts_kiemsoat "
        "WHERE ngay_xu_ly='2026-06-25' AND nvkt_db='NV1'"
    ).fetchone()
    conn.close()
    assert row[0] == '2020-01-01 00:00:00'
    assert row[1] == 'mới'


def test_kiemsoat_luu_empty_deletes(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO shc_cts_kiemsoat (ngay_xu_ly, nvkt_db, noi_dung_kiem_soat) "
            "VALUES ('2026-06-25','NV1','cũ')"
        )
        conn.commit()
    _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'NV1', 'noi_dung': ''},
    )
    conn = sqlite3.connect(db_path)
    count = conn.execute(
        "SELECT COUNT(*) FROM shc_cts_kiemsoat "
        "WHERE ngay_xu_ly='2026-06-25' AND nvkt_db='NV1'"
    ).fetchone()[0]
    conn.close()
    assert count == 0


def test_kiemsoat_luu_rejects_missing_keys(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    response = _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'noi_dung': 'x'},
    )
    assert response.status_code == 400


# --- Detail ---

def test_detail_today_reads_live_excel_and_syncs(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail')
    assert response.status_code == 200
    body = response.get_json()
    assert body['selected_date'] == '2026-06-25'
    assert body['is_today_live'] is True
    assert '2026-06-25' in body['available_dates']
    don_vi = 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
    assert don_vi in body['sheets']
    assert body['sheets'][don_vi]['data'][0]['NVKT_DB'] == 'Nguyễn Văn A'
    conn = sqlite3.connect(db_path)
    n = conn.execute('SELECT COUNT(*) FROM shc_cts_tien_do').fetchone()[0]
    conn.close()
    assert n == 2


def test_detail_past_date_reads_from_db(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    quality_routes._sync_shc_cts_tien_do_to_db()
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260626.xlsx')
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail?date=2026-06-25')
    assert response.status_code == 200
    body = response.get_json()
    assert body['selected_date'] == '2026-06-25'
    assert body['is_today_live'] is False
    don_vi = 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
    assert body['sheets'][don_vi]['data'][0]['NVKT_DB'] == 'Nguyễn Văn A'


def test_detail_joins_kiemsoat_annotation(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'Nguyễn Văn A',
              'don_vi': 'Tổ Kỹ thuật Địa bàn Phúc Thọ', 'noi_dung': 'Ghi chú KS'},
    )
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail')
    body = response.get_json()
    don_vi = 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
    row = body['sheets'][don_vi]['data'][0]
    assert row['kiemsoat_noi_dung'] == 'Ghi chú KS'
    assert row['kiemsoat_da_nhap'] is True


def test_detail_returns_404_when_no_data(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(quality_routes, 'SHC_CTS_INTRADAY_REPORT_DIR', str(tmp_path))
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail?date=2099-01-01')
    assert response.status_code == 404
