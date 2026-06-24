import sqlite3
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import operations_routes


def _prepare_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'brcd_kiemsoat.db'
    monkeypatch.setattr(operations_routes, 'BRCD_KIEMSOAT_DB_PATH', str(db_path))
    operations_routes._brcd_kiemsoat_schema_ready_path = None
    operations_routes._ensure_brcd_kiemsoat_schema()
    return db_path


def _logged_in_client():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'
    return client


_PTTB_COLUMNS = [
    'ma_thue_bao', 'ten_thuebao', 'diachi_lapdat', 'loaihinh_tb',
    'nhanvien_tiepthi', 'doi_vt', 'ten_kv', 'ngayhen_den',
    'noidung_hen', 'chitieu_tg', 'gio_conlai', 'trang_thai',
]


def _write_fake_pttb(tmp_path, monkeypatch, rows):
    """rows: list[dict] cho 1 sheet ToKT_SonTay."""
    excel_path = tmp_path / 'baoCaoPTTB.xlsx'
    df = pd.DataFrame(rows, columns=_PTTB_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ToKT_SonTay', index=False)
    monkeypatch.setattr(operations_routes, 'PTTB_SUMMARY_FILE', str(excel_path))
    return excel_path


def _pttb_rows():
    return [
        {
            'ma_thue_bao': 'fbr00ec1h', 'ten_thuebao': 'Nguyễn Văn A',
            'diachi_lapdat': 'DC1', 'loaihinh_tb': 'Fiber',
            'nhanvien_tiepthi': 'NV1', 'doi_vt': 'ToKT_SonTay',
            'ten_kv': 'KV1', 'ngayhen_den': '2026-06-25',
            'noidung_hen': 'Hẹn lắp', 'chitieu_tg': 24,
            'gio_conlai': -50.0, 'trang_thai': 'Quá giờ',
        },
        {
            'ma_thue_bao': 'cam00ai4o', 'ten_thuebao': 'Trần Văn B',
            'diachi_lapdat': 'DC2', 'loaihinh_tb': 'MyTV',
            'nhanvien_tiepthi': 'NV2', 'doi_vt': 'ToKT_SonTay',
            'ten_kv': 'KV2', 'ngayhen_den': '2026-06-26',
            'noidung_hen': 'Hẹn sửa', 'chitieu_tg': 24,
            'gio_conlai': 5.0, 'trang_thai': 'Bình thường',
        },
    ]


# ---------------------------------------------------------------------------
# Task 1: Schema + sync + POST save
# ---------------------------------------------------------------------------


def test_ensure_schema_creates_pttb_tables(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)

    conn = sqlite3.connect(db_path)
    cols_ks = {row[1] for row in conn.execute('PRAGMA table_info(pttb_kiemsoat)')}
    cols_phieu = {row[1] for row in conn.execute('PRAGMA table_info(pttb_phieu)')}
    conn.close()

    assert {'ma_thue_bao', 'loaihinh_tb', 'doi_vt', 'nhanvien_tiepthi',
            'noi_dung_kiem_soat', 'nguoi_nhap', 'thoi_diem_nhap',
            'thoi_diem_cap_nhat'} <= cols_ks
    assert {'ma_thue_bao', 'ten_thuebao', 'diachi_lapdat', 'loaihinh_tb',
            'nhanvien_tiepthi', 'doi_vt', 'ten_kv', 'ngayhen_den',
            'noidung_hen', 'chitieu_tg', 'gio_conlai', 'trang_thai',
            'sheet', 'first_seen', 'last_seen'} <= cols_phieu


def test_sync_pttb_first_time(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    result = operations_routes._sync_pttb_phieu_to_db()

    assert result == {'synced': 2, 'new': 2, 'updated': 0, 'skipped': 0, 'reason': ''}
    conn = sqlite3.connect(db_path)
    count = conn.execute('SELECT COUNT(*) FROM pttb_phieu').fetchone()[0]
    conn.close()
    assert count == 2


def test_sync_pttb_preserves_first_seen(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    operations_routes._sync_pttb_phieu_to_db()
    conn = sqlite3.connect(db_path)
    conn.execute("UPDATE pttb_phieu SET first_seen = '2020-01-01 00:00:00'")
    conn.commit()
    conn.close()

    operations_routes._sync_pttb_phieu_to_db()
    conn = sqlite3.connect(db_path)
    rows = {r[0]: (r[1], r[2]) for r in conn.execute(
        'SELECT ma_thue_bao, first_seen, last_seen FROM pttb_phieu'
    ).fetchall()}
    conn.close()
    for _, (first, last) in rows.items():
        assert first == '2020-01-01 00:00:00'
        assert last != '2020-01-01 00:00:00'


def test_sync_pttb_skips_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'PTTB_SUMMARY_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    result = operations_routes._sync_pttb_phieu_to_db()
    assert result['skipped'] == 1
    assert result['reason'] == 'excel_missing'


def test_pttb_kiemsoat_luu_creates(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)

    response = _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={
            'ma_thue_bao': 'fbr00ec1h',
            'loaihinh_tb': 'Fiber',
            'doi_vt': 'ToKT_SonTay',
            'nhanvien_tiepthi': 'NV1',
            'noi_dung': 'KH đi vắng hẹn 18h',
        },
    )
    assert response.status_code == 200
    body = response.get_json()
    assert body == {
        'ok': True,
        'ma_thue_bao': 'fbr00ec1h',
        'noi_dung': 'KH đi vắng hẹn 18h',
        'nguoi_nhap': 'test-user',
    }


def test_pttb_kiemsoat_luu_update_keeps_thoi_diem_nhap(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(db_path)
    conn.execute(
        '''INSERT INTO pttb_kiemsoat
            (ma_thue_bao, noi_dung_kiem_soat, nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat)
            VALUES (?, ?, ?, ?, ?)''',
        ('fbr00ec1h', 'cũ', 'old', '2020-01-01 00:00:00', '2020-01-01 00:00:00'),
    )
    conn.commit()
    conn.close()

    response = _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'noi_dung': 'mới'},
    )
    assert response.status_code == 200
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        'SELECT thoi_diem_nhap, thoi_diem_cap_nhat FROM pttb_kiemsoat WHERE ma_thue_bao = ?',
        ('fbr00ec1h',),
    ).fetchone()
    conn.close()
    assert row[0] == '2020-01-01 00:00:00'
    assert row[1] != '2020-01-01 00:00:00'


def test_pttb_kiemsoat_luu_empty_deletes(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(str(tmp_path / 'brcd_kiemsoat.db'))
    conn.execute(
        'INSERT INTO pttb_kiemsoat (ma_thue_bao, noi_dung_kiem_soat) VALUES (?, ?)',
        ('fbr00ec1h', 'cũ'),
    )
    conn.commit()
    conn.close()

    response = _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'noi_dung': ''},
    )
    assert response.status_code == 200
    conn = sqlite3.connect(str(tmp_path / 'brcd_kiemsoat.db'))
    count = conn.execute(
        'SELECT COUNT(*) FROM pttb_kiemsoat WHERE ma_thue_bao = ?', ('fbr00ec1h',)
    ).fetchone()[0]
    conn.close()
    assert count == 0


# ---------------------------------------------------------------------------
# Task 2: Detail endpoint + on-load sync
# ---------------------------------------------------------------------------


def test_pttb_detail_joins_annotation(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'doi_vt': 'ToKT_SonTay',
              'nhanvien_tiepthi': 'NV1', 'noi_dung': 'Đã liên hệ'},
    )

    response = _logged_in_client().get('/api/pttb-kiemsoat/detail')
    assert response.status_code == 200
    payload = response.get_json()
    assert 'ToKT_SonTay' in payload['sheets']
    rows = payload['sheets']['ToKT_SonTay']['data']
    by_id = {r['ma_thue_bao']: r for r in rows}
    assert by_id['fbr00ec1h']['kiemsoat_da_nhap'] is True
    assert by_id['fbr00ec1h']['kiemsoat_noi_dung'] == 'Đã liên hệ'
    assert by_id['cam00ai4o']['kiemsoat_da_nhap'] is False


def test_pttb_load_triggers_sync(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM pttb_phieu').fetchone()[0] == 0
    conn.close()

    _logged_in_client().get('/api/pttb-kiemsoat/detail')
    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM pttb_phieu').fetchone()[0] == 2
    conn.close()


def test_pttb_detail_returns_404_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'PTTB_SUMMARY_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    response = _logged_in_client().get('/api/pttb-kiemsoat/detail')
    assert response.status_code == 404


# ---------------------------------------------------------------------------
# Task 3: Thongke + lich_su
# ---------------------------------------------------------------------------


def test_pttb_thongke_returns_summary(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())
    _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'doi_vt': 'ToKT_SonTay',
              'nhanvien_tiepthi': 'NV1', 'noi_dung': 'done'},
    )

    response = _logged_in_client().get('/api/pttb-kiemsoat/thongke')
    assert response.status_code == 200
    payload = response.get_json()
    assert payload['summary'] == {'total': 2, 'da_kiem_soat': 1, 'chua': 1, 'ty_le': 50.0}
    assert 'lich_su' in payload


def test_pttb_thongke_by_doi_and_by_nvkt(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    rows = _pttb_rows() + [
        {
            'ma_thue_bao': 'ext00zz99', 'ten_thuebao': 'Lê Văn C',
            'diachi_lapdat': 'DC3', 'loaihinh_tb': 'Fiber',
            'nhanvien_tiepthi': 'NV1', 'doi_vt': 'ToKT_SuoiHai',
            'ten_kv': 'KV3', 'ngayhen_den': '2026-06-27',
            'noidung_hen': 'Hẹn mới', 'chitieu_tg': 48,
            'gio_conlai': 10.0, 'trang_thai': 'Bình thường',
        },
    ]
    _write_fake_pttb(tmp_path, monkeypatch, rows)

    response = _logged_in_client().get('/api/pttb-kiemsoat/thongke')
    payload = response.get_json()

    doi_names = [d['doi_vt'] for d in payload['by_doi']]
    assert 'ToKT_SonTay' in doi_names
    assert 'ToKT_SuoiHai' in doi_names

    nvkt_names = [d['nhanvien_tiepthi'] for d in payload['by_nvkt']]
    assert 'NV1' in nvkt_names
    assert 'NV2' in nvkt_names


def test_pttb_thongke_filter_by_doi(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    rows = _pttb_rows() + [
        {
            'ma_thue_bao': 'ext00zz99', 'ten_thuebao': 'Lê Văn C',
            'diachi_lapdat': 'DC3', 'loaihinh_tb': 'Fiber',
            'nhanvien_tiepthi': 'NV3', 'doi_vt': 'ToKT_SuoiHai',
            'ten_kv': 'KV3', 'ngayhen_den': '2026-06-27',
            'noidung_hen': 'Hẹn mới', 'chitieu_tg': 48,
            'gio_conlai': 10.0, 'trang_thai': 'Bình thường',
        },
    ]
    _write_fake_pttb(tmp_path, monkeypatch, rows)

    response = _logged_in_client().get('/api/pttb-kiemsoat/thongke?doi=ToKT_SonTay')
    payload = response.get_json()
    assert payload['summary']['total'] == 2


def test_pttb_thongke_filter_nhom_qua_gio(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    response = _logged_in_client().get('/api/pttb-kiemsoat/thongke?nhom=qua_gio')
    payload = response.get_json()
    assert payload['summary']['total'] == 1


def test_pttb_thongke_lich_su_present(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())
    _logged_in_client().get('/api/pttb-kiemsoat/detail')

    conn = sqlite3.connect(str(tmp_path / 'brcd_kiemsoat.db'))
    conn.execute(
        "UPDATE pttb_phieu SET last_seen = '2020-01-01 00:00:00' WHERE ma_thue_bao = 'fbr00ec1h'"
    )
    conn.commit()
    conn.close()

    response = _logged_in_client().get('/api/pttb-kiemsoat/thongke?khoang=nam_nay')
    payload = response.get_json()
    lich_su = payload['lich_su']
    assert 'tu_ngay' in lich_su
    assert 'den_ngay' in lich_su
    assert isinstance(lich_su['roi_da_ks'], int)
    assert isinstance(lich_su['roi_chua_ks'], int)


def test_pttb_thongke_returns_404_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'PTTB_SUMMARY_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    response = _logged_in_client().get('/api/pttb-kiemsoat/thongke')
    assert response.status_code == 404


# ---------------------------------------------------------------------------
# Task 4: Excel report download
# ---------------------------------------------------------------------------


def test_pttb_report_download_returns_xlsx(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())
    _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'doi_vt': 'ToKT_SonTay',
              'nhanvien_tiepthi': 'NV1', 'noi_dung': 'done'},
    )

    response = _logged_in_client().get('/download/pttb-kiemsoat-report')
    assert response.status_code == 200
    assert 'spreadsheetml' in response.headers['Content-Type']


def test_pttb_report_download_returns_404_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'PTTB_SUMMARY_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    response = _logged_in_client().get('/download/pttb-kiemsoat-report')
    assert response.status_code == 404
