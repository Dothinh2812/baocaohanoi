import io
import sqlite3
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import operations_routes


def _prepare_kiemsoat_db(tmp_path, monkeypatch):
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


# ---------------------------------------------------------------------------
# Task 1: storage layer
# ---------------------------------------------------------------------------


def test_get_brcd_kiemsoat_map_returns_empty_for_empty_list(tmp_path, monkeypatch):
    _prepare_kiemsoat_db(tmp_path, monkeypatch)

    assert operations_routes.get_brcd_kiemsoat_map([]) == {}


def test_get_brcd_kiemsoat_map_returns_inserted_rows(tmp_path, monkeypatch):
    db_path = _prepare_kiemsoat_db(tmp_path, monkeypatch)
    operations_routes._ensure_brcd_kiemsoat_schema()

    conn = sqlite3.connect(db_path)
    conn.execute(
        '''
        INSERT INTO brcd_kiemsoat
            (baohong_id, ma_tb, doi_vt, nvkt, noi_dung_kiem_soat, nguoi_nhap,
             thoi_diem_nhap, thoi_diem_cap_nhat)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''',
        (11795318, 'vanhung98879', 'ToKT_PhucTho', 'Khuất Duy Hiệp',
         'KH đi vắng', 'to_truong', '2026-06-24 08:00:00', '2026-06-24 09:00:00'),
    )
    conn.commit()
    conn.close()

    result = operations_routes.get_brcd_kiemsoat_map([11795318, 99999999])

    assert set(result.keys()) == {11795318}
    row = result[11795318]
    assert row['ma_tb'] == 'vanhung98879'
    assert row['noi_dung_kiem_soat'] == 'KH đi vắng'
    assert row['nguoi_nhap'] == 'to_truong'


def test_ensure_brcd_kiemsoat_schema_is_idempotent(tmp_path, monkeypatch):
    _prepare_kiemsoat_db(tmp_path, monkeypatch)

    operations_routes._ensure_brcd_kiemsoat_schema()
    operations_routes._ensure_brcd_kiemsoat_schema()

    conn = sqlite3.connect(operations_routes.BRCD_KIEMSOAT_DB_PATH)
    cols = {row[1] for row in conn.execute('PRAGMA table_info(brcd_kiemsoat)')}
    conn.close()
    assert {'baohong_id', 'ma_tb', 'doi_vt', 'nvkt', 'noi_dung_kiem_soat',
            'nguoi_nhap', 'thoi_diem_nhap', 'thoi_diem_cap_nhat'} <= cols


# ---------------------------------------------------------------------------
# Task 2: POST save
# ---------------------------------------------------------------------------


def test_kiemsoat_luu_creates_new_note(tmp_path, monkeypatch):
    db_path = _prepare_kiemsoat_db(tmp_path, monkeypatch)

    response = _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={
            'baohong_id': 11795318,
            'ma_tb': 'vanhung98879',
            'doi_vt': 'ToKT_PhucTho',
            'nvkt': 'Khuất Duy Hiệp',
            'noi_dung': 'KH đi vắng hẹn 18h',
        },
    )

    assert response.status_code == 200
    body = response.get_json()
    assert body == {
        'ok': True,
        'baohong_id': 11795318,
        'noi_dung': 'KH đi vắng hẹn 18h',
        'nguoi_nhap': 'test-user',
    }

    conn = sqlite3.connect(db_path)
    row = conn.execute(
        'SELECT noi_dung_kiem_soat, nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat '
        'FROM brcd_kiemsoat WHERE baohong_id = 11795318'
    ).fetchone()
    conn.close()
    assert row[0] == 'KH đi vắng hẹn 18h'
    assert row[1] == 'test-user'
    assert row[2] == row[3]  # tạo mới: thoi_diem_nhap == cap_nhat


def test_kiemsoat_luu_update_keeps_thoi_diem_nhap(tmp_path, monkeypatch):
    db_path = _prepare_kiemsoat_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(db_path)
    conn.execute(
        '''
        INSERT INTO brcd_kiemsoat
            (baohong_id, ma_tb, doi_vt, nvkt, noi_dung_kiem_soat, nguoi_nhap,
             thoi_diem_nhap, thoi_diem_cap_nhat)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''',
        (11795318, 'x', 'ToKT_PhucTho', 'NV1', 'cũ', 'old_user',
         '2026-06-20 08:00:00', '2026-06-20 08:00:00'),
    )
    conn.commit()
    conn.close()

    response = _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'noi_dung': 'mới'},
    )

    assert response.status_code == 200
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        'SELECT thoi_diem_nhap, thoi_diem_cap_nhat FROM brcd_kiemsoat WHERE baohong_id = 11795318'
    ).fetchone()
    conn.close()
    assert row[0] == '2026-06-20 08:00:00'  # giữ nguyên
    assert row[1] != '2026-06-20 08:00:00'  # cập nhật mới


def test_kiemsoat_luu_empty_note_deletes_row(tmp_path, monkeypatch):
    db_path = _prepare_kiemsoat_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(db_path)
    conn.execute(
        'INSERT INTO brcd_kiemsoat (baohong_id, noi_dung_kiem_soat, thoi_diem_nhap) VALUES (?, ?, ?)',
        (11795318, 'cũ', '2026-06-20 08:00:00'),
    )
    conn.commit()
    conn.close()

    response = _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'noi_dung': ''},
    )

    assert response.status_code == 200
    conn = sqlite3.connect(db_path)
    count = conn.execute(
        'SELECT COUNT(*) FROM brcd_kiemsoat WHERE baohong_id = 11795318'
    ).fetchone()[0]
    conn.close()
    assert count == 0


def test_kiemsoat_luu_rejects_invalid_baohong_id(tmp_path, monkeypatch):
    _prepare_kiemsoat_db(tmp_path, monkeypatch)

    response = _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 'khong-phai-so', 'noi_dung': 'x'},
    )

    assert response.status_code == 400
    assert response.get_json()['ok'] is False


def test_kiemsoat_luu_rejects_too_long_note(tmp_path, monkeypatch):
    _prepare_kiemsoat_db(tmp_path, monkeypatch)

    response = _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 1, 'noi_dung': 'x' * 2001},
    )

    assert response.status_code == 400


# ---------------------------------------------------------------------------
# Task 3: GET detail (join)
# ---------------------------------------------------------------------------


_DETAIL_COLUMNS = [
    'baohong_id', 'ma_tb', 'TEN_TB', 'DIACHI_LD', 'LOAIHINH_TB', 'GHICHU_HONG',
    'NVKT', 'DOI_VT', 'ngay_bh', 'Trạng thái cổng', 'ttvt_ton', 'chitieu_tg',
    'thời gian tồn thực', 'giờ còn lại thực',
]


def _write_fake_brcd_detail(tmp_path, monkeypatch, rows):
    """rows: list[dict] cho 1 sheet ToKT_SonTay (đủ _DETAIL_COLUMNS)."""
    excel_path = tmp_path / 'chiTietBrcd5Doi.xlsx'
    df = pd.DataFrame(rows, columns=_DETAIL_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ToKT_SonTay', index=False)
        # thêm 1 sheet _rut_gon để xác nhận loader bỏ qua đúng
        df.drop(columns=['baohong_id']).to_excel(writer, sheet_name='ToKT_SonTay_rut_gon', index=False)
    monkeypatch.setattr(operations_routes, 'BRCD_DETAIL_MAIN_FILE', str(excel_path))
    return excel_path


def _detail_rows():
    return [
        {
            'baohong_id': 11795318, 'ma_tb': 'tb1', 'TEN_TB': 'NV A', 'DIACHI_LD': 'DC1',
            'LOAIHINH_TB': 'Fiber', 'GHICHU_HONG': 'hong1', 'NVKT': 'NV1', 'DOI_VT': 'ToKT_SonTay',
            'ngay_bh': '2026-06-23 10:59', 'Trạng thái cổng': 'ON', 'ttvt_ton': 'KH đi vắng',
            'chitieu_tg': 8, 'thời gian tồn thực': 21.6, 'giờ còn lại thực': -13.6,
        },
        {
            'baohong_id': 11796115, 'ma_tb': 'tb2', 'TEN_TB': 'NV B', 'DIACHI_LD': 'DC2',
            'LOAIHINH_TB': 'MyTV', 'GHICHU_HONG': 'hong2', 'NVKT': 'NV2', 'DOI_VT': 'ToKT_SonTay',
            'ngay_bh': '2026-06-23 17:41', 'Trạng thái cổng': 'OFF', 'ttvt_ton': 'Đứt cáp',
            'chitieu_tg': 8, 'thời gian tồn thực': 4.0, 'giờ còn lại thực': 4.0,
        },
    ]


def test_kiemsoat_detail_joins_annotation(tmp_path, monkeypatch):
    _prepare_kiemsoat_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())

    # chỉ phiếu 11795318 đã được kiểm soát
    _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'doi_vt': 'ToKT_SonTay', 'nvkt': 'NV1', 'noi_dung': 'Đã liên hệ KH'},
    )

    response = _logged_in_client().get('/api/brcd-kiemsoat/detail')

    assert response.status_code == 200
    payload = response.get_json()
    assert 'ToKT_SonTay' in payload['sheets']
    rows = payload['sheets']['ToKT_SonTay']['data']
    cols = payload['sheets']['ToKT_SonTay']['columns']
    assert 'kiemsoat_noi_dung' in cols
    assert 'kiemsoat_da_nhap' in cols

    by_id = {r['baohong_id']: r for r in rows}
    assert by_id[11795318]['kiemsoat_da_nhap'] is True
    assert by_id[11795318]['kiemsoat_noi_dung'] == 'Đã liên hệ KH'
    assert by_id[11795318]['kiemsoat_nguoi_nhap'] == 'test-user'
    assert by_id[11796115]['kiemsoat_da_nhap'] is False
    assert by_id[11796115]['kiemsoat_noi_dung'] == ''


def test_kiemsoat_detail_404_when_excel_missing(tmp_path, monkeypatch):
    _prepare_kiemsoat_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'BRCD_DETAIL_MAIN_FILE', str(tmp_path / 'khong_co.xlsx'))

    response = _logged_in_client().get('/api/brcd-kiemsoat/detail')

    assert response.status_code == 404


# ---------------------------------------------------------------------------
# Task 4: GET thống kê (filter + aggregate)
# ---------------------------------------------------------------------------


def _prepare_thongke(tmp_path, monkeypatch):
    _prepare_kiemsoat_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())
    # phiếu 11795318 (quá giờ) đã kiểm soát; 11796115 (trong giờ) chưa
    _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'doi_vt': 'ToKT_SonTay', 'nvkt': 'NV1', 'noi_dung': 'done'},
    )


def test_kiemsoat_thongke_default_summary(tmp_path, monkeypatch):
    _prepare_thongke(tmp_path, monkeypatch)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke')

    assert response.status_code == 200
    payload = response.get_json()
    assert payload['summary'] == {'total': 2, 'da_kiem_soat': 1, 'chua': 1, 'ty_le': 50.0}
    assert payload['by_doi'] == [
        {'DOI_VT': 'ToKT_SonTay', 'total': 2, 'da_kiem_soat': 1, 'chua': 1, 'ty_le': 50.0}
    ]
    assert len(payload['chi_tiet']) == 2


def test_kiemsoat_thongke_filter_nhom_qua_gio(tmp_path, monkeypatch):
    _prepare_thongke(tmp_path, monkeypatch)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke?nhom=qua_gio')

    payload = response.get_json()
    assert payload['summary']['total'] == 1
    assert payload['summary']['da_kiem_soat'] == 1  # phiếu quá giờ là phiếu đã kiểm soát
    assert [r['baohong_id'] for r in payload['chi_tiet']] == [11795318]


def test_kiemsoat_thongke_filter_trangthai_chua(tmp_path, monkeypatch):
    _prepare_thongke(tmp_path, monkeypatch)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke?trangthai=chua')

    payload = response.get_json()
    assert payload['summary']['total'] == 1
    assert payload['summary']['da_kiem_soat'] == 0
    assert [r['baohong_id'] for r in payload['chi_tiet']] == [11796115]


def test_kiemsoat_thongke_filter_doi_uses_raw_doi_vt(tmp_path, monkeypatch):
    """Bộ lọc Đội phải khớp giá trị DOI_VT gốc (ToKT_...), không phải tên hiển thị."""
    _prepare_thongke(tmp_path, monkeypatch)

    match = _logged_in_client().get('/api/brcd-kiemsoat/thongke?doi=ToKT_SonTay')
    nomatch = _logged_in_client().get('/api/brcd-kiemsoat/thongke?doi=ToKT_PhucTho')
    display_name = _logged_in_client().get('/api/brcd-kiemsoat/thongke?doi=S%C3%B3n%20T%C3%A2y')

    assert match.get_json()['summary']['total'] == 2          # đúng mã đội → có dữ liệu
    assert nomatch.get_json()['summary']['total'] == 0        # mã đội không tồn tại → rỗng
    assert display_name.get_json()['summary']['total'] == 0   # tên hiển thị không khớp cột


# ---------------------------------------------------------------------------
# Task 5: Excel report download
# ---------------------------------------------------------------------------


def test_kiemsoat_report_download_returns_xlsx(tmp_path, monkeypatch):
    _prepare_thongke(tmp_path, monkeypatch)

    response = _logged_in_client().get('/download/brcd-kiemsoat-report')

    assert response.status_code == 200
    assert 'spreadsheetml' in response.headers['Content-Type']
    assert response.data[:2] == b'PK'  # ZIP magic (xlsx)
    cd = response.headers['Content-Disposition']
    assert 'brcd_kiemsoat_' in cd and '.xlsx' in cd

    df = pd.read_excel(io.BytesIO(response.data), sheet_name='Kiem_soat_BRCD')
    assert len(df) == 2
    assert 'kiemsoat_noi_dung' in df.columns


def test_kiemsoat_report_download_applies_filter(tmp_path, monkeypatch):
    _prepare_thongke(tmp_path, monkeypatch)

    response = _logged_in_client().get('/download/brcd-kiemsoat-report?trangthai=chua')

    assert response.status_code == 200
    df = pd.read_excel(io.BytesIO(response.data), sheet_name='Kiem_soat_BRCD')
    assert len(df) == 1
    assert int(df.iloc[0]['baohong_id']) == 11796115


# ---------------------------------------------------------------------------
# Task 6: Frontend (HTML assertions)
# ---------------------------------------------------------------------------


def test_brcd_page_has_kiemsoat_section():
    response = _logged_in_client().get('/brcd')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    # bảng chi tiết đầu tiên giờ là nơi nhập kiểm soát
    assert 'id="excel-tables-container-main"' in html
    assert 'Nội dung kiểm soát' in html  # mô tả dưới tiêu đề bảng đầu
    # section thống kê
    assert 'id="kiemsoat-stats"' in html
    assert 'id="kiemsoat-chitiet-container"' in html
    assert 'id="kiemsoat-filter-nhom"' in html
    assert 'Thống kê kiểm soát' in html
    # handlers wires trong template
    assert 'exportBrcdKiemSoat' in html
    assert 'reloadKiemSoatThongKe' in html
    # JS render cột kiểm soát + endpoint
    brcd_js = Path(__file__).resolve().parents[1].joinpath('static', 'js', 'pages', 'brcd.js').read_text('utf-8')
    assert 'getBrcdKiemSoatDetail' in brcd_js
    assert 'saveKiemSoat' in brcd_js
    assert '/download/brcd-kiemsoat-report' in brcd_js




