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


_DETAIL_COLUMNS = [
    'baohong_id', 'ma_tb', 'TEN_TB', 'DIACHI_LD', 'LOAIHINH_TB', 'GHICHU_HONG',
    'NVKT', 'DOI_VT', 'ngay_bh', 'Trạng thái cổng', 'ttvt_ton', 'chitieu_tg',
    'thời gian tồn thực', 'giờ còn lại thực',
]


def _write_fake_brcd_detail(tmp_path, monkeypatch, rows, sheet_name='ToKT_SonTay'):
    """rows: list[dict] cho 1 sheet ToKT_<...>."""
    excel_path = tmp_path / 'chiTietBrcd5Doi.xlsx'
    df = pd.DataFrame(rows, columns=_DETAIL_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        df.drop(columns=['baohong_id']).to_excel(writer, sheet_name=sheet_name + '_rut_gon', index=False)
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


# ---------------------------------------------------------------------------
# Task 1: schema + sync function
# ---------------------------------------------------------------------------


def test_ensure_schema_creates_brcd_phieu_table(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)

    conn = sqlite3.connect(db_path)
    cols = {row[1] for row in conn.execute('PRAGMA table_info(brcd_phieu)')}
    conn.close()

    assert {'baohong_id', 'ma_tb', 'ten_tb', 'diachi_ld', 'loaihinh_tb',
            'ghichu_hong', 'nvkt', 'doi_vt', 'ngay_bh', 'trang_thai_cong',
            'ttvt_ton', 'chitieu_tg', 'thoi_gian_ton_thuc', 'gio_con_lai_thuc',
            'sa', 'sheet', 'first_seen', 'last_seen'} <= cols


def test_ensure_schema_is_idempotent(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)

    operations_routes._ensure_brcd_kiemsoat_schema()
    operations_routes._ensure_brcd_kiemsoat_schema()


def test_sync_first_time_inserts_with_first_seen_equal_last_seen(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())

    result = operations_routes._sync_brcd_phieu_to_db()

    assert result == {'synced': 2, 'new': 2, 'updated': 0, 'skipped': 0, 'reason': ''}

    conn = sqlite3.connect(db_path)
    rows = {r[0]: r for r in conn.execute(
        'SELECT baohong_id, ma_tb, nvkt, doi_vt, sheet, first_seen, last_seen FROM brcd_phieu'
    ).fetchall()}
    conn.close()

    assert set(rows.keys()) == {11795318, 11796115}
    for baohong_id, (_, ma_tb, nvkt, doi_vt, sheet, first_seen, last_seen) in rows.items():
        assert first_seen == last_seen  # lần đầu: first_seen = last_seen
        assert first_seen is not None
        assert sheet == 'ToKT_SonTay'


def test_sync_second_time_preserves_first_seen_updates_last_seen(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())

    operations_routes._sync_brcd_phieu_to_db()

    # Fake time: set first_seen thủ công về quá khứ
    conn = sqlite3.connect(db_path)
    conn.execute("UPDATE brcd_phieu SET first_seen = '2020-01-01 00:00:00'")
    conn.commit()
    conn.close()

    # Sync lại
    result = operations_routes._sync_brcd_phieu_to_db()
    assert result == {'synced': 2, 'new': 0, 'updated': 2, 'skipped': 0, 'reason': ''}

    conn = sqlite3.connect(db_path)
    rows = {r[0]: (r[1], r[2]) for r in conn.execute(
        'SELECT baohong_id, first_seen, last_seen FROM brcd_phieu'
    ).fetchall()}
    conn.close()

    for baohong_id, (first_seen, last_seen) in rows.items():
        assert first_seen == '2020-01-01 00:00:00'  # giữ nguyên
        assert last_seen != '2020-01-01 00:00:00'   # cập nhật mới
        assert last_seen.startswith('20')  # ISO timestamp


def test_sync_skips_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'BRCD_DETAIL_MAIN_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    result = operations_routes._sync_brcd_phieu_to_db()

    assert result == {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'excel_missing'}


def test_sync_keeps_phieu_that_left_universe(tmp_path, monkeypatch):
    """Phiếu biến mất khỏi Excel → dòng DB vẫn giữ, không bị xóa."""
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())  # 2 phiếu

    operations_routes._sync_brcd_phieu_to_db()

    # Lần 2: Excel chỉ còn 1 phiếu (phiếu 11796115 đã xử lý xong)
    _write_fake_brcd_detail(tmp_path, monkeypatch, [_detail_rows()[0]])

    result = operations_routes._sync_brcd_phieu_to_db()
    assert result == {'synced': 1, 'new': 0, 'updated': 1, 'skipped': 0, 'reason': ''}

    conn = sqlite3.connect(db_path)
    count = conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0]
    ids = {r[0] for r in conn.execute('SELECT baohong_id FROM brcd_phieu').fetchall()}
    conn.close()

    assert count == 2  # vẫn 2 dòng, không xóa
    assert ids == {11795318, 11796115}  # phiếu đã rời tồn vẫn còn trong DB


def test_sync_updates_field_changes(tmp_path, monkeypatch):
    """Khi Excel update NVKT của 1 phiếu, snapshot cũng update."""
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())
    operations_routes._sync_brcd_phieu_to_db()

    # Excel giờ đổi NVKT của phiếu 11795318 từ NV1 → NV_NEW
    new_rows = _detail_rows()
    new_rows[0] = {**new_rows[0], 'NVKT': 'NV_NEW'}
    _write_fake_brcd_detail(tmp_path, monkeypatch, new_rows)

    operations_routes._sync_brcd_phieu_to_db()

    conn = sqlite3.connect(db_path)
    nvkt = conn.execute(
        'SELECT nvkt FROM brcd_phieu WHERE baohong_id = 11795318'
    ).fetchone()[0]
    conn.close()
    assert nvkt == 'NV_NEW'
