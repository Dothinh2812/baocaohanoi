import os
import sqlite3
import sys
from datetime import date
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


# ---------------------------------------------------------------------------
# Task 2: on-load sync
# ---------------------------------------------------------------------------


def test_load_brcd_kiemsoat_df_triggers_sync(tmp_path, monkeypatch):
    """Khi detail endpoint load, brcd_phieu tự sync từ Excel."""
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())

    # Trước khi gọi endpoint: brcd_phieu rỗng
    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0] == 0
    conn.close()

    response = _logged_in_client().get('/api/brcd-kiemsoat/detail')

    assert response.status_code == 200
    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0] == 2
    conn.close()


def test_load_brcd_kiemsoat_df_does_not_crash_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'BRCD_DETAIL_MAIN_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    # Detail endpoint trả 404 (Excel thiếu) — không crash
    response = _logged_in_client().get('/api/brcd-kiemsoat/detail')
    assert response.status_code == 404


# ---------------------------------------------------------------------------
# Task 3: cron script
# ---------------------------------------------------------------------------


def test_sync_script_runs_via_subprocess(tmp_path, monkeypatch):
    """Script scripts/sync_brcd_phieu.py chạy được với env DASHV4_*."""
    import subprocess

    db_path = tmp_path / 'brcd_kiemsoat.db'
    excel_path = tmp_path / 'chiTietBrcd5Doi.xlsx'

    df = pd.DataFrame(_detail_rows(), columns=_DETAIL_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ToKT_SonTay', index=False)

    env = {
        **os.environ,
        'DASHV4_BRCD_KIEMSOAT_DB_PATH': str(db_path),
    }
    wrapper = tmp_path / 'run_sync.py'
    wrapper.write_text(f'''
import sys
sys.path.insert(0, {repr(str(Path(__file__).resolve().parents[1]))})
import runtime_limits
from blueprints import operations_routes
operations_routes.BRCD_DETAIL_MAIN_FILE = {repr(str(excel_path))}
operations_routes.BRCD_KIEMSOAT_DB_PATH = {repr(str(db_path))}
operations_routes._brcd_kiemsoat_schema_ready_path = None
result = operations_routes._sync_brcd_phieu_to_db()
print(result)
''')

    completed = subprocess.run(
        ['python3', str(wrapper)],
        env=env,
        capture_output=True,
        text=True,
        timeout=30,
    )

    assert completed.returncode == 0, completed.stderr
    assert "'synced': 2" in completed.stdout

    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0] == 2
    conn.close()


# ---------------------------------------------------------------------------
# Task 4: thongke lich_su
# ---------------------------------------------------------------------------


def _seed_phieu_with_history(tmp_path, monkeypatch, db_path):
    """Tạo kịch bản:
    - Phiếu 11795318: rời tồn trong tháng này, ĐÃ kiểm soát
    - Phiếu 11796115: rời tồn trong tháng này, CHƯA kiểm soát
    - Phiếu 11797000: vẫn còn trong tồn (current universe)
    """
    rows = _detail_rows() + [{
        'baohong_id': 11797000, 'ma_tb': 'tb3', 'TEN_TB': 'NV C', 'DIACHI_LD': 'DC3',
        'LOAIHINH_TB': 'Fiber', 'GHICHU_HONG': 'h3', 'NVKT': 'NV3', 'DOI_VT': 'ToKT_SonTay',
        'ngay_bh': '2026-06-24 08:00', 'Trạng thái cổng': 'ON', 'ttvt_ton': 'x',
        'chitieu_tg': 8, 'thời gian tồn thực': 1.0, 'giờ còn lại thực': 7.0,
    }]
    _write_fake_brcd_detail(tmp_path, monkeypatch, rows)

    # Trigger sync (tất cả 3 phiếu có first_seen=last_seen=now)
    operations_routes._sync_brcd_phieu_to_db()

    # Fake last_seen của 2 phiếu "rời tồn" về quá khứ (trong tháng này)
    today = date.today()
    past_iso = today.replace(day=max(1, today.day - 3)).strftime('%Y-%m-%d 10:00:00')
    conn = sqlite3.connect(db_path)
    conn.execute(
        "UPDATE brcd_phieu SET last_seen = ? WHERE baohong_id IN (?, ?)",
        (past_iso, 11795318, 11796115),
    )
    conn.commit()
    conn.close()

    # Excel "hiện tại" giờ chỉ còn phiếu 11797000 — mô phỏng 2 phiếu kia đã xử lý xong
    _write_fake_brcd_detail(tmp_path, monkeypatch, [rows[2]])

    # Trigger sync lại để cập nhật last_seen của 11797000 = now
    operations_routes._sync_brcd_phieu_to_db()

    # Add kiemsoat cho phiếu 11795318 (đã kiểm soát trước khi rời)
    _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'doi_vt': 'ToKT_SonTay', 'nvkt': 'NV1', 'noi_dung': 'done'},
    )
    # Excel cần chứa 11795318 để POST luu ghi annotation được — viết lại Excel với cả 3
    _write_fake_brcd_detail(tmp_path, monkeypatch, rows)
    _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'doi_vt': 'ToKT_SonTay', 'nvkt': 'NV1', 'noi_dung': 'done'},
    )

    # Cuối cùng: Excel chỉ còn 11797000 (mô phỏng real state)
    _write_fake_brcd_detail(tmp_path, monkeypatch, [rows[2]])


def test_thongke_returns_lich_su_default_thang_nay(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _seed_phieu_with_history(tmp_path, monkeypatch, db_path)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke')

    assert response.status_code == 200
    payload = response.get_json()
    assert 'lich_su' in payload
    ls = payload['lich_su']
    assert ls['roi_da_ks'] == 1   # 11795318 đã KS rồi rời
    assert ls['roi_chua_ks'] == 1  # 11796115 rời mà chưa KS


def test_thongke_lich_su_filter_khoang_tat_ca(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _seed_phieu_with_history(tmp_path, monkeypatch, db_path)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke?khoang=tat_ca')

    payload = response.get_json()
    ls = payload['lich_su']
    # tat_ca không giới hạn date → 2 phiếu rời tồn vẫn đếm
    assert ls['roi_da_ks'] + ls['roi_chua_ks'] == 2


def test_thongke_lich_su_filter_doi_applies(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _seed_phieu_with_history(tmp_path, monkeypatch, db_path)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke?doi=ToKT_SonTay')

    payload = response.get_json()
    ls = payload['lich_su']
    assert ls['roi_da_ks'] + ls['roi_chua_ks'] == 2

    # Đội khác → 0
    response2 = _logged_in_client().get('/api/brcd-kiemsoat/thongke?doi=ToKT_PhucTho')
    ls2 = response2.get_json()['lich_su']
    assert ls2['roi_da_ks'] + ls2['roi_chua_ks'] == 0


# ---------------------------------------------------------------------------
# Task 5: UI
# ---------------------------------------------------------------------------


def test_brcd_page_has_khoang_dropdown_and_lich_su_cards():
    response = _logged_in_client().get('/brcd')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    # Dropdown khoảng thời gian
    assert 'id="kiemsoat-filter-khoang"' in html
    assert 'value="thang_nay"' in html
    assert 'value="tuan_nay"' in html
    assert 'value="nam_nay"' in html
    assert 'value="tat_ca"' in html

    # JS render 2 thẻ mới
    brcd_js = Path(__file__).resolve().parents[1].joinpath(
        'static', 'js', 'pages', 'brcd.js').read_text('utf-8')
    assert 'roi_da_ks' in brcd_js
    assert 'roi_chua_ks' in brcd_js
    assert 'Rời tồn' in brcd_js
    # _kiemSoatQuery gửi param khoang
    assert "'khoang'" in brcd_js or '"khoang"' in brcd_js
