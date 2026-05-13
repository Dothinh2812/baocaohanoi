import copy
import json
import os
import sqlite3
import tempfile
import time
from datetime import date, datetime, timedelta
from threading import Event, Lock, Thread

import pandas as pd

from app_helpers import read_excel_sheet_cached
from config import (
    BASE_DATA_PATH,
    QUANG_CHU_DONG_CACHE_FILE,
    QUANG_CHU_DONG_CACHE_LOCK_FILE,
    QUANG_CHU_DONG_CACHE_REFRESH_SECONDS,
    QUANG_CHU_DONG_CACHE_WAIT_SECONDS,
    QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE,
    QUANG_CHU_DONG_DB_LIST,
    QUANG_CHU_DONG_DB_PATH,
)

try:
    import fcntl
except ImportError:  # pragma: no cover
    fcntl = None


def _extract_sa(port_id):
    if port_id and '_' in port_id:
        parts = port_id.split('_')
        if len(parts) > 1:
            return parts[1].split(':')[0] if ':' in parts[1] else parts[1]
    return ''


def _fetch_rows(database_path, query, params=(), *, row_factory=False):
    last_error = None
    for attempt in range(3):
        conn = None
        try:
            conn = sqlite3.connect(database_path, timeout=5)
            if row_factory:
                conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, params)
            return cursor.fetchall()
        except sqlite3.OperationalError as exc:
            last_error = exc
            message = str(exc).lower()
            if attempt == 2 or ('unable to open database file' not in message and 'database is locked' not in message):
                raise
            time.sleep(0.2 * (attempt + 1))
        finally:
            if conn is not None:
                conn.close()

    if last_error is not None:
        raise last_error


def _normalize_value(value):
    if isinstance(value, dict):
        return {key: _normalize_value(val) for key, val in value.items()}
    if isinstance(value, (list, tuple)):
        return [_normalize_value(item) for item in value]
    if isinstance(value, (datetime, date)):
        return value.strftime('%Y-%m-%d %H:%M:%S') if isinstance(value, datetime) else value.isoformat()
    if hasattr(value, 'item') and not isinstance(value, (str, bytes)):
        try:
            return _normalize_value(value.item())
        except Exception:
            pass
    try:
        if pd.isna(value):
            return ''
    except Exception:
        pass
    return value


def _sort_duration_key(item):
    duration = item.get('duration_minutes')
    try:
        if duration not in (None, ''):
            return (0, float(duration), str(item.get('ma_tb', '')))
    except (TypeError, ValueError):
        pass

    event_time = item.get('first_off_time') or item.get('alert_time') or ''
    normalized = str(event_time).replace('T', ' ')
    for time_format in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S'):
        try:
            parsed = datetime.strptime(normalized, time_format)
            return (1, parsed.timestamp(), str(item.get('ma_tb', '')))
        except ValueError:
            continue

    return (2, str(event_time), str(item.get('ma_tb', '')))


class QuangChuDongCache:
    def __init__(self):
        self._app = None
        self._state_lock = Lock()
        self._refresh_lock = Lock()
        self._stop_event = Event()
        self._thread = None
        self._pid = None
        self._snapshot = None
        self._snapshot_mtime = 0
        self._source_snapshot_mtime = 0
        self._last_error = None
        self._ngay_bh_cache = None
        self._ngay_bh_cache_mtime = 0
        self._last_source_check_time = 0
        self._last_disk_check_time = 0

    def initialize(self, app, *, warm=False):
        current_pid = os.getpid()
        with self._state_lock:
            self._app = app
            if self._pid != current_pid:
                self._pid = current_pid
                self._stop_event = Event()
                self._thread = None
                self._snapshot = None
                self._snapshot_mtime = 0
                self._source_snapshot_mtime = 0
                self._last_error = None

        if warm:
            self.ensure_snapshot_loaded(wait_for_refresh=True)

        with self._state_lock:
            if self._thread and self._thread.is_alive():
                return

            self._thread = Thread(
                target=self._refresh_loop,
                name='quangchudong-cache-refresh',
                daemon=True,
            )
            self._thread.start()

    def get_active_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        payload = copy.deepcopy(dict(snapshot.get('active', {'alerts': [], 'excluded': [], 'sources': []})))
        payload['cache'] = copy.deepcopy(dict(snapshot.get('cache_meta', self._fallback_cache_meta())))
        return payload

    def get_recovered_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        return copy.deepcopy(snapshot.get('recovered', []))

    def get_wide_area_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        return copy.deepcopy(snapshot.get('wide_area', []))

    def get_wide_area_groups_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        payload = snapshot.get('wide_area_groups')
        if payload is not None:
            return copy.deepcopy(payload)
        return self._fallback_wide_area_groups(snapshot)

    def get_exclusion_list_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        return copy.deepcopy(snapshot.get('exclusion_list', []))

    def get_pattern_exclusions_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        payload = snapshot.get('pattern_exclusions')
        if payload is not None:
            return copy.deepcopy(payload)
        return self._fallback_pattern_exclusions(snapshot)

    def get_port_down_groups_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        payload = snapshot.get('port_down_groups')
        if payload is not None:
            return copy.deepcopy(payload)
        return []

    def get_stats_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        return copy.deepcopy(snapshot.get('stats', {
            'active_by_doi_vt': {},
            'recovered_24h_by_doi_vt': {},
        }))

    def get_dashboard_payload(self):
        snapshot = self._get_snapshot(wait_for_refresh_if_empty=True)
        active = copy.deepcopy(dict(snapshot.get('active', {'alerts': [], 'excluded': [], 'sources': []})))
        active['cache'] = copy.deepcopy(dict(snapshot.get('cache_meta', self._fallback_cache_meta())))
        return {
            'active': active,
            'wide_area_groups': copy.deepcopy(
                snapshot.get('wide_area_groups')
                if snapshot.get('wide_area_groups') is not None
                else self._fallback_wide_area_groups(snapshot)
            ),
            'pattern_exclusions': copy.deepcopy(
                snapshot.get('pattern_exclusions')
                if snapshot.get('pattern_exclusions') is not None
                else self._fallback_pattern_exclusions(snapshot)
            ),
            'port_down_groups': copy.deepcopy(snapshot.get('port_down_groups') or []),
        }

    def ensure_snapshot_loaded(self, *, wait_for_refresh=False):
        if self._load_source_snapshot_if_newer():
            return True

        if self._load_snapshot_from_disk_if_newer():
            return True

        if self._refresh_shared_snapshot():
            return True

        if wait_for_refresh:
            deadline = time.time() + QUANG_CHU_DONG_CACHE_WAIT_SECONDS
            while time.time() < deadline:
                if self._load_snapshot_from_disk_if_newer():
                    return True
                time.sleep(0.2)

        with self._state_lock:
            return self._snapshot is not None

    def _get_snapshot(self, *, wait_for_refresh_if_empty=False):
        now = time.time()
        with self._state_lock:
            source_check_age = now - self._last_source_check_time
            disk_check_age = now - self._last_disk_check_time

        if source_check_age >= 5:
            self._load_source_snapshot_if_newer()
            with self._state_lock:
                self._last_source_check_time = now
        if disk_check_age >= 5:
            self._load_snapshot_from_disk_if_newer()
            with self._state_lock:
                self._last_disk_check_time = now

        with self._state_lock:
            snapshot = self._snapshot

        if snapshot is not None:
            return snapshot

        self.ensure_snapshot_loaded(wait_for_refresh=wait_for_refresh_if_empty)
        with self._state_lock:
            return self._snapshot or {
                'active': {'alerts': [], 'excluded': [], 'sources': []},
                'recovered': [],
                'wide_area': [],
                'wide_area_groups': [],
                'exclusion_list': [],
                'pattern_exclusions': [],
                'port_down_groups': [],
                'stats': {'active_by_doi_vt': {}, 'recovered_24h_by_doi_vt': {}},
                'cache_meta': self._fallback_cache_meta(),
            }

    def _refresh_loop(self):
        while not self._stop_event.wait(QUANG_CHU_DONG_CACHE_REFRESH_SECONDS):
            try:
                self._load_source_snapshot_if_newer()
                self._load_snapshot_from_disk_if_newer()
                self._refresh_shared_snapshot()
                self._load_source_snapshot_if_newer()
                self._load_snapshot_from_disk_if_newer()
            except Exception as exc:  # pragma: no cover
                self._log_warning('Khong the refresh cache quang chu dong nen', exc)

    def _refresh_shared_snapshot(self):
        if QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE and os.path.exists(QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE):
            return self._load_source_snapshot_if_newer()

        if not self._refresh_lock.acquire(blocking=False):
            return False

        try:
            os.makedirs(os.path.dirname(QUANG_CHU_DONG_CACHE_FILE), exist_ok=True)
            with open(QUANG_CHU_DONG_CACHE_LOCK_FILE, 'a+', encoding='utf-8') as lock_file:
                if not self._acquire_file_lock(lock_file):
                    return False

                try:
                    if self._snapshot_file_is_fresh():
                        return self._load_snapshot_from_disk_if_newer()

                    started = time.perf_counter()
                    snapshot = self._build_snapshot()
                    partial_error = self._summarize_source_errors(snapshot)
                    snapshot['cache_meta'] = {
                        'status': 'degraded' if partial_error else 'ready',
                        'refreshed_at': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                        'refresh_interval_seconds': QUANG_CHU_DONG_CACHE_REFRESH_SECONDS,
                        'duration_ms': int((time.perf_counter() - started) * 1000),
                        'last_error': partial_error,
                        'pid': os.getpid(),
                    }
                    self._write_snapshot(snapshot)
                    self._set_snapshot(snapshot, os.path.getmtime(QUANG_CHU_DONG_CACHE_FILE))
                    self._last_error = None
                    return True
                except Exception as exc:
                    self._last_error = str(exc)
                    self._log_warning('Khong the tao snapshot quang chu dong', exc)
                    return self._load_snapshot_from_disk_if_newer()
                finally:
                    self._release_file_lock(lock_file)
        finally:
            self._refresh_lock.release()

    def _build_snapshot(self):
        return {
            'active': self._build_active_payload(),
            'recovered': self._build_recovered_payload(),
            'wide_area': self._build_wide_area_payload(),
            'wide_area_groups': [],
            'exclusion_list': self._build_exclusion_payload(),
            'pattern_exclusions': [],
            'port_down_groups': [],
            'stats': self._build_stats_payload(),
        }

    def _load_source_snapshot_if_newer(self):
        if not QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE or not os.path.exists(QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE):
            return False

        file_mtime = os.path.getmtime(QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE)
        with self._state_lock:
            if self._snapshot is not None and file_mtime <= self._source_snapshot_mtime:
                return True

        try:
            with open(QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE, 'r', encoding='utf-8') as snapshot_file:
                source_snapshot = json.load(snapshot_file)
            snapshot = self._build_snapshot_from_source(source_snapshot)
            self._set_snapshot(snapshot, file_mtime, source_snapshot_mtime=file_mtime)
            self._last_error = None
            return True
        except Exception as exc:
            self._last_error = str(exc)
            self._log_warning('Khong the nap snapshot quang chu dong tu source json', exc)
            return False

    def _build_snapshot_from_source(self, source_snapshot):
        subscribers = source_snapshot.get('subscribers') or []
        summary = source_snapshot.get('summary') or {}
        measured_at = source_snapshot.get('measured_at') or source_snapshot.get('generated_at') or ''
        generated_at = source_snapshot.get('generated_at') or measured_at
        ngay_bh_map = self._load_ngay_bh_map()
        alerts = []
        excluded = []
        active_counts = {}

        for raw_subscriber in subscribers:
            subscriber = _normalize_value(dict(raw_subscriber))
            subscriber['source'] = 'STY snapshot'
            subscriber['sa'] = _extract_sa(subscriber.get('port_id', ''))
            subscriber['ngay_bh'] = ngay_bh_map.get(str(subscriber.get('ma_tb', '')).strip(), '')

            exclude_reason = self._get_source_exclude_reason(subscriber)
            if exclude_reason:
                subscriber['exclude_reason'] = exclude_reason
                excluded.append(subscriber)
                continue

            doi_vt = subscriber.get('doi_vt') or 'Khác'
            active_counts[doi_vt] = active_counts.get(doi_vt, 0) + 1
            alerts.append(subscriber)

        wide_area = []
        exclusion_list = []
        seen_exclusion = set()
        for subscriber in excluded:
            if self._is_wide_area_row(subscriber):
                wide_area.append(subscriber)
            if subscriber.get('suppressed_by_pattern'):
                ma_tb = str(subscriber.get('ma_tb', '')).strip()
                if ma_tb and ma_tb in seen_exclusion:
                    continue
                if ma_tb:
                    seen_exclusion.add(ma_tb)
                exclusion_list.append({
                    'ma_tb': subscriber.get('ma_tb', ''),
                    'ten_tb': subscriber.get('ten_tb', ''),
                    'pattern_type': 'SOURCE_PATTERN',
                    'pattern_score': 1,
                    'total_events': '',
                    'first_detected': subscriber.get('first_off_time', ''),
                    'last_updated': generated_at,
                    'notes': subscriber.get('suppression_reason', 'Trong danh sách tắt chủ động'),
                    'source': 'STY snapshot',
                })

        wide_area_groups = self._group_wide_area_subscribers(excluded)
        pattern_exclusions = self._collect_pattern_exclusions(excluded)
        port_down_groups = self._group_port_down_subscribers(subscribers)

        alerts.sort(key=_sort_duration_key)

        return {
            'active': {
                'alerts': alerts,
                'excluded': excluded,
                'sources': [{
                    'name': 'STY snapshot',
                    'status': 'ok',
                    'count': len(subscribers),
                    'measured_at': measured_at,
                }],
            },
            'recovered': [],
            'wide_area': wide_area,
            'wide_area_groups': wide_area_groups,
            'exclusion_list': exclusion_list,
            'pattern_exclusions': pattern_exclusions,
            'port_down_groups': port_down_groups,
            'stats': {
                'active_by_doi_vt': active_counts,
                'recovered_24h_by_doi_vt': {},
                'summary': summary,
            },
            'cache_meta': {
                'status': 'ready',
                'refreshed_at': self._format_display_time(generated_at),
                'refresh_interval_seconds': QUANG_CHU_DONG_CACHE_REFRESH_SECONDS,
                'duration_ms': 0,
                'last_error': '',
                'pid': os.getpid(),
                'source_file': QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE,
                'measured_at': measured_at,
            },
        }

    def _get_source_exclude_reason(self, subscriber):
        if subscriber.get('suppression_reason'):
            return subscriber.get('suppression_reason')
        if subscriber.get('suppressed_by_wide_area'):
            return 'Sự cố diện rộng'
        if subscriber.get('suppressed_by_pattern'):
            return 'Trong danh sách tắt chủ động'
        return None

    def _group_wide_area_subscribers(self, excluded_rows):
        groups = {}
        for subscriber in excluded_rows:
            if not self._is_wide_area_row(subscriber):
                continue

            parent_port_key = self._resolve_parent_port_key(subscriber)
            group = groups.setdefault(parent_port_key, {
                'parent_port_key': parent_port_key,
                'olt_name': subscriber.get('olt_name', ''),
                'sa': _extract_sa(parent_port_key or subscriber.get('port_id', '')),
                'down_count': 0,
                'start_time': subscriber.get('first_off_time') or subscriber.get('alert_time') or '',
            })
            group['down_count'] += 1

            current_start = subscriber.get('first_off_time') or subscriber.get('alert_time') or ''
            if self._compare_times(current_start, group.get('start_time')) < 0:
                group['start_time'] = current_start
            if not group.get('olt_name') and subscriber.get('olt_name'):
                group['olt_name'] = subscriber.get('olt_name')

        return self._sort_wide_area_groups(groups.values())

    def _collect_pattern_exclusions(self, excluded_rows):
        rows = []
        for subscriber in excluded_rows:
            if not self._is_pattern_exclusion_row(subscriber):
                continue

            row = dict(subscriber)
            row['sa'] = row.get('sa') or _extract_sa(
                row.get('parent_port_key') or row.get('port_id', '')
            )
            rows.append(row)

        rows.sort(key=_sort_duration_key)
        return rows

    def _group_port_down_subscribers(self, rows):
        groups = {}
        for subscriber in rows:
            if str(subscriber.get('current_status') or '').strip().upper() != 'PORT_DOWN':
                continue

            parent_port_key = self._resolve_parent_port_key(subscriber)
            group = groups.setdefault(parent_port_key, {
                'parent_port_key': parent_port_key,
                'olt_name': subscriber.get('olt_name', ''),
                'sa': _extract_sa(parent_port_key or subscriber.get('port_id', '')),
                'down_count': 0,
                'start_time': subscriber.get('first_off_time') or subscriber.get('alert_time') or '',
            })
            group['down_count'] += 1

            current_start = subscriber.get('first_off_time') or subscriber.get('alert_time') or ''
            if self._compare_times(current_start, group.get('start_time')) < 0:
                group['start_time'] = current_start
            if not group.get('olt_name') and subscriber.get('olt_name'):
                group['olt_name'] = subscriber.get('olt_name')

        return self._sort_wide_area_groups(groups.values())

    def _is_wide_area_row(self, row):
        if str(row.get('current_status') or '').strip().upper() == 'PORT_DOWN':
            return False
        reason = str(row.get('suppression_reason') or row.get('exclude_reason') or '').strip().lower()
        return bool(row.get('suppressed_by_wide_area') or row.get('wide_area_sent') or reason in {'wide_area', 'sự cố diện rộng'})

    def _is_pattern_exclusion_row(self, row):
        reason = str(row.get('suppression_reason') or row.get('exclude_reason') or '').strip().lower()
        return bool(row.get('suppressed_by_pattern') or reason in {
            'pattern_exclusion',
            'trong danh sách tắt chủ động',
        })

    def _resolve_parent_port_key(self, row):
        parent_port_key = str(row.get('parent_port_key') or '').strip()
        if parent_port_key:
            return parent_port_key

        port_id = str(row.get('port_id') or '').strip()
        if ':' in port_id:
            return port_id.split(':', 1)[0]
        return port_id

    def _sort_wide_area_groups(self, groups):
        return sorted(
            groups,
            key=lambda item: (
                self._time_sort_key(item.get('start_time')),
                -int(item.get('down_count') or 0),
                str(item.get('parent_port_key') or ''),
            ),
        )

    def _time_sort_key(self, value):
        parsed = self._parse_datetime(value)
        if parsed is not None:
            return (0, parsed.timestamp())
        return (1, str(value or ''))

    def _compare_times(self, left, right):
        left_parsed = self._parse_datetime(left)
        right_parsed = self._parse_datetime(right)

        if left_parsed and right_parsed:
            if left_parsed < right_parsed:
                return -1
            if left_parsed > right_parsed:
                return 1
            return 0
        if left_parsed:
            return -1
        if right_parsed:
            return 1

        left_value = str(left or '')
        right_value = str(right or '')
        if left_value < right_value:
            return -1
        if left_value > right_value:
            return 1
        return 0

    def _parse_datetime(self, value):
        normalized = str(value or '').strip().replace('T', ' ')
        if not normalized:
            return None

        for time_format in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S'):
            try:
                return datetime.strptime(normalized, time_format)
            except ValueError:
                continue
        return None

    def _format_display_time(self, value):
        if not value:
            return ''

        normalized = str(value).replace('T', ' ')
        for time_format in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S'):
            try:
                return datetime.strptime(normalized, time_format).strftime('%d/%m/%Y %H:%M:%S')
            except ValueError:
                continue
        return normalized

    def _build_active_payload(self):
        all_outages = []
        exclusion_set = set()
        sources_info = []

        for db_config in QUANG_CHU_DONG_DB_LIST:
            db_name = db_config['name']
            db_path = db_config['path']

            if not os.path.exists(db_path):
                sources_info.append({'name': db_name, 'status': 'not_found', 'count': 0})
                continue

            try:
                outages = _fetch_rows(
                    db_path,
                    """
                    SELECT o.*
                    FROM outage_alerts o
                    WHERE NOT EXISTS (
                        SELECT 1 FROM recovery_alerts r
                        WHERE r.port_id = o.port_id
                        AND r.recovery_time > o.alert_time
                    )
                    ORDER BY o.alert_time DESC
                    """,
                    row_factory=True,
                )
                normalized_outages = []
                for row in outages:
                    record = _normalize_value(dict(row))
                    record['source'] = db_name
                    normalized_outages.append(record)
                all_outages.extend(normalized_outages)
                source_info = {'name': db_name, 'status': 'ok', 'count': len(normalized_outages)}
            except Exception as exc:
                self._log_warning(f'Khong the doc database quang chu dong {db_name}', exc)
                sources_info.append({'name': db_name, 'status': 'error', 'error': str(exc), 'count': 0})
                continue

            try:
                exclusion_rows = _fetch_rows(
                    db_path,
                    "SELECT ma_tb FROM pattern_exclusion_list WHERE is_active = 1",
                )
                exclusion_set.update(str(row[0]).strip() for row in exclusion_rows if row and row[0])
            except Exception as exc:
                self._log_warning(f'Khong the doc exclusion list quang chu dong {db_name}', exc)
                source_info['status'] = 'partial'
                source_info['error'] = f'exclusion_list: {exc}'

            sources_info.append(source_info)

        alerts = []
        excluded = []
        now = datetime.now()
        ngay_bh_map = self._load_ngay_bh_map()

        for outage in all_outages:
            exclude_reason = self._get_exclude_reason(outage, now, exclusion_set)
            outage['sa'] = _extract_sa(outage.get('port_id', ''))
            outage['ngay_bh'] = ngay_bh_map.get(str(outage.get('ma_tb', '')).strip(), '')

            if exclude_reason:
                outage['exclude_reason'] = exclude_reason
                excluded.append(outage)
            else:
                alerts.append(outage)

        return {
            'alerts': alerts,
            'excluded': excluded,
            'sources': sources_info,
        }

    def _build_recovered_payload(self):
        all_recovered = []
        cutoff_time = (datetime.now() - timedelta(hours=24)).strftime('%Y-%m-%d %H:%M:%S')

        for db_config in QUANG_CHU_DONG_DB_LIST:
            db_path = db_config['path']
            if not os.path.exists(db_path):
                continue

            try:
                rows = _fetch_rows(
                    db_path,
                    """
                    SELECT *
                    FROM recovery_alerts
                    WHERE recovery_time > ?
                    ORDER BY recovery_time DESC
                    """,
                    (cutoff_time,),
                    row_factory=True,
                )
                for row in rows:
                    record = _normalize_value(dict(row))
                    record['source'] = db_config['name']
                    all_recovered.append(record)
            except Exception as exc:
                self._log_warning(f'Khong the doc recovered alerts tu {db_config["name"]}', exc)

        return all_recovered

    def _build_wide_area_payload(self):
        all_wide_area = []
        cutoff_time = (datetime.now() - timedelta(hours=24)).strftime('%Y-%m-%d %H:%M:%S')

        for db_config in QUANG_CHU_DONG_DB_LIST:
            db_path = db_config['path']
            if not os.path.exists(db_path):
                continue

            try:
                rows = _fetch_rows(
                    db_path,
                    """
                    SELECT *
                    FROM wide_area_outage_alerts
                    WHERE alert_time > ?
                    ORDER BY alert_time DESC
                    """,
                    (cutoff_time,),
                    row_factory=True,
                )
                for row in rows:
                    record = _normalize_value(dict(row))
                    record['source'] = db_config['name']
                    all_wide_area.append(record)
            except Exception as exc:
                self._log_warning(f'Khong the doc wide area alerts tu {db_config["name"]}', exc)

        return all_wide_area

    def _build_exclusion_payload(self):
        exclusion_rows = []
        seen_ma_tb = set()

        for db_config in QUANG_CHU_DONG_DB_LIST:
            db_path = db_config['path']
            if not os.path.exists(db_path):
                continue

            try:
                rows = _fetch_rows(
                    db_path,
                    """
                    SELECT ma_tb, ten_tb, pattern_type, pattern_score,
                           total_events, first_detected, last_updated, notes
                    FROM pattern_exclusion_list
                    WHERE is_active = 1
                    ORDER BY pattern_score DESC
                    """,
                    row_factory=True,
                )
            except Exception as exc:
                self._log_warning(f'Khong the doc exclusion list tu {db_config["name"]}', exc)
                continue

            for row in rows:
                record = _normalize_value(dict(row))
                ma_tb = str(record.get('ma_tb', '')).strip()
                if ma_tb and ma_tb in seen_ma_tb:
                    continue
                if ma_tb:
                    seen_ma_tb.add(ma_tb)
                record['source'] = db_config['name']
                exclusion_rows.append(record)

        exclusion_rows.sort(
            key=lambda item: (
                float(item.get('pattern_score') or 0),
                str(item.get('last_updated') or ''),
            ),
            reverse=True,
        )
        return exclusion_rows

    def _build_stats_payload(self):
        active_counts = {}
        recovered_counts = {}

        for db_config in QUANG_CHU_DONG_DB_LIST:
            db_path = db_config['path']
            if not os.path.exists(db_path):
                continue

            try:
                active_rows = _fetch_rows(
                    db_path,
                    """
                    SELECT doi_vt, COUNT(*) as count
                    FROM outage_alerts o
                    WHERE NOT EXISTS (
                        SELECT 1 FROM recovery_alerts r
                        WHERE r.port_id = o.port_id
                        AND r.recovery_time > o.alert_time
                    )
                    GROUP BY doi_vt
                    """,
                )
                recovered_rows = _fetch_rows(
                    db_path,
                    """
                    SELECT doi_vt, COUNT(*) as count
                    FROM recovery_alerts
                    WHERE recovery_time > datetime('now', '-24 hours')
                    GROUP BY doi_vt
                    """,
                )
            except Exception as exc:
                self._log_warning(f'Khong the doc thong ke quang chu dong tu {db_config["name"]}', exc)
                continue

            for doi_vt, count in active_rows:
                active_counts[doi_vt] = active_counts.get(doi_vt, 0) + count
            for doi_vt, count in recovered_rows:
                recovered_counts[doi_vt] = recovered_counts.get(doi_vt, 0) + count

        return {
            'active_by_doi_vt': active_counts,
            'recovered_24h_by_doi_vt': recovered_counts,
        }

    def _summarize_source_errors(self, snapshot):
        source_errors = []
        for source in snapshot.get('active', {}).get('sources', []):
            if source.get('status') in {'error', 'partial'}:
                source_errors.append(f"{source.get('name')}: {source.get('error')}")
            elif source.get('status') == 'not_found':
                source_errors.append(f"{source.get('name')}: database not found")

        return '; '.join(source_errors)

    def _fallback_wide_area_groups(self, snapshot):
        groups = {}
        for row in snapshot.get('active', {}).get('excluded', []):
            if not self._is_wide_area_row(row):
                continue

            parent_port_key = self._resolve_parent_port_key(row)
            group = groups.setdefault(parent_port_key, {
                'parent_port_key': parent_port_key,
                'olt_name': row.get('olt_name', ''),
                'sa': _extract_sa(parent_port_key or row.get('port_id', '')),
                'down_count': 0,
                'start_time': row.get('first_off_time') or row.get('alert_time') or '',
            })
            group['down_count'] += 1
            current_start = row.get('first_off_time') or row.get('alert_time') or ''
            if self._compare_times(current_start, group.get('start_time')) < 0:
                group['start_time'] = current_start
            if not group.get('olt_name') and row.get('olt_name'):
                group['olt_name'] = row.get('olt_name')

        return self._sort_wide_area_groups(groups.values())

    def _fallback_pattern_exclusions(self, snapshot):
        rows = []
        for row in snapshot.get('active', {}).get('excluded', []):
            if not self._is_pattern_exclusion_row(row):
                continue

            normalized = dict(row)
            normalized['sa'] = normalized.get('sa') or _extract_sa(
                normalized.get('parent_port_key') or normalized.get('port_id', '')
            )
            rows.append(normalized)

        rows.sort(key=_sort_duration_key)
        return rows

    def _load_ngay_bh_map(self):
        brcd_excel_path = os.path.join(BASE_DATA_PATH, 'downloads', 'kq_dhsc', 'bc_BRCD.xlsx')
        if not os.path.exists(brcd_excel_path):
            return {}

        try:
            current_mtime = os.path.getmtime(brcd_excel_path)
        except OSError:
            return {}

        with self._state_lock:
            if self._ngay_bh_cache is not None and current_mtime == self._ngay_bh_cache_mtime:
                return self._ngay_bh_cache

        ngay_bh_map = {}
        try:
            df_brcd = read_excel_sheet_cached(brcd_excel_path, 'chi_tiet_ton_brcd')
            if 'ma_tb' not in df_brcd.columns or 'ngay_bh' not in df_brcd.columns:
                return {}

            for _, row in df_brcd.iterrows():
                ma_tb = str(row.get('ma_tb', '')).strip()
                if ma_tb:
                    ngay_bh_map[ma_tb] = _normalize_value(row.get('ngay_bh', ''))
        except Exception as exc:
            self._log_warning('Khong the doc bc_BRCD.xlsx de bo sung ngay_bh', exc)

        with self._state_lock:
            self._ngay_bh_cache = ngay_bh_map
            self._ngay_bh_cache_mtime = current_mtime

        return ngay_bh_map

    def _get_exclude_reason(self, outage, now, exclusion_set):
        if outage.get('wide_area_sent'):
            return 'Sự cố diện rộng'
        if str(outage.get('ma_tb', '')).strip() in exclusion_set:
            return 'Trong danh sách tắt chủ động'

        first_off_time = outage.get('first_off_time') or outage.get('alert_time')
        if not first_off_time:
            return None

        try:
            time_format = '%Y-%m-%d %H:%M:%S.%f' if '.' in str(first_off_time) else '%Y-%m-%d %H:%M:%S'
            off_time = datetime.strptime(str(first_off_time), time_format)
            off_hour = off_time.hour

            if off_hour < 6:
                return f'OFF trước 6h sáng ({off_hour}h)'
            if off_hour >= 18:
                return f'OFF sau 18h tối ({off_hour}h)'

            off_duration_hours = (now - off_time).total_seconds() / 3600
            if off_duration_hours >= 12:
                return f'OFF >= 12 tiếng ({off_duration_hours:.1f}h)'
        except Exception as exc:
            self._log_warning('Khong parse duoc first_off_time', exc)

        return None

    def _snapshot_file_is_fresh(self):
        if QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE and os.path.exists(QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE):
            return self._load_source_snapshot_if_newer()
        if not os.path.exists(QUANG_CHU_DONG_CACHE_FILE):
            return False
        return (time.time() - os.path.getmtime(QUANG_CHU_DONG_CACHE_FILE)) < QUANG_CHU_DONG_CACHE_REFRESH_SECONDS

    def _load_snapshot_from_disk_if_newer(self):
        if QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE and os.path.exists(QUANG_CHU_DONG_SOURCE_SNAPSHOT_FILE):
            return False
        if not os.path.exists(QUANG_CHU_DONG_CACHE_FILE):
            return False

        file_mtime = os.path.getmtime(QUANG_CHU_DONG_CACHE_FILE)
        with self._state_lock:
            if self._snapshot is not None and file_mtime <= self._snapshot_mtime:
                return True

        try:
            with open(QUANG_CHU_DONG_CACHE_FILE, 'r', encoding='utf-8') as cache_file:
                snapshot = json.load(cache_file)
            self._set_snapshot(snapshot, file_mtime)
            return True
        except Exception as exc:
            self._log_warning('Khong the nap snapshot quang chu dong tu dia', exc)
            return False

    def _write_snapshot(self, snapshot):
        cache_dir = os.path.dirname(QUANG_CHU_DONG_CACHE_FILE)
        os.makedirs(cache_dir, exist_ok=True)
        fd, temp_path = tempfile.mkstemp(prefix='quangchudong_', suffix='.json', dir=cache_dir)
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as temp_file:
                json.dump(_normalize_value(snapshot), temp_file, ensure_ascii=False)
            os.replace(temp_path, QUANG_CHU_DONG_CACHE_FILE)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def _set_snapshot(self, snapshot, snapshot_mtime, *, source_snapshot_mtime=0):
        with self._state_lock:
            self._snapshot = snapshot
            self._snapshot_mtime = snapshot_mtime
            self._source_snapshot_mtime = source_snapshot_mtime

    def _fallback_cache_meta(self):
        return {
            'status': 'warming' if self._snapshot is None else 'stale',
            'refreshed_at': '',
            'refresh_interval_seconds': QUANG_CHU_DONG_CACHE_REFRESH_SECONDS,
            'duration_ms': 0,
            'last_error': self._last_error or '',
            'pid': os.getpid(),
        }

    def _acquire_file_lock(self, lock_file):
        if fcntl is None:
            return True
        try:
            fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            return True
        except BlockingIOError:
            return False

    def _release_file_lock(self, lock_file):
        if fcntl is None:
            return
        try:
            fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass

    def _log_warning(self, message, exc):
        if self._app is not None:
            self._app.logger.warning('%s: %s', message, exc)


_QUANG_CHU_DONG_CACHE = QuangChuDongCache()


def initialize_quangchudong_cache(app, *, warm=False):
    _QUANG_CHU_DONG_CACHE.initialize(app, warm=warm)


def get_quangchudong_cache():
    return _QUANG_CHU_DONG_CACHE
