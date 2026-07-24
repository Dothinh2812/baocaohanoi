# AGENTS.md

Single Flask dashboard app (`dashboard.py` → `app`) for Vietnam Telecom units (TTVT). UI and comments are Vietnamese. Forked from `dashv3`, backend rewired to read-only SQLite from `baocaohanoi/api_transition`.

## Commands

Run everything **from the repo root**. The repo is a flat module tree (no `setup.py`/`pyproject.toml`); modules import each other as top-level (`import config`, `from repositories...`, `from blueprints...`).

- **Tests:** `python3 -m pytest tests/` — uses the *system* pytest (`~/.local/bin/pytest`, 9.x). The bundled `venv/` has **no pytest**; do not run `venv/bin/pytest`.
- **Single test:** `python3 -m pytest tests/test_route_policy.py -q` or `-k <name>`.
- Tests are fast (~2s, 47 tests) and self-contained: they build sqlite DBs in `tmp_path` and `monkeypatch` config attrs — they do **not** touch the real `report_history.db` or the external host data dirs.
- There is **no lint, typecheck, or formatter configured** — don't invent one. Verify Python syntax with `python3 -m py_compile <file>`.
- **Dev server:** `python3 dashboard.py` (binds `0.0.0.0:5011` by default).
- **Production:** `./start_dashboard.sh` → `gunicorn -c gunicorn_config.py dashboard:app` using `/home/vtst/.local/bin/gunicorn`. Do not run both at once (port conflict on 5011).

## Dependencies

No lockfile. The runtime dependency set lives only in `venv/`: Flask 3.1.2, pandas 2.2.3, openpyxl, flask_session, xlsxwriter, gunicorn 23 (Python 3.10). If `venv/` is missing or broken, recreate manually — there is no `requirements.txt` to restore from.

## Critical import ordering (will break subtly if wrong)

- `runtime_limits.py` **must** import before pandas/numpy. `dashboard.py` does `import runtime_limits` first; preserve this when adding new entrypoints. It caps native BLAS threads (`DASHV4_NATIVE_THREADS`, default 1) so multi-worker gunicorn doesn't oversubscribe.
- `config.py` calls `time.tzset()` at import time and resolves all paths. Side effects happen on import — tests that change env vars `importlib.reload(config)` then `importlib.reload(dashboard)` (see `tests/test_route_policy.py`).

## Config & runtime environment

All config is env-driven via `DASHV4_*` vars (legacy `DASH_*` fallbacks) read in `config.py`. Key ones: `DASHV4_DB_PATH` (report_history.db), `DASHV4_UNIT_CODE`, `DASHV4_PORT`, `DASHV4_RUNTIME_DIR`, `DASHV4_SECRET_KEY`.

The app reads many **hardcoded host paths outside the repo** under `/home/vtst/baocaohanoi/...`, `/home/vtst/baocao-vattu/...`, `/home/vtst/do_chu_dong_api/...`, etc. (see `config.py`). These exist only on the deploy host; the app will return 404/empty if they're missing. Override via the matching `DASH_*` env var when needed.

`runtime_app/`, `flask_session/`, `logs/`, `*.db`, `username.xlsx`, `dashboard_data_paths.xlsx` are gitignored — per-instance runtime dirs are created at startup, not committed.

## Architecture

- `dashboard.py` — Flask app factory entrypoint. Registers 9 blueprints; `@app.before_request` enforces auth + route policy globally.
- `blueprints/` — route handlers (`*_routes.py`), one Flask blueprint per domain. The endpoint names here (e.g. `quality.page_chatluong`, `quangchudong.get_dashboard_payload`) are the canonical keys used by route policy and `deploy/units.yaml`.
- `repositories/` — data access. `sqlite_runtime.py` opens `report_history.db` **read-only** via `file:...?mode=ro&immutable=1` URI. `dashboard_views.py` = consumer views; `report_history_by_date.py` = canonical date-filtered access. `tiep_thi_source.py` = separate tiep_thi DB.
- `services/` — background services (`quangchudong_cache.py`), off by default (`DASHV4_ENABLE_BACKGROUND_SERVICES=0`); warmed in gunicorn `post_fork`.
- `app_helpers.py` — Excel read cache (`read_excel_sheet_cached`), DataFrame→payload serializers, CSRF helpers (`csrf_protect` decorator), safe file serving.
- `auth.py` — users in `username.xlsx` (werkzeug password hashes), flask_session filesystem sessions.
- `route_policy.py` + `config.py` `PUBLIC_ENDPOINTS` / `DISABLED_*` — see below.

## Route policy & multi-instance deploy

One shared codebase runs **one gunicorn process per unit** (`dashv4@<unit_code>` systemd template, ports 5011–5028), fronted by a Cloudflare tunnel. Generated from `deploy/units.yaml`:

```bash
python3 scripts/generate_instances.py --units-file deploy/units.yaml --output-dir deploy/generated
```

Per-instance route policy uses Flask endpoint names (CSV) via `DASHV4_DISABLED_ENDPOINTS` / `DASHV4_ENABLED_ENDPOINTS`:
- globally-disabled endpoints live in `config.py` (`DISABLED_PAGE_ENDPOINTS`, `DISABLED_NONPAGE_ENDPOINTS`) — these return a 501 pending screen / JSON with a `required_display_contract`.
- `enabled_endpoints` in `units.yaml` re-opens a globally-disabled route **only for that instance** (verify the per-unit data source first).
- enforcement + sidebar hiding (`is_endpoint_enabled` injected into Jinja) is centralized in `dashboard.py` `enforce_auth_policy` and `route_policy.py`.

List all endpoint names: see the snippet in `deploy/README-multi-instance.md`. Durable policy goes in `units.yaml`; direct edits to `/etc/dashv4/*.env` get overwritten on regenerate.

## Date-filtering contract (docs/09 — mandatory for any route reading report_history.db)

This is the most error-prone area. Before touching a route that reads `report_history.db`:
- Canonical date source is **`bao_cao_tong_hop_ngay.ngay_du_lieu`** only. Never `__imported_at`, `thoi_gian_tao/cap_nhat`, or request time.
- Date-enabled routes **must not** read `v_*_moi_nhat` / latest-snapshot views as the primary source. Query raw tables joined through `__sheet_id → sheet_bao_cao_tong_hop → bao_cao_tong_hop_ngay`.
- Use the shared helpers in `repositories/report_history_by_date.py` (`resolve_date_context`, `load_table_by_date`, `get_available_dates`). One page = one `selected_date`. Never silently fall back to a different date when the requested one is missing.
- Payload contract: include `selected_date`, `latest_available_date`, `available_dates`, `date_has_data`.
- Hide technical columns (`id`, `__snapshot_id`, `__sheet_id`, `__row_num`, `__row_hash`, `__imported_at`) via the serializer layer, not per-route. Set is `_HIDDEN_DASHBOARD_COLUMNS` in `app_helpers.py`.

## Doc-sync rule (enforced by README)

Any change touching data source, route→data mapping, or date filtering must update these in the same change:
- `docs/08-trang-thai-thuc-thi.md` (current route status: migrated / 501 / date-supported)
- `docs/04-mapping-route-va-du-lieu.md` (route ↔ `ma_bao_cao` ↔ `ten_bang_du_lieu` ↔ `supports_date`)
- `docs/09-nguyen-tac-loc-ngay-report-history.md` (only if the shared date principle changes)

Start unfamiliar work by reading `docs/00-doc-index.md`, then `docs/08` (status) and `docs/04` (mapping). Don't migrate a route halfway — if a screen lacks a data contract, mark it 501 rather than faking with patched-together Excel data.
