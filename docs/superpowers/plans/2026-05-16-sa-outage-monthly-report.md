# SA Outage Monthly Report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a month-filtered full SA outage report below the current `/su_co_sa` live and cleared-today sections.

**Architecture:** Extend the existing `blueprints/sa_outage_routes.py` API payload with a monthly report section sourced from `sa_outage_incidents`. Update `templates/pages/su_co_sa.html` to pass the selected month to the API and render summary stats plus a table below the existing sections.

**Tech Stack:** Flask blueprint, SQLite, Jinja template, vanilla JavaScript, pytest.

---

### Task 1: API Monthly Report

**Files:**
- Modify: `blueprints/sa_outage_routes.py`
- Test: `tests/test_sa_outage_routes.py`

- [ ] Write failing tests for default month, explicit `YYYY-MM` month, and invalid month fallback.
- [ ] Run the route tests and confirm the monthly report assertions fail before implementation.
- [ ] Add month parsing, cache keying by month, and a query using `strftime('%Y-%m', started_at) = ?`.
- [ ] Return `monthly_report` with `selected_month`, `summary`, and `incidents`.
- [ ] Run the route tests and confirm they pass.

### Task 2: Page Rendering

**Files:**
- Modify: `templates/pages/su_co_sa.html`

- [ ] Add a month input and monthly report section below the existing sections.
- [ ] Update `loadData()` to request `/api/su-co-sa/data?month=YYYY-MM`.
- [ ] Render monthly summary counts and rows using `max_off_count` as affected subscriber count.
- [ ] Keep existing refresh behavior and make refresh reload the selected month.

### Task 3: Verification

**Files:**
- Run tests that cover SA outage route behavior.
- Optionally run the full test suite if the focused tests pass quickly.

- [ ] Run `pytest tests/test_sa_outage_routes.py -v`.
- [ ] Inspect `git diff` to verify only intended files changed.

### Task 4: Per-Incident User Notes

**Files:**
- Modify: `blueprints/sa_outage_routes.py`
- Modify: `templates/pages/su_co_sa.html`
- Test: `tests/test_sa_outage_routes.py`

- [ ] Write failing tests for saving user notes when the database does not yet have note columns.
- [ ] Add idempotent nullable columns `user_notes`, `user_notes_updated_at`, and `user_notes_updated_by`.
- [ ] Add `POST /api/su-co-sa/incidents/<id>/notes` to update only the dashboard-owned note columns.
- [ ] Include `user_notes` fields in monthly report payload rows.
- [ ] Render a textarea and per-row `Luu` button in the monthly report.
- [ ] Run `pytest tests/test_sa_outage_routes.py -v`.
