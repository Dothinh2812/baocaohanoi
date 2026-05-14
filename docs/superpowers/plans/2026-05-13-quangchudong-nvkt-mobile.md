# Quang Chu Dong NVKT Mobile Routes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `/quangchudong/<ten-nvkt>` mobile-first pages that show only the selected technician's DOWN alerts from 06:00 today.

**Architecture:** Keep the existing `/quangchudong` dashboard unchanged. Add small helper functions in the Quang Chu Dong blueprint for slug matching and time filtering, expose a focused JSON API for a single technician, and render a standalone mobile page that fetches that API.

**Tech Stack:** Flask blueprint, Jinja template, vanilla JavaScript, pytest.

---

### Task 1: Route and API Behavior

**Files:**
- Modify: `blueprints/quangchudong_routes.py`
- Test: `tests/test_quangchudong_routes.py`

- [ ] **Step 1: Write failing tests**

Add tests covering Vietnamese slug matching and filtering `first_off_time`/`alert_time` from 06:00 today.

- [ ] **Step 2: Run route tests and verify failure**

Run: `pytest tests/test_quangchudong_routes.py -v`

- [ ] **Step 3: Implement helpers, API route, and page route**

Add `_slugify_nvkt_name`, `_today_six_am`, `_filter_nvkt_alerts`, `/api/quangchudong/nvkt/<nvkt_slug>/alerts`, and `/quangchudong/<nvkt_slug>`.

- [ ] **Step 4: Run route tests and verify pass**

Run: `pytest tests/test_quangchudong_routes.py -v`

### Task 2: Mobile Template

**Files:**
- Create: `templates/quangchudong_nvkt_mobile.html`

- [ ] **Step 1: Build standalone HTML**

Use no `base.html`; include compact CSS and JavaScript that fetches `/api/quangchudong/nvkt/<slug>/alerts`.

- [ ] **Step 2: Verify template route renders**

Add/keep route test asserting `/quangchudong/le-van-tuan` contains the API URL and no base sidebar-specific structure.

### Task 3: Final Verification

**Files:**
- Test: `tests/test_quangchudong_routes.py`
- Optional broader check: `tests/test_route_policy.py`

- [ ] **Step 1: Run focused tests**

Run: `pytest tests/test_quangchudong_routes.py tests/test_route_policy.py -v`

- [ ] **Step 2: Inspect diff**

Run: `git diff -- blueprints/quangchudong_routes.py templates/quangchudong_nvkt_mobile.html tests/test_quangchudong_routes.py docs/superpowers/plans/2026-05-13-quangchudong-nvkt-mobile.md`
