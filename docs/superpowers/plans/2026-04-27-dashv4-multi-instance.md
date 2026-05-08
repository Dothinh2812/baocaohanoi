# Dashv4 Multi Instance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make one dashv4 codebase run as many isolated unit instances, each with its own DB, port, Cloudflare hostname, session/cache/logs, and user file.

**Architecture:** Keep Flask code shared and move all instance-specific paths into environment variables generated from `deploy/units.yaml`. Run one systemd template service per unit and route public hostnames through one Cloudflare Tunnel ingress file.

**Tech Stack:** Flask, Gunicorn, systemd, Cloudflare Tunnel, Python stdlib, pytest.

---

### Task 1: Instance Runtime Configuration

**Files:**
- Modify: `config.py`
- Modify: `auth.py`
- Modify: `gunicorn_config.py`
- Test: `tests/test_multi_instance_config.py`

- [ ] Write tests that reload `config` and `auth` with `DASHV4_UNIT_CODE`, `DASHV4_SESSION_FILE_DIR`, `DASHV4_CACHE_DIR`, `DASHV4_USER_FILE`, `DASHV4_LOGIN_LOG_FILE`, `DASHV4_LOG_DIR`, and `DASHV4_PID_FILE`.
- [ ] Verify the tests fail because these env vars are not fully supported yet.
- [ ] Add env-backed paths to `config.py`, including safe defaults under `runtime_app/<unit_code>` and `logs/<unit_code>`.
- [ ] Update `auth.py` to use env-backed user and login log files, creating the parent log directory.
- [ ] Update `gunicorn_config.py` to use env-backed log dir, pid file, and proc name.
- [ ] Run the new tests and existing dashboard tests.

### Task 2: Deployment Generator

**Files:**
- Create: `deploy/units.yaml`
- Create: `scripts/generate_instances.py`
- Test: `tests/test_generate_instances.py`

- [ ] Write tests for parsing unit config, rendering env files, rendering systemd service, and rendering Cloudflare ingress.
- [ ] Verify the tests fail because the generator does not exist.
- [ ] Implement a small YAML reader for the simple list-of-dicts format used in `deploy/units.yaml`.
- [ ] Implement validation for duplicate `code`, `port`, and `hostname`.
- [ ] Implement render functions that do not require root writes during tests.
- [ ] Add a dry-run CLI that writes generated files to an output directory.
- [ ] Run generator tests.

### Task 3: Operational Documentation And Verification

**Files:**
- Create: `deploy/systemd/dashv4@.service`
- Create: `deploy/cloudflared/config.generated.example.yml`
- Create: `deploy/README-multi-instance.md`

- [ ] Add a systemd template example using `/etc/dashv4/%i.env`.
- [ ] Add a Cloudflare Tunnel ingress example generated from sample units.
- [ ] Document how to generate env files, install service, update tunnel config, start instances, and smoke test.
- [ ] Run full test suite with `pytest`.
