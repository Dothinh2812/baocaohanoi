# Dashv4 Multi-Instance Deployment

## Model

One shared codebase runs one process per unit:

```text
Cloudflare Tunnel
  son-tay.dashboard.example.vn  -> http://127.0.0.1:5011 -> dashv4@son_tay
  ba-dinh.dashboard.example.vn  -> http://127.0.0.1:5012 -> dashv4@ba_dinh
```

Each instance has its own:

- `DASHV4_DB_PATH`
- `DASHV4_PORT`
- session directory
- cache directory
- user file
- login log
- gunicorn logs
- pid file

## Configure Units

Edit `deploy/units.yaml` and replace `dashboard.example.vn` with the real domain suffix.

Keep these fields unique:

- `code`
- `port`
- `hostname`

## Generate Files

Dry-run into the repo:

```bash
python3 scripts/generate_instances.py \
  --units-file deploy/units.yaml \
  --output-dir deploy/generated \
  --tunnel-id YOUR_TUNNEL_ID \
  --credentials-file /root/.cloudflared/YOUR_TUNNEL_ID.json
```

The generated tree contains:

```text
deploy/generated/env/*.env
deploy/generated/systemd/dashv4@.service
deploy/generated/cloudflared/config.yml
deploy/generated/smoke_test.sh
```

## Install On Server

Copy env files and service template:

```bash
sudo mkdir -p /etc/dashv4
sudo cp deploy/generated/env/*.env /etc/dashv4/
sudo cp deploy/generated/systemd/dashv4@.service /etc/systemd/system/
sudo systemctl daemon-reload
```

Before starting services, replace every `DASHV4_SECRET_KEY=change-me-*` in `/etc/dashv4/*.env` with a stable private value.

Create initial user files for each unit:

```bash
mkdir -p /home/vtst/dashv4/runtime_app/son_tay
cp /home/vtst/dashv4/username.xlsx /home/vtst/dashv4/runtime_app/son_tay/users.xlsx
```

Repeat for each unit, or copy a prepared per-unit user list.

Start services:

```bash
sudo systemctl enable --now dashv4@son_tay
sudo systemctl enable --now dashv4@ba_dinh
```

Check logs:

```bash
journalctl -u dashv4@son_tay -f
tail -f /home/vtst/dashv4/logs/son_tay/gunicorn_error.log
```

## Cloudflare Tunnel

Merge `deploy/generated/cloudflared/config.yml` into the active cloudflared config, then restart cloudflared:

```bash
sudo systemctl restart cloudflared
```

The app instances should bind to `127.0.0.1`, so no dashboard port needs to be opened publicly.

## Smoke Test

Local ports:

```bash
deploy/generated/smoke_test.sh
```

Public hostnames:

```bash
curl -I https://son-tay.dashboard.example.vn/
curl -I https://ba-dinh.dashboard.example.vn/
```

Then verify login, `/admin/users`, password change, and one data page per unit.
