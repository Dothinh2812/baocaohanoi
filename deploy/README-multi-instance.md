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

## Per-Instance Route Policy

Each unit in `deploy/units.yaml` can define route policy by Flask endpoint name:

```yaml
disabled_endpoints:
  - quangchudong.page_quangchudong
  - sa_outage.page_su_co_sa
enabled_endpoints:
  - quality.page_shc_processing
```

`disabled_endpoints` blocks a route only for that instance. Page routes return the pending-feature screen and are hidden from the sidebar. API routes return HTTP 501 with `{"error": "route_disabled"}`.

`enabled_endpoints` re-opens a route that is globally disabled in `config.py`. Use this only after confirming that the route reads data from the correct per-unit source.

List all route endpoint names:

```bash
cd /home/vtst/dashv4
python3 -c "import dashboard; rows=[]; [rows.append('{:<45s} {:<10s} {}'.format(r.endpoint, ','.join(sorted(r.methods - {'HEAD','OPTIONS'})), r.rule)) for r in dashboard.app.url_map.iter_rules()]; print('\n'.join(sorted(rows)))"
```

For a quick runtime-only change, edit the instance env file directly and restart only that instance:

```bash
sudo editor /etc/dashv4/hoai_duc.env
```

```env
DASHV4_DISABLED_ENDPOINTS=quangchudong.page_quangchudong,sa_outage.page_su_co_sa,operations.page_pttb,operations.page_brcd
```

```bash
sudo systemctl restart dashv4@hoai_duc
```

Direct edits in `/etc/dashv4/*.env` can be overwritten by regenerated env files. Record durable policy in `deploy/units.yaml`.

Common page endpoints:

- `operations.page_brcd`
- `operations.page_pttb`
- `operations.page_cau_hinh_tu_dong`
- `operations.page_tong_hop_bsc_kpi`
- `quality.page_chatluong`
- `quality.page_i15`
- `quality.page_i15k2`
- `quality.page_shc_processing`
- `quangchudong.page_quangchudong`
- `sa_outage.page_su_co_sa`
- `growth.page_tiepthi`
- `growth.page_ngungpsc`
- `retention.page_giahan`
- `inventory.page_tra_cuu_nhanh_vat_tu`
- `inventory.page_tong_hop_tien`
- `inventory.page_tam_dung_khoi_phuc`

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

• Chưa sẵn sàng để chạy thật bằng:

  sudo systemctl start dashv4@son_tay

  Tôi vừa kiểm tra trực tiếp, hiện còn thiếu các thành phần systemd production:

  - Thiếu /etc/systemd/system/dashv4@.service
  - Thiếu /etc/dashv4/son_tay.env
  - Chưa có unit systemd dashv4@son_tay
  - Service template generated đang dùng /usr/bin/gunicorn, nhưng máy hiện chỉ có /home/vtst/.local/bin/gunicorn
  - File generated deploy/generated/env/son_tay.env vẫn có DASHV4_SECRET_KEY=change-me-son_tay, chưa phù hợp production

  Các phần đã có đủ:

  - deploy/generated/env/son_tay.env đã tồn tại
  - /home/vtst/dashv4/runtime_app/son_tay/users.xlsx đã có
  - /home/vtst/dashv4/logs/son_tay đã có
  - DB Sơn Tây đã có: /home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db
  - Python packages chính có đủ: flask, flask_session, pandas, openpyxl, gunicorn

  Để chạy thật, cần làm tối thiểu:

  sudo mkdir -p /etc/dashv4
  sudo cp deploy/generated/env/son_tay.env /etc/dashv4/son_tay.env
  sudo cp deploy/generated/systemd/dashv4@.service /etc/systemd/system/

  Sau đó sửa 2 điểm trước khi start:

  1. Đổi DASHV4_SECRET_KEY=change-me-son_tay trong /etc/dashv4/son_tay.env thành secret thật.
  2. Sửa service dùng đúng gunicorn, vì /usr/bin/gunicorn chưa có. Hoặc cài gunicorn vào /usr/bin, hoặc sửa ExecStart thành:

  ExecStart=/home/vtst/.local/bin/gunicorn -c gunicorn_config.py dashboard:app

  Rồi chạy:

  sudo systemctl daemon-reload
  sudo systemctl start dashv4@son_tay
  sudo systemctl status dashv4@son_tay --no-pager

  Tóm lại: cấu hình trong repo đã có, runtime Sơn Tây đã có, nhưng systemd production trên máy này chưa được cài đặt hoàn chỉnh.

  1. Chạy thử bằng Gunicorn
  Trong repo:

  cd /home/vtst/dashv4
  ./start_dashboard.sh

  Script này đã chạy Gunicorn, không còn chạy python3 dashboard.py.

  Kiểm tra:

  ps -ef | grep 'gunicorn.*dashboard:app' | grep -v grep
  curl -I http://127.0.0.1:5011/

  Dừng bằng Ctrl+C nếu đang chạy trong terminal/tmux.

  2. Chạy bền bằng systemd
  Trước hết dừng bản đang chạy trong terminal/tmux để tránh trùng port 5011.

  sudo mkdir -p /etc/dashv4

  sudo tee /etc/dashv4/son_tay.env >/dev/null <<'EOF'
  DASHV4_UNIT_CODE=son_tay
  DASHV4_UNIT_NAME=TTVT Sơn Tây
  DASHV4_HOST=0.0.0.0
  DASHV4_PORT=5011
  DASHV4_DB_PATH=/home/vtst/bchn/runtime/son_tay/sqlite_history/report_history.db
  DASHV4_RUNTIME_DIR=/home/vtst/dashv4/runtime_app/son_tay
  DASHV4_SESSION_FILE_DIR=/home/vtst/dashv4/runtime_app/son_tay/flask_session
  DASHV4_CACHE_DIR=/home/vtst/dashv4/runtime_app/son_tay/cache
  DASHV4_LOG_DIR=/home/vtst/dashv4/logs/son_tay
  DASHV4_PID_FILE=/tmp/dashv4-son_tay.pid
  DASHV4_NATIVE_THREADS=1
  DASHV4_WORKERS=2
  DASHV4_GUNICORN_TIMEOUT=30
  DASHV4_LOG_REQUESTS=1
  EOF

  Cài unit systemd:

  sudo cp /home/vtst/dashv4/deploy/systemd/dashv4@.service /etc/systemd/system/
  sudo systemctl daemon-reload

  Start service:

  sudo systemctl enable --now dashv4@son_tay

  Kiểm tra:

  sudo systemctl status dashv4@son_tay --no-pager
  journalctl -u dashv4@son_tay -f
  tail -f /home/vtst/dashv4/logs/son_tay/gunicorn_error.log

  Các lệnh vận hành thường dùng:

  sudo systemctl restart dashv4@son_tay
  sudo systemctl stop dashv4@son_tay
  sudo systemctl start dashv4@son_tay

  Quan trọng: sau khi dùng systemd thì không chạy thêm ./start_dashboard.sh hoặc python3 dashboard.py cùng lúc, vì sẽ tranh port 5011.
