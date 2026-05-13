# Triển khai multi-instance cho dashv4

## Mục tiêu

Tài liệu này ghi lại hiện trạng multi-instance của `dashv4` tại ngày `2026-05-12` và cách chạy nhiều instance song song từ cùng một codebase.

Mô hình mục tiêu:

```text
Một thư mục code /home/vtst/dashv4
  -> nhiều process gunicorn
  -> mỗi process đọc một file env riêng
  -> mỗi đơn vị có port, DB, session, cache, user, log và pid riêng
  -> Cloudflare Tunnel hoặc reverse proxy route hostname về port nội bộ
```

Ví dụ:

```text
son-tay.dashboard.example.vn  -> http://127.0.0.1:5011 -> dashv4@son_tay
ba-dinh.dashboard.example.vn  -> http://127.0.0.1:5012 -> dashv4@ba_dinh
```

## Hiện trạng đã có

### Cấu hình instance

File nguồn khai báo instance:

- `deploy/units.yaml`

Hiện file này đã khai báo 18 đơn vị, port từ `5011` đến `5028`, mỗi đơn vị có:

- `code`: mã instance dùng cho systemd và tên env file
- `slug`: dạng URL-friendly
- `name`: tên hiển thị
- `port`: port nội bộ riêng
- `hostname`: hostname public dự kiến
- `db_path`: đường dẫn `report_history.db` riêng của đơn vị

Các hostname trong repo vẫn đang là dạng mẫu `*.dashboard.example.vn`; khi triển khai thật phải đổi sang domain thực tế.

### Runtime path đã tách theo instance

`config.py` đã hỗ trợ các biến môi trường sau:

```bash
DASHV4_UNIT_CODE
DASHV4_UNIT_NAME
DASHV4_APP_NAME
DASHV4_HOST
DASHV4_PORT
DASHV4_DB_PATH
DASHV4_SECRET_KEY
DASHV4_TIMEZONE
DASHV4_RUNTIME_DIR
DASHV4_SESSION_FILE_DIR
DASHV4_CACHE_DIR
DASHV4_USER_FILE
DASHV4_LOGIN_LOG_FILE
DASHV4_LOG_DIR
DASHV4_PID_FILE
DASHV4_SESSION_COOKIE_SECURE
```

Generator sẽ sinh mặc định theo cấu trúc:

```text
/home/vtst/dashv4/runtime_app/<unit_code>/
  users.xlsx
  flask_session/
  cache/

/home/vtst/dashv4/logs/<unit_code>/
  login_history.csv
  gunicorn_access.log
  gunicorn_error.log

/tmp/dashv4-<unit_code>.pid
```

### Deployment generator

Script đã có:

- `scripts/generate_instances.py`

Script này đọc `deploy/units.yaml` và sinh:

```text
deploy/generated/env/*.env
deploy/generated/systemd/dashv4@.service
deploy/generated/cloudflared/config.yml
deploy/generated/smoke_test.sh
```

Script có kiểm tra trùng:

- `code`
- `port`
- `hostname`

### Systemd template

Service template đã có:

- `deploy/systemd/dashv4@.service`
- `deploy/generated/systemd/dashv4@.service`

Service dùng file env theo instance:

```ini
EnvironmentFile=/etc/dashv4/%i.env
ExecStart=/usr/bin/gunicorn -c gunicorn_config.py dashboard:app
```

Khi chạy:

```bash
sudo systemctl enable --now dashv4@son_tay
sudo systemctl enable --now dashv4@ba_dinh
```

systemd sẽ đọc:

```text
/etc/dashv4/son_tay.env
/etc/dashv4/ba_dinh.env
```

### Gunicorn

`gunicorn_config.py` đã lấy bind, process name, pid và log từ config:

```text
bind      = DASHV4_HOST:DASHV4_PORT
proc_name = dashv4_<unit_code>
pidfile   = DASHV4_PID_FILE
accesslog = DASHV4_LOG_DIR/gunicorn_access.log
errorlog  = DASHV4_LOG_DIR/gunicorn_error.log
```

Lưu ý vận hành: số worker mặc định là 2 cho từng instance. Có thể chỉnh bằng `DASHV4_WORKERS` trong file env của instance.

### Cloudflare Tunnel

Generator sinh file:

```text
deploy/generated/cloudflared/config.yml
```

Mỗi hostname trỏ về port nội bộ tương ứng:

```yaml
ingress:
  - hostname: son-tay.dashboard.example.vn
    service: http://127.0.0.1:5011

  - hostname: ba-dinh.dashboard.example.vn
    service: http://127.0.0.1:5012

  - service: http_status:404
```

## Cách chạy kiểm thử cục bộ

### 1. Chạy một instance bằng biến môi trường

Ví dụ chạy Sơn Tây bằng Flask dev server:

```bash
cd /home/vtst/dashv4

DASHV4_UNIT_CODE=son_tay \
DASHV4_UNIT_NAME="TTVT Sơn Tây" \
DASHV4_HOST=127.0.0.1 \
DASHV4_PORT=5011 \
DASHV4_DB_PATH=/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db \
DASHV4_RUNTIME_DIR=/home/vtst/dashv4/runtime_app/son_tay \
DASHV4_SESSION_FILE_DIR=/home/vtst/dashv4/runtime_app/son_tay/flask_session \
DASHV4_CACHE_DIR=/home/vtst/dashv4/runtime_app/son_tay/cache \
DASHV4_USER_FILE=/home/vtst/dashv4/runtime_app/son_tay/users.xlsx \
DASHV4_LOGIN_LOG_FILE=/home/vtst/dashv4/logs/son_tay/login_history.csv \
DASHV4_LOG_DIR=/home/vtst/dashv4/logs/son_tay \
DASHV4_PID_FILE=/tmp/dashv4-son_tay.pid \
python3 dashboard.py
```

Nếu chưa có user file:

```bash
mkdir -p /home/vtst/dashv4/runtime_app/son_tay
cp /home/vtst/dashv4/username.xlsx /home/vtst/dashv4/runtime_app/son_tay/users.xlsx
```

### 2. Chạy nhiều instance bằng nhiều terminal

Terminal 1:

```bash
cd /home/vtst/dashv4
DASHV4_UNIT_CODE=son_tay DASHV4_PORT=5011 DASHV4_DB_PATH=/path/to/son_tay/report_history.db python3 dashboard.py
```

Terminal 2:

```bash
cd /home/vtst/dashv4
DASHV4_UNIT_CODE=ba_dinh DASHV4_PORT=5012 DASHV4_DB_PATH=/path/to/ba_dinh/report_history.db python3 dashboard.py
```

Cách này chỉ dùng để kiểm thử nhanh. Khi production, dùng systemd template.

### 3. Sinh thử artifact triển khai

Không cần quyền root nếu sinh ra `/tmp`:

```bash
cd /home/vtst/dashv4

python3 scripts/generate_instances.py \
  --units-file deploy/units.yaml \
  --output-dir /tmp/dashv4-generated-check \
  --tunnel-id CHECK-TUNNEL \
  --credentials-file /tmp/check-tunnel.json
```

Kiểm tra kết quả:

```bash
find /tmp/dashv4-generated-check -maxdepth 3 -type f | sort
sed -n '1,80p' /tmp/dashv4-generated-check/env/son_tay.env
sed -n '1,120p' /tmp/dashv4-generated-check/cloudflared/config.yml
```

## Cách triển khai production

### 1. Chuẩn bị `deploy/units.yaml`

Cập nhật các trường sau trước khi sinh file:

```yaml
- code: son_tay
  slug: son-tay
  name: "TTVT Sơn Tây"
  port: 5011
  hostname: son-tay.<domain-thuc-te>
  db_path: /home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db
```

Yêu cầu:

- `code` không trùng
- `port` không trùng và chưa bị service khác dùng
- `hostname` không trùng
- `db_path` tồn tại trên server production
- mỗi instance có user file riêng hoặc danh sách user được copy riêng

### 2. Sinh file triển khai

```bash
cd /home/vtst/dashv4

python3 scripts/generate_instances.py \
  --units-file deploy/units.yaml \
  --output-dir deploy/generated \
  --tunnel-id <CLOUDFLARED_TUNNEL_ID> \
  --credentials-file /root/.cloudflared/<CLOUDFLARED_TUNNEL_ID>.json
```

### 3. Cài env và systemd service

```bash
sudo mkdir -p /etc/dashv4
sudo cp deploy/generated/env/*.env /etc/dashv4/
sudo cp deploy/generated/systemd/dashv4@.service /etc/systemd/system/
sudo systemctl daemon-reload
```

Sửa secret key trong từng file:

```bash
sudo editor /etc/dashv4/son_tay.env
sudo editor /etc/dashv4/ba_dinh.env
```

Không để production dùng giá trị:

```text
DASHV4_SECRET_KEY=change-me-...
```

Secret phải ổn định qua restart; nếu đổi secret, session đăng nhập cũ sẽ mất hiệu lực.

### 4. Chuẩn bị thư mục runtime và user file

Ví dụ cho Sơn Tây:

```bash
mkdir -p /home/vtst/dashv4/runtime_app/son_tay
mkdir -p /home/vtst/dashv4/logs/son_tay
cp /home/vtst/dashv4/username.xlsx /home/vtst/dashv4/runtime_app/son_tay/users.xlsx
```

Lặp lại cho từng đơn vị, hoặc chuẩn bị file `users.xlsx` riêng từng đơn vị.

Kiểm tra quyền:

```bash
test -r /home/vtst/dashv4/runtime_app/son_tay/users.xlsx
test -w /home/vtst/dashv4/runtime_app/son_tay/users.xlsx
test -d /home/vtst/dashv4/logs/son_tay
```

### 5. Kiểm tra DB trước khi start

Ví dụ:

```bash
test -r /home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db
test -r /home/vtst/baocaohanoi/api_transition/runtime/ba_dinh/sqlite_history/report_history.db
```

Nếu DB không tồn tại hoặc user chạy service không đọc được, các route đọc `report_history.db` sẽ lỗi.

### 6. Start service

Start một instance trước:

```bash
sudo systemctl enable --now dashv4@son_tay
sudo systemctl status dashv4@son_tay --no-pager
```

Xem log:

```bash
journalctl -u dashv4@son_tay -f
tail -f /home/vtst/dashv4/logs/son_tay/gunicorn_error.log
```

Khi instance đầu chạy ổn, start các instance còn lại:

```bash
sudo systemctl enable --now dashv4@ba_dinh
sudo systemctl enable --now dashv4@cau_giay
```

### 7. Cấu hình Cloudflare Tunnel

Merge nội dung:

```text
deploy/generated/cloudflared/config.yml
```

vào config cloudflared đang chạy trên server.

Sau đó restart:

```bash
sudo systemctl restart cloudflared
sudo systemctl status cloudflared --no-pager
```

Port dashboard nên bind `127.0.0.1`, không mở trực tiếp ra public.

### 8. Smoke test

Kiểm tra local port:

```bash
deploy/generated/smoke_test.sh
```

Kiểm tra từng hostname:

```bash
curl -I https://son-tay.<domain-thuc-te>/
curl -I https://ba-dinh.<domain-thuc-te>/
```

Kiểm tra bằng trình duyệt:

- mở trang chủ từng hostname
- đăng nhập
- mở `/admin/users`
- đổi mật khẩu thử với một user test
- mở ít nhất một page đọc `report_history.db`
- xác nhận tên đơn vị và dữ liệu đúng instance

## Checklist production

Trước go-live:

- `deploy/units.yaml` đã dùng hostname thật
- `deploy/generated/env/*.env` đã copy vào `/etc/dashv4`
- mọi `DASHV4_SECRET_KEY=change-me-*` đã được thay
- mỗi instance có `users.xlsx` riêng
- mỗi `DASHV4_DB_PATH` tồn tại và đọc được
- không trùng port với service khác
- systemd service đã `daemon-reload`
- Cloudflare Tunnel route đúng hostname -> port
- smoke test local port đạt
- smoke test public hostname đạt

Sau go-live:

- theo dõi `journalctl -u dashv4@<unit>`
- theo dõi `logs/<unit>/gunicorn_error.log`
- kiểm tra dung lượng `runtime_app/<unit>/flask_session`
- kiểm tra `logs/<unit>/login_history.csv`
- định kỳ backup hoặc quản lý riêng `runtime_app/<unit>/users.xlsx`

## Rủi ro và giới hạn hiện tại

### Worker gunicorn có thể quá nhiều

`gunicorn_config.py` đang dùng `DASHV4_WORKERS`, mặc định 2 worker cho mỗi instance. Nếu server chạy nhiều instance và tải thấp, có thể giảm xuống 1 worker/instance.

Khuyến nghị production:

- đặt `DASHV4_WORKERS=1` hoặc `DASHV4_WORKERS=2` tùy tải thực tế
- giám sát RAM sau khi bật nhiều instance

### Auth chỉ tách hoàn toàn khi env đầy đủ

`auth.py` có hỗ trợ `DASHV4_USER_FILE` và `DASHV4_LOGIN_LOG_FILE`. Khi chạy bằng env generator thì an toàn.

Không nên chỉ set `DASHV4_UNIT_CODE` rồi chạy app, vì khi thiếu hai biến trên, auth có thể rơi về:

```text
username.xlsx
logs/login_history.csv
```

### Một số nguồn dữ liệu vẫn là nguồn chung

Một số route/nguồn chưa tách theo đơn vị hoặc vẫn phụ thuộc runtime ngoài `report_history.db`, ví dụ:

- sự cố SA đọc `/home/vtst/1bss/runtime/default/sqlite/sa_outage.db`
- quang chủ động đọc snapshot `/home/vtst/do_chu_dong_api/runtime/current_off_snapshot.json`
- DB quang chủ động đang hard-code STY/SHI
- một số nguồn SHC, vật tư, Excel legacy phụ thuộc đường dẫn chung

Điều này không ngăn app chạy nhiều process song song, nhưng có thể khiến một số page hiển thị dữ liệu chung hoặc dữ liệu không đúng đơn vị nếu mở cho tất cả instance.

### Khóa/mở route theo từng instance

#### Tra cứu danh sách route

Route policy dùng Flask endpoint name, không dùng URL path. List toàn bộ route bằng lệnh:

```bash
cd /home/vtst/dashv4

python3 -c "import dashboard; rows=[]; [rows.append('{:<45s} {:<10s} {}'.format(r.endpoint, ','.join(sorted(r.methods - {'HEAD','OPTIONS'})), r.rule)) for r in dashboard.app.url_map.iter_rules()]; print('\n'.join(sorted(rows)))"
```

Output có 3 cột:

```text
endpoint_name                                  method     url_path
```

Ví dụ:

```text
quangchudong.page_quangchudong                GET        /quangchudong
sa_outage.page_su_co_sa                       GET        /su_co_sa
operations.page_pttb                          GET        /pttb
operations.page_brcd                          GET        /brcd
```

Tên đưa vào `disabled_endpoints`, `enabled_endpoints`, `DASHV4_DISABLED_ENDPOINTS`, `DASHV4_ENABLED_ENDPOINTS` là cột `endpoint_name`.

Nếu chỉ muốn xem các page route:

```bash
python3 -c "import dashboard; print('\n'.join(sorted(f'{r.endpoint} {r.rule}' for r in dashboard.app.url_map.iter_rules() if '.page_' in r.endpoint)))"
```

Nếu chỉ muốn tìm route theo URL hoặc tên:

```bash
python3 -c "import dashboard; print('\n'.join(sorted(f'{r.endpoint} {r.rule}' for r in dashboard.app.url_map.iter_rules() if 'quangchudong' in r.endpoint or 'quangchudong' in r.rule)))"
```

#### Cách bền vững: sửa `deploy/units.yaml`

Với các route chưa chắc chắn đúng dữ liệu theo đơn vị, cấu hình trong `deploy/units.yaml`:

```yaml
- code: hoai_duc
  slug: hoai-duc
  name: "TTVT Hoài Đức"
  port: 5019
  hostname: hoai-duc.dashboard.example.vn
  db_path: /home/vtst/baocaohanoi/api_transition/runtime/hoai_duc/sqlite_history/report_history.db
  disabled_endpoints:
    - quangchudong.page_quangchudong
    - sa_outage.page_su_co_sa
    - operations.page_pttb
    - operations.page_brcd
```

Sau khi chạy `scripts/generate_instances.py`, env sinh ra sẽ có:

```text
DASHV4_DISABLED_ENDPOINTS=quangchudong.page_quangchudong,sa_outage.page_su_co_sa,operations.page_pttb,operations.page_brcd
```

Sinh lại env và copy env cho instance liên quan:

```bash
cd /home/vtst/dashv4

python3 scripts/generate_instances.py \
  --units-file deploy/units.yaml \
  --output-dir deploy/generated

sudo cp deploy/generated/env/hoai_duc.env /etc/dashv4/hoai_duc.env
sudo systemctl restart dashv4@hoai_duc
```

Nếu một route đang bị khóa global trong `config.py` nhưng một instance đã có dữ liệu đúng, mở riêng bằng:

```yaml
enabled_endpoints:
  - quality.page_shc_processing
```

#### Cách sửa nhanh: sửa trực tiếp file env

Có thể sửa trực tiếp `/etc/dashv4/<unit>.env` khi cần thao tác nhanh trên server:

```bash
sudo editor /etc/dashv4/hoai_duc.env
```

Thêm hoặc sửa dòng:

```env
DASHV4_DISABLED_ENDPOINTS=quangchudong.page_quangchudong,sa_outage.page_su_co_sa,operations.page_pttb,operations.page_brcd
```

Mở riêng route global-disabled:

```env
DASHV4_ENABLED_ENDPOINTS=quality.page_shc_processing
```

Sau khi sửa file env, phải restart đúng instance vì app chỉ đọc env khi process khởi động:

```bash
sudo systemctl restart dashv4@hoai_duc
```

Kiểm tra env đang áp dụng trên file runtime:

```bash
grep DASHV4_DISABLED_ENDPOINTS /etc/dashv4/hoai_duc.env
grep DASHV4_ENABLED_ENDPOINTS /etc/dashv4/hoai_duc.env
```

Lưu ý: sửa trực tiếp `/etc/dashv4/*.env` có thể bị mất nếu sau này regenerate từ `deploy/units.yaml` rồi copy đè env mới. Thay đổi cần lưu lâu dài nên ghi vào `deploy/units.yaml`.

Quy tắc vận hành:

- khóa route đọc nguồn chung cho mọi instance chưa được kiểm chứng
- chỉ mở route khi route đó đọc `DASHV4_DB_PATH` hoặc nguồn per-unit tương ứng
- sau mỗi thay đổi policy, regenerate env và restart service instance liên quan

Nguyên tắc vận hành:

- page nào đã đọc `report_history.db` theo `DASHV4_DB_PATH` thì phù hợp multi-instance hơn
- page nào đọc nguồn hard-code cần được kiểm tra riêng trước khi công bố cho toàn bộ đơn vị
- nếu chưa có nguồn per-unit, nên ẩn hoặc ghi rõ phạm vi dữ liệu

### Script dev không phải cơ chế production

`start_dashboard.sh` vẫn là script chạy nhanh một instance, không thay thế systemd template.

Production nên dùng:

```text
/etc/dashv4/<unit>.env
dashv4@<unit>.service
gunicorn_config.py
cloudflared ingress
```

## Lệnh kiểm tra kỹ thuật

Chạy test riêng cho multi-instance:

```bash
pytest tests/test_multi_instance_config.py tests/test_generate_instances.py
```

Chạy toàn bộ test suite:

```bash
pytest
```

Sinh lại artifact kiểm tra:

```bash
python3 scripts/generate_instances.py \
  --units-file deploy/units.yaml \
  --output-dir /tmp/dashv4-generated-check \
  --tunnel-id CHECK-TUNNEL \
  --credentials-file /tmp/check-tunnel.json
```

Kiểm tra service:

```bash
systemctl status dashv4@son_tay --no-pager
journalctl -u dashv4@son_tay -n 100 --no-pager
```

Kiểm tra port:

```bash
curl -fsS http://127.0.0.1:5011/ >/dev/null
curl -fsS http://127.0.0.1:5012/ >/dev/null
```
