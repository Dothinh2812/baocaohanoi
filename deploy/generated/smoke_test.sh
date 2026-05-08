#!/usr/bin/env bash
set -euo pipefail

echo "Checking son_tay on port 5011"
curl -fsS http://127.0.0.1:5011/ >/dev/null

echo "Checking ba_dinh on port 5012"
curl -fsS http://127.0.0.1:5012/ >/dev/null

echo "Checking cau_giay on port 5013"
curl -fsS http://127.0.0.1:5013/ >/dev/null

echo "Checking dong_anh on port 5014"
curl -fsS http://127.0.0.1:5014/ >/dev/null

echo "Checking dong_da on port 5015"
curl -fsS http://127.0.0.1:5015/ >/dev/null

echo "Checking gia_lam on port 5016"
curl -fsS http://127.0.0.1:5016/ >/dev/null

echo "Checking giai_phong on port 5017"
curl -fsS http://127.0.0.1:5017/ >/dev/null

echo "Checking ha_dong on port 5018"
curl -fsS http://127.0.0.1:5018/ >/dev/null

echo "Checking hoai_duc on port 5019"
curl -fsS http://127.0.0.1:5019/ >/dev/null

echo "Checking hoan_kiem on port 5020"
curl -fsS http://127.0.0.1:5020/ >/dev/null

echo "Checking hoang_mai on port 5021"
curl -fsS http://127.0.0.1:5021/ >/dev/null

echo "Checking long_bien on port 5022"
curl -fsS http://127.0.0.1:5022/ >/dev/null

echo "Checking phu_xuyen on port 5023"
curl -fsS http://127.0.0.1:5023/ >/dev/null

echo "Checking soc_son on port 5024"
curl -fsS http://127.0.0.1:5024/ >/dev/null

echo "Checking tay_ho on port 5025"
curl -fsS http://127.0.0.1:5025/ >/dev/null

echo "Checking thach_that on port 5026"
curl -fsS http://127.0.0.1:5026/ >/dev/null

echo "Checking thanh_tri on port 5027"
curl -fsS http://127.0.0.1:5027/ >/dev/null

echo "Checking tu_liem on port 5028"
curl -fsS http://127.0.0.1:5028/ >/dev/null
