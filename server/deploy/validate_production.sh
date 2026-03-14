#!/usr/bin/env bash
set -euo pipefail

DOMAIN="${1:-pdf-input-studio.kanazawa-application-support.jp}"

echo "[1] service status"
sudo systemctl is-active inputstudio
sudo systemctl is-active nginx

echo "[2] nginx config test"
sudo nginx -t >/dev/null
echo "nginx config: ok"

echo "[3] redirect check (http -> https)"
curl -sSI "http://${DOMAIN}/" | sed -n '1,5p'

echo "[4] https reachability"
curl -sSI "https://${DOMAIN}/" | sed -n '1,5p'

echo "[5] ad config"
curl -s "https://${DOMAIN}/ad-config.js" | sed -n '1,3p'

echo "[6] ad env vars"
systemctl show inputstudio -p Environment | tr ' ' '\n' | grep '^INPUTSTUDIO_AD' || true

echo "[7] upload limit in nginx"
grep -n "client_max_body_size" /etc/nginx/conf.d/inputstudio.conf || true

echo "[8] recent app logs"
sudo journalctl -u inputstudio -n 60 --no-pager
