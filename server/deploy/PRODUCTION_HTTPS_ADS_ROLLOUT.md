# Production Rollout (ConoHa VPS): HTTPS + Live Ads

Target domain: `pdf-input-studio.kanazawa-application-support.jp`

This runbook enables:
- HTTPS (Let's Encrypt)
- HTTP -> HTTPS redirect
- AdSense live ads for all 4 slots

## 0) Preconditions

- DNS `A` record points to the ConoHa VPS global IP
- App source is at `/opt/inputstudio`
- Python virtualenv is at `/opt/inputstudio/.venv`
- Service name is `inputstudio`

## 1) Baseline Verification

```bash
sudo systemctl is-active inputstudio
sudo systemctl status inputstudio --no-pager -l
sudo nginx -t
sudo systemctl is-active nginx
grep -n "server_name" /etc/nginx/conf.d/inputstudio.conf
curl -I http://pdf-input-studio.kanazawa-application-support.jp/
```

Expected:
- `inputstudio` and `nginx` are `active`
- `server_name` contains `pdf-input-studio.kanazawa-application-support.jp`
- HTTP returns a normal response (before redirect setup)

## 2) Backup Existing Config

```bash
sudo cp -a /etc/nginx/conf.d/inputstudio.conf /etc/nginx/conf.d/inputstudio.conf.bak.$(date +%Y%m%d%H%M%S)
sudo cp -a /etc/systemd/system/inputstudio.service /etc/systemd/system/inputstudio.service.bak.$(date +%Y%m%d%H%M%S)
```

## 3) Apply systemd Service (with ad env vars)

```bash
sudo cp /opt/inputstudio/server/deploy/inputstudio.service.example /etc/systemd/system/inputstudio.service
sudo systemctl daemon-reload
sudo systemctl restart inputstudio
sudo systemctl is-active inputstudio
```

Validate env vars:

```bash
systemctl show inputstudio -p Environment | tr ' ' '\n' | grep '^INPUTSTUDIO_AD'
systemctl show inputstudio -p Environment | tr ' ' '\n' | grep '^INPUTSTUDIO_UNLOCK_AD_SECONDS='
systemctl show inputstudio -p Environment | tr ' ' '\n' | grep '^INPUTSTUDIO_MAX_UPLOAD_MB='
```

## 4) Apply Nginx Config and Enable HTTPS

Install certbot once:

```bash
sudo apt-get update
sudo apt-get install -y certbot python3-certbot-nginx
```

Deploy Nginx config template:

```bash
sudo mkdir -p /var/www/certbot
sudo cp /opt/inputstudio/server/deploy/nginx.inputstudio.conf.example /etc/nginx/conf.d/inputstudio.conf
sudo nginx -t
sudo systemctl reload nginx
```

Issue certificate:

```bash
sudo certbot --nginx -d pdf-input-studio.kanazawa-application-support.jp --agree-tos -m your-email@example.com --no-eff-email -n
```

Final reload:

```bash
sudo nginx -t
sudo systemctl reload nginx
```

## 5) Production Validation

```bash
curl -I http://pdf-input-studio.kanazawa-application-support.jp/
curl -I https://pdf-input-studio.kanazawa-application-support.jp/
curl -s https://pdf-input-studio.kanazawa-application-support.jp/ad-config.js | sed -n '1,3p'
grep -n "client_max_body_size" /etc/nginx/conf.d/inputstudio.conf
sudo journalctl -u inputstudio -n 100 --no-pager
```

One-shot validation script (optional):

```bash
cd /opt/inputstudio
chmod +x server/deploy/validate_production.sh
server/deploy/validate_production.sh
```

Expected:
- HTTP is `301` and redirects to HTTPS
- HTTPS is reachable
- `/ad-config.js` includes:
  - `enabled: true`
  - `client: "ca-pub-3765496771126537"`
  - slots for `gate`, `panel`, `panelBottom`, `unlock`
- `client_max_body_size 500M;` is present

Browser checks:
- `https://pdf-input-studio.kanazawa-application-support.jp/` opens normally
- Ad appears in:
  - Gate screen
  - Left panel (top and bottom)
  - Unlock modal (ZIP open / PDF append)

## 6) Fast Rollback

Disable ads immediately:

```bash
sudo sed -i 's/^Environment=INPUTSTUDIO_ADS_ENABLED=.*/Environment=INPUTSTUDIO_ADS_ENABLED=0/' /etc/systemd/system/inputstudio.service
sudo systemctl daemon-reload
sudo systemctl restart inputstudio
```

Restore previous configs:

```bash
sudo cp -a /etc/nginx/conf.d/inputstudio.conf.bak.* /etc/nginx/conf.d/inputstudio.conf
sudo cp -a /etc/systemd/system/inputstudio.service.bak.* /etc/systemd/system/inputstudio.service
sudo systemctl daemon-reload
sudo nginx -t && sudo systemctl reload nginx
sudo systemctl restart inputstudio
```

