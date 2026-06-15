# CLAUDE.md — project + deployment notes

Maxx Health Commission Statement Generator. Flask app that turns a monthly
commission worksheet into a review workbook (Step 1, openpyxl) and then a PDF
zip bundle (Step 2, LibreOffice). Live at https://commissions.phillyshah.com.

Read this before touching deployment — these are hard-won, non-obvious facts
about the production environment.

## Production environment (Hostinger Ubuntu 24.04 VPS)

- **Host:** `srv1373951`, IPv4 `72.62.174.193`. App dir: `/opt/commission-app`.
- **Service:** systemd `commission-app.service` runs gunicorn from the repo's
  `./venv`. Logs in `/var/log/commission-app/` and `journalctl -u commission-app`.
- **Reverse proxy is Traefik, NOT nginx.** Traefik runs as a Docker container in
  `/opt/traefik` and owns ports 80/443 (Let's Encrypt via ACME). Route is a file
  provider at `/opt/traefik/dynamic/commissions-phillyshah-com.yml` →
  `http://host.docker.internal:5002`. Reference copy: `deploy/traefik-commissions.yml`.
  - Do **not** install nginx/certbot. `deploy.sh` and `nginx-commissions.conf`
    are LEGACY and must not be run (deploy.sh exits early on purpose).

## Ports — three apps share this VPS (do not collide)

| Port | App | Domain |
|------|-----|--------|
| 5000 | tmcheck | tm.phillyshah.com |
| 5001 | MO Commission | mo-commissions.90ten.life |
| 5002 | **THIS app** | **commissions.phillyshah.com** |

This app must stay on **5002** (`gunicorn.conf.py` binds `0.0.0.0:5002`; bind on
`0.0.0.0` so the Traefik container can reach it via `host.docker.internal`). If
two apps share a port, only one binds and the other crash-loops — that is how
`commissions.phillyshah.com` once served the wrong site.

## Deployment — self-hosted runner (NOT SSH)

- Push to `main` → `.github/workflows/deploy.yml` runs on a **self-hosted runner
  on the VPS** (labels `self-hosted`, `commission-app`).
- **GitHub's cloud runners cannot SSH to this VPS** — port 22 is firewalled from
  GitHub's IP ranges (`dial tcp 72.62.174.193:22: i/o timeout`). The old
  `appleboy/ssh-action` workflow failed on every push. Do not bring it back.
- Deploy does `git fetch && git reset --hard origin/main`, then
  `pip install -r requirements.txt`, then `systemctl restart commission-app`.

## ⚠️ Never hand-edit files on the server

`/opt/commission-app` mirrors `origin/main` and the deploy **resets hard** to it,
so any local edit on the box is discarded. A past hand-edit to `gunicorn.conf.py`
also caused `git pull` to abort ("local changes would be overwritten"). Make all
changes in git and let them deploy.

## PDF generation depends on LibreOffice

Step 2 shells out to `soffice` (`libreoffice-calc`). It must be installed on the
host. Note: `soffice` exits 0 even on failure, so `generate_pdfs()` verifies each
expected PDF was actually written and raises a clear error if conversion fails
rather than returning an empty zip.

## Verify a deploy

```bash
sudo systemctl status commission-app --no-pager      # active (running)
curl -s http://127.0.0.1:5002/ | grep -o 'v[0-9.]*'  # footer version
which soffice                                         # PDF dependency present
```
