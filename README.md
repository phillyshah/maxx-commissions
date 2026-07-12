# Maxx Health — Commission Statement Generator

Internal tool for generating distributor commission statements from invoice worksheets.
automatically pushed to commissions.phillyshah.com

## What It Does

1. Step 1: Upload a Commission Worksheet (`.xlsx`) containing:
   - **Invoice List** sheet — all invoices grouped by distributor code
   - **Dist Lookup** sheet — distributor code → name → contact mapping
   - **Trauma** sheet (optional) — trauma-specific invoices

2. The app automatically creates a review workbook:
   - Detects the sales month/year from the Invoice List
   - Creates a **Summary** sheet with all distributors and totals
   - Creates **individual distributor tabs** with Maxx branding, logo, formatted headers, footer
   - Outputs a downloadable **Excel workbook** for accounting review

3. Step 2: Upload the verified Excel workbook to:
   - Generate **individual PDFs** for each distributor (landscape, print-ready)
   - Package all PDFs into a downloadable **zip file**

4. Download the completed workbook and final PDF bundle.

## Tech Stack

- **Python 3 / Flask** — backend processing
- **openpyxl** — Excel file generation
- **LibreOffice Calc** (headless) — PDF conversion
- **Gunicorn** — production WSGI server (binds `0.0.0.0:5002`)
- **Traefik v3** — reverse proxy + automatic Let's Encrypt SSL

## Deployment (Hostinger Ubuntu VPS)

> **This server uses Traefik, not nginx.** Do **not** run `nginx`/`certbot` here
> — Traefik (running as a Docker container in `/opt/traefik`) already owns ports
> 80/443 and issues certs via ACME. Installing nginx will fail to bind port 80
> and conflict with Traefik. The legacy `deploy.sh` / `nginx-commissions.conf`
> in this repo are kept only for reference and are **not** used in production.

The app runs as a systemd service (`commission-app.service`) behind Traefik:

```text
# 1. Code lives in /opt/commission-app and runs gunicorn on 0.0.0.0:5002
#    (port 5002 — see gunicorn.conf.py and "Ports" below for why)

# 2. Traefik routes the domain via a file-provider config:
#    /opt/traefik/dynamic/commissions-phillyshah-com.yml
#      Host(`commissions.phillyshah.com`) → http://host.docker.internal:5002
#    See deploy/traefik-commissions.yml in this repo for the reference route.

# 3. PDF generation (Step 2) shells out to LibreOffice (`soffice`). It MUST be
#    installed on the host: `sudo apt-get install -y libreoffice-calc`.
```

### Auto-deploy: self-hosted GitHub Actions runner

Pushes to `main` auto-deploy via `.github/workflows/deploy.yml`, which runs on a
**self-hosted runner installed on the VPS** (labels: `self-hosted`,
`commission-app`).

> **Why self-hosted and not the usual SSH deploy?** GitHub's cloud runners
> **cannot reach this VPS on port 22** — the host firewall blocks inbound SSH
> from GitHub's runner IP ranges, so the old `appleboy/ssh-action` workflow
> failed on every push with `dial tcp 72.62.174.193:22: i/o timeout` and nothing
> deployed. A self-hosted runner connects **outbound** to GitHub, so no inbound
> firewall hole is needed. Do **not** revert to the SSH approach unless the
> firewall is opened to GitHub's IPs (not recommended).

The deploy job does `git fetch && git reset --hard origin/main` (not `git pull`)
so the box always mirrors `main` exactly — see the warning below about not
hand-editing files on the server.

**One-time runner install on the VPS** (as root). Get a registration `<TOKEN>`
from the repo: **Settings → Actions → Runners → New self-hosted runner**.

```bash
mkdir -p /opt/actions-runner && cd /opt/actions-runner
RUNNER_VERSION=$(curl -fsSL https://api.github.com/repos/actions/runner/releases/latest \
  | grep -oP '"tag_name": "v\K[^"]+')
curl -fsSL -o runner.tar.gz \
  "https://github.com/actions/runner/releases/download/v${RUNNER_VERSION}/actions-runner-linux-x64-${RUNNER_VERSION}.tar.gz"
tar xzf runner.tar.gz

export RUNNER_ALLOW_RUNASROOT=1   # the deploy needs systemctl; runner runs as root
./config.sh --url https://github.com/phillyshah/maxx-commissions \
  --token <TOKEN> --name commission-vps --labels commission-app \
  --unattended --replace

./svc.sh install      # install + run as a systemd service (survives reboot)
./svc.sh start
./svc.sh status       # should show "active (running)"
```

After this, the (now unused) `SERVER_HOST` / `SERVER_USER` / `SSH_PRIVATE_KEY`
repo secrets can be deleted.

**Manual deploy / restart** (if you ever need to bypass the runner):

```bash
cd /opt/commission-app && git fetch origin main && git reset --hard origin/main && \
  source venv/bin/activate && pip install -r requirements.txt && \
  sudo systemctl restart commission-app
```

### Ports (important)

The same VPS runs three separate gunicorn apps. Each must have its own port:

| Port  | App                              | Domain                     |
|-------|----------------------------------|----------------------------|
| 5000  | tmcheck (`tmcheck.service`)      | tm.phillyshah.com          |
| 5001  | MO Commission (`mo-commission-app.service`) | mo-commissions.90ten.life |
| 5002  | **This app** (`commission-app.service`) | **commissions.phillyshah.com** |

If two apps share a port, only one binds and the other crash-loops — which is
exactly how `commissions.phillyshah.com` ended up serving the tmcheck site (then
the MO site). This app must stay on **5002** and Traefik must route to 5002.

### ⚠️ Do not hand-edit files on the server

`/opt/commission-app` is a **deploy target that mirrors `origin/main`**. The
deploy resets hard to `main`, so any local edit on the box is **discarded** on
the next deploy — and before this was automated, a hand-edit to
`gunicorn.conf.py` (changing the port to 5002) caused `git pull` to abort with
*"Your local changes would be overwritten by merge."* Make every change in git
and let it deploy; never edit files directly on the VPS.

## Local Development

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
# Open http://localhost:5000  (Flask dev server; gunicorn binds 0.0.0.0:5002 in prod)
```

## File Structure

```
commission-app/
├── app.py                    # Flask app + all processing logic
├── requirements.txt
├── gunicorn.conf.py          # Production server config (binds 0.0.0.0:5002)
├── commission-app.service    # systemd service
├── .github/workflows/
│   └── deploy.yml            # Auto-deploy on push to main (self-hosted runner)
├── deploy/
│   └── traefik-commissions.yml  # Reference Traefik route (→ :5002)
├── nginx-commissions.conf    # LEGACY — not used (server runs Traefik)
├── deploy.sh                 # LEGACY — not used (installs nginx; do not run)
├── cleanup.sh                # Cron job for old file cleanup
├── static/
│   └── maxx_logo.png         # Maxx Health logo
├── templates/
│   └── index.html            # Web UI
├── uploads/                  # Temporary upload storage
└── outputs/                  # Generated files (auto-cleaned)
```
