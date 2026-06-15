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
- **Gunicorn** — production WSGI server (binds `0.0.0.0:5001`)
- **Traefik v3** — reverse proxy + automatic Let's Encrypt SSL

## Deployment (Hostinger Ubuntu VPS)

> **This server uses Traefik, not nginx.** Do **not** run `nginx`/`certbot` here
> — Traefik (running as a Docker container in `/opt/traefik`) already owns ports
> 80/443 and issues certs via ACME. Installing nginx will fail to bind port 80
> and conflict with Traefik. The legacy `deploy.sh` / `nginx-commissions.conf`
> in this repo are kept only for reference and are **not** used in production.

The app runs as a systemd service (`commission-app.service`) behind Traefik:

```bash
# 1. Code lives in /opt/commission-app and runs gunicorn on 0.0.0.0:5001
#    (port 5001, NOT 5000 — see gunicorn.conf.py for why)

# 2. Traefik routes the domain via a file-provider config:
#    /opt/traefik/dynamic/commissions-phillyshah-com.yml
#      Host(`commissions.phillyshah.com`) → http://host.docker.internal:5001
#    See deploy/traefik-commissions.yml in this repo for the reference route.

# 3. Deploys happen automatically via .github/workflows/deploy.yml on push to
#    main (git pull + pip install + systemctl restart commission-app).
#    Requires repo secrets: SERVER_HOST, SERVER_USER, SSH_PRIVATE_KEY.

# Manual deploy / restart:
cd /opt/commission-app && git pull && \
  source venv/bin/activate && pip install -r requirements.txt && \
  sudo systemctl restart commission-app
```

### Port 5001 (important)

The same VPS also hosts an unrelated **tmcheck** app (`tm.phillyshah.com`) on
`0.0.0.0:5000`. If this app is also put on 5000, only one can bind the port and
the other crash-loops — which previously made `commissions.phillyshah.com` serve
the tmcheck site. This app must stay on **5001** and Traefik must route to 5001.

## Local Development

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
# Open http://localhost:5000  (Flask dev server; gunicorn uses 5001 in prod)
```

## File Structure

```
commission-app/
├── app.py                    # Flask app + all processing logic
├── requirements.txt
├── gunicorn.conf.py          # Production server config (binds 0.0.0.0:5001)
├── commission-app.service    # systemd service
├── .github/workflows/
│   └── deploy.yml            # Auto-deploy on push to main (SSH)
├── deploy/
│   └── traefik-commissions.yml  # Reference Traefik route (→ :5001)
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
