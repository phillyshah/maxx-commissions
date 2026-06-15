# Port 5002 — this app needs its OWN port. The production VPS runs three
# gunicorn apps and the lower ports are already taken:
#   :5000 = tmcheck         (tm.phillyshah.com)
#   :5001 = MO Commission   (mo-commissions.90ten.life)
#   :5002 = THIS app, Maxx Health Commission Statements (commissions.phillyshah.com)
#
# History: tmcheck was deployed on 5000 — the port this app originally used —
# so this service crash-looped ("Address already in use") and
# commissions.phillyshah.com silently served the tmcheck site. Keeping this app
# on its own dedicated port (5002) and routing Traefik to 5002 fixes it.
#
# Bind on 0.0.0.0 (not 127.0.0.1) so the Traefik container can reach it via
# host.docker.internal (host-gateway).
bind = "0.0.0.0:5002"
workers = 2
timeout = 300  # 5 min for large workbooks + PDF generation
accesslog = "/var/log/commission-app/access.log"
errorlog = "/var/log/commission-app/error.log"
