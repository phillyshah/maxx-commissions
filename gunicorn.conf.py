# Port 5001 (NOT 5000): the production VPS also runs the unrelated "tmcheck"
# app (tm.phillyshah.com) on 0.0.0.0:5000. Two services on the same port means
# whichever binds first wins and the other crash-loops — historically this app
# lost, so commissions.phillyshah.com silently served the tmcheck site. Keep
# this app on its own port and route Traefik to it accordingly.
#
# Bind on 0.0.0.0 (not 127.0.0.1) so the Traefik container can reach it via
# host.docker.internal (host-gateway).
bind = "0.0.0.0:5001"
workers = 2
timeout = 300  # 5 min for large workbooks + PDF generation
accesslog = "/var/log/commission-app/access.log"
errorlog = "/var/log/commission-app/error.log"
