#!/bin/bash
# SessionStart hook — provisions a Claude Code on the web session so the
# commission app can be run, tested, and its PDFs rendered/inspected.
#
# Why these packages:
#   - libreoffice-calc : the app converts xlsx -> PDF via `soffice` (Step 2).
#     The base web image ships libreoffice-core but NOT the Calc module
#     (missing share/registry/calc.xcd), so `soffice --convert-to pdf` fails
#     with "source file could not be loaded" for every file. Installing
#     libreoffice-calc registers the Calc module and its filters.
#   - poppler-utils    : provides `pdftoppm`, which lets PDFs be rasterized so
#     generated statements can be visually verified.
#   - Python deps      : installed into ./venv so `import app`, openpyxl, etc.
#     and the test suite work.
set -euo pipefail

# Local machines already have their own environment; only provision the web sandbox.
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

# Run in the background so the session starts immediately. The install still
# runs, but doesn't block startup; the idempotency checks below make it a no-op
# once the container has the tools cached.
echo '{"async": true, "asyncTimeout": 300000}'

SUDO=""
[ "$(id -u)" -ne 0 ] && SUDO="sudo"

# --- System packages for PDF generation + inspection (idempotent) ---
if ! command -v soffice >/dev/null 2>&1 \
   || [ ! -f /usr/lib/libreoffice/share/registry/calc.xcd ] \
   || ! command -v pdftoppm >/dev/null 2>&1; then
  # The base image's apt index can be stale (exact point-release .deb pruned),
  # so refresh first. Tolerate unrelated third-party PPAs that 403 via the proxy.
  $SUDO apt-get update -y || true
  $SUDO DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
    libreoffice-calc poppler-utils
fi

# --- Python dependencies in a venv ---
if [ ! -d "$CLAUDE_PROJECT_DIR/venv" ]; then
  python3 -m venv "$CLAUDE_PROJECT_DIR/venv"
fi
"$CLAUDE_PROJECT_DIR/venv/bin/pip" install -q --upgrade pip
"$CLAUDE_PROJECT_DIR/venv/bin/pip" install -q -r "$CLAUDE_PROJECT_DIR/requirements.txt"

# Make the venv the default python/pip for the session.
echo "export PATH=\"$CLAUDE_PROJECT_DIR/venv/bin:\$PATH\"" >> "$CLAUDE_ENV_FILE"

echo "session-start: libreoffice-calc + poppler-utils + python deps ready."
