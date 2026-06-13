#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT_DIR"

if [[ -f deploy/linux/.env ]]; then
  set -a
  # shellcheck disable=SC1091
  source deploy/linux/.env
  set +a
fi

export PDF_BACKEND="${PDF_BACKEND:-libreoffice}"
export WEBDAV_BASE_URL="${WEBDAV_BASE_URL:-http://localhost:${WEBDAV_PORT:-8080}}"

if ! command -v "${LIBREOFFICE_BIN:-soffice}" >/dev/null 2>&1; then
  echo "LibreOffice non trovato (${LIBREOFFICE_BIN:-soffice}). Installalo, ad es.: sudo apt install libreoffice"
  exit 1
fi

python3 -m venv .venv
# shellcheck disable=SC1091
source .venv/bin/activate
pip install -r requirements-linux.txt

python webdav_server.py &
WEBDAV_PID=$!
python app.py &
APP_PID=$!

cleanup() {
  kill "$WEBDAV_PID" "$APP_PID" 2>/dev/null || true
}
trap cleanup EXIT INT TERM

echo "Flask: http://localhost:${FLASK_PORT:-5000}/"
echo "WebDAV: ${WEBDAV_BASE_URL}/"
wait
