#!/usr/bin/env bash
set -euo pipefail

PORT="${1:-8000}"
HOST="${2:-0.0.0.0}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

cd "$ROOT_DIR"

if command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="python"
else
  echo "Python is not installed or not on PATH."
  exit 1
fi

echo "Serving Kindle Notes Viewer from: $ROOT_DIR"
echo "Local URL:   http://localhost:$PORT"
echo "Network URL: http://<your-ip>:$PORT"
echo "Press Ctrl+C to stop."

exec "$PYTHON_BIN" -m http.server --bind "$HOST" "$PORT"
