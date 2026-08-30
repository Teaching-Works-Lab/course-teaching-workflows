#!/usr/bin/env sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
if command -v python3 >/dev/null 2>&1; then
  exec python3 "$SCRIPT_DIR/update_markdown_cache.py" "$@"
elif command -v python >/dev/null 2>&1; then
  exec python "$SCRIPT_DIR/update_markdown_cache.py" "$@"
else
  echo "Python 3 is required." >&2
  exit 1
fi
