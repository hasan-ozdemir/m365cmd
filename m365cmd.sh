#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MAIN="$SCRIPT_DIR/m365cmd-main.ps1"

if ! command -v pwsh >/dev/null 2>&1; then
  echo "ERROR: PowerShell Core is required." >&2
  exit 1
fi

exec pwsh -NoProfile -File "$MAIN" "$@"
