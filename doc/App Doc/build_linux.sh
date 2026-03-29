#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$ROOT_DIR/.venv"
export MPLCONFIGDIR="$ROOT_DIR/.mplconfig"

cd "$ROOT_DIR"

if [[ ! -d "$VENV_DIR" ]]; then
  python3 -m venv "$VENV_DIR"
fi

"$VENV_DIR/bin/pip" install -r "$ROOT_DIR/requirements-linux.txt"
"$VENV_DIR/bin/pyinstaller" \
  --onefile \
  --windowed \
  --name PBIX_Analyzer_Linux \
  pbix_analyzer_gui.py

printf '\nBuild concluido em: %s\n' "$ROOT_DIR/dist/PBIX_Analyzer_Linux"
