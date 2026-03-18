#!/bin/bash
# quick_action_run.sh
# Called by macOS Quick Action (Finder right-click menu)
# Receives selected file(s) as arguments.

INSTALL_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_SCRIPT="$INSTALL_DIR/generate_overview.py"

for filepath in "$@"; do
    python3 "$PYTHON_SCRIPT" "$filepath" --notify
done
