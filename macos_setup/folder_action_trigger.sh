#!/bin/bash
# folder_action_trigger.sh
# Called by macOS Folder Action when new files land in the Inbox folder.
# Reads DONE_DIR from config.env installed alongside this script.

INSTALL_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_SCRIPT="$INSTALL_DIR/generate_overview.py"
CONSOLIDATE_SCRIPT="$INSTALL_DIR/consolidate_annual.py"

# Load config (sets XLS_DONE_DIR)
# shellcheck source=/dev/null
source "$INSTALL_DIR/config.env" 2>/dev/null || XLS_DONE_DIR="$HOME/Documents/XLS-Done"

mkdir -p "$XLS_DONE_DIR"

for filepath in "$@"; do
    filename=$(basename "$filepath")
    if [[ "$filename" == Details_*.xlsx ]]; then
        python3 "$PYTHON_SCRIPT" "$filepath" --notify
        EXIT_CODE=$?
        if [ $EXIT_CODE -eq 0 ]; then
            mv "$filepath" "$XLS_DONE_DIR/$filename"
            python3 "$CONSOLIDATE_SCRIPT" "$XLS_DONE_DIR" --notify
        else
            osascript -e "display notification \"Failed: $filename\" with title \"XLS Reporter — Error\""
        fi
    fi
done
