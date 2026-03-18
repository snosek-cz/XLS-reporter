#!/bin/bash
# folder_action_trigger.sh
# Called by macOS Folder Action when new files land in the Inbox folder.
# 1. Generates the overview sheet in the Details file
# 2. Moves the processed file to Done/
# 3. Updates the Annual_YEAR.xlsx consolidation file

INSTALL_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_SCRIPT="$INSTALL_DIR/generate_overview.py"
CONSOLIDATE_SCRIPT="$INSTALL_DIR/consolidate_annual.py"
DONE_DIR="$HOME/XLS-Reports/Done"

mkdir -p "$DONE_DIR"

for filepath in "$@"; do
    filename=$(basename "$filepath")
    if [[ "$filename" == Details_*.xlsx ]]; then
        # Step 1: Generate overview sheet
        python3 "$PYTHON_SCRIPT" "$filepath" --notify
        EXIT_CODE=$?

        if [ $EXIT_CODE -eq 0 ]; then
            # Step 2: Move to Done folder
            mv "$filepath" "$DONE_DIR/$filename"

            # Step 3: Update annual consolidation
            python3 "$CONSOLIDATE_SCRIPT" "$DONE_DIR" --notify
        else
            osascript -e "display notification \"Failed to process $filename\" with title \"XLS Reporter — Error\""
        fi
    fi
done
