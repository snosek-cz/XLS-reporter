#!/bin/bash
# folder_action_trigger.sh
# Called by macOS Folder Action when new files land in the Inbox folder.
# Receives file paths as arguments from Automator.
#
# Install location: ~/Library/Scripts/XLS-Reporter/

INSTALL_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_SCRIPT="$INSTALL_DIR/generate_overview.py"
DONE_DIR="$HOME/XLS-Reports/Done"

mkdir -p "$DONE_DIR"

for filepath in "$@"; do
    # Only process .xlsx files named Details_*
    filename=$(basename "$filepath")
    if [[ "$filename" == Details_*.xlsx ]] || [[ "$filename" == Details*.xlsx ]]; then
        # Run the Python overview generator
        python3 "$PYTHON_SCRIPT" "$filepath" --notify
        EXIT_CODE=$?

        if [ $EXIT_CODE -eq 0 ]; then
            # Move processed file to Done folder
            mv "$filepath" "$DONE_DIR/$filename"
        else
            # Notify error - file stays in Inbox
            osascript -e "display notification \"Failed to process $filename\" with title \"XLS Reporter — Error\""
        fi
    fi
done
