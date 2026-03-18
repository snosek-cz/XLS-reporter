#!/bin/bash
# folder_action_trigger.sh
# Called by macOS Folder Action when new files land in the Inbox folder.

INSTALL_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_SCRIPT="$INSTALL_DIR/generate_overview.py"
CONSOLIDATE_SCRIPT="$INSTALL_DIR/consolidate_annual.py"
LOG_FILE="$HOME/Library/Logs/XLS-Reporter.log"

# Load config (sets XLS_DONE_DIR)
source "$INSTALL_DIR/config.env" 2>/dev/null || XLS_DONE_DIR="$HOME/Documents/XLS-Done"

mkdir -p "$XLS_DONE_DIR"
mkdir -p "$(dirname "$LOG_FILE")"

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" | tee -a "$LOG_FILE"
}

log "=== Folder Action triggered ==="
log "Install dir: $INSTALL_DIR"
log "Done dir: $XLS_DONE_DIR"
log "Arguments received: $#"
for arg in "$@"; do
    log "  File: $arg"
done

if [ $# -eq 0 ]; then
    log "WARNING: No files received as arguments"
    exit 0
fi

for filepath in "$@"; do
    filename=$(basename "$filepath")
    log "Processing: $filename"

    if [[ "$filename" == Details_*.xlsx ]]; then
        # Run overview generator
        log "Running generate_overview.py..."
        python3 "$PYTHON_SCRIPT" "$filepath" --notify >> "$LOG_FILE" 2>&1
        EXIT_CODE=$?
        log "generate_overview exit code: $EXIT_CODE"

        if [ $EXIT_CODE -eq 0 ]; then
            mv "$filepath" "$XLS_DONE_DIR/$filename"
            log "Moved to: $XLS_DONE_DIR/$filename"

            log "Running consolidate_annual.py..."
            python3 "$CONSOLIDATE_SCRIPT" "$XLS_DONE_DIR" --notify >> "$LOG_FILE" 2>&1
            log "consolidate_annual exit code: $?"
        else
            log "ERROR: generate_overview.py failed"
            osascript -e "display notification \"Failed: $filename\" with title \"XLS Reporter — Error\""
        fi
    else
        log "Skipping (not a Details_*.xlsx file): $filename"
    fi
done

log "=== Done ==="
