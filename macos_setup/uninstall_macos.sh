#!/bin/bash
# uninstall_macos.sh
# Removes the XLS Reporter macOS automation installation.
# Does NOT remove ~/XLS-Reports/Done/ files or openpyxl.
#
# Usage: bash uninstall_macos.sh

echo "======================================"
echo "  XLS Reporter — macOS Uninstall"
echo "======================================"
echo

INSTALL_DIR="$HOME/Library/Scripts/XLS-Reporter"
ACTIONS_DIR="$HOME/Library/Scripts/Folder Action Scripts"
SERVICES_DIR="$HOME/Library/Services"
INBOX_DIR="$HOME/XLS-Reports/Inbox"

# ── Step 1: Detach Folder Action ──────────────────────────────────────────────
echo "[1/4] Detaching Folder Action from Inbox..."
osascript << APPLESCRIPT 2>/dev/null
tell application "System Events"
  try
    set inboxAlias to POSIX file "$INBOX_DIR" as alias
    set theActions to every folder action of inboxAlias
    repeat with fa in theActions
      if name of fa contains "XLS-Reporter" then
        delete fa
      end if
    end repeat
  end try
end tell
APPLESCRIPT
echo "  ✅ Folder Action detached"

# ── Step 2: Remove Folder Action script ───────────────────────────────────────
echo "[2/4] Removing Folder Action script..."
rm -f "$ACTIONS_DIR/XLS-Reporter-Folder-Action.scpt"
rm -f "$ACTIONS_DIR/XLS-Reporter-Folder-Action.scpt.scpt"
echo "  ✅ Folder Action script removed"

# ── Step 3: Remove Quick Action ───────────────────────────────────────────────
echo "[3/4] Removing Quick Action..."
rm -rf "$SERVICES_DIR/Generate XLS Overview.workflow"
echo "  ✅ Quick Action removed"

# ── Step 4: Remove installed scripts ──────────────────────────────────────────
echo "[4/4] Removing installed scripts..."
rm -rf "$INSTALL_DIR"
echo "  ✅ Scripts removed from: $INSTALL_DIR"

# ── Optional: Remove Inbox folder ─────────────────────────────────────────────
echo
if [ -d "$INBOX_DIR" ]; then
    read -p "Remove ~/XLS-Reports/Inbox/ folder? (Done/ folder kept) [y/N]: " answer
    if [[ "$answer" =~ ^[Yy]$ ]]; then
        rm -rf "$INBOX_DIR"
        echo "  ✅ Inbox folder removed"
    else
        echo "  ⏭️  Inbox folder kept"
    fi
fi

# ── Done ──────────────────────────────────────────────────────────────────────
echo
echo "======================================"
echo "  Uninstall Complete"
echo "======================================"
echo
echo "Note: The following were NOT removed:"
echo "  • ~/XLS-Reports/Done/   ← your processed files are safe"
echo "  • openpyxl Python package (may be used by other tools)"
echo "  • This repository/scripts folder"
echo
echo "To reinstall at any time: bash setup_macos.sh"
echo
