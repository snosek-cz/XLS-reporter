#!/bin/bash
# setup_macos.sh
# Run this ONCE on your Mac to install the XLS Reporter automation.
# Usage: bash setup_macos.sh

set -e

echo "======================================"
echo "  XLS Reporter — macOS Setup"
echo "======================================"
echo

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
INSTALL_DIR="$HOME/Library/Scripts/XLS-Reporter"
VENV_DIR="$INSTALL_DIR/venv"
INBOX_DIR="$HOME/XLS-Reports/Inbox"
DONE_DIR="$HOME/XLS-Reports/Done"
ACTIONS_DIR="$HOME/Library/Scripts/Folder Action Scripts"

# ── Step 1: Create folders ────────────────────────────────────────────────────
echo "[1/5] Creating folders..."
mkdir -p "$INSTALL_DIR"
mkdir -p "$INBOX_DIR"
mkdir -p "$DONE_DIR"
mkdir -p "$ACTIONS_DIR"
echo "  ✅ ~/XLS-Reports/Inbox    ← drop your Details files here"
echo "  ✅ ~/XLS-Reports/Done     ← processed files land here"
echo

# ── Step 2: Install scripts ───────────────────────────────────────────────────
echo "[2/5] Installing scripts..."
cp "$SCRIPT_DIR/../generate_overview.py"    "$INSTALL_DIR/"
cp "$SCRIPT_DIR/../consolidate_annual.py"   "$INSTALL_DIR/"
cp "$SCRIPT_DIR/folder_action_trigger.sh"   "$INSTALL_DIR/"
cp "$SCRIPT_DIR/quick_action_run.sh"        "$INSTALL_DIR/"
chmod +x "$INSTALL_DIR/folder_action_trigger.sh"
chmod +x "$INSTALL_DIR/quick_action_run.sh"
echo "  ✅ Scripts installed to: $INSTALL_DIR"
echo

# ── Step 3: Python venv + openpyxl ───────────────────────────────────────────
echo "[3/5] Setting up Python environment..."
if ! command -v python3 &>/dev/null; then
    echo "  ❌ python3 not found. Install from https://www.python.org or: brew install python"
    exit 1
fi
PYTHON_VERSION=$(python3 --version)
echo "  ✅ $PYTHON_VERSION found"

# Create isolated venv inside install dir
echo "  Creating virtual environment..."
python3 -m venv "$VENV_DIR"
"$VENV_DIR/bin/pip" install --quiet --upgrade pip
"$VENV_DIR/bin/pip" install --quiet openpyxl
echo "  ✅ openpyxl installed in isolated venv: $VENV_DIR"
echo

# ── Step 4: Install Folder Action ────────────────────────────────────────────
echo "[4/5] Installing macOS Folder Action..."

FA_SCRIPT="$ACTIONS_DIR/XLS-Reporter-Folder-Action.scpt"
PYTHON_BIN="$VENV_DIR/bin/python3"

# Write AppleScript that calls our shell trigger
cat > "/tmp/xls_fa.applescript" << APPLESCRIPT
on adding folder items to thisFolder after receiving theItems
    set installDir to "$INSTALL_DIR/"
    set triggerScript to installDir & "folder_action_trigger.sh"
    repeat with theItem in theItems
        set itemPath to POSIX path of theItem
        do shell script "bash " & quoted form of triggerScript & " " & quoted form of itemPath
    end repeat
end adding folder items to
APPLESCRIPT

osacompile -o "$FA_SCRIPT" "/tmp/xls_fa.applescript" 2>/dev/null && \
    echo "  ✅ Folder Action script compiled" || \
    echo "  ⚠️  Could not compile AppleScript (Folder Action will need manual setup)"
rm -f /tmp/xls_fa.applescript

# Attach Folder Action to Inbox
osascript << APPLESCRIPT2 2>/dev/null
tell application "System Events"
    try
        set inboxFolder to POSIX file "$INBOX_DIR" as alias
        make new folder action at inboxFolder with properties {name:"XLS-Reporter-Folder-Action", script name:"XLS-Reporter-Folder-Action"}
    end try
end tell
APPLESCRIPT2
echo "  ✅ Folder Action attached to: $INBOX_DIR"
echo

# Write the venv python path into folder_action_trigger.sh and quick_action_run.sh
# so they use the correct Python automatically
sed -i '' "s|python3 |\"$VENV_DIR/bin/python3\" |g" "$INSTALL_DIR/folder_action_trigger.sh"
sed -i '' "s|python3 |\"$VENV_DIR/bin/python3\" |g" "$INSTALL_DIR/quick_action_run.sh"

# ── Step 5: Quick Action ──────────────────────────────────────────────────────
echo "[5/5] Installing Quick Action (right-click)..."

QA_DIR="$HOME/Library/Services/Generate XLS Overview.workflow"
mkdir -p "$QA_DIR/Contents"

cat > "$QA_DIR/Contents/document.wflow" << WORKFLOW
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>AMApplicationBuild</key><string>523</string>
    <key>AMApplicationVersion</key><string>2.10</string>
    <key>AMDocumentVersion</key><string>2</string>
    <key>actions</key>
    <array>
        <dict>
            <key>action</key>
            <dict>
                <key>AMAccepts</key>
                <dict><key>Container</key><string>List</string><key>Optional</key><true/><key>Types</key><array><string>com.apple.cocoa.path</string></array></dict>
                <key>AMActionVersion</key><string>2.0.3</string>
                <key>AMApplication</key><array><string>Finder</string></array>
                <key>AMParameterProperties</key><dict><key>COMMAND_STRING</key><dict/><key>inputMethod</key><dict/><key>shell</key><dict/><key>source</key><dict/></dict>
                <key>AMProvides</key><dict><key>Container</key><string>List</string><key>Types</key><array><string>com.apple.cocoa.path</string></array></dict>
                <key>ActionBundlePath</key><string>/System/Library/Automator/Run Shell Script.action</string>
                <key>ActionName</key><string>Run Shell Script</string>
                <key>ActionParameters</key>
                <dict>
                    <key>COMMAND_STRING</key>
                    <string>for f in "$@"; do
  "$HOME/Library/Scripts/XLS-Reporter/venv/bin/python3" "$HOME/Library/Scripts/XLS-Reporter/generate_overview.py" "$f" --notify
done</string>
                    <key>inputMethod</key><integer>1</integer>
                    <key>shell</key><string>/bin/bash</string>
                    <key>source</key><string></string>
                </dict>
                <key>BundleIdentifier</key><string>com.apple.RunShellScript</string>
                <key>CFBundleVersion</key><string>2.0.3</string>
                <key>CanShowSelectedItemsWhenRun</key><false/>
                <key>CanShowWhenRun</key><true/>
                <key>Category</key><array><string>AMCategoryUtilities</string></array>
                <key>Class Name</key><string>RunShellScriptAction</string>
                <key>InputUUID</key><string>F6A6E5A1-1B2C-3D4E-5F6A-7B8C9D0E1F2A</string>
                <key>Keywords</key><array><string>Shell</string><string>Script</string></array>
                <key>OutputUUID</key><string>A1B2C3D4-E5F6-7890-ABCD-EF1234567890</string>
                <key>UUID</key><string>12345678-1234-1234-1234-123456789012</string>
                <key>UnlockTimeout</key><integer>0</integer>
                <key>arguments</key><dict/>
                <key>isViewVisible</key><integer>1</integer>
                <key>location</key><string>309.500000:253.000000</string>
                <key>nickname</key><string>Run Shell Script</string>
            </dict>
        </dict>
    </array>
    <key>connectors</key><dict/>
    <key>workflowMetaData</key>
    <dict>
        <key>serviceApplicationBundleID</key><string>com.apple.finder</string>
        <key>serviceApplicationPath</key><string>/System/Library/CoreServices/Finder.app</string>
        <key>serviceInputTypeIdentifier</key><string>com.apple.Automator.fileSystemObject.folder</string>
        <key>serviceOutputTypeIdentifier</key><string>com.apple.Automator.nothing</string>
        <key>serviceProcessesInput</key><integer>0</integer>
        <key>workflowTypeIdentifier</key><string>com.apple.Automator.servicesMenu</string>
    </dict>
</dict>
</plist>
WORKFLOW

cat > "$QA_DIR/Contents/Info.plist" << 'INFOPLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>NSServices</key>
    <array>
        <dict>
            <key>NSMenuItem</key><dict><key>default</key><string>Generate XLS Overview</string></dict>
            <key>NSMessage</key><string>runWorkflowAsService</string>
            <key>NSRequiredContext</key><dict><key>NSApplicationIdentifier</key><string>com.apple.finder</string></dict>
            <key>NSSendFileTypes</key><array><string>org.openxmlformats.spreadsheetml.sheet</string></array>
        </dict>
    </array>
</dict>
</plist>
INFOPLIST

echo "  ✅ Quick Action installed"
echo

# ── Done ──────────────────────────────────────────────────────────────────────
echo "======================================"
echo "  Setup Complete! 🎉"
echo "======================================"
echo
echo "HOW TO USE:"
echo
echo "  📂 AUTOMATIC (Folder Action):"
echo "     Drop any Details_*.xlsx into:"
echo "     ~/XLS-Reports/Inbox/"
echo "     → Overview generated, file moves to ~/XLS-Reports/Done/"
echo "     → Annual_YEAR.xlsx updated automatically"
echo
echo "  🖱️  MANUAL (Quick Action):"
echo "     Right-click any Details_*.xlsx in Finder"
echo "     → Quick Actions → Generate XLS Overview"
echo
echo "  📋 NOTE: You may need to enable Folder Actions:"
echo "     Right-click Desktop → Services → Folder Actions Setup"
echo "     Make sure 'Enable Folder Actions' is checked."
echo
