#!/usr/bin/env bash
set -e

PLUGIN_GUID="{B7A5C98E-892F-4E44-9B05-6D4B892FE01A}"
SRC_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

DEST_APP="/Applications/ONLYOFFICE.app/Contents/Resources/editors/sdkjs-plugins/$PLUGIN_GUID"
DEST_USER="$HOME/Library/Application Support/asc.onlyoffice.ONLYOFFICE/sdkjs-plugins/$PLUGIN_GUID"

echo "Installing Zotero QuickCite..."

# 1. Install to User Application Support
mkdir -p "$HOME/Library/Application Support/asc.onlyoffice.ONLYOFFICE/sdkjs-plugins"
rm -rf "$DEST_USER"
cp -R "$SRC_DIR" "$DEST_USER"
echo "✓ Installed to $DEST_USER"

# 2. Try install to ONLYOFFICE App Bundle (if writeable and permitted)
if [ -d "/Applications/ONLYOFFICE.app/Contents/Resources/editors/sdkjs-plugins" ]; then
    rm -rf "$DEST_APP" 2>/dev/null || true
    cp -R "$SRC_DIR" "$DEST_APP" 2>/dev/null && echo "✓ Installed to $DEST_APP" || echo "ℹ Note: User directory installation active (App bundle write-protected by macOS)"
fi

echo "Done! Restart ONLYOFFICE to see 'Zotero QuickCite' under the Plugins tab."
