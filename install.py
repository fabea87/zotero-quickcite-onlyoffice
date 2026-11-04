#!/usr/bin/env python3
"""
Zotero QuickCite - Universal Cross-Platform Installer for ONLYOFFICE
Works out of the box on macOS, Windows, and Linux (Native, Flatpak, Snap).

Usage:
    python3 install.py              # Install or update plugin
    python3 install.py --uninstall  # Remove plugin
    python3 install.py --package    # Generate .plugin release bundle
"""

import os
import sys
import shutil
import zipfile
import platform
import argparse

PLUGIN_GUID = "{B7A5C98E-892F-4E44-9B05-6D4B892FE01A}"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def get_possible_plugin_dirs():
    system = platform.system()
    home = os.path.expanduser("~")
    dirs = []

    if system == "Darwin":  # macOS
        dirs.extend([
            os.path.join(home, "Library", "Application Support", "asc.onlyoffice.ONLYOFFICE", "data", "sdkjs-plugins"),
            os.path.join(home, "Library", "Application Support", "asc.onlyoffice.ONLYOFFICE", "sdkjs-plugins"),
        ])
    elif system == "Windows":  # Windows
        appdata = os.environ.get("APPDATA", "")
        localappdata = os.environ.get("LOCALAPPDATA", "")
        if appdata:
            dirs.extend([
                os.path.join(appdata, "ONLYOFFICE", "DesktopEditors", "data", "sdkjs-plugins"),
                os.path.join(appdata, "ONLYOFFICE", "DesktopEditors", "sdkjs-plugins"),
            ])
        if localappdata:
            dirs.extend([
                os.path.join(localappdata, "ONLYOFFICE", "DesktopEditors", "data", "sdkjs-plugins"),
                os.path.join(localappdata, "ONLYOFFICE", "DesktopEditors", "sdkjs-plugins"),
            ])
    elif system == "Linux":  # Linux
        dirs.extend([
            # Flatpak
            os.path.join(home, ".var", "app", "org.onlyoffice.desktopeditors", "data", "onlyoffice", "desktopeditors", "data", "sdkjs-plugins"),
            os.path.join(home, ".var", "app", "org.onlyoffice.desktopeditors", "data", "onlyoffice", "desktopeditors", "sdkjs-plugins"),
            # Native / Deb / RPM
            os.path.join(home, ".local", "share", "onlyoffice", "desktopeditors", "data", "sdkjs-plugins"),
            os.path.join(home, ".local", "share", "onlyoffice", "desktopeditors", "sdkjs-plugins"),
            # Snap
            os.path.join(home, "snap", "onlyoffice-desktopeditors", "current", ".local", "share", "onlyoffice", "desktopeditors", "data", "sdkjs-plugins"),
        ])

    return dirs

def create_package():
    package_path = os.path.join(SCRIPT_DIR, "ZoteroQuickCite.plugin")
    if os.path.exists(package_path):
        os.remove(package_path)

    include_files = ["config.json", "index.html", "README.md"]
    include_dirs = ["resources", "scripts", "styles"]

    print("Creating ZoteroQuickCite.plugin package...")
    with zipfile.ZipFile(package_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in include_files:
            fp = os.path.join(SCRIPT_DIR, f)
            if os.path.exists(fp):
                z.write(fp, f)

        for d in include_dirs:
            dp = os.path.join(SCRIPT_DIR, d)
            if os.path.exists(dp):
                for root, _, files in os.walk(dp):
                    for file in files:
                        if file == ".DS_Store":
                            continue
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, SCRIPT_DIR)
                        z.write(full_path, rel_path)

    print(f"✓ Created: {package_path}")
    return package_path

def install_plugin():
    target_dirs = get_possible_plugin_dirs()
    installed_count = 0

    print("=" * 60)
    print("Zotero QuickCite - ONLYOFFICE Installer")
    print(f"Platform: {platform.system()} ({platform.machine()})")
    print("=" * 60)

    for base_dir in target_dirs:
        dest_dir = os.path.join(base_dir, PLUGIN_GUID)
        try:
            os.makedirs(base_dir, exist_ok=True)
            if os.path.exists(dest_dir):
                shutil.rmtree(dest_dir)

            # Copy plugin tree
            shutil.copytree(
                SCRIPT_DIR,
                dest_dir,
                ignore=shutil.ignore_patterns(".git*", "*.py", "*.sh", "*.plugin", ".DS_Store")
            )
            print(f"✓ Installed to: {dest_dir}")
            installed_count += 1
        except Exception as e:
            # Skip paths that cannot be accessed or created
            pass

    # Also build the .plugin package
    create_package()

    print("=" * 60)
    if installed_count > 0:
        print("Installation complete!")
        print("Please restart ONLYOFFICE. You will find 'Zotero QuickCite'")
        print("under the 'Plugins' tab on the ribbon.")
    else:
        print("Could not auto-detect ONLYOFFICE plugin directory.")
        print("Alternative 1-Click Install:")
        print("1. Open ONLYOFFICE -> Plugins -> Plugin Manager.")
        print("2. Click 'Add plugin' / 'Install plugin manually'.")
        print("3. Select the 'ZoteroQuickCite.plugin' file.")
    print("=" * 60)

def uninstall_plugin():
    target_dirs = get_possible_plugin_dirs()
    removed = 0
    for base_dir in target_dirs:
        dest_dir = os.path.join(base_dir, PLUGIN_GUID)
        if os.path.exists(dest_dir):
            try:
                shutil.rmtree(dest_dir)
                print(f"✓ Removed from: {dest_dir}")
                removed += 1
            except Exception as e:
                print(f"Error removing {dest_dir}: {e}")

    if removed > 0:
        print("Uninstallation complete. Restart ONLYOFFICE.")
    else:
        print("No installation found to remove.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Zotero QuickCite Installer for ONLYOFFICE")
    parser.add_argument("--uninstall", action="store_true", help="Remove plugin")
    parser.add_argument("--package", action="store_true", help="Only build the .plugin file")

    args = parser.parse_args()

    if args.uninstall:
        uninstall_plugin()
    elif args.package:
        create_package()
    else:
        install_plugin()
