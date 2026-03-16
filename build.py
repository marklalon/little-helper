"""
Build script for Clipboard Image Paster
Creates executable using PyInstaller
"""

import os
import sys
import shutil
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(SCRIPT_DIR, "dist")
BUILD_DIR = os.path.join(SCRIPT_DIR, "build")


def clean():
    """Remove build artifacts."""
    for path in [DIST_DIR, BUILD_DIR]:
        if os.path.exists(path):
            shutil.rmtree(path)
            print(f"Removed: {path}")
    
    spec_file = os.path.join(SCRIPT_DIR, "ClipboardImagePaster.spec")
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"Removed: {spec_file}")


def build():
    """Build executable with PyInstaller."""
    icon_path = os.path.join(SCRIPT_DIR, "icon.ico")
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=ClipboardImagePaster",
        "--onefile",
        "--windowed",  # No console window
        f"--icon={icon_path}",
        f"--add-data={icon_path};.",  # Include icon.ico in the bundle
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=win32timezone",
        os.path.join(SCRIPT_DIR, "clipboard_image.pyw"),
    ]
    
    print("Running PyInstaller...")
    print(" ".join(cmd))
    subprocess.run(cmd, cwd=SCRIPT_DIR, check=True)
    
    exe_path = os.path.join(DIST_DIR, "ClipboardImagePaster.exe")
    if os.path.exists(exe_path):
        print(f"\nBuild successful: {exe_path}")
        return exe_path
    else:
        print("Build failed!")
        return None


def main():
    print("=" * 50)
    print("Clipboard Image Paster - Build Script")
    print("=" * 50)
    
    if len(sys.argv) > 1 and sys.argv[1] == "clean":
        clean()
        return
    
    clean()
    exe_path = build()
    
    if exe_path:
        print("\n" + "=" * 50)
        print("Build complete!")
        print(f"Executable: {exe_path}")
        print("\nNext step: Run Inno Setup on setup.iss to create installer")
        print("=" * 50)


if __name__ == "__main__":
    main()
