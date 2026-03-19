"""
Build script for Little Helper
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
    
    spec_file = os.path.join(SCRIPT_DIR, "LittleHelper.spec")
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"Removed: {spec_file}")


def build():
    """Build executable with PyInstaller."""
    icon_path = os.path.join(SCRIPT_DIR, "res", "icon.ico")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=LittleHelper",
        "--onefile",
        "--windowed",  # No console window
        f"--icon={icon_path}",
        f"--add-data={icon_path};.",  # Include icon.ico in the bundle
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=win32timezone",
        "--hidden-import=psutil",
        "--hidden-import=pynvml",
        "--collect-all=wmi",        # wmi needs full collect for its extensions
        "--hidden-import=config",
        "--hidden-import=clipboard_paste",
        "--hidden-import=screenshot",
        "--hidden-import=hotkey",
        "--hidden-import=gpu_power",
        "--hidden-import=system_overlay",
        os.path.join(SCRIPT_DIR, "src", "main.pyw"),
    ]
    
    print("Running PyInstaller...")
    print(" ".join(cmd))
    subprocess.run(cmd, cwd=SCRIPT_DIR, check=True)
    
    exe_path = os.path.join(DIST_DIR, "LittleHelper.exe")
    if os.path.exists(exe_path):
        print(f"\nBuild successful: {exe_path}")
        return exe_path
    else:
        print("Build failed!")
        return None


def main():
    print("=" * 50)
    print("Little Helper - Build Script")
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
