"""
Build script for Triple Duty Bond Generator EXE
Run this script to create the standalone executable.
"""

import subprocess
import sys
import os

def main():
    print("=" * 60)
    print("Building Triple Duty Bond Generator EXE")
    print("=" * 60)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print("✓ PyInstaller is installed")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed")
    
    # Build command
    # --onefile: Single EXE file
    # --windowed: No console window (GUI app)
    # --add-data: Include the template file
    # --name: Output EXE name
    # --icon: Optional icon file
    
    template_file = "Triple_Duty_Bond_template__IR53112.docx"
    
    if not os.path.exists(template_file):
        print(f"ERROR: Template file not found: {template_file}")
        print("Make sure the template is in the same directory as this script.")
        return
    
    print(f"\n✓ Template found: {template_file}")
    print("\nBuilding EXE... (this may take a few minutes)\n")
    
    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # Single EXE
        "--windowed",                   # No console window
        "--name", "Triple_Duty_Bond_Generator",  # EXE name
        f"--add-data", f"{template_file};.",     # Bundle template (Windows uses ;)
        "--clean",                      # Clean build
        "generate_bond.py"
    ]
    
    print(f"Running: {' '.join(cmd)}\n")
    
    result = subprocess.run(cmd)
    
    if result.returncode == 0:
        print("\n" + "=" * 60)
        print("✅ BUILD SUCCESSFUL!")
        print("=" * 60)
        print(f"\nEXE Location: dist\\Triple_Duty_Bond_Generator.exe")
        print("\nYou can now distribute the single EXE file.")
        print("The template is bundled inside - no extra files needed!")
    else:
        print("\n❌ Build failed. Check the errors above.")

if __name__ == "__main__":
    main()
