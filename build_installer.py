#!/usr/bin/env python3
"""
TranscriptFlow Build Script

This script automates the build process for TranscriptFlow:
1. Cleans previous builds
2. Runs PyInstaller to create executable
3. Optionally compiles Inno Setup installer

Usage:
    python build_installer.py              # Build executable only
    python build_installer.py --installer  # Build executable and installer
    python build_installer.py --clean      # Clean build directories only

Note: This script preserves the 'installer_output' directory to keep previous versions.
      You should ignore/archive old versions manually as needed.
"""

import os
import sys
import shutil
import subprocess
import argparse
from pathlib import Path

# Colors for terminal output
class Colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def print_step(message):
    """Print a build step message"""
    print(f"\n{Colors.HEADER}{Colors.BOLD}==> {message}{Colors.ENDC}")

def print_success(message):
    """Print a success message"""
    print(f"{Colors.OKGREEN}[OK] {message}{Colors.ENDC}")

def print_error(message):
    """Print an error message"""
    print(f"{Colors.FAIL}[ERROR] {message}{Colors.ENDC}")

def print_warning(message):
    """Print a warning message"""
    print(f"{Colors.WARNING}[WARNING] {message}{Colors.ENDC}")

def clean_build():
    """Remove previous build artifacts"""
    print_step("Cleaning previous builds...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['TranscriptFlow.spec.bak']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"  Removing {dir_name}/")
            try:
                shutil.rmtree(dir_name)
            except Exception as e:
                print(f"  [Warning] Could not fully remove {dir_name}: {e}")
                print(f"  Continuing anyway...")
    
    for file_name in files_to_clean:
        if os.path.exists(file_name):
            print(f"  Removing {file_name}")
            os.remove(file_name)
    
    print_success("Build directories cleaned")

def check_dependencies():
    """Check if required tools are installed"""
    print_step("Checking dependencies...")
    
    # Check Python packages
    required_packages = {
        'PyInstaller': 'PyInstaller',
        'PyQt6': 'PyQt6',
        'enchant': 'enchant'
    }
    missing_packages = []
    
    for display_name, import_name in required_packages.items():
        try:
            __import__(import_name)
            print_success(f"{display_name} is installed")
        except ImportError:
            missing_packages.append(display_name)
            print_error(f"{display_name} is NOT installed")
    
    if missing_packages:
        print_error(f"Missing packages: {', '.join(missing_packages)}")
        print(f"\nInstall missing packages with:")
        print(f"  pip install {' '.join(missing_packages)}")
        return False
    
    return True

def build_executable():
    """Build the executable using PyInstaller"""
    print_step("Building executable with PyInstaller...")
    
    if not os.path.exists('TranscriptFlow.spec'):
        print_error("TranscriptFlow.spec not found!")
        return False
    
    try:
        # Run PyInstaller
        result = subprocess.run(
            [sys.executable, '-m', 'PyInstaller', 'TranscriptFlow.spec', '--clean'],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            print_success("Executable built successfully")
            print(f"  Output: dist/TranscriptFlow/")
            return True
        else:
            print_error("PyInstaller build failed")
            print(result.stderr)
            return False
            
    except FileNotFoundError:
        print_error("PyInstaller not found. Install it with: pip install pyinstaller")
        return False
    except Exception as e:
        print_error(f"Build failed: {e}")
        return False

def verify_build():
    """Verify the build output"""
    print_step("Verifying build...")
    
    exe_path = Path('dist/TranscriptFlow/TranscriptFlow.exe')
    
    if not exe_path.exists():
        print_error("TranscriptFlow.exe not found in dist/TranscriptFlow/")
        return False
    
    print_success(f"Executable found: {exe_path}")
    print(f"  Size: {exe_path.stat().st_size / (1024*1024):.2f} MB")
    
    # Check for required files - check both root and _internal (PyInstaller 6+)
    required_files = ['app_icon.ico', 'mpv-1.dll', 'dicts']
    
    all_present = True
    internal_dir = Path('dist/TranscriptFlow/_internal')
    root_dir = Path('dist/TranscriptFlow')
    
    for filename in required_files:
        path_internal = internal_dir / filename
        path_root = root_dir / filename
        
        if path_root.exists() or path_internal.exists():
            found_path = path_root if path_root.exists() else path_internal
            print_success(f"Found: {found_path}")
        else:
            print_warning(f"Missing: {filename} (checked root and _internal)")
            all_present = False
    
    return all_present

def build_installer():
    """Build the installer using Inno Setup"""
    print_step("Building installer with Inno Setup...")
    
    if not os.path.exists('TranscriptFlow.iss'):
        print_error("TranscriptFlow.iss not found!")
        return False
    
    # Common Inno Setup installation paths
    inno_paths = [
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
        r"C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
        r"C:\Program Files\Inno Setup 5\ISCC.exe",
    ]
    
    iscc_path = None
    for path in inno_paths:
        if os.path.exists(path):
            iscc_path = path
            break
    
    if not iscc_path:
        print_error("Inno Setup not found!")
        print("Download from: https://jrsoftware.org/isdl.php")
        return False
    
    try:
        # Create output directory
        os.makedirs('installer_output', exist_ok=True)
        
        # Run Inno Setup compiler
        result = subprocess.run(
            [iscc_path, 'TranscriptFlow.iss'],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            print_success("Installer built successfully")
            
            # Find the newest installer file
            installer_files = list(Path('installer_output').glob('*.exe'))
            if installer_files:
                # Sort by modification time, newest first
                installer_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                installer_path = installer_files[0]
                print(f"  Output: {installer_path}")
                print(f"  Size: {installer_path.stat().st_size / (1024*1024):.2f} MB")
            
            return True
        else:
            print_error("Inno Setup compilation failed")
            print(result.stderr)
            return False
            
    except Exception as e:
        print_error(f"Installer build failed: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Build TranscriptFlow installer')
    parser.add_argument('--clean', action='store_true', help='Clean build directories only')
    parser.add_argument('--installer', action='store_true', help='Build installer after executable')
    parser.add_argument('--skip-deps', action='store_true', help='Skip dependency check')
    
    args = parser.parse_args()
    
    print(f"{Colors.BOLD}{Colors.OKCYAN}")
    print("=" * 60)
    print("  TranscriptFlow Build Script")
    print("=" * 60)
    print(f"{Colors.ENDC}")
    
    # Clean if requested
    if args.clean:
        clean_build()
        return
    
    # Check dependencies
    if not args.skip_deps:
        if not check_dependencies():
            sys.exit(1)
    
    # Clean previous builds
    clean_build()
    
    # Build executable
    if not build_executable():
        print_error("\nBuild failed!")
        sys.exit(1)
    
    # Verify build
    if not verify_build():
        print_warning("\nBuild verification found issues")
    
    # Build installer if requested
    if args.installer:
        if not build_installer():
            print_error("\nInstaller build failed!")
            sys.exit(1)
    
    print(f"\n{Colors.OKGREEN}{Colors.BOLD}Build completed successfully!{Colors.ENDC}")
    
    if args.installer:
        print("\nNext steps:")
        print("  1. Test the installer in installer_output/")
        print("  2. Install on a clean Windows system")
        print("  3. Verify all features work correctly")
    else:
        print("\nNext steps:")
        print("  1. Test the executable in dist/TranscriptFlow/")
        print("  2. Run with --installer flag to create installer")

if __name__ == '__main__':
    main()
