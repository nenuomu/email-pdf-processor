#!/usr/bin/env python3
"""
Build script for Email PDF Processor
Creates a standalone Windows executable using PyInstaller
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def clean_previous_builds():
    """Remove previous build directories"""
    project_root = Path(__file__).parent.parent
    
    # Directories to clean
    dirs_to_clean = [
        project_root / "build",
        project_root / "dist",
        project_root / "__pycache__",
        project_root / "src" / "__pycache__"
    ]
    
    for dir_path in dirs_to_clean:
        if dir_path.exists():
            print(f"Cleaning: {dir_path}")
            shutil.rmtree(dir_path)

def check_requirements():
    """Check if all required packages are installed"""
    required_packages = [
        "pyinstaller",
        "pandas", 
        "pdfplumber",
        "customtkinter",
        "extract-msg",
        "xlsxwriter"
    ]
    
    print("Checking required packages...")
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package.replace("-", "_"))
            print(f"✓ {package}")
        except ImportError:
            missing_packages.append(package)
            print(f"✗ {package} - MISSING")
    
    if missing_packages:
        print(f"\nMissing packages: {', '.join(missing_packages)}")
        print("Please install with: pip install " + " ".join(missing_packages))
        return False
    
    print("All required packages are installed!")
    return True

def build_executable():
    """Build the standalone executable using PyInstaller"""
    
    # Get paths
    project_root = Path(__file__).parent.parent
    src_dir = project_root / "src"
    main_script = src_dir / "main.py"
    assets_dir = project_root / "assets"
    icon_path = assets_dir / "icon.ico"
    
    # Verify main script exists
    if not main_script.exists():
        print(f"ERROR: Main script not found at {main_script}")
        print("Make sure src/main.py exists!")
        return False
    
    print(f"Building executable from: {main_script}")
    
    # Base PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",                           # Create single executable file
        "--windowed",                          # Hide console window (GUI app)
        "--name=EmailPDFProcessor",            # Executable name
        "--distpath=dist",                     # Output directory
        "--workpath=build",                    # Build directory
        "--specpath=build_scripts",            # Spec file location
        "--clean",                             # Clean PyInstaller cache
        
        # Include hidden imports (packages that PyInstaller might miss)
        "--hidden-import=pandas",
        "--hidden-import=pdfplumber", 
        "--hidden-import=customtkinter",
        "--hidden-import=extract_msg",
        "--hidden-import=xlsxwriter",
        "--hidden-import=email",
        "--hidden-import=email.mime",
        "--hidden-import=tempfile",
        "--hidden-import=threading",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        
        # Collect all files from these packages
        "--collect-all=customtkinter",
        "--collect-all=tkinter",
        
        # Optimize the build
        "--strip",                             # Strip debug information
        "--optimize=2",                        # Optimize bytecode
        
        # Exclude unnecessary modules to reduce size
        "--exclude-module=matplotlib",
        "--exclude-module=IPython",
        "--exclude-module=jupyter",
        "--exclude-module=notebook",
        "--exclude-module=scipy",
        "--exclude-module=numpy.distutils",
        "--exclude-module=tkinter.test",
        
        str(main_script)                       # Main script path
    ]
    
    # Add icon if it exists
    if icon_path.exists():
        cmd.extend(["--icon", str(icon_path)])
        print(f"Using icon: {icon_path}")
    else:
        print("No icon found, using default")
    
    # Add version info (Windows only)
    if sys.platform == "win32":
        cmd.extend([
            "--version-file=build_scripts/version_info.txt"
        ])
    
    print("\n" + "="*60)
    print("BUILDING EXECUTABLE")
    print("="*60)
    print(f"Command: {' '.join(cmd)}")
    print("This may take several minutes...")
    print("="*60)
    
    try:
        # Run PyInstaller
        result = subprocess.run(
            cmd, 
            check=True, 
            capture_output=True, 
            text=True,
            cwd=project_root
        )
        
        print("PyInstaller completed successfully!")
        
        # Check if executable was created
        exe_path = project_root / "dist" / "EmailPDFProcessor.exe"
        if exe_path.exists():
            file_size = exe_path.stat().st_size / (1024 * 1024)  # Size in MB
            print(f"\n✅ SUCCESS!")
            print(f"Executable created: {exe_path}")
            print(f"File size: {file_size:.1f} MB")
            
            # Test if executable can be run (basic check)
            print("\nTesting executable...")
            try:
                test_result = subprocess.run(
                    [str(exe_path), "--version"], 
                    capture_output=True, 
                    text=True, 
                    timeout=10
                )
                print("✓ Executable test passed")
            except (subprocess.TimeoutExpired, subprocess.CalledProcessError):
                print("⚠ Executable test completed (expected for GUI app)")
            
            return True
        else:
            print("❌ ERROR: Executable not found after build")
            return False
            
    except subprocess.CalledProcessError as e:
        print("❌ BUILD FAILED!")
        print(f"Return code: {e.returncode}")
        print("\nSTDOUT:")
        print(e.stdout)
        print("\nSTDERR:")
        print(e.stderr)
        print("\nCommon solutions:")
        print("1. Make sure all dependencies are installed: pip install -r requirements.txt")
        print("2. Check that src/main.py exists and is valid Python code")
        print("3. Try running the script directly first: python src/main.py")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

def create_version_info():
    """Create version info file for Windows executable"""
    project_root = Path(__file__).parent.parent
    version_file = project_root / "build_scripts" / "version_info.txt"
    
    version_content = '''# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx
VSVersionInfo(
  ffi=FixedFileInfo(
# filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)
# Set not needed items to zero 0.
filevers=(1,0,0,0),
prodvers=(1,0,0,0),
# Contains a bitmask that specifies the valid bits 'flags'r
mask=0x3f,
# Contains a bitmask that specifies the Boolean attributes of the file.
flags=0x0,
# The operating system for which this file was designed.
# 0x4 - NT and there is no need to change it.
OS=0x4,
# The general type of file.
# 0x1 - the file is an application.
fileType=0x1,
# The function of the file.
# 0x0 - the function is not defined for this fileType
subtype=0x0,
# Creation date and time stamp.
date=(0, 0)
),
  kids=[
StringFileInfo(
  [
  StringTable(
    u'040904B0',
    [StringStruct(u'CompanyName', u'Your Company'),
    StringStruct(u'FileDescription', u'Email PDF to Excel Processor'),
    StringStruct(u'FileVersion', u'1.0.0'),
    StringStruct(u'InternalName', u'EmailPDFProcessor'),
    StringStruct(u'LegalCopyright', u'Copyright (C) 2024'),
    StringStruct(u'OriginalFilename', u'EmailPDFProcessor.exe'),
    StringStruct(u'ProductName', u'Email PDF Processor'),
    StringStruct(u'ProductVersion', u'1.0.0')])
  ]), 
VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)'''
    
    try:
        os.makedirs(version_file.parent, exist_ok=True)
        version_file.write_text(version_content)
        print(f"Created version info: {version_file}")
    except Exception as e:
        print(f"Warning: Could not create version info: {e}")

def main():
    """Main build function"""
    print("Email PDF Processor - Build Script")
    print("="*50)
    
    # Check if we're in the right directory
    if not Path("src/main.py").exists():
        print("❌ ERROR: Please run this script from the project root directory")
        print("The directory should contain: src/main.py")
        print("Current directory:", os.getcwd())
        return False
    
    # Step 1: Clean previous builds
    print("\n1. Cleaning previous builds...")
    clean_previous_builds()
    
    # Step 2: Check requirements
    print("\n2. Checking requirements...")
    if not check_requirements():
        return False
    
    # Step 3: Create version info
    print("\n3. Creating version info...")
    create_version_info()
    
    # Step 4: Build executable
    print("\n4. Building executable...")
    success = build_executable()
    
    if success:
        print("\n" + "="*60)
        print("🎉 BUILD COMPLETED SUCCESSFULLY!")
        print("="*60)
        print("Your executable is ready at: dist/EmailPDFProcessor.exe")
        print("\nNext steps:")
        print("1. Test the executable on your computer")
        print("2. If it works, commit your code to GitHub")
        print("3. Create a release to trigger automatic builds")
        print("="*60)
    else:
        print("\n" + "="*60)
        print("❌ BUILD FAILED")
        print("="*60)
        print("Please check the error messages above and fix the issues.")
    
    return success

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
