#!/usr/bin/env python3
"""
Build comparison script for docpdf
Builds with both PyInstaller and Nuitka for comparison
"""

import subprocess
import time
import os
import sys

def run_command(command, description):
    """Run a command and measure execution time"""
    print(f"\n{'='*50}")
    print(f"Starting: {description}")
    print(f"Command: {command}")
    print(f"{'='*50}")
    
    start_time = time.time()
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        end_time = time.time()
        
        print(f"✅ Success! Time taken: {end_time - start_time:.2f} seconds")
        if result.stdout:
            print("Output:", result.stdout[-500:])  # Last 500 chars
        return True
    except subprocess.CalledProcessError as e:
        end_time = time.time()
        print(f"❌ Failed! Time taken: {end_time - start_time:.2f} seconds")
        print("Error:", e.stderr)
        return False

def get_file_size(filepath):
    """Get file size in MB"""
    if os.path.exists(filepath):
        size_bytes = os.path.getsize(filepath)
        return size_bytes / (1024 * 1024)
    return 0

def main():
    print("DocPDF Build Comparison Tool")
    print("This will build your app with both PyInstaller and Nuitka")
    
    # Clean previous builds
    print("\nCleaning previous builds...")
    if os.path.exists("dist"):
        subprocess.run("rmdir /s /q dist", shell=True)
    if os.path.exists("build"):
        subprocess.run("rmdir /s /q build", shell=True)
    
    # Install requirements
    print("\nInstalling requirements...")
    subprocess.run("pip install -r requirements.txt", shell=True)
    
    # Build with PyInstaller
    pyinstaller_success = run_command(
        "pyinstaller --onefile --windowed --name docpdf_pyinstaller docpdf.py",
        "PyInstaller Build"
    )
    
    # Move PyInstaller result
    if pyinstaller_success and os.path.exists("dist/docpdf_pyinstaller.exe"):
        subprocess.run("move dist\\docpdf_pyinstaller.exe dist\\docpdf_pyinstaller_backup.exe", shell=True)
    
    # Build with Nuitka
    nuitka_success = run_command(
        "python -m nuitka --onefile --windows-disable-console --enable-plugin=tk-inter --output-filename=docpdf_nuitka.exe --output-dir=dist --remove-output docpdf.py",
        "Nuitka Build"
    )
    
    # Compare results
    print(f"\n{'='*50}")
    print("BUILD COMPARISON RESULTS")
    print(f"{'='*50}")
    
    if pyinstaller_success:
        pyinstaller_size = get_file_size("dist/docpdf_pyinstaller_backup.exe")
        print(f"PyInstaller: ✅ Success - Size: {pyinstaller_size:.2f} MB")
    else:
        print("PyInstaller: ❌ Failed")
    
    if nuitka_success:
        nuitka_size = get_file_size("dist/docpdf_nuitka.exe")
        print(f"Nuitka: ✅ Success - Size: {nuitka_size:.2f} MB")
    else:
        print("Nuitka: ❌ Failed")
    
    if pyinstaller_success and nuitka_success:
        print(f"\nSize difference: {abs(pyinstaller_size - nuitka_size):.2f} MB")
        if nuitka_size < pyinstaller_size:
            print(f"Nuitka is {((pyinstaller_size - nuitka_size) / pyinstaller_size * 100):.1f}% smaller")
        else:
            print(f"PyInstaller is {((nuitka_size - pyinstaller_size) / nuitka_size * 100):.1f}% smaller")

if __name__ == "__main__":
    main()