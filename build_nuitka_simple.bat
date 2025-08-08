@echo off
echo Installing required packages...
pip install -r requirements.txt

echo Building executable with Nuitka (simplified)...
python -m nuitka ^
    --onefile ^
    --windows-disable-console ^
    --enable-plugin=tk-inter ^
    --output-filename=DocumentConverter.exe ^
    --output-dir=dist ^
    --assume-yes-for-downloads ^
    docpdf.py

echo Build complete! Check the 'dist' folder for DocumentConverter.exe
pause