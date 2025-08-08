@echo off
echo Installing required packages...
pip install -r requirements.txt

echo Building executable with Nuitka...
python -m nuitka ^
    --onefile ^
    --windows-disable-console ^
    --enable-plugin=tk-inter ^
    --output-filename=docpdf.exe ^
    --output-dir=dist ^
    --remove-output ^
    docpdf.py

echo Build complete! Check the 'dist' folder for docpdf.exe
pause