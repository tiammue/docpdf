@echo off
echo Installing required packages...
pip install -r requirements.txt

echo Building executable...
pyinstaller --onefile --windowed --name "docpdf" docpdf.py

echo Build complete! Check the 'dist' folder for docpdf.exe
pause