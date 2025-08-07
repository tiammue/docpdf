@echo off
echo Installing required packages...
pip install -r requirements.txt

echo Building executable...
pyinstaller --onefile --windowed --name "DocumentConverter" document_converter.py

echo Build complete! Check the 'dist' folder for DocumentConverter.exe
pause