@echo off
echo Installing required packages...
pip install -r requirements.txt
pip install cx_freeze

echo Building executable with cx_Freeze...
python setup_cx_freeze.py build

echo Moving executable to dist folder...
if not exist dist mkdir dist
copy build\exe.win-amd64-3.13\*.exe dist\
copy build\exe.win-amd64-3.13\*.dll dist\

echo Build complete! Check the 'dist' folder for DocumentConverter.exe
pause