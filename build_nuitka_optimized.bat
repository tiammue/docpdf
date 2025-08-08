@echo off
echo Installing required packages...
pip install -r requirements.txt

echo Building optimized executable with Nuitka...
python -m nuitka ^
    --onefile ^
    --windows-disable-console ^
    --enable-plugin=tk-inter ^
    --output-filename=DocumentConverter.exe ^
    --output-dir=dist ^
    --remove-output ^
    --assume-yes-for-downloads ^
    --msvc=latest ^
    --jobs=4 ^
    --show-progress ^
    --python-flag=-O ^
    --python-flag=no_asserts ^
    --include-package=docx ^
    --include-package=docx2pdf ^
    --include-package=pypdf ^
    docpdf.py

echo Build complete! Check the 'dist' folder for DocumentConverter.exe
pause