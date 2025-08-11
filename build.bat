@echo off
echo Building docpdf C++ application...

REM Create build directory
if not exist build mkdir build
cd build

REM Configure with CMake
cmake .. -G "Visual Studio 17 2022" -A x64

REM Build the project
cmake --build . --config Release

echo.
echo Build complete! Executable is in build\Release\docpdf.exe
pause