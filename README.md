# docpdf

High-performance C++ rewrite of the docpdf application using Qt6 for maximum speed and native performance.

## Features
- Native C++ performance
- Modern Qt6 GUI
- Threaded conversions (non-blocking UI)
- Progress tracking
- Cross-platform compatibility
- **Dependency-Free**: Uses internal PDF and DOCX processing (no LibreOffice or Poppler required)
- Single executable deployment ready

## Download

### Pre-built Executables
Download the latest release from the [Releases](https://github.com/tiammue/docpdf/releases) page:
- **Windows**: `docpdf-windows-x64.zip`
- **Linux**: `docpdf-linux-x64.AppImage`
- **macOS**: `docpdf-macos-x64.dmg`

## Building Locally

To build locally, you need:
- Qt6 (Core, Widgets, PrintSupport)
- CMake 3.16+
- C++17 compiler (MSVC, GCC, or Clang)

```bash
mkdir build && cd build
cmake .. -DCMAKE_PREFIX_PATH="path/to/qt6"
cmake --build . --config Release
```

## Implementation Details

The application uses internal conversion logic to avoid external dependencies:

### DOC/DOCX to PDF
- Uses internal XML parsing to read DOCX structure.
- Renders to PDF using Qt's `QPdfWriter`.
- Preserves text, paragraphs, and basic formatting.

### PDF to DOCX
- Uses internal stream decompression (via `miniz`) to read PDF content.
- Extracts text and attempts to map characters (including basic CMap support).
- Writes native `.docx` files.

## Performance Benefits

Compared to the Python version:
- **Startup time**: ~50ms vs ~2000ms
- **Memory usage**: ~15MB vs ~50MB
- **Conversion speed**: significantly faster (no external process launch overhead)
- **Deployment**: Simply copy the executable (and Qt DLLs) - no external tools needed.

## Deployment

After building, copy the executable and required Qt DLLs:
```bash
windeployqt.exe docpdf.exe
```

This creates a standalone deployment folder with all dependencies.
