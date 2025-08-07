# Changelog

All notable changes to this project will be documented in this file.

## [1.0.0] - 2025-01-08

### Added
- Initial release of Document Converter
- GUI application with tkinter interface
- DOC/DOCX to PDF conversion using docx2pdf library
- PDF to DOCX conversion using pypdf library
- Batch processing of all files in directory
- Progress tracking and status updates
- Threaded conversion to prevent GUI freezing
- Standalone executable creation with PyInstaller
- Automatic detection of executable directory
- Filtering of temporary Word files (~$ files)

### Features
- Simple two-button interface
- Drag-and-drop usage (place exe in folder)
- Real-time conversion progress
- Error handling and user feedback
- Preserves original files
- Works with Microsoft Word integration