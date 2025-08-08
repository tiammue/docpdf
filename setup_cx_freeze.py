import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_options = {
    'packages': ['docx', 'docx2pdf', 'pypdf', 'tkinter'],
    'excludes': [],
    'include_files': []
}

base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('docpdf.py', base=base, target_name='DocumentConverter.exe')
]

setup(
    name='DocumentConverter',
    version='1.0',
    description='Document Converter - DOC/DOCX to PDF and PDF to DOCX',
    options={'build_exe': build_options},
    executables=executables
)