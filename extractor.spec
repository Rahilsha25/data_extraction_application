# -*- mode: python ; coding: utf-8 -*-

import os

tesseract_path = r'C:\Users\em010\AppData\Local\Programs\Tesseract-OCR'
tess_binaries = [
    (os.path.join(tesseract_path, 'tesseract.exe'), '.'),
    (os.path.join(tesseract_path, 'tessdata'), 'tessdata')  # include language data if needed
]

a = Analysis(
    ['extractor.py'],
    pathex=[],
    binaries=tess_binaries,
    datas=[('electron-app/Input_Data.xlsx', '.')],
    hiddenimports=[
        'pymupdf',
        'fitz',
        'pandas',
        'docx',
        'openpyxl',
        'google.generativeai',
        'pytesseract',
        'PIL',
        'bs4',
        'bs4.element',
        'jinja2',  # used internally by bs4 sometimes
        'pdfplumber',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='extractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)