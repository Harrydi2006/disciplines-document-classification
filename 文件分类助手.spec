# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['file_classifier.py'],
    pathex=[],
    binaries=[('C:\\Users\\35024\\AppData\\Local\\Programs\\Python\\Python311\\python311.dll', '.')],
    datas=[('gui', 'gui')],
    hiddenimports=['PIL', 'pytesseract', 'docx', 'PyPDF2', 'speech_recognition', 'pydub', 'py7zr', 'rarfile', 'zipfile'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='文件分类助手',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['resources\\icon.ico'],
)
