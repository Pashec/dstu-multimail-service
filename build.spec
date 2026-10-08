# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Автоматический поиск папки со словарями pymorphy3
try:
    import pymorphy3_dicts_ru
    dict_data_path = os.path.join(os.path.dirname(pymorphy3_dicts_ru.__file__), "data")
except Exception:
    dict_data_path = None

datas = [
    ('templates', 'templates'),
    ('dstu-logo.png', '.'),
    ('dstu-logo.png', 'static'),
]

if dict_data_path and os.path.exists(dict_data_path):
    datas.append((dict_data_path, 'pymorphy3_dicts_ru/data'))

binaries = []

hiddenimports = [
    'uvicorn.logging',
    'uvicorn.loops',
    'uvicorn.loops.auto',
    'uvicorn.protocols',
    'uvicorn.protocols.http',
    'uvicorn.protocols.http.auto',
    'uvicorn.protocols.http.httptools_impl',
    'uvicorn.protocols.websockets',
    'uvicorn.protocols.websockets.auto',
    'uvicorn.lifespan',
    'uvicorn.lifespan.on',
    'uvicorn.lifespan.off',
    'aiosmtplib',
    'email_validator',
    'jinja2',
    'openpyxl',
    'docx',
    'pymorphy3',
    'pypdf',
    'pypdfium2',
    'PIL',
]

for pkg in ['uvicorn', 'webview', 'pypdf', 'pypdfium2', 'PIL']:
    try:
        pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(pkg)
        datas += pkg_datas
        binaries += pkg_binaries
        hiddenimports += [h for h in pkg_hiddenimports if 'android' not in h and 'ios' not in h]
    except Exception:
        pass

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['android', 'tkinter'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DGTU_Rassylka',
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
    icon='app_icon.ico',
)
