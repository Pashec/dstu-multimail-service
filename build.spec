# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Добавляем шаблоны и файлы из корня
datas = [
    ('templates', 'templates'),
    ('dstu-logo.png', '.'),      # Кладется прямо в корень распакованного приложения
]

# Если в коде где-то используется обращение к static/DSTU-logo.png, 
# продублируем логотип и в виртуальную папку static, чтобы не было 404:
datas.append(('dstu-logo.png', 'static'))

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
]

# Собираем зависимости pywebview и uvicorn без мобильных платформ
for pkg in ['uvicorn', 'webview']:
    try:
        pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(pkg)
        datas += pkg_datas
        hiddenimports += [h for h in pkg_hiddenimports if 'android' not in h and 'ios' not in h]
    except Exception:
        pass

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
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
    name='DSTU_MultiMailer_Pro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # Без черного окна консоли
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',     # Настоящая иконка приложения из корня!
)