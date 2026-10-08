import os
import sys
import PyInstaller.__main__
import pymorphy3_dicts_ru

dict_data_path = os.path.join(os.path.dirname(pymorphy3_dicts_ru.__file__), "data")

print(f"[СБОРКА] Найдены словари pymorphy3: {dict_data_path}")

separator = ";" if sys.platform.startswith("win") else ":"

PyInstaller.__main__.run([
    'main.py',
    '--name=DGTU_Rassylka',
    '--onefile',
    '--noconsole',
    '--clean',
    f'--add-data=templates{separator}templates',
    f'--add-data=dstu-logo.png{separator}.',
    f'--add-data=dstu-logo.png{separator}static',
    f'--add-data={dict_data_path}{separator}pymorphy3_dicts_ru/data',
    '--icon=app_icon.ico',
    '--hidden-import=uvicorn.logging',
    '--hidden-import=uvicorn.loops',
    '--hidden-import=uvicorn.loops.auto',
    '--hidden-import=uvicorn.protocols',
    '--hidden-import=uvicorn.protocols.http',
    '--hidden-import=uvicorn.protocols.http.auto',
    '--hidden-import=uvicorn.protocols.http.httptools_impl',
    '--hidden-import=uvicorn.protocols.websockets',
    '--hidden-import=uvicorn.protocols.websockets.auto',
    '--hidden-import=uvicorn.lifespan',
    '--hidden-import=uvicorn.lifespan.on',
    '--hidden-import=uvicorn.lifespan.off',
    '--hidden-import=aiosmtplib',
    '--hidden-import=openpyxl',
    '--hidden-import=jinja2',
    '--hidden-import=pymorphy3',
    '--hidden-import=pypdf',
    '--hidden-import=pypdfium2',
    '--hidden-import=PIL',
    '--hidden-import=docx',
    '--hidden-import=email_validator',
    '--collect-all=webview',
    '--collect-all=uvicorn',
    '--collect-all=pypdf',
    '--collect-all=pypdfium2',
    '--collect-all=PIL',
])
