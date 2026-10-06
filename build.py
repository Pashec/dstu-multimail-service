import os
import PyInstaller.__main__
import pymorphy3_dicts_ru

# Получаем путь к папке data со словарями
dict_data_path = os.path.join(os.path.dirname(pymorphy3_dicts_ru.__file__), "data")

print(f"[СБОРКА] Найдены словари pymorphy3: {dict_data_path}")

PyInstaller.__main__.run([
    'main.py',
    '--name=DSTU_Mail_Sender',
    '--onefile',
    '--noconsole',
    '--clean',
    # Включаем шаблоны HTML
    '--add-data=templates;templates',
    # Включаем словари русского языка
    f'--add-data={dict_data_path};pymorphy3_dicts_ru/data',
    # Скрытые импорты для Uvicorn и асинхронных библиотек
    '--hidden-import=uvicorn.logging',
    '--hidden-import=uvicorn.loops',
    '--hidden-import=uvicorn.loops.auto',
    '--hidden-import=uvicorn.protocols',
    '--hidden-import=uvicorn.protocols.http',
    '--hidden-import=uvicorn.protocols.http.auto',
    '--hidden-import=uvicorn.lifespans',
    '--hidden-import=uvicorn.lifespans.on',
    '--hidden-import=aiosmtplib',
    '--hidden-import=openpyxl',
    '--hidden-import=jinja2',
    '--hidden-import=pymorphy3',
])