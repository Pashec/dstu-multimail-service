import os
import sys

# =========================================================================
# КРИТИЧЕСКИ ВАЖНО ДЛЯ WINDOWS EXE С console=False (NOCONSOLE РЕЖИМ):
# Если консоль отключена, sys.stdout и sys.stderr равны None.
# Uvicorn при попытке записать логи падает с ошибкой 'NoneType' has no attribute 'write'.
# Перенаправляем их в devnull:
# =========================================================================
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")
if sys.stdin is None:
    sys.stdin = open(os.devnull, "r", encoding="utf-8")

import time
import threading
import multiprocessing
import ctypes
import uvicorn
import webview
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from routers.broadcast import router as broadcast_router
from config import BASE_DIR, STATIC_DIR

# Инициализация FastAPI приложения
app = FastAPI(title="Email Broadcast Service")
app.include_router(broadcast_router)

# Безопасное монтирование статики
if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
else:
    app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")


def run_server():
    """Запуск Uvicorn-сервера в фоновом потоке"""
    # Полностью отключаем встроенное консольное логирование Uvicorn для noconsole
    config = uvicorn.Config(
        app=app,
        host="127.0.0.1",
        port=8000,
        log_level="critical",  # Чтобы не спамил в закрытые дескрипторы
        access_log=False,
        loop="asyncio"
    )
    server = uvicorn.Server(config)
    server.run()


if __name__ == "__main__":
    multiprocessing.freeze_support()

    try:
        myappid = "dstu.mailer.pro.v2"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    # 1. Запуск сервера Uvicorn
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()

    # 2. Небольшая пауза на инициализацию FastAPI
    time.sleep(1.0)

    # 3. Иконка
    icon_path = os.path.join(BASE_DIR, "app_icon.ico")
    if not os.path.exists(icon_path):
        icon_path = None

    # 4. Окно программы
    window = webview.create_window(
        title="DSTU Multi-mailer Pro",
        url="http://127.0.0.1:8000",
        width=1500,
        height=1100,
        resizable=True,
        min_size=(1200, 800)
    )

    webview.start(icon=icon_path if icon_path else None)

    # Принудительный чистый выход при закрытии окна
    os._exit(0)