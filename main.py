import os
import sys

# =========================================================================
# КРИТИЧЕСКИ ВАЖНО ДЛЯ WINDOWS EXE С console=False (NOCONSOLE РЕЖИМ):
# Если консоль отключена, sys.stdout и sys.stderr равны None.
# Uvicorn при попытке записать логи падает с ошибкой 'NoneType' has no attribute 'write'.
# =========================================================================
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")
if sys.stdin is None:
    sys.stdin = open(os.devnull, "r", encoding="utf-8")

import time
import socket
import threading
import multiprocessing
import ctypes
import uvicorn
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

try:
    import webview
except ImportError:
    webview = None

from routers.broadcast import router as broadcast_router
from config import BASE_DIR, STATIC_DIR, DATA_DIR, load_window_state, save_window_state, APP_NAME

# Настройка постоянной папки профиля WebView2 (устраняет WinError 5 при закрытии приложения)
WEBVIEW_STORAGE_DIR = os.path.join(DATA_DIR, "webview_storage")
try:
    os.makedirs(WEBVIEW_STORAGE_DIR, exist_ok=True)
    os.environ["WEBVIEW2_USER_DATA_FOLDER"] = WEBVIEW_STORAGE_DIR
except Exception:
    pass

app = FastAPI(title=APP_NAME)
app.include_router(broadcast_router)

if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
else:
    app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")


def wait_for_server(host: str = "127.0.0.1", port: int = 8000, timeout: float = 10.0) -> bool:
    """Ожидание готовности порта Uvicorn перед открытием WebView окна."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.3)
            if sock.connect_ex((host, port)) == 0:
                return True
        time.sleep(0.15)
    return False


def run_server():
    """Запуск Uvicorn-сервера в фоновом потоке."""
    srv_config = uvicorn.Config(
        app=app,
        host="127.0.0.1",
        port=8000,
        log_level="error",
        access_log=False,
        loop="asyncio"
    )
    server = uvicorn.Server(srv_config)
    server.run()


if __name__ == "__main__":
    multiprocessing.freeze_support()

    try:
        myappid = "dstu.rassylka.v3"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    if webview is None:
        print(f"[СЕРВЕР] Запуск «{APP_NAME}» в режиме браузера (pywebview не обнаружен)...")
        print("[СЕРВЕР] Откройте в браузере: http://127.0.0.1:8000")
        run_server()
    else:
        # 1. Запуск сервера Uvicorn в фоне
        server_thread = threading.Thread(target=run_server, daemon=True)
        server_thread.start()

        # 2. Ожидание фактического старта сервера
        wait_for_server("127.0.0.1", 8000, timeout=10.0)

        # 3. Иконка приложения
        icon_path = os.path.join(BASE_DIR, "app_icon.ico")
        if not os.path.exists(icon_path):
            icon_path = None

        # 4. Загрузка сохраненного положения и размеров окна
        w_state = load_window_state()
        window_kwargs = {
            "title": APP_NAME,
            "url": "http://127.0.0.1:8000",
            "width": w_state.get("width", 1460),
            "height": w_state.get("height", 960),
            "resizable": True,
            "min_size": (1150, 750)
        }
        if w_state.get("x") is not None and w_state.get("y") is not None:
            window_kwargs["x"] = w_state["x"]
            window_kwargs["y"] = w_state["y"]

        window = webview.create_window(**window_kwargs)

        # 5. Сохранение положения при закрытии окна
        def on_window_closing():
            try:
                save_window_state({
                    "x": int(window.x),
                    "y": int(window.y),
                    "width": int(window.width),
                    "height": int(window.height)
                })
            except Exception:
                pass

        try:
            window.events.closing += on_window_closing
        except Exception:
            pass

        # Инициализация с постоянным хранилищем WebView2 для предотвращения ошибок удаления temp-папки при закрытии
        start_kwargs = {
            "private_mode": False,
            "storage_path": WEBVIEW_STORAGE_DIR
        }
        if icon_path:
            start_kwargs["icon"] = icon_path

        try:
            webview.start(**start_kwargs)
        except TypeError:
            # Fallback для старых версий pywebview
            webview.start(icon=icon_path if icon_path else None)
        except Exception as e:
            pass
        finally:
            time.sleep(0.15)
            os._exit(0)
