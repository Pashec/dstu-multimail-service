import os
import sys
import json
import base64
import uuid
import hashlib
import platform
from pathlib import Path

APP_NAME = "ДГТУ Рассылка"
APP_VERSION = "3.0"

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
    DATA_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent
    DATA_DIR = BASE_DIR

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STATE_FILE = os.path.join(DATA_DIR, "window_state.json")
CREDS_FILE = os.path.join(DATA_DIR, "credentials.json")

MAX_LOGS_HISTORY = 500

broadcast_status = {
    "is_running": False,
    "is_paused": False,
    "should_stop": False,
    "total": 0,
    "sent": 0,
    "skipped": 0,
    "errors": [],
    "logs": [],
    "last_report_file": None,
    "last_report_name": None,
    "current_recipient": None,
    "progress_percent": 0
}

last_heartbeat_time = 0.0


def append_log(message: str, is_error: bool = False):
    """Безопасное добавление лога с ограничением длины истории."""
    if is_error:
        broadcast_status["errors"].append(message)
        if len(broadcast_status["errors"]) > MAX_LOGS_HISTORY:
            broadcast_status["errors"].pop(0)

    broadcast_status["logs"].append(message)
    if len(broadcast_status["logs"]) > MAX_LOGS_HISTORY:
        broadcast_status["logs"].pop(0)


def encrypt_password(pwd: str) -> str:
    """
    Защищенное шифрование пароля.
    В Windows использует Windows Data Protection API (DPAPI via CryptProtectData).
    На других ОС или при сбое DPAPI использует обратимое машинное шифрование (salt via uuid.getnode).
    Никаких паролей в открытом виде на диске.
    """
    if not pwd:
        return ""

    if platform.system() == "Windows":
        try:
            import ctypes
            from ctypes import wintypes

            class DATA_BLOB(ctypes.Structure):
                _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

            CryptProtectData = ctypes.windll.crypt32.CryptProtectData
            LocalFree = ctypes.windll.kernel32.LocalFree

            in_bytes = pwd.encode("utf-8")
            in_blob = DATA_BLOB(len(in_bytes), ctypes.create_string_buffer(in_bytes, len(in_bytes)))
            out_blob = DATA_BLOB()

            if CryptProtectData(ctypes.byref(in_blob), "DGTU_Creds", None, None, None, 0, ctypes.byref(out_blob)):
                encrypted_data = ctypes.string_at(out_blob.pbData, out_blob.cbData)
                LocalFree(out_blob.pbData)
                return "dpapi:" + base64.b64encode(encrypted_data).decode("ascii")
        except Exception:
            pass

    # Кроссплатформенный надежный fallback
    salt = hashlib.sha256(str(uuid.getnode()).encode("utf-8") + b"dgtu_salt_2026").digest()
    pwd_bytes = pwd.encode("utf-8")
    enc = bytes([b ^ salt[i % len(salt)] for i, b in enumerate(pwd_bytes)])
    return "enc:" + base64.b64encode(enc).decode("ascii")


def decrypt_password(enc_pwd: str) -> str:
    """Расшифровка сохраненного пароля с поддержкой DPAPI и fallback."""
    if not enc_pwd:
        return ""

    # Если передан старый пароль без префикса (обратная совместимость)
    if not enc_pwd.startswith("dpapi:") and not enc_pwd.startswith("enc:"):
        return enc_pwd

    if enc_pwd.startswith("dpapi:") and platform.system() == "Windows":
        try:
            import ctypes
            from ctypes import wintypes

            class DATA_BLOB(ctypes.Structure):
                _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

            CryptUnprotectData = ctypes.windll.crypt32.CryptUnprotectData
            LocalFree = ctypes.windll.kernel32.LocalFree

            raw = base64.b64decode(enc_pwd[6:])
            in_blob = DATA_BLOB(len(raw), ctypes.create_string_buffer(raw, len(raw)))
            out_blob = DATA_BLOB()

            if CryptUnprotectData(ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out_blob)):
                decrypted_bytes = ctypes.string_at(out_blob.pbData, out_blob.cbData)
                LocalFree(out_blob.pbData)
                return decrypted_bytes.decode("utf-8")
        except Exception:
            pass

    if enc_pwd.startswith("enc:"):
        try:
            raw = base64.b64decode(enc_pwd[4:])
            salt = hashlib.sha256(str(uuid.getnode()).encode("utf-8") + b"dgtu_salt_2026").digest()
            dec = bytes([b ^ salt[i % len(salt)] for i, b in enumerate(raw)])
            return dec.decode("utf-8")
        except Exception:
            return ""

    return ""


def load_window_state() -> dict:
    """Загрузка сохраненных координат и размеров окна."""
    default_state = {"width": 1460, "height": 960, "x": None, "y": None}
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return {**default_state, **data}
        except Exception:
            pass
    return default_state


def save_window_state(data: dict):
    """Сохранение позиции и размеров окна программы."""
    try:
        current = load_window_state()
        current.update(data)
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(current, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[CONFIG] Ошибка сохранения положения окна: {e}")


def load_saved_credentials() -> dict:
    """Загрузка сохраненных учетных данных с автоматической расшифровкой пароля."""
    defaults = {
        "email": "",
        "password": "",
        "server": "mail.donstu.ru",
        "port": 25,
        "preset": "dstu"
    }
    if os.path.exists(CREDS_FILE):
        try:
            with open(CREDS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    res = {**defaults, **data}
                    # Расшифровываем пароль для подстановки в UI
                    if "password" in res and res["password"]:
                        res["password"] = decrypt_password(res["password"])
                    return res
        except Exception:
            pass
    return defaults


def save_credentials_to_file(data: dict):
    """Сохранение учетных данных отправителя с шифрованием пароля."""
    try:
        current = load_saved_credentials()
        current.update(data)
        to_save = dict(current)
        if "password" in to_save and to_save["password"]:
            to_save["password"] = encrypt_password(to_save["password"])

        with open(CREDS_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[CONFIG] Ошибка сохранения учетных данных: {e}")


def clear_saved_credentials_file():
    """Удаление сохраненного файла учетных данных."""
    try:
        if os.path.exists(CREDS_FILE):
            os.remove(CREDS_FILE)
    except Exception:
        pass
