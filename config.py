import os
import sys
from pathlib import Path

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).resolve().parent

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")

broadcast_status = {
    "is_running": False,
    "should_stop": False,
    "total": 0,
    "sent": 0,
    "skipped": 0,
    "errors": [],
    "logs": []
}

last_heartbeat_time = 0.0