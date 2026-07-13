import os
import sys
from datetime import datetime

def get_app_directory():
    # Jika dijalankan sebagai .exe
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)

    # Jika dijalankan sebagai script python
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = get_app_directory()

FILES_DIR = os.path.join(APP_DIR, "files")
TEMPLATE_WORD = os.path.join(FILES_DIR, "TEMPLATE_TPP.docx")

JAM_MASUK_NORM = datetime.strptime("07:30:00", "%H:%M:%S")