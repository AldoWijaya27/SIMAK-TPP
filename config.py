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

def get_files_dir():
    """
    Mengembalikan path folder 'files'.
    - Prioritas utama: di samping .exe (lokasi yang bisa diedit user)
    - Fallback: di dalam _internal/ (bundled oleh PyInstaller v6+, via sys._MEIPASS)
    """
    # Lokasi di samping .exe (user-editable override)
    beside_exe = os.path.join(APP_DIR, "files")
    if os.path.isdir(beside_exe):
        return beside_exe

    # Fallback: PyInstaller v6+ menyimpan datas di _internal/ (sys._MEIPASS)
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        internal = os.path.join(sys._MEIPASS, "files")
        if os.path.isdir(internal):
            return internal

    return beside_exe  # default, meski belum ada

FILES_DIR = get_files_dir()
TEMPLATE_WORD = os.path.join(FILES_DIR, "TEMPLATE_TPP.docx")

JAM_MASUK_NORM = datetime.strptime("07:30:00", "%H:%M:%S")