# ===== BOOTSTRAP PACKAGE (HARUS PALING ATAS) =====
import os
import sys
import types

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)

if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

if 'Aplikasi' not in sys.modules:
    aplikasi_mod = types.ModuleType('Aplikasi')
    aplikasi_mod.__path__ = [CURRENT_DIR]
    sys.modules['Aplikasi'] = aplikasi_mod

if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)
# ==================================================

import tkinter as tk
from gui import App   # <-- pakai import absolut

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()