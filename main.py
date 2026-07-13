# ===== BOOTSTRAP PACKAGE (HARUS PALING ATAS) =====
import os
import sys

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)

if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)
# ==================================================

import tkinter as tk
from Aplikasi.gui import App   # <-- pakai import absolut

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()