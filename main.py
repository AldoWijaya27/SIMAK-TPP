# ===== BOOTSTRAP PACKAGE (HARUS PALING ATAS) =====
import os
import sys

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)
# ==================================================

import tkinter as tk
from gui import App   # <-- import lokal, kompatibel dengan PyInstaller

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()