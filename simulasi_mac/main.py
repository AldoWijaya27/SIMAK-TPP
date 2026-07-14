# ===== BOOTSTRAP PACKAGE SIMULASI =====
import os
import sys
import types

# Dapatkan folder root utama (satu tingkat di atas simulasi_mac)
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Masukkan folder root utama ke path pencarian Python
sys.path.insert(0, ROOT_DIR)

# Alias 'Aplikasi' ke folder root utama
if 'Aplikasi' not in sys.modules:
    aplikasi_mod = types.ModuleType('Aplikasi')
    aplikasi_mod.__path__ = [ROOT_DIR]
    sys.modules['Aplikasi'] = aplikasi_mod
# =======================================

import tkinter as tk
from Aplikasi.gui import App

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
