# ===== BOOTSTRAP PACKAGE (HARUS PALING ATAS) =====
import os
import sys

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

# Pastikan path browser Playwright terarah dengan benar jika berjalan sebagai exe
if getattr(sys, 'frozen', False):
    if 'PLAYWRIGHT_BROWSERS_PATH' not in os.environ:
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        if local_appdata:
            pw_cache = os.path.join(local_appdata, 'ms-playwright')
            if os.path.isdir(pw_cache):
                os.environ['PLAYWRIGHT_BROWSERS_PATH'] = pw_cache
# ==================================================

import customtkinter as ctk
from gui import App

def main():
    # Coba gunakan TkinterDnD wrapper untuk dukungan penuh drag and drop di OS
    try:
        from tkinterdnd2 import TkinterDnD
        class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.TkdndVersion = TkinterDnD._require(self)
        
        root = TkinterDnD_CTk(className="SIMAK-TPP")
    except Exception as e:
        print(f"TkinterDnD failed to load, drag and drop disabled: {e}")
        root = ctk.CTk(className="SIMAK-TPP")

    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()