# ===== BOOTSTRAP PACKAGE (HARUS PALING ATAS) =====
import os
import sys

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)
# ==================================================

import customtkinter as ctk
from gui import App

def main():
    # First create the root
    root = ctk.CTk(className="SIMAK-TPP")

    # Try to inject TkinterDnD for drag-and-drop support
    try:
        from tkinterdnd2 import TkinterDnD
        TkinterDnD._require(root)
    except Exception as e:
        print(f"TkinterDnD failed to load, drag and drop disabled: {e}")

    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()