# ===== BOOTSTRAP PACKAGE (HARUS PALING ATAS) =====
import os
import sys

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)

if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)
# ==================================================

import customtkinter as ctk
from gui import App

def main():
    # Use TkinterDnD for drag-and-drop support if available
    try:
        from tkinterdnd2 import TkinterDnD
        root = ctk.CTk(className="SIMAK-TPP")
        # Inject TkinterDnD capabilities into the CTk root
        TkinterDnD._require(root)
    except Exception:
        root = ctk.CTk()

    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()