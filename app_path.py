import os
import sys

def base_dir():
    # saat jadi exe
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)

    # saat python biasa
    return os.path.dirname(os.path.abspath(__file__))

def get_file(*paths):
    return os.path.join(base_dir(), *paths)