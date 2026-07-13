import re
import glob
import os
from PyPDF2 import PdfMerger

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))

def find_rekap_pdf(dir_rekap, bidang, nama):
    folder = os.path.join(dir_rekap, bidang)
    files = glob.glob(os.path.join(folder, f"{nama}*.pdf"))
    return files[0] if files else None


def merge_pdf(files, output):
    merger = PdfMerger()
    for f in files:
        if f and os.path.exists(f):
            merger.append(f)
    merger.write(output)
    merger.close()
