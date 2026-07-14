import re
import glob
import os
import fitz  # PyMuPDF — mendukung kompresi saat merge

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))

def find_rekap_pdf(dir_rekap, bidang, nama):
    folder = os.path.join(dir_rekap, bidang)
    files = glob.glob(os.path.join(folder, f"{nama}*.pdf"))
    return files[0] if files else None


def merge_pdf(files, output):
    """Gabungkan beberapa file PDF menjadi satu dan simpan dengan kompresi."""
    merged = fitz.open()

    for f in files:
        if f and os.path.exists(f):
            with fitz.open(f) as src:
                merged.insert_pdf(src)

    # Simpan dengan kompresi untuk memperkecil ukuran file final
    merged.save(output,
        garbage=4,          # hapus objek PDF tidak terpakai secara agresif
        deflate=True,       # kompres stream konten teks
        deflate_images=True,# kompres stream gambar
        deflate_fonts=True, # kompres font tertanam
        clean=True          # rapikan syntax internal PDF
    )
    merged.close()
