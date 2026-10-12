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


def find_skp_pdf(dir_skp, nama):
    """
    Cari file PDF SKP Ekin pegawai di dir_skp.
    Format nama prioritas:
      1. SKP_*_{nama}.pdf (format baru: SKP_KADIS_Nama dengan gelar.pdf)
      2. SKP_*_{nama}*.pdf (format bertanggal/periode)
      3. *{nama}*.pdf (fallback pencarian umum)
    Abaikan file 'gabungan_*' dan file '*_TTD.pdf'.
    """
    if not dir_skp or not os.path.exists(dir_skp):
        return None

    nama_safe = sanitize_filename(nama)

    # 1. Exact match dengan prefix SKP_
    pattern_exact = os.path.join(dir_skp, f"SKP_*_{nama_safe}.pdf")
    matches = glob.glob(pattern_exact)
    matches = [m for m in matches if not os.path.basename(m).startswith("gabungan_") and not os.path.basename(m).endswith("_TTD.pdf")]
    if matches:
        return matches[0]

    # 2. Match dengan prefix SKP_ dan suffix apapun
    pattern_prefix = os.path.join(dir_skp, f"SKP_*_{nama_safe}*.pdf")
    matches = glob.glob(pattern_prefix)
    matches = [m for m in matches if not os.path.basename(m).startswith("gabungan_") and not os.path.basename(m).endswith("_TTD.pdf")]
    if matches:
        return matches[0]

    # 3. Fallback: file yang memuat nama_safe
    pattern_fallback = os.path.join(dir_skp, f"*{nama_safe}*.pdf")
    matches = glob.glob(pattern_fallback)
    matches = [m for m in matches if not os.path.basename(m).startswith("gabungan_") and not os.path.basename(m).endswith("_TTD.pdf")]
    if matches:
        return matches[0]

    return None


def is_skp_already_included(pdf_pegawai_path, skp_path):
    """
    Memeriksa apakah dokumen SKP sudah tergabung di bagian akhir pdf_pegawai.
    """
    if not pdf_pegawai_path or not skp_path:
        return False
    if not os.path.exists(pdf_pegawai_path) or not os.path.exists(skp_path):
        return False
    try:
        doc_pegawai = fitz.open(pdf_pegawai_path)
        doc_skp = fitz.open(skp_path)
        len_p = len(doc_pegawai)
        len_s = len(doc_skp)
        if len_p <= len_s or len_s == 0:
            doc_pegawai.close()
            doc_skp.close()
            return False

        # Bandingkan cuplikan teks halaman terakhir
        txt_p = doc_pegawai[-1].get_text().strip()
        txt_s = doc_skp[-1].get_text().strip()
        doc_pegawai.close()
        doc_skp.close()

        if txt_s and txt_p:
            s_snippet = "".join(txt_s.split())[:80]
            p_snippet = "".join(txt_p.split())[:80]
            if s_snippet and s_snippet == p_snippet:
                return True
        return False
    except Exception:
        return False


def is_skp_closing_page(page_text, nama="", nip=""):
    """
    Mendeteksi apakah suatu halaman merupakan halaman terakhir/penutup SKP
    (halaman yang memuat tanda tangan Pegawai yang Dinilai dan Pejabat Penilai Kinerja).
    """
    if not page_text or not page_text.strip():
        return False

    text_upper = page_text.upper()

    # Kriteria Wajib: Memuat teks Penilai Kinerja dan Pegawai yang Dinilai
    has_penilai = ("PEJABAT PENILAI" in text_upper)
    has_dinilai = ("PEGAWAI YANG DINILAI" in text_upper)

    if not (has_penilai and has_dinilai):
        return False

    # 1. Jika NIP diberikan: cek apakah NIP pegawai ada di halaman (paling akurat)
    if nip:
        nip_clean = re.sub(r'\D', '', str(nip))
        text_digits = re.sub(r'\D', '', page_text)
        if len(nip_clean) >= 9:
            if nip_clean in text_digits:
                return True
            return False

    # 2. Jika NIP tidak ada tapi Nama ada: cek apakah token nama utama ada di halaman
    if nama:
        nama_bersih = re.sub(r'[^A-Z\s]', ' ', nama.upper())
        gelar_stopwords = {
            "DR", "DRA", "DRS", "SH", "SE", "SOS", "SI", "MM", "MH", "MSI",
            "MSC", "HUT", "KOM", "ST", "IP", "AP", "SPD", "MPD", "SKM", "MKES"
        }
        tokens = [t for t in nama_bersih.split() if len(t) >= 3 and t not in gelar_stopwords]
        if tokens:
            return all(tok in text_upper for tok in tokens)

    # 3. Fallback jika nama dan NIP keduanya kosong
    has_ttd_marker = (
        ("${TTD_PENGIRIM}" in text_upper) or
        ("BANDAR LAMPUNG" in text_upper) or
        ("LAMPUNG" in text_upper)
    )
    return has_ttd_marker

