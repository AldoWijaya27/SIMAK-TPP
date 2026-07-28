import os
import fitz  # PyMuPDF
import pandas as pd
from utils import sanitize_filename
from steps.merge_pdf_jabatan import KELOMPOK_JABATAN, _find_pdf_for_pegawai, NAMA_FILE_OUTPUT

def split_pdf_jabatan(dir_output, csv_file, bulan, tahun, log, pdf_gabungan_path=""):
    """
    Memecah file PDF Gabungan (yang telah ditandatangani) kembali menjadi file individu,
    berdasarkan urutan saat proses merge, lalu menyimpan ke folder bidang masing-masing.
    """
    log("── Memulai proses Pecah & Distribusi Dokumen TTD ──")

    if not csv_file or not os.path.exists(csv_file):
        log("❌ File CSV tidak ditemukan. Pastikan step Generate CSV sudah dijalankan.")
        return

    periode_str = f"{bulan}_{tahun}"
    
    # Menentukan file PDF Gabungan
    pdf_gabungan_path = pdf_gabungan_path.strip() if pdf_gabungan_path else ""
    if not pdf_gabungan_path or not os.path.exists(pdf_gabungan_path):
        nama_file_default = f"{NAMA_FILE_OUTPUT}_{periode_str}.pdf"
        pdf_gabungan_path = os.path.join(dir_output, nama_file_default)
        if not os.path.exists(pdf_gabungan_path):
            log(f"❌ File PDF Gabungan tidak ditemukan: {pdf_gabungan_path}")
            log("Harap isi input 'File PDF Gabungan (TTD)' jika menggunakan file kustom.")
            return

    log(f"  Membaca PDF Gabungan: {os.path.basename(pdf_gabungan_path)}")
    try:
        doc_gabungan = fitz.open(pdf_gabungan_path)
    except Exception as e:
        log(f"❌ Gagal membuka PDF Gabungan: {e}")
        return
        
    total_pages_gabungan = len(doc_gabungan)
    log(f"  Total halaman PDF Gabungan: {total_pages_gabungan}")

    # --- Baca CSV ---
    try:
        df = pd.read_csv(csv_file, sep=";", encoding="utf-8", dtype=str)
        df.columns = df.columns.str.strip().str.upper()
    except Exception as e:
        log(f"❌ Gagal membaca CSV: {e}")
        doc_gabungan.close()
        return

    if "NAMA" not in df.columns or "JABATAN" not in df.columns:
        log("❌ Kolom 'NAMA' atau 'JABATAN' tidak ditemukan di CSV.")
        doc_gabungan.close()
        return

    df["NAMA"]    = df["NAMA"].fillna("").str.strip()
    df["JABATAN"] = df["JABATAN"].fillna("").str.strip()
    df = df[df["NAMA"] != ""]

    dir_output = os.path.abspath(dir_output)
    
    current_page_idx = 0
    berhasil_split = 0

    for nama_kelompok, kata_kunci_list in KELOMPOK_JABATAN:
        from steps.download import stop_event
        if stop_event.is_set():
            log("Proses Split PDF dibatalkan oleh pengguna.")
            doc_gabungan.close()
            raise SystemExit()

        # Filter pegawai yang jabatannya cocok
        def _cocok(jabatan_pegawai, kw_list=kata_kunci_list):
            j = jabatan_pegawai.lower()
            return any(kw.lower() in j for kw in kw_list)

        pegawai_cocok = df[df["JABATAN"].apply(_cocok)].copy()

        if pegawai_cocok.empty:
            continue

        pegawai_cocok = pegawai_cocok.sort_values("NAMA")
        log(f"\n  📂 Ekstrak {nama_kelompok} ({len(pegawai_cocok)} pegawai):")

        for _, row in pegawai_cocok.iterrows():
            nama_raw  = row["NAMA"]
            jabatan   = row["JABATAN"]
            nama_safe = sanitize_filename(nama_raw)

            pdf_asli_path = _find_pdf_for_pegawai(dir_output, nama_safe)

            if pdf_asli_path:
                # Cek jumlah halaman PDF asli untuk memotong dengan jumlah yang persis sama
                try:
                    doc_asli = fitz.open(pdf_asli_path)
                    pages_needed = len(doc_asli)
                    doc_asli.close()
                except Exception as e:
                    log(f"    ⚠ Gagal cek hlmn PDF asli '{nama_raw}': {e}. Asumsi 1 hlmn.")
                    pages_needed = 1
                    
                if current_page_idx + pages_needed > total_pages_gabungan:
                    log(f"    ❌ Halaman di PDF Gabungan kurang untuk '{nama_raw}'! Halaman ke-{current_page_idx} s/d {current_page_idx + pages_needed - 1} melebihi total {total_pages_gabungan}.")
                    log("    ❌ PROSES EKSTRAKSI TERHENTI KARENA KETIDAKSESUAIAN HALAMAN.")
                    doc_gabungan.close()
                    return
                    
                # Ekstrak halaman
                doc_baru = fitz.open()
                doc_baru.insert_pdf(doc_gabungan, from_page=current_page_idx, to_page=current_page_idx + pages_needed - 1)
                
                folder_bidang = os.path.dirname(pdf_asli_path)
                nama_file_baru = f"{nama_safe}_TTD.pdf"
                path_simpan = os.path.join(folder_bidang, nama_file_baru)
                
                try:
                    doc_baru.save(path_simpan)
                    berhasil_split += 1
                    
                    # Hapus file PDF aslinya karena diminta replace dengan nama baru
                    if os.path.exists(pdf_asli_path) and pdf_asli_path != path_simpan:
                        try:
                            os.remove(pdf_asli_path)
                        except OSError as e:
                            log(f"    ⚠ Gagal menghapus file lama '{pdf_asli_path}': {e}")
                        
                    log(f"    ✔ {nama_raw} ({pages_needed} hlmn) -> {nama_file_baru} (File asli diganti)")
                except Exception as e:
                    log(f"    ❌ Gagal menyimpan '{nama_file_baru}': {e}")
                finally:
                    doc_baru.close()
                
                current_page_idx += pages_needed
            else:
                log(f"    ⚠ PDF Asli tidak ditemukan untuk: {nama_raw}")
                log(f"    ⚠ PERINGATAN: Urutan halaman akan meleset! Hentikan proses jika ini tidak disengaja.")

    doc_gabungan.close()
    
    if current_page_idx < total_pages_gabungan:
        log(f"\n  ⚠ Ada {total_pages_gabungan - current_page_idx} sisa halaman di PDF Gabungan yang tidak terdistribusi.")
        
    log(f"\n✅ Berhasil mendistribusikan {berhasil_split} dokumen kembali ke folder bidang.")
