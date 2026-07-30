import os
import fitz  # PyMuPDF
import pandas as pd
from utils import sanitize_filename
from steps.merge_pdf_jabatan import KELOMPOK_PIMPINAN, _find_pdf_for_pegawai


def _detect_kelompok(pdf_filename):
    """
    Deteksi kelompok pimpinan dari nama file PDF Gabungan.
    Mengembalikan (nama_kelompok, nilai_jabatan_pimpinan) atau None.
    """
    nama_lower = pdf_filename.lower()
    for nama_kelompok, nilai_jp, nama_file_base in KELOMPOK_PIMPINAN:
        if nama_file_base.lower() in nama_lower:
            return (nama_kelompok, nilai_jp)
    return None


def split_pdf_jabatan(dir_output, csv_file, bulan, tahun, log, pdf_gabungan_path=""):
    """
    Memecah file PDF Gabungan (yang telah ditandatangani) kembali menjadi file individu,
    berdasarkan urutan saat proses merge, lalu menyimpan ke folder bidang masing-masing.

    Otomatis mendeteksi kelompok (KaDis / Sekretaris) dari nama file PDF yang dimasukkan.
    User bisa menjalankan step ini 2x: sekali untuk file KaDis, sekali untuk Sekretaris.
    """
    log("── Memulai proses Pecah & Distribusi Dokumen TTD ──")

    if not csv_file or not os.path.exists(csv_file):
        log("❌ File CSV tidak ditemukan. Pastikan step Generate CSV sudah dijalankan.")
        return

    periode_str = f"{bulan}_{tahun}"

    # Menentukan file PDF Gabungan
    pdf_gabungan_path = pdf_gabungan_path.strip() if pdf_gabungan_path else ""
    if not pdf_gabungan_path or not os.path.exists(pdf_gabungan_path):
        log("❌ File PDF Gabungan tidak ditemukan.")
        log("Harap isi input 'File PDF Gabungan (TTD)' dengan file hasil scan yang sudah ditandatangani.")
        return

    pdf_filename = os.path.basename(pdf_gabungan_path)
    log(f"  Membaca PDF Gabungan: {pdf_filename}")

    # --- Deteksi kelompok dari nama file ---
    kelompok = _detect_kelompok(pdf_filename)
    if kelompok:
        nama_kelompok_aktif, nilai_jp_aktif = kelompok
        log(f"  🔍 Terdeteksi sebagai kelompok: {nama_kelompok_aktif}")
    else:
        # Fallback: jika nama file tidak cocok, coba semua kelompok
        log(f"  ⚠ Nama file tidak mengandung penanda kelompok (KaDis/Sekretaris).")
        log(f"    Akan memproses semua kelompok pimpinan secara berurutan.")
        nama_kelompok_aktif = None
        nilai_jp_aktif = None

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

    col_jp = "JABATAN PIMPINAN"
    for col_wajib in ["NAMA", col_jp]:
        if col_wajib not in df.columns:
            log(f"❌ Kolom '{col_wajib}' tidak ditemukan di CSV.")
            doc_gabungan.close()
            return

    df["NAMA"] = df["NAMA"].fillna("").str.strip()
    df[col_jp] = df[col_jp].fillna("").str.strip()
    if "JABATAN" in df.columns:
        df["JABATAN"] = df["JABATAN"].fillna("").str.strip()
    df = df[df["NAMA"] != ""]

    dir_output = os.path.abspath(dir_output)

    # --- Tentukan kelompok mana yang akan diproses ---
    if nama_kelompok_aktif is not None:
        # Hanya proses 1 kelompok yang terdeteksi
        kelompok_proses = [(nama_kelompok_aktif, nilai_jp_aktif)]
    else:
        # Fallback: proses semua kelompok
        kelompok_proses = [(nk, njp) for nk, njp, _ in KELOMPOK_PIMPINAN]

    current_page_idx = 0
    berhasil_split = 0

    for nama_kel, nilai_jp in kelompok_proses:
        from steps.download import stop_event
        if stop_event.is_set():
            log("Proses Split PDF dibatalkan oleh pengguna.")
            doc_gabungan.close()
            raise SystemExit()

        # Filter pegawai berdasarkan JABATAN PIMPINAN (sama seperti saat merge)
        pegawai_cocok = df[df[col_jp].str.upper() == nilai_jp.upper()].copy()

        if pegawai_cocok.empty:
            continue

        pegawai_cocok = pegawai_cocok.sort_values("NAMA")
        log(f"\n  📂 Ekstrak {nama_kel} ({len(pegawai_cocok)} pegawai):")

        for _, row in pegawai_cocok.iterrows():
            if stop_event.is_set():
                log("Proses Split PDF dibatalkan oleh pengguna.")
                doc_gabungan.close()
                raise SystemExit()

            nama_raw  = row["NAMA"]
            nama_safe = sanitize_filename(nama_raw)

            jabatan = row.get("JABATAN", "") if "JABATAN" in df.columns else ""
            if pd.isna(jabatan):
                jabatan = ""

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
                    log(f"    ❌ Halaman di PDF Gabungan kurang untuk '{nama_raw}'! "
                        f"Halaman ke-{current_page_idx} s/d {current_page_idx + pages_needed - 1} "
                        f"melebihi total {total_pages_gabungan}.")
                    log("    ❌ PROSES EKSTRAKSI TERHENTI KARENA KETIDAKSESUAIAN HALAMAN.")
                    doc_gabungan.close()
                    return

                # Ekstrak halaman
                doc_baru = fitz.open()
                doc_baru.insert_pdf(doc_gabungan,
                                    from_page=current_page_idx,
                                    to_page=current_page_idx + pages_needed - 1)

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

                    log(f"    ✔ {nama_raw} ({pages_needed} hlmn) -> {nama_file_baru}")
                except Exception as e:
                    log(f"    ❌ Gagal menyimpan '{nama_file_baru}': {e}")
                finally:
                    doc_baru.close()

                current_page_idx += pages_needed
            else:
                log(f"    ⚠ PDF Asli tidak ditemukan untuk: {nama_raw}")
                log(f"    ⚠ PERINGATAN: Urutan halaman akan meleset!")

    doc_gabungan.close()

    if current_page_idx < total_pages_gabungan:
        log(f"\n  ⚠ Ada {total_pages_gabungan - current_page_idx} sisa halaman di PDF Gabungan yang tidak terdistribusi.")

    log(f"\n✅ Berhasil mendistribusikan {berhasil_split} dokumen kembali ke folder bidang.")
