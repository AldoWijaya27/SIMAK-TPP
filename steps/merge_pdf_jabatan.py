import os
import glob
import pandas as pd

from utils import merge_pdf, sanitize_filename


# ============================================================
# DEFINISI JABATAN YANG DIMASUKKAN KE DALAM MERGE
# Setiap entry: (nama_kelompok, [kata_kunci_substring])
# Pencocokan dilakukan secara case-insensitive substring match.
# Tambahkan kelompok baru di sini sesuai kebutuhan.
# ============================================================
KELOMPOK_JABATAN = [
    ("Kepala Dinas", ["kepala dinas"]),
    ("Sekretaris",   ["sekretaris"]),
    ("Madya",        ["madya"]),
]

# Nama file output tunggal (tanpa ekstensi)
NAMA_FILE_OUTPUT = "Merge_Pejabat"


def _find_pdf_for_pegawai(dir_output, nama_sanitized):
    """
    Cari file PDF pegawai di semua subfolder DIR_OUTPUT.
    Struktur folder: DIR_OUTPUT/<BIDANG>/<NAMA>.pdf
    Kembalikan path lengkap jika ditemukan, atau None.
    """
    pattern = os.path.join(dir_output, "**", f"{nama_sanitized}.pdf")
    matches = glob.glob(pattern, recursive=True)

    # Abaikan file yang ada di root DIR_OUTPUT (itu file hasil merge)
    matches = [m for m in matches if os.path.dirname(m) != os.path.abspath(dir_output)]

    return matches[0] if matches else None


def merge_pdf_by_jabatan(dir_output, csv_file, bulan, tahun, log):
    """
    Gabungkan semua PDF dari jabatan yang ditentukan menjadi SATU file PDF.

    Args:
        dir_output (str): Folder root PERHITUNGAN TPP
        csv_file   (str): Path ke Disiplin_TPP_Lengkap.csv
        bulan      (str): Bulan periode (mis. "06")
        tahun      (int): Tahun periode (mis. 2026)
        log        (callable): Fungsi log ke GUI
    """
    log("── Memulai proses Merge PDF per Jabatan ──")

    # --- Validasi CSV ---
    if not csv_file or not os.path.exists(csv_file):
        log("❌ File CSV tidak ditemukan. Pastikan step Generate CSV sudah dijalankan.")
        return

    # --- Baca CSV ---
    try:
        df = pd.read_csv(csv_file, sep=";", encoding="utf-8", dtype=str)
        df.columns = df.columns.str.strip().str.upper()
    except Exception as e:
        log(f"❌ Gagal membaca CSV: {e}")
        return

    if "NAMA" not in df.columns:
        log("❌ Kolom 'NAMA' tidak ditemukan di CSV.")
        return

    if "JABATAN" not in df.columns:
        log("❌ Kolom 'JABATAN' tidak ditemukan di CSV. Pastikan data pegawai memiliki kolom JABATAN.")
        return

    # Bersihkan NaN
    df["NAMA"]    = df["NAMA"].fillna("").str.strip()
    df["JABATAN"] = df["JABATAN"].fillna("").str.strip()

    # Filter baris yang punya nama
    df = df[df["NAMA"] != ""]

    log(f"  Total pegawai di CSV: {len(df)}")

    dir_output  = os.path.abspath(dir_output)
    periode_str = f"{bulan}_{tahun}"

    # --- Kumpulkan semua PDF dari seluruh kelompok jabatan ---
    # Urutan: per kelompok (sesuai urutan KELOMPOK_JABATAN), lalu nama alfabetis
    semua_pdf   = []
    total_cocok = 0

    for nama_kelompok, kata_kunci_list in KELOMPOK_JABATAN:
        from steps.download import stop_event
        if stop_event.is_set():
            log("Proses Merge PDF per Jabatan dibatalkan oleh pengguna.")
            raise SystemExit()

        # Filter pegawai yang jabatannya cocok
        def _cocok(jabatan_pegawai, kw_list=kata_kunci_list):
            j = jabatan_pegawai.lower()
            return any(kw.lower() in j for kw in kw_list)

        pegawai_cocok = df[df["JABATAN"].apply(_cocok)].copy()

        if pegawai_cocok.empty:
            log(f"  ⚠ Tidak ada pegawai dengan jabatan '{nama_kelompok}'. Dilewati.")
            continue

        pegawai_cocok = pegawai_cocok.sort_values("NAMA")
        log(f"\n  📂 {nama_kelompok} ({len(pegawai_cocok)} pegawai):")

        for _, row in pegawai_cocok.iterrows():
            nama_raw  = row["NAMA"]
            jabatan   = row["JABATAN"]
            nama_safe = sanitize_filename(nama_raw)

            pdf_path = _find_pdf_for_pegawai(dir_output, nama_safe)

            if pdf_path:
                log(f"    ✔ {nama_raw} ({jabatan})")
                semua_pdf.append(pdf_path)
                total_cocok += 1
            else:
                log(f"    ⚠ PDF tidak ditemukan untuk: {nama_raw}")

    # --- Merge semua PDF menjadi 1 file ---
    if not semua_pdf:
        log("\n❌ Tidak ada PDF yang bisa digabungkan. Periksa data jabatan dan file PDF.")
        return

    nama_file_output = f"{NAMA_FILE_OUTPUT}_{periode_str}.pdf"
    output_path      = os.path.join(dir_output, nama_file_output)

    log(f"\n  Menggabungkan {total_cocok} PDF menjadi satu file...")

    try:
        merge_pdf(semua_pdf, output_path)
        log(f"✅ Berhasil! File tersimpan → {nama_file_output}")
    except Exception as e:
        log(f"❌ Gagal merge PDF: {e}")
