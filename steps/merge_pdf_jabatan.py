import os
import glob
import pandas as pd

from utils import merge_pdf, sanitize_filename


# ============================================================
# DEFINISI KELOMPOK BERDASARKAN JABATAN PIMPINAN
# Setiap entry: (nama_kelompok, nilai_jabatan_pimpinan, nama_file_output)
# Pencocokan: case-insensitive exact match pada kolom JABATAN PIMPINAN.
# ============================================================
KELOMPOK_PIMPINAN = [
    ("Kepala Dinas", "Plt. KEPALA DINAS KEHUTANAN", "Merge_Pejabat_KaDis"),
    ("Sekretaris",   "SEKRETARIS",                   "Merge_Pejabat_Sekretaris"),
]

# Legacy constants — dipertahankan agar split_pdf_jabatan.py tidak error import
KELOMPOK_JABATAN = [
    ("Kepala Bidang", ["kepala bidang"]),
    ("Kepala UPTD",   ["kepala uptd"]),
    ("Sekretaris",    ["sekretaris"]),
    ("Madya",         ["madya"]),
]
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
    Gabungkan PDF berdasarkan JABATAN PIMPINAN menjadi file-file terpisah.

    Output:
      - Merge_Pejabat_KaDis_{bulan}_{tahun}.pdf
        → pegawai yang pimpinannya Plt. KEPALA DINAS KEHUTANAN
      - Merge_Pejabat_Sekretaris_{bulan}_{tahun}.pdf
        → pegawai yang pimpinannya SEKRETARIS
    """
    log("── Memulai proses Merge PDF per Jabatan Pimpinan ──")

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

    col_jp = "JABATAN PIMPINAN"
    if col_jp not in df.columns:
        log(f"❌ Kolom '{col_jp}' tidak ditemukan di CSV.")
        return

    # Bersihkan NaN
    df["NAMA"] = df["NAMA"].fillna("").str.strip()
    df[col_jp] = df[col_jp].fillna("").str.strip()
    if "JABATAN" in df.columns:
        df["JABATAN"] = df["JABATAN"].fillna("").str.strip()

    # Filter baris yang punya nama
    df = df[df["NAMA"] != ""]

    log(f"  Total pegawai di CSV: {len(df)}")

    dir_output  = os.path.abspath(dir_output)
    periode_str = f"{bulan}_{tahun}"

    total_berhasil = 0

    # --- Proses setiap kelompok pimpinan ---
    for nama_kelompok, nilai_jp, nama_file_base in KELOMPOK_PIMPINAN:
        from steps.download import stop_event
        if stop_event.is_set():
            log("Proses Merge PDF per Jabatan Pimpinan dibatalkan oleh pengguna.")
            raise SystemExit()

        # Filter: case-insensitive exact match pada JABATAN PIMPINAN
        pegawai_cocok = df[df[col_jp].str.upper() == nilai_jp.upper()].copy()

        if pegawai_cocok.empty:
            log(f"  ⚠ Tidak ada pegawai dengan pimpinan '{nama_kelompok}'. Dilewati.")
            continue

        pegawai_cocok = pegawai_cocok.sort_values("NAMA")
        log(f"\n  📂 Pimpinan: {nama_kelompok} ({len(pegawai_cocok)} pegawai):")

        semua_pdf = []

        for _, row in pegawai_cocok.iterrows():
            if stop_event.is_set():
                log("Proses Merge PDF per Jabatan Pimpinan dibatalkan oleh pengguna.")
                raise SystemExit()

            nama_raw  = row["NAMA"]
            nama_safe = sanitize_filename(nama_raw)

            jabatan = row.get("JABATAN", "") if "JABATAN" in df.columns else ""
            if pd.isna(jabatan):
                jabatan = ""

            pdf_path = _find_pdf_for_pegawai(dir_output, nama_safe)

            if pdf_path:
                log(f"    ✔ {nama_raw} ({jabatan})")
                semua_pdf.append(pdf_path)
            else:
                log(f"    ⚠ PDF tidak ditemukan untuk: {nama_raw}")

        # --- Merge PDF untuk kelompok ini ---
        if not semua_pdf:
            log(f"\n  ❌ Tidak ada PDF yang bisa digabungkan untuk kelompok '{nama_kelompok}'.")
            continue

        nama_file_output = f"{nama_file_base}_{periode_str}.pdf"
        output_path = os.path.join(dir_output, nama_file_output)

        log(f"\n  Menggabungkan {len(semua_pdf)} PDF menjadi {nama_file_output}...")

        try:
            merge_pdf(semua_pdf, output_path)
            log(f"  ✅ Berhasil! → {nama_file_output}")
            total_berhasil += len(semua_pdf)
        except Exception as e:
            log(f"  ❌ Gagal merge PDF untuk '{nama_kelompok}': {e}")

    log(f"\n✅ Proses Merge PDF per Jabatan Pimpinan selesai. Total {total_berhasil} dokumen digabungkan.")
