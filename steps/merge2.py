from logging import log
import os
import time
import pandas as pd
import urllib.parse
import platform
import shlex

try:
    import pythoncom
    import win32com.client as win32
    WINDOWS_COM_AVAILABLE = True
except ImportError:
    WINDOWS_COM_AVAILABLE = False

try:
    from mailmerge import MailMerge
    MAILMERGE_AVAILABLE = True
except ImportError:
    MAILMERGE_AVAILABLE = False

from utils import sanitize_filename, find_rekap_pdf, merge_pdf

# Monkeypatch shlex.split to handle unclosed quotes in Word mergefield formatting instructions
original_split = shlex.split
def patched_split(s, *args, **kwargs):
    if s.count('"') % 2 != 0:
        s = s.replace('"', '')
    if s.count("'") % 2 != 0:
        s = s.replace("'", "")
    return original_split(s, *args, **kwargs)

shlex.split = patched_split


def _run_windows_mail_merge(DIR_REKAP, DIR_OUTPUT, TEMP_DIR, csv_file, template_word, log):
    pythoncom.CoInitialize()

    try:
        try:
            word = win32.Dispatch("Word.Application")
        except Exception:
            try:
                # Coba gunakan WPS Office jika Microsoft Word tidak ada
                word = win32.Dispatch("KWps.Application")
            except Exception:
                raise Exception("Aplikasi Microsoft Word atau WPS Office tidak ditemukan atau belum disetting default. Fitur Mail Merge ini wajib membutuhkan MS Word/WPS Office.")
        
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(template_word)
        doc.MailMerge.OpenDataSource(
            Name=csv_file,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Revert=False,
            Format=0,
        )

        df = pd.read_csv(csv_file, sep=";", encoding="utf-8")
        df.columns = df.columns.str.strip().str.upper()
        total_records = len(df)

        log(f"Total record CSV: {total_records}")

        t_start = time.time()

        for idx in range(total_records):
            from steps.download import stop_event
            if stop_event.is_set():
                doc.Close(False)
                word.Quit()
                log("Proses Mail Merge dibatalkan oleh pengguna.")
                raise SystemExit()

            row = df.iloc[idx]

            nama = sanitize_filename(str(row["NAMA"]))
            bidang = str(row["BIDANG"])

            # Estimasi waktu tersisa
            elapsed = time.time() - t_start
            if idx > 0:
                avg_per_item = elapsed / idx
                sisa = avg_per_item * (total_records - idx)
                if sisa >= 60:
                    est_text = f"~{int(sisa // 60)} menit {int(sisa % 60)} detik tersisa"
                else:
                    est_text = f"~{int(sisa)} detik tersisa"
            else:
                est_text = "menghitung..."

            log(f"  [{idx + 1}/{total_records}] {nama} — {est_text}")

            doc.MailMerge.Destination = 0  # New document
            doc.MailMerge.DataSource.FirstRecord = idx + 1
            doc.MailMerge.DataSource.LastRecord = idx + 1
            doc.MailMerge.Execute(False)

            result = word.ActiveDocument

            temp_pdf = os.path.abspath(os.path.join(TEMP_DIR, f"{nama}.pdf"))

            result.ExportAsFixedFormat(
                OutputFileName=temp_pdf,
                ExportFormat=17
            )

            result.Close(False)

            rekap = find_rekap_pdf(DIR_REKAP, bidang, nama)

            final_folder = os.path.join(DIR_OUTPUT, bidang)
            os.makedirs(final_folder, exist_ok=True)

            final = os.path.join(final_folder, f"{nama}.pdf")

            merge_pdf([temp_pdf, rekap], final)

            log(f"Selesai: {final}")

        doc.Close(False)
        word.Quit()
        log("Semua proses mail merge selesai.")

    finally:
        pythoncom.CoUninitialize()


def _run_macos_mail_merge(DIR_REKAP, DIR_OUTPUT, TEMP_DIR, csv_file, template_word, log):
    log("[MACOS] Memulai proses Mail Merge nyata via docx-mailmerge2 + LibreOffice...")
    
    if not MAILMERGE_AVAILABLE:
        raise Exception("Pustaka 'docx-mailmerge2' tidak terinstal di macOS. Silakan jalankan 'pip install docx-mailmerge2'.")
        
    soffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not os.path.exists(soffice_path):
        raise Exception("LibreOffice tidak ditemukan di /Applications/LibreOffice.app. Silakan instal LibreOffice terlebih dahulu.")

    # Baca data CSV dengan dtype=str untuk menjaga format teks
    try:
        df = pd.read_csv(csv_file, sep=";", encoding="utf-8", dtype=str)
    except Exception as e:
        raise Exception(f"Gagal membaca CSV: {e}")
        
    df.columns = df.columns.str.strip().str.upper()
    total_records = len(df)
    log(f"[MACOS] Total record CSV: {total_records}")

    t_start = time.time()
    loop_idx = 0

    for idx, row in df.iterrows():
        from steps.download import stop_event
        if stop_event.is_set():
            log("Proses Mail Merge dibatalkan oleh pengguna.")
            raise SystemExit()

        # Bersihkan NaN
        row_dict = row.to_dict()
        row_dict = {k: (v if pd.notna(v) else "") for k, v in row_dict.items()}

        # --- Normalisasi key CSV agar cocok dengan merge field di template ---
        # Masalah: kolom CSV pakai spasi (dari Excel), tapi merge field di Word
        # pakai underscore. Selain itu, casing-nya juga bisa beda.
        # Solusi: baca merge field dari template, lalu cocokkan via case-insensitive.
        import re

        # Bangun mapping: key CSV yang dinormalisasi -> merge field asli di template
        if not hasattr(_run_macos_mail_merge, '_template_fields'):
            with MailMerge(template_word) as tmp_doc:
                _run_macos_mail_merge._template_fields = tmp_doc.get_merge_fields()

        template_fields = _run_macos_mail_merge._template_fields

        # Buat lookup dari template fields: lowercase -> nama asli
        field_lookup = {f.lower(): f for f in template_fields}

        mapped_dict = {}
        for csv_key, val in row_dict.items():
            # Normalisasi key CSV: ganti karakter non-alfanumerik jadi underscore
            norm_key = re.sub(r'[^A-Za-z0-9_]', '_', csv_key)
            norm_key = re.sub(r'_+', '_', norm_key)
            norm_key = norm_key.strip('_')

            # Coba cocokkan (case-insensitive) ke merge field template
            real_field = field_lookup.get(norm_key.lower())
            if real_field:
                mapped_dict[real_field] = val
            else:
                # Coba juga dengan trailing underscore (untuk kasus "Skor Kehadiran (%)")
                real_field_trail = field_lookup.get((norm_key + '_').lower())
                if real_field_trail:
                    mapped_dict[real_field_trail] = val
                else:
                    # Simpan apa adanya untuk field lain (NAMA, BIDANG, dll)
                    mapped_dict[csv_key] = val

        # --- Mapping manual untuk edge case ---
        # Word mengkonversi "< 15 menit" menjadi "M__15_menit" dst.
        _manual_map = {
            "< 15 menit": "M__15_menit",
            "< 30 menit": "M__30_menit",
            "< 60 menit": "M__60_menit",
            "> 60 menit": "M__60_menit1",
        }
        for csv_col, merge_field in _manual_map.items():
            # Cari kolom CSV (case-insensitive)
            for orig_key, val in row.to_dict().items():
                if orig_key.strip().lower() == csv_col.lower():
                    if merge_field in template_fields:
                        mapped_dict[merge_field] = val if pd.notna(val) else ""
                    break

        # Duplikat field: template punya "Bulan" dan "BULAN" yang merujuk data sama
        if "Bulan" in template_fields and "Bulan" not in mapped_dict:
            bulan_val = mapped_dict.get("BULAN", "")
            mapped_dict["Bulan"] = bulan_val

        row_dict = mapped_dict

        nama = sanitize_filename(str(row_dict.get("NAMA", f"Pegawai_{idx+1}")))
        bidang = str(row_dict.get("BIDANG", "Umum"))

        # Estimasi waktu tersisa
        elapsed = time.time() - t_start
        if loop_idx > 0:
            avg_per_item = elapsed / loop_idx
            sisa = avg_per_item * (total_records - loop_idx)
            if sisa >= 60:
                est_text = f"~{int(sisa // 60)} menit {int(sisa % 60)} detik tersisa"
            else:
                est_text = f"~{int(sisa)} detik tersisa"
        else:
            est_text = "menghitung..."

        log(f"  [{loop_idx + 1}/{total_records}] {nama} ({bidang}) — {est_text}")
        loop_idx += 1

        # 1. Jalankan mail merge di memory dan simpan sebagai docx sementara
        temp_docx = os.path.abspath(os.path.join(TEMP_DIR, f"{nama}.docx"))
        try:
            with MailMerge(template_word) as document:
                document.merge(**row_dict)
                document.write(temp_docx)
        except Exception as e:
            log(f"    ❌ Gagal Mail Merge docx untuk {nama}: {e}")
            continue

        # 2. Konversi docx sementara menjadi PDF via LibreOffice
        temp_pdf = os.path.abspath(os.path.join(TEMP_DIR, f"{nama}.pdf"))
        if os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except OSError:
                pass

        import subprocess
        cmd = [
            soffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            TEMP_DIR,
            temp_docx
        ]
        
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            log(f"    ❌ Gagal konversi PDF via LibreOffice untuk {nama}: {e}")
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except OSError:
                    pass
            continue

        if os.path.exists(temp_docx):
            try:
                os.remove(temp_docx)
            except OSError:
                pass

        # 3. Cari file rekap PDF
        rekap = find_rekap_pdf(DIR_REKAP, bidang, nama)

        # 4. Gabungkan PDF hasil mail merge dengan PDF rekap kehadiran
        final_folder = os.path.join(DIR_OUTPUT, bidang)
        os.makedirs(final_folder, exist_ok=True)
        final = os.path.join(final_folder, f"{nama}.pdf")

        try:
            merge_pdf([temp_pdf, rekap], final)
            log(f"    ✔ Selesai: {final}")
        except Exception as e:
            log(f"    ❌ Gagal menggabungkan PDF untuk {nama}: {e}")
        finally:
            if os.path.exists(temp_pdf):
                try:
                    os.remove(temp_pdf)
                except OSError:
                    pass

    log("[MACOS] Semua proses mail merge selesai.")


def process_mail_merge(DIR_REKAP, DIR_OUTPUT, TEMP_DIR, csv_file, template_word, log):
    template_word = urllib.parse.unquote(template_word)
    template_word = os.path.normpath(template_word)
    
    if not os.path.exists(template_word):
        raise Exception(f"Template tidak ditemukan: {template_word}")

    if platform.system() != "Darwin" and WINDOWS_COM_AVAILABLE:
        _run_windows_mail_merge(DIR_REKAP, DIR_OUTPUT, TEMP_DIR, csv_file, template_word, log)
    else:
        _run_macos_mail_merge(DIR_REKAP, DIR_OUTPUT, TEMP_DIR, csv_file, template_word, log)
