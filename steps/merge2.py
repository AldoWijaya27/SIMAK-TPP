from logging import log
import os
import pandas as pd
import pythoncom
import win32com.client as win32
import urllib.parse

from utils import sanitize_filename, find_rekap_pdf, merge_pdf


def process_mail_merge(DIR_REKAP, DIR_OUTPUT, TEMP_DIR, csv_file, template_word, log):

    pythoncom.CoInitialize()

    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        template_word = urllib.parse.unquote(template_word)
        template_word = os.path.normpath(template_word)

        if not os.path.exists(template_word):
            raise Exception(f"Template tidak ditemukan: {template_word}")

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

        for idx in range(total_records):
            row = df.iloc[idx]

            nama = sanitize_filename(str(row["NAMA"]))
            bidang = str(row["BIDANG"])

            log(f"Processing {nama}")

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


        # =============================
        # 6. CLEANUP
        # =============================
        doc.Close(False)
        word.Quit()

        log("Semua proses mail merge selesai.")

    finally:
        pythoncom.CoUninitialize()
