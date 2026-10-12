import os
import fitz

from steps.download import stop_event

def process_pdfs(folder_path, log=print):
    """
    Menambahkan placeholder teks '${ttd_pengirim}' untuk TTE Srikandi
    pada dokumen SKP hasil unduhan di posisi 'Pejabat Penilai Kinerja',
    serta menghasilkan berkas gabungan kompresi untuk Kadis & Sekdis.
    """
    if not os.path.exists(folder_path):
        log(f"❌ Folder target tidak ditemukan: {folder_path}")
        return

    # Dapatkan semua file PDF di folder (kecuali file gabungan jika sudah ada)
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf') and not f.startswith('gabungan_')]
    pdf_files.sort()

    if not pdf_files:
        log(f"⚠️ Tidak ada file PDF ditemukan di {folder_path}")
        return

    log(f"[*] Ditemukan {len(pdf_files)} file PDF SKP. Memulai penambahan placeholder TTD...")

    sukses = 0
    merged_doc_kadis = fitz.open()
    merged_doc_sekdis = fitz.open()

    for file_name in pdf_files:
        if stop_event.is_set():
            log("Proses penambahan placeholder TTD dibatalkan oleh pengguna.")
            break

        file_path = os.path.join(folder_path, file_name)
        try:
            doc = fitz.open(file_path)

            # Cari dari halaman terakhir ke halaman pertama
            found = False
            for page_num in range(len(doc) - 1, -1, -1):
                page = doc[page_num]

                # Cari teks Pejabat Penilai Kinerja
                rects = page.search_for("Pejabat Penilai Kinerja")

                if rects:
                    rect = rects[-1]
                    y_insert = rect.y1 + 35
                    x_insert = rect.x0 + 40

                    page.insert_text(
                        fitz.Point(x_insert, y_insert),
                        "${ttd_pengirim}",
                        fontsize=8,
                        fontname="helv",
                        color=(0, 0, 0)
                    )

                    found = True
                    break

            # Masukkan ke dalam merged document yang sesuai
            if "KADIS" in file_name.upper():
                merged_doc_kadis.insert_pdf(doc)
            elif "SEKDIS" in file_name.upper() or "SEKRETARIS" in file_name.upper():
                merged_doc_sekdis.insert_pdf(doc)

            if found:
                temp_path = file_path + ".tmp"
                doc.save(temp_path)
                doc.close()
                os.replace(temp_path, file_path)
                sukses += 1
                log(f"  [v] Placeholder TTD ditambahkan: {file_name}")
            else:
                log(f"  [x] Teks 'Pejabat Penilai Kinerja' tidak ditemukan di: {file_name}")
                doc.close()

        except Exception as e:
            log(f"  ❌ Error memproses {file_name}: {e}")

    log(f"✅ Selesai! Berhasil memodifikasi {sukses} dari {len(pdf_files)} file SKP.")

    # Simpan dokumen gabungan
    def simpan_gabungan(merged_doc, output_filename):
        if len(merged_doc) > 0:
            log(f"  Menggabungkan dan mengkompresi PDF menjadi {output_filename}...")
            output_path = os.path.join(folder_path, output_filename)
            merged_doc.save(
                output_path,
                garbage=4,
                deflate=True,
                clean=True
            )
            log(f"  ✔ File gabungan berhasil disimpan di: {output_path}")
        merged_doc.close()

    simpan_gabungan(merged_doc_kadis, "gabungan_skp_kadis_compressed.pdf")
    simpan_gabungan(merged_doc_sekdis, "gabungan_skp_sekdis_compressed.pdf")

if __name__ == "__main__":
    folder_target = "hasil_download_skp"
    process_pdfs(folder_target)
