import fitz
import os
from app_path import get_file

def input_spesimen(dir_rekap, dir_rekap_tandatangan, log):

    # arahkan ke folder files menggunakan get_file untuk mendukung mode frozen .exe
    TTD_IMAGE = get_file("files", "spesimen-compressed.png")
    if not os.path.exists(TTD_IMAGE):
        TTD_IMAGE = get_file("files", "spesimen.png")

    if not os.path.exists(TTD_IMAGE):
        log(f"❌ File gambar spesimen tanda tangan tidak ditemukan di folder 'files'. Pastikan 'files/spesimen-compressed.png' tersedia.")
        return

    os.makedirs(dir_rekap_tandatangan, exist_ok=True)

    total_file = 0

    # 🔥 Loop semua subfolder
    for root, dirs, files in os.walk(dir_rekap):
        for filename in files:
            if not filename.lower().endswith(".pdf"):
                continue

            total_file += 1

            from steps.download import stop_event
            if stop_event.is_set():
                log("Input spesimen dibatalkan oleh pengguna.")
                raise SystemExit()

            input_path = os.path.join(root, filename)

            # 🔥 Jaga struktur folder output
            relative_path = os.path.relpath(root, dir_rekap)
            output_dir = os.path.join(dir_rekap_tandatangan, relative_path)
            os.makedirs(output_dir, exist_ok=True)

            output_path = os.path.join(output_dir, filename)

            log(f"Proses: {filename}")

            try:
                keywords = ["KEPALA DINAS KEHUTANAN", "KEPALA DINAS", "PLT. KEPALA DINAS", "PLH. KEPALA DINAS"]

                for page in doc:
                    areas = []
                    for kw in keywords:
                        areas = page.search_for(kw)
                        if areas:
                            break

                    if not areas:
                        continue

                    # Ambil yang paling bawah (lebih aman)
                    rect = sorted(areas, key=lambda r: r.y1)[-1]

                    x0, y0, x1, y1 = rect

                    # 🔥 ukuran lebih kecil
                    ttd_width = 90
                    ttd_height = 45

                    # 🔥 posisi tengah
                    center_x = (x0 + x1) / 2

                    # 🔥 geser supaya center image pas di tengah teks
                    ttd_rect = fitz.Rect(
                        center_x - (ttd_width / 2),   # kiri
                        y1 + 3,                       # jarak lebih dekat
                        center_x + (ttd_width / 2),   # kanan
                        y1 + 5 + ttd_height           # bawah
                    )

                    page.insert_image(ttd_rect, filename=TTD_IMAGE)

                # Simpan dengan kompresi untuk memperkecil ukuran file
                # garbage=4: hapus objek tidak terpakai sepenuhnya
                # deflate=True: kompres stream konten
                # deflate_images=True: kompres gambar (termasuk spesimen)
                # deflate_fonts=True: kompres font
                # clean=True: bersihkan syntax PDF agar lebih padat
                doc.save(output_path,
                    garbage=4,
                    deflate=True,
                    deflate_images=True,
                    deflate_fonts=True,
                    clean=True
                )
            finally:
                doc.close()

    log(f"✔ Selesai. Total file diproses: {total_file}")