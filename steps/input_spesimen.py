import fitz
import os

def input_spesimen(dir_rekap, dir_rekap_tandatangan, log):

    # lokasi file ini (steps/)
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # naik ke folder Aplikasi
    APP_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))

    # arahkan ke folder files
    TTD_IMAGE = os.path.join(APP_DIR, "files", "spesimen.png")

    os.makedirs(dir_rekap_tandatangan, exist_ok=True)

    total_file = 0

    # 🔥 Loop semua subfolder
    for root, dirs, files in os.walk(dir_rekap):
        for filename in files:
            if not filename.lower().endswith(".pdf"):
                continue

            total_file += 1

            input_path = os.path.join(root, filename)

            # 🔥 Jaga struktur folder output
            relative_path = os.path.relpath(root, dir_rekap)
            output_dir = os.path.join(dir_rekap_tandatangan, relative_path)
            os.makedirs(output_dir, exist_ok=True)

            output_path = os.path.join(output_dir, filename)

            log(f"Proses: {filename}")

            doc = fitz.open(input_path)

            for page in doc:
                areas = page.search_for("KEPALA DINAS KEHUTANAN")

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

            doc.save(output_path)
            doc.close()

    log(f"✔ Selesai. Total file diproses: {total_file}")