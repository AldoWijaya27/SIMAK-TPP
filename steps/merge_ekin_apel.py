import pandas as pd
from openpyxl import load_workbook
import shutil


def merge_ekin_apel(file_sumber, file_utama, output_path):
    try:
        shutil.copy(file_utama, output_path)
        df_sumber = pd.read_excel(file_sumber)
        df_sumber.columns = df_sumber.columns.str.strip()

        required_cols = ["NIP", "Predikat Kinerja", "TMA", 'TMA Lain']
        for col in required_cols:
            if col not in df_sumber.columns:
                raise ValueError(f"Kolom '{col}' tidak ditemukan di file sumber.")

        # Bersihkan NIP
        df_sumber["NIP"] = (
            df_sumber["NIP"]
            .astype(str)
            .str.replace(" ", "", regex=False)
            .str.strip()
        )

        # Buat mapping berdasarkan NIP
        mapping = df_sumber.set_index("NIP")[["Predikat Kinerja", "TMA", "TMA Lain"]].to_dict("index")

        # ==========================================================
        # 3️⃣ Load file output pakai openpyxl (bukan pandas!)
        # ==========================================================
        wb = load_workbook(output_path)
        ws = wb["DISIPLIN"]  # jika sheet tertentu, bisa diganti wb["NamaSheet"]

        # ==========================================================
        # 4️⃣ Identifikasi posisi kolom berdasarkan header
        # ==========================================================
        header_row = 1
        headers = {}

        for col in range(1, ws.max_column + 1):
            header_value = ws.cell(row=header_row, column=col).value
            if header_value:
                headers[str(header_value).strip()] = col

        if "NIP" not in headers:
            raise ValueError("Kolom 'NIP' tidak ditemukan di file utama.")

        if "Predikat Kinerja" not in headers:
            raise ValueError("Kolom 'Predikat Kinerja' tidak ditemukan di file utama.")

        if "TMA" not in headers:
            raise ValueError("Kolom 'TMA' tidak ditemukan di file utama.")
        
        if "TMA Lain" not in headers:
            raise ValueError("Kolom 'TMA Lain' tidak ditemukan di file utama.")

        col_nip = headers["NIP"]
        col_predikat = headers["Predikat Kinerja"]
        col_tma = headers["TMA"]
        col_tma_lain = headers["TMA Lain"]

        # ==========================================================
        # 5️⃣ Update nilai berdasarkan NIP (tanpa merusak rumus lain)
        # ==========================================================
        updated_count = 0

        for row in range(2, ws.max_row + 1):
            nip_cell = ws.cell(row=row, column=col_nip).value

            if nip_cell is None:
                continue

            nip_clean = str(nip_cell).replace(" ", "").strip()

            if nip_clean in mapping:
                ws.cell(row=row, column=col_predikat).value = mapping[nip_clean]["Predikat Kinerja"]
                ws.cell(row=row, column=col_tma).value = mapping[nip_clean]["TMA"]
                ws.cell(row=row, column=col_tma_lain).value = mapping[nip_clean]["TMA Lain"]
                updated_count += 1

        # ==========================================================
        # 6️⃣ Simpan file (struktur tetap aman)
        # ==========================================================
        wb.save(output_path)

        return True, f"Berhasil update {updated_count} data berdasarkan NIP."

    except Exception as e:
        return False, str(e)
