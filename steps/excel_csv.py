import os
import pandas as pd

def recalculate_disiplin_sheet(ws):
    headers = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(1, col).value
        if val:
            headers[str(val).strip().lower()] = col

    def get_val(row, name, default=0):
        c = headers.get(name.lower())
        if not c:
            return default
        v = ws.cell(row, c).value
        if v is None or str(v).startswith("="):
            return default
        try:
            return float(v)
        except (ValueError, TypeError):
            return str(v)

    def set_val(row, name, val):
        c = headers.get(name.lower())
        if c:
            if isinstance(val, float):
                val = round(val, 2)
                if val.is_integer():
                    val = int(val)
            ws.cell(row, c).value = val

    for r in range(2, ws.max_row + 1):
        nama = get_val(r, "NAMA", "")
        if not nama or str(nama).strip() == "":
            continue

        tmk = get_val(r, "TMK", 0)
        tmk_persen = tmk * 2
        set_val(r, "TMK Persen", tmk_persen)

        tma = get_val(r, "TMA", 0)
        tma_persen = tma * 2
        set_val(r, "TMA Persen", tma_persen)

        tma_lain = get_val(r, "TMA Lain", 0)
        tma_lain_persen = tma_lain * 5
        set_val(r, "TMA Lain Persen", tma_lain_persen)

        p15 = get_val(r, "< 15 menit", 0)
        persen_a = p15 * 0.25
        set_val(r, "Persen a", persen_a)

        p30 = get_val(r, "< 30 menit", 0)
        persen_b = p30 * 0.5
        set_val(r, "Persen b", persen_b)

        p60 = get_val(r, "< 60 menit", 0)
        persen_c = p60 * 1.0
        set_val(r, "Persen c", persen_c)

        p60_gt = get_val(r, "> 60 menit", 0)
        persen_d = p60_gt * 2.5
        set_val(r, "Persen d", persen_d)

        skor_tdk_disiplin = tma_persen + tma_lain_persen + persen_a + persen_b + persen_c + persen_d
        set_val(r, "Skor Tidak Disiplin", skor_tdk_disiplin)

        total_hk = get_val(r, "Total HK", 0)
        hadir_hk = total_hk
        set_val(r, "Hadir HK", hadir_hk)

        persen_hadir = (hadir_hk / total_hk * 100) if total_hk > 0 else 100
        set_val(r, "persentase hadir", persen_hadir)

        predikat = str(get_val(r, "Predikat Kinerja", "Baik/Sangat Baik")).strip().lower()
        if any(w in predikat for w in ["sangat baik", "baik"]):
            persen_kinerja = 100
        elif "butuh" in predikat or "perbaikan" in predikat:
            persen_kinerja = 80
        elif "kurang" in predikat and "sangat" not in predikat:
            persen_kinerja = 60
        elif "sangat kurang" in predikat:
            persen_kinerja = 40
        else:
            persen_kinerja = 100
        set_val(r, "Persentase Kinerja", persen_kinerja)

        skor_kehadiran = (100 - skor_tdk_disiplin) * 0.40
        set_val(r, "Skor Kehadiran (%)", skor_kehadiran)

        skor_kinerja = persen_kinerja * 0.60
        set_val(r, "Skor Kinerja (%)", skor_kinerja)

        skor_total = skor_kehadiran + skor_kinerja
        set_val(r, "Skor Total (%)", skor_total)

        # Bulatkan rupiah ke integer utuh agar seluruh nominal TPP bersih tanpa koma desimal
        tpp_asli = round(get_val(r, "TPP Asli", 0))
        set_val(r, "TPP Asli", tpp_asli)

        persen_total = round(skor_total - tmk_persen, 2)
        set_val(r, "Persentase Total", persen_total)

        jumlah_tp = round((persen_total / 100.0) * tpp_asli)
        set_val(r, "jumlah TP", jumlah_tp)

        tambahan_20 = round(get_val(r, "Tambahan 20%", 0))
        tpp_kotor = jumlah_tp + tambahan_20
        set_val(r, "TPP Kotor", tpp_kotor)

        gol = str(get_val(r, "GOLONGAN", "")).strip().upper()
        if "IV" in gol:
            pph21 = round(tpp_kotor * 0.15)
        elif "III" in gol:
            pph21 = round(tpp_kotor * 0.05)
        else:
            pph21 = 0
        set_val(r, "PPh21", pph21)

        bpjs = round(get_val(r, "BPJS", 0))
        set_val(r, "BPJS", bpjs)

        jumlah_bersih = tpp_kotor - pph21 - bpjs
        set_val(r, "jumlah bersih", jumlah_bersih)

        if jumlah_bersih <= 0:
            zakat = 0
        else:
            jabatan = str(get_val(r, "JABATAN", "")).strip().lower()
            keywords_zakat = [
                "ahli pertama", "mahir", "penyelia", "terampil",
                "penelaah teknis", "penata kelola", "pengolah",
                "pengadministrasi", "operator"
            ]
            if any(kw in jabatan for kw in keywords_zakat):
                if "IV" in gol:
                    zakat = 50000
                elif "III" in gol:
                    zakat = 30000
                elif "II" in gol:
                    zakat = 20000
                else:
                    zakat = round(jumlah_bersih * 0.025)
            else:
                zakat = round(jumlah_bersih * 0.025)
        set_val(r, "zakat", zakat)

        tpp_bersih = jumlah_bersih - zakat
        set_val(r, "TPP bersih", tpp_bersih)

        pengurangan_disiplin = tpp_asli - jumlah_tp
        set_val(r, "pengurangan disiplin", pengurangan_disiplin)

        jumlah_potongan = pph21 + bpjs + zakat + pengurangan_disiplin
        set_val(r, "jumlah potongan", jumlah_potongan)


def excel_sheet_disiplin_ke_csv(file_excel, output_folder):
    file_excel = os.path.abspath(file_excel)
    output_folder = os.path.abspath(output_folder)

    filename = os.path.splitext(os.path.basename(file_excel))[0] + ".csv"
    output_csv = os.path.join(output_folder, filename)

    # Evaluasi seluruh rumus openpyxl dan simpan ke excel terlebih dahulu
    from openpyxl import load_workbook
    wb = load_workbook(file_excel, data_only=False)
    sheet_name = "DISIPLIN" if "DISIPLIN" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    recalculate_disiplin_sheet(ws)
    wb.save(file_excel)
    wb.close()

    # Baca sheet "DISIPLIN" dari excel dan ekspor ke CSV
    df = pd.read_excel(file_excel, sheet_name=sheet_name)
    df.to_csv(output_csv, sep=';', encoding='utf-8', index=False)
