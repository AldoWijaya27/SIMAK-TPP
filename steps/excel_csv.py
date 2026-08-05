import os
import pandas as pd

def recalculate_disiplin_sheet(ws, write_formulas=True):
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
        if isinstance(v, (int, float)):
            return float(v)

        s = str(v).strip()
        import re
        s = re.sub(r"(?i)rp\.?\s*", "", s)
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "." in s:
            parts = s.split(".")
            if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3 and parts[0].isdigit()):
                s = s.replace(".", "")
        elif "," in s:
            s = s.replace(",", ".")

        try:
            return float(s)
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
        if "sangat kurang" in predikat:
            persen_kinerja = 40
        elif "kurang" in predikat:
            persen_kinerja = 60
        elif "butuh" in predikat or "perbaikan" in predikat:
            persen_kinerja = 80
        elif "sangat baik" in predikat or "baik" in predikat:
            persen_kinerja = 100
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

        if write_formulas:
            set_val(r, "TMK Persen", f"=O{r}*2")
            set_val(r, "TMA Persen", f"=Q{r}*2")
            set_val(r, "TMA Lain Persen", f"=S{r}*5")
            set_val(r, "Persen a", f"=U{r}*0.25")
            set_val(r, "Persen b", f"=W{r}*0.5")
            set_val(r, "Persen c", f"=Y{r}*1")
            set_val(r, "Persen d", f"=AA{r}*2.5")
            set_val(r, "Skor Tidak Disiplin", f"=R{r}+T{r}+V{r}+X{r}+Z{r}+AB{r}")
            set_val(r, "persentase hadir", f"=IF(AD{r}>0,(AE{r}/AD{r})*100,100)")
            set_val(r, "Skor Kehadiran (%)", f"=(100-AC{r})*0.4")
            set_val(r, "Skor Kinerja (%)", f"=AG{r}*0.6")
            set_val(r, "Skor Total (%)", f"=AI{r}+AJ{r}")
            set_val(r, "Persentase Total", f"=AK{r}-P{r}")
            set_val(r, "jumlah TP", f"=ROUND((AM{r}/100)*AL{r},0)")
            set_val(r, "TPP Kotor", f"=AN{r}+AO{r}")
            set_val(r, "jumlah bersih", f"=AP{r}-AQ{r}-AR{r}")
            set_val(r, "TPP bersih", f"=AS{r}-AT{r}")
            set_val(r, "pengurangan disiplin", f"=AL{r}-AN{r}")
            set_val(r, "jumlah potongan", f"=AQ{r}+AR{r}+AT{r}+AV{r}")


def format_rupiah(val):
    if val is None or pd.isna(val) or str(val).strip() == "":
        return "0"
    try:
        s = str(val).strip()
        import re
        s = re.sub(r"(?i)rp\.?\s*", "", s)
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "." in s:
            parts = s.split(".")
            if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3 and parts[0].isdigit()):
                s = s.replace(".", "")
        elif "," in s:
            s = s.replace(",", ".")

        n = int(round(float(s)))
        return f"{n:,}".replace(",", ".")
    except (ValueError, TypeError):
        return str(val)


def format_persen(val):
    if val is None or pd.isna(val) or str(val).strip() == "":
        return "0,00"
    try:
        s = str(val).strip().replace("%", "")
        import re
        s = re.sub(r"(?i)rp\.?\s*", "", s)
        f = float(s)
        formatted = f"{f:.2f}".replace(".", ",")
        return formatted
    except (ValueError, TypeError):
        return str(val)


def release_file_lock_if_needed(file_path):
    """
    Jika file sedang terkunci di Windows oleh WINWORD.EXE atau wps.exe (sisa Mail Merge sebelumnya),
    secara otomatis mematikan proses background Word/WPS agar file dapat ditulis ulang.
    """
    import sys, os
    if sys.platform != "win32" or not os.path.exists(file_path):
        return
    import subprocess, time
    try:
        with open(file_path, "a"):
            pass
        return
    except IOError:
        pass

    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "WINWORD.EXE"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=0x08000000
        )
        subprocess.run(
            ["taskkill", "/F", "/IM", "kwps.exe"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=0x08000000
        )
        subprocess.run(
            ["taskkill", "/F", "/IM", "wps.exe"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=0x08000000
        )
        time.sleep(0.5)
    except Exception:
        pass


def excel_sheet_disiplin_ke_csv(file_excel, output_folder):
    file_excel = os.path.abspath(file_excel)
    output_folder = os.path.abspath(output_folder)

    filename = os.path.splitext(os.path.basename(file_excel))[0] + ".csv"
    output_csv = os.path.join(output_folder, filename)

    # Lepaskan penguncian file jika ada proses background Word yang masih memegang file
    release_file_lock_if_needed(file_excel)
    release_file_lock_if_needed(output_csv)

    from openpyxl import load_workbook
    wb = load_workbook(file_excel, data_only=False)
    sheet_name = "DISIPLIN" if "DISIPLIN" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    # 1. Hitung nilai numerik terlebih dahulu untuk ekspor ke DataFrame / CSV
    recalculate_disiplin_sheet(ws, write_formulas=False)

    # Evaluasi kolom TTE secara presisi per baris berdasarkan Atasan / Penandatangan TTE
    tte_col = None
    nip_p_col = None
    jab_p_col = None
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(1, c).value or "").strip().lower()
        if h == "tte":
            tte_col = c
        elif h == "nip pimpinan":
            nip_p_col = c
        elif h == "jabatan pimpinan":
            jab_p_col = c

    if tte_col:
        auth_nips = {"196902051989101002", "198204182006042012"}
        for r in range(2, ws.max_row + 1):
            nip_p = str(ws.cell(r, nip_p_col).value or "").replace(" ", "").strip() if nip_p_col else ""
            jab_p = str(ws.cell(r, jab_p_col).value or "").strip().upper() if jab_p_col else ""

            is_auth = (nip_p in auth_nips) or (("KEPALA DINAS" in jab_p or "SEKRETARIS" in jab_p) and "SEKRETARIS DAERAH" not in jab_p)
            if is_auth:
                ws.cell(r, tte_col).value = "${ttd_pengirim}"
            else:
                ws.cell(r, tte_col).value = ""

    # Ambil nilai numerik untuk DataFrame CSV
    data_rows = list(ws.values)
    headers_list = data_rows[0] if data_rows else []
    df = pd.DataFrame(data_rows[1:], columns=headers_list) if len(data_rows) > 1 else pd.DataFrame()

    # 2. Tulis rumus Excel asli (live formulas) ke worksheet untuk disimpan sebagai file .xlsx
    recalculate_disiplin_sheet(ws, write_formulas=True)

    try:
        wb.save(file_excel)
    except PermissionError:
        release_file_lock_if_needed(file_excel)
        try:
            wb.save(file_excel)
        except PermissionError:
            raise Exception(f"File Excel '{os.path.basename(file_excel)}' sedang dibuka di Microsoft Excel atau aplikasi lain. Silakan tutup file tersebut terlebih dahulu.")
    finally:
        wb.close()

    # Format kolom nominal rupiah dengan pemisah titik ribuan (misal 4.949.000)
    currency_cols = [
        "TPP Asli", "jumlah TP", "Tambahan 20%", "TPP Kotor",
        "PPh21", "BPJS", "jumlah bersih", "zakat", "TPP bersih",
        "pengurangan disiplin", "jumlah potongan"
    ]
    for col in currency_cols:
        if col in df.columns:
            df[col] = df[col].apply(format_rupiah)

    # Format persentase dengan 2 desimal koma (misal 60,00%, 40,00%, 2,00%, 98,00%)
    pct_cols = [
        "Skor Kinerja (%)", "Skor Kehadiran (%)", "TMK Persen",
        "Persentase Total", "Skor Total (%)", "Persentase Kinerja",
        "TMA Persen", "TMA Lain Persen", "Persen a", "Persen b",
        "Persen c", "Persen d", "persentase hadir", "Skor Tidak Disiplin"
    ]
    for col in pct_cols:
        if col in df.columns:
            df[col] = df[col].apply(format_persen)

    if "tte" in df.columns:
        df["tte"] = df["tte"].fillna("")

    try:
        df.to_csv(output_csv, sep=';', encoding='utf-8', index=False)
    except PermissionError:
        release_file_lock_if_needed(output_csv)
        try:
            df.to_csv(output_csv, sep=';', encoding='utf-8', index=False)
        except PermissionError:
            raise Exception(f"File CSV '{os.path.basename(output_csv)}' sedang dibuka di Microsoft Excel atau aplikasi lain. Silakan tutup jendela Excel/Word Anda terlebih dahulu lalu coba lagi.")
