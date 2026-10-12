import requests
import time
import os
import pandas as pd
import concurrent.futures
from playwright.sync_api import sync_playwright

from steps.download import stop_event

# --- KONFIGURASI BAWAAN ---
LOKASI = "BANDAR LAMPUNG"
TANGGAL_PEGAWAI = "2026-10-01"
PERIODE_KONSTAN = "2026-490b2797ecac489cb01f6cf6712de843"

PERIODE_SKP_MAP = {
    "JANUARI": "b8bada1f688d42039644033e0a2a080d",
    "FEBRUARI": "5eb30cf66faf4a14a2635783244636ee",
    "MARET": "eadc309311d94d1c9b6d15d6730ed141",
    "APRIL": "42d3774d8dbe4ad0b56f6b8e054ec4f9",
    "MEI": "7a79e16dc8ef4776a8edb7c50fa94376",
    "JUNI": "feb0ac042e704052b67c4b7ba4bb2348",
    "JULI": "3d11a771e5114e3daef8c5a19dec8db2",
    "AGUSTUS": "fbbf789c6ba24ea38bf65a068c6130f6",
    "SEPTEMBER": "490b2797ecac489cb01f6cf6712de843",
    "OKTOBER": "4ff8e27d772846f9a9ae745485632292",
    "NOVEMBER": "4aa160f44f284dc89a4ad3e8f1f35c15",
    "DESEMBER": "1f4638b5a7a24046a2e85887254b4d8e",
    "PENILAIAN AKHIR": "dd11363dac0047e18ea9c3787200ec79",
}
# --------------------------

def get_app_dir():
    """Mendapatkan direktori dasar aplikasi."""
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def download_skp(skp_id, periode, output_filename, auth_token, label="", tanggal_pegawai=None, log=print):
    if stop_event.is_set():
        return False

    prefix = f"[{label}] " if label else ""
    tgl_peg = tanggal_pegawai if tanggal_pegawai else TANGGAL_PEGAWAI
    headers = {
        "Authorization": auth_token,
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        "Origin": "https://kinerja.bkn.go.id",
        "Referer": "https://kinerja.bkn.go.id/"
    }

    log(f"{prefix}[*] Memulai download SKP ID: {skp_id}")

    # 1. Request Cetak (Antre)
    url_cetak = f"https://kinerja.bkn.go.id/api/skp/{skp_id}/penilaian/{periode}/cetak"
    payload = {
        "lokasi": LOKASI,
        "periode": periode,
        "skp_id": skp_id,
        "tanggal_pegawai": tgl_peg
    }

    try:
        if stop_event.is_set():
            return False
        resp_cetak = requests.post(url_cetak, json=payload, headers=headers, timeout=30)
        if resp_cetak.status_code != 200:
            log(f"{prefix}❌ Error request cetak: {resp_cetak.text}")
            return False

        data_cetak = resp_cetak.json()
        job_id = data_cetak.get("job_id")
        if not job_id:
            log(f"{prefix}❌ Tidak mendapatkan job_id dari response!")
            return False

    except Exception as e:
        log(f"{prefix}❌ Terjadi kesalahan jaringan saat request cetak: {e}")
        return False

    # 2. Skip Polling Status (jeda singkat)
    time.sleep(3)
    if stop_event.is_set():
        return False

    # 3. Download Dokumen (Biner)
    url_download = f"https://kinerja.bkn.go.id/api/skp/cetak_download/{job_id}"
    try:
        resp_download = requests.get(url_download, headers=headers, timeout=60)

        if resp_download.status_code == 200:
            with open(output_filename, 'wb') as f:
                f.write(resp_download.content)
            log(f"{prefix}✔ Berhasil menyimpan: {os.path.basename(output_filename)}")
            return True
        else:
            log(f"{prefix}❌ Gagal download dokumen. Status code: {resp_download.status_code}")
            return False
    except Exception as e:
        log(f"{prefix}❌ Terjadi kesalahan saat download dokumen: {e}")
        return False

def dapatkan_token_otomatis(label, log=print):
    log(f"── Otentikasi Akun {label.upper()} ──")
    log(f"  Membuka browser Playwright untuk {label}...")
    log("  Jika belum login, silakan login di browser. Browser otomatis tertutup begitu token didapatkan.")

    token = None
    app_dir = get_app_dir()
    user_data_dir = os.path.join(app_dir, f"profil_chrome_{label.lower()}")

    # Fallback jika folder profil berada di root workspace / auto-download-ekin
    if not os.path.exists(user_data_dir):
        alt_path = os.path.join(os.path.dirname(app_dir), "auto-download-ekin", f"profil_chrome_{label.lower()}")
        if os.path.exists(alt_path):
            user_data_dir = alt_path

    with sync_playwright() as p:
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                viewport={'width': 1280, 'height': 720}
            )
        except Exception as e:
            log(f"❌ Gagal meluncurkan browser untuk {label}: {e}")
            return None

        page = context.pages[0] if context.pages else context.new_page()

        def cek_request(request):
            nonlocal token
            if "kinerja.bkn.go.id/api/" in request.url:
                auth = request.headers.get("authorization", "")
                if auth.startswith("Bearer ") and not token:
                    token = auth
                    log(f"  [+] Token {label} berhasil ditangkap secara otomatis!")

        page.on("request", cek_request)

        try:
            page.goto("https://kinerja.bkn.go.id/")

            waktu_tunggu = 0
            while not token and waktu_tunggu < 300:
                if stop_event.is_set():
                    log(f"  ⚠ Otentikasi {label} dibatalkan pengguna.")
                    break
                page.wait_for_timeout(1000)
                waktu_tunggu += 1

            if not token and not stop_event.is_set():
                log(f"  ❌ Gagal mendapatkan token untuk {label} dalam 5 menit.")

        except Exception as e:
            if not stop_event.is_set():
                log(f"  ❌ Error saat membuka halaman login: {e}")
        finally:
            try:
                context.close()
            except Exception:
                pass

    return token

def run_download_skp(excel_file, folder_output, log=print, periode_konstan=None, tanggal_pegawai=None):
    """
    Fungsi utama untuk menjalankan alur download SKP Kadis & Sekdis dari Excel.
    """
    from utils import sanitize_filename

    periode_final = periode_konstan if periode_konstan else PERIODE_KONSTAN
    tanggal_final = tanggal_pegawai if tanggal_pegawai else TANGGAL_PEGAWAI

    log("── Memulai Proses Download SKP E-Kinerja ──")
    log(f"  📌 Kode Periode SKP : {periode_final}")
    log(f"  📌 Tanggal Pegawai  : {tanggal_final}")

    if stop_event.is_set():
        return False

    auth_token_kadis = dapatkan_token_otomatis("KADIS", log=log)
    if stop_event.is_set():
        return False
    if not auth_token_kadis:
        log("❌ Token Kadis tidak didapatkan. Pembatalan proses.")
        return False

    auth_token_sekdis = dapatkan_token_otomatis("SEKRETARIS", log=log)
    if stop_event.is_set():
        return False
    if not auth_token_sekdis:
        log("❌ Token Sekretaris tidak didapatkan. Pembatalan proses.")
        return False

    os.makedirs(folder_output, exist_ok=True)

    log(f"[*] Membaca data pegawai dari: {os.path.basename(excel_file)}")
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        log(f"❌ Gagal membaca file Excel: {e}")
        return False

    # Standarisasi nama kolom case-insensitive
    col_map = {str(c).strip().lower(): c for c in df.columns}
    col_name = col_map.get("name") or col_map.get("nama")
    col_skp = col_map.get("skp_id")
    col_pimpinan = col_map.get("jab_pimpinan") or col_map.get("jabatan_pimpinan")

    if not col_name or not col_skp or not col_pimpinan:
        log("❌ Kolom yang dibutuhkan ('name/nama', 'skp_id', 'jab_pimpinan') tidak lengkap di file Excel!")
        return False

    antrean_kadis = []
    antrean_sekdis = []

    for _, row in df.iterrows():
        nama_raw = str(row.get(col_name, 'TanpaNama')).strip()
        nama = sanitize_filename(nama_raw)
        skp_id = row.get(col_skp)
        jab_pimpinan = str(row.get(col_pimpinan, ''))

        if pd.isna(skp_id) or not str(skp_id).strip():
            continue

        item = {'nama': nama, 'skp_id': str(skp_id).strip()}

        if 'KEPALA DINAS' in jab_pimpinan.upper():
            antrean_kadis.append(item)
        elif 'SEKRETARIS' in jab_pimpinan.upper():
            antrean_sekdis.append(item)

    log(f"  Antrean ditemukan:")
    log(f"    - Bawahan Kepala Dinas: {len(antrean_kadis)} orang")
    log(f"    - Bawahan Sekretaris: {len(antrean_sekdis)} orang")

    if not antrean_kadis and not antrean_sekdis:
        log("⚠️ Tidak ada antrean SKP dengan skp_id yang valid.")
        return True

    maks_percobaan = 3

    def proses_antrean(antrean, token, label):
        if not antrean or stop_event.is_set():
            return []

        log(f"[*] Memulai download untuk {label.upper()} ({len(antrean)} orang)...")
        antrean_berjalan = antrean
        percobaan = 1

        while antrean_berjalan and percobaan <= maks_percobaan:
            if stop_event.is_set():
                break

            if percobaan > 1:
                log(f"[*] Mengulang yang gagal - {label} (Percobaan {percobaan}/{maks_percobaan})...")
                time.sleep(5)

            antrean_gagal = []

            for bawahan in antrean_berjalan:
                if stop_event.is_set():
                    break

                nama = bawahan['nama']
                skp_id = bawahan['skp_id']
                filename = os.path.join(folder_output, f"SKP_{label}_{nama}_{periode_final}.pdf")

                berhasil = download_skp(skp_id, periode_final, filename, token, label=label, tanggal_pegawai=tanggal_final, log=log)
                if not berhasil:
                    antrean_gagal.append(bawahan)

            antrean_berjalan = antrean_gagal
            percobaan += 1

        return antrean_berjalan

    log("[*] Memulai download paralel untuk Kadis & Sekdis...")
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        future_kadis = executor.submit(proses_antrean, antrean_kadis, auth_token_kadis, "KADIS")
        future_sekdis = executor.submit(proses_antrean, antrean_sekdis, auth_token_sekdis, "SEKDIS")

        gagal_kadis = future_kadis.result()
        gagal_sekdis = future_sekdis.result()

    if stop_event.is_set():
        log("Proses download SKP dihentikan oleh pengguna.")
        return False

    total_gagal = len(gagal_kadis) + len(gagal_sekdis)
    if total_gagal > 0:
        log(f"⚠️ Terdapat {total_gagal} dokumen SKP yang gagal didownload.")
    else:
        log("✅ Seluruh dokumen SKP Kadis & Sekdis berhasil didownload 100%!")

    return True

if __name__ == "__main__":
    test_excel = "data-pegawai.xlsx"
    test_output = "hasil_download_skp"
    run_download_skp(test_excel, test_output)
