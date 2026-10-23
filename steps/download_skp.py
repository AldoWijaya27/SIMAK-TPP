import requests
import time
import os
import sys
import random
import subprocess
import pandas as pd
import concurrent.futures
from playwright.sync_api import sync_playwright

from steps.download import stop_event

def interruptible_sleep(seconds):
    """Tidur dengan pengecekan stop_event secara berkala agar responsif terhadap pembatalan."""
    step = 0.25
    elapsed = 0.0
    while elapsed < seconds:
        if stop_event.is_set():
            return False
        time.sleep(min(step, seconds - elapsed))
        elapsed += step
    return True

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
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def download_skp(skp_id, periode, output_filename, auth_token, label="", tanggal_pegawai=None, log=print, session=None, error_info=None):
    if stop_event.is_set():
        if error_info is not None:
            error_info['reason'] = "Dibatalkan oleh pengguna"
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

    # 1. Request Cetak (Antre) dengan Queue-Clearing Backoff & Fast-Skip Error Permanen
    url_cetak = f"https://kinerja.bkn.go.id/api/skp/{skp_id}/penilaian/{periode}/cetak"
    payload = {
        "lokasi": LOKASI,
        "periode": periode,
        "skp_id": skp_id,
        "tanggal_pegawai": tgl_peg
    }

    http = session if session else requests
    job_id = None
    maks_retry_job = 4

    for attempt in range(1, maks_retry_job + 1):
        if stop_event.is_set():
            if error_info is not None:
                error_info['reason'] = "Dibatalkan oleh pengguna"
            return False

        try:
            resp_cetak = http.post(url_cetak, json=payload, headers=headers, timeout=30)
            
            if resp_cetak.status_code == 200:
                try:
                    data_cetak = resp_cetak.json()
                    job_id = data_cetak.get("job_id")
                    if job_id:
                        break  # Berhasil memperoleh job_id
                    else:
                        msg = data_cetak.get("message") or resp_cetak.text[:120]
                        log(f"{prefix}⚠️ BKN belum memberikan job_id: {msg}")

                        # Cek apakah error bersifat permanen (misal periode penilaian tidak ditemukan / belum dinilai)
                        pesan_lower = msg.lower()
                        kata_kunci_permanen = ["tidak ditemukan", "belum dinilai", "tidak ada", "belum disetujui", "belum final", "belum dibuat", "tidak berhak"]
                        if any(k in pesan_lower for k in kata_kunci_permanen):
                            log(f"{prefix}ℹ️ Error permanen dari BKN ({msg}). Melewati percobaan ulang.")
                            if error_info is not None:
                                error_info['reason'] = msg
                                error_info['is_permanent'] = True
                            return False
                except Exception as je:
                    log(f"{prefix}⚠️ [Coba {attempt}/{maks_retry_job}] Format JSON respons cetak tidak sesuai: {je}")
                    if error_info is not None:
                        error_info['reason'] = f"Gagal membaca respons BKN: {je}"
            elif resp_cetak.status_code == 429:
                log(f"{prefix}⚠️ [Coba {attempt}/{maks_retry_job}] Terkena Rate Limit BKN (HTTP 429), jeda sejenak...")
                if error_info is not None:
                    error_info['reason'] = "Rate limit server BKN (HTTP 429)"
                if not interruptible_sleep(6 + random.uniform(1.0, 3.0)):
                    return False
                continue
            elif resp_cetak.status_code in [400, 404, 422]:
                msg = resp_cetak.text[:120]
                log(f"{prefix}⚠️ Server BKN HTTP {resp_cetak.status_code}: {msg}")
                pesan_lower = msg.lower()
                kata_kunci_permanen = ["tidak ditemukan", "belum dinilai", "tidak ada", "periode", "tidak berhak"]
                if any(k in pesan_lower for k in kata_kunci_permanen):
                    log(f"{prefix}ℹ️ Data tidak ditemukan di BKN ({msg}). Melewati percobaan ulang.")
                    if error_info is not None:
                        error_info['reason'] = msg
                        error_info['is_permanent'] = True
                    return False
                if error_info is not None:
                    error_info['reason'] = f"HTTP {resp_cetak.status_code}: {msg}"
            else:
                msg = resp_cetak.text[:120]
                log(f"{prefix}⚠️ [Coba {attempt}/{maks_retry_job}] Server BKN HTTP {resp_cetak.status_code}: {msg}")
                if error_info is not None:
                    error_info['reason'] = f"HTTP {resp_cetak.status_code}: {msg}"

        except requests.exceptions.Timeout:
            log(f"{prefix}⚠️ [Coba {attempt}/{maks_retry_job}] Request antrean cetak timeout (>30s)")
            if error_info is not None:
                error_info['reason'] = "Timeout request antrean ke server BKN (>30s)"
        except Exception as e:
            log(f"{prefix}⚠️ [Coba {attempt}/{maks_retry_job}] Kesalahan jaringan request cetak: {e}")
            if error_info is not None:
                error_info['reason'] = f"Kesalahan jaringan: {e}"

        if attempt < maks_retry_job:
            # Beri jeda agar worker server BKN menyelesaikan antrean yang menggantung
            jeda = attempt * 3 + random.uniform(0.5, 1.5)
            log(f"{prefix}⏳ Menunggu {jeda:.1f}s sebelum meminta ulang job_id...")
            if not interruptible_sleep(jeda):
                return False

    if not job_id:
        log(f"{prefix}❌ Gagal mendapatkan job_id setelah {maks_retry_job}x percobaan.")
        if error_info is not None and 'reason' not in error_info:
            error_info['reason'] = "Gagal mendapatkan job_id dari server BKN"
        return False

    # 2. Polling Status Antrean Cetak hingga status == "ready" (ready: true)
    url_status = f"https://kinerja.bkn.go.id/api/skp/cetak_status/{job_id}"
    maks_wait_status = 35  # Polling hingga ~45 detik
    is_ready = False

    for s_attempt in range(1, maks_wait_status + 1):
        if stop_event.is_set():
            if error_info is not None:
                error_info['reason'] = "Dibatalkan oleh pengguna"
            return False

        if not interruptible_sleep(1.5):
            return False

        try:
            resp_status = http.get(url_status, headers=headers, timeout=20)
            if resp_status.status_code == 200:
                data_status = resp_status.json()
                # Respons resmi: {"status":"ready","ready":true,"error":null}
                if data_status.get("ready") is True or data_status.get("status") == "ready":
                    is_ready = True
                    break

                err_bkn = data_status.get("error")
                if err_bkn:
                    log(f"{prefix}❌ Error rendering dokumen dari BKN: {err_bkn}")
                    if error_info is not None:
                        error_info['reason'] = f"BKN render error: {err_bkn}"
                        error_info['is_permanent'] = True
                    return False

                status_str = data_status.get("status", "queued")
                if s_attempt % 5 == 0:
                    log(f"{prefix}⏳ Menunggu antrean dokumen selesai di server BKN ({status_str})...")
            else:
                if s_attempt % 5 == 0:
                    log(f"{prefix}⚠️ Cek status cetak HTTP {resp_status.status_code} (Coba {s_attempt}/{maks_wait_status})")
        except Exception:
            pass  # Jaringan fluktuatif sesaat saat polling

    if not is_ready:
        log(f"{prefix}❌ Dokumen belum siap (timeout) di server BKN untuk job_id: {job_id}")
        if error_info is not None:
            error_info['reason'] = "Timeout menunggu dokumen ready di server BKN"
        return False

    # 3. Download Dokumen PDF (setelah dipastikan ready: true)
    url_download = f"https://kinerja.bkn.go.id/api/skp/cetak_download/{job_id}"
    maks_retry_dl = 3

    for dl_attempt in range(1, maks_retry_dl + 1):
        if stop_event.is_set():
            if error_info is not None:
                error_info['reason'] = "Dibatalkan oleh pengguna"
            return False

        try:
            resp_download = http.get(url_download, headers=headers, timeout=60)
            if resp_download.status_code == 200:
                content = resp_download.content
                if content.startswith(b"%PDF") or (len(content) > 1000 and not content.strip().startswith(b"{")):
                    with open(output_filename, 'wb') as f:
                        f.write(content)
                    log(f"{prefix}✔ Berhasil menyimpan: {os.path.basename(output_filename)}")
                    return True
                else:
                    log(f"{prefix}⚠️ File download bukan biner PDF valid (Coba {dl_attempt}/{maks_retry_dl})")
            else:
                log(f"{prefix}⚠️ Gagal download HTTP {resp_download.status_code} (Coba {dl_attempt}/{maks_retry_dl})")
        except Exception as e:
            log(f"{prefix}⚠️ Error download dokumen: {e}")

        if dl_attempt < maks_retry_dl:
            if not interruptible_sleep(2.0):
                return False

    log(f"{prefix}❌ Gagal download dokumen SKP untuk job_id: {job_id}")
    if error_info is not None and 'reason' not in error_info:
        error_info['reason'] = "Gagal mengunduh file biner PDF"
    return False

def setup_playwright_env():
    """Memastikan path browser Playwright terkonfigurasi dengan benar di environment PyInstaller / Windows."""
    if getattr(sys, "frozen", False):
        if "PLAYWRIGHT_BROWSERS_PATH" not in os.environ:
            local_appdata = os.environ.get("LOCALAPPDATA", "")
            if local_appdata:
                user_pw_dir = os.path.join(local_appdata, "ms-playwright")
                if os.path.isdir(user_pw_dir):
                    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = user_pw_dir

def install_playwright_chromium(log=print):
    """Mengunduh browser Chromium menggunakan CLI Playwright bawaan jika tidak ada browser lain yang tersedia."""
    try:
        from playwright._impl._driver import compute_driver_executable

        node, cli = compute_driver_executable()
        env = os.environ.copy()
        local_appdata = os.environ.get("LOCALAPPDATA", "")
        if local_appdata:
            user_pw_dir = os.path.join(local_appdata, "ms-playwright")
            env["PLAYWRIGHT_BROWSERS_PATH"] = user_pw_dir
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = user_pw_dir

        log("  📥 Sedang mengunduh browser Chromium untuk Playwright (~150MB). Mohon tunggu...")
        proc = subprocess.run([node, cli, "install", "chromium"], env=env, capture_output=True, text=True)
        if proc.returncode == 0:
            log("  ✅ Berhasil mengunduh browser Chromium.")
            return True
        else:
            err_msg = proc.stderr.strip() or proc.stdout.strip()
            log(f"  ❌ Gagal mengunduh Chromium: {err_msg}")
            return False
    except Exception as e:
        log(f"  ❌ Error saat instalasi browser Chromium: {e}")
        return False

def launch_browser_context(p, user_data_dir, log=print):
    """
    Meluncurkan persistent browser context dengan prioritas:
    1. Google Chrome (channel='chrome') - browser utama jika terpasang di sistem
    2. Microsoft Edge (channel='msedge') - browser bawaan Windows 10 & 11
    3. Chromium bawaan Playwright (channel=None)
    4. Auto-download Chromium jika belum ada browser yang bisa dipakai
    """
    setup_playwright_env()
    channels = [
        ("chrome", "Google Chrome"),
        ("msedge", "Microsoft Edge"),
        (None, "Chromium (Playwright)")
    ]

    last_error = None
    for channel, browser_name in channels:
        try:
            kwargs = {
                "user_data_dir": user_data_dir,
                "headless": False,
                "viewport": {"width": 1280, "height": 720}
            }
            if channel:
                kwargs["channel"] = channel

            ctx = p.chromium.launch_persistent_context(**kwargs)
            log(f"  Browser berhasil dibuka ({browser_name}).")
            return ctx
        except Exception as e:
            last_error = e
            continue

    # Jika semua browser di atas tidak ditemukan, coba unduh Chromium otomatis
    log("  ⚠ Browser Google Chrome & Microsoft Edge tidak dapat diluncurkan. Mencoba mengunduh Chromium...")
    if install_playwright_chromium(log=log):
        try:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                viewport={"width": 1280, "height": 720}
            )
            log("  Browser berhasil dibuka (Chromium terinstal).")
            return ctx
        except Exception as e:
            last_error = e

    log(f"❌ Gagal meluncurkan browser: {last_error}")
    return None

def dapatkan_token_otomatis(label, log=print):
    log(f"── Otentikasi Akun {label.upper()} ──")
    log(f"  Membuka browser untuk {label}...")
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
        context = launch_browser_context(p, user_data_dir, log=log)
        if not context:
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
        session = requests.Session()
        semua_gagal_permanen = []

        try:
            while antrean_berjalan and percobaan <= maks_percobaan:
                if stop_event.is_set():
                    break

                if percobaan > 1:
                    log(f"[*] Mengulang yang gagal sementara - {label} (Percobaan {percobaan}/{maks_percobaan}, {len(antrean_berjalan)} orang)...")
                    if not interruptible_sleep(5):
                        break

                antrean_gagal_siklus = []

                for bawahan in antrean_berjalan:
                    if stop_event.is_set():
                        break

                    nama = bawahan['nama']
                    skp_id = bawahan['skp_id']
                    filename = os.path.join(folder_output, f"SKP_{label}_{nama}.pdf")

                    err_info = {}
                    berhasil = download_skp(
                        skp_id, periode_final, filename, token,
                        label=label, tanggal_pegawai=tanggal_final, log=log, session=session,
                        error_info=err_info
                    )
                    if not berhasil:
                        item_gagal = dict(bawahan)
                        item_gagal['alasan'] = err_info.get('reason', 'Gagal memproses')
                        item_gagal['permanen'] = err_info.get('is_permanent', False)

                        if item_gagal['permanen']:
                            # Langsung tandai permanen, JANGAN diulang di percobaan 2 & 3
                            semua_gagal_permanen.append(item_gagal)
                        else:
                            antrean_gagal_siklus.append(item_gagal)

                    # Jeda santai antar-pegawai agar antrean di server BKN benar-benar bersih
                    if not stop_event.is_set():
                        interruptible_sleep(random.uniform(1.5, 2.5))

                antrean_berjalan = antrean_gagal_siklus
                percobaan += 1
        finally:
            session.close()

        total_list_gagal = semua_gagal_permanen + antrean_berjalan
        return total_list_gagal

    log("[*] Memulai download paralel untuk Kadis & Sekdis...")
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        future_kadis = executor.submit(proses_antrean, antrean_kadis, auth_token_kadis, "KADIS")
        time.sleep(1)  # Jeda inisiasi agar kedua thread tidak menembak request pertama di milidetik yang persis sama
        future_sekdis = executor.submit(proses_antrean, antrean_sekdis, auth_token_sekdis, "SEKDIS")

        gagal_kadis = future_kadis.result()
        gagal_sekdis = future_sekdis.result()

    if stop_event.is_set():
        log("Proses download SKP dihentikan oleh pengguna.")
        return False

    total_gagal = len(gagal_kadis) + len(gagal_sekdis)
    if total_gagal > 0:
        log("\n" + "=" * 65)
        log(f"⚠️ REKAPITULASI: Terdapat {total_gagal} pegawai yang gagal didownload:")
        log("=" * 65)
        if gagal_kadis:
            log("📌 Bawahan Kepala Dinas (KADIS):")
            for idx, p in enumerate(gagal_kadis, 1):
                alasan = p.get('alasan', 'Gagal memproses')
                log(f"   {idx}. {p['nama']} (Alasan: {alasan})")
        if gagal_sekdis:
            log("📌 Bawahan Sekretaris (SEKDIS):")
            for idx, p in enumerate(gagal_sekdis, 1):
                alasan = p.get('alasan', 'Gagal memproses')
                log(f"   {idx}. {p['nama']} (Alasan: {alasan})")
        log("=" * 65 + "\n")
    else:
        log("\n✅ Seluruh dokumen SKP Kadis & Sekdis berhasil didownload 100%!\n")

    return True

if __name__ == "__main__":
    test_excel = "data-pegawai.xlsx"
    test_output = "EKIN SKP"
    run_download_skp(test_excel, test_output)
