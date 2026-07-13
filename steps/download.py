import json
import os
import sys
import time
import base64
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from utils import sanitize_filename

thread_local = threading.local()
drivers = []
drivers_lock = threading.Lock()
shared_cookies = []

def perform_manual_login(login_url, log):
    log("Cepetan loginnya, jangan kelamaan, cepetan!!!")
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    
    try:
        driver.get(login_url)
        # Tunggu sampai elemen table muncul (tanda bahwa login sudah berhasil dan masuk ke halaman)
        wait = WebDriverWait(driver, 120)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        time.sleep(1)
        
        # Ambil cookies untuk diberikan ke background browser
        cookies = driver.get_cookies()
        return cookies
    finally:
        driver.quit()

def get_driver():
    if not hasattr(thread_local, "driver"):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new") # Jalan di background
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--window-size=1920,1080")
        
        driver = webdriver.Chrome(options=options)
        
        # Pindah ke domain root dulu agar selenium mengizinkan injeksi cookies
        driver.get("https://dev1.sikap.lampungprov.go.id/favicon.ico")
        
        global shared_cookies
        for cookie in shared_cookies:
            driver.add_cookie(cookie)
            
        thread_local.driver = driver
        with drivers_lock:
            drivers.append(driver)
    return thread_local.driver

def process_pegawai(obj, base_url, dir_rekap, log):
    id_peg = obj["id"]
    nama = sanitize_filename(obj["name"])
    bidang = sanitize_filename(obj.get("bidang", "Lainnya"))

    folder = os.path.join(dir_rekap, bidang)
    os.makedirs(folder, exist_ok=True)

    out = os.path.join(folder, f"{nama}.pdf")

    log(f"Download {nama}")

    driver = get_driver()
    wait = WebDriverWait(driver, 30)

    try:
        driver.get(base_url.format(id=id_peg))
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        time.sleep(1)

        pdf = driver.execute_cdp_cmd("Page.printToPDF", {"scale": 0.5})

        with open(out, "wb") as f:
            f.write(base64.b64decode(pdf["data"]))
    except Exception as e:
        log(f"Error download {nama}: {str(e)}")

def download_rekap(json_pegawai, dir_rekap, bulan, tahun, log):
    with open(json_pegawai, encoding="utf-8") as f:
        items = json.load(f)

    if not items:
        log("Tidak ada data pegawai untuk didownload.")
        return

    BASE_URL = f"https://dev1.sikap.lampungprov.go.id/app/cetak-laporan/data-harian-bulanan-pegawai?bulan={bulan}&tahun={tahun}&id_peg={{id}}"
    
    global drivers, shared_cookies
    drivers = []
    
    # 1. Buka browser normal untuk memancing user login
    first_url = BASE_URL.format(id=items[0]["id"])
    try:
        shared_cookies = perform_manual_login(first_url, log)
        log("Login terdeteksi!")
    except Exception as e:
        log(f"Gagal memverifikasi login atau waktu tunggu habis. Pesan error: {e}")
        return
    
    # 2. Mulai download secara paralel dengan session yang sudah valid
    max_workers = 5 # Menjalankan 5 browser secara paralel
    log(f"Memulai download. {max_workers}x lebih cepat dari biasanya!")

    try:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for obj in items:
                futures.append(executor.submit(process_pegawai, obj, BASE_URL, dir_rekap, log))
            
            for future in as_completed(futures):
                future.result() # Tangkap error jika ada thread yang gagal
    finally:
        # Bersihkan dan tutup semua browser
        for d in drivers:
            try:
                d.quit()
            except:
                pass

    log("Download selesai")
