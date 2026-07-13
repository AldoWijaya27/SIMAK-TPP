import json
from datetime import datetime, time, date
from Aplikasi.app_path import get_file

BULAN_ID = {
    "januari":1, "februari":2, "maret":3, "april":4,
    "mei":5, "juni":6, "juli":7, "agustus":8,
    "september":9, "oktober":10, "november":11, "desember":12
}


# =========================
# PARSE TANGGAL INDONESIA
# =========================
def parse_tanggal_indonesia(teks):
    try:
        bagian = teks.strip().lower().split()
        hari = int(bagian[0])
        bulan = BULAN_ID[bagian[1]]
        tahun = int(bagian[2])
        return date(tahun, bulan, hari)
    except Exception:
        return None


# =========================
# WORK CALENDAR ENGINE
# =========================
class WorkCalendar:

    def __init__(self, json_path=""):

        json_path = json_path.strip() if json_path else ""
        self.rules = []

        if json_path:
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    self.rules = json.load(f)

                for r in self.rules:
                    r["mulai"] = datetime.strptime(r["mulai"], "%Y-%m-%d").date()
                    r["selesai"] = datetime.strptime(r["selesai"], "%Y-%m-%d").date()
                    r["jam_masuk"] = datetime.strptime(r["jam_masuk"], "%H:%M:%S").time()
            except Exception as e:
                print(f"Gagal membaca atau memproses file '{json_path}': {e}. Menggunakan default jam masuk 07.30.")


    def get_jam_masuk(self, tanggal: date) -> time:

        for r in self.rules:
            if r["mulai"] <= tanggal <= r["selesai"]:
                return r["jam_masuk"]

        # default nasional ASN
        return time(7,30,0)