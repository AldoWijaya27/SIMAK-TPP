import json
from datetime import datetime, time, date
from app_path import get_file

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
                if json_path.lower().endswith(('.xlsx', '.xls')):
                    import pandas as pd
                    # Baca file Excel
                    df = pd.read_excel(json_path)
                    
                    # Bersihkan spasi di nama kolom, ubah ke lowercase
                    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
                    
                    # Pastikan kolom-kolom penting ada
                    required = ["nama", "mulai", "selesai", "jam_masuk"]
                    if not all(col in df.columns for col in required):
                        missing = [col for col in required if col not in df.columns]
                        raise ValueError(f"Kolom wajib tidak ditemukan di Excel: {missing}")
                    
                    # Konversi baris ke list of dict
                    raw_rules = df.to_dict(orient="records")
                    
                    for r in raw_rules:
                        # Membersihkan data kosong/NaN
                        if pd.isna(r["nama"]) or pd.isna(r["mulai"]) or pd.isna(r["selesai"]) or pd.isna(r["jam_masuk"]):
                            continue
                        
                        rule = {
                            "nama": str(r["nama"]).strip(),
                            "mulai": None,
                            "selesai": None,
                            "jam_masuk": None
                        }
                        
                        # Parsing tanggal mulai
                        m = r["mulai"]
                        if isinstance(m, datetime):
                            rule["mulai"] = m.date()
                        elif isinstance(m, date):
                            rule["mulai"] = m
                        else:
                            rule["mulai"] = datetime.strptime(str(m).strip()[:10], "%Y-%m-%d").date()
                            
                        # Parsing tanggal selesai
                        s = r["selesai"]
                        if isinstance(s, datetime):
                            rule["selesai"] = s.date()
                        elif isinstance(s, date):
                            rule["selesai"] = s
                        else:
                            rule["selesai"] = datetime.strptime(str(s).strip()[:10], "%Y-%m-%d").date()
                            
                        # Parsing jam masuk
                        jm = r["jam_masuk"]
                        if isinstance(jm, time):
                            rule["jam_masuk"] = jm
                        elif isinstance(jm, datetime):
                            rule["jam_masuk"] = jm.time()
                        else:
                            # Coba parsing string
                            jm_str = str(jm).strip()
                            # Menangani format "HH:MM:SS" atau "HH:MM"
                            parts = jm_str.split(':')
                            if len(parts) >= 2:
                                h = int(parts[0])
                                m_val = int(parts[1])
                                s_val = int(parts[2]) if len(parts) >= 3 else 0
                                rule["jam_masuk"] = time(h, m_val, s_val)
                            else:
                                raise ValueError(f"Format jam masuk tidak dikenal: {jm_str}")
                                
                        self.rules.append(rule)
                else:
                    # Parse as JSON
                    with open(json_path, "r", encoding="utf-8") as f:
                        self.rules = json.load(f)

                    for r in self.rules:
                        r["mulai"] = datetime.strptime(r["mulai"], "%Y-%m-%d").date()
                        r["selesai"] = datetime.strptime(r["selesai"], "%Y-%m-%d").date()
                        
                        # Handle jam_masuk jika berbentuk string
                        jm = r["jam_masuk"]
                        if isinstance(jm, str):
                            r["jam_masuk"] = datetime.strptime(jm, "%H:%M:%S").time()
            except Exception as e:
                print(f"Gagal membaca atau memproses file '{json_path}': {e}. Menggunakan default jam masuk 07.30.")


    def get_jam_masuk(self, tanggal: date) -> time:

        for r in self.rules:
            if r["mulai"] <= tanggal <= r["selesai"]:
                return r["jam_masuk"]

        # default nasional ASN
        return time(7,30,0)