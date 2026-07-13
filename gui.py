import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from datetime import datetime
import os
import traceback

from steps.merge_ekin_apel import merge_ekin_apel
from steps.excel_csv import excel_sheet_disiplin_ke_csv
from steps.download import download_rekap
from steps.input_spesimen import input_spesimen
from steps.analisis import analisis_kehadiran
from steps.merge2 import process_mail_merge
from config import TEMPLATE_WORD
from services.firebase_access import FirebaseRealtimeUserAccessGateway, ValidateUserAccess

if getattr(sys, 'frozen', False):
    os.environ['WDM_LOCAL'] = '1'

FIREBASE_URL = "https://simak-tpp-default-rtdb.asia-southeast1.firebasedatabase.app"
APP_USERNAME = "dinaskehutanan"

class App:
    def __init__(self, root):

        self.root = root
        root.title("SIMAK-TPP")

        main_frame = tk.Frame(root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        left_frame = tk.Frame(main_frame)
        left_frame.pack(side="left", fill="y", padx=10)

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=10)

        tk.Label(left_frame, text="Pilih Proses Yang Dijalankan").pack(pady=5)

        self.var_download = tk.BooleanVar(value=True)
        self.var_analisis = tk.BooleanVar(value=True)
        self.var_spesimen = tk.BooleanVar(value=True)
        self.var_merge = tk.BooleanVar(value=True)
        self.var_csv = tk.BooleanVar(value=True)
        self.var_mailmerge = tk.BooleanVar(value=True)

        tk.Checkbutton(left_frame, text="Download Rekap Kehadiran", variable=self.var_download).pack(anchor="w")
        tk.Checkbutton(left_frame, text="Input Spesimen", variable=self.var_spesimen).pack(anchor="w")
        tk.Checkbutton(left_frame, text="Analisis Kehadiran", variable=self.var_analisis).pack(anchor="w")
        tk.Checkbutton(left_frame, text="Gabung Ekin & Apel", variable=self.var_merge).pack(anchor="w")
        tk.Checkbutton(left_frame, text="Generate CSV", variable=self.var_csv).pack(anchor="w")
        tk.Checkbutton(left_frame, text="Mail Merge TPP", variable=self.var_mailmerge).pack(anchor="w")

        tk.Label(right_frame, text="Template Excel").pack()
        self.entry_template_excel = tk.Entry(right_frame, width=60)
        self.entry_template_excel.pack()

        tk.Button(
            right_frame, 
            text="Browse Excel",
            command=lambda: self.entry_template_excel.insert(
                0,
                filedialog.askopenfilename(
                    title="Pilih File Excel",
                    filetypes=[
                        ("Excel Workbook", "*.xlsx")
                    ]
                )
            )
        ).pack()

        tk.Label(right_frame, text="Template JSON").pack()
        self.entry_template_json = tk.Entry(right_frame, width=60)
        self.entry_template_json.pack()

        tk.Button(
            right_frame, 
            text="Browse JSON",
            command=lambda: self.entry_template_json.insert(
                0,
                filedialog.askopenfilename(
                    title="Pilih File JSON",
                    filetypes=[
                        ("JSON File", "*.json")
                    ]
                )
            )
        ).pack()

        tk.Label(right_frame, text="Data Ekin dan Apel").pack()
        self.entry_ekin_apel = tk.Entry(right_frame, width=60)
        self.entry_ekin_apel.pack()

        tk.Button(
            right_frame, 
            text="Browse Excel",
            command=lambda: self.entry_ekin_apel.insert(
                0,
                filedialog.askopenfilename(
                    title="Pilih File Excel",
                    filetypes=[
                        ("Excel Workbook", "*.xlsx")
                    ]
                )
            )
        ).pack()
        
        tk.Label(right_frame,  text="Folder Penyimpanan").pack()
        self.entry_base_dir = tk.Entry(right_frame,  width=60)
        self.entry_base_dir.pack()

        tk.Button(
            right_frame, 
            text="Pilih Folder",
            command=lambda: self.entry_base_dir.insert(
                0,
                filedialog.askdirectory(title="Pilih Folder Penyimpanan")
            )
        ).pack()

        tk.Label(right_frame, text="Kalender Kerja JSON (Opsional)").pack()
        self.entry_kalender_json = tk.Entry(right_frame, width=60)
        self.entry_kalender_json.pack()

        tk.Button(
            right_frame, 
            text="Browse JSON",
            command=lambda: self.entry_kalender_json.insert(
                0,
                filedialog.askopenfilename(
                    title="Pilih File Kalender Kerja JSON",
                    filetypes=[
                        ("JSON File", "*.json")
                    ]
                )
            )
        ).pack()

        # Periode
        frame = tk.Frame(right_frame)
        frame.pack()

        tk.Label(frame,text="Bulan").grid(row=0,column=0)
        self.entry_bulan = tk.Entry(frame,width=5)
        self.entry_bulan.grid(row=0,column=1)

        tk.Label(frame,text="Tahun").grid(row=0,column=2)
        self.entry_tahun = tk.Entry(frame,width=8)
        self.entry_tahun.grid(row=0,column=3)

        now = datetime.now()
        bulan = now.month - 1 if now.month > 1 else 12
        self.entry_bulan.insert(0, f"{bulan:02d}")

        self.entry_tahun.insert(0, now.year)
        self.btn_action = tk.Button(right_frame, text="JALANKAN",bg="green",fg="white",
                  command=self.toggle_process)
        self.btn_action.pack(pady=10)

        self.log_box = tk.Text(root,height=15,width=80)
        self.log_box.pack()

        self.access_validator = ValidateUserAccess(
            gateway=FirebaseRealtimeUserAccessGateway(base_url=FIREBASE_URL)
        )

    def log(self,text):
        def _update():
            self.log_box.insert(tk.END,str(text)+"\n")
            self.log_box.see(tk.END)
            self.root.update()
        self.root.after(0, _update)

    def jalankan(self):
        try:
            base_dir = self.entry_base_dir.get()

            # NORMALISASI KERAS
            base_dir = os.path.abspath(base_dir)
            base_dir = os.path.normpath(base_dir)
            base_dir = base_dir.strip()

            # Buat struktur folder
            DIR_REKAP = os.path.join(base_dir, "REKAP KEHADIRAN")
            DIR_REKAP_DITANDATANGANI = os.path.join(base_dir, "REKAP KEHADIRAN DITANDATANGANI")
            DIR_OUTPUT = os.path.join(base_dir, "PERHITUNGAN TPP")
            TEMP_DIR = os.path.join(base_dir, "TEMP")

            DIR_REKAP_DITANDATANGANI = os.path.abspath(DIR_REKAP_DITANDATANGANI)
            DIR_OUTPUT = os.path.abspath(DIR_OUTPUT)
            TEMP_DIR = os.path.abspath(TEMP_DIR)

            os.makedirs(DIR_REKAP, exist_ok=True)
            os.makedirs(DIR_REKAP_DITANDATANGANI, exist_ok=True)
            os.makedirs(DIR_OUTPUT, exist_ok=True)
            os.makedirs(TEMP_DIR, exist_ok=True)

            excel = self.entry_template_excel.get()
            word = TEMPLATE_WORD
            ekin_apel = self.entry_ekin_apel.get()
            json_pegawai = self.entry_template_json.get()
            json_kalender = self.entry_kalender_json.get()

            bulan = self.entry_bulan.get().zfill(2)
            tahun = int(self.entry_tahun.get())

            if not excel or not word:
                self.root.after(0, lambda: messagebox.showerror("Error","Template belum dipilih"))
                return

            output_excel = os.path.join(base_dir,"Disiplin_TPP.xlsx")
            output_template_ready = os.path.join(base_dir,"Disiplin_TPP_Lengkap.xlsx")
            csv_output = os.path.join(base_dir,"Disiplin_TPP_Lengkap.csv")


            # STEP 1 — DOWNLOAD
            if self.var_download.get():
                self.log("== Download Rekap Kehadiran ==")
                download_rekap(json_pegawai, DIR_REKAP, bulan, tahun, self.log)

            # STEP 1.5 — INPUT SPESIMEN
            if self.var_spesimen.get():
                self.log("== Input Spesimen Tanda Tangan ==")
                input_spesimen(DIR_REKAP, DIR_REKAP_DITANDATANGANI, self.log)

            # STEP 2 — ANALISIS
            if self.var_analisis.get():
                self.log("== Analisis Kehadiran ==")
                analisis_kehadiran(DIR_REKAP, excel, output_excel, self.log, json_kalender)

            # STEP 3 — MERGE EKIN APEL
            if self.var_merge.get():
                self.log("== Merge Ekin & Apel ==")
                status, pesan = merge_ekin_apel(
                    ekin_apel,
                    output_excel,
                    output_template_ready
                )
                self.log(pesan)

                if not status:
                    messagebox.showerror("Error", pesan)
                    return

            # STEP 4 — CSV
            if self.var_csv.get():
                self.log("== Generate CSV ==")
                excel_sheet_disiplin_ke_csv(output_template_ready, base_dir)

            # STEP 5 — MAIL MERGE
            if self.var_mailmerge.get():
                self.log("== Mail Merge TPP ==")
                process_mail_merge(DIR_REKAP_DITANDATANGANI, DIR_OUTPUT, TEMP_DIR, csv_output, word, self.log)

            # Jika berhasil semua
            self.log("Semua proses selesai dengan sukses.")
            self.root.after(0, lambda: messagebox.showinfo("Selesai","Semua proses selesai"))

        # BAGIAN INI AKAN MENANGKAP ERROR APAPUN YANG BIKIN STUCK
        except SystemExit:
            # Jika diberhentikan paksa oleh pengguna
            pass
        except Exception as e:
            error_detail = traceback.format_exc() # Mengambil detail baris yang error
            self.log("\n================ ERROR TERJADI ================")
            self.log(error_detail)
            self.log("===============================================")
            
            # Memunculkan pop-up error ke layar pengguna
            error_msg = f"Aplikasi berhenti karena error:\n{str(e)}\n\nSilakan cek kotak log di aplikasi untuk detailnya."
            self.root.after(0, lambda: messagebox.showerror("Terjadi Kesalahan Kritis", error_msg))
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_action.config(text="JALANKAN", bg="green"))

    def toggle_process(self):
        if getattr(self, 'is_running', False):
            self.stop_process()
        else:
            self.start_process()

    def start_process(self):
        is_allowed = self.access_validator.execute(APP_USERNAME)
        if not is_allowed:
            self.root.after(0, lambda: messagebox.showerror("Error", "ERROR!"))
            return

        self.is_running = True
        self.btn_action.config(text="STOP", bg="red")
        self.log_box.delete('1.0', tk.END) # Bersihkan log sebelumnya
        self.process_thread = threading.Thread(target=self.jalankan, daemon=True)
        self.process_thread.start()

    def stop_process(self):
        if hasattr(self, 'process_thread') and self.process_thread.is_alive():
            import ctypes
            exc = ctypes.py_object(SystemExit)
            res = ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_long(self.process_thread.ident), exc)
            if res > 1:
                ctypes.pythonapi.PyThreadState_SetAsyncExc(self.process_thread.ident, None)
        
        self.is_running = False
        self.btn_action.config(text="JALANKAN", bg="green")
        self.log("\n================ PROSES DIBERHENTIKAN ================")
        self.log("Proses telah diberhentikan oleh pengguna.")
        messagebox.showinfo("Berhenti", "Proses berhasil diberhentikan. Silahkan mulai lagi.")