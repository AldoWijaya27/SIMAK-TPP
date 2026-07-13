import subprocess
import os

def excel_sheet_disiplin_ke_csv(file_excel, output_folder):

    file_excel = os.path.abspath(file_excel)
    output_folder = os.path.abspath(output_folder)

    subprocess.run([
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "--headless",
        "--calc",
        "--convert-to",
        "csv:Text - txt - csv (StarCalc):59,34,76,1",
        "--outdir",
        output_folder,
        file_excel
    ], check=True)
