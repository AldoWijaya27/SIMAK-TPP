import subprocess
import os
import platform

def excel_sheet_disiplin_ke_csv(file_excel, output_folder):

    file_excel = os.path.abspath(file_excel)
    output_folder = os.path.abspath(output_folder)

    if platform.system() == "Darwin":
        soffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:
        soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"

    subprocess.run([
        soffice_path,
        "--headless",
        "--calc",
        "--convert-to",
        "csv:Text - txt - csv (StarCalc):59,34,76,1",
        "--outdir",
        output_folder,
        file_excel
    ], check=True)
