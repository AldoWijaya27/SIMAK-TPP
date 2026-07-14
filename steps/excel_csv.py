import os
import pandas as pd

def excel_sheet_disiplin_ke_csv(file_excel, output_folder):
    file_excel = os.path.abspath(file_excel)
    output_folder = os.path.abspath(output_folder)

    # Nama file output csv didasarkan pada nama file excel
    filename = os.path.splitext(os.path.basename(file_excel))[0] + ".csv"
    output_csv = os.path.join(output_folder, filename)

    # Baca sheet "DISIPLIN" dari excel
    df = pd.read_excel(file_excel, sheet_name="DISIPLIN")

    # Tulis ke CSV dengan separator ';' dan encoding utf-8
    df.to_csv(output_csv, sep=';', encoding='utf-8', index=False)
