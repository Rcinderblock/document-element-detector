import subprocess
import os

DOCX_DIR = 'data/docx/'
PDF_DIR = 'data/pdf/'
os.makedirs(PDF_DIR, exist_ok=True)

docx_files = [f for f in os.listdir(DOCX_DIR) if f.endswith('.docx')]

for docx_file in docx_files:
    input_path = os.path.join(DOCX_DIR, docx_file)
    output_path = os.path.join(PDF_DIR, os.path.splitext(docx_file)[0] + '.pdf')
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', input_path, '--outdir', PDF_DIR], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Ошибка конвертации {docx_file}: {e}")