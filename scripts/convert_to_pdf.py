import os
from docx2pdf import convert
from tqdm import tqdm

DOCX_DIR = 'data/docx/'
PDF_DIR = 'data/pdf/'
os.makedirs(PDF_DIR, exist_ok=True)

docx_files = [f for f in os.listdir(DOCX_DIR) if f.endswith('.docx')]

for docx_file in tqdm(docx_files, desc="Конвертация DOCX в PDF"):
    input_path = os.path.join(DOCX_DIR, docx_file)
    output_path = os.path.join(PDF_DIR, os.path.splitext(docx_file)[0] + '.pdf')
    try:
        convert(input_path, output_path)
    except Exception as e:
        print(f"Ошибка конвертации {docx_file}: {e}")