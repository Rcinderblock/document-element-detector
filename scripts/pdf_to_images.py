import os
import fitz
from tqdm import tqdm

PDF_DIR = 'data/pdf/'
IMAGE_DIR = 'data/images/'
os.makedirs(IMAGE_DIR, exist_ok=True)

pdf_files = [f for f in os.listdir(PDF_DIR) if f.endswith('.pdf')]

for pdf_file in tqdm(pdf_files, desc="Конвертация PDF в изображения"):
    pdf_path = os.path.join(PDF_DIR, pdf_file)
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=300)
            image_name = f"{os.path.splitext(pdf_file)[0]}_{page_num + 1}.png"
            pix.save(os.path.join(IMAGE_DIR, image_name))
        doc.close()
    except Exception as e:
        print(f"Ошибка конвертации {pdf_file}: {e}")