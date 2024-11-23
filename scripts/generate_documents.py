import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from doc_generator.doc_generator import generate_document
from tqdm import tqdm

DOCX_DIR = 'data/docx/'
os.makedirs(DOCX_DIR, exist_ok=True)

NUM_DOCUMENTS = 1  # ПОМЕНЯТЬ НА 200 ПРИ ОСНОВНОЙ ГЕНЕРАЦИИ

NUM_OF_ALREADY_DONE_DOCS = 250  # ПОКА НЕИЗВЕСТНО, ПОСТАВЛЕНО ЗАРАНЕЕ 250

for i in tqdm(range(NUM_OF_ALREADY_DONE_DOCS, NUM_OF_ALREADY_DONE_DOCS + NUM_DOCUMENTS), desc="Генерация документов"):
    doc_path = os.path.join(DOCX_DIR, f'demo{i}.docx')
    generate_document(doc_path)