import os
import re
import json
import fitz  # PyMuPDF
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBoxHorizontal, LTImage, LTChar
from PIL import Image
from tqdm import tqdm
import camelot

# Папки с данными
PDF_DIR = 'data/pdf/'
IMAGE_DIR = 'data/images/'
ANNOTATIONS_DIR = 'data/annotations/'

# Создание директории для аннотаций, если она не существует
os.makedirs(ANNOTATIONS_DIR, exist_ok=True)

def is_header(element, page_height, threshold=50):
    """
    Определяет, является ли элемент колонтитулом (header).
    """
    # Если y0 элемента выше определённого порога, считаем его колонтитулом
    return element[1] > (page_height - threshold)

def is_footer(element, page_height, threshold=50):
    """
    Определяет, является ли элемент нижним колонтитулом (footer).
    """
    # Если y1 элемента ниже определённого порога, считаем его нижним колонтитулом
    return element[3] < threshold

def extract_tables_camelot(pdf_path, page_num):
    """
    Извлекает таблицы с помощью Camelot и возвращает их координаты.
    """
    try:
        tables = camelot.read_pdf(pdf_path, pages=str(page_num + 1), flavor='stream')  # 'stream' для таблиц без явных границ
        table_coords = []
        for table in tables:
            bbox = table._bbox  # (x1, y1, x2, y2)
            table_coords.append([bbox[0], bbox[1], bbox[2], bbox[3]])
        return table_coords
    except Exception as e:
        print(f"Ошибка извлечения таблиц из {pdf_path}, страница {page_num + 1}: {e}")
        return []

def extract_elements(pdf_path, page_num):
    """
    Извлекает элементы из конкретной страницы PDF-файла.
    """
    elements = {
        "title": [],
        "paragraph": [],
        "table": [],
        "picture": [],
        "table_signature": [],
        "picture_signature": [],
        "numbered_list": [],
        "marked_list": [],
        "header": [],
        "footer": [],
        "footnote": [],
        "formula": []
    }
    try:
        # Открываем PDF с помощью fitz для получения размеров страницы
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_num)
        page_height = page.rect.height
        doc.close()

        # Извлечение таблиц с помощью Camelot
        table_coords = extract_tables_camelot(pdf_path, page_num)
        elements["table"].extend(table_coords)

        # Извлечение текстовых элементов с помощью pdfminer
        for page_layout in extract_pages(pdf_path, page_numbers=[page_num]):
            for element in page_layout:
                if isinstance(element, LTTextBoxHorizontal):
                    text = element.get_text().strip()

                    # Получаем информацию о шрифтах и размерах
                    fonts_sizes = set()
                    for obj in element:
                        if isinstance(obj, LTChar):
                            fonts_sizes.add(obj.size)
                    avg_font_size = sum(fonts_sizes) / len(fonts_sizes) if fonts_sizes else 12

                    # Регулярные выражения для распознавания элементов
                    footnote_pattern = re.compile(r".+[\u2070\u00B9\u00B2\u00B3\u2074-\u2079]$")
                    picture_signature_pattern = re.compile(r"^(Рис\.|Рисунок|Figure)\s+\d+", re.IGNORECASE)
                    table_signature_pattern = re.compile(r"^(Табл\.|Таблица|Table)\s+\d+", re.IGNORECASE)
                    numbered_list_pattern = re.compile(r"^\d+\.")
                    marked_list_pattern = re.compile(r"^-")
                    formula_pattern = re.compile(r'=|\\int|\\sum|\\frac')

                    # Определение типа элемента
                    if footnote_pattern.match(text):
                        elements["footnote"].append([element.x0, element.y0, element.x1, element.y1])
                    elif picture_signature_pattern.match(text):
                        elements["picture_signature"].append([element.x0, element.y0, element.x1, element.y1])
                    elif table_signature_pattern.match(text):
                        elements["table_signature"].append([element.x0, element.y0, element.x1, element.y1])
                    elif numbered_list_pattern.match(text):
                        elements["numbered_list"].append([element.x0, element.y0, element.x1, element.y1])
                    elif marked_list_pattern.match(text):
                        elements["marked_list"].append([element.x0, element.y0, element.x1, element.y1])
                    elif formula_pattern.search(text):
                        elements["formula"].append([element.x0, element.y0, element.x1, element.y1])
                    elif avg_font_size >= 16:
                        elements["title"].append([element.x0, element.y0, element.x1, element.y1])
                    elif avg_font_size >= 14:
                        elements["title"].append([element.x0, element.y0, element.x1, element.y1])
                    elif len(text) > 100:
                        elements["paragraph"].append([element.x0, element.y0, element.x1, element.y1])
                    else:
                        # Дополнительная логика для определения заголовков второго уровня или других элементов
                        if len(text) > 50:
                            elements["paragraph"].append([element.x0, element.y0, element.x1, element.y1])
                        else:
                            elements["title"].append([element.x0, element.y0, element.x1, element.y1])

                elif isinstance(element, LTImage):
                    elements["picture"].append([element.x0, element.y0, element.x1, element.y1])

        # Вторичный проход для определения колонтитулов
        for page_layout in extract_pages(pdf_path, page_numbers=[page_num]):
            for element in page_layout:
                if isinstance(element, LTTextBoxHorizontal):
                    text = element.get_text().strip()
                    # Получаем информацию о шрифтах и размерах
                    fonts_sizes = set()
                    for obj in element:
                        if isinstance(obj, LTChar):
                            fonts_sizes.add(obj.size)
                    avg_font_size = sum(fonts_sizes) / len(fonts_sizes) if fonts_sizes else 12

                    # Определение колонтитулов
                    if is_header([element.x0, element.y0, element.x1, element.y1], page_height):
                        elements["header"].append([element.x0, element.y0, element.x1, element.y1])
                    elif is_footer([element.x0, element.y0, element.x1, element.y1], page_height):
                        elements["footer"].append([element.x0, element.y0, element.x1, element.y1])

        return elements

    except Exception as e:
        print(f"Ошибка извлечения элементов из {pdf_path}, страница {page_num + 1}: {e}")
        return elements

def extract_annotations():
    """
    Основная функция для извлечения аннотаций из всех PDF-файлов.
    """
    pdf_files = [f for f in os.listdir(PDF_DIR) if f.endswith('.pdf')]

    for pdf_file in tqdm(pdf_files, desc="Извлечение разметки из PDF"):
        pdf_path = os.path.join(PDF_DIR, pdf_file)
        try:
            doc = fitz.open(pdf_path)
            num_pages = len(doc)
            doc.close()

            for page_num in range(num_pages):
                annotations = extract_elements(pdf_path, page_num)

                # Определение пути к изображению страницы
                image_name = f"{os.path.splitext(pdf_file)[0]}_{page_num + 1}.png"
                image_path = os.path.join(IMAGE_DIR, image_name)

                if not os.path.exists(image_path):
                    print(f"Изображение {image_name} не найдено, пропускаем")
                    continue

                with Image.open(image_path) as img:
                    width, height = img.size
                    annotations["image_height"] = height
                    annotations["image_width"] = width
                    annotations["image_path"] = image_path

                # Создание имени JSON-файла
                json_name = f"{os.path.splitext(pdf_file)[0]}_{page_num + 1}.json"
                with open(os.path.join(ANNOTATIONS_DIR, json_name), 'w', encoding='utf-8') as f:
                    json.dump(annotations, f, ensure_ascii=False, indent=4)

        except Exception as e:
            print(f"Ошибка обработки {pdf_file}: {e}")

if __name__ == "__main__":
    extract_annotations()