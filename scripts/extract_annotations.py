import os
import re
import json
import fitz  # PyMuPDF
from PIL import Image, ImageDraw


# Папки с данными
PDF_DIR = 'data/pdf/'
ANNOTATIONS_DIR = 'data/annotations/'

# Создание директории для аннотаций, если она не существует
os.makedirs(ANNOTATIONS_DIR, exist_ok=True)


class DocumentAnalyzer:

    def extract_text_from_block(self, block):
        text = ""
        for line in block.get('lines', []):
            for span in line.get('spans', []):
                text += span.get('text', '') + " "
        return text.strip()


    def __init__(self):
        self.colors = {
            'title': (255, 0, 0),  # Red
            'paragraph': (0, 255, 0),  # Green
            'picture': (255, 165, 0),  # Orange
            'table_signature': (128, 0, 128),  # Purple
            'picture_signature': (255, 192, 203),  # Pink
            'numbered_list': (165, 42, 42),  # Brown
            'marked_list': (0, 255, 255),  # Cyan
            'header': (255, 255, 0),  # Yellow
            'footer': (128, 128, 128),  # Gray
            'footnote': (0, 128, 0),  # Dark Green
            'formula': (228, 228, 100)  # Yellow
        }

    def extract_coordinates(self, pdf_path):
        """Извлекает координаты всех элементов, классифицируя их"""
        doc = fitz.open(pdf_path)
        page_data = []

        for page_num in range(len(doc)):
            page = doc[page_num]
            page_dict = {
                'image_height': int(page.rect.height),
                'image_width': int(page.rect.width),
                'image_path': f'page_{page_num + 1}.png',
                'title': [],
                'paragraph': [],
                'table': [],
                'picture': [],
                'table_signature': [],
                'picture_signature': [],
                'numbered_list': [],
                'marked_list': [],
                'header': [],
                'footer': [],
                'footnote': [],
                'formula': [],
                'multicolumn_text': []
            }

            # Все содержимое документа разбивается на блоки
            blocks = page.get_text("dict")["blocks"]

            for block in blocks:
                bbox = tuple(block['bbox'])
                text = self.extract_text_from_block(block)
                block_type = block['type']
                # Если текст пустой (например пропуск строки) и это не картинка
                if not text and block_type != 1:
                    print('NOTICE: text is empty')
                    continue
                # Todo: add TABLE checking and MULTICOLUMN, change TITLE logic
                if self._is_title(block):
                    page_dict['title'].append(self._convert_coordinates(bbox))
                if block_type == 1:  # is formula
                    page_dict['picture'].append(self._convert_coordinates(bbox))
                elif self._is_table_signature(text):
                    page_dict['table_signature'].append(self._convert_coordinates(bbox))
                elif self._is_picture_signature(text):
                    page_dict['picture_signature'].append(self._convert_coordinates(bbox))
                elif self._is_numbered_list(text):
                    page_dict['numbered_list'].append(self._convert_coordinates(bbox))
                elif self._is_marked_list(text):
                    page_dict['marked_list'].append(self._convert_coordinates(bbox))
                elif self._is_footer(bbox, page.rect.height):
                    page_dict['footer'].append(self._convert_coordinates(bbox))
                elif self._is_header(bbox):
                    page_dict['header'].append(self._convert_coordinates(bbox))
                elif self._is_formula(block):
                    page_dict['formula'].append(self._convert_coordinates(bbox))
                elif self._is_footnote(block, page.rect.height):
                    page_dict['footnote'].append(self._convert_coordinates(bbox))
                else:
                    page_dict['paragraph'].append(self._convert_coordinates(bbox))

            page_data.append(page_dict)

        doc.close()
        return page_data


    def _convert_coordinates(self, bbox):
        """Конвертирует координаты в [x1, y1, x2, y2] формат."""
        return [int(bbox[0]), int(bbox[1]), int(bbox[2]), int(bbox[3])]


    def _is_title(self, block):
        return block.get('lines', [{}])[0].get('spans', [{}])[0].get('size', 0) > 12

    def _is_table_signature(self, text):
        return bool(re.match(r'(Табл\.|Таблица|Table)\s*\d+', text, re.IGNORECASE))

    def _is_picture_signature(self, text):
        return bool(re.match(r'(Рис\.|Рисунок|Figure)\s*\d+', text,  re.IGNORECASE))

    def _is_numbered_list(self, text):
        return bool(re.match(r'^\d+\.\s', text))

    def _is_marked_list(self, text):
        return bool(re.match(r'^(\s*[-•*]\s+)', text))


    def _is_footer(self, bbox, page_height):
        return bbox[1] > page_height - 50

    def _is_header(self, bbox):
        return bbox[1] < 50

    import re

    def _is_formula(self, block):
        """
        Проверяет, является ли текст математической формулой.
        Проверка шрифта (например, Cambria Math), а также
        поиск математических символов, греческих букв, чисел и операторов.
        """
        text = self.extract_text_from_block(block)
        fonts = []
        for line in block['lines']:
            for span in line['spans']:
                font_name = span['font']
                fonts.append(font_name)
        # 1. Проверка на использование математического шрифта
        if fonts and 'CambriaMath' in fonts:
            return True

        # 2. Проверка на наличие чисел и математических операторов
        math_operators = r'[+\-*/=()]'
        if re.search(math_operators, text) and any(c.isdigit() for c in text):
            return True

        # 3. Дополнительная проверка на наличие математических символов и греческих букв
        math_symbols = r'[∪∩≈≠∞∑∏√∂∇⊕⊗≡⊂∈∉∫]'  # Символы объединений, пересечений, суммы и другие
        greek_letters = r'[αβγδεζηθικλμνξοπρστυφχψω]'  # Греческие буквы
        if re.search(math_symbols, text) or re.search(greek_letters, text):
            return True

        # 4. Дополнительная проверка для выражений с дробями, квадратными корнями и экспонентами
        advanced_math_pattern = r'(\d+[\+\-\*/^()])+\d+|\d+\.\d+'  # Пример: 3.14+2, 2*(3+4), 3^2
        if re.search(advanced_math_pattern, text):
            return True

        # 5. Если текст содержит логарифмы или интегралы
        advanced_math_keywords = r'log|ln|∫|∑|∏|lim'
        if re.search(advanced_math_keywords, text):
            return True

        return False

    def _is_footnote(self, block, page_height):
        """
        Проверяет, является ли текст сноской.
        Сноски часто содержат номера, могут быть внизу страницы и иметь меньший размер шрифта.
        Также проверяется стиль текста, например, курсив и подчеркивание.
        """

        text = self.extract_text_from_block(block)
        bbox = tuple(block['bbox'])
        fonts = []
        for line in block['lines']:
            for span in line['spans']:
                font_name = span['font']
                fonts.append(font_name)

        # Проверка на присутствие номеров сносок (например, [1], 1), [1.], 1], маленькие цифры и цифры в скобках
        footnote_pattern = r'^\[\^?\d+\]|\d+\)|\d+\.$|\[\d+\]|\d+[\)\.\]]'  # Шаблон для номеров сносок
        if re.search(footnote_pattern, text.strip()):
            return True

        # Проверка на наличие верхних индексов (например, ¹, ², ³, ...)
        superscript_pattern = r'[¹²³⁴⁵⁶⁷⁸⁹]'  # Символы для верхних индексов
        if re.search(superscript_pattern, text):
            return True

        # Проверка на позицию (сноски обычно находятся внизу страницы)
        if bbox and bbox[1] > 0.9 * page_height:  # Нижняя часть страницы
            return True

        # Проверка на стиль текста (курсив, подчеркивание, структура вида [слова]* - [слова])
        italic_pattern = r'\b[а-яА-Яa-zA-Z]*[а-яА-Яa-zA-Z]*\b[\s-]*[а-яА-Яa-zA-Z]*'
        if re.search(italic_pattern, text) and (
                'italic' in fonts or 'underline' in fonts):
            return True

        return False


    def generate_annotated_images(self, pdf_dir, output_dir):
        """Генерирует аннотации, т.е. координаты элементов"""
        os.makedirs(output_dir, exist_ok=True)

        for pdf_file in os.listdir(pdf_dir):
            if not pdf_file.endswith('.pdf'):
                continue

            pdf_path = os.path.join(pdf_dir, pdf_file)
            doc = fitz.open(pdf_path)
            page_data = self.extract_coordinates(pdf_path)

            for page_num, page_info in enumerate(page_data):
                page = doc[page_num]
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                draw = ImageDraw.Draw(img)

                # Рисует прямоугольники для просмотра результата
                for element_type, coordinates_list in page_info.items():
                    if element_type not in ['image_height', 'image_width', 'image_path']:
                        color = self.colors.get(element_type, (0, 0, 0))
                        for coords in coordinates_list:
                            draw.rectangle(coords, outline=color, width=2)

                # Сохраняет картинки
                output_path = os.path.join(output_dir,
                                           f'{os.path.splitext(pdf_file)[0]}_annotated_page_{page_num + 1}.png')
                img.save(output_path)

                # Сохраняет JSON
                json_path = os.path.join(output_dir, f'{os.path.splitext(pdf_file)[0]}_page_{page_num + 1}.json')
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(page_info, f, indent=2)

            doc.close()


def main():
    analyzer = DocumentAnalyzer()
    analyzer.generate_annotated_images(PDF_DIR, ANNOTATIONS_DIR)


if __name__ == "__main__":
    main()
