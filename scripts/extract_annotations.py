import os
import re
import json
from collections import Counter

# не менять на fitz, иначе у Тимофея не зафурычит
import pymupdf
from PIL import Image, ImageDraw

# Папки с данными
PDF_DIR = 'data/pdf/'
ANNOTATIONS_DIR = 'data/annotations/'


class DocumentAnalyzer:

    def extract_text_from_block(self, block):
        text = ""
        for line in block.get('lines', []):
            for span in line['spans']:
                text += span['text'] + " "
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
            'formula': (228, 228, 100),  # Yellow
        }
        self.base_font_size = None
        self.list_mark_left_indent = round(133.3000030517578, 3)
        self.list_text_left_indent = round(151.3000030517578, 3)
        self.line_spacing = None
        self.prev_element = None

    def extract_coordinates(self, pdf_path):
        """Извлекает координаты всех элементов, классифицируя их"""
        doc = pymupdf.open(pdf_path)
        page_data = []
        tables = self.tables_find(pdf_path)
        paragraph_lines_count = 0
        paragraphs_font_sizes = []

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
                'formula': []
            }

            table_areas = tables.get(page_num, [])
            if page_num in tables:
                for bbox in tables[page_num]:
                    page_dict['table'].append(self._convert_coordinates(bbox))

            # Сначала извлекаем текст ПЕРВОЙ страницы для вычисления базового размера шрифта
            if page_num == 0:
                paragraphs = []
                page = doc[0]
                blocks = page.get_text("dict")["blocks"]
                prev_line = None
                line = None

                for block in blocks:
                    block_bbox = block["bbox"]

                    if not self.line_spacing:
                        if not self._is_header(block['bbox']) and not self._is_footer(block['bbox'], page.rect.height) and not self._is_footnote(block, page.rect.height):
                            line = block
                            if not prev_line:
                                prev_line = line
                            else:
                                self.line_spacing = line['bbox'][1] - prev_line['bbox'][3]

                    # Проверяем, находится ли блок в одной из областей таблицы
                    is_in_table_area = any(
                        self._is_within_table(block_bbox, table_area) for table_area in table_areas
                    )

                    # Если блок не в таблице, добавляем его текст для анализа
                    if not is_in_table_area:
                        text = self.extract_text_from_block(block)
                        if text:
                            paragraphs.append(block)

                # Теперь вычисляем базовый размер шрифта на основе этих параграфов
                if self.base_font_size is None:
                    self.base_font_size = self.get_base_font_size(paragraphs)

            # Все содержимое документа разбивается на блоки
            blocks = page.get_text("dict")["blocks"]

            for block in blocks:
                bbox = tuple(block['bbox'])
                text = self.extract_text_from_block(block)
                block_type = block['type']

                # Если текст пустой (например пропуск строки) и это не картинка
                if not text and block_type != 1:
                    continue

                if any(self._is_within_table(bbox, table) for table in table_areas):
                    continue
                    
                element_name = None
                if block_type == 1:  # is_picture
                    page_dict['picture'].append(self._convert_coordinates(bbox))
                    element_name = 'picture'
                elif self._is_table_signature(text):
                    page_dict['table_signature'].append(self._convert_coordinates(bbox))
                    element_name = 'table_signature'
                elif self._is_picture_signature(text):
                    page_dict['picture_signature'].append(self._convert_coordinates(bbox))
                    element_name = 'picture_signature'
                elif self._is_header(bbox):
                    page_dict['header'].append(self._convert_coordinates(bbox))
                    element_name = 'header'
                elif self._is_numbered_list(block):
                    page_dict['numbered_list'].append(self._convert_coordinates(bbox))
                    element_name = 'numbered_list'
                elif self._is_marked_list(block):
                    page_dict['marked_list'].append(self._convert_coordinates(bbox))
                    element_name = 'marked_list'
                elif self._is_footer(bbox, page.rect.height):
                    page_dict['footer'].append(self._convert_coordinates(bbox))
                    element_name = 'footer'
                elif self._is_formula(block)[0]:
                    coords = self._convert_coordinates(bbox)
                    # Если в формуле есть знак суммы, интегралы или знак произведения
                    # Тогда раздвинем границы вверх и вниз на 5, чтобы они вместились в бокс
                    if self._is_formula(block)[1]:
                        coords[1] -= 5
                        coords[-1] += 5
                    page_dict['formula'].append(coords)
                    element_name = 'formula'
                elif self._is_footnote(block, page.rect.height):
                    page_dict['footnote'].append(self._convert_coordinates(bbox))
                    element_name = 'footnote'
                elif self._is_title(block):
                    page_dict['title'].append(self._convert_coordinates(bbox))
                    element_name = 'title'
                else:
                    page_dict['paragraph'].append(self._convert_coordinates(bbox))
                    element_name = 'paragraph'
                    # Складываем шрифты параграфов
                    for line in block['lines']:
                        for span in line['spans']:
                            paragraphs_font_sizes.append(round(span.get('size')))
                        paragraph_lines_count += 1
                    # Обновляем базовый размер шрифта, если собрали уже достаточно и если обновление было не так давно
                    if paragraph_lines_count > 50 and len(paragraphs_font_sizes) > 5:
                        self.base_font_size = self.update_base_font_size(paragraphs_font_sizes)

                self.prev_element = { 
                    'name': element_name,
                    'text': text,
                    'bbox': bbox,
                }

            # Слияние боксов формул если они рядом или пересекаются
            page_dict['formula'] = self.merge_rects(page_dict['formula'], max_y_distance=5, max_x_distance=5)

            # Слияние боксов текстов (УДАЛИТЬ ЕСЛИ НЕОБХОДИМО)
            page_dict['paragraph'] = self.merge_rects(page_dict['paragraph'], max_y_distance=6, max_x_distance=0)

            # Слияние боксов сносок (УДАЛИТЬ ЕСЛИ НЕОБХОДИМО)
            page_dict['footnote'] = self.merge_rects(page_dict['footnote'], max_y_distance=5, max_x_distance=0)

            page_dict['numbered_list'] = self.merge_rects(page_dict['numbered_list'], max_y_distance=3, max_x_distance=0)

            page_dict['marked_list'] = self.merge_rects(page_dict['marked_list'], max_y_distance=3, max_x_distance=0)

            page_data.append(page_dict)

        doc.close()
        return page_data

    def merge_rects(self, rectangles, max_y_distance=10, max_x_distance=20):
        def rectangles_intersect_or_nearby(rect1, rect2):
            # Проверка стандартного пересечения
            if not (rect1[2] < rect2[0] or rect2[2] < rect1[0] or rect1[3] < rect2[1] or rect2[3] < rect1[1]):
                return True

            # Проверка расстояния по оси Y
            if abs(rect1[3] - rect2[1]) <= max_y_distance or abs(rect2[3] - rect1[1]) <= max_y_distance:
                return True

            # Проверка расстояния по оси X
            if abs(rect1[2] - rect2[0]) <= max_x_distance or abs(rect2[2] - rect1[0]) <= max_x_distance:
                return True

            return False

        def merge_rectangles(rect1, rect2):
            # Объединение двух пересекающихся или близко расположенных прямоугольников
            x1 = min(rect1[0], rect2[0])
            y1 = min(rect1[1], rect2[1])
            x2 = max(rect1[2], rect2[2])
            y2 = max(rect1[3], rect2[3])
            return [x1, y1, x2, y2]

        merged_rects = []
        while rectangles:
            # Берем первый прямоугольник и ищем для него все пересекающиеся или близкие
            base_rect = rectangles.pop(0)
            merged = True
            while merged:
                merged = False
                i = 0
                while i < len(rectangles):
                    if rectangles_intersect_or_nearby(base_rect, rectangles[i]):
                        # Если пересечение или близость найдены, объединяем и убираем из списка
                        base_rect = merge_rectangles(base_rect, rectangles.pop(i))
                        merged = True  # Указываем, что произошло объединение
                    else:
                        i += 1
            # Добавляем объединенный прямоугольник в результат
            merged_rects.append(base_rect)
        return merged_rects

    def _convert_coordinates(self, bbox):
        """Конвертирует координаты в [x1, y1, x2, y2] формат."""
        return [int(bbox[0]), int(bbox[1]), int(bbox[2]), int(bbox[3])]

    def _is_title(self, block):
        """Определяет, является ли блок заголовком на основе его шрифта."""
        title_font_size = self.base_font_size + 2
        is_bold = False
        font_sizes = []
        # Собираем все размеры шрифта
        for line in block['lines']:
            for span in line['spans']:
                font_size = span.get('size')
                if not is_bold:
                    is_bold = 'bold' in span['font'].lower()
                if font_size:  # Если размер шрифта найден
                    font_size = round(font_size)
                    font_sizes.append(font_size)

        # Если нет шрифтов, возвращаем False
        if not font_sizes or not is_bold:
            return False

        # Находим самый популярный размер шрифта
        font_size_counter = Counter(font_sizes)
        most_common_font_size, most_common_count = font_size_counter.most_common(1)[0]
        # Проверяем, является ли этот размер шрифта заголовком
        # размер шрифта = 26 здесь как флаг заголовка 0 уровня. Иначе какие-то баги с ним
        if most_common_font_size >= title_font_size or 26 in font_size_counter:
            return True
        return False

    def get_base_font_size(self, paragraphs):
        """Вычисляет наиболее часто встречающийся размер шрифта для первых параграфов."""
        font_sizes = []
        for block in paragraphs:
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    font_sizes.append(round(span.get('size')))

        # Используем Counter для подсчета частоты встречающихся размеров шрифта
        font_size_counts = Counter(font_sizes)

        # Находим наиболее часто встречающийся размер шрифта
        if font_size_counts:
            most_common_size = font_size_counts.most_common(1)[0][0]
            return most_common_size

        return 12

    def update_base_font_size(self, paragraphs_font_sizes):
        """Обновляет размер шрифта, если уже классифицированы параграфы"""
        font_size_counts = Counter(paragraphs_font_sizes)

        # Находим наиболее часто встречающийся размер шрифта
        if font_size_counts:
            most_common_size = font_size_counts.most_common(1)[0][0]
            return most_common_size

        return None

    def _is_table_signature(self, text):
        pattern = r'(Табл\.|Таблица|Table|Таблица\.|Табл)\s*\d*'
        if text[0].isupper() and bool(re.match(pattern, text, re.IGNORECASE)):
            modified_text = re.sub(pattern, '', text, flags=re.IGNORECASE).strip()
            if modified_text[0] in ['-', '.', '—']:
                return True
        return False

    def _is_picture_signature(self, text):
        pattern = r'(Рис\.|Рисунок|Figure|Рис|Рисунок\.)\s*\d*'
        if text[0].isupper() and bool(re.match(pattern, text, re.IGNORECASE)):
            modified_text = re.sub(pattern, '', text, flags=re.IGNORECASE).strip()
            if modified_text[0] in ['-', '.', '—']:
                return True
        return False

    def _is_numbered_list(self, block):
        is_bold = False
        for line in block['lines']:
            for span in line['spans']:
                if not is_bold:
                    is_bold = 'bold' in span['font'].lower()
        if is_bold:
            return False
        
        text = self.extract_text_from_block(block)
        bbox = tuple(block['bbox'])
        numbered_pattern = r'^\d+\.\s'
        left_indent = round(bbox[0], 3)
        if bool(re.match(numbered_pattern, text)) and left_indent == self.list_mark_left_indent:
           return True
        
        if not self.prev_element:
            return False

        space = bbox[1] - self.prev_element['bbox'][3]

        if self.prev_element['name'] == 'numbered_list' and space < self.line_spacing and left_indent == self.list_text_left_indent:
            return True
        
        return False

    def _is_marked_list(self, block):
        text = self.extract_text_from_block(block)
        bbox = tuple(block['bbox'])
        numbered_pattern = r'^(\s*[-•*–·]\s+)'
        left_indent = round(bbox[0], 3)
        if bool(re.match(numbered_pattern, text)) and left_indent == self.list_mark_left_indent:
           return True
        
        if not self.prev_element:
            return False

        space = bbox[1] - self.prev_element['bbox'][3]

        if self.prev_element['name'] == 'marked_list' and space < self.line_spacing and left_indent == self.list_text_left_indent:
            return True
        
        return False

    def _is_footer(self, bbox, page_height):
        return bbox[1] > page_height - 60

    def _is_header(self, bbox):
        return bbox[1] < 60

    def _is_formula(self, block) -> tuple[bool, bool]:
        """
        Проверяет, является ли текст математической формулой.
        Проверка шрифта (например, Cambria Math), а также
        поиск математических символов, греческих букв, чисел и операторов.

        Returns:
            Первое bool значение -- является ли блок формулой
            Второе bool значение -- нужно ли поднимать границы

        """
        text = self.extract_text_from_block(block)
        fonts = []
        up_borders_symbols = r'[∑∏∫]'

        for line in block['lines']:
            for span in line['spans']:
                font_name = span['font']
                fonts.append(font_name)
        # 1. Проверка на использование математического шрифта
        if fonts and 'CambriaMath' in fonts:
            if re.search(up_borders_symbols, text):
                return True, True
            return True, False

        # 2. Проверка на наличие чисел и математических операторов
        math_operators = r'[+\*/=()]'
        if re.search(math_operators, text) and any(c.isdigit() for c in text):
            if re.search(up_borders_symbols, text):
                return True, True
            return True, False

        # 3. Дополнительная проверка на наличие математических символов и греческих букв
        math_symbols = r'[∪∩≈≠∞∑∏√∂∇⊕⊗≡⊂∈∉∫]'  # Символы объединений, пересечений, суммы и другие
        greek_letters = r'[αβγδεζηθικλμνξοπρστυφχψω]'  # Греческие буквы
        if re.search(math_symbols, text) or re.search(greek_letters, text):
            if re.search(up_borders_symbols, text):
                return True, True
            return True, False

        return False, False

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
        footnote_pattern = r'^\[\^?\d+\]|\d+\)|\d+\.$|\[\d+\]|\d+[\)\]]'  # Шаблон для номеров сносок
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

    def tables_find(self, file_path):
        """Находит таблицы в документе и возвращает их координаты по страницам."""
        tables = {}
        doc = pymupdf.open(file_path)

        for page_num in range(len(doc)):
            page = doc[page_num]
            tabs = page.find_tables()  # Находим таблицы на текущей странице
            coordinates = [tab.bbox for tab in tabs]  # Получаем координаты таблиц
            if coordinates:
                tables[page_num] = coordinates
        doc.close()
        return tables

    def _is_within_table(self, bbox, table_bbox):
        """Проверяет, находится ли блок в пределах координат таблицы."""
        bx0, by0, bx1, by1 = bbox
        tx0, ty0, tx1, ty1 = table_bbox
        return bx0 >= tx0 and bx1 <= tx1 and by0 >= ty0 and by1 <= ty1

    def generate_annotated_images(self, pdf_dir, output_dir):
        """Генерирует аннотации, т.е. координаты элементов"""
        os.makedirs(output_dir, exist_ok=True)

        for pdf_file in os.listdir(pdf_dir):
            if not pdf_file.endswith('.pdf'):
                continue

            pdf_path = os.path.join(pdf_dir, pdf_file)
            doc = pymupdf.open(pdf_path)
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
