import os
import re
import json
from collections import Counter

# import fitz  # PyMuPDF # не трогать, иначе у тимофея не зафурычит
import pymupdf
from PIL import Image, ImageDraw

# Папки с данными
PDF_DIR = 'data/pdf/'
ANNOTATIONS_DIR = 'data/annotations/'

# Создание директории для аннотаций, если она не существует
os.makedirs(ANNOTATIONS_DIR, exist_ok=True)


def column_boxes(page, footer_margin=50, header_margin=50, no_image_text=True):
    """Determine bboxes which wrap a column."""
    paths = page.get_drawings()
    bboxes = []

    # path rectangles
    path_rects = []

    # image bboxes
    img_bboxes = []

    # bboxes of non-horizontal text
    # avoid when expanding horizontal text boxes
    vert_bboxes = []

    # compute relevant page area
    clip = +page.rect
    clip.y1 -= footer_margin  # Remove footer area
    clip.y0 += header_margin  # Remove header area

    def can_extend(temp, bb, bboxlist):
        """Determines whether rectangle 'temp' can be extended by 'bb'
        without intersecting any of the rectangles contained in 'bboxlist'.

        Items of bboxlist may be None if they have been removed.

        Returns:
            True if 'temp' has no intersections with items of 'bboxlist'.
        """
        for b in bboxlist:
            if not intersects_bboxes(temp, vert_bboxes) and (
                    b == None or b == bb or (temp & b).is_empty
            ):
                continue
            return False

        return True

    def in_bbox(bb, bboxes):
        """Return 1-based number if a bbox contains bb, else return 0."""
        for i, bbox in enumerate(bboxes):
            if bb in bbox:
                return i + 1
        return 0

    def intersects_bboxes(bb, bboxes):
        """Return True if a bbox intersects bb, else return False."""
        for bbox in bboxes:
            if not (bb & bbox).is_empty:
                return True
        return False

    def extend_right(bboxes, width, path_bboxes, vert_bboxes, img_bboxes):
        """Extend a bbox to the right page border.

        Whenever there is no text to the right of a bbox, enlarge it up
        to the right page border.

        Args:
            bboxes: (list[IRect]) bboxes to check
            width: (int) page width
            path_bboxes: (list[IRect]) bboxes with a background color
            vert_bboxes: (list[IRect]) bboxes with vertical text
            img_bboxes: (list[IRect]) bboxes of images
        Returns:
            Potentially modified bboxes.
        """
        for i, bb in enumerate(bboxes):
            # do not extend text with background color
            if in_bbox(bb, path_bboxes):
                continue

            # do not extend text in images
            if in_bbox(bb, img_bboxes):
                continue

            # temp extends bb to the right page border
            temp = +bb
            temp.x1 = width

            # do not cut through colored background or images
            if intersects_bboxes(temp, path_bboxes + vert_bboxes + img_bboxes):
                continue

            # also, do not intersect other text bboxes
            check = can_extend(temp, bb, bboxes)
            if check:
                bboxes[i] = temp  # replace with enlarged bbox

        return [b for b in bboxes if b != None]

    def clean_nblocks(nblocks):
        """Do some elementary cleaning."""

        # 1. remove any duplicate blocks.
        blen = len(nblocks)
        if blen < 2:
            return nblocks
        start = blen - 1
        for i in range(start, -1, -1):
            bb1 = nblocks[i]
            bb0 = nblocks[i - 1]
            if bb0 == bb1:
                del nblocks[i]

        # 2. repair sequence in special cases:
        # consecutive bboxes with almost same bottom value are sorted ascending
        # by x-coordinate.
        y1 = nblocks[0].y1  # first bottom coordinate
        i0 = 0  # its index
        i1 = -1  # index of last bbox with same bottom

        # Iterate over bboxes, identifying segments with approx. same bottom value.
        # Replace every segment by its sorted version.
        for i in range(1, len(nblocks)):
            b1 = nblocks[i]
            if abs(b1.y1 - y1) > 10:  # different bottom
                if i1 > i0:  # segment length > 1? Sort it!
                    nblocks[i0: i1 + 1] = sorted(
                        nblocks[i0: i1 + 1], key=lambda b: b.x0
                    )
                y1 = b1.y1  # store new bottom value
                i0 = i  # store its start index
            i1 = i  # store current index
        if i1 > i0:  # segment waiting to be sorted
            nblocks[i0: i1 + 1] = sorted(nblocks[i0: i1 + 1], key=lambda b: b.x0)
        return nblocks

    # extract vector graphics
    for p in paths:
        path_rects.append(p["rect"].irect)
    path_bboxes = path_rects

    # sort path bboxes by ascending top, then left coordinates
    path_bboxes.sort(key=lambda b: (b.y0, b.x0))

    # bboxes of images on page, no need to sort them
    for item in page.get_images():
        img_bboxes.extend(page.get_image_rects(item[0]))

    # blocks of text on page
    blocks = page.get_text(
        "dict",
        flags=pymupdf.TEXTFLAGS_TEXT,
        clip=clip,
    )["blocks"]

    # Make block rectangles, ignoring non-horizontal text
    for b in blocks:
        bbox = pymupdf.IRect(b["bbox"])  # bbox of the block

        # ignore text written upon images
        if no_image_text and in_bbox(bbox, img_bboxes):
            continue

        # confirm first line to be horizontal
        line0 = b["lines"][0]  # get first line
        if line0["dir"] != (1, 0):  # only accept horizontal text
            vert_bboxes.append(bbox)
            continue

        srect = pymupdf.EMPTY_IRECT()
        for line in b["lines"]:
            lbbox = pymupdf.IRect(line["bbox"])
            text = "".join([s["text"].strip() for s in line["spans"]])
            if len(text) > 1:
                srect |= lbbox
        bbox = +srect

        if not bbox.is_empty:
            bboxes.append(bbox)

    # Sort text bboxes by ascending background, top, then left coordinates
    bboxes.sort(key=lambda k: (in_bbox(k, path_bboxes), k.y0, k.x0))

    # Extend bboxes to the right where possible
    bboxes = extend_right(
        bboxes, int(page.rect.width), path_bboxes, vert_bboxes, img_bboxes
    )

    # immediately return of no text found
    if bboxes == []:
        return []

    # --------------------------------------------------------------------
    # Join bboxes to establish some column structure
    # --------------------------------------------------------------------
    # the final block bboxes on page
    nblocks = [bboxes[0]]  # pre-fill with first bbox
    bboxes = bboxes[1:]  # remaining old bboxes

    for i, bb in enumerate(bboxes):  # iterate old bboxes
        check = False  # indicates unwanted joins

        # check if bb can extend one of the new blocks
        for j in range(len(nblocks)):
            nbb = nblocks[j]  # a new block

            # never join across columns
            if bb == None or nbb.x1 < bb.x0 or bb.x1 < nbb.x0:
                continue

            # never join across different background colors
            if in_bbox(nbb, path_bboxes) != in_bbox(bb, path_bboxes):
                continue

            temp = bb | nbb  # temporary extension of new block
            check = can_extend(temp, nbb, nblocks)
            if check == True:
                break

        if not check:  # bb cannot be used to extend any of the new bboxes
            nblocks.append(bb)  # so add it to the list
            j = len(nblocks) - 1  # index of it
            temp = nblocks[j]  # new bbox added

        # check if some remaining bbox is contained in temp
        check = can_extend(temp, bb, bboxes)
        if check == False:
            nblocks.append(bb)
        else:
            nblocks[j] = temp
        bboxes[i] = None

    # do some elementary cleaning
    nblocks = clean_nblocks(nblocks)

    # return identified text bboxes
    return nblocks

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
            'multicolumn_text': (0, 0, 0) # Black
        }
        self.base_font_size = None

    def extract_coordinates(self, pdf_path):
        """Извлекает координаты всех элементов, классифицируя их"""
        doc = pymupdf.open(pdf_path)
        page_data = []
        tables = self.tables_find(pdf_path)
        multi_columns = self.check_multi_column_text(pdf_path)
        paragraph_lines_count = 0

        for page_num in range(len(doc)):
            paragraphs_font_sizes = []
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

            table_areas = tables.get(page_num, [])
            if page_num in tables:
                for bbox in tables[page_num]:
                    page_dict['table'].append(self._convert_coordinates(bbox))

            text_areas = multi_columns.get(page_num, [])
            if page_num in multi_columns:
                for bbox in multi_columns[page_num]:
                    page_dict['multicolumn_text'].append(self._convert_coordinates(bbox))

            # Сначала извлекаем текст ПЕРВОЙ страницы для вычисления базового размера шрифта
            if page_num == 0:
                paragraphs = []
                page = doc[0]
                blocks = page.get_text("dict")["blocks"]

                for block in blocks:
                    block_bbox = block["bbox"]

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

                if any(self._is_within_table(bbox, table) for table in table_areas):
                    print('NOTICE: block in table')
                    continue

                # Если текст пустой (например пропуск строки) и это не картинка
                if not text and block_type != 1:
                    print('NOTICE: text is empty')
                    continue

                # Todo: add MULTICOLUMN
                if block_type == 1:  # is_picture
                    page_dict['picture'].append(self._convert_coordinates(bbox))
                elif self._is_table_signature(text):
                    page_dict['table_signature'].append(self._convert_coordinates(bbox))
                elif self._is_picture_signature(text):
                    page_dict['picture_signature'].append(self._convert_coordinates(bbox))
                elif self._is_header(bbox):
                    page_dict['header'].append(self._convert_coordinates(bbox))
                elif self._is_numbered_list(text):
                    page_dict['numbered_list'].append(self._convert_coordinates(bbox))
                elif self._is_marked_list(text):
                    page_dict['marked_list'].append(self._convert_coordinates(bbox))
                elif self._is_footer(bbox, page.rect.height):
                    page_dict['footer'].append(self._convert_coordinates(bbox))
                elif self._is_formula(block):
                    page_dict['formula'].append(self._convert_coordinates(bbox))
                elif self._is_footnote(block, page.rect.height):
                    page_dict['footnote'].append(self._convert_coordinates(bbox))
                elif self._is_title(block):
                    page_dict['title'].append(self._convert_coordinates(bbox))
                else:
                    page_dict['paragraph'].append(self._convert_coordinates(bbox))
                    # Складываем шрифты параграфов
                    for line in block['lines']:
                        for span in line['spans']:
                            paragraphs_font_sizes.append(round(span.get('size')))
                        paragraph_lines_count += 1
                    # Обновляем базовый размер шрифта, если собрали уже достаточно и если обновление было не так давно
                    if paragraph_lines_count > 50 and len(paragraphs_font_sizes) > 5:
                        self.base_font_size = self.update_base_font_size(paragraphs_font_sizes)
                        paragraph_lines_count = 0
                        print(self.base_font_size)

            page_data.append(page_dict)

        doc.close()
        return page_data

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
        if not font_sizes:
            return False

        # Находим самый популярный размер шрифта
        font_size_counter = Counter(font_sizes)
        most_common_font_size, most_common_count = font_size_counter.most_common(1)[0]

        # Проверяем, является ли этот размер шрифта заголовком
        # размер шрифта = 26 здесь как флаг заголовка 0 уровня. Иначе какие-то баги с ним
        if is_bold and (most_common_font_size >= title_font_size or 26 in font_size_counter):
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
            print('ПОПУЛЯРНЕЙШИЙ ШРИФТ СНАЧАЛА:', font_size_counts.items())
            return most_common_size

        return 12

    def update_base_font_size(self, paragraphs_font_sizes):
        """Обновляет размер шрифта, если уже классифицированы параграфы"""
        font_size_counts = Counter(paragraphs_font_sizes)

        # Находим наиболее часто встречающийся размер шрифта
        if font_size_counts:
            most_common_size = font_size_counts.most_common(1)[0][0]
            print('ПОПУЛЯРНЕЙШИЙ ШРИФТ ПАРАГРАФОВ:', font_size_counts.items())
            return most_common_size

        return None

    def _is_table_signature(self, text):
        return bool(re.match(r'(Табл\.|Таблица|Table|Таблица\.|Табл)\s*\d*', text, re.IGNORECASE))

    def _is_picture_signature(self, text):
        return bool(re.match(r'(Рис\.|Рисунок|Figure|Рис|Рисунок\.)\s*\d*', text, re.IGNORECASE))

    def _is_numbered_list(self, text):
        return bool(re.match(r'^\d+\.\s', text))

    def _is_marked_list(self, text):
        return bool(re.match(r'^(\s*[-•*]\s+)', text))

    def _is_footer(self, bbox, page_height):
        return bbox[1] > page_height - 50

    def _is_header(self, bbox):
        return bbox[1] < 50

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

    def check_multi_column_text(self, file_path):
        doc = pymupdf.open(file_path)
        text = dict()
        for page in range(len(doc)):
            bboxes = column_boxes(doc[page], footer_margin=50, no_image_text=True)
            p = []
            for rect in range(len(bboxes) - 1):
                if 85 < bboxes[rect][0] < 95 and 295 < bboxes[rect][2] < 315 and 305 < bboxes[rect + 1][
                    0] < 315 and 510 < \
                        bboxes[rect + 1][2] < 530 and abs(bboxes[rect + 1][1] - bboxes[rect][1]) < 5:
                    p.append(bboxes[rect])
                    p.append(bboxes[rect + 1])
                if rect < len(bboxes) - 2:
                    if 85 < bboxes[rect][0] < 95 and 210 < bboxes[rect][2] < 230 and \
                            230 < bboxes[rect + 1][0] < 250 and 370 < bboxes[rect + 1][2] < 385 and \
                            370 < bboxes[rect + 2][0] < 390 and 505 < bboxes[rect + 2][2] < 530 and \
                            abs(max(bboxes[rect + 1][1] - bboxes[rect][1], bboxes[rect + 2][1] - bboxes[rect][1],
                                    bboxes[rect + 1][1] - bboxes[rect + 2][1])) < 5:
                        p.append(bboxes[rect])
                        p.append(bboxes[rect + 1])
                        p.append(bboxes[rect + 2])
            if len(p) > 0:
                text[page] = p
        # return словарь - страница: точки прямоульника
        return text

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
