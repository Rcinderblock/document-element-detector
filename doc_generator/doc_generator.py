import os
import random
import io
from collections import OrderedDict
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import RGBColor, Pt, Inches
from faker import Faker
from mimesis import Text
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
import sympy as sp
import uuid
from docx.enum.style import WD_STYLE_TYPE

# Инициализация Faker для английского и русского языков
fake = Faker(['en_US', 'ru_RU'])
text_mimesis = Text('ru')

# Цвета для таблиц
COLORS = """000000 000080 00008B 0000CD 0000FF 006400 008000 008080 008B8B 00BFFF 00CED1 
00FA9A 00FF00 00FF7F 00FFFF 191970 1E90FF 20B2AA 228B22 2E8B57 2F4F4F 32CD32 
3CB371 40E0E0 4169E1 4682B4 483D8B 48D1CC 4B0082 556B2F 5F9EA0 6495ED 66CDAA 
696969 6A5ACD 6B8E23 708090 778899 7B68EE 7CFC00 7FFF00 7FFFD4 800000 800080 
808000 808080 87CEEB 87CEFA 8A2BE2 8B0000 8B008B 8B4513 8FBC8F 90EE90 9370D8 
9400D3 98FB98 9932CC 9ACD32 A0522D A52A2A A9A9A9 ADD8E6 ADFF2F AFEEEE B0C4DE 
B0E0E6 B22222 B8860B BA55D3 BC8F8F BDB76B C0C0C0 C71585 CD5C5C CD853F D2691E 
D2B48C D3D3D3 D87093 D8BFD8 DA70D6 DAA520 DC143C DCDCDC DDA0DD DEB887 E0FFFF 
E6E6FA E9967A EE82EE EEE8AA F08080 F0E68C F0F8FF F0FFF0 F0FFFF F4A460 F5DEB3 
F5F5DC F5F5F5 F5FFFA F8F8FF FA8072 FAEBD7 FAF0E6 FAFAD2 FDF5E6 FF0000 FF00FF 
FF00FF FF1493 FF4500 FF6347 FF69B4 FF7F50 FF8C00 FFA07A FFA500 FFB6C1 FFC0CB 
FFD700 FFDAB9 FFDEAD FFE4B5 FFE4C4 FFE4E1 FFEBCD FFEFD5 FFF0F5 FFF5EE FFF8DC 
FFFACD FFFAF0 FFFAFA FFFF00 FFFFE0 FFFFF0 FFFFFF"""


def change_orientation(document, landscape=True):
    section = document.sections[-1]
    if landscape and section.orientation == WD_ORIENT.PORTRAIT:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    elif not landscape and section.orientation == WD_ORIENT.LANDSCAPE:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width, section.page_height = section.page_height, section.page_width


def add_header_footer(document, base_font_size):
    section = document.sections[-1]
    header = section.header
    footer = section.footer
    # Генерация текста для колонтитулов
    header_text = fake.sentence(nb_words=5)
    footer_text = fake.sentence(nb_words=5)

    # Верхний
    header_paragraph = header.paragraphs[0]
    header_paragraph.text = header_text
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header_paragraph.runs[0]
    run.font.size = Pt(base_font_size - 4)

    # Нижний
    footer_paragraph = footer.paragraphs[0]
    footer_paragraph.text = footer_text
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_paragraph.runs[0]
    run.font.size = Pt(base_font_size - 4)


def add_heading(document, base_font_size):
    # Генерация заголовка
    level = random.randint(1, 3)  # Уровень заголовка
    if level == 1:
        font_size = base_font_size + random.choice([2, 4, 6, 8])
    elif level == 2:
        font_size = base_font_size + random.choice([2, 4, 6])
    else:
        font_size = base_font_size + random.choice([2, 4])

    heading_text = fake.sentence(nb_words=random.randint(3, 7)) if random.choice([True, False]) else \
    text_mimesis.text(quantity=1).split('.')[0]
    heading = document.add_heading(level=level)
    run = heading.add_run(heading_text)
    run.bold = random.random() < 0.7
    run.italic = random.random() < 0.3
    run.font.size = Pt(font_size)

    # Выравнивание заголовка
    alignment = random.choices(['left', 'center', 'right'], weights=[30, 70, 30])[0]
    if alignment == 'left':
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif alignment == 'center':
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Добавление нумерации заголовка с вероятностью 50%
    if random.random() < 0.5 and level == 1:
        numbering = f"{level}."
        # Вставляем номер перед текстом заголовка
        heading.text = numbering + " " + heading_text


def add_paragraph(document, base_font_size):
    # Генерация абзаца
    paragraph = document.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.left_indent = Inches(0.25) if random.choice([True, False]) else Inches(0)
    paragraph_format.space_before = Pt(random.randint(0, 10))
    paragraph_format.space_after = Pt(random.randint(0, 10))
    paragraph_format.first_line_indent = Inches(0.25) if random.choice([True, False]) else Inches(0)

    # Выравнивание абзаца
    alignment = random.choice([WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY])
    paragraph.alignment = alignment

    # Генерация текста
    if random.choice([True, False]):
        text = fake.text(max_nb_chars=random.randint(1000, 1500))
    else:
        text = text_mimesis.text(quantity=random.randint(15, 20))
    run = paragraph.add_run(text)
    run.font.size = Pt(base_font_size)


def add_table_with_caption(document, base_font_size, table_styles):
    # Генерация подписи для таблицы
    caption_text = fake.sentence(nb_words=random.randint(3, 7))
    caption_size = base_font_size - random.randint(1, 2)  # Размер шрифта подписи на 1-2 меньше основного
    # Размер подписи не меньше 8
    caption_size = max(caption_size, 8)

    if random.choice([True, False]):
        caption_text = f"Табл. {random.randint(1, 100)} - {caption_text}"
    elif random.choice([True, False]):
        caption_text = f"Таблица {random.randint(1, 100)} - {caption_text}"
    else:
        caption_text = f"Таблица. {caption_text}"

    caption_paragraph = document.add_paragraph(caption_text)
    caption_paragraph.style = 'Caption'
    run = caption_paragraph.runs[0]
    run.font.size = Pt(caption_size)
    run.italic = True

    # Выравнивание подписи
    alignment = random.choices(['left', 'center', 'right'], weights=[30, 70, 30])[0]
    if alignment == 'left':
        caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif alignment == 'center':
        caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    num_rows = random.randint(3, 8)
    num_cols = random.randint(3, 6)
    table = document.add_table(rows=num_rows, cols=num_cols)
    table.alignment = random.choice([WD_TABLE_ALIGNMENT.LEFT, WD_TABLE_ALIGNMENT.CENTER, WD_TABLE_ALIGNMENT.RIGHT])
    # Выбор стиля таблицы
    if random.choice([True, False]) and table_styles:
        table.style = random.choice(table_styles)
    else:
        table.style = 'Table Grid'
    for row in table.rows:
        for cell in row.cells:
            text = str(fake.random_number(digits=random.randint(1, 5))) if random.random() < 0.3 else fake.sentence(
                nb_words=random.randint(2, 5))
            cell.text = text
            # Заливка ячеек
            if table.style == 'Table Grid':
                color = random.choice(COLORS.split())
                cell_xml = cell._tc
                cell_properties = cell_xml.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:fill'), color.lower())
                cell_properties.append(shd)
                # Цвет текста
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(0, 0, 0) if random.choice([True, False]) else RGBColor(255,
                                                                                                             255,
                                                                                                                 255)


def add_picture_with_caption(document, base_font_size):
    # Создание графика
    plt.figure(figsize=(4, 3))
    plt.plot([random.randint(0, 10) for _ in range(10)])
    plt.title(fake.sentence(nb_words=random.randint(2, 5)))
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='PNG')
    plt.close()
    img_stream.seek(0)

    # Создание нового параграфа для рисунка
    paragraph = document.add_paragraph()  # Новый параграф для рисунка
    document.add_picture(img_stream, width=Inches(4))

    # Определение позиции подписи (80% - снизу, 20% - сверху)
    position = random.choices(['below', 'above'], weights=[0.8, 0.2])[0]

    # Создание подписи
    caption_prefix = random.choice(["Рис.", "Рисунок", "Figure"])
    picture_number = random.randint(1, 100)
    caption_text = f"{caption_prefix} {picture_number}. {fake.sentence(nb_words=random.randint(3, 7))}"

    # Добавляем подпись в тот же параграф
    if position == 'below':
        # Подпись после рисунка
        caption_paragraph = document.add_paragraph(caption_text)  # Добавляем подпись
    else:
        # Подпись перед рисунком
        caption_paragraph = document.add_paragraph(caption_text)  # Добавляем подпись
        # Добавляем перевод строки для отделения
        paragraph.add_run('\n')

    # Устанавливаем стиль для подписи
    caption_paragraph.style = 'Caption'
    run = caption_paragraph.runs[0]  # Получаем доступ к тексту подписи
    run.font.size = Pt(max(base_font_size - 2, 8))
    run.italic = True

    # Устанавливаем выравнивание для подписи и рисунка
    alignment = random.choices(['left', 'center', 'right'], weights=[30, 70, 30])[0]
    caption_paragraph.alignment = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT
    }[alignment]

    # Выравнивание параграфа с рисунком
    paragraph.alignment = caption_paragraph.alignment


def generate_formula():
    # Генерация случайной формулы с использованием sympy
    formula_type = random.choice(['polynomial', 'trig', 'exponential', 'fraction'])
    x = sp.symbols('x')
    if formula_type == 'polynomial':
        expr = sp.random_poly(x, 3, inf=-10, sup=10).as_expr()
    elif formula_type == 'trig':
        trig_function = random.choice([sp.sin, sp.cos, sp.tan])
        expr = trig_function(x * random.randint(1, 5))
    elif formula_type == 'exponential':
        base = random.randint(2, 5)
        expr = base ** x
    else:  # fraction
        numerator = sp.random_poly(x, 1, inf=-5, sup=5).as_expr()
        denominator = sp.random_poly(x, 1, inf=1, sup=5).as_expr()
        expr = numerator / denominator
    return expr


def add_formula(document, base_font_size):
    if random.choice([True, False]):
        formula_expr = generate_formula()
        # Поскольку python-docx не поддерживает LaTeX, используем простой текст с форматированием
        formula_str = str(formula_expr)

        # Создаем параграф для формулы
        paragraph = document.add_paragraph()
        run = paragraph.add_run()
        run.font.size = Pt(base_font_size)
        run.font.name = 'Cambria Math'  # Специальный шрифт для формул
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Обработка строки формулы для добавления надстрочных символов
        import re
        tokens = re.findall(r'\^(\{[^}]+}|.)', formula_str)
        last_index = 0
        for match in re.finditer(r'\^(\{[^}]+}|.)', formula_str):
            start, end = match.span()
            if start > last_index:
                run.add_text(formula_str[last_index:start])
            superscript_text = match.group(1)
            if superscript_text.startswith('{') and superscript_text.endswith('}'):
                superscript_text = superscript_text[1:-1]
            run_sup = paragraph.add_run(superscript_text)
            run_sup.font.size = Pt(base_font_size * 0.8)  # Уменьшаем размер шрифта для надстрочных
            run_sup.font.superscript = True
            last_index = end
        # Добавляем оставшуюся часть формулы
        if last_index < len(formula_str):
            run.add_text(formula_str[last_index:])


def add_numbered_list(document):
    if random.choice([True, False]):
        list_type = 'List Number'
        num_items = random.randint(3, 7)
        for _ in range(num_items):
            list_item = fake.sentence(nb_words=random.randint(3, 10)) if random.choice([True, False]) else \
            text_mimesis.text(quantity=1).split('.')[0]
            paragraph = document.add_paragraph(list_item, style=list_type)
            paragraph.paragraph_format.left_indent = Inches(0.25)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)


def add_bulleted_list(document):
    if random.choice([True, False]):
        list_type = 'List Bullet'
        num_items = random.randint(3, 7)
        for _ in range(num_items):
            list_item = fake.sentence(nb_words=random.randint(3, 10)) if random.choice([True, False]) else \
            text_mimesis.text(quantity=1).split('.')[0]
            paragraph = document.add_paragraph(list_item, style=list_type)
            paragraph.paragraph_format.left_indent = Inches(0.25)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)




def add_footnote(document, base_font_size):
    # Генерация случайного слова и случайного текста
    random_word = fake.word()
    random_sentence = fake.sentence()

    # Создаем текст сноски в нужном формате
    footnote_text = f"{random_word}* -- {random_sentence}"

    # Создаем новый параграф для сноски
    footnote_paragraph = document.add_paragraph()
    run = footnote_paragraph.add_run(footnote_text)

    # Устанавливаем параметры шрифта
    run.font.size = Pt(max(base_font_size - 2, 8))  # Уменьшаем размер шрифта
    run.italic = True  # Устанавливаем курсив
    run.underline = True  # Подчеркиваем текст

    # Выравниваем сноску влево
    footnote_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

def set_multicolumn(section, num_columns=2):
    sectPr = section._sectPr
    # Создаем элемент для колонок
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), str(num_columns))
    cols.set(qn('w:space'), '720')  # 720 твипов = 0.5 дюйма
    cols.set(qn('w:equalWidth'), '1')  # Колонки равной ширины
    # Удаляем существующие колонки, если есть
    existing_cols = sectPr.find(qn('w:cols'))
    if existing_cols is not None:
        sectPr.remove(existing_cols)
    sectPr.append(cols)


def reset_multicolumn(section):
    sectPr = section._sectPr
    # Удаляем элемент колонок, если существует
    existing_cols = sectPr.find(qn('w:cols'))
    if existing_cols is not None:
        sectPr.remove(existing_cols)
    # Добавляем обратно одну колонку по умолчанию
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '1')
    cols.set(qn('w:space'), '720')  # 720 твипов = 0.5 дюйма
    sectPr.append(cols)


def add_multi_column_text(document, base_font_size):
    # Добавляем секцию с многоколоночным текстом
    multi_col_section = document.add_section(WD_SECTION.NEW_PAGE)
    num_columns = random.randint(2, 3)  # Случайное количество колонок от 2 до 3
    set_multicolumn(multi_col_section, num_columns=num_columns)

    for _ in range(num_columns):
        # Генерация текста для колонки с размером от 1000 до 3000 символов
        column_text = fake.text(max_nb_chars=random.randint(600, 1500))

        # Добавляем новый параграф в документ для текущей колонки
        paragraph = document.add_paragraph(column_text)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY  # Выравнивание по ширине

        # Устанавливаем размер шрифта
        for run in paragraph.runs:
            run.font.size = Pt(base_font_size)




def generate_document(path):
    document = Document()
    # Получение доступных стилей таблиц
    table_styles = [style.name for style in document.styles if style.type == WD_STYLE_TYPE.TABLE]

    # Установка начального шрифта
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    base_font_size = random.randint(8, 16)
    font.size = Pt(base_font_size)

    num_iterations = random.randint(10, 15)  # Количество итераций для генерации содержимого
    for i in range(num_iterations):
        # Добавление новой секции на новой странице, кроме первой
        if i != 0:
            new_sect = document.add_section(WD_SECTION.NEW_PAGE)
            reset_multicolumn(new_sect)

        # Изменение ориентации страницы на вертикальную с вероятностью 85%
        if random.random() < 0.85:
            landscape = random.choice([True, False])
            change_orientation(document, landscape=landscape)

        # Добавление заголовка
        add_heading(document, base_font_size)

        # Добавление колонтитулов
        add_header_footer(document, base_font_size)

        # Добавление абзаца
        add_paragraph(document, base_font_size)

        # Добавление таблицы с подписью
        add_table_with_caption(document, base_font_size, table_styles)


        # Добавление списков
        if random.random() < 0.3:
            add_numbered_list(document)
        elif random.random() < 0.7:
            add_bulleted_list(document)

        # Добавление рисунка с подписью
        add_picture_with_caption(document, base_font_size)

        # Добавление формул
        add_formula(document, base_font_size)

        # Добавление сносок
        add_footnote(document, base_font_size)

        # Добавление много колоночного текста с вероятностью 20%
        if random.random() < 0.4:
            add_multi_column_text(document, base_font_size)

    # Сохранение документа
    document.save(path)