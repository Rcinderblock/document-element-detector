from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from mimesis import Text
from docx.enum.style import WD_STYLE_TYPE
import math2docx
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import io
import random
from faker import Faker
import pandas as pd

from mimesis import Generic
from mimesis.locales import Locale
from mimesis.random import Random

# Faker будем использовать и для английского и для русского, т.к. в примере документа оба языка
fake = Faker(['en_US', 'ru_RU'])
text_mimesis = Text('ru')
NUM_ITERATIONS = 100  # Количество итераций для генерации содержимого

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
    """ Функция смены ориентации листа"""
    section = document.sections[-1]
    if landscape and section.orientation == WD_ORIENT.PORTRAIT:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    elif not landscape and section.orientation == WD_ORIENT.LANDSCAPE:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width, section.page_height = section.page_height, section.page_width


def add_header_footer(document, base_font_size):
    """
    Функция добавления колонтитулов. Была рассчитана на многократный вызов, но в итоге применяется один раз.
    :param document: cсылка на документ
    :param base_font_size: базовый шрифт для стандартного текста листа
    От него будем выбирать размер шрифта колонтитулов.
    """
    section = document.sections[-1]
    header = section.header
    footer = section.footer
    header_text = fake.sentence(nb_words=5)
    footer_text = fake.sentence(nb_words=5)

    # Верхний колонтитул
    header_paragraph = header.paragraphs[0]
    header_paragraph.text = header_text
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header_paragraph.runs[0]
    run.font.size = Pt(base_font_size - 4)

    # Нижний колонтитул
    footer_paragraph = footer.paragraphs[0]
    footer_paragraph.text = footer_text
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_paragraph.runs[0]
    run.font.size = Pt(base_font_size - 4)


class RandomHeadingFormat:
    def __init__(self, random_paragraph_format):
        self.font_size_1 = random_paragraph_format.font_size + random.choice([2, 4, 6, 8])
        self.font_size_2 = random_paragraph_format.font_size + random.choice([2, 4, 6])
        self.font_size_3 = random_paragraph_format.font_size + random.choice([2, 4])
        self.bold = random.random() < 0.7
        self.italic = random.random() < 0.3
        self.alignment = random.choices([WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT], weights=[30, 70, 30])[0]


def add_heading(document, random_heading_format):
    """
    Функция с добавлением заголовка
    """
    # Генерация заголовка
    level = random.randint(0, 3)  # Уровень заголовка
    if level == 1:
        font_size = random_heading_format.font_size_1
    elif level == 2:
        font_size = random_heading_format.font_size_2
    else:
        font_size = random_heading_format.font_size_3

    heading_text = fake.sentence(nb_words=random.randint(3, 7)) if random.choice([True, False]) else \
        text_mimesis.text(quantity=1).split('.')[0]  # Рандомная генерация фейкером / мимезисом
    heading = document.add_heading(level=level)
    run = heading.add_run(heading_text)
    run.bold = random_heading_format.bold
    run.italic = random_heading_format.italic
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(0, 0, 0)

    # Выравнивание заголовка (по условию задачи чаще центр)
    alignment = random_heading_format.alignment

    # Добавление нумерации заголовка с вероятностью 50%
    if random.random() < 0.5 and level == 1:
        numbering = f"{random.randint(1, 15)}."
        # Вставляем номер перед текстом заголовка
        heading.text = numbering + " " + heading_text


class RandomParagraphFormat:
    def __init__(self):
        self.left_indent = Inches(0.25) if random.choice([True, False]) else Inches(0)
        self.space_before = random.randint(0, 10)
        self.space_after = random.randint(0, 10)
        self.first_line_indent = Inches(0.25) if random.choice([True, False]) else Inches(0)
        self.alignment = random.choice([WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.JUSTIFY])
        self.font_size = random.randint(8, 14)
        self.font_name = random.choice(['Times New Roman', 'Arial', 'Calibri', 'Georgia', 'Verdana', 'Tahoma', 'Garamond', 'Helvetica', 'Courier New', 'Trebuchet MS', 'Comic Sans'])


def add_paragraph(document, random_paragraph_format):
    """ Функция генерации обычного текста """
    # Генерация абзаца
    paragraph = document.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.left_indent = random_paragraph_format.left_indent  # Рандомная красная строка
    paragraph_format.space_before = Pt(random_paragraph_format.space_before)
    paragraph_format.space_after = Pt(random_paragraph_format.space_after)
    paragraph_format.first_line_indent = random_paragraph_format.first_line_indent

    # Выравнивание абзаца
    paragraph.alignment = random_paragraph_format.alignment

    # Генерация текста для абзаца
    if random.choice([True, False]):
        text = fake.text(max_nb_chars=400)
    else:
        text = text_mimesis.text(quantity=random.randint(5, 10))
    run = paragraph.add_run(text)
    run.font.size = Pt(random_paragraph_format.font_size)
    run.font.name = random_paragraph_format.font_name


class RandomTableFormat:
    def __init__(self, random_paragraph_format, table_styles):
        self.caption_size = max(random_paragraph_format.font_size - random.randint(1, 2), 8)
        self.caption_alignment = random.choice([WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.JUSTIFY])
        self.caption_position = random.choice([0, 1])
        self.caption_name = random.choice(['Times New Roman', 'Arial', 'Calibri', 'Georgia', 'Verdana', 'Tahoma', 'Garamond', 'Helvetica', 'Courier New', 'Trebuchet MS', 'Comic Sans'])
        self.alignment = random.choices([WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER], weights=[30, 70, 30])[0]
        if random.choice([True, False]) and table_styles:
            self.table_style = random.choice(table_styles)
        else:
            self.table_style = 'Table Grid'


def add_table_with_caption(document, random_table_format):
    """ Генерация таблицы вместе с подписью сразу """
    # Генерация подписи для таблицы (первее, т.к. сначала подпись, потом -- таблица)
    caption_text = fake.sentence(nb_words=random.randint(3, 7))

    if random.choice([True, False]):
        caption_text = f"Табл. {random.randint(1, 100)} - {caption_text}"
    elif random.choice([True, False]):
        caption_text = f"Таблица {random.randint(1, 100)} - {caption_text}"
    else:
        caption_text = f"Таблица. {caption_text}"

    if random_table_format.caption_position == 0:
        caption_paragraph = document.add_paragraph(caption_text)
        caption_paragraph.style = 'Caption'
        run = caption_paragraph.runs[0]
        run.font.size = Pt(random_table_format.caption_size)
        run.alignment = random_table_format.caption_alignment
        run.font.name = random_table_format.caption_name
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.alignment = random_table_format.alignment

    num_rows = random.randint(3, 8)
    num_cols = random.randint(3, 6)
    table = document.add_table(rows=num_rows, cols=num_cols)
    table.alignment = random_table_format.alignment

    if random_table_format.caption_position == 1:
        caption_paragraph = document.add_paragraph(caption_text)
        caption_paragraph.style = 'Caption'
        run = caption_paragraph.runs[0]
        run.font.size = Pt(random_table_format.caption_size)
        run.alignment = random_table_format.caption_alignment
        run.font.name = random_table_format.caption_name
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.alignment = random_table_format.alignment

    # Выбор стиля таблицы
    table_style = random_table_format.table_style

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


def generate_plot_or_chart():
    chart_type = random.choice(['line', 'bar', 'scatter', 'pie'])
    fig, ax = plt.subplots()
    if chart_type == 'line':
        x = range(10)
        y = [random.randint(1, 100) for _ in x]
        ax.plot(x, y, color=random.choice(['red', 'green', 'blue']), linestyle=random.choice(['-', '--', '-.', ':']),
                linewidth=random.uniform(0.5, 2.5))
        ax.set_title(fake.sentence(nb_words=random.randint(1, 4)))
        ax.set_xlabel(fake.word())
        ax.set_ylabel(fake.word())

    elif chart_type == 'bar':
        x = range(5)
        y = [random.randint(1, 100) for _ in x]
        ax.bar(x, y, color=random.choice(['red', 'green', 'blue']))
        ax.set_title(fake.sentence(nb_words=random.randint(1, 4)))
        ax.set_xlabel(fake.word())
        ax.set_ylabel(fake.word())

    elif chart_type == 'scatter':
        x = [random.uniform(0, 100) for _ in range(20)]
        y = [random.uniform(0, 100) for _ in range(20)]
        ax.scatter(x, y, color=random.choice(['red', 'green', 'blue']), s=random.randint(10, 200))
        ax.set_title(fake.sentence(nb_words=random.randint(1, 4)))
        ax.set_xlabel(fake.word())
        ax.set_ylabel(fake.word())

    elif chart_type == 'pie':
        sizes = [random.randint(1, 10) for _ in range(4)]
        labels = [fake.word(), fake.word(), fake.word(), fake.word()]
        colors = ['red', 'green', 'blue', 'yellow']
        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=140)
        ax.set_title(fake.sentence(nb_words=random.randint(1, 4)))


def add_picture_with_caption(document, base_font_size):
    """ Функция добавления картинки вместе с подписью сразу """
    # Создание графика (будет вместо рисунков)
    generate_plot_or_chart()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='PNG')
    plt.close()
    img_stream.seek(0)

    # Определяем выравнивание
    alignment = random.choices(['left', 'center', 'right'], weights=[30, 70, 30])[0]
    alignment_value = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT
    }[alignment]

    # Создание нового параграфа для рисунка
    paragraph = document.add_paragraph()  # Новый параграф для рисунка
    run = paragraph.add_run()  # Добавляем пустой run для размещения изображения

    # Добавление рисунка
    run.add_picture(img_stream, width=Inches(4))

    # Устанавливаем выравнивание для рисунка
    paragraph.alignment = alignment_value

    # Определение позиции подписи (80% - снизу, 20% - сверху)
    position = random.choices(['below', 'above'], weights=[0.8, 0.2])[0]

    # Создание подписи
    caption_prefix = random.choice(["Рис.", "Рисунок", "Figure"])
    picture_number = random.randint(1, 100)
    caption_text = f"{caption_prefix} {picture_number}. {fake.sentence(nb_words=random.randint(3, 7))}"

    # Создаем подпись в новом параграфе
    caption_paragraph = document.add_paragraph(caption_text)  # Добавляем подпись
    caption_paragraph.alignment = alignment_value  # Устанавливаем выравнивание для подписи

    # Устанавливаем стиль для подписи
    caption_paragraph.style = 'Caption'
    run = caption_paragraph.runs[0]  # Получаем доступ к тексту подписи
    run.font.size = Pt(max(base_font_size - 2, 8))
    run.italic = True
    run.font.color.rgb = RGBColor(0, 0, 0)

    # Если подпись выше рисунка, нужно вставить перевод строки перед подписью
    if position == 'above':
        document.add_paragraph()  # Добавляем пустой параграф для отделения

    # Убедимся, что у нас есть отступ между рисунком и подписью
    document.add_paragraph()  # Добавляем пустой параграф для отделения


def generate_formula():
    """ Функция генерирует рандомную формулу в виде a {+|-|*|:} b = c"""
    # Генерация случайных чисел для выражений
    operand1 = random.randint(-10000, 10000)
    operand2 = random.randint(1, 10000)  # Избегаем деления на 0

    # Случайный выбор операции
    operation = random.choice(['+', '-', '*', '/'])

    # Формируем выражение в зависимости от операции
    if operation == '+':
        expr = f"{operand1} + {operand2}"
        result = operand1 + operand2
    elif operation == '-':
        expr = f"{operand1} - {operand2}"
        result = operand1 - operand2
    elif operation == '*':
        expr = f"{operand1} * {operand2}"
        result = operand1 * operand2
    else:
        expr = f"{operand1} / {operand2}"
        result = operand1 / operand2

    # Формируем итоговое выражение в формате "что-то (операция) с чем-то = что-то"
    return f"{expr} = {result}"


def get_formula(doc, latex_data):
    """ Вызвать функцию == добавить формулу в документ"""
    paragraph = doc.add_paragraph()
    latex_ = r'{}'.format(latex_data[random.randint(a=0, b=2371)])

    math2docx.add_math(paragraph, latex_)


def add_numbered_list(document):
    """ Вызвать функцию == добавить нумерованный список в документ"""
    list_type = 'List Number'
    num_items = random.randint(3, 7)
    for _ in range(num_items):
        list_item = fake.sentence(nb_words=random.randint(3, 10)) if random.choice([True, False]) else \
            text_mimesis.text(quantity=1).split('.')[0]
        paragraph = document.add_paragraph(list_item, style=list_type)
        paragraph.paragraph_format.left_indent = Inches(0.85)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)


def add_bulleted_list(document):
    """ Вызвать функцию == добавить маркированный список в документ"""
    list_type = 'List Bullet'
    num_items = random.randint(3, 7)
    for _ in range(num_items):
        list_item = fake.sentence(nb_words=random.randint(3, 10)) if random.choice([True, False]) else \
            text_mimesis.text(quantity=1).split('.')[0]
        paragraph = document.add_paragraph(list_item, style=list_type)
        paragraph.paragraph_format.left_indent = Inches(0.85)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)


def add_paragraph_with_footnote(document, footnotes, random_paragraph_format):
    text = fake.text(max_nb_chars=random.randint(50, 300))
    footnote_text = fake.sentence(nb_words=random.randint(5, 10))
    paragraph = document.add_paragraph()
    run = paragraph.add_run(f'{fake.text(max_nb_chars=random.randint(50, 300))} [^{len(footnotes) + 1}]')
    run.font.name = random_paragraph_format.font_name
    run.font.size = Pt(random_paragraph_format.font_size)
    paragraph_format = paragraph.paragraph_format
    paragraph_format.left_indent = random_paragraph_format.left_indent
    paragraph_format.space_before = Pt(random_paragraph_format.space_before)
    paragraph_format.space_after = Pt(random_paragraph_format.space_after)
    paragraph_format.first_line_indent = random_paragraph_format.first_line_indent
    paragraph.alignment = random_paragraph_format.alignment
    footnotes.append((len(footnotes) + 1, footnote_text))


def add_footnotes_section(document, footnotes, ):
    document.add_paragraph()
    for fn_count, fn_text in footnotes:
        fn_paragraph = document.add_paragraph()
        index_run = fn_paragraph.add_run(f'[^ {fn_count}] ')
        index_run.bold = True
        fn_paragraph.add_run(fn_text)
        fn_paragraph.style = document.styles['Normal']
        fn_paragraph.paragraph_format.left_indent = Pt(18)
        fn_paragraph.paragraph_format.space_after = Pt(2)
        fn_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_footnote(document, base_font_size):
    """ Вызвать функцию == добавить сноску в документ
    Сноска в данном понимании -- маленький текст в конце главы,
    с пониженным размером шрифта, курсивом и подчеркиванием.
    """
    random_word = fake.word()
    random_sentence = fake.sentence()

    # Создаем текст сноски в нужном формате (word* -- definition)
    footnote_text = f"{random_word}* — {random_sentence}"

    footnote_paragraph = document.add_paragraph()
    run = footnote_paragraph.add_run(footnote_text)

    run.font.size = Pt(max(base_font_size - 2, 8))  # Уменьшаем размер шрифта
    run.italic = True  # Устанавливаем курсив
    run.underline = True  # Подчеркиваем текст

    # Выравниваем сноску влево
    footnote_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def set_multicolumn(section, num_columns=2):
    """Установить режим многоколонного текста"""
    sectPr = section._sectPr
    # Создаем элемент для колонок
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), str(num_columns))
    cols.set(qn('w:space'), '720')
    cols.set(qn('w:equalWidth'), '1')

    # Удаляем существующие колонки, если есть
    # Делал, когда исправлял, не знаю, нужно ли сейчас
    existing_cols = sectPr.find(qn('w:cols'))
    if existing_cols is not None:
        sectPr.remove(existing_cols)
    sectPr.append(cols)


def reset_multicolumn(section):
    """ Сбросить с многоколонного режима до обычного """
    sectPr = section._sectPr
    # Удаляем элемент колонок, если существует
    existing_cols = sectPr.find(qn('w:cols'))
    if existing_cols is not None:
        sectPr.remove(existing_cols)
    # Добавляем обратно одну колонку по умолчанию
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '1')
    cols.set(qn('w:space'), '720')
    sectPr.append(cols)


def add_multi_column_text(document, base_font_size):
    """ Вызвать функцию == добавить многоколонный (2-3 колонны) текст"""
    # Добавляем секцию с многоколонным текстом
    multi_col_section = document.add_section(WD_SECTION.NEW_PAGE)
    num_columns = random.randint(2, 3)  # Случайное количество колонок от 2 до 3
    set_multicolumn(multi_col_section, num_columns=num_columns)

    # Рассчитывалось как вставка определенного количества текста в каждую колонну,
    # но оказалось, что текст вставляется просто в одну колонну, а потом избыток переносится.
    # TODO: Может быть оптимизировано.
    for _ in range(num_columns):
        column_text = fake.text(
            max_nb_chars=10000)  # Генерируем намного больше чем 1000 текст, чтобы после обрезать до 1000
        if num_columns == 2:
            column_text = column_text[:1000]  # Такая практика нужна была, чтобы во всех колонах было по 1000 символов.
        else:
            column_text = column_text[:633]  # Но оказалось, что это не работает, как должно.
        # Добавляем новый параграф в документ для текущей колонки
        paragraph = document.add_paragraph(column_text)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY  # Выравнивание по ширине

        # Устанавливаем размер шрифта
        for run in paragraph.runs:
            run.font.size = Pt(base_font_size)


def get_paragraph(doc, based_font_size):
    generic = Generic(locale=Locale.RU)
    random = Random()

    colums = random.weighted_choice({'2': 0.5, '3': 0.5})
    if colums == '2':
        table = doc.add_table(rows=1, cols=2)

        left_cell = table.cell(0, 0)
        right_cell = table.cell(0, 1)

        quantity = random.randint(a=30, b=60)
        left_cell.text = ' '.join(generic.text.words(quantity=quantity))
        left_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        left_cell.paragraphs[0].runs[0].font.size = Pt(based_font_size)
        right_cell.text = ' '.join(generic.text.words(quantity=quantity))
        right_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        right_cell.paragraphs[0].runs[0].font.size = Pt(based_font_size)

    elif colums == '3':
        table = doc.add_table(rows=1, cols=3)

        left_cell = table.cell(0, 0)
        center_cell = table.cell(0, 1)
        right_cell = table.cell(0, 2)

        quantity = random.randint(a=20, b=40)
        left_cell.text = ' '.join(generic.text.words(quantity=quantity))
        left_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        left_cell.paragraphs[0].runs[0].font.size = Pt(based_font_size)
        right_cell.text = ' '.join(generic.text.words(quantity=quantity))
        right_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        right_cell.paragraphs[0].runs[0].font.size = Pt(based_font_size)
        center_cell.text = ' '.join(generic.text.words(quantity=quantity))
        center_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        center_cell.paragraphs[0].runs[0].font.size = Pt(based_font_size)


def get_footnote(doc, based_font_size):
    generic = Generic(locale=Locale.RU)
    random = Random()

    section = doc.sections[0]
    paragraph = section.footer.paragraphs[0]
    paragraph.style.font.size = Pt(based_font_size)

    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    power_utf8 = {1: '\u00B9', 2: '\u00B2', 3: '\u00B3', 4: '\u2074'}
    for note in range(1, random.randint(a=2, b=5)):
        random_word = generic.text.words(quantity=1)
        random_sentence = generic.text.words(quantity=random.randint(a=3, b=6))

        footnote_text = f"{power_utf8[note]}{random_word[0]} — {' '.join(random_sentence)} \n"
        paragraph.add_run(footnote_text).italic = True


def get_image(doc, based_font_size, num_of_pictures):
    random = Random()
    generic = Generic(locale=Locale.RU)

    image_note = random.randint(a=0, b=1)
    if image_note == 0:
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(str('Рисунок ' + str(num_of_pictures) + ' — ' + ' '.join(
            generic.text.words(quantity=random.randint(a=1, b=3)))))
        run.font.size = Pt(based_font_size)
        num_of_pictures += 1
    if random.weighted_choice({0: 0.5, 1: 0.5}) > 0:
        doc.add_picture(
            'doc_generator/Caltech101/' + str(random.randint(a=1, b=9142)) + '.jpg',
            width=Inches(random.choice_enum_item((3.5, 4, 4.5))))
    else:
        doc.add_picture('doc_generator/PlotQA/' + str(random.randint(a=0, b=33656)) + '.png',
                             width=Inches(random.choice_enum_item((5, 5.5, 6))))
    if image_note == 1:
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(str('Рисунок ' + str(num_of_pictures) + ' — ' + ' '.join(
            generic.text.words(quantity=random.randint(a=1, b=3)))))
        run.font.size = Pt(based_font_size)
        num_of_pictures += 1


def generate_document(path):
    """
    Запуск генерации документа.
    Создает документ, выбирает базовый размер шрифта и сам шрифт.
    После этого выбирает количество итераций генерации элементов.
    Генерирует NUM_ITERATIONS раз все элементы, вызывая функции выше.
    """
    document = Document()
    table_styles = [style.name for style in document.styles if style.type == WD_STYLE_TYPE.TABLE]
    num_of_pictures = 1


    # Установка начального шрифта
    style = document.styles['Normal']
    latex_data = pd.read_csv(r'latex_for_formulas.csv', index_col=False)['formula']

    for i in range(NUM_ITERATIONS):

        random_paragraph_format = RandomParagraphFormat()
        random_table_format = RandomTableFormat(random_paragraph_format, table_styles)
        random_heading_format = RandomHeadingFormat(random_paragraph_format)
        footnotes_type = random.choice([0, 1, 2, 3])
        footnotes = []

        # Добавление новой секции на новой странице, кроме первой
        # А на первой добавляем колонтитулы -- раз и на все страницы.
        if i != 0:
            new_sect = document.add_section(WD_SECTION.NEW_PAGE)
            reset_multicolumn(new_sect)  # TODO [check actuality]: проверить надобность этой строки
        else:
            add_header_footer(document, random_paragraph_format.font_size)

        # Изменение ориентации страницы на вертикальную с вероятностью 85%
        if random.random() < 0.85:
            landscape = random.choice([True, False])
            change_orientation(document, landscape=landscape)

        # Добавление заголовка
        add_heading(document, random_heading_format)

        # Добавление абзаца
        add_paragraph(document, random_paragraph_format)

        # Добавление таблицы с подписью
        add_table_with_caption(document, random_table_format)

        # Добавление абзаца
        add_paragraph(document, random_paragraph_format)

        if footnotes_type == 2:
            for _ in range(random.randint(0, 5)):
                add_paragraph_with_footnote(document, footnotes, random_paragraph_format)

        # Добавление списков
        if random.random() < 0.3:
            add_numbered_list(document)
        elif random.random() < 0.7:
            add_bulleted_list(document)

        # Добавление рисунка с подписью
        add_picture_with_caption(document, random_paragraph_format.font_size)

        # Добавление формул
        get_formula(document, latex_data)

        # Добавление абзаца
        add_paragraph(document, random_paragraph_format)

        # Добавление сносок
        if random.random() < 0.2 and footnotes_type == 1:
            add_footnote(document, random_paragraph_format.font_size)
        elif random.random() < 0.3 and footnotes_type == 3:
            get_footnote(document, random_paragraph_format.font_size)
            
        if footnotes_type == 2 and len(footnotes) > 0:
            add_footnotes_section(document, footnotes)

        get_paragraph(document, random_paragraph_format.font_size)

    # Сохранение документа
    document.save(path)
