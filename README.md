# Document Element Detector

## Обзор

**Document Element Detector** — это инструмент на Python, предназначенный для генерации документов и извлечения из них определенных видов элементов.

## Установка

Чтобы установить Document Element Detector, выполните следующие шаги:

1. Клонируйте репозиторий:
   ```bash
   git clone https://github.com/Rcinderblock/document-element-detector.git
   ```
2. Перейдите в каталог проекта:
   ```bash
   cd document-element-detector
   ```
3. Установите необходимые зависимости:
  ```bash
  pip install -r requirements.txt
  ```
## Использования
1. Для генерации документов следует выполнить скрипт:
  ```bash
  python scripts/generate_documents.py
  ```
В этом случае сгенерируется n документов. Константа n изменяется внутри кода в срипте:
  ```bash
  NUM_DOCUMENTS = 1000
  ```
2. Для перевода документов в PDF используйте скрипт:
   ```bash
   python scripst/convert_to_pdf.py
   ```
3. Для перевода из PDF в картинки (для будущего обучения модели) используйте:
   ```bash
   python scripts/pdf_to_images.py
   ```
4. Для извлечения элементов используйте скрипт:
   ```bash
   python scripts/extract_annotations.py
   ```
