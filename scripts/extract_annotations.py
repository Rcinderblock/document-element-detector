import os
import re
import json
import fitz  # PyMuPDF
from PIL import Image, ImageDraw


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
            'formula': (139, 69, 19)  # Saddle Brown
        }

    def extract_coordinates(self, pdf_path):
        """Extract coordinates of different elements from PDF."""
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

            # Extract text blocks and their coordinates
            blocks = page.get_text("dict")["blocks"]
            for block in blocks:
                bbox = tuple(block['bbox'])
                text = self.extract_text_from_block(block)
                if not text:
                    print('WARNING: text is empty.')
                    continue
                # Detect element type based on formatting and content
                if self._is_title(block):
                    page_dict['title'].append(self._convert_coordinates(bbox))
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
                elif self._is_header(bbox, page.rect.height):
                    page_dict['header'].append(self._convert_coordinates(bbox))
                elif self._is_formula(text):
                    page_dict['formula'].append(self._convert_coordinates(bbox))
                else:
                    page_dict['paragraph'].append(self._convert_coordinates(bbox))

            # Extract images
            images = page.get_images(full=True)
            for img in images:
                xref = img[0]
                pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
                if pix:
                    page_dict['picture'].append(self._convert_coordinates(pix.irect))

            page_data.append(page_dict)

        doc.close()
        return page_data


    def _convert_coordinates(self, bbox):
        """Convert coordinates to [x1, y1, x2, y2] format."""
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
        """Check if text is part of a bulleted list."""
        return bool(re.match(r'^(\s*[-•*]\s+)', text))


    def _is_footer(self, bbox, page_height):
        return bbox[1] > page_height - 50

    def _is_header(self, bbox, page_height):
        return bbox[1] < 50

    def _is_formula(self, text):
        return '=' in text and any(c.isdigit() for c in text)

    def generate_annotated_images(self, pdf_dir, output_dir):
        """Generate annotated images with bounding boxes for each PDF in pdf_dir."""
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

                # Draw bounding boxes for each element type
                for element_type, coordinates_list in page_info.items():
                    if element_type not in ['image_height', 'image_width', 'image_path']:
                        color = self.colors.get(element_type, (0, 0, 0))
                        for coords in coordinates_list:
                            draw.rectangle(coords, outline=color, width=2)

                # Save annotated image
                output_path = os.path.join(output_dir,
                                           f'{os.path.splitext(pdf_file)[0]}_annotated_page_{page_num + 1}.png')
                img.save(output_path)

                # Save JSON data
                json_path = os.path.join(output_dir, f'{os.path.splitext(pdf_file)[0]}_page_{page_num + 1}.json')
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(page_info, f, indent=2)

            doc.close()


# Папки с данными
PDF_DIR = 'data/pdf/'
ANNOTATIONS_DIR = 'data/annotations/'

# Создание директории для аннотаций, если она не существует
os.makedirs(ANNOTATIONS_DIR, exist_ok=True)


def main():
    analyzer = DocumentAnalyzer()
    analyzer.generate_annotated_images(PDF_DIR, ANNOTATIONS_DIR)


if __name__ == "__main__":
    main()
