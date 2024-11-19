import torch
import numpy as np
import cv2
import json
from ultralytics import YOLO


def load_model():
    # Загрузка модели YOLO
    model = YOLO('../models/best.pt')
    return model


def serialize_elements(elements):
    """Преобразует элементы в JSON-сериализуемый формат."""
    for element in elements:
        for key, value in element.items():
            if isinstance(value, np.ndarray):
                element[key] = value.tolist()
    return elements


def process_image(image, model):
    image_np = np.array(image)
    image_cv = cv2.cvtColor(image_np, cv2.COLOR_RGB2BGR)
    results = model(image_cv)  # Получаем результаты от модели

    # Рисуем боксы и метки на изображении
    annotated_image = results[0].plot()  # Используем метод plot() для отрисовки боксов

    # Сериализация элементов (координат и меток объектов)
    elements = []
    for result in results:
        boxes = result.boxes.xyxy.cpu().numpy()
        confidences = result.boxes.conf.cpu().numpy()
        classes = result.boxes.cls.cpu().numpy()

        for box, confidence, cls in zip(boxes, confidences, classes):
            x1, y1, x2, y2 = map(int, box)
            elements.append({
                "x1": x1,
                "y1": y1,
                "x2": x2,
                "y2": y2,
                "confidence": float(confidence),
                "class": int(cls)
            })

    # Преобразуем изображение обратно в RGB для отображения в Streamlit
    annotated_image_rgb = cv2.cvtColor(annotated_image, cv2.COLOR_BGR2RGB)

    return {
        "elements": serialize_elements(elements),
        "annotated_image": annotated_image_rgb
    }


def process_image_debug(image, model):
    image_np = np.array(image)
    image_cv = cv2.cvtColor(image_np, cv2.COLOR_RGB2BGR)
    results = model(image_cv)  # Получаем "сырые" результаты

    # Добавляем debug-вывод
    print("Raw results:", results)  # В терминал для проверки
    return results
