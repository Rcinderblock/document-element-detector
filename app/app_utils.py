import torch
import numpy as np
import cv2
import json
from ultralytics import YOLO


def load_model(model_path):
    try:
        model = YOLO(model_path)
        print("Модель успешно загружена")
    except Exception as e:
        print(f"Ошибка загрузки модели: {e}")
        model = None
    return model


def serialize_elements(elements):
    """Преобразует элементы в JSON-сериализуемый формат."""
    for element in elements:
        for key, value in element.items():
            if isinstance(value, np.ndarray):
                element[key] = value.tolist()
    return elements


def process_image(image, model):
    convert_to_612_792 = True  # СМЕНА РАСШИРЕНИЯ, ФЛАГ МОЖНО УБРАТЬ, КОГДА ПОФИКСИМ БАГИ
    if convert_to_612_792:
        width, height = image.size
        if width > height:  # Горизонтальное изображение
            image = image.resize((792, 612))
        else:  # Вертикальное изображение
            image = image.resize((612, 792))

    image_np = np.array(image)  # Преобразуем PIL Image в numpy-формат

    image_np = cv2.cvtColor(image_np, cv2.COLOR_RGB2BGR)
    results = model(image_np)  # Передаем изображение напрямую в модель

    # Рисуем боксы и метки на изображении
    annotated_image = results[0].plot()  # Метод plot() рисует боксы и метки

    # Сериализация элементов (координат и меток объектов)
    elements = []
    for result in results:
        boxes = result.boxes.xyxy.cpu().numpy()  # Координаты боксов
        confidences = result.boxes.conf.cpu().numpy()  # Уверенность
        classes = result.boxes.cls.cpu().numpy()  # Классы

        for box, confidence, cls in zip(boxes, confidences, classes):
            x1, y1, x2, y2 = map(int, box)
            elements.append({
                "x1": x1,
                "y1": y1,
                "x2": x2,
                "y2": y2,
                "confidence": float(confidence),
                "class": result.names[int(cls)]
            })

    return {
        "elements": serialize_elements(elements),
        "annotated_image": annotated_image
    }
