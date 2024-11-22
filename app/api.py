import os

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from .app_utils import load_model, process_image
from PIL import Image

app = FastAPI()

models = {}


@app.post("/process/")
async def process(
        file: UploadFile = File(...),
        model_name: str = Form(...),
):
    image = Image.open(file.file)

    model_path = os.path.join(os.getcwd(), "models", model_name)

    if not os.path.exists(model_path):
        raise HTTPException(status_code=404, detail=f"Модель не найдена по пути {model_path}")

    try:
        model = load_model(model_path)
        if model is None:
            raise HTTPException(status_code=500, detail="Не удалось загрузить модель")
    except Exception as e:
        print(f"Ошибка при загрузке модели: {e}")  # Логируем ошибку загрузки
        raise HTTPException(status_code=500, detail="Ошибка загрузки модели")

    results = process_image(image, model)

    return {
        "elements": results["elements"],
        "annotated_image": results["annotated_image"].tolist()
    }