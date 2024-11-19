import streamlit as st
from PIL import Image
import json
import os
from app_utils import load_model, process_image

# Загрузка модели
model = load_model()

st.title("Компьютерное зрение: определение элементов документа")

# Загрузка изображения
uploaded_file = st.file_uploader("Загрузите изображение", type=["jpg", "png", "jpeg"])

if uploaded_file:
    # Открываем изображение
    image = Image.open(uploaded_file)
    st.image(image, caption="Загруженное изображение", use_container_width=True)

    # Обработка изображения
    with st.spinner("Обработка изображения..."):
        results = process_image(image, model)

    st.success("Обработка завершена!")

    # Создаем уникальное имя для файла JSON
    base_name = os.path.splitext(uploaded_file.name)[0]
    json_file_name = f"{base_name}_output.json"

    # Сохранение JSON
    with open(json_file_name, "w") as f:
        json.dump(results["elements"], f, indent=4)

    # Отображение JSON
    st.json(results["elements"])

    # Отображение размеченного изображения
    annotated_image = Image.fromarray(results["annotated_image"])
    st.image(annotated_image, caption="Размеченное изображение", use_container_width=True)

    # Кнопка для скачивания JSON
    with open(json_file_name, "r") as f:
        st.download_button(
            label="Скачать результат в формате JSON",
            data=f,
            file_name=json_file_name,
            mime="application/json"
        )
