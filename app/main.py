import streamlit as st
from PIL import Image
import json
import os
from app_utils import load_model, process_image


# Определение корневой директории проекта
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Загрузка модели
models_dir = os.path.join(BASE_DIR, "../models/")
model_files = [f for f in os.listdir(models_dir) if f.endswith('.pt')]

# Выбор модели через Streamlit
selected_model_file = st.selectbox("Выберите модель", model_files)

# Загрузка выбранной модели
model_path = os.path.join(models_dir, selected_model_file)
model = load_model(model_path)

st.title("Компьютерное зрение: определение элементов документа")

# Загрузка изображения
uploaded_file = st.file_uploader("Загрузите изображение", type=["png"])

if uploaded_file:
    # Открываем изображение
    with st.expander("Показать загруженное изображение"):
        image = Image.open(uploaded_file)
        st.image(image, caption="Загруженное изображение", use_container_width=True)

    # Обработка изображения
    with st.spinner("Обработка изображения..."):
        results = process_image(image, model)
        if results is None:
            st.error("Ошибка обработки изображения. Пожалуйста, попробуйте еще раз.")
            st.stop()
        st.success("Обработка завершена!")  # Отображаем сообщение внутри блока spinner

    # Сохранение JSON в директорию ./annotations
    json_file_name = os.path.join(BASE_DIR, "latest_output.json")
    with open(json_file_name, "w") as f:
        json.dump(results["elements"], f, indent=4)

    # Отображение JSON
    with st.expander("Показать результаты в формате JSON"):
        st.json(results["elements"])

    # Отображение размеченного изображения
    annotated_image = Image.fromarray(results["annotated_image"])
    st.image(annotated_image, caption="Размеченное изображение", use_container_width=True)

    # Кнопка для скачивания JSON
    with open(json_file_name, "r") as f:
        st.download_button(
            label="Скачать результат в формате JSON",
            data=f,
            file_name="latest_output.json",
            mime="application/json"
        )