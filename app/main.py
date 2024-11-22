import json

import streamlit as st
import requests
from PIL import Image
import os
import numpy as np

models_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../models")

model_files = [f for f in os.listdir(models_dir) if f.endswith('.pt')]

API_URL = "http://localhost:8000/process/"

st.title("Компьютерное зрение: определение элементов документа")
selected_model = st.selectbox("Выберите модель", model_files)

uploaded_file = st.file_uploader("Загрузите изображение", type=["png", "jpg", "jpeg"])

if uploaded_file and selected_model:
    files = {"file": uploaded_file}
    data = {"model_name": selected_model}


    # Отправляем запрос к API для обработки изображения
    with st.spinner(f"Обработка изображения..."):
        response = requests.post(API_URL, files=files, data=data)
        if response.status_code == 200:
            results = response.json()
            st.success("Обработка завершена!")

            # Отображение JSON с результатами
            with st.expander("Показать результаты в формате JSON"):
                st.json(results["elements"])


            annotated_image_array = np.array(results["annotated_image"])

            annotated_image_array = annotated_image_array.astype(np.uint8)

            # Создаем изображение
            annotated_image = Image.fromarray(annotated_image_array)
            st.image(annotated_image, caption="Размеченное изображение", use_container_width=True)
            json_filename = "latest_output.json"
            json_data = json.dumps(results["elements"], indent=4)

            # Кнопка для скачивания JSON
            st.download_button(
                label="Скачать результаты в формате JSON",
                data=json_data,
                file_name=json_filename,
                mime="application/json"
            )
        else:
            st.error("Ошибка обработки изображения. Проверьте модель и файл.")
