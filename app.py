import streamlit as st
import openpyxl
import io

st.set_page_config(page_title="Excel Обработчик", page_icon="📊")

st.title("📊 Обработка Excel файлов")
st.write("Загрузите файл .xlsx для обработки")

uploaded_file = st.file_uploader("Выберите файл", type=['xlsx'])

if uploaded_file is not None:
    st.success(f"✅ Файл загружен: {uploaded_file.name}")
    st.balloons()
