import streamlit as st
import openpyxl
import io
from collections import defaultdict

st.set_page_config(page_title="Excel Обработчик", page_icon="📊")

st.title("📊 Обработка Excel файлов")
st.markdown("---")

def add_statistics_table(ws):
    """Добавляет таблицу статистики"""
    try:
        # Находим строку с "общее количество"
        summary_row = None
        for row in range(1, ws.max_row + 1):
            if ws[f'A{row}'].value == "общее количество":
                summary_row = row
                break
        
        if not summary_row:
            return
        
        # Создаем таблицу через 1 строку
        table_start_row = summary_row + 2
        
        # Заголовки таблицы
        headers = [
            "номер депеши",
            "кол-во всего",
            "кол-во посылки", 
            "кол-во мешки",
            "вес всего",
            "вес посылки",
            "вес мешки"
        ]
        
        for col_idx, header in enumerate(headers, start=1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws[f'{col_letter}{table_start_row}'] = header
        
        # Собираем данные
        depesh_codes = set()
        data_by_depesh = defaultdict(lambda: {'pos_count': 0, 'mesh_count': 0, 'pos_weight': 0.0, 'mesh_weight': 0.0})
        
        # Анализируем данные Посылок (A-C)
        for row in range(2, ws.max_row + 1):
            code = ws[f'A{row}'].value
            depesh = ws[f'B{row}'].value
            weight = ws[f'C{row}'].value
            
            if code and depesh and weight is not None:
                if isinstance(depesh, str) and len(depesh) == 4 and depesh.isdigit():
                    depesh_codes.add(depesh)
                    data_by_depesh[depesh]['pos_count'] += 1
                    data_by_depesh[depesh]['pos_weight'] += float(weight)
        
        # Анализируем данные Мешков (D-F)
        for row in range(2, ws.max_row + 1):
            code = ws[f'D{row}'].value
            depesh = ws[f'E{row}'].value
            weight = ws[f'F{row}'].value
            
            if code and depesh and weight is not None:
                if isinstance(depesh, str) and len(depesh) == 4 and depesh.isdigit():
                    depesh_codes.add(depesh)
                    data_by_depesh[depesh]['mesh_count'] += 1
                    data_by_depesh[depesh]['mesh_weight'] += float(weight)
        
        # Заполняем таблицу
        sorted_depesh_codes = sorted(depesh_codes)
        for idx, depesh in enumerate(sorted_depesh_codes, start=1):
            table_row = table_start_row + idx
            data = data_by_depesh[depesh]
            
            ws[f'A{table_row}'] = depesh
            ws[f'B{table_row}'] = data['pos_count'] + data['mesh_count']
            ws[f'C{table_row}'] = data['pos_count']
            ws[f'D{table_row}'] = data['mesh_count']
            ws[f'E{table_row}'] = data['pos_weight'] + data['mesh_weight']
            ws[f'F{table_row}'] = data['pos_weight']
            ws[f'G{table_row}'] = data['mesh_weight']
            
    except Exception as e:
        st.error(f"Ошибка при добавлении статистики: {e}")

def process_excel_file(file_bytes):
    """Полная обработка Excel файла"""
    try:
        # Загружаем книгу
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        
        # ШАГ 1: Вставить два столбца между A и B
        ws.insert_cols(2, 2)
        
        # ШАГ 2: Заполнить столбец B (номер депеши из A)
        ws['B1'] = "номер депеши"
        for row in range(2, ws.max_row + 1):
            cell_a = ws[f'A{row}'].value
            if cell_a and len(str(cell_a)) == 29:
                ws[f'B{row}'] = str(cell_a)[16:20]
        
        # ШАГ 3: Заполнить столбец C (вес из A)
        ws['C1'] = "вес"
        for row in range(2, ws.max_row + 1):
            cell_a = ws[f'A{row}'].value
            if cell_a and len(str(cell_a)) == 29:
                try:
                    ws[f'C{row}'] = int(str(cell_a)[-4:]) / 10
                except:
                    pass
        
        # ШАГ 4: Заполнить столбец E (номер депеши из D)
        ws['E1'] = "номер депеши"
        for row in range(2, ws.max_row + 1):
            cell_d = ws[f'D{row}'].value
            if cell_d and len(str(cell_d)) == 29:
                ws[f'E{row}'] = str(cell_d)[16:20]
        
        # ШАГ 5: Заполнить столбец F (вес из D)
        ws['F1'] = "вес"
        for row in range(2, ws.max_row + 1):
            cell_d = ws[f'D{row}'].value
            if cell_d and len(str(cell_d)) == 29:
                try:
                    ws[f'F{row}'] = int(str(cell_d)[-4:]) / 10
                except:
                    pass
        
        # Найти последнюю строку с данными
        last_row = 1
        for row in range(2, ws.max_row + 1):
            cell_a = ws[f'A{row}'].value
            cell_d = ws[f'D{row}'].value
            if (cell_a and len(str(cell_a)) == 29) or (cell_d and len(str(cell_d)) == 29):
                last_row = row
        
        # ШАГ 6: Общий вес
        start_row = last_row + 4
        ws[f'A{start_row}'] = "общий вес"
        total_weight = 0
        for row in range(2, ws.max_row + 1):
            weight_c = ws[f'C{row}'].value
            weight_f = ws[f'F{row}'].value
            if isinstance(weight_c, (int, float)):
                total_weight += weight_c
            if isinstance(weight_f, (int, float)):
                total_weight += weight_f
        ws[f'B{start_row}'] = round(total_weight, 1)
        
        # ШАГ 7: Общее количество
        ws[f'A{start_row + 1}'] = "общее количество"
        total_count = 0
        for row in range(2, ws.max_row + 1):
            cell_a = ws[f'A{row}'].value
            cell_d = ws[f'D{row}'].value
            if cell_a and len(str(cell_a)) == 29:
                total_count += 1
            if cell_d and len(str(cell_d)) == 29:
                total_count += 1
        ws[f'B{start_row + 1}'] = total_count
        
        # ШАГ 8: Добавляем таблицу статистики
        add_statistics_table(ws)
        
        # Сохраняем в память
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output
        
    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
        return None

# Интерфейс пользователя
uploaded_file = st.file_uploader("Выберите Excel файл", type=['xlsx'])

if uploaded_file is not None:
    st.info(f"📁 Загружен: {uploaded_file.name} ({uploaded_file.size / 1024:.1f} KB)")
    
    if st.button("🚀 ОБРАБОТАТЬ ФАЙЛ", type="primary"):
        with st.spinner("⏳ Идет обработка..."):
            # Читаем файл
            file_bytes = uploaded_file.getvalue()
            
            # Обрабатываем
            processed_file = process_excel_file(file_bytes)
            
            if processed_file:
                st.success("✅ Файл успешно обработан!")
                
                # Кнопка скачивания
                st.download_button(
                    label="📥 СКАЧАТЬ ОБРАБОТАННЫЙ ФАЙЛ",
                    data=processed_file,
                    file_name=f"processed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.balloons()
else:
    st.info("👆 Загрузите файл .xlsx для начала обработки")

st.markdown("---")
st.caption("🔹 Программа добавляет номера депеш, веса, итоги и таблицу статистики")
