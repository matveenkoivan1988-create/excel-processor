import streamlit as st
import openpyxl
import io
from collections import defaultdict

st.set_page_config(page_title="Excel Обработчик", page_icon="📊")

def process_excel_file(file_stream):
    """Обрабатывает Excel файл"""
    wb = openpyxl.load_workbook(file_stream)
    ws = wb.active
    
    # Шаг 1: Вставить два столбца между A и B
    ws.insert_cols(2, 2)
    
    # Шаг 2: Заполнить столбец B
    ws['B1'] = "номер депеши"
    for row in range(2, ws.max_row + 1):
        cell_a = ws[f'A{row}'].value
        if cell_a and len(str(cell_a)) == 29:
            code = str(cell_a)
            ws[f'B{row}'] = code[16:20]
    
    # Шаг 3: Заполнить столбец C
    ws['C1'] = "вес"
    for row in range(2, ws.max_row + 1):
        cell_a = ws[f'A{row}'].value
        if cell_a and len(str(cell_a)) == 29:
            code = str(cell_a)
            try:
                weight = int(code[-4:]) / 10
                ws[f'C{row}'] = weight
            except:
                pass
    
    # Шаг 4: Заполнить столбец E
    ws['E1'] = "номер депеши"
    for row in range(2, ws.max_row + 1):
        cell_d = ws[f'D{row}'].value
        if cell_d and len(str(cell_d)) == 29:
            code = str(cell_d)
            ws[f'E{row}'] = code[16:20]
    
    # Шаг 5: Заполнить столбец F
    ws['F1'] = "вес"
    for row in range(2, ws.max_row + 1):
        cell_d = ws[f'D{row}'].value
        if cell_d and len(str(cell_d)) == 29:
            code = str(cell_d)
            try:
                weight = int(code[-4:]) / 10
                ws[f'F{row}'] = weight
            except:
                pass
    
    # Найти последнюю строку с кодом
    last_row = 1
    for row in range(2, ws.max_row + 1):
        cell_a = ws[f'A{row}'].value
        cell_d = ws[f'D{row}'].value
        if (cell_a and len(str(cell_a)) == 29) or (cell_d and len(str(cell_d)) == 29):
            last_row = row
    
    # Шаг 6: Общий вес
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
    ws[f'B{start_row}'] = total_weight
    
    # Шаг 7: Общее количество
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
    
    # Добавляем статистику
    add_statistics_table(ws)
    
    return wb

def add_statistics_table(ws):
    """Добавляет таблицу статистики"""
    summary_row = None
    for row in range(1, ws.max_row + 1):
        if ws[f'A{row}'].value == "общее количество":
            summary_row = row
            break
    
    if not summary_row:
        return
    
    table_start_row = summary_row + 2
    
    headers = [
        "номер депеши", "кол-во всего", "кол-во посылки", 
        "кол-во мешки", "вес всего", "вес посылки", "вес мешки"
    ]
    
    for col_idx, header in enumerate(headers, start=1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws[f'{col_letter}{table_start_row}'] = header
    
    depesh_codes = set()
    data_by_depesh = defaultdict(lambda: {'pos_count': 0, 'mesh_count': 0, 'pos_weight': 0.0, 'mesh_weight': 0.0})
    
    for row in range(2, ws.max_row + 1):
        depesh = ws[f'B{row}'].value
        weight = ws[f'C{row}'].value
        if depesh and weight is not None:
            if isinstance(depesh, str) and len(depesh) == 4 and depesh.isdigit():
                depesh_codes.add(depesh)
                data_by_depesh[depesh]['pos_count'] += 1
                data_by_depesh[depesh]['pos_weight'] += float(weight)
    
    for row in range(2, ws.max_row + 1):
        depesh = ws[f'E{row}'].value
        weight = ws[f'F{row}'].value
        if depesh and weight is not None:
            if isinstance(depesh, str) and len(depesh) == 4 and depesh.isdigit():
                depesh_codes.add(depesh)
                data_by_depesh[depesh]['mesh_count'] += 1
                data_by_depesh[depesh]['mesh_weight'] += float(weight)
    
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

def main():
    st.title("📊 Обработка Excel файлов")
    st.markdown("---")
    
    st.markdown("""
    ### Инструкция:
    1. Загрузите Excel файл (.xlsx)
    2. Нажмите кнопку "Обработать"
    3. Скачайте готовый файл
    """)
    
    uploaded_file = st.file_uploader("Выберите файл", type=['xlsx'])
    
    if uploaded_file is not None:
        if st.button("🚀 Обработать файл", type="primary"):
            with st.spinner("⏳ Обработка файла..."):
                try:
                    file_stream = io.BytesIO(uploaded_file.read())
                    processed_wb = process_excel_file(file_stream)
                    
                    output = io.BytesIO()
                    processed_wb.save(output)
                    output.seek(0)
                    
                    st.success("✅ Файл успешно обработан!")
                    
                    st.download_button(
                        label="📥 Скачать обработанный файл",
                        data=output,
                        file_name=f"processed_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"❌ Ошибка: {str(e)}")
    
    st.markdown("---")
    st.markdown("🔹 Поддерживаются только файлы .xlsx")

if __name__ == "__main__":
    main()