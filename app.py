import streamlit as st
import pandas as pd
from datetime import timedelta, datetime, time
import io
import os

st.set_page_config(page_title="N4 Playlist Generator", page_icon="📺")
st.title("📺 Генератор для N4 (TXT + Excel)")

def format_time_str(x):
    if isinstance(x, (datetime, time)):
        return x.strftime('%H:%M:%S')
    if isinstance(x, (int, float)):
        total_seconds = int(round(x * 86400))
        return str(timedelta(seconds=total_seconds)).zfill(8)
    return str(x)

def seconds_to_hms(total_seconds):
    # Преобразуем секунды в формат 00:01:23
    return str(timedelta(seconds=int(round(total_seconds)))).zfill(8)

uploaded_file = st.file_uploader("Загрузите Excel файл (N4)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        base_name = os.path.splitext(uploaded_file.name)[0]
        
        # Читаем исходный файл
        df = pd.read_excel(uploaded_file, skiprows=6)
        
        # Столбцы: Время (2), Длительность (7), ID (9)
        df_res = df.iloc[:, [2, 7, 9]].copy()
        df_res.columns = ['Block_Time', 'Dur', 'ID']
        
        df_res['Block_Time'] = df_res['Block_Time'].ffill()
        df_res = df_res.dropna(subset=['ID'])
        
        grouped = df_res.groupby('Block_Time', sort=False)

        # Подготовка данных для TXT
        out_id = io.StringIO()
        
        # Подготовка данных для Excel
        xls_data = []

        for i, (block_time, items) in enumerate(grouped, 1):
            time_label = format_time_str(block_time)
            
            # 1. Формируем текстовый файл (ID)
            out_id.write(f"{time_label}\n")
            id_elements = [f'"{str(row["ID"]).split(".")[0]}"' for _, row in items.iterrows()]
            out_id.write("".join(id_elements) + "\n\n")

            # 2. Формируем данные для таблицы (Сумма + 20 сек)
            total_block_seconds = items['Dur'].sum() + 20.0
            hms_duration = seconds_to_hms(total_block_seconds)
            
            xls_data.append({
                "Время выхода блока": time_label,
                "Длительность": hms_duration
            })

        # Создаем DataFrame для Excel
        df_export = pd.DataFrame(xls_data)
        
        # Записываем Excel в буфер
        buffer_xlsx = io.BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Playlist_N4')
        
        # Интерфейс
        st.subheader("Результаты генерации:")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.info("Текстовый файл")
            st.download_button(
                label="📥 Скачать IDs (.txt)",
                data=out_id.getvalue(),
                file_name=f"N4_IDs_{base_name}.txt",
                mime="text/plain"
            )
            
        with col2:
            st.success("Таблица Excel")
            st.download_button(
                label="📥 Скачать Тайминги (.xlsx)",
                data=buffer_xlsx.getvalue(),
                file_name=f"N4_Timings_{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.divider()
        st.write("**Предпросмотр данных для Excel:**")
        st.table(df_export.head(10))

    except Exception as e:
        st.error(f"Произошла ошибка: {e}")
