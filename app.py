import streamlit as st
import pandas as pd
from datetime import timedelta, datetime, time
import io
import os

st.set_page_config(page_title="N4 Playlist Generator", page_icon="📺")
st.title("📺 Генератор для канала N4")

def format_time(x):
    if isinstance(x, (datetime, time)):
        return x.strftime('%H:%M:%S')
    if isinstance(x, (int, float)):
        total_seconds = int(round(x * 86400))
        return str(timedelta(seconds=total_seconds)).zfill(8)
    return str(x)

uploaded_file = st.file_uploader("Загрузите Excel файл (N4)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        base_name = os.path.splitext(uploaded_file.name)[0]
        
        # Читаем Excel, пропускаем шапку
        df = pd.read_excel(uploaded_file, skiprows=6)
        
        # Колонки: Время (2), Длительность (7), ID (9)
        df_res = df.iloc[:, [2, 7, 9]].copy()
        df_res.columns = ['Block_Time', 'Dur', 'ID']
        
        df_res['Block_Time'] = df_res['Block_Time'].ffill()
        df_res = df_res.dropna(subset=['ID'])
        
        grouped = df_res.groupby('Block_Time', sort=False)

        out_id = io.StringIO()
        out_time_format = io.StringIO()

        for i, (block_time, items) in enumerate(grouped, 1):
            time_str = format_time(block_time)
            
            # ФАЙЛ 1: Время + ID в кавычках (слитно)
            out_id.write(f"{time_str}\n")
            id_elements = [f'"{str(row["ID"]).split(".")[0]}"' for _, row in items.iterrows()]
            out_id.write("".join(id_elements) + "\n\n")

            # ФАЙЛ 2: Время + Формат МИНУТЫ.СЕКУНДЫ (Сумма + 20 сек)
            out_time_format.write(f"{time_str}\n")
            
            # Считаем чистую сумму из Excel и добавляем 20 секунд
            total_seconds_raw = int(round(items['Dur'].sum() + 20.0))
            
            # Переводим в понятные минуты и секунды
            minutes = total_seconds_raw // 60
            seconds = total_seconds_raw % 60
            
            # Записываем (например, 1.23)
            out_time_format.write(f"{minutes}.{seconds:02d}\n\n")

        # Интерфейс скачивания
        st.subheader("Готовые файлы для N4:")
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Скачать IDs (N4)",
                data=out_id.getvalue(),
                file_name=f"N4_IDs_{base_name}.txt",
                mime="text/plain"
            )
            
        with col2:
            st.download_button(
                label="📥 Скачать Время M.SS (N4)",
                data=out_time_format.getvalue(),
                file_name=f"N4_Time_{base_name}.txt",
                mime="text/plain"
            )

        # Предпросмотр для контроля
        st.divider()
        st.write(f"**Предпросмотр таймингов для {base_name}:**")
        st.text(out_time_format.getvalue()[:300])

    except Exception as e:
        st.error(f"Ошибка: {e}")
