import streamlit as st
import pandas as pd
from datetime import timedelta, datetime, time
import io
import os

st.set_page_config(page_title="Global24 TXT Generator", page_icon="📝")
st.title("📝 Генератор ID и Таймингов (Мин.Сек)")

def format_time(x):
    if isinstance(x, (datetime, time)):
        return x.strftime('%H:%M:%S')
    if isinstance(x, (int, float)):
        total_seconds = int(round(x * 86400))
        return str(timedelta(seconds=total_seconds)).zfill(8)
    return str(x)

uploaded_file = st.file_uploader("Загрузите Excel", type=["xls", "xlsx"])

if uploaded_file:
    try:
        base_name = os.path.splitext(uploaded_file.name)[0]
        
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
            
            # ФАЙЛ 1: Время + ID слитно
            out_id.write(f"{time_str}\n")
            id_elements = [f'"{str(row["ID"]).split(".")[0]}"' for _, row in items.iterrows()]
            out_id.write("".join(id_elements) + "\n\n")

            # ФАЙЛ 2: Время + Формат МИНУТЫ.СЕКУНДЫ
            out_time_format.write(f"{time_str}\n")
            
            # Считаем сумму секунд + ваши 20 секунд
            total_seconds_raw = int(round(items['Dur'].sum() + 20.0))
            
            # Переводим в минуты и остаток секунд
            minutes = total_seconds_raw // 60
            seconds = total_seconds_raw % 60
            
            # Записываем как 1.23 (где 23 — это реальные секунды)
            out_time_format.write(f"{minutes}.{seconds:02d}\n\n")

        # Кнопки
        st.subheader("Скачать результаты")
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Скачать ID",
                data=out_id.getvalue(),
                file_name=f"IDs_{base_name}.txt"
            )
            
        with col2:
            st.download_button(
                label="📥 Скачать Время (М.СС)",
                data=out_time_format.getvalue(),
                file_name=f"Time_MS_{base_name}.txt"
            )

        # Превью
        st.divider()
        st.write("**Пример второго файла (Минуты.Секунды):**")
        st.text(out_time_format.getvalue()[:300])

    except Exception as e:
        st.error(f"Ошибка: {e}")
