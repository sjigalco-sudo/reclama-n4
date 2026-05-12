import streamlit as st
import pandas as pd
from datetime import timedelta, datetime, time
import io
import os

st.set_page_config(page_title="N4 Multi-Format Generator", page_icon="📺")
st.title("📺 Генератор для N4 (Все форматы)")

def format_time_str(x):
    if isinstance(x, (datetime, time)):
        return x.strftime('%H:%M:%S')
    if isinstance(x, (int, float)):
        total_seconds = int(round(x * 86400))
        return str(timedelta(seconds=total_seconds)).zfill(8)
    return str(x)

def seconds_to_hms(total_seconds):
    return str(timedelta(seconds=int(round(total_seconds)))).zfill(8)

uploaded_file = st.file_uploader("Загрузите Excel файл (N4)", type=["xls", "xlsx"])

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

        # Списки для сбора данных
        txt_id_content = io.StringIO()
        xlsx_id_data = []
        xlsx_timing_data = []

        for i, (block_time, items) in enumerate(grouped, 1):
            time_label = format_time_str(block_time)
            
            # Собираем ID в строку вида "ID1""ID2"
            id_list_str = "".join([f'"{str(row["ID"]).split(".")[0]}"' for _, row in items.iterrows()])
            
            # 1. Данные для TXT и XLSX вариантов списка ID
            txt_id_content.write(f"{time_label}\n{id_list_str}\n\n")
            xlsx_id_data.append({
                "Время выхода": time_label,
                "Список ID": id_list_str
            })

            # 2. Данные для XLSX таймингов (Сумма + 20 сек)
            total_block_seconds = items['Dur'].sum() + 20.0
            xlsx_timing_data.append({
                "Время выхода блока": time_label,
                "Длительность": seconds_to_hms(total_block_seconds)
            })

        # Создаем Excel-буферы
        def to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        # Интерфейс
        st.subheader("Скачать файлы:")
        
        # Секция ID
        st.write("### 📂 Списки ID")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 IDs (.txt)", txt_id_content.getvalue(), f"N4_IDs_{base_name}.txt")
        with c2:
            st.download_button("📥 IDs (.xlsx)", to_excel(pd.DataFrame(xlsx_id_data)), f"N4_IDs_Table_{base_name}.xlsx")
            
        # Секция Таймингов
        st.write("### ⏳ Тайминги (Длительность + 20с)")
        st.download_button("📥 Timings (.xlsx)", to_excel(pd.DataFrame(xlsx_timing_data)), f"N4_Timings_{base_name}.xlsx")

        # Предпросмотр нового формата
        st.divider()
        st.write("**Предпросмотр таблицы ID (Новый формат):**")
        st.table(pd.DataFrame(xlsx_id_data).head(5))

    except Exception as e:
        st.error(f"Ошибка: {e}")
