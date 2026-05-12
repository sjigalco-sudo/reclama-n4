import streamlit as st
import pandas as pd
from datetime import timedelta, datetime, time
import io
import os

st.set_page_config(page_title="Global24 Multi-TXT", page_icon="📝")
st.title("📝 Генератор ID и Длительности (в минутах)")

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
        df_res = df.iloc[:, [2, 6, 7, 9]].copy()
        df_res.columns = ['Block_Time', 'Name', 'Dur', 'ID']
        df_res['Block_Time'] = df_res['Block_Time'].ffill()
        df_res = df_res.dropna(subset=['Name', 'ID'])
        
        grouped = df_res.groupby('Block_Time', sort=False)

        out_id = io.StringIO()
        out_min = io.StringIO()

        for i, (block_time, items) in enumerate(grouped, 1):
            time_str = format_time(block_time)
            
            # ФАЙЛ 1: Только ID (без заставок, без пробелов)
            out_id.write(f"{time_str}\n")
            id_elements = [f'"{str(row["ID"]).split(".")[0]}"' for _, row in items.iterrows()]
            out_id.write("".join(id_elements) + "\n\n")

            # ФАЙЛ 2: (Сумма роликов + 20 сек) / 60 = МИНУТЫ
            out_min.write(f"{time_str}\n")
            
            sum_dur_only_ads = items['Dur'].sum()
            total_seconds = sum_dur_only_ads + 20.0
            total_minutes = total_seconds / 60.0
            
            # Записываем результат (минуты с дробной частью, например 1.50)
            out_min.write(f"{total_minutes:.2f}\n\n")

        # Интерфейс
        st.subheader("Скачать файлы:")
        c1, c2 = st.columns(2)
        
        with c1:
            st.download_button(
                label="📥 Скачать ID",
                data=out_id.getvalue(),
                file_name=f"IDs_{base_name}.txt",
                mime="text/plain"
            )
            
        with c2:
            st.download_button(
                label="📥 Скачать Минуты (+20с)",
                data=out_min.getvalue(),
                file_name=f"Minutes_{base_name}.txt",
                mime="text/plain"
            )

        # Превью для проверки
        st.divider()
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.write("**Превью ID:**")
            st.text(out_id.getvalue()[:200])
        with col_p2:
            st.write("**Превью минут:**")
            st.text(out_min.getvalue()[:200])

    except Exception as e:
        st.error(f"Ошибка: {e}")
