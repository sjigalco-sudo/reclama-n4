import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import timedelta, datetime, time

# Конфигурация страницы
st.set_page_config(page_title="N4 | Ultimate Generator", page_icon="📺", layout="wide")

# Кастомный стиль N4
st.markdown('''
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #E11D48; color: white; border: none; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; }
    h1 { color: #E11D48; border-bottom: 2px solid #E11D48; padding-bottom: 10px; }
    </style>
    ''', unsafe_allow_html=True)

st.title("📺 N4: ПОЛНЫЙ ГЕНЕРАТОР (ЭФИР + ID + ТАЙМИНГИ)")

# --- Вспомогательные функции ---
def format_time_hh_mm(x):
    if isinstance(x, (datetime, time)): return x.strftime('%H-%M')
    if isinstance(x, (int, float)):
        total_seconds = int(round(x * 86400))
        return str(timedelta(seconds=total_seconds)).zfill(8)[:5].replace(':', '-')
    return str(x)[:5].replace(':', '-')

def format_time_full(x):
    if isinstance(x, (datetime, time)): return x.strftime('%H:%M:%S')
    if isinstance(x, (int, float)):
        total_seconds = int(round(x * 86400))
        return str(timedelta(seconds=total_seconds)).zfill(8)
    return str(x)

def seconds_to_hms(total_seconds):
    return str(timedelta(seconds=int(round(total_seconds)))).zfill(8)

def xml_escape(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- Sidebar ---
with st.sidebar:
    st.header("⚙️ Настройки N4")
    user_path = st.text_input("Путь к роликам на сервере:", value=r"I:\RECLAMA 2026")
    st.divider()
    uploaded_file = st.file_uploader("Загрузите медиа-план (Excel)", type=["xls", "xlsx"])
    
    if user_path.endswith('\\'):
        user_path = user_path[:-1]

# --- Основная логика ---
if uploaded_file:
    try:
        # Извлекаем имя файла (дату) для именования выходных данных
        base_name = os.path.splitext(uploaded_file.name)[0]
        
        df = pd.read_excel(uploaded_file, skiprows=6)
        
        # Колонки: Время(2), Название(6), Длит(7), ID(9)
        df_res = df.iloc[:, [2, 6, 7, 9]].copy()
        df_res.columns = ['Block_Time', 'Name', 'Dur', 'ID']
        df_res['Block_Time'] = df_res['Block_Time'].ffill()
        df_res = df_res.dropna(subset=['ID'])
        
        grouped = df_res.groupby('Block_Time', sort=False)
        
        zip_slblock_buffer = io.BytesIO()
        txt_id_content = io.StringIO()
        xlsx_id_data = []
        xlsx_timing_data = []
        
        with zipfile.ZipFile(zip_slblock_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (block_time, items) in enumerate(grouped, 1):
                time_filename = format_time_hh_mm(block_time)
                time_label = format_time_full(block_time)
                
                # 1. ГЕНЕРАЦИЯ SLBLOCK
                total_dur_pure = items['Dur'].sum()
                xml_lines = [
                    f'<slblock Source="list" Type="accurate" Sec="{total_dur_pure:.3f}" Include_subfolders="no" Path="" cptn_start_file="" cptn_end_file="" cptn_between_file="" cptn_start_en="no" cptn_end_en="no" cptn_between_en="no">version 2'
                ]
                
                for _, row in items.iterrows():
                    id_clean = str(row['ID']).split(".")[0]
                    nm = str(row['Name']).strip()
                    ext = "" if any(nm.lower().endswith(e) for e in ['.mov', '.mp4', '.tga', '.mpg']) else ".mov"
                    
                    full_path = f"{user_path}\\{id_clean}_{nm}{ext}"
                    xml_lines.append(f'  <item file="{xml_escape(full_path)}" in="0.000" dur="{float(row["Dur"]):.3f}" />')
                
                xml_lines.append('</slblock>')
                zip_file.writestr(f"{time_filename}.slblock", "\r\n".join(xml_lines).encode('utf-16'))

                # 2. СБОР ID
                id_list_str = "".join([f'"{str(row["ID"]).split(".")[0]}"' for _, row in items.iterrows()])
                txt_id_content.write(f"{time_label}\n{id_list_str}\n\n")
                xlsx_id_data.append({"Время выхода": time_label, "Список ID": id_list_str})
                
                # 3. ТАЙМИНГИ (+20с)
                total_timing_plus_20 = total_dur_pure + 20.0
                xlsx_timing_data.append({
                    "Время выхода блока": time_label, 
                    "Длительность (+20с)": seconds_to_hms(total_timing_plus_20)
                })

        st.success(f"✅ Обработка файла '{base_name}' завершена!")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🚀 Эфирные файлы")
            st.download_button(f"📥 Скачать SLBlocks ({base_name}.zip)", zip_slblock_buffer.getvalue(), f"N4_SLBLOCKS_{base_name}.zip")
        
        with col2:
            st.subheader("📊 Отчеты")
            # Теперь во всех именах файлов присутствует base_name (дата из исходного файла)
            st.download_button(f"📥 Список ID (.txt) - {base_name}", txt_id_content.getvalue(), f"N4_IDs_{base_name}.txt")
            st.download_button(f"📥 Список ID (.xlsx) - {base_name}", to_excel(pd.DataFrame(xlsx_id_data)), f"N4_IDs_Table_{base_name}.xlsx")
            st.download_button(f"📥 Тайминги (.xlsx) - {base_name}", to_excel(pd.DataFrame(xlsx_timing_data)), f"N4_Timings_{base_name}.xlsx")

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
else:
    st.info("👈 Загрузите медиа-план в боковой панели.")
