import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import timedelta, datetime, time

# Конфигурация страницы
st.set_page_config(page_title="N4 | Clean Generator", page_icon="📺", layout="wide")

# Кастомный стиль N4
st.markdown('''
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #E11D48; color: white; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; }
    h1 { color: #E11D48; border-bottom: 2px solid #E11D48; }
    </style>
    ''', unsafe_allow_html=True)

st.title("📺 N4: ГЕНЕРАТОР (БЕЗ ЗАСТАВОК)")

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
    user_path = st.text_input("Путь к роликам:", value=r"I:\RECLAMA 2026")
    st.divider()
    uploaded_file = st.file_uploader("Загрузите Excel (N4)", type=["xls", "xlsx"])
    if user_path.endswith('\\'): user_path = user_path[:-1]

# --- Логика ---
if uploaded_file:
    try:
        base_name = os.path.splitext(uploaded_file.name)[0]
        df = pd.read_excel(uploaded_file, skiprows=6)
        
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
                
                # 1. ГЕНЕРАЦИЯ SLBLOCK (ТОЛЬКО РОЛИКИ)
                total_dur_pure = items['Dur'].sum()
                xml_lines = [
                    f'<slblock Source="list" Type="accurate" Sec="{total_dur_pure:.3f}" Include_subfolders="no" Path="" cptn_start_file="" cptn_end_file="" cptn_between_file="" cptn_start_en="no" cptn_end_en="no" cptn_between_en="no">version 2'
                ]
                
                for _, row in items.iterrows():
                    id_clean = str(row['ID']).split(".")[0]
                    nm = str(row['Name']).strip()
                    ext = "" if any(nm.lower().endswith(e) for e in ['.mov', '.mp4', '.tga', '.mpg']) else ".mov"
                    xml_lines.append(f'  <item file="{xml_escape(user_path)}\\ {id_clean}_{nm}{ext}" in="0.000" dur="{float(row["Dur"]):.3f}" />')
                
                xml_lines.append('</slblock>')
                zip_file.writestr(f"{time_filename}.slblock", "\r\n".join(xml_lines).encode('utf-16'))

                # 2. ОТЧЕТЫ
                id_list_str = "".join([f'"{str(row["ID"]).split(".")[0]}"' for _, row in items.iterrows()])
                txt_id_content.write(f"{time_label}\n{id_list_str}\n\n")
                xlsx_id_data.append({"Время": time_label, "Список ID": id_list_str})
                
                # Тайминг + 20с
                xlsx_timing_data.append({"Время блока": time_label, "Длительность (+20с)": seconds_to_hms(total_dur_pure + 20.0)})

        st.success(f"✅ Файлы для N4 созданы (без заставок). Путь: {user_path}")
        
        c1, col_ отчет = st.columns(2)
        with c1:
            st.download_button(f"🚀 Скачать SLBlocks ({base_name}.zip)", zip_slblock_buffer.getvalue(), f"N4_CLEAN_{base_name}.zip")
        with col_ отчет:
            st.download_button("📥 Список ID (.xlsx)", to_excel(pd.DataFrame(xlsx_id_data)), f"N4_IDs_{base_name}.xlsx")
            st.download_button("📥 Тайминги +20с (.xlsx)", to_excel(pd.DataFrame(xlsx_timing_data)), f"N4_Timings_{base_name}.xlsx")

    except Exception as e:
        st.error(f"Ошибка: {e}")
