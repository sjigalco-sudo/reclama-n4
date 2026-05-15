import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import timedelta, datetime, time

# Конфигурация страницы
st.set_page_config(page_title="N4 | Full Generator", page_icon="📺", layout="wide")

# Кастомный стиль N4
st.markdown('''
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #E11D48; color: white; border: none; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; }
    h1 { color: #E11D48; border-bottom: 2px solid #E11D48; padding-bottom: 10px; }
    div[data-testid="stExpander"], .stFileUploader { border: 1px solid #30363d; border-radius: 8px; }
    </style>
    ''', unsafe_allow_html=True)

st.title("📺 N4: ПОЛНЫЙ ГЕНЕРАТОР (Blocks + IDs + Timings)")

# --- Константы N4 ---
DEFAULT_REKLAMA_PATH = r"D:\AIR\REKLAMA 2026"
DEFAULT_SOCIAL_PATH = r"D:\AIR\REKLAMA 2025"
PUB_FILE = "PUBLICITATE_HD.mp4"
PUB_DUR = 5.000

SOCIAL_ADS = {
    1: {"file": "1_APA_HD.mpg", "dur": 10.440},
    2: {"file": "2_FRUCTE_HD.mpg", "dur": 10.440},
    3: {"file": "3_MESE HD.mpg", "dur": 10.440},
    4: {"file": "4_MISCARE_HD.mpg", "dur": 10.440},
    5: {"file": "5_SARE_HD.mpg", "dur": 10.440}
}

def xml_escape(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def seconds_to_hms(total_seconds):
    """Преобразует секунды в формат Ч:ММ:СС"""
    td = timedelta(seconds=int(round(total_seconds)))
    return str(td)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- Sidebar ---
with st.sidebar:
    st.header("⚙️ Настройки")
    path_ads = st.text_input("Путь к роликам рекламы:", value=DEFAULT_REKLAMA_PATH)
    path_soc = st.text_input("Путь к отбивкам/соц:", value=DEFAULT_SOCIAL_PATH)
    
    st.divider()
    st.header("🎥 Форматы")
    mp4_ids_input = st.text_area("ID для MP4 (через запятую):", value="6856, 7142")
    mp4_ids = [x.strip() for x in mp4_ids_input.split(",") if x.strip()]
    
    st.divider()
    uploaded_file = st.file_uploader("Загрузите план (XLS, XLSX)", type=["xls", "xlsx"])
    
    if path_ads.endswith('\\'): path_ads = path_ads[:-1]
    if path_soc.endswith('\\'): path_soc = path_soc[:-1]

# --- Логика ---
if uploaded_file:
    try:
        base_name = os.path.splitext(uploaded_file.name)[0]
        df = pd.read_excel(uploaded_file, skiprows=6)
        
        df_res = df.iloc[:, [2, 6, 7, 9]].copy()
        df_res.columns = ['Block_Time', 'Name', 'Dur', 'ID']
        
        df_res['Block_Time'] = pd.to_datetime(df_res['Block_Time'], format='%H:%M:%S', errors='coerce').dt.time
        df_res['Block_Time'] = df_res['Block_Time'].ffill()
        df_res = df_res.dropna(subset=['ID'])
        
        grouped = df_res.groupby('Block_Time', sort=False)
        
        zip_buffer = io.BytesIO()
        txt_id_content = io.StringIO()
        xlsx_id_data = []
        timing_data = []
        hour_counts = {}
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (block_time, items) in enumerate(grouped, 1):
                h = block_time.hour
                hour_counts[h] = hour_counts.get(h, 0) + 1
                file_name = f"{h:02d}-{hour_counts[h]}"
                time_str = block_time.strftime('%H:%M:%S')
                
                soc_idx = ((i - 1) % 5) + 1
                soc = SOCIAL_ADS[soc_idx]
                
                # Полная длительность (реклама + 20.440с заставок)
                total_block_dur = PUB_DUR + items['Dur'].sum() + PUB_DUR + soc['dur']
                
                # 1. Данные для Excel Таймингов
                timing_data.append({
                    "Блок": file_name,
                    "Время выхода": time_str,
                    "Длительность (Ч:ММ:СС)": seconds_to_hms(total_block_dur)
                })

                # 2. Данные для отчетов по ID
                id_list_raw = [str(row['ID']).split(".")[0] for _, row in items.iterrows()]
                id_list_str = "".join([f'"{x}"' for x in id_list_raw])
                
                txt_id_content.write(f"{time_str}\n{id_list_str}\n\n")
                xlsx_id_data.append({"Время": time_str, "Список ID": id_list_str})

                # 3. XML структура SLBlock
                xml = [
                    f'<slblock\r\n      Source="list"\r\n      Type="accurate"\r\n      Image_using_type="Video files"\r\n      Sec="{total_block_dur:.3f}"\r\n      Include_subfolders="no"\r\n      Path=""\r\n      cptn_start_file=""\r\n      cptn_end_file=""\r\n      cptn_between_file=""\r\n      cptn_start_en="no"\r\n      cptn_end_en="no"\r\n      cptn_between_en="no"\r\n      Image_Duration="1.000">\r\n',
                    '      version 3\r\n',
                    f'      <item\r\n            file="{xml_escape(path_soc)}\\{PUB_FILE}"\r\n            in="0.000"\r\n            dur="{PUB_DUR:.3f}"/>\r\n'
                ]
                
                for _, row in items.iterrows():
                    id_clean = str(row['ID']).split(".")[0]
                    ext = ".mp4" if id_clean in mp4_ids else ".mov"
                    xml.append(f'      <item\r\n            file="{xml_escape(path_ads)}\\{id_clean}{ext}"\r\n            in="0.000"\r\n            dur="{float(row["Dur"]):.3f}"/>\r\n')
                
                xml.append(f'      <item\r\n            file="{xml_escape(path_soc)}\\{PUB_FILE}"\r\n            in="0.000"\r\n            dur="{PUB_DUR:.3f}"/>\r\n')
                xml.append(f'      <item\r\n            file="{xml_escape(path_soc)}\\{soc["file"]}"\r\n            in="0.000"\r\n            dur="{soc["dur"]:.3f}"/>\r\n')
                xml.append('</slblock>')
                
                zip_file.writestr(f"{file_name}.slblock", "".join(xml).encode('utf-16'))

        st.success(f"✅ Обработка завершена! Блоков: {len(timing_data)}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🚀 Эфирные файлы")
            st.download_button(f"📥 SLBlocks ({base_name}).zip", zip_buffer.getvalue(), f"N4_Blocks_{base_name}.zip")
        
        with col2:
            st.subheader("📊 Отчетность")
            st.download_button("📥 Тайминги Ч:ММ:СС (.xlsx)", to_excel(pd.DataFrame(timing_data)), f"N4_Timings_{base_name}.xlsx")
            st.download_button("📥 Список ID (.xlsx)", to_excel(pd.DataFrame(xlsx_id_data)), f"N4_IDs_{base_name}.xlsx")
            st.download_button("📥 Список ID (.txt)", txt_id_content.getvalue(), f"N4_IDs_{base_name}.txt")

    except Exception as e:
        st.error(f"Ошибка: {e}")
