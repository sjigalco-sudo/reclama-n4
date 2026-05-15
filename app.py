import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import timedelta

# Конфигурация
st.set_page_config(page_title="N4 | Full Generator", page_icon="📺", layout="wide")

st.markdown('''
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #E11D48; color: white; border: none; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; }
    h1 { color: #E11D48; border-bottom: 2px solid #E11D48; padding-bottom: 10px; }
    </style>
    ''', unsafe_allow_html=True)

st.title("📺 N4: ИСПРАВЛЕННЫЙ ГЕНЕРАТОР")

# --- Константы и БД ---
DB_FILE = "mp4_database.txt"
PUB_FILE = "PUBLICITATE_HD.mp4"
PUB_DUR = 5.000
SOCIAL_ADS = {
    1: {"file": "1_APA_HD.mpg", "dur": 10.440},
    2: {"file": "2_FRUCTE_HD.mpg", "dur": 10.440},
    3: {"file": "3_MESE HD.mpg", "dur": 10.440},
    4: {"file": "4_MISCARE_HD.mpg", "dur": 10.440},
    5: {"file": "5_SARE_HD.mpg", "dur": 10.440}
}

def load_mp4_ids():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return sorted(list(set([line.strip() for line in f if line.strip()])))
    return ["6856", "7142"]

def save_mp4_ids(id_list):
    cleaned = sorted(list(set([x.strip() for x in id_list if x.strip()])))
    with open(DB_FILE, "w", encoding="utf-8") as f:
        for x in cleaned: f.write(f"{x}\n")
    return cleaned

saved_mp4 = load_mp4_ids()

def xml_escape(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def seconds_to_hms(total_seconds):
    td = timedelta(seconds=int(round(total_seconds)))
    return str(td)

# --- Sidebar ---
with st.sidebar:
    st.header("⚙️ Настройки")
    path_ads = st.text_input("Путь к рекламе:", value=r"D:\AIR\REKLAMA 2026")
    path_soc = st.text_input("Путь к заставкам:", value=r"D:\AIR\REKLAMA 2025")
    
    st.divider()
    current_input = st.text_area("ID для MP4 (через запятую):", value=", ".join(saved_mp4))
    mp4_ids = [x.strip() for x in current_input.split(",") if x.strip()]
    if sorted(mp4_ids) != sorted(saved_mp4):
        saved_mp4 = save_mp4_ids(mp4_ids)
        st.toast("💾 База MP4 обновлена!")

    uploaded_file = st.file_uploader("Загрузите медиаплан", type=["xls", "xlsx"])

# --- Обработка ---
if uploaded_file:
    try:
        base_name = os.path.splitext(uploaded_file.name)[0]
        df = pd.read_excel(uploaded_file, skiprows=6)
        
        # Берем Колонку 2 (Время), 7 (Длительность), 9 (ID)
        df_res = df.iloc[:, [2, 7, 9]].copy()
        df_res.columns = ['Block_Time', 'Dur', 'ID']
        
        # Наполняем пустые ячейки времени (FFILL)
        df_res['Block_Time'] = pd.to_datetime(df_res['Block_Time'], format='%H:%M:%S', errors='coerce').dt.time
        df_res['Block_Time'] = df_res['Block_Time'].ffill()
        df_res = df_res.dropna(subset=['ID'])
        
        grouped = df_res.groupby('Block_Time', sort=False)
        
        zip_buffer = io.BytesIO()
        timing_data = []
        hour_counts = {}
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (block_time, items) in enumerate(grouped, 1):
                h = block_time.hour
                hour_counts[h] = hour_counts.get(h, 0) + 1
                file_name = f"{h:02d}-{hour_counts[h]}"
                
                soc = SOCIAL_ADS[((i - 1) % 5) + 1]
                
                # ВАЖНО: Считаем длительность на основе суммы Dur из медиаплана
                # Плюс реальные длительности заставок
                pure_ads_seconds = items['Dur'].sum()
                total_block_dur = PUB_DUR + pure_ads_seconds + PUB_DUR + soc['dur']
                
                # Данные для Excel отчета
                timing_data.append({
                    "Блок": file_name,
                    "Время (план)": block_time.strftime('%H:%M:%S'),
                    "Длительность (Ч:ММ:СС)": seconds_to_hms(total_block_dur)
                })

                # Формируем XML (UTF-16)
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

        st.success(f"✅ Готово! Параметр Sec в slblock теперь равен сумме длин роликов.")
        
        # Кнопки скачивания
        st.download_button(f"📥 Скачать SLBlocks ({base_name})", zip_buffer.getvalue(), f"N4_Blocks_{base_name}.zip")
        
        df_timing = pd.DataFrame(timing_data)
        output_ex = io.BytesIO()
        with pd.ExcelWriter(output_ex, engine='openpyxl') as writer:
            df_timing.to_excel(writer, index=False)
        st.download_button("📥 Скачать Тайминги (.xlsx)", output_ex.getvalue(), f"N4_Timings_{base_name}.xlsx")

    except Exception as e:
        st.error(f"Ошибка: {e}")
