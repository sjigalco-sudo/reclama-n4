import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import timedelta, datetime

# Конфигурация страницы
st.set_page_config(page_title="N4 | Full Generator", page_icon="📺", layout="wide")

st.markdown('''
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #E11D48; color: white; border: none; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; }
    h1 { color: #E11D48; border-bottom: 2px solid #E11D48; padding-bottom: 10px; }
    </style>
    ''', unsafe_allow_html=True)

st.title("📺 N4: ПОЛНЫЙ ГЕНЕРАТОР С ТОЧНЫМ ОТЧЕТОМ")

# --- Константы и База Данных ---
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
    """Преобразует секунды в строгий формат ЧЧ:ММ:СС"""
    total_secs = int(round(total_seconds))
    h = total_secs // 3600
    m = (total_secs % 3600) // 60
    s = total_secs % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

def make_styled_timing_excel(timing_rows, file_date):
    """Генерирует Excel-отчет строго по шаблону пользователя"""
    output = io.BytesIO()
    
    # Строим структуру как в файле пользователя
    # Строка 0: Дата
    # Строка 1: Канал (N4)
    # Строка 5: Заголовок "Длительность рекламных блоков"
    # Строка 6: Названия колонок
    
    data_dict = {
        "": [
            file_date, 
            "N4", 
            "", 
            "", 
            "", 
            "Длительность рекламных блоков", 
            "Длина ролика"
        ],
        " ": [
            "", 
            "", 
            "", 
            "", 
            "", 
            "", 
            ""
        ],
        "  ": [
            "", 
            "", 
            "", 
            "", 
            "", 
            "", 
            "Название блока"
        ]
    }
    
    # Добавляем данные блоков
    for row in timing_rows:
        data_dict[""].append(row["dur"])
        data_dict[" "].append("")
        data_dict["  "].append(row["name"])
        
    df_template = pd.DataFrame(data_dict)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, header=False)
        
    return output.getvalue()

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

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
        st.toast("💾 База MP4 сохранена!")

    st.divider()
    uploaded_file = st.file_uploader("Загрузите медиаплан", type=["xls", "xlsx"])

# --- Логика ---
if uploaded_file:
    try:
        base_name = os.path.splitext(uploaded_file.name)[0]
        
        # Попробуем вытащить дату из оригинального плана (обычно в шапке) или возьмем текущую
        file_date = datetime.now().strftime("%Y-%m-%d")
        
        df = pd.read_excel(uploaded_file, skiprows=6)
        
        # Колонки: 2 (Время), 7 (Длительность), 9 (ID)
        df_res = df.iloc[:, [2, 7, 9]].copy()
        df_res.columns = ['Block_Time', 'Dur', 'ID']
        
        df_res['Block_Time'] = pd.to_datetime(df_res['Block_Time'], format='%H:%M:%S', errors='coerce').dt.time
        df_res['Block_Time'] = df_res['Block_Time'].ffill()
        df_res = df_res.dropna(subset=['ID'])
        
        grouped = df_res.groupby('Block_Time', sort=False)
        
        zip_buffer = io.BytesIO()
        txt_id_content = io.StringIO()
        xlsx_id_data = []
        timing_rows_formatted = [] # Для нового красивого отчета
        hour_counts = {}
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (block_time, items) in enumerate(grouped, 1):
                h = block_time.hour
                hour_counts[h] = hour_counts.get(h, 0) + 1
                
                # Имена для файлов по-прежнему технические (06-1.slblock)
                file_name = f"{h:02d}-{hour_counts[h]}"
                time_str = block_time.strftime('%H:%M:%S')
                
                # Имя для отчета в формате "Реклама 6.1"
                report_block_name = f"Реклама {h}.{hour_counts[h]}"
                
                soc = SOCIAL_ADS[((i - 1) % 5) + 1]
                
                # Расчет полной длительности блока (план + заставки)
                pure_ads_seconds = items['Dur'].sum()
                total_block_dur = PUB_DUR + pure_ads_seconds + PUB_DUR + soc['dur']
                
                # Сохраняем в список для красивого Excel по твоему шаблону
                timing_rows_formatted.append({
                    "dur": seconds_to_hms(total_block_dur),
                    "name": report_block_name
                })

                # Сохраняем данные для списков ID
                id_list_raw = [str(row['ID']).split(".")[0] for _, row in items.iterrows()]
                id_list_str = "".join([f'"{x}"' for x in id_list_raw])
                
                txt_id_content.write(f"{time_str}\n{id_list_str}\n\n")
                xlsx_id_data.append({"Время": time_str, "Список ID": id_list_str})

                # XML SLBlock
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

        st.success(f"✅ Готово! Сформировано блоков: {len(timing_rows_formatted)}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🚀 Эфирные файлы")
            st.download_button(f"📥 SLBlocks ({base_name}).zip", zip_buffer.getvalue(), f"N4_Blocks_{base_name}.zip")
        
        with col2:
            st.subheader("📊 Отчетность")
            # Выдача нового отчета по шаблону
            st.download_button("📥 Скачать Тайминги (по шаблону .xlsx)", make_styled_timing_excel(timing_rows_formatted, file_date), f"N4_Timings_{base_name}.xlsx")
            st.download_button("📥 Список ID (.xlsx)", to_excel(pd.DataFrame(xlsx_id_data)), f"N4_IDs_{base_name}.xlsx")
            st.download_button("📥 Список ID (.txt)", txt_id_content.getvalue(), f"N4_IDs_{base_name}.txt")

    except Exception as e:
        st.error(f"Ошибка: {e}")
