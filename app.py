import streamlit as st
import pandas as pd
import io
import zipfile
import os

# --- КОНСТАНТЫ N4 ---
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

def inject_custom_css():
    st.markdown("""
        <style>
        .main { background-color: #0d1117; }
        h1 { color: #e6edf3; font-weight: 700; border-bottom: 2px solid #30363d; padding-bottom: 10px; }
        section[data-testid="stSidebar"] { background-color: #161b22 !important; border-right: 1px solid #30363d; }
        .stButton>button { width: 100%; border-radius: 6px; background-color: #21262d; color: #c9d1d9; }
        .stDownloadButton>button { width: 100%; background-color: #238636 !important; color: white !important; font-weight: 700; }
        div[data-testid="stExpander"], .stFileUploader { border: 1px solid #30363d; border-radius: 8px; background-color: #0d1117; }
        </style>
    """, unsafe_allow_html=True)

def xml_escape(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

st.set_page_config(page_title="N4 Generator", page_icon="📺", layout="wide")
inject_custom_css()

st.markdown("<h1 style='text-align: center;'>N4: ГЕНЕРАТОР РЕКЛАМНЫХ БЛОКОВ</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ НАСТРОЙКИ ПУТЕЙ")
    path_ads = st.text_input("Путь к роликам (REKLAMA 2026):", value=DEFAULT_REKLAMA_PATH)
    path_soc = st.text_input("Путь к отбивкам (REKLAMA 2025):", value=DEFAULT_SOCIAL_PATH)
    
    st.divider()
    st.markdown("### 🎥 ФОРМАТЫ")
    # Список исключений для MP4 как в Global 24
    mp4_ids_input = st.text_area("ID для MP4 (через запятую):", value="6856, 7142")
    mp4_ids = [x.strip() for x in mp4_ids_input.split(",") if x.strip()]
    
    st.divider()
    uploaded_file = st.file_uploader("Загрузите план N4 (XLS, XLSX)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        source_name = os.path.splitext(uploaded_file.name)[0]
        df = pd.read_excel(uploaded_file, skiprows=6)
        df_res = df.iloc[:, [2, 6, 7, 9]].copy()
        df_res.columns = ['Time', 'Name', 'Dur', 'ID']
        
        # Обработка времени и заполнение пропусков
        df_res['Time'] = pd.to_datetime(df_res['Time'], format='%H:%M:%S', errors='coerce').dt.time
        df_res['Time'] = df_res['Time'].ffill()
        df_res = df_res.dropna(subset=['ID'])

        zip_buffer = io.BytesIO()
        summary_data = []
        hour_counts = {}

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            groups = df_res.groupby('Time', sort=False)
            
            for i, (block_time, items) in enumerate(groups, 1):
                h = block_time.hour
                hour_counts[h] = hour_counts.get(h, 0) + 1
                
                # Название файла 06-1, 07-2 и т.д.
                file_name_slblock = f"{h:02d}-{hour_counts[h]}"
                
                soc_idx = ((i - 1) % 5) + 1
                soc = SOCIAL_ADS[soc_idx]
                total_sec = PUB_DUR + items['Dur'].sum() + PUB_DUR + soc['dur']
                
                # Формирование структуры XML v3 с отступами
                xml = [
                    f'<slblock\r\n      Source="list"\r\n      Type="accurate"\r\n      Image_using_type="Video files"\r\n      Sec="{total_sec:.3f}"\r\n      Include_subfolders="no"\r\n      Path=""\r\n      cptn_start_file=""\r\n      cptn_end_file=""\r\n      cptn_between_file=""\r\n      cptn_start_en="no"\r\n      cptn_end_en="no"\r\n      cptn_between_en="no"\r\n      Image_Duration="1.000">\r\n',
                    '      version 3\r\n',
                    f'      <item\r\n            file="{xml_escape(path_soc)}\\{PUB_FILE}"\r\n            in="0.000"\r\n            dur="{PUB_DUR:.3f}"/>\r\n'
                ]

                for _, r in items.iterrows():
                    raw_id = str(r['ID']).split(".")[0]
                    # Проверка исключения: если ID в списке, ставим .mp4, иначе .mov
                    ext = ".mp4" if raw_id in mp4_ids else ".mov"
                    xml.append(f'      <item\r\n            file="{xml_escape(path_ads)}\\{raw_id}{ext}"\r\n            in="0.000"\r\n            dur="{float(r["Dur"]):.3f}"/>\r\n')

                # Закрывающая часть
                xml.append(f'      <item\r\n            file="{xml_escape(path_soc)}\\{PUB_FILE}"\r\n            in="0.000"\r\n            dur="{PUB_DUR:.3f}"/>\r\n')
                xml.append(f'      <item\r\n            file="{xml_escape(path_soc)}\\{soc["file"]}"\r\n            in="0.000"\r\n            dur="{soc["dur"]:.3f}"/>\r\n')
                xml.append('</slblock>')

                zip_file.writestr(f"{file_name_slblock}.slblock", "".join(xml).encode('utf-16'))
                summary_data.append([file_name_slblock, block_time.strftime('%H:%M'), f"{total_sec:.3f}"])

        st.success(f"Архив N4 успешно создан!")
        
        if summary_data:
            with st.expander("📝 Список сгенерированных файлов", expanded=True):
                st.table(pd.DataFrame(summary_data, columns=["Файл", "Время", "Длительность (сек)"]))

        st.download_button(
            label=f"📥 СКАЧАТЬ АРХИВ: {source_name}.zip",
            data=zip_buffer.getvalue(),
            file_name=f"N4_{source_name}.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"Произошла ошибка: {e}")
