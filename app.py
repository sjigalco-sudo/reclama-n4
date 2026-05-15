import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import timedelta, datetime
import openpyxl
from openpyxl.styles import Font, Alignment

# Конфигурация страницы
st.set_page_config(page_title="N4 | Full Generator", page_icon="📺", layout="wide")

st.markdown('''
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #E11D48; color: white; border: none; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; }
    h1 { color: #E11D48; border-bottom: 2px solid #E11D48; padding-bottom: 10px; }
    </style>
    ''', unsafe_allow_html=True)

st.title("📺 N4: ПОЛНЫЙ ГЕНЕРАТОР (Точный Визуальный Отчет)")

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
    """Преобразует секунды в формат ЧЧ:ММ:СС"""
    total_secs = int(round(total_seconds))
    h = total_secs // 3600
    m = (total_secs % 3600) // 60
    s = total_secs % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

def generate_exact_report(timing_rows, file_date):
    """Создает Excel-файл, полностью копируя визуальную структуру оригинала"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Лист1"
    
    # Включаем сетку (опционально, можно выключить: ws.views.sheetView[0].showGridLines = False)
    ws.views.sheetView[0].showGridLines = True
    
    # Стили
    font_regular = Font(name="Calibri", size=11)
    font_bold = Font(name="Calibri", size=11, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    # 1. Шапка отчета (строки 1 и 2 в колонке E)
    ws["E1"] = file_date
    ws["E1"].font = font_regular
    
    ws["E2"] = "N4"
    ws["E2"].font = font_bold
    
    # 2. Заголовок таблицы (строка 6, колонка D)
    ws["D6"] = "Длительность рекламных блоков"
    ws["D6"].font = font_bold
    
    # 3. Названия колонок (строка 7)
    ws["C7"] = "Длина ролика"
    ws["C7"].font = font_bold
    ws["C7"].alignment = align_center
    
    ws["E7"] = "Название блока"
    ws["E7"].font = font_bold
    ws["E7"].alignment = align_center
    
    # 4. Заполнение данными (начиная со строки 9)
    current_row = 9
    for row in timing_rows:
        ws[f"C{current_row}"] = row["dur"]
        ws[f"C{current_row}"].font = font_regular
        ws[f"C{current_row}"].alignment = align_center
        
        ws[f"E{current_row}"] = row["name"]
        ws[f"E{current_row}"].font = font_regular
        ws[f"E{current_row}"].alignment = align_left
        
        current_row += 1
        
    # Выставляем красивую ширину столбцов, чтобы текст не влезал на соседние ячейки
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 3
    ws.column_dimensions['C'].width = 16  # Для таймингов "00:06:15"
    ws.column_dimensions['D'].width = 5   # Пустой разделитель
    ws.column_dimensions['E'].width = 18  # Для названий "Реклама 6.1"
    
    # Сохраняем в байтовый поток
    output = io.BytesIO()
    wb.save(output)
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
        
        # Пробуем автоматически достать дату из названия файла (например, "publicitate 17.05.2026")
        # Если в названии есть дата, запишем её, иначе текущую дату на компьютере
        file_date = datetime.now().strftime("%Y-%m-%d")
        for part in base_name.split():
            if len(part) == 10 and part.count('.') == 2:
                file_date = part  # Найдена дата формата ДД.ММ.ГГГГ
        
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
        timing_rows_formatted = []
        hour_counts = {}
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (block_time, items) in enumerate(grouped, 1):
                h = block_time.hour
                hour_counts[h] = hour_counts.get(h, 0) + 1
                
                file_name = f"{h:02d}-{hour_counts[h]}"
                time_str = block_time.strftime('%H:%M:%S')
                report_block_name = f"Реклама {h}.{hour_counts[h]}"
                
                soc = SOCIAL_ADS[((i - 1) % 5) + 1]
                
                pure_ads_seconds = items['Dur'].sum()
                total_block_dur = PUB_DUR + pure_ads_seconds + PUB_DUR + soc['dur']
                
                timing_rows_formatted.append({
                    "dur": seconds_to_hms(total_block_dur),
                    "name": report_block_name
                })

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

        st.success(f"✅ Успешно обработано! Файлы готовы.")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🚀 Эфирные файлы")
            st.download_button(f"📥 SLBlocks ({base_name}).zip", zip_buffer.getvalue(), f"N4_Blocks_{base_name}.zip")
        
        with col2:
            st.subheader("📊 Отчетность")
            # Кнопка отчета в оригинальном стиле
            st.download_button("📥 Визуальный Отчет Таймингов (.xlsx)", generate_exact_report(timing_rows_formatted, file_date), f"N4_Timings_{base_name}.xlsx")
            st.download_button("📥 Список ID (.xlsx)", to_excel(pd.DataFrame(xlsx_id_data)), f"N4_IDs_{base_name}.xlsx")
            st.download_button("📥 Список ID (.txt)", txt_id_content.getvalue(), f"N4_IDs_{base_name}.txt")

    except Exception as e:
        st.error(f"Ошибка: {e}")
