import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# Конфигурация
st.set_page_config(page_title="N4 | Generator", page_icon="📺", layout="wide")
st.markdown('''<style>.stButton>button, .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; margin-bottom: 10px; }</style>''', unsafe_allow_html=True)

st.title("📺 N4: ПОЛНЫЙ ГЕНЕРАТОР")

# Константы
DB_FILE = "mp4_database.txt"
PUB_FILE = "PUBLICITATE_HD.mp4"
PUB_DUR = 5.000
SOCIAL_ADS = {1: {"file": "1_APA_HD.mpg", "dur": 10.440}, 2: {"file": "2_FRUCTE_HD.mpg", "dur": 10.440}, 3: {"file": "3_MESE HD.mpg", "dur": 10.440}, 4: {"file": "4_MISCARE_HD.mpg", "dur": 10.440}, 5: {"file": "5_SARE_HD.mpg", "dur": 10.440}}

# Функции БД
def load_mp4_ids():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f: return sorted(list(set([line.strip() for line in f if line.strip()])))
    return ["6856", "7142"]

def save_mp4_ids(id_list):
    cleaned = sorted(list(set([x.strip() for x in id_list if x.strip()])))
    with open(DB_FILE, "w", encoding="utf-8") as f:
        for x in cleaned: f.write(f"{x}\n")
    return cleaned

# Исправленная функция отчета
def generate_exact_report(timing_rows, file_date):
    wb = openpyxl.Workbook(); ws = wb.active
    f_reg, f_bold = Font(name="Calibri", size=11), Font(name="Calibri", size=11, bold=True)
    align_c, align_l = Alignment(horizontal="center", vertical="center"), Alignment(horizontal="left", vertical="center", indent=1.0)
    fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    # Размеры
    ws.column_dimensions['A'].width = 8.09
    ws.column_dimensions['B'].width = 8.09
    ws.column_dimensions['C'].width = 8.09
    ws.column_dimensions['D'].width = 8.09
    ws.column_dimensions['E'].width = 62.34
    
    ws["E1"], ws["E2"] = file_date, "N4"
    ws.merge_cells("C6:E6"); ws["C6"] = "Длительность рекламных блоков"
    ws["C6"].font, ws["C6"].alignment = f_bold, align_c
    for col in ["C", "D", "E"]: ws[f"{col}6"].border = border

    # Шапка
    for col, val in [("C", "Длина ролика"), ("D", ""), ("E", "Название блока")]:
        cell = ws[f"{col}8"]
        cell.value = val
        cell.font, cell.alignment, cell.border, cell.fill = f_bold, align_c, border, fill
        
    # Данные
    for idx, row in enumerate(timing_rows, start=9):
        ws.row_dimensions[idx].height = 15.0
        for col in ["C", "D", "E"]:
            cell = ws[f"{col}{idx}"]
            cell.border = border
            if col == "C": cell.value = row["dur"]; cell.alignment = align_c
            elif col == "E": cell.value = row["name"]; cell.alignment = align_l
            
    out = io.BytesIO(); wb.save(out); return out.getvalue()

# Sidebar
with st.sidebar:
    path_air = st.text_input("Путь к файлам:", value=r"D:\AIR\REKLAMA 2026")
    saved_mp4 = load_mp4_ids()
    current_input = st.text_area("ID для MP4:", value=", ".join(saved_mp4))
    new_ids = [x.strip() for x in current_input.split(",") if x.strip()]
    if new_ids != saved_mp4: saved_mp4 = save_mp4_ids(new_ids)
    uploaded_file = st.file_uploader("Загрузите медиаплан", type=["xls", "xlsx"])

# Логика
if uploaded_file:
    df = pd.read_excel(uploaded_file, skiprows=6)
    df_res = df.iloc[:, [2, 7, 9]].copy()
    df_res.columns = ['Block_Time', 'Dur', 'ID']
    df_res['Block_Time'] = pd.to_datetime(df_res['Block_Time'], format='%H:%M:%S', errors='coerce').dt.time.ffill()
    grouped = df_res.groupby('Block_Time', sort=False)
    
    zip_buffer, txt_buffer, xlsx_id_list, timing_rows = io.BytesIO(), io.StringIO(), [], []
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for i, (bt, items) in enumerate(grouped, 1):
            h = bt.hour; report_hour = 24 if h == 0 else h
            name = f"Реклама {report_hour}.{i}"; time_s = f"{report_hour}:{bt.strftime('%M:%S')}"
            valid = items.dropna(subset=['ID'])
            if not valid.empty:
                total_dur = PUB_DUR * 2 + valid['Dur'].sum() + SOCIAL_ADS[((i-1)%5)+1]['dur']
                xml = [f'<slblock Source="list" Type="accurate" Sec="{total_dur:.3f}">\nversion 3\n', f'<item file="{path_air}\\{PUB_FILE}" dur="{PUB_DUR:.3f}"/>\n']
                for _, r in valid.iterrows(): xml.append(f'<item file="{path_air}\\{str(r["ID"]).split(".")[0]}{".mp4" if str(r["ID"]).split(".")[0] in saved_mp4 else ".mov"}" dur="{float(r["Dur"]):.3f}"/>\n')
                xml.append(f'<item file="{path_air}\\{PUB_FILE}" dur="{PUB_DUR:.3f}"/>\n<item file="{path_air}\\{SOCIAL_ADS[((i-1)%5)+1]["file"]}" dur="{SOCIAL_ADS[((i-1)%5)+1]["dur"]:.3f}"/>\n</slblock>')
                zip_file.writestr(f"{report_hour:02d}-{i}.slblock", "".join(xml).encode('utf-16le'))
                timing_rows.append({"dur": f"{int(total_dur)//3600:02d}:{int(total_dur)%3600//60:02d}:{int(total_dur)%60:02d}", "name": name})
                id_str = "".join([f'"{str(r['ID']).split('.')[0]}"' for _, r in valid.iterrows()])
            else:
                timing_rows.append({"dur": "00:00:00", "name": name}); id_str = ""
            txt_buffer.write(f"{time_s}\n{id_str}\n\n"); xlsx_id_list.append({"Время": time_s, "Список ID": id_str})

    st.success("✅ Готово!")
    xlsx_id_buffer = io.BytesIO()
    pd.DataFrame(xlsx_id_list).to_excel(xlsx_id_buffer, index=False)
    
    col_l, col_r = st.columns([3, 1])
    with col_l:
        st.download_button("📥 СКАЧАТЬ АРХИВ БЛОКОВ (.zip)", zip_buffer.getvalue(), f"Blocks_{uploaded_file.name}.zip")
    with col_r:
        st.download_button("📥 Отчет: Тайминги (.xlsx)", generate_exact_report(timing_rows, datetime.now().strftime("%d.%m.%Y")), f"Timings_{uploaded_file.name}.xlsx")
        st.download_button("📥 Отчет: Список ID (.xlsx)", xlsx_id_buffer.getvalue(), f"IDs_{uploaded_file.name}.xlsx")
        st.download_button("📥 Отчет: Список ID (.txt)", txt_buffer.getvalue(), f"IDs_{uploaded_file.name}.txt")
