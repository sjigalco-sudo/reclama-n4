import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# Конфигурация страницы
st.set_page_config(page_title="N4 | Generator", page_icon="📺", layout="wide")
st.markdown('''<style>.stButton>button, .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; }</style>''', unsafe_allow_html=True)

st.title("📺 N4: ПОЛНЫЙ ГЕНЕРАТОР")

# Константы и функции БД
DB_FILE = "mp4_database.txt"
PUB_FILE = "PUBLICITATE_HD.mp4"
PUB_DUR = 5.000
SOCIAL_ADS = {1: {"file": "1_APA_HD.mpg", "dur": 10.440}, 2: {"file": "2_FRUCTE_HD.mpg", "dur": 10.440}, 3: {"file": "3_MESE HD.mpg", "dur": 10.440}, 4: {"file": "4_MISCARE_HD.mpg", "dur": 10.440}, 5: {"file": "5_SARE_HD.mpg", "dur": 10.440}}

def load_mp4_ids():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f: return sorted(list(set([line.strip() for line in f if line.strip()])))
    return ["6856", "7142"]

def save_mp4_ids(id_list):
    cleaned = sorted(list(set([x.strip() for x in id_list if x.strip()])))
    with open(DB_FILE, "w", encoding="utf-8") as f:
        for x in cleaned: f.write(f"{x}\n")
    return cleaned

# Функция генерации Excel
def generate_exact_report(timing_rows, file_date):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.views.sheetView[0].showGridLines = False 
    f_reg, f_bold = Font(name="Calibri", size=11), Font(name="Calibri", size=11, bold=True)
    align_c, align_l = Alignment(horizontal="center", vertical="center"), Alignment(horizontal="left", vertical="center", indent=1.0)
    fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    ws["E1"], ws["E2"] = file_date, "N4"
    ws.merge_cells("C6:E6"); ws["C6"] = "Длительность рекламных блоков"
    ws["C6"].font, ws["C6"].alignment, ws["C6"].border = f_bold, align_c, border
    ws["C8"], ws["E8"] = "Длина ролика", "Название блока"
    for c in ["C8", "D8", "E8"]: ws[c].font, ws[c].alignment, ws[c].border, ws[c].fill = f_bold, align_c, border, fill
    for idx, row in enumerate(timing_rows, start=9):
        ws.row_dimensions[idx].height = 15.0
        ws[f"C{idx}"], ws[f"E{idx}"] = row["dur"], row["name"]
        for col in ["C", "D", "E"]: ws[f"{col}{idx}"].border = border
        ws[f"C{idx}"].alignment, ws[f"E{idx}"].alignment = align_c, align_l
    ws.column_dimensions['C'].width, ws.column_dimensions['D'].width, ws.column_dimensions['E'].width = 14.5, 2.5, 21.0
    out = io.BytesIO(); wb.save(out); return out.getvalue()

# Sidebar
with st.sidebar:
    path_air = st.text_input("Путь к файлам:", value=r"D:\AIR\REKLAMA 2026")
    saved_mp4 = load_mp4_ids()
    current_input = st.text_area("ID для MP4:", value=", ".join(saved_mp4))
    new_ids = [x.strip() for x in current_input.split(",") if x.strip()]
    if new_ids != saved_mp4: saved_mp4 = save_mp4_ids(new_ids)
    uploaded_file = st.file_uploader("Загрузите медиаплан", type=["xls", "xlsx"])

# Основная логика
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
    
    # Исправление ошибки с Excel
    xlsx_id_buffer = io.BytesIO()
    pd.DataFrame(xlsx_id_list).to_excel(xlsx_id_buffer, index=False)
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.download_button("📥 Блоки (.zip)", zip_buffer.getvalue(), f"Blocks_{uploaded_file.name}.zip")
    with c2: st.download_button("📥 Отчет таймингов", generate_exact_report(timing_rows, datetime.now().strftime("%d.%m.%Y")), f"Timings_{uploaded_file.name}.xlsx")
    with c3: st.download_button("📥 Список ID (.xlsx)", xlsx_id_buffer.getvalue(), f"IDs_{uploaded_file.name}.xlsx")
    with c4: st.download_button("📥 Список ID (.txt)", txt_buffer.getvalue(), f"IDs_{uploaded_file.name}.txt")
