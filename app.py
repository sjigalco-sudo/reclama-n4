import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# Конфигурация страницы
st.set_page_config(page_title="N4 | Full Generator", page_icon="📺", layout="wide")

st.markdown('''
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #E11D48; color: white; border: none; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #2563EB; color: white; border: none; }
    h1 { color: #E11D48; border-bottom: 2px solid #E11D48; padding-bottom: 10px; }
    </style>
    ''', unsafe_allow_html=True)

st.title("📺 N4: ПОЛНЫЙ ГЕНЕРАТОР")

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

def seconds_to_hms_custom(total_seconds):
    """Преобразует секунды в формат Ч:ММ:СС (без ведущего нуля для часов)"""
    total_secs = int(round(total_seconds))
    h = total_secs // 3600
    m = (total_secs % 3600) // 60
    s = total_secs % 60
    return f"{h}:{m:02d}:{s:02d}"

def generate_exact_report(timing_rows, file_date):
    """Создает Excel-файл с точным воссозданием разметки, шрифтов и обрамления таблиц"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Лист1"
    
    ws.views.sheetView[0].showGridLines = True
    
    font_regular = Font(name="Calibri", size=11)
    font_bold = Font(name="Calibri", size=11, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_side = Side(border_style="thin", color="000000")
    table_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # 1. Шапка (Дата и Канал в столбце E)
    ws["E1"] = file_date
    ws["E1"].font = font_regular
    
    ws["E2"] = "N4"
    ws["E2"].font = font_bold
    
    # 2. Заголовок над таблицей (строка 6, колонка D)
    ws["D6"] = "Длительность рекламных блоков"
    ws["D6"].font = font_bold
    
    # 3. Шапка таблицы (строка 7) с обрамлением
    ws["C7"] = "Длина ролика"
    ws["C7"].font = font_bold
    ws["C7"].alignment = align_center
    ws["C7"].border = table_border
    
    ws["E7"] = "Название блока"
    ws["E7"].font = font_bold
    ws["E7"].alignment = align_center
    ws["E7"].border = table_border
    
    ws["D7"].border = table_border 
    
    # 4. Заполнение данными (с 9-й строки)
    current_row = 9
    for row in timing_rows:
        ws[f"C{current_row}"] = row["dur"]
        ws[f"C{current_row}"].font = font_regular
        ws[f"C{current_row}"].alignment = align_center
        ws[f"C{current_row}"].border = table_border
        
        ws[f"D{current_row}"].border = table_border
        
        ws[f"E{current_row}"] = row["name"]
        ws[f"E{current_row}"].font = font_regular
        ws[f"E{current_row}"].alignment = align_left
        ws[f"E{current_row}"].border = table_border
        
        current_row += 1
        
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 3
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 5
    ws.column_dimensions['E'].width = 18
