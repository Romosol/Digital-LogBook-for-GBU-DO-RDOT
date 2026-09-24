#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Графический интерфейс на Python (Tkinter + ttk) для формирования обложки классного журнала.
Формат страницы: А4 под подшивку (левое поле 30 мм).
Единственный инпут: Начальный / Конечный год учебного года.

Особенности:
- Единственный инпут даты: выпадающий список истории прямо при нажатии на поле ввода (ttk.Combobox)
- Хранение истории в файле config.json
- Формирование Excel вынесено в отдельную вкладку приложения
- Аппаратная поддержка High-DPI (Windows 10/11, macOS Retina, Linux)
- Экспорт в Excel (.xlsx) через openpyxl

Установка зависимостей:
    pip install openpyxl

Запуск:
    python app_gui.py
"""

import os
import sys
import re
import json
import platform
import subprocess
import threading
import urllib.request
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Версия приложения и репозиторий GitHub для проверки обновлений
APP_VERSION = "1.0.1"
DEFAULT_GITHUB_REPO = "Romosol/Digital-LogBook-for-GBU-DO-RDOT"


def get_app_directory() -> str:
    """
    Возвращает постоянную директорию, где находится исполняемый файл:
    - для скомпилированного .exe (PyInstaller): папка, где лежит сам .exe
    - для обычного скрипта Python: папка со скриптом app_gui.py
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_directory()
CONFIG_FILE = os.path.join(APP_DIR, "config.json")


def get_app_version() -> str:
    """Определяет версию приложения (из version.txt или константы APP_VERSION)."""
    candidates = []
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        candidates.append(os.path.join(sys._MEIPASS, "version.txt"))
    candidates.append(os.path.join(APP_DIR, "version.txt"))

    for c in candidates:
        if os.path.exists(c):
            try:
                with open(c, "r", encoding="utf-8") as f:
                    v = f.read().strip()
                    if v:
                        return v.lstrip("v")
            except Exception:
                pass
    return APP_VERSION


APP_VERSION = get_app_version()


def parse_version_tuple(v_str: str) -> tuple:
    """Извлекает кортеж чисел из строки версии, например 'v1.2.3' -> (1, 2, 3)."""
    if not v_str:
        return (0, 0, 0)
    nums = re.findall(r'\d+', str(v_str))
    if not nums:
        return (0, 0, 0)
    return tuple(int(n) for n in nums)


def setup_high_dpi():
    """Активирует режим аппаратной поддержки High-DPI в Windows."""
    if platform.system() == "Windows":
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
            return
        except Exception:
            pass

        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            return
        except Exception:
            pass

        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


setup_high_dpi()

# Проверка openpyxl
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.worksheet.page import PageMargins
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# =========================================================================
# КОНФИГУРАЦИЯ ВСЕХ 12 МЕСЯЦЕВ ЖУРНАЛА (с Сентября по Август)
# =========================================================================
MONTHS_CONFIG = [
    {"key": "september", "name": "Сентябрь", "p1": 4, "p2": 5, "tab_num": 5},
    {"key": "october",   "name": "Октябрь",  "p1": 6, "p2": 7, "tab_num": 6},
    {"key": "november",  "name": "Ноябрь",   "p1": 8, "p2": 9, "tab_num": 7},
    {"key": "december",  "name": "Декабрь",  "p1": 10, "p2": 11, "tab_num": 8},
    {"key": "january",   "name": "Январь",   "p1": 12, "p2": 13, "tab_num": 9},
    {"key": "february",  "name": "Февраль",  "p1": 14, "p2": 15, "tab_num": 10},
    {"key": "march",     "name": "Март",     "p1": 16, "p2": 17, "tab_num": 11},
    {"key": "april",     "name": "Апрель",   "p1": 18, "p2": 19, "tab_num": 12},
    {"key": "may",       "name": "Май",      "p1": 20, "p2": 21, "tab_num": 13},
    {"key": "june",      "name": "Июнь",     "p1": 22, "p2": 23, "tab_num": 14},
    {"key": "july",      "name": "Июль",     "p1": 24, "p2": 25, "tab_num": 15},
    {"key": "august",    "name": "Август",   "p1": 26, "p2": 27, "tab_num": 16},
]


def load_config() -> dict:
    """Загружает конфигурацию и историю из config.json."""
    default_config = {
        "github_repo": DEFAULT_GITHUB_REPO,
        "last_academic_year": "2024 / 2025",
        "history": ["2024 / 2025", "2023 / 2024", "2022 / 2023"],
        "teacher_name": "",
        "teacher_history": [],
        "org_name": "ГБУ ДО Республиканский детский образовательный технопарк",
        "org_history": ["ГБУ ДО Республиканский детский образовательный технопарк"],
        "title_academic_year": "2024 / 2025",
        "start_day_val": "",
        "start_month": "сентября",
        "start_year_val": "2024",
        "end_day_val": "",
        "end_month": "мая",
        "end_year_val": "2025",
        # Основные данные (стр. 3 журнала)
        "main_org_name": "ГБУ ДО Республиканский детский образовательный технопарк",
        "department": "",
        "association": "",
        "group_name": "",
        "study_year": "1-й год",
        "schedule": [
            {"day": "", "time": ""},
            {"day": "", "time": ""},
            {"day": "", "time": ""},
            {"day": "", "time": ""},
            {"day": "", "time": ""},
            {"day": "", "time": ""}
        ],
        "schedule_changes": [],
        "rukovoditel": "",
        "starosta": "",
        "accompanist": "",
        "accompanist_schedule": "",
        "accompanist_changes": "",
        # Сентябрь (стр. 4 и стр. 5 журнала)
        "september_students": [""] * 30,
        "september_dates": [""] * 15,
        "september_attendance": [[""] * 15 for _ in range(30)],
        "september_topics": [
            {"date": "", "content": "", "hours_teacher": "", "sign_teacher": "", "hours_acc": "", "sign_acc": ""}
            for _ in range(16)
        ],
        # Учёт массовых мероприятий с обучающимися (стр. 30 и стр. 31 журнала, по 26 строк на каждой странице)
        "mass_events_p1": [
            {"date": "", "content": "", "count": "", "location": "", "conducted_by": ""}
            for _ in range(26)
        ],
        "mass_events_p2": [
            {"date": "", "content": "", "count": "", "location": "", "conducted_by": ""}
            for _ in range(26)
        ],
        # Творческие достижения обучающихся (стр. 32 и стр. 33 журнала, по 26 строк на каждой странице)
        "creative_achievements_p1": [
            {"student": "", "event": ""}
            for _ in range(26)
        ],
        "creative_achievements_p2": [
            {"results": "", "works": ""}
            for _ in range(26)
        ],
        # Список обучающихся (стр. 34-39 журнала, 6 страниц по 10 строк)
        # Страницы 1, 3, 5: 6 столбцов (№, ФИО, Год рождения, Школа/класс, Район, Заключение врача)
        "students_list_p1": [
            {"student": "", "birth_year": "", "school_class": "", "district": "", "doctor_conclusion": ""}
            for _ in range(10)
        ],
        "students_list_p3": [
            {"student": "", "birth_year": "", "school_class": "", "district": "", "doctor_conclusion": ""}
            for _ in range(10)
        ],
        "students_list_p5": [
            {"student": "", "birth_year": "", "school_class": "", "district": "", "doctor_conclusion": ""}
            for _ in range(10)
        ],
        # Страницы 2, 4, 6: 5 столбцов (Адрес/телефон, Родители/телефон, Дата вступления, Когда/почему выбыл, Примечания)
        "students_list_p2": [
            {"address_phone": "", "parents_info": "", "join_date": "", "leave_info": "", "notes": ""}
            for _ in range(10)
        ],
        "students_list_p4": [
            {"address_phone": "", "parents_info": "", "join_date": "", "leave_info": "", "notes": ""}
            for _ in range(10)
        ],
        "students_list_p6": [
            {"address_phone": "", "parents_info": "", "join_date": "", "leave_info": "", "notes": ""}
            for _ in range(10)
        ],
        # Подсчёт отработанного времени (12 месяцев)
        "work_hours_report": [
            {
                "month_key": m["key"],
                "month_name": m["name"],
                "lessons_count": "",
                "hours_teacher": "",
                "hours_acc": "",
                "hours_total": "",
                "notes": ""
            }
            for m in MONTHS_CONFIG
        ]
    }
    if not os.path.exists(CONFIG_FILE):
        save_config(default_config)
        return default_config

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if not isinstance(data, dict):
                return default_config
            if "history" not in data or not isinstance(data["history"], list):
                data["history"] = default_config["history"]
            if "last_academic_year" not in data:
                data["last_academic_year"] = data["history"][0] if data["history"] else "2024 / 2025"
            if "teacher_name" not in data:
                data["teacher_name"] = default_config["teacher_name"]
            if "teacher_history" not in data or not isinstance(data["teacher_history"], list):
                data["teacher_history"] = default_config["teacher_history"]
            if "org_name" not in data:
                data["org_name"] = default_config["org_name"]
            if "org_history" not in data or not isinstance(data["org_history"], list):
                data["org_history"] = default_config["org_history"]
            if "title_academic_year" not in data:
                data["title_academic_year"] = default_config["title_academic_year"]
            if "start_day_val" not in data:
                data["start_day_val"] = default_config["start_day_val"]
            if "start_month" not in data:
                data["start_month"] = default_config["start_month"]
            if "start_year_val" not in data:
                data["start_year_val"] = default_config["start_year_val"]
            if "end_day_val" not in data:
                data["end_day_val"] = default_config["end_day_val"]
            if "end_month" not in data:
                data["end_month"] = default_config["end_month"]
            if "end_year_val" not in data:
                data["end_year_val"] = default_config["end_year_val"]

            # Проверка ключей основных данных
            if "main_org_name" not in data:
                data["main_org_name"] = data.get("org_name", default_config["main_org_name"])
            if "department" not in data:
                data["department"] = default_config["department"]
            if "association" not in data:
                data["association"] = default_config["association"]
            if "group_name" not in data:
                data["group_name"] = default_config["group_name"]
            if "study_year" not in data:
                data["study_year"] = default_config["study_year"]
            if "schedule" not in data or not isinstance(data["schedule"], list):
                data["schedule"] = default_config["schedule"]
            else:
                # Гарантируем наличие хотя бы 6 строк
                while len(data["schedule"]) < 6:
                    data["schedule"].append({"day": "", "time": ""})
            if "schedule_changes" not in data or not isinstance(data["schedule_changes"], list):
                data["schedule_changes"] = default_config["schedule_changes"]
            if "rukovoditel" not in data:
                data["rukovoditel"] = data.get("teacher_name", default_config["rukovoditel"])
            if "starosta" not in data:
                data["starosta"] = default_config["starosta"]
            if "accompanist" not in data:
                data["accompanist"] = default_config["accompanist"]
            if "accompanist_schedule" not in data:
                data["accompanist_schedule"] = default_config["accompanist_schedule"]
            if "accompanist_changes" not in data:
                data["accompanist_changes"] = default_config["accompanist_changes"]

            # Проверка ключей Сентября
            if "september_students" not in data or not isinstance(data["september_students"], list):
                data["september_students"] = default_config["september_students"]
            if "september_dates" not in data or not isinstance(data["september_dates"], list):
                data["september_dates"] = default_config["september_dates"]
            if "september_attendance" not in data or not isinstance(data["september_attendance"], list):
                data["september_attendance"] = default_config["september_attendance"]
            if "september_topics" not in data or not isinstance(data["september_topics"], list):
                data["september_topics"] = default_config["september_topics"]

            # Проверка ключей Учёта массовых мероприятий (стр. 30 и стр. 31)
            if "mass_events_p1" not in data or not isinstance(data["mass_events_p1"], list):
                data["mass_events_p1"] = default_config["mass_events_p1"]
            else:
                while len(data["mass_events_p1"]) < 26:
                    data["mass_events_p1"].append({"date": "", "content": "", "count": "", "location": "", "conducted_by": ""})

            if "mass_events_p2" not in data or not isinstance(data["mass_events_p2"], list):
                data["mass_events_p2"] = default_config["mass_events_p2"]
            else:
                while len(data["mass_events_p2"]) < 26:
                    data["mass_events_p2"].append({"date": "", "content": "", "count": "", "location": "", "conducted_by": ""})

            # Проверка ключей Творческих достижений (стр. 32 и стр. 33)
            if "creative_achievements_p1" not in data or not isinstance(data["creative_achievements_p1"], list):
                data["creative_achievements_p1"] = default_config["creative_achievements_p1"]
            else:
                while len(data["creative_achievements_p1"]) < 26:
                    data["creative_achievements_p1"].append({"student": "", "event": ""})

            if "creative_achievements_p2" not in data or not isinstance(data["creative_achievements_p2"], list):
                data["creative_achievements_p2"] = default_config["creative_achievements_p2"]
            else:
                while len(data["creative_achievements_p2"]) < 26:
                    data["creative_achievements_p2"].append({"results": "", "works": ""})

            # Проверка ключей Списка обучающихся (стр. 34-39 журнала, 6 страниц по 10 строк)
            for odd_k in ("students_list_p1", "students_list_p3", "students_list_p5"):
                if odd_k not in data or not isinstance(data[odd_k], list):
                    data[odd_k] = default_config[odd_k]
                else:
                    while len(data[odd_k]) < 10:
                        data[odd_k].append({"student": "", "birth_year": "", "school_class": "", "district": "", "doctor_conclusion": ""})

            for even_k in ("students_list_p2", "students_list_p4", "students_list_p6"):
                if even_k not in data or not isinstance(data[even_k], list):
                    data[even_k] = default_config[even_k]
                else:
                    while len(data[even_k]) < 10:
                        data[even_k].append({"address_phone": "", "parents_info": "", "join_date": "", "leave_info": "", "notes": ""})

            # Проверка данных отчёта об отработанном времени (12 месяцев)
            if "work_hours_report" not in data or not isinstance(data["work_hours_report"], list):
                data["work_hours_report"] = default_config["work_hours_report"]
            else:
                while len(data["work_hours_report"]) < 12:
                    idx = len(data["work_hours_report"])
                    m = MONTHS_CONFIG[idx]
                    data["work_hours_report"].append({
                        "month_key": m["key"],
                        "month_name": m["name"],
                        "lessons_count": "",
                        "hours_teacher": "",
                        "hours_acc": "",
                        "hours_total": "",
                        "notes": ""
                    })

            return data
    except Exception:
        return default_config


def save_config(config_data: dict) -> None:
    """Сохраняет конфигурацию и историю в config.json."""
    try:
        os.makedirs(APP_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Предупреждение: Не удалось сохранить {CONFIG_FILE}: {e}")


def populate_cover_sheet(ws, start_year: str, end_year: str):
    """Форматирует рабочий лист под Обложку журнала А4 (внешняя)."""
    ws.title = "Обложка"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Левое поле 30 мм (~1.18 дюйма)
    ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.78, bottom=0.78, header=0.2, footer=0.2)
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 75

    font_ministry = Font(name="Times New Roman", size=12, bold=True)
    font_title = Font(name="Times New Roman", size=28, bold=True)
    font_desc = Font(name="Times New Roman", size=13, bold=True)
    font_year = Font(name="Times New Roman", size=13, bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[5].height = 20
    ws['B5'] = "МИНИСТЕРСТВО ПРОСВЕЩЕНИЯ"
    ws['B5'].font = font_ministry
    ws['B5'].alignment = center

    ws.row_dimensions[6].height = 20
    ws['B6'] = "РОССИЙСКОЙ ФЕДЕРАЦИИ"
    ws['B6'].font = font_ministry
    ws['B6'].alignment = center

    ws.row_dimensions[13].height = 46
    ws['B13'] = "ЖУРНАЛ"
    ws['B13'].font = font_title
    ws['B13'].alignment = center

    ws.row_dimensions[17].height = 24
    ws['B17'] = "УЧЁТА РАБОТЫ ПЕДАГОГА"
    ws['B17'].font = font_desc
    ws['B17'].alignment = center

    ws.row_dimensions[19].height = 24
    ws['B19'] = "ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ"
    ws['B19'].font = font_desc
    ws['B19'].alignment = center

    ws.row_dimensions[21].height = 24
    ws['B21'] = "В ОБЪЕДИНЕНИИ (секции, клубе, кружке)"
    ws['B21'].font = font_desc
    ws['B21'].alignment = center

    ws.row_dimensions[25].height = 28
    st_val = start_year.strip() if start_year.strip() else "____"
    en_val = end_year.strip() if end_year.strip() else "____"
    ws['B25'] = f"на {st_val} / {en_val} учебный год"
    ws['B25'].font = font_year
    ws['B25'].alignment = center


def populate_empty_cover_back_sheet(ws):
    """Форматирует пустой рабочий лист под Оборот обложки А4 для правильной двусторонней печати."""
    ws.title = "Оборот обложки"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Четная страница: поле под переплет 30 мм СПРАВА
    ws.page_margins = PageMargins(left=0.59, right=1.18, top=0.78, bottom=0.78, header=0.0, footer=0.0)

    # Чтобы Excel гарантированно распечатал эту страницу как пустую при дуплексной печати книги:
    ws['A1'] = " "
    ws.print_area = 'A1:A1'


def populate_title_page_sheet(
    ws,
    org_name: str,
    academic_year: str,
    start_month: str,
    start_yr: str = "",
    end_month: str = "",
    end_yr: str = "",
    start_day: str = "",
    end_day: str = "",
    **kwargs
):
    """Форматирует рабочий лист под Титульный лист А4 (страница 1 журнала)."""
    ws.title = "Титульный лист (стр. 1)"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Страница 1 (нечетная): поле 30 мм СЛЕВА
    ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.59, bottom=0.59, header=0.2, footer=0.2)
    ws.column_dimensions['A'].width = 92

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_org = Font(name="Times New Roman", size=12, bold=True)
    font_sub_org = Font(name="Times New Roman", size=8, italic=True)
    font_fgos = Font(name="Times New Roman", size=11, italic=True)
    # В реальном журнале высота заглавных букв слова «ЖУРНАЛ» составляет 8.5 мм.
    # В Times New Roman Bold высота прописных знаков составляет ~68% от кегля.
    # Кегль 36 pt (12.7 мм) дает высоту заглавных букв: 36 * 0.68 * 0.3528 мм = 8.63 мм (~8.5 мм).
    font_title = Font(name="Times New Roman", size=36, bold=True)
    font_desc = Font(name="Times New Roman", size=13, bold=True)
    font_year = Font(name="Times New Roman", size=13, bold=True)
    font_dates = Font(name="Times New Roman", size=12, bold=True)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    # 1. Номер страницы "1" вверху справа
    ws['A1'] = "1"
    ws['A1'].font = font_page
    ws['A1'].alignment = right

    # 2. Название организации в самом верху (над линейками)
    org = org_name.strip() if org_name.strip() else ""
    line_rule = "______________________________________________________________________________"
    if org:
        ws['A3'] = org
        ws['A3'].font = font_org
        ws['A3'].alignment = center
        ws.row_dimensions[3].height = 24

        ws['A4'] = line_rule
        ws['A4'].alignment = center

        ws['A5'] = "(наименование образовательной организации)"
        ws['A5'].font = font_sub_org
        ws['A5'].alignment = center
    else:
        ws['A3'] = line_rule
        ws['A3'].alignment = center
        ws['A4'] = "(наименование образовательной организации)"
        ws['A4'].font = font_sub_org
        ws['A4'].alignment = center
        ws['A5'] = line_rule
        ws['A5'].alignment = center

    # 3. «Соответствует ФГОС» справа
    ws['A7'] = "Соответствует ФГОС"
    ws['A7'].font = font_fgos
    ws['A7'].alignment = right

    # 4. Заголовок «ЖУРНАЛ» (высота букв 8.5 мм при печати)
    ws['A11'] = "ЖУРНАЛ"
    ws['A11'].font = font_title
    ws['A11'].alignment = center
    ws.row_dimensions[11].height = 44

    # 5. Учёта работы педагога...
    ws['A13'] = "УЧЁТА РАБОТЫ ПЕДАГОГА"
    ws['A13'].font = font_desc
    ws['A13'].alignment = center
    ws.row_dimensions[13].height = 20

    ws['A15'] = "ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ"
    ws['A15'].font = font_desc
    ws['A15'].alignment = center
    ws.row_dimensions[15].height = 20

    ws['A17'] = "В ОБЪЕДИНЕНИИ (СЕКЦИИ, КЛУБЕ, КРУЖКЕ)"
    ws['A17'].font = font_desc
    ws['A17'].alignment = center
    ws.row_dimensions[17].height = 20

    # 6. Учебный год в середине
    y_str = academic_year.strip() if academic_year.strip() else "2024 / 2025"
    ws['A20'] = f"на   {y_str}   учебный год"
    ws['A20'].font = font_year
    ws['A20'].alignment = center
    ws.row_dimensions[20].height = 26

    # 7. Сроки внизу справа
    sm = start_month.strip() if start_month.strip() else "____________"
    sy = start_yr.strip() if start_yr.strip() else ""
    if len(sy) == 2:
        sy = f"20{sy}"
    elif not sy:
        sy = "20___"

    em = end_month.strip() if end_month.strip() else "____________"
    ey = end_yr.strip() if end_yr.strip() else ""
    if len(ey) == 2:
        ey = f"20{ey}"
    elif not ey:
        ey = "20___"

    sd = (start_day or kwargs.get("start_day_val", "")).strip()
    ed = (end_day or kwargs.get("end_day_val", "")).strip()

    sd_str = f"«  {sd}  »" if sd else "« ___ »"
    ed_str = f"«  {ed}  »" if ed else "« ___ »"

    ws['A26'] = f"Начат   {sd_str}   {sm}   {sy} г."
    ws['A26'].font = font_dates
    ws['A26'].alignment = right
    ws.row_dimensions[26].height = 22

    ws['A28'] = f"Окончен   {ed_str}   {em}   {ey} г."
    ws['A28'].font = font_dates
    ws['A28'].alignment = right
    ws.row_dimensions[28].height = 22


def populate_inside_cover_sheet(ws, teacher_name: str):
    """Форматирует рабочий лист под Оборот титульного листа А4 (страница 2 журнала)."""
    ws.title = "Оборот титульного (стр. 2)"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Четная страница (левая страница разворота): поле под корешок 30 мм СПРАВА
    ws.page_margins = PageMargins(left=0.59, right=1.18, top=0.78, bottom=0.78, header=0.2, footer=0.2)
    ws.column_dimensions['A'].width = 110

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_law_h = Font(name="Times New Roman", size=11, bold=True)
    font_law_p = Font(name="Times New Roman", size=9.5)
    font_instr_h = Font(name="Times New Roman", size=10.5, bold=True)
    font_rule = Font(name="Times New Roman", size=9)
    font_sign = Font(name="Times New Roman", size=10.5, bold=True)
    font_sub = Font(name="Times New Roman", size=8, italic=True)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    justify = Alignment(horizontal="justify", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws['A1'] = "2"
    ws['A1'].font = font_page

    ws['A3'] = "Федеральный закон Российской Федерации от 29 декабря 2012 г. № 273-ФЗ\n«Об образовании в Российской Федерации» (извлечения)"
    ws['A3'].font = font_law_h
    ws['A3'].alignment = center
    ws.row_dimensions[3].height = 28

    ws['A5'] = "«Дополнительное образование – вид образования, который направлен на всестороннее удовлетворение образовательных потребностей человека в интеллектуальном, духовно-нравственном, физическом и (или) профессиональном совершенствовании и не сопровождается повышением уровня образования»."
    ws['A5'].font = font_law_p
    ws['A5'].alignment = justify
    ws.row_dimensions[5].height = 36

    ws['A7'] = "«Дополнительное образование детей и взрослых направлено на формирование и развитие творческих способностей детей и взрослых, удовлетворение их индивидуальных потребностей в интеллектуальном, нравственном и физическом совершенствовании, формирование культуры здорового и безопасного образа жизни, укрепление здоровья, а также на организацию их свободного времени. Дополнительное образование детей обеспечивает их адаптацию к жизни в обществе, профессиональную ориентацию, а также выявление и поддержку детей, проявивших выдающиеся способности. Дополнительные общеобразовательные программы для детей должны учитывать возрастные и индивидуальные особенности детей»."
    ws['A7'].font = font_law_p
    ws['A7'].alignment = justify
    ws.row_dimensions[7].height = 72

    ws['A9'] = "УКАЗАНИЯ\nК ВЕДЕНИЮ ЖУРНАЛА УЧЁТА РАБОТЫ ПЕДАГОГА ДОПОЛНИТЕЛЬНОГО\nОБРАЗОВАНИЯ В ОБЪЕДИНЕНИИ (СЕКЦИИ, КЛУБЕ, КРУЖКЕ)"
    ws['A9'].font = font_instr_h
    ws['A9'].alignment = center
    ws.row_dimensions[9].height = 42

    rules = [
        "1. Журнал учёта работы педагога дополнительного образования в объединении (секции, клубе, кружке) является государственным учётным, финансовым документом. Его обязан вести каждый педагог дополнительного образования.",
        "2. Заведующий отделом, заместитель директора обязаны систематически контролировать правильность ведения журнала, внося соответствующие замечания, предложения по его ведению.",
        "3. Журнал рассчитан на учебный год и ведётся в каждом объединении.",
        "4. Записи в журнале должны вестись регулярно, чётко, аккуратно, чернилами синего цвета.",
        "5. На первой странице журнала педагог дополнительного образования записывает название объединения, расписание занятий, свои Ф. И. О. (полностью). Изменения расписания производятся в порядке, установленном в образовательной организации, и также с указанием числа (с какого произошло изменение) и названия документа (на основании которого изменение произошло).",
        "6. Для учёта работы объединения в журнале на каждый месяц отводится отдельная страница, где указываются Ф. И. обучающихся, дата проведения занятий, содержание занятий (тема), количество часов в соответствии с дополнительной образовательной программой и утверждённым расписанием занятий, ставится подпись педагога (при необходимости концертмейстера).",
        "7. Педагог объединения в дни занятий проверяет явку членов объединения и отмечает в журнале неявившихся буквой «н» (в графе, соответствующей дате занятий).",
        "8. Педагог объединения в конце первого месяца занятий составляет «Список обучающихся объединения», заполняет соответствующие графы. В случаях изменения состава объединения отмечает выбывших (дата, причина), вносит в журнал вновь принятых с указанием даты вступления в объединение.",
        "9. Педагог заполняет статистические данные о составе объединения на 01 октября, 01 января, 01 июня учебного года.",
        "10. Педагог ведёт учёт достижений обучающихся и заполняет соответствующие графы в журнале."
    ]

    cur_row = 11
    for r in rules:
        ws.cell(row=cur_row, column=1, value=r)
        ws.cell(row=cur_row, column=1).font = font_rule
        ws.cell(row=cur_row, column=1).alignment = justify
        cur_row += 1

    cur_row += 1
    t_name = teacher_name.strip() if teacher_name.strip() else "                                                           "
    ws.cell(row=cur_row, column=1, value=f"С указаниями по ведению журнала ознакомлен(-а)  {t_name}")
    ws.cell(row=cur_row, column=1).font = font_sign
    ws.cell(row=cur_row, column=1).alignment = left

    cur_row += 1
    ws.cell(row=cur_row, column=1, value="                                                                                Ф. И. О., подпись педагога")
    ws.cell(row=cur_row, column=1).font = font_sub
    ws.cell(row=cur_row, column=1).alignment = left


def generate_excel_cover(start_year: str, end_year: str, save_path: str) -> bool:
    """Генерация печатного листа обложки А4 под подшивку через openpyxl."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_cover_sheet(ws, start_year, end_year)
    wb.save(save_path)
    return True


def generate_excel_title_page(
    org_name: str,
    academic_year: str,
    start_month: str,
    start_yr: str = "",
    end_month: str = "",
    end_yr: str = "",
    save_path: str = None,
    start_day: str = "",
    end_day: str = "",
    **kwargs
) -> bool:
    """Генерация печатного листа 'Титульный лист' (стр. 1 журнала) А4 через openpyxl."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    actual_path = save_path or kwargs.get("filename")
    sy = start_yr or kwargs.get("start_year", "") or kwargs.get("start_year_val", "")
    ey = end_yr or kwargs.get("end_year", "") or kwargs.get("end_year_val", "")
    sd = start_day or kwargs.get("start_day", "") or kwargs.get("start_day_val", "")
    ed = end_day or kwargs.get("end_day", "") or kwargs.get("end_day_val", "")
    wb = Workbook()
    ws = wb.active
    populate_title_page_sheet(
        ws,
        org_name,
        academic_year,
        start_month,
        sy,
        end_month,
        ey,
        start_day=sd,
        end_day=ed
    )
    wb.save(actual_path)
    return True


def populate_main_data_sheet(
    ws,
    org_name: str = "",
    department: str = "",
    association: str = "",
    group_name: str = "",
    study_year: str = "",
    schedule: list = None,
    schedule_changes: list = None,
    rukovoditel: str = "",
    starosta: str = "",
    accompanist: str = "",
    accompanist_schedule: str = "",
    accompanist_changes: str = "",
    **kwargs
):
    """Форматирует рабочий лист под 'Основные данные' (страница 3 журнала) А4 через openpyxl."""
    ws.title = "Основные данные (стр. 3)"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Страница 3 (нечетная): отступ под скоросшиватель 30 мм СЛЕВА
    ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.59, bottom=0.59, header=0.2, footer=0.2)

    # 2 колонки: колонка A (38) и B (72) -> общая ширина 110
    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 72

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_bold_title = Font(name="Times New Roman", size=11, bold=True)
    font_label_bold = Font(name="Times New Roman", size=10, bold=True)
    font_label = Font(name="Times New Roman", size=10)
    font_value = Font(name="Times New Roman", size=10.5, bold=True)
    font_sub = Font(name="Times New Roman", size=8, italic=True)
    font_table_h = Font(name="Times New Roman", size=10, bold=True)
    font_table_cell = Font(name="Times New Roman", size=10)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    bottom_line = Border(bottom=thin_side)

    left_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    # 1. Номер страницы "3" в правом верхнем углу
    ws.row_dimensions[1].height = 18
    ws['B1'] = "3"
    ws['B1'].font = font_page
    ws['B1'].alignment = right_align

    ws.row_dimensions[2].height = 6

    # 2. Наименование организации
    ws.merge_cells("A3:B3")
    ws.row_dimensions[3].height = 20
    org_clean = org_name.strip() if org_name else ""
    ws['A3'] = f"Наименование организации   {org_clean}" if org_clean else "Наименование организации"
    ws['A3'].font = font_value if org_clean else font_label
    ws['A3'].alignment = left_align
    ws['A3'].border = bottom_line
    ws['B3'].border = bottom_line

    ws.merge_cells("A4:B4")
    ws.row_dimensions[4].height = 16
    ws['A4'].border = bottom_line
    ws['B4'].border = bottom_line

    ws.row_dimensions[5].height = 4

    # 3. Отдел
    ws.merge_cells("A6:B6")
    ws.row_dimensions[6].height = 20
    dept_clean = department.strip() if department else ""
    ws['A6'] = f"Отдел   {dept_clean}" if dept_clean else "Отдел"
    ws['A6'].font = font_value if dept_clean else font_label
    ws['A6'].alignment = left_align
    ws['A6'].border = bottom_line
    ws['B6'].border = bottom_line

    ws.row_dimensions[7].height = 4

    # 4. Объединение + (название)
    ws.merge_cells("A8:B8")
    ws.row_dimensions[8].height = 20
    assoc_clean = association.strip() if association else ""
    ws['A8'] = f"Объединение   {assoc_clean}" if assoc_clean else "Объединение"
    ws['A8'].font = font_value if assoc_clean else font_label
    ws['A8'].alignment = left_align
    ws['A8'].border = bottom_line
    ws['B8'].border = bottom_line

    ws.merge_cells("A9:B9")
    ws.row_dimensions[9].height = 12
    ws['A9'] = "(название)"
    ws['A9'].font = font_sub
    ws['A9'].alignment = Alignment(horizontal="center", vertical="top")

    # 5. Группа
    ws.merge_cells("A10:B10")
    ws.row_dimensions[10].height = 20
    grp_clean = group_name.strip() if group_name else ""
    ws['A10'] = f"Группа   {grp_clean}" if grp_clean else "Группа"
    ws['A10'].font = font_value if grp_clean else font_label
    ws['A10'].alignment = left_align
    ws['A10'].border = bottom_line
    ws['B10'].border = bottom_line

    ws.row_dimensions[11].height = 4

    # 6. Год обучения
    ws.merge_cells("A12:B12")
    ws.row_dimensions[12].height = 20
    yr_clean = study_year.strip() if study_year else ""
    ws['A12'] = f"Год обучения   {yr_clean}" if yr_clean else "Год обучения"
    ws['A12'].font = font_value if yr_clean else font_label
    ws['A12'].alignment = left_align
    ws['A12'].border = bottom_line
    ws['B12'].border = bottom_line

    ws.row_dimensions[13].height = 8

    # 7. РАСПИСАНИЕ ЗАНЯТИЙ
    ws.merge_cells("A14:B14")
    ws.row_dimensions[14].height = 22
    ws['A14'] = "РАСПИСАНИЕ ЗАНЯТИЙ"
    ws['A14'].font = font_bold_title
    ws['A14'].alignment = center_align

    # 8. Таблица расписания занятий (шапка)
    ws.row_dimensions[15].height = 20
    ws['A15'] = "Дни недели"
    ws['A15'].font = font_table_h
    ws['A15'].alignment = center_align
    ws['A15'].border = thin_border

    ws['B15'] = "Время (часы)"
    ws['B15'].font = font_table_h
    ws['B15'].alignment = center_align
    ws['B15'].border = thin_border

    # 9. 6 строк таблицы расписания занятий
    sch = schedule or []
    for idx in range(6):
        r = 16 + idx
        ws.row_dimensions[r].height = 20
        d_val = sch[idx].get("day", "") if idx < len(sch) and isinstance(sch[idx], dict) else ""
        t_val = sch[idx].get("time", "") if idx < len(sch) and isinstance(sch[idx], dict) else ""
        c1 = ws.cell(row=r, column=1, value=d_val)
        c2 = ws.cell(row=r, column=2, value=t_val)
        c1.font = font_table_cell
        c2.font = font_table_cell
        c1.alignment = center_align
        c2.alignment = center_align
        c1.border = thin_border
        c2.border = thin_border

    ws.row_dimensions[22].height = 6

    # 10. Изменения расписания занятий
    ws.merge_cells("A23:B23")
    ws.row_dimensions[23].height = 18
    ws['A23'] = "Изменения расписания занятий:"
    ws['A23'].font = font_label_bold
    ws['A23'].alignment = left_align

    # 11. Строки изменений (3-4 строки)
    chg = schedule_changes or []
    num_change_lines = max(3, len(chg))
    for idx in range(num_change_lines):
        r = 24 + idx
        ws.row_dimensions[r].height = 18
        if idx < len(chg) and isinstance(chg[idx], dict):
            d_val = chg[idx].get("day", "").strip()
            t_val = chg[idx].get("time", "").strip()
            val = f"{d_val}:  {t_val}" if d_val and t_val else (d_val or t_val)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            c = ws.cell(row=r, column=1, value=val)
            c.font = font_table_cell
            c.alignment = left_align
            c.border = bottom_line
            ws.cell(row=r, column=2).border = bottom_line
        else:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.cell(row=r, column=1).border = bottom_line
            ws.cell(row=r, column=2).border = bottom_line

    cur_row = 24 + num_change_lines
    ws.row_dimensions[cur_row].height = 6
    cur_row += 1

    # 12. РУКОВОДИТЕЛЬ + (фамилия, имя, отчество полностью)
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 20
    ruk_clean = rukovoditel.strip() if rukovoditel else ""
    ws.cell(row=cur_row, column=1, value=f"РУКОВОДИТЕЛЬ   {ruk_clean}" if ruk_clean else "РУКОВОДИТЕЛЬ").font = font_value if ruk_clean else font_label_bold
    ws.cell(row=cur_row, column=1).alignment = left_align
    ws.cell(row=cur_row, column=1).border = bottom_line
    ws.cell(row=cur_row, column=2).border = bottom_line
    cur_row += 1

    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 12
    c_sub_ruk = ws.cell(row=cur_row, column=1, value="(фамилия, имя, отчество полностью)")
    c_sub_ruk.font = font_sub
    c_sub_ruk.alignment = Alignment(horizontal="center", vertical="top")
    cur_row += 1

    # 13. СТАРОСТА
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 20
    star_clean = starosta.strip() if starosta else ""
    ws.cell(row=cur_row, column=1, value=f"СТАРОСТА   {star_clean}" if star_clean else "СТАРОСТА").font = font_value if star_clean else font_label_bold
    ws.cell(row=cur_row, column=1).alignment = left_align
    ws.cell(row=cur_row, column=1).border = bottom_line
    ws.cell(row=cur_row, column=2).border = bottom_line
    cur_row += 1

    ws.row_dimensions[cur_row].height = 4
    cur_row += 1

    # 14. АККОМПАНИАТОР (КОНЦЕРТМЕЙСТЕР)
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 20
    acc_clean = accompanist.strip() if accompanist else ""
    ws.cell(row=cur_row, column=1, value=f"АККОМПАНИАТОР (КОНЦЕРТМЕЙСТЕР)   {acc_clean}" if acc_clean else "АККОМПАНИАТОР (КОНЦЕРТМЕЙСТЕР)").font = font_value if acc_clean else font_label_bold
    ws.cell(row=cur_row, column=1).alignment = left_align
    ws.cell(row=cur_row, column=1).border = bottom_line
    ws.cell(row=cur_row, column=2).border = bottom_line
    cur_row += 1

    ws.row_dimensions[cur_row].height = 4
    cur_row += 1

    # 15. Расписание работы аккомпаниатора (концертмейстера)
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 18
    acc_s_clean = accompanist_schedule.strip() if accompanist_schedule else ""
    ws.cell(row=cur_row, column=1, value=f"Расписание работы аккомпаниатора (концертмейстера)   {acc_s_clean}" if acc_s_clean else "Расписание работы аккомпаниатора (концертмейстера)").font = font_label
    ws.cell(row=cur_row, column=1).alignment = left_align
    ws.cell(row=cur_row, column=1).border = bottom_line
    ws.cell(row=cur_row, column=2).border = bottom_line
    cur_row += 1

    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 16
    ws.cell(row=cur_row, column=1).border = bottom_line
    ws.cell(row=cur_row, column=2).border = bottom_line
    cur_row += 1

    ws.row_dimensions[cur_row].height = 4
    cur_row += 1

    # 16. Изменение расписания работы аккомпаниатора (концертмейстера)
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 18
    acc_c_clean = accompanist_changes.strip() if accompanist_changes else ""
    ws.cell(row=cur_row, column=1, value=f"Изменение расписания работы аккомпаниатора (концертмейстера)   {acc_c_clean}" if acc_c_clean else "Изменение расписания работы аккомпаниатора (концертмейстера)").font = font_label
    ws.cell(row=cur_row, column=1).alignment = left_align
    ws.cell(row=cur_row, column=1).border = bottom_line
    ws.cell(row=cur_row, column=2).border = bottom_line
    cur_row += 1

    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=2)
    ws.row_dimensions[cur_row].height = 16
    ws.cell(row=cur_row, column=1).border = bottom_line
    ws.cell(row=cur_row, column=2).border = bottom_line


def generate_excel_inside_cover(teacher_name: str, save_path: str) -> bool:
    """Генерация печатного листа 'Оборот титульного листа' (стр. 2 журнала) А4 через openpyxl."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_inside_cover_sheet(ws, teacher_name)
    wb.save(save_path)
    return True


def generate_excel_main_data(
    org_name: str = "",
    department: str = "",
    association: str = "",
    group_name: str = "",
    study_year: str = "",
    schedule: list = None,
    schedule_changes: list = None,
    rukovoditel: str = "",
    starosta: str = "",
    accompanist: str = "",
    accompanist_schedule: str = "",
    accompanist_changes: str = "",
    save_path: str = None,
    **kwargs
) -> bool:
    """Генерация печатного листа 'Основные данные' (стр. 3 журнала) А4 через openpyxl."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    actual_path = save_path or kwargs.get("filename")
    wb = Workbook()
    ws = wb.active
    populate_main_data_sheet(
        ws,
        org_name=org_name,
        department=department,
        association=association,
        group_name=group_name,
        study_year=study_year,
        schedule=schedule,
        schedule_changes=schedule_changes,
        rukovoditel=rukovoditel,
        starosta=starosta,
        accompanist=accompanist,
        accompanist_schedule=accompanist_schedule,
        accompanist_changes=accompanist_changes
    )
    wb.save(actual_path)
    return True


def populate_month_page1_sheet(
    ws,
    month_name: str = "Сентябрь",
    page_number: int = 4,
    students: list = None,
    dates: list = None,
    attendance: list = None
):
    """
    Универсальное форматирование листа '{month_name} - Учёт посещаемости и выполнения' (стр. {page_number} журнала) А4.
    30 обучающихся, 15 колонок для дат. Четная страница разворота (корешок 30 мм СПРАВА).
    """
    ws.title = f"{month_name} (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Четная страница (левая часть разворота): отступ 30 мм СПРАВА под корешок (1.18 дюйма), слева 10 мм (0.39 дюйма)
    ws.page_margins = PageMargins(left=0.39, right=1.18, top=0.39, bottom=0.39, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 4.5
    ws.column_dimensions['B'].width = 28.0
    date_cols = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    for c in date_cols:
        ws.column_dimensions[c].width = 3.5

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_title = Font(name="Times New Roman", size=11, bold=True)
    font_table_h = Font(name="Times New Roman", size=9.5, bold=True)
    font_num = Font(name="Times New Roman", size=9)
    font_name = Font(name="Times New Roman", size=9.5)
    font_date = Font(name="Times New Roman", size=8.5, bold=True)
    font_att = Font(name="Times New Roman", size=9)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    # 1. Номер страницы слева вверху и заголовок "УЧЁТ ПОСЕЩАЕМОСТИ И ВЫПОЛНЕНИЯ" справа
    ws.row_dimensions[1].height = 18
    ws['A1'] = str(page_number)
    ws['A1'].font = font_page
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("C1:Q1")
    ws['C1'] = "УЧЁТ ПОСЕЩАЕМОСТИ И ВЫПОЛНЕНИЯ"
    ws['C1'].font = font_header_title
    ws['C1'].alignment = right_align

    ws.row_dimensions[2].height = 5

    # 2. Шапка таблицы (строки 3, 4, 5)
    ws.merge_cells("A3:A5")
    ws['A3'] = "№\nп/п"
    ws['A3'].font = font_table_h
    ws['A3'].alignment = center_align

    ws.merge_cells("B3:B5")
    ws['B3'] = "Фамилия, имя обучающегося"
    ws['B3'].font = font_table_h
    ws['B3'].alignment = center_align

    spaced_month = " ".join(month_name)
    ws.merge_cells("C3:Q3")
    ws['C3'] = f"Месяц:  {spaced_month}"
    ws['C3'].font = font_table_h
    ws['C3'].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("C4:Q4")
    ws['C4'] = "Д а т а"
    ws['C4'].font = font_table_h
    ws['C4'].alignment = center_align

    dates_list = dates if dates is not None else []
    for idx, c in enumerate(date_cols):
        d_val = str(dates_list[idx]).strip() if idx < len(dates_list) and dates_list[idx] else ""
        cell = ws[f"{c}5"]
        cell.value = d_val
        cell.font = font_date
        cell.alignment = center_align

    all_cols = ['A', 'B'] + date_cols
    for r in range(3, 6):
        ws.row_dimensions[r].height = 16
        for c in all_cols:
            ws[f"{c}{r}"].border = thin_border

    # 3. 30 строк обучающихся (строки 6..35)
    students_list = students if students is not None else []
    att_matrix = attendance if attendance is not None else []

    for i in range(30):
        row_num = 6 + i
        ws.row_dimensions[row_num].height = 17.5

        cell_a = ws[f"A{row_num}"]
        cell_a.value = i + 1
        cell_a.font = font_num
        cell_a.alignment = center_align
        cell_a.border = thin_border

        name_val = str(students_list[i]).strip() if i < len(students_list) and students_list[i] else ""
        cell_b = ws[f"B{row_num}"]
        cell_b.value = name_val
        cell_b.font = font_name
        cell_b.alignment = left_align
        cell_b.border = thin_border

        row_att = att_matrix[i] if i < len(att_matrix) and isinstance(att_matrix[i], list) else []
        for c_idx, c in enumerate(date_cols):
            att_val = str(row_att[c_idx]).strip() if c_idx < len(row_att) and row_att[c_idx] else ""
            cell_d = ws[f"{c}{row_num}"]
            cell_d.value = att_val
            cell_d.font = font_att
            cell_d.alignment = center_align
            cell_d.border = thin_border


def populate_month_page2_sheet(
    ws,
    month_name: str = "Сентябрь",
    page_number: int = 5,
    topics: list = None
):
    """
    Универсальное форматирование листа '{month_name} - Выполнение ДОП' (стр. {page_number} журнала) А4.
    Таблица из 6 столбцов и 16 строк. Нечетная страница разворота (корешок 30 мм СЛЕВА).
    """
    ws.title = f"{month_name} (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Нечетная страница (правая часть разворота): отступ 30 мм СЛЕВА под корешок (1.18 дюйма), справа 10 мм (0.39 дюйма)
    ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.39, bottom=0.39, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 11.5
    ws.column_dimensions['B'].width = 34.5
    ws.column_dimensions['C'].width = 6.0
    ws.column_dimensions['D'].width = 10.0
    ws.column_dimensions['E'].width = 6.0
    ws.column_dimensions['F'].width = 17.0

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_title = Font(name="Times New Roman", size=11, bold=True)
    font_table_h = Font(name="Times New Roman", size=9, bold=True)
    font_table_h_acc = Font(name="Times New Roman", size=8, bold=True)
    font_cell = Font(name="Times New Roman", size=9)
    font_cell_bold = Font(name="Times New Roman", size=9.5, bold=True)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    # 1. Заголовок "ДОПОЛНИТЕЛЬНОЙ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ (ДОП)" слева и номер страницы справа
    ws.row_dimensions[1].height = 18
    ws.merge_cells("A1:E1")
    ws['A1'] = "ДОПОЛНИТЕЛЬНОЙ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ (ДОП)"
    ws['A1'].font = font_header_title
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    ws['F1'] = str(page_number)
    ws['F1'].font = font_page
    ws['F1'].alignment = right_align

    ws.row_dimensions[2].height = 5

    # 2. Шапка таблицы (строка 3)
    ws.row_dimensions[3].height = 46
    cols_header = [
        ('A', "Даты занятий\nобъединения"),
        ('B', "Содержание занятий согласно ДОП"),
        ('C', "Часы"),
        ('D', "Подпись\nпедагога"),
        ('E', "Часы"),
        ('F', "Подпись\nаккомпаниатора\n(концертмейстера)")
    ]
    for col_letter, title_text in cols_header:
        c_cell = ws[f"{col_letter}3"]
        c_cell.value = title_text
        c_cell.font = font_table_h_acc if col_letter == 'F' else font_table_h
        c_cell.alignment = center_align
        c_cell.border = thin_border

    # 3. 16 строк занятий (строки 4..19)
    topics_list = topics if topics is not None else []
    for i in range(16):
        row_num = 4 + i
        ws.row_dimensions[row_num].height = 33

        t_data = topics_list[i] if i < len(topics_list) and isinstance(topics_list[i], dict) else {}
        d_val = str(t_data.get("date", "")).strip()
        cnt_val = str(t_data.get("content", "")).strip()
        ht_val = str(t_data.get("hours_teacher", "")).strip()
        st_val = str(t_data.get("sign_teacher", "")).strip()
        ha_val = str(t_data.get("hours_acc", "")).strip()
        sa_val = str(t_data.get("sign_acc", "")).strip()

        c_a = ws[f"A{row_num}"]
        c_a.value = d_val
        c_a.font = font_cell
        c_a.alignment = center_align
        c_a.border = thin_border

        c_b = ws[f"B{row_num}"]
        c_b.value = cnt_val
        c_b.font = font_cell
        c_b.alignment = left_align
        c_b.border = thin_border

        c_c = ws[f"C{row_num}"]
        c_c.value = ht_val
        c_c.font = font_cell_bold if ht_val else font_cell
        c_c.alignment = center_align
        c_c.border = thin_border

        c_d = ws[f"D{row_num}"]
        c_d.value = st_val
        c_d.font = font_cell
        c_d.alignment = center_align
        c_d.border = thin_border

        c_e = ws[f"E{row_num}"]
        c_e.value = ha_val
        c_e.font = font_cell_bold if ha_val else font_cell
        c_e.alignment = center_align
        c_e.border = thin_border

        c_f = ws[f"F{row_num}"]
        c_f.value = sa_val
        c_f.font = font_cell
        c_f.alignment = center_align
        c_f.border = thin_border


def populate_september_page1_sheet(ws, students: list = None, dates: list = None, attendance: list = None):
    """Совместимость: форматирование листа Сентябрь (стр. 4)."""
    return populate_month_page1_sheet(ws, month_name="Сентябрь", page_number=4, students=students, dates=dates, attendance=attendance)


def populate_september_page2_sheet(ws, topics: list = None):
    """Совместимость: форматирование листа Сентябрь (стр. 5)."""
    return populate_month_page2_sheet(ws, month_name="Сентябрь", page_number=5, topics=topics)


def generate_excel_month_p1(save_path: str, month_name: str = None, page_number: int = None, month_key: str = None, students: list = None, dates: list = None, attendance: list = None) -> bool:
    """Генерация печатного листа месяца (стр. 1 - Посещаемость) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    if month_key:
        for m in MONTHS_CONFIG:
            if m["key"] == month_key:
                month_name = month_name or m["name"]
                page_number = page_number or m["p1"]
                break
    month_name = month_name or "Сентябрь"
    page_number = page_number or 4
    wb = Workbook()
    ws = wb.active
    populate_month_page1_sheet(ws, month_name=month_name, page_number=page_number, students=students, dates=dates, attendance=attendance)
    wb.save(save_path)
    return True


def generate_excel_month_p2(save_path: str, month_name: str = None, page_number: int = None, month_key: str = None, topics: list = None) -> bool:
    """Генерация печатного листа месяца (стр. 2 - Выполнение ДОП) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    if month_key:
        for m in MONTHS_CONFIG:
            if m["key"] == month_key:
                month_name = month_name or m["name"]
                page_number = page_number or m["p2"]
                break
    month_name = month_name or "Сентябрь"
    page_number = page_number or 5
    wb = Workbook()
    ws = wb.active
    populate_month_page2_sheet(ws, month_name=month_name, page_number=page_number, topics=topics)
    wb.save(save_path)
    return True


def generate_excel_month_spread(
    save_path: str,
    month_name: str = None,
    p1_no: int = None,
    p2_no: int = None,
    month_key: str = None,
    students: list = None,
    dates: list = None,
    attendance: list = None,
    topics: list = None
) -> bool:
    """Генерация полного разворота месяца (2 страницы журнала) в один файл Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    if month_key:
        for m in MONTHS_CONFIG:
            if m["key"] == month_key:
                month_name = month_name or m["name"]
                p1_no = p1_no or m["p1"]
                p2_no = p2_no or m["p2"]
                break
    month_name = month_name or "Сентябрь"
    p1_no = p1_no or 4
    p2_no = p2_no or 5
    wb = Workbook()
    ws1 = wb.active
    populate_month_page1_sheet(ws1, month_name=month_name, page_number=p1_no, students=students, dates=dates, attendance=attendance)
    ws2 = wb.create_sheet(title=f"{month_name} (стр. {p2_no})")
    populate_month_page2_sheet(ws2, month_name=month_name, page_number=p2_no, topics=topics)
    wb.save(save_path)
    return True


def generate_excel_september_p1(save_path: str, students: list = None, dates: list = None, attendance: list = None) -> bool:
    """Совместимость: генерация листа Сентябрь (стр. 4)."""
    return generate_excel_month_p1(save_path, "Сентябрь", 4, students=students, dates=dates, attendance=attendance)


def generate_excel_september_p2(save_path: str, topics: list = None) -> bool:
    """Совместимость: генерация листа Сентябрь (стр. 5)."""
    return generate_excel_month_p2(save_path, "Сентябрь", 5, topics=topics)


def generate_excel_september_spread(
    save_path: str,
    students: list = None,
    dates: list = None,
    attendance: list = None,
    topics: list = None
) -> bool:
    """Совместимость: генерация разворота Сентябрь (стр. 4 и 5)."""
    return generate_excel_month_spread(save_path, "Сентябрь", 4, 5, students=students, dates=dates, attendance=attendance, topics=topics)


def populate_mass_events_sheet(
    ws,
    page_number: int = 30,
    events: list = None
):
    """
    Форматирование листа 'Учёт массовых мероприятий с обучающимися' (стр. 30 или 31 журнала) А4.
    Таблица из 5 столбцов и 26 строк.
    Стр. 30: четная страница разворота (корешок 30 мм СПРАВА, номер страницы 30 СЛЕВА).
    Стр. 31: нечетная страница разворота (корешок 30 мм СЛЕВА, номер страницы 31 СПРАВА).
    """
    ws.title = f"Мероприятия (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    if page_number % 2 == 0:
        # Четная страница (левая полоса разворота, стр. 30):
        # отступ 30 мм СПРАВА под корешок (1.18 дюйма), слева 10 мм (0.39 дюйма)
        ws.page_margins = PageMargins(left=0.39, right=1.18, top=0.39, bottom=0.39, header=0.15, footer=0.15)
    else:
        # Нечетная страница (правая полоса разворота, стр. 31):
        # отступ 30 мм СЛЕВА под корешок (1.18 дюйма), справа 10 мм (0.39 дюйма)
        ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.39, bottom=0.39, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 11.0
    ws.column_dimensions['B'].width = 44.0
    ws.column_dimensions['C'].width = 12.0
    ws.column_dimensions['D'].width = 18.0
    ws.column_dimensions['E'].width = 18.0

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_title = Font(name="Times New Roman", size=11, bold=True)
    font_table_h = Font(name="Times New Roman", size=9, bold=True)
    font_cell = Font(name="Times New Roman", size=9)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 1. Заголовок и номер страницы (строка 1)
    ws.row_dimensions[1].height = 20
    if page_number % 2 == 0:
        # Четная страница: номер страницы слева (A1), заголовок по центру (B1:E1)
        ws['A1'] = str(page_number)
        ws['A1'].font = font_page
        ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells("B1:E1")
        ws['B1'] = "УЧЁТ МАССОВЫХ МЕРОПРИЯТИЙ С ОБУЧАЮЩИМИСЯ"
        ws['B1'].font = font_header_title
        ws['B1'].alignment = center_align
    else:
        # Нечетная страница: заголовок по центру (A1:D1), номер страницы справа (E1)
        ws.merge_cells("A1:D1")
        ws['A1'] = "УЧЁТ МАССОВЫХ МЕРОПРИЯТИЙ С ОБУЧАЮЩИМИСЯ"
        ws['A1'].font = font_header_title
        ws['A1'].alignment = center_align

        ws['E1'] = str(page_number)
        ws['E1'].font = font_page
        ws['E1'].alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[2].height = 6

    # 2. Шапка таблицы (строка 3)
    ws.row_dimensions[3].height = 42
    cols_header = [
        ('A', "Дата"),
        ('B', "Название и краткое содержание\nпроведенного мероприятия"),
        ('C', "Количе-\nство уча-\nстников"),
        ('D', "Место\nпроведения\nмероприятия"),
        ('E', "Кто проводил")
    ]
    for col_letter, title_text in cols_header:
        c_cell = ws[f"{col_letter}3"]
        c_cell.value = title_text
        c_cell.font = font_table_h
        c_cell.alignment = center_align
        c_cell.border = thin_border

    # 3. 26 строк данных (строки 4..29)
    events_list = events if events is not None else []
    for i in range(26):
        row_num = 4 + i
        ws.row_dimensions[row_num].height = 22.5

        item_data = events_list[i] if i < len(events_list) and isinstance(events_list[i], dict) else {}
        d_val = str(item_data.get("date", "")).strip()
        cnt_val = str(item_data.get("content", "")).strip()
        count_val = str(item_data.get("count", "")).strip()
        loc_val = str(item_data.get("location", "")).strip()
        who_val = str(item_data.get("conducted_by", "")).strip()

        # A: Дата
        ca = ws[f"A{row_num}"]
        ca.value = d_val
        ca.font = font_cell
        ca.alignment = center_align
        ca.border = thin_border

        # B: Название и содержание
        cb = ws[f"B{row_num}"]
        cb.value = cnt_val
        cb.font = font_cell
        cb.alignment = left_align
        cb.border = thin_border

        # C: Количество участников
        cc = ws[f"C{row_num}"]
        cc.value = count_val
        cc.font = font_cell
        cc.alignment = center_align
        cc.border = thin_border

        # D: Место проведения
        cd = ws[f"D{row_num}"]
        cd.value = loc_val
        cd.font = font_cell
        cd.alignment = left_align
        cd.border = thin_border

        # E: Кто проводил
        ce = ws[f"E{row_num}"]
        ce.value = who_val
        ce.font = font_cell
        ce.alignment = left_align
        ce.border = thin_border


def generate_excel_mass_events_p1(save_path: str, events: list = None) -> bool:
    """Генерация печатного листа 'Учёт массовых мероприятий' (стр. 30) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_mass_events_sheet(ws, page_number=30, events=events)
    wb.save(save_path)
    return True


def generate_excel_mass_events_p2(save_path: str, events: list = None) -> bool:
    """Генерация печатного листа 'Учёт массовых мероприятий' (стр. 31) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_mass_events_sheet(ws, page_number=31, events=events)
    wb.save(save_path)
    return True


def generate_excel_mass_events_spread(save_path: str, events_p1: list = None, events_p2: list = None) -> bool:
    """Генерация разворота 'Учёт массовых мероприятий' (стр. 30 и 31) в один файл Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws1 = wb.active
    populate_mass_events_sheet(ws1, page_number=30, events=events_p1)
    ws2 = wb.create_sheet(title="Мероприятия (стр. 31)")
    populate_mass_events_sheet(ws2, page_number=31, events=events_p2)
    wb.save(save_path)
    return True


def populate_creative_achievements_p1_sheet(
    ws,
    page_number: int = 32,
    items: list = None
):
    """
    Форматирование листа 'Творческие достижения' (стр. 32 журнала) А4.
    Таблица из 3 столбцов и 26 строк:
    1) № п/п
    2) Фамилия, имя обучающегося
    3) В каких соревнованиях, смотрах, спектаклях и др. мероприятиях участвовал
    Четная страница разворота (корешок 30 мм СПРАВА, номер страницы 32 СЛЕВА).
    """
    ws.title = f"Достижения (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Четная страница (левая полоса разворота, стр. 32):
    # отступ 30 мм СПРАВА под корешок (1.18 дюйма), слева 10 мм (0.39 дюйма)
    ws.page_margins = PageMargins(left=0.39, right=1.18, top=0.39, bottom=0.39, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 6.0
    ws.column_dimensions['B'].width = 32.0
    ws.column_dimensions['C'].width = 62.0

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_title = Font(name="Times New Roman", size=11, bold=True)
    font_table_h = Font(name="Times New Roman", size=9, bold=True)
    font_cell = Font(name="Times New Roman", size=9)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 1. Заголовок и номер страницы (строка 1)
    ws.row_dimensions[1].height = 20
    ws['A1'] = str(page_number)
    ws['A1'].font = font_page
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    ws['C1'] = "ТВОРЧЕСКИЕ ДОСТИЖЕНИЯ"
    ws['C1'].font = font_header_title
    ws['C1'].alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[2].height = 6

    # 2. Шапка таблицы (строка 3)
    ws.row_dimensions[3].height = 42
    cols_header = [
        ('A', "№\nп/п"),
        ('B', "Фамилия, имя обучающегося"),
        ('C', "В каких соревнованиях, смотрах,\nспектаклях и др. мероприятиях\nучаствовал")
    ]
    for col_letter, title_text in cols_header:
        c_cell = ws[f"{col_letter}3"]
        c_cell.value = title_text
        c_cell.font = font_table_h
        c_cell.alignment = center_align
        c_cell.border = thin_border

    # 3. 26 строк данных (строки 4..29)
    items_list = items if items is not None else []
    for i in range(26):
        row_num = 4 + i
        ws.row_dimensions[row_num].height = 22.5

        item_data = items_list[i] if i < len(items_list) and isinstance(items_list[i], dict) else {}
        student_val = str(item_data.get("student", "")).strip()
        event_val = str(item_data.get("event", "")).strip()

        # A: № п/п
        ca = ws[f"A{row_num}"]
        ca.value = i + 1
        ca.font = font_cell
        ca.alignment = center_align
        ca.border = thin_border

        # B: Фамилия, имя обучающегося
        cb = ws[f"B{row_num}"]
        cb.value = student_val
        cb.font = font_cell
        cb.alignment = left_align
        cb.border = thin_border

        # C: В каких соревнованиях... участвовал
        cc = ws[f"C{row_num}"]
        cc.value = event_val
        cc.font = font_cell
        cc.alignment = left_align
        cc.border = thin_border


def populate_creative_achievements_p2_sheet(
    ws,
    page_number: int = 33,
    items: list = None
):
    """
    Форматирование листа 'Творческие достижения - Обучающихся' (стр. 33 журнала) А4.
    Таблица из 2 столбцов и 26 строк:
    1) Результаты (полученное звание, разряд и другие результаты)
    2) Работы, выполненные объединением по заказам или инициативно
    Нечетная страница разворота (корешок 30 мм СЛЕВА, номер страницы 33 СПРАВА).
    """
    ws.title = f"Достижения (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Нечетная страница (правая полоса разворота, стр. 33):
    # отступ 30 мм СЛЕВА под корешок (1.18 дюйма), справа 10 мм (0.39 дюйма)
    ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.39, bottom=0.39, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 48.0
    ws.column_dimensions['B'].width = 52.0

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_title = Font(name="Times New Roman", size=11, bold=True)
    font_table_h = Font(name="Times New Roman", size=9, bold=True)
    font_cell = Font(name="Times New Roman", size=9)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 1. Заголовок и номер страницы (строка 1)
    ws.row_dimensions[1].height = 20
    ws['A1'] = "ОБУЧАЮЩИХСЯ"
    ws['A1'].font = font_header_title
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    ws['B1'] = str(page_number)
    ws['B1'].font = font_page
    ws['B1'].alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[2].height = 6

    # 2. Шапка таблицы (строка 3)
    ws.row_dimensions[3].height = 42
    cols_header = [
        ('A', "Результаты (полученное звание,\nразряд и другие результаты)"),
        ('B', "Работы, выполненные объединением\nпо заказам или инициативно")
    ]
    for col_letter, title_text in cols_header:
        c_cell = ws[f"{col_letter}3"]
        c_cell.value = title_text
        c_cell.font = font_table_h
        c_cell.alignment = center_align
        c_cell.border = thin_border

    # 3. 26 строк данных (строки 4..29)
    items_list = items if items is not None else []
    for i in range(26):
        row_num = 4 + i
        ws.row_dimensions[row_num].height = 22.5

        item_data = items_list[i] if i < len(items_list) and isinstance(items_list[i], dict) else {}
        results_val = str(item_data.get("results", "")).strip()
        works_val = str(item_data.get("works", "")).strip()

        # A: Результаты
        ca = ws[f"A{row_num}"]
        ca.value = results_val
        ca.font = font_cell
        ca.alignment = left_align
        ca.border = thin_border

        # B: Работы...
        cb = ws[f"B{row_num}"]
        cb.value = works_val
        cb.font = font_cell
        cb.alignment = left_align
        cb.border = thin_border


def generate_excel_creative_achievements_p1(save_path: str, items: list = None) -> bool:
    """Генерация печатного листа 'Творческие достижения' (стр. 32) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_creative_achievements_p1_sheet(ws, page_number=32, items=items)
    wb.save(save_path)
    return True


def generate_excel_creative_achievements_p2(save_path: str, items: list = None) -> bool:
    """Генерация печатного листа 'Творческие достижения - Обучающихся' (стр. 33) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_creative_achievements_p2_sheet(ws, page_number=33, items=items)
    wb.save(save_path)
    return True


def generate_excel_creative_achievements_spread(save_path: str, items_p1: list = None, items_p2: list = None) -> bool:
    """Генерация разворота 'Творческие достижения' (стр. 32 и 33) в один файл Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws1 = wb.active
    populate_creative_achievements_p1_sheet(ws1, page_number=32, items=items_p1)
    ws2 = wb.create_sheet(title="Достижения (стр. 33)")
    populate_creative_achievements_p2_sheet(ws2, page_number=33, items=items_p2)
    wb.save(save_path)
    return True


def populate_students_list_odd_sheet(
    ws,
    page_number: int = 34,
    start_student_no: int = 1,
    items: list = None
):
    """
    Форматирование листа 'Список обучающихся в объединении' (стр. 34, 36, 38 журнала) А4.
    Таблица из 6 столбцов и 10 строк:
    1) № п/п
    2) Фамилия, имя обучающегося
    3) Год рождения
    4) Школа, класс
    5) Район
    6) Заключение врача о допуске к занятиям
    Четная страница журнала (левая сторона разворота: корешок 30 мм СПРАВА, номер страницы слева).
    """
    ws.title = f"Обучающиеся (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Четная страница журнала (левая полоса разворота):
    # отступ 30 мм СПРАВА под корешок (1.18 дюйма), слева 10 мм (0.39 дюйма)
    ws.page_margins = PageMargins(left=0.39, right=1.18, top=0.39, bottom=0.39, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 5.5
    ws.column_dimensions['B'].width = 26.0
    ws.column_dimensions['C'].width = 10.0
    ws.column_dimensions['D'].width = 15.0
    ws.column_dimensions['E'].width = 16.0
    ws.column_dimensions['F'].width = 20.0

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_title = Font(name="Times New Roman", size=10, bold=True)
    font_table_h = Font(name="Times New Roman", size=9, bold=True)
    font_cell = Font(name="Times New Roman", size=9)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 1. Номер страницы (слева) и Заголовок (справа) (строка 1)
    ws.row_dimensions[1].height = 22
    ws['A1'] = str(page_number)
    ws['A1'].font = font_page
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    ws['F1'] = "СПИСОК ОБУЧАЮЩИХСЯ В ОБЪЕДИНЕНИИ;"
    ws['F1'].font = font_header_title
    ws['F1'].alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[2].height = 6

    # 2. Шапка таблицы (строка 3)
    ws.row_dimensions[3].height = 44
    cols_header = [
        ('A', "№\nп/п"),
        ('B', "Фамилия, имя обучающегося"),
        ('C', "Год\nрождения"),
        ('D', "Школа, класс"),
        ('E', "Район"),
        ('F', "Заключение врача\nо допуске к занятиям")
    ]
    for col_letter, title_text in cols_header:
        c_cell = ws[f"{col_letter}3"]
        c_cell.value = title_text
        c_cell.font = font_table_h
        c_cell.alignment = center_align
        c_cell.border = thin_border

    # 3. 10 строк данных (строки 4..13)
    items_list = items if items is not None else []
    for i in range(10):
        row_num = 4 + i
        ws.row_dimensions[row_num].height = 52.0

        item_data = items_list[i] if i < len(items_list) and isinstance(items_list[i], dict) else {}
        student_val = str(item_data.get("student", "")).strip()
        birth_year_val = str(item_data.get("birth_year", "")).strip()
        school_class_val = str(item_data.get("school_class", "")).strip()
        district_val = str(item_data.get("district", "")).strip()
        doctor_val = str(item_data.get("doctor_conclusion", "")).strip()

        # A: № п/п
        ca = ws[f"A{row_num}"]
        ca.value = start_student_no + i
        ca.font = font_cell
        ca.alignment = center_align
        ca.border = thin_border

        # B: Фамилия, имя обучающегося
        cb = ws[f"B{row_num}"]
        cb.value = student_val
        cb.font = font_cell
        cb.alignment = left_align
        cb.border = thin_border

        # C: Год рождения
        cc = ws[f"C{row_num}"]
        cc.value = birth_year_val
        cc.font = font_cell
        cc.alignment = center_align
        cc.border = thin_border

        # D: Школа, класс
        cd = ws[f"D{row_num}"]
        cd.value = school_class_val
        cd.font = font_cell
        cd.alignment = center_align
        cd.border = thin_border

        # E: Район
        ce = ws[f"E{row_num}"]
        ce.value = district_val
        ce.font = font_cell
        ce.alignment = center_align
        ce.border = thin_border

        # F: Заключение врача о допуске к занятиям
        cf = ws[f"F{row_num}"]
        cf.value = doctor_val
        cf.font = font_cell
        cf.alignment = center_align
        cf.border = thin_border


def populate_students_list_even_sheet(
    ws,
    page_number: int = 35,
    items: list = None
):
    """
    Форматирование листа 'Сведения о родителях' (стр. 35, 37, 39 журнала) А4.
    Таблица из 5 столбцов и 10 строк:
    1) Домашний адрес, телефон
    2) Фамилия, имя, отчество родителей, телефон
    3) Дата вступления в объединение
    4) Когда и почему выбыл
    5) Примечания
    Нечетная страница журнала (правая сторона разворота: корешок 30 мм СЛЕВА, номер страницы справа).
    """
    ws.title = f"Обучающиеся (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Нечетная страница журнала (правая полоса разворота):
    # отступ 30 мм СЛЕВА под корешок (1.18 дюйма), справа 10 мм (0.39 дюйма)
    ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.39, bottom=0.39, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 22.0
    ws.column_dimensions['B'].width = 25.0
    ws.column_dimensions['C'].width = 14.0
    ws.column_dimensions['D'].width = 15.0
    ws.column_dimensions['E'].width = 16.5

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_title = Font(name="Times New Roman", size=11, bold=True)
    font_table_h = Font(name="Times New Roman", size=9, bold=True)
    font_cell = Font(name="Times New Roman", size=9)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 1. Заголовок (слева) и Номер страницы (справа) (строка 1)
    ws.row_dimensions[1].height = 22
    ws['A1'] = "СВЕДЕНИЯ О РОДИТЕЛЯХ"
    ws['A1'].font = font_header_title
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    ws['E1'] = str(page_number)
    ws['E1'].font = font_page
    ws['E1'].alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[2].height = 6

    # 2. Шапка таблицы (строка 3)
    ws.row_dimensions[3].height = 44
    cols_header = [
        ('A', "Домашний адрес,\nтелефон"),
        ('B', "Фамилия, имя, отчество\nродителей, телефон"),
        ('C', "Дата\nвступления\nв объеди-\nнение"),
        ('D', "Когда\nи почему\nвыбыл"),
        ('E', "Примечания")
    ]
    for col_letter, title_text in cols_header:
        c_cell = ws[f"{col_letter}3"]
        c_cell.value = title_text
        c_cell.font = font_table_h
        c_cell.alignment = center_align
        c_cell.border = thin_border

    # 3. 10 строк данных (строки 4..13)
    items_list = items if items is not None else []
    for i in range(10):
        row_num = 4 + i
        ws.row_dimensions[row_num].height = 52.0

        item_data = items_list[i] if i < len(items_list) and isinstance(items_list[i], dict) else {}
        addr_val = str(item_data.get("address_phone", "")).strip()
        parent_val = str(item_data.get("parents_info", "")).strip()
        join_val = str(item_data.get("join_date", "")).strip()
        leave_val = str(item_data.get("leave_info", "")).strip()
        notes_val = str(item_data.get("notes", "")).strip()

        # A: Домашний адрес, телефон
        ca = ws[f"A{row_num}"]
        ca.value = addr_val
        ca.font = font_cell
        ca.alignment = left_align
        ca.border = thin_border

        # B: Фамилия, имя, отчество родителей, телефон
        cb = ws[f"B{row_num}"]
        cb.value = parent_val
        cb.font = font_cell
        cb.alignment = left_align
        cb.border = thin_border

        # C: Дата вступления в объединение
        cc = ws[f"C{row_num}"]
        cc.value = join_val
        cc.font = font_cell
        cc.alignment = center_align
        cc.border = thin_border

        # D: Когда и почему выбыл
        cd = ws[f"D{row_num}"]
        cd.value = leave_val
        cd.font = font_cell
        cd.alignment = left_align
        cd.border = thin_border

        # E: Примечания
        ce = ws[f"E{row_num}"]
        ce.value = notes_val
        ce.font = font_cell
        ce.alignment = left_align
        ce.border = thin_border


def generate_excel_students_list_page(save_path: str, page_num: int = 1, items: list = None) -> bool:
    """Генерация отдельного листа страницы 'Список обучающихся' (стр. 1..6 вкладки, в журнале стр. 34..39) в Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    journal_page_no = 33 + page_num
    if page_num in (1, 3, 5):
        start_no = 1 if page_num == 1 else (11 if page_num == 3 else 21)
        populate_students_list_odd_sheet(ws, page_number=journal_page_no, start_student_no=start_no, items=items)
    else:
        populate_students_list_even_sheet(ws, page_number=journal_page_no, items=items)
    wb.save(save_path)
    return True


def generate_excel_students_list_spread(save_path: str, spread_idx: int = 1, items_odd: list = None, items_even: list = None) -> bool:
    """Генерация разворота 'Список обучающихся' (разворот 1: стр. 34-35, разворот 2: стр. 36-37, разворот 3: стр. 38-39)."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws1 = wb.active
    if spread_idx == 1:
        p_odd = 34
        p_even = 35
        start_no = 1
    elif spread_idx == 2:
        p_odd = 36
        p_even = 37
        start_no = 11
    else:
        p_odd = 38
        p_even = 39
        start_no = 21

    populate_students_list_odd_sheet(ws1, page_number=p_odd, start_student_no=start_no, items=items_odd)
    ws2 = wb.create_sheet(title=f"Обучающиеся (стр. {p_even})")
    populate_students_list_even_sheet(ws2, page_number=p_even, items=items_even)
    wb.save(save_path)
    return True


def generate_excel_students_list_all(save_path: str, all_pages: dict = None) -> bool:
    """Генерация всех 6 страниц 'Список обучающихся' (стр. 34..39 журнала) в один файл Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    pages = all_pages or {}

    ws1 = wb.active
    populate_students_list_odd_sheet(ws1, page_number=34, start_student_no=1, items=pages.get("students_list_p1"))

    ws2 = wb.create_sheet(title="Обучающиеся (стр. 35)")
    populate_students_list_even_sheet(ws2, page_number=35, items=pages.get("students_list_p2"))

    ws3 = wb.create_sheet(title="Обучающиеся (стр. 36)")
    populate_students_list_odd_sheet(ws3, page_number=36, start_student_no=11, items=pages.get("students_list_p3"))

    ws4 = wb.create_sheet(title="Обучающиеся (стр. 37)")
    populate_students_list_even_sheet(ws4, page_number=37, items=pages.get("students_list_p4"))

    ws5 = wb.create_sheet(title="Обучающиеся (стр. 38)")
    populate_students_list_odd_sheet(ws5, page_number=38, start_student_no=21, items=pages.get("students_list_p5"))

    ws6 = wb.create_sheet(title="Обучающиеся (стр. 39)")
    populate_students_list_even_sheet(ws6, page_number=39, items=pages.get("students_list_p6"))

    wb.save(save_path)
    return True


def populate_safety_briefing_sheet(
    ws,
    page_number: int = 38,
    items: list = None
):
    """
    Форматирование листа 'Список обучающихся в объединении, прошедших инструктаж по технике безопасности' А4.
    Таблица из 5 столбцов и 28 строк:
    1) № п/п
    2) Фамилия, имя обучающегося
    3) Дата проведения инструктажа
    4) Краткое содержание инструктажа
    5) Подпись проводившего инструктаж (разборчиво)
    Четная страница (стр. 38): корешок 30 мм СПРАВА, номер страницы 38 СЛЕВА.
    Нечетная страница (стр. 39): корешок 30 мм СЛЕВА, номер страницы 39 СПРАВА.
    """
    ws.title = f"Инструктаж по ТБ (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    if page_number % 2 == 0:
        # Четная страница (левая полоса разворота, стр. 38):
        # отступ 30 мм СПРАВА под корешок (1.18 дюйма), слева 10 мм (0.39 дюйма)
        ws.page_margins = PageMargins(left=0.39, right=1.18, top=0.35, bottom=0.35, header=0.15, footer=0.15)
    else:
        # Нечетная страница (правая полоса разворота, стр. 39):
        # отступ 30 мм СЛЕВА под корешок (1.18 дюйма), справа 10 мм (0.39 дюйма)
        ws.page_margins = PageMargins(left=1.18, right=0.39, top=0.35, bottom=0.35, header=0.15, footer=0.15)

    ws.column_dimensions['A'].width = 5.5
    ws.column_dimensions['B'].width = 28.0
    ws.column_dimensions['C'].width = 14.5
    ws.column_dimensions['D'].width = 31.0
    ws.column_dimensions['E'].width = 19.5

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_header_1 = Font(name="Times New Roman", size=11, bold=True)
    font_header_2 = Font(name="Times New Roman", size=9.5, bold=True)
    font_table_h = Font(name="Times New Roman", size=8.5, bold=True)
    font_cell = Font(name="Times New Roman", size=9)
    font_cell_small = Font(name="Times New Roman", size=8.5)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    top_left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # 1. Заголовок и номер страницы (строки 1 и 2)
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 4

    if page_number % 2 == 0:
        # Четная страница: номер страницы слева вверху (A1)
        ws['A1'] = str(page_number)
        ws['A1'].font = font_page
        ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells("B1:E1")
        ws['B1'] = "СПИСОК ОБУЧАЮЩИХСЯ"
        ws['B1'].font = font_header_1
        ws['B1'].alignment = center_align
    else:
        # Нечетная страница: номер страницы справа вверху (E1)
        ws.merge_cells("A1:D1")
        ws['A1'] = "СПИСОК ОБУЧАЮЩИХСЯ"
        ws['A1'].font = font_header_1
        ws['A1'].alignment = center_align

        ws['E1'] = str(page_number)
        ws['E1'].font = font_page
        ws['E1'].alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A2:E2")
    ws['A2'] = "В ОБЪЕДИНЕНИИ, ПРОШЕДШИХ ИНСТРУКТАЖ ПО ТЕХНИКЕ БЕЗОПАСНОСТИ"
    ws['A2'].font = font_header_2
    ws['A2'].alignment = center_align

    # 2. Шапка таблицы (строка 4)
    ws.row_dimensions[4].height = 48
    cols_header = [
        ('A', "№\nп/п"),
        ('B', "Фамилия, имя обучающегося"),
        ('C', "Дата\nпроведения\nинструктажа"),
        ('D', "Краткое содержание инструктажа"),
        ('E', "Подпись\nпроводившего инструктаж\n(разборчиво)")
    ]
    for col_letter, text in cols_header:
        cell = ws[f"{col_letter}4"]
        cell.value = text
        cell.font = font_table_h
        cell.alignment = center_align
        cell.border = thin_border

    # 3. 28 строк данных (строки 5..32)
    items_list = items or []

    # Находим общее содержание и общую подпись (первое непустое значение)
    common_content = ""
    common_signature = ""
    for it in items_list:
        if isinstance(it, dict):
            if not common_content and it.get("content", "").strip():
                common_content = it.get("content", "").strip()
            if not common_signature and it.get("signature", "").strip():
                common_signature = it.get("signature", "").strip()

    for i in range(28):
        row_idx = 5 + i
        ws.row_dimensions[row_idx].height = 19.0
        data = items_list[i] if i < len(items_list) and isinstance(items_list[i], dict) else {}

        # № п/п
        c_a = ws[f"A{row_idx}"]
        c_a.value = i + 1
        c_a.font = font_cell
        c_a.alignment = center_align
        c_a.border = thin_border

        # Фамилия, имя обучающегося
        c_b = ws[f"B{row_idx}"]
        c_b.value = data.get("student", "")
        c_b.font = font_cell
        c_b.alignment = left_align
        c_b.border = thin_border

        # Дата проведения инструктажа
        c_c = ws[f"C{row_idx}"]
        c_c.value = data.get("date", "")
        c_c.font = font_cell
        c_c.alignment = center_align
        c_c.border = thin_border

        # Границы для ячеек D и E во всех строках
        ws[f"D{row_idx}"].border = thin_border
        ws[f"E{row_idx}"].border = thin_border

    # Объединение всех строк под столбцом 'Краткое содержание инструктажа' (D5:D32)
    # Выравнивание по верхнему и левому краю
    ws.merge_cells("D5:D32")
    ws['D5'].value = common_content
    ws['D5'].font = font_cell
    ws['D5'].alignment = top_left_align

    # Объединение всех строк под столбцом 'Подпись проводившего инструктаж (разборчиво)' (E5:E32)
    # Единая подпись, выравнивание по центру
    ws.merge_cells("E5:E32")
    ws['E5'].value = common_signature
    ws['E5'].font = font_cell
    ws['E5'].alignment = center_align


def generate_excel_safety_briefing_p1(save_path: str, items: list = None) -> bool:
    """Генерация печатного листа 'Инструктаж по технике безопасности' (стр. 38) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_safety_briefing_sheet(ws, page_number=38, items=items)
    wb.save(save_path)
    return True


def generate_excel_safety_briefing_p2(save_path: str, items: list = None) -> bool:
    """Генерация печатного листа 'Инструктаж по технике безопасности' (стр. 39) А4."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_safety_briefing_sheet(ws, page_number=39, items=items)
    wb.save(save_path)
    return True


def generate_excel_safety_briefing_spread(save_path: str, items_p1: list = None, items_p2: list = None) -> bool:
    """Генерация разворота 'Инструктаж по технике безопасности' (стр. 38 и 39) в один файл Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws1 = wb.active
    populate_safety_briefing_sheet(ws1, page_number=38, items=items_p1)
    ws2 = wb.create_sheet(title="Инструктаж по ТБ (стр. 39)")
    populate_safety_briefing_sheet(ws2, page_number=39, items=items_p2)
    wb.save(save_path)
    return True


def populate_annual_report_sheet(
    ws,
    page_number: int = 40,
    data: list = None
):
    """
    Форматирование листа 'Годовой цифровой отчёт' (стр. 40 журнала) А4.
    Таблица: 6 больших столбцов, последние 2 поделены на меньшие столбцы. Всего 18 элементарных столбцов.
    4 строки данных (I полугодие, II полугодие, За год, дополнительная).
    Ниже таблицы: блок текста 'Требования к руководителям объединений...' (пункты 1-10).
    Четная страница (стр. 40): отступ 30 мм СПРАВА под переплет, номер страницы слева.
    """
    ws.title = f"Годовой отчёт (стр. {page_number})"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Четная страница 40: корешок 30 мм СПРАВА (1.18 дюйма), слева 10 мм (0.39 дюйма)
    ws.page_margins = PageMargins(left=0.39, right=1.18, top=0.35, bottom=0.35, header=0.15, footer=0.15)

    # 18 колонок: A..R
    col_widths = {
        'A': 15.5,  # Учебный период (с запасом для 'I полугодие' и 'II полугодие')
        'B': 8.5,   # Всего в объединении
        'C': 7.5,   # Мальчиков
        'D': 7.5,   # Девочек
        'E': 4.3,   # I
        'F': 4.3,   # II
        'G': 4.3,   # III
        'H': 4.3,   # IV
        'I': 4.3,   # V
        'J': 4.3,   # VI
        'K': 4.3,   # VII
        'L': 4.3,   # VIII
        'M': 4.3,   # IX
        'N': 4.3,   # X
        'O': 4.3,   # XI
        'P': 4.5,   # 1
        'Q': 4.5,   # 2
        'R': 6.8,   # 3 и более
    }
    for col_letter, w in col_widths.items():
        ws.column_dimensions[col_letter].width = w

    font_page = Font(name="Times New Roman", size=14, bold=True)
    font_title = Font(name="Times New Roman", size=12, bold=True)
    font_th = Font(name="Times New Roman", size=8, bold=True)
    font_cell = Font(name="Times New Roman", size=9)
    font_period = Font(name="Times New Roman", size=8.5, bold=True)
    font_req_h1 = Font(name="Times New Roman", size=10.5, bold=True)
    font_req_h2 = Font(name="Times New Roman", size=9, bold=True)
    font_req_lead = Font(name="Times New Roman", size=8.5, bold=True)
    font_req_body = Font(name="Times New Roman", size=8)
    font_footer = Font(name="Times New Roman", size=6.5, italic=True, color="475569")

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Строка 1: номер страницы '40' слева и заголовок 'ГОДОВОЙ ЦИФРОВОЙ ОТЧЁТ'
    ws.row_dimensions[1].height = 20
    ws['A1'] = str(page_number)
    ws['A1'].font = font_page
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("B1:R1")
    ws['B1'] = "ГОДОВОЙ ЦИФРОВОЙ ОТЧЁТ"
    ws['B1'].font = font_title
    ws['B1'].alignment = center_align

    # Строка 2: небольшой разделитель
    ws.row_dimensions[2].height = 4

    # Шапка таблицы (строки 3 и 4)
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 22

    # A: Учебный период (A3:A4)
    ws.merge_cells("A3:A4")
    ws['A3'] = "Учебный\nпериод"

    # B: Всего в объединении (B3:B4)
    ws.merge_cells("B3:B4")
    ws['B3'] = "Всего\nв объеди-\nнении"

    # C: Мальчиков (C3:C4)
    ws.merge_cells("C3:C4")
    ws['C3'] = "Маль-\nчиков"

    # D: Девочек (D3:D4)
    ws.merge_cells("D3:D4")
    ws['D3'] = "Девочек"

    # E..O: Количество обучающихся по классам (E3:O3)
    ws.merge_cells("E3:O3")
    ws['E3'] = "Количество обучающихся по классам"

    # P..R: Сколько лет посещает объединение (P3:R3)
    ws.merge_cells("P3:R3")
    ws['P3'] = "Сколько лет\nпосещает объединение"

    # Подшапка (строка 4)
    roman_classes = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI"]
    for idx, r_txt in enumerate(roman_classes):
        col_char = chr(ord('E') + idx)
        ws[f"{col_char}4"] = r_txt

    ws["P4"] = "1"
    ws["Q4"] = "2"
    ws["R4"] = "3\nи более"

    # Рамки и выравнивание для шапки таблицы
    all_col_letters = list(col_widths.keys())
    for r_num in (3, 4):
        for c_let in all_col_letters:
            c = ws[f"{c_let}{r_num}"]
            c.font = font_th
            c.alignment = center_align
            c.border = thin_border

    # Строки данных (строки 5, 6, 7, 8) - всего 4 строки
    default_periods = ["I полугодие", "II полугодие", "За год", ""]

    data_rows = data or []

    for row_idx in range(4):
        sheet_row = 5 + row_idx
        ws.row_dimensions[sheet_row].height = 24

        item = data_rows[row_idx] if row_idx < len(data_rows) and isinstance(data_rows[row_idx], dict) else {}

        # Название периода
        p_val = item.get("period", "").strip()
        if not p_val and row_idx < len(default_periods):
            p_val = default_periods[row_idx]

        c_a = ws[f"A{sheet_row}"]
        c_a.value = p_val
        c_a.font = font_period
        c_a.alignment = center_align
        c_a.border = thin_border

        ws[f"B{sheet_row}"] = item.get("total", "")
        ws[f"C{sheet_row}"] = item.get("boys", "")
        ws[f"D{sheet_row}"] = item.get("girls", "")

        # Классы E..O (11 классов)
        item_classes = item.get("classes", [])
        for c_i in range(11):
            c_char = chr(ord('E') + c_i)
            c_val = item_classes[c_i] if c_i < len(item_classes) else ""
            ws[f"{c_char}{sheet_row}"] = c_val

        # Сколько лет P..R (3 колонки)
        item_years = item.get("years_in_org", [])
        y_cols = ['P', 'Q', 'R']
        for y_i, y_col in enumerate(y_cols):
            y_val = item_years[y_i] if y_i < len(item_years) else ""
            ws[f"{y_col}{sheet_row}"] = y_val

        # Рамки и выравнивание для строки данных (столбцы B..R)
        for c_let in all_col_letters:
            if c_let == 'A':
                continue
            c = ws[f"{c_let}{sheet_row}"]
            c.font = font_cell
            c.alignment = center_align
            c.border = thin_border

    # Отступ после таблицы
    ws.row_dimensions[9].height = 10

    # Блок текста: ТРЕБОВАНИЯ
    ws.row_dimensions[10].height = 16
    ws.merge_cells("A10:R10")
    ws['A10'] = "ТРЕБОВАНИЯ"
    ws['A10'].font = font_req_h1
    ws['A10'].alignment = center_align

    ws.row_dimensions[11].height = 24
    ws.merge_cells("A11:R11")
    ws['A11'] = "К РУКОВОДИТЕЛЯМ ОБЪЕДИНЕНИЙ И ПОДРАЗДЕЛЕНИЙ ОРГАНИЗАЦИЙ ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ ДЕТЕЙ ПО ОХРАНЕ ТРУДА, ТЕХНИКЕ БЕЗОПАСНОСТИ И ПРОИЗВОДСТВЕННОЙ САНИТАРИИ"
    ws['A11'].font = font_req_h2
    ws['A11'].alignment = center_align

    ws.row_dimensions[12].height = 4

    ws.row_dimensions[13].height = 16
    ws.merge_cells("A13:R13")
    ws['A13'] = "РУКОВОДИТЕЛЬ ОБЪЕДИНЕНИЯ при непосредственном участии и помощи заведующего лабораторией, кабинетом, мастерской:"
    ws['A13'].font = font_req_lead
    ws['A13'].alignment = left_align

    requirements_text = [
        ("1. Принимает необходимые меры для создания здоровых и безопасных условий проведения занятий.", 14),
        ("2. Обеспечивает выполнение действующих правил и инструкций по технике безопасности и производственной санитарии.", 14),
        ("3. Проводит занятия и работы при наличии соответствующего оборудования и других условий, предусмотренных правилами и нормами по технике безопасности.", 20),
        ("4. Обеспечивает безопасное состояние рабочих мест, оборудования, приборов, инструментов и санитарное состояние помещений.", 20),
        ("5. Проводит инструктаж обучающихся в объединении по технике безопасности с соответствующим оформлением инструктажа в журнале (см. «Список обучающихся в объединении, прошедших инструктаж по ТБ»).", 22),
        ("6. Разрабатывает мероприятия по технике безопасности для включения их в план и соглашение по охране труда.", 18),
        ("7. Не допускает обучающихся в объединении к проведению работы или занятий без предусмотренной спецодежды и защитных приспособлений.", 18),
        ("8. Приостанавливает проведение работы и занятий, сопряжённых с опасностью для жизни, и докладывает об этом руководителю организации.", 18),
        ("9. Немедленно извещает руководителя организации о каждом несчастном случае.", 14),
        ("10. Несет ответственность за несчастные случаи, происшедшие в результате невыполнения им обязанностей, возложенных настоящими требованиями, ФЗ № 273 от 29.12.2012 «Об образовании в РФ» (п. 4 ч. 4 ст. 41), Приказом Минобрнауки РФ № 602 от 27.06.2017 «Об утверждении Порядка расследования и учета несчастных случаев с обучающимися во время пребывания в организации, осуществляющей образовательную деятельность».", 42),
    ]

    for req_idx, (req_text, req_height) in enumerate(requirements_text):
        cur_row = 14 + req_idx
        ws.row_dimensions[cur_row].height = req_height
        ws.merge_cells(f"A{cur_row}:R{cur_row}")
        ws[f"A{cur_row}"] = req_text
        ws[f"A{cur_row}"].font = font_req_body
        ws[f"A{cur_row}"].alignment = left_align

    # Выходные данные внизу листа (строка 25)
    ws.row_dimensions[24].height = 8
    ws.row_dimensions[25].height = 14
    ws.merge_cells("A25:R25")
    ws['A25'] = "Журнал учёта работы педагога дополнительного образования в объединении (секции, клубе, кружке). Формат 60×84/8. 40 с."
    ws['A25'].font = font_footer
    ws['A25'].alignment = center_align


def generate_excel_annual_report(save_path: str, data: list = None) -> bool:
    """Генерация страницы 'Годовой цифровой отчёт' (стр. 40 журнала) в отдельный файл Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_annual_report_sheet(ws, page_number=40, data=data)
    wb.save(save_path)
    return True


def populate_work_hours_sheet(
    ws,
    data: list = None,
    org_name: str = "",
    academic_year: str = "",
    association: str = "",
    group_name: str = "",
    study_year: str = "",
    teacher_name: str = "",
    accompanist_name: str = ""
):
    """
    Форматирует рабочий лист под 'Отчёт об отработанном времени (подсчёт часов)' А4 (портретная ориентация).
    Таблица содержит 12 месяцев (Сентябрь – Август), количество занятий, часы педагога,
    часы аккомпаниатора, общую сумму часов за месяц и итоговую строку за учебный год.
    """
    ws.title = "Отработанное время"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Поля для печати А4 (левое под подшивку 25 мм = 0.98", правое 10 мм = 0.39", верх/низ 15 мм = 0.59")
    ws.page_margins = PageMargins(left=0.98, right=0.39, top=0.59, bottom=0.59, header=0.2, footer=0.2)

    # Ширина столбцов (A..G):
    ws.column_dimensions['A'].width = 6.0
    ws.column_dimensions['B'].width = 22.0
    ws.column_dimensions['C'].width = 13.0
    ws.column_dimensions['D'].width = 13.0
    ws.column_dimensions['E'].width = 14.0
    ws.column_dimensions['F'].width = 13.0
    ws.column_dimensions['G'].width = 20.0

    font_org = Font(name="Times New Roman", size=10, italic=True)
    font_title = Font(name="Times New Roman", size=13, bold=True)
    font_subtitle = Font(name="Times New Roman", size=11, bold=True)
    font_meta = Font(name="Times New Roman", size=10)
    font_meta_bold = Font(name="Times New Roman", size=10, bold=True)
    font_th = Font(name="Times New Roman", size=9.5, bold=True)
    font_cell = Font(name="Times New Roman", size=9.5)
    font_cell_bold = Font(name="Times New Roman", size=10, bold=True)
    font_sign = Font(name="Times New Roman", size=10)

    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    th_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") if OPENPYXL_AVAILABLE else None
    tot_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid") if OPENPYXL_AVAILABLE else None

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 1. Шапка документа
    ws.row_dimensions[1].height = 16
    ws.merge_cells("A1:G1")
    ws['A1'] = org_name or "Образовательная организация дополнительного образования"
    ws['A1'].font = font_org
    ws['A1'].alignment = center_align

    ws.row_dimensions[2].height = 22
    ws.merge_cells("A2:G2")
    ws['A2'] = "ОТЧЁТ ОБ ОТРАБОТАННОМ ВРЕМЕНИ"
    ws['A2'].font = font_title
    ws['A2'].alignment = center_align

    ws.row_dimensions[3].height = 18
    ws.merge_cells("A3:G3")
    ay_str = f"за {academic_year} учебный год" if academic_year else "за учебный год"
    ws['A3'] = f"учёта педагогических часов {ay_str}"
    ws['A3'].font = font_subtitle
    ws['A3'].alignment = center_align

    ws.row_dimensions[4].height = 6

    ws.row_dimensions[5].height = 18
    ws.merge_cells("A5:G5")
    assoc_info = f"Объединение: {association or '—'}     Группа: {group_name or '—'}     Год обучения: {study_year or '—'}"
    ws['A5'] = assoc_info
    ws['A5'].font = font_meta
    ws['A5'].alignment = left_align

    ws.row_dimensions[6].height = 18
    ws.merge_cells("A6:G6")
    ws['A6'] = f"Педагог (руководитель объединения): {teacher_name or '—'}"
    ws['A6'].font = font_meta_bold
    ws['A6'].alignment = left_align

    ws.row_dimensions[7].height = 18
    ws.merge_cells("A7:G7")
    if accompanist_name:
        ws['A7'] = f"Аккомпаниатор (концертмейстер): {accompanist_name}"
        ws['A7'].font = font_meta
    else:
        ws['A7'] = ""
    ws['A7'].alignment = left_align

    ws.row_dimensions[8].height = 6

    # 2. Таблица: Двухуровневая шапка (строки 9 и 10)
    ws.row_dimensions[9].height = 20
    ws.row_dimensions[10].height = 20

    ws.merge_cells("A9:A10")
    ws['A9'] = "№\nп/п"

    ws.merge_cells("B9:B10")
    ws['B9'] = "Учебный месяц"

    ws.merge_cells("C9:C10")
    ws['C9'] = "Количество\nзанятий"

    ws.merge_cells("D9:F9")
    ws['D9'] = "Отработано часов"

    ws['D10'] = "Педагог"
    ws['E10'] = "Аккомпаниатор"
    ws['F10'] = "Всего"

    ws.merge_cells("G9:G10")
    ws['G9'] = "Примечание /\nподпись"

    for r in range(9, 11):
        for col_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            cell = ws[f"{col_letter}{r}"]
            cell.font = font_th
            cell.alignment = center_align
            cell.border = thin_border
            if th_fill:
                cell.fill = th_fill

    # 3. Строки 12 месяцев (строки 11..22)
    months_list = data if (data and isinstance(data, list)) else []
    total_lessons = 0
    total_ht = 0.0
    total_ha = 0.0
    total_all = 0.0

    for i in range(12):
        row_num = 11 + i
        ws.row_dimensions[row_num].height = 20
        m_item = months_list[i] if i < len(months_list) and isinstance(months_list[i], dict) else {}

        m_name = m_item.get("month_name", MONTHS_CONFIG[i]["name"] if i < len(MONTHS_CONFIG) else f"Месяц {i+1}")
        les_val = str(m_item.get("lessons_count", "")).strip()
        ht_val = str(m_item.get("hours_teacher", "")).strip()
        ha_val = str(m_item.get("hours_acc", "")).strip()
        tot_m_val = str(m_item.get("hours_total", "")).strip()
        note_val = str(m_item.get("notes", "")).strip()

        try:
            if les_val:
                total_lessons += int(les_val)
        except ValueError:
            pass

        try:
            if ht_val:
                total_ht += float(ht_val.replace(',', '.'))
        except ValueError:
            pass

        try:
            if ha_val:
                total_ha += float(ha_val.replace(',', '.'))
        except ValueError:
            pass

        try:
            if tot_m_val:
                total_all += float(tot_m_val.replace(',', '.'))
            elif ht_val or ha_val:
                ht_f = float(ht_val.replace(',', '.')) if ht_val else 0.0
                ha_f = float(ha_val.replace(',', '.')) if ha_val else 0.0
                calc_tot = ht_f + ha_f
                tot_m_val = f"{int(calc_tot)}" if calc_tot == int(calc_tot) else f"{calc_tot:.1f}"
                total_all += calc_tot
        except ValueError:
            pass

        row_vals = [
            ('A', str(i + 1), center_align, font_cell),
            ('B', m_name, left_align, font_cell_bold),
            ('C', les_val, center_align, font_cell),
            ('D', ht_val, center_align, font_cell),
            ('E', ha_val, center_align, font_cell),
            ('F', tot_m_val, center_align, font_cell_bold),
            ('G', note_val, left_align, font_cell),
        ]

        for col_l, val, align, fnt in row_vals:
            c = ws[f"{col_l}{row_num}"]
            c.value = val
            c.font = fnt
            c.alignment = align
            c.border = thin_border

    # 4. Строка ИТОГО ЗА ГОД (строка 23)
    ws.row_dimensions[23].height = 24
    ws.merge_cells("A23:B23")
    ws['A23'] = "ИТОГО ЗА УЧЕБНЫЙ ГОД:"
    ws['A23'].font = font_cell_bold
    ws['A23'].alignment = Alignment(horizontal="right", vertical="center")
    ws['A23'].border = thin_border
    ws['B23'].border = thin_border

    tot_les_str = str(total_lessons) if total_lessons > 0 else ""
    tot_ht_str = f"{int(total_ht)}" if total_ht == int(total_ht) else f"{total_ht:.1f}"
    tot_ha_str = f"{int(total_ha)}" if total_ha == int(total_ha) else f"{total_ha:.1f}"
    tot_all_str = f"{int(total_all)}" if total_all == int(total_all) else f"{total_all:.1f}"

    ws['C23'] = tot_les_str if total_lessons > 0 else "0"
    ws['C23'].font = font_cell_bold
    ws['C23'].alignment = center_align
    ws['C23'].border = thin_border

    ws['D23'] = tot_ht_str if total_ht > 0 else "0"
    ws['D23'].font = font_cell_bold
    ws['D23'].alignment = center_align
    ws['D23'].border = thin_border

    ws['E23'] = tot_ha_str if total_ha > 0 else ""
    ws['E23'].font = font_cell_bold
    ws['E23'].alignment = center_align
    ws['E23'].border = thin_border

    ws['F23'] = tot_all_str if total_all > 0 else "0"
    ws['F23'].font = font_cell_bold
    ws['F23'].alignment = center_align
    ws['F23'].border = thin_border

    ws['G23'] = ""
    ws['G23'].border = thin_border

    if tot_fill:
        for cl in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            ws[f"{cl}23"].fill = tot_fill

    # 5. Блок подписей (строки 25..29)
    ws.row_dimensions[24].height = 14

    ws.row_dimensions[25].height = 20
    ws.merge_cells("A25:G25")
    t_fio = teacher_name or "________________________"
    ws['A25'] = f"Руководитель объединения (педагог): _______________________ / {t_fio} /"
    ws['A25'].font = font_sign
    ws['A25'].alignment = left_align

    ws.row_dimensions[26].height = 8

    ws.row_dimensions[27].height = 20
    ws.merge_cells("A27:G27")
    ws['A27'] = "Руководитель организации (зав. отделом): _______________________ / ________________________ /"
    ws['A27'].font = font_sign
    ws['A27'].alignment = left_align

    ws.row_dimensions[28].height = 8

    ws.row_dimensions[29].height = 18
    ws.merge_cells("A29:G29")
    ws['A29'] = "«____» ________________ 202___ г."
    ws['A29'].font = font_sign
    ws['A29'].alignment = left_align


def generate_excel_work_hours(
    save_path: str,
    data: list = None,
    org_name: str = "",
    academic_year: str = "",
    association: str = "",
    group_name: str = "",
    study_year: str = "",
    teacher_name: str = "",
    accompanist_name: str = ""
) -> bool:
    """Генерация отдельного Excel файла 'Отчёт об отработанном времени (подсчёт часов)'."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    wb = Workbook()
    ws = wb.active
    populate_work_hours_sheet(
        ws,
        data=data,
        org_name=org_name,
        academic_year=academic_year,
        association=association,
        group_name=group_name,
        study_year=study_year,
        teacher_name=teacher_name,
        accompanist_name=accompanist_name
    )
    wb.save(save_path)
    return True


def generate_excel_full_journal(
    start_year: str = "",
    end_year: str = "",
    org_name: str = "",
    academic_year: str = "",
    start_month: str = "",
    start_yr: str = "",
    end_month: str = "",
    end_yr: str = "",
    teacher_name: str = "",
    department: str = "",
    association: str = "",
    group_name: str = "",
    study_year: str = "",
    schedule: list = None,
    schedule_changes: list = None,
    rukovoditel: str = "",
    starosta: str = "",
    accompanist: str = "",
    accompanist_schedule: str = "",
    accompanist_changes: str = "",
    save_path: str = None,
    start_day: str = "",
    end_day: str = "",
    **kwargs
) -> bool:
    """Генерация полного комплекта журнала (Обложка + Титульный лист + Оборот + Основные данные + 12 месяцев) в один файл Excel."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Библиотека openpyxl не установлена.\nВыполните в терминале:\npip install openpyxl")
    actual_path = save_path or kwargs.get("filename")
    st = start_year or kwargs.get("start_yr", "")
    en = end_year or kwargs.get("end_yr", "")
    ay = academic_year or kwargs.get("title_academic_year", "")
    sy = start_yr or kwargs.get("start_year_val", "") or kwargs.get("start_year", "")
    ey = end_yr or kwargs.get("end_year_val", "") or kwargs.get("end_year", "")
    sd = start_day or kwargs.get("start_day_val", "") or kwargs.get("start_day", "")
    ed = end_day or kwargs.get("end_day_val", "") or kwargs.get("end_day", "")
    t_name = teacher_name or kwargs.get("teacher", "")

    m_org = kwargs.get("main_org_name") or org_name
    dept = department or kwargs.get("dept", "")
    assoc = association or kwargs.get("assoc", "")
    grp = group_name or kwargs.get("group", "")
    s_yr = study_year or kwargs.get("s_year", "")
    sch = schedule if schedule is not None else kwargs.get("sch", [])
    sch_chg = schedule_changes if schedule_changes is not None else kwargs.get("sch_chg", [])
    ruk = rukovoditel or kwargs.get("ruk", "") or t_name
    star = starosta or kwargs.get("star", "")
    acc = accompanist or kwargs.get("acc", "")
    acc_s = accompanist_schedule or kwargs.get("acc_s", "")
    acc_c = accompanist_changes or kwargs.get("acc_c", "")

    wb = Workbook()
    # 1. Лист Обложки (внешняя сторона, нечетная страница)
    ws1 = wb.active
    populate_cover_sheet(ws1, st, en)

    # 1.1 Оборот обложки (чистый лист, четная страница для двусторонней печати)
    ws_empty = wb.create_sheet(title="Оборот обложки")
    populate_empty_cover_back_sheet(ws_empty)

    # 2. Титульный лист (стр. 1)
    ws2 = wb.create_sheet(title="Титульный лист (стр. 1)")
    populate_title_page_sheet(ws2, org_name, ay, start_month, sy, end_month, ey, start_day=sd, end_day=ed)

    # 3. Оборот титульного листа (стр. 2)
    ws3 = wb.create_sheet(title="Оборот титульного (стр. 2)")
    populate_inside_cover_sheet(ws3, t_name)

    # 4. Основные данные (стр. 3)
    ws4 = wb.create_sheet(title="Основные данные (стр. 3)")
    populate_main_data_sheet(
        ws4,
        org_name=m_org,
        department=dept,
        association=assoc,
        group_name=grp,
        study_year=s_yr,
        schedule=sch,
        schedule_changes=sch_chg,
        rukovoditel=ruk,
        starosta=star,
        accompanist=acc,
        accompanist_schedule=acc_s,
        accompanist_changes=acc_c
    )

    # 5..28. Все 12 месяцев (по 2 страницы на каждый месяц)
    for m in MONTHS_CONFIG:
        k = m["key"]
        m_name = m["name"]
        p1_no = m["p1"]
        p2_no = m["p2"]

        m_st = kwargs.get(f"{k}_students")
        m_dt = kwargs.get(f"{k}_dates")
        m_att = kwargs.get(f"{k}_attendance")
        m_top = kwargs.get(f"{k}_topics")

        ws_p1 = wb.create_sheet(title=f"{m_name} (стр. {p1_no})")
        populate_month_page1_sheet(ws_p1, month_name=m_name, page_number=p1_no, students=m_st, dates=m_dt, attendance=m_att)

        ws_p2 = wb.create_sheet(title=f"{m_name} (стр. {p2_no})")
        populate_month_page2_sheet(ws_p2, month_name=m_name, page_number=p2_no, topics=m_top)

    # 29. Учёт массовых мероприятий с обучающимися (стр. 30)
    ev_p1 = kwargs.get("mass_events_p1")
    ws_ev1 = wb.create_sheet(title="Мероприятия (стр. 30)")
    populate_mass_events_sheet(ws_ev1, page_number=30, events=ev_p1)

    # 30. Учёт массовых мероприятий с обучающимися (стр. 31)
    ev_p2 = kwargs.get("mass_events_p2")
    ws_ev2 = wb.create_sheet(title="Мероприятия (стр. 31)")
    populate_mass_events_sheet(ws_ev2, page_number=31, events=ev_p2)

    # 31. Творческие достижения обучающихся (стр. 32)
    cr_p1 = kwargs.get("creative_achievements_p1")
    ws_cr1 = wb.create_sheet(title="Достижения (стр. 32)")
    populate_creative_achievements_p1_sheet(ws_cr1, page_number=32, items=cr_p1)

    # 32. Творческие достижения обучающихся (стр. 33)
    cr_p2 = kwargs.get("creative_achievements_p2")
    ws_cr2 = wb.create_sheet(title="Достижения (стр. 33)")
    populate_creative_achievements_p2_sheet(ws_cr2, page_number=33, items=cr_p2)

    # 33..38. Список обучающихся (стр. 34-39)
    st_p1 = kwargs.get("students_list_p1")
    st_p2 = kwargs.get("students_list_p2")
    st_p3 = kwargs.get("students_list_p3")
    st_p4 = kwargs.get("students_list_p4")
    st_p5 = kwargs.get("students_list_p5")
    st_p6 = kwargs.get("students_list_p6")

    ws_st1 = wb.create_sheet(title="Обучающиеся (стр. 34)")
    populate_students_list_odd_sheet(ws_st1, page_number=34, start_student_no=1, items=st_p1)

    ws_st2 = wb.create_sheet(title="Обучающиеся (стр. 35)")
    populate_students_list_even_sheet(ws_st2, page_number=35, items=st_p2)

    ws_st3 = wb.create_sheet(title="Обучающиеся (стр. 36)")
    populate_students_list_odd_sheet(ws_st3, page_number=36, start_student_no=11, items=st_p3)

    ws_st4 = wb.create_sheet(title="Обучающиеся (стр. 37)")
    populate_students_list_even_sheet(ws_st4, page_number=37, items=st_p4)

    ws_st5 = wb.create_sheet(title="Обучающиеся (стр. 38)")
    populate_students_list_odd_sheet(ws_st5, page_number=38, start_student_no=21, items=st_p5)

    ws_st6 = wb.create_sheet(title="Обучающиеся (стр. 39)")
    populate_students_list_even_sheet(ws_st6, page_number=39, items=st_p6)

    # 39..40. Список обучающихся, прошедших инструктаж по технике безопасности (стр. 38 и 39)
    sb_p1 = kwargs.get("safety_briefing_p1")
    sb_p2 = kwargs.get("safety_briefing_p2")
    ws_sb1 = wb.create_sheet(title="Инструктаж по ТБ (стр. 38)")
    populate_safety_briefing_sheet(ws_sb1, page_number=38, items=sb_p1)
    ws_sb2 = wb.create_sheet(title="Инструктаж по ТБ (стр. 39)")
    populate_safety_briefing_sheet(ws_sb2, page_number=39, items=sb_p2)

    # 41. Годовой цифровой отчёт (стр. 40 журнала)
    ar_data = kwargs.get("annual_report")
    ws_ar = wb.create_sheet(title="Годовой отчёт (стр. 40)")
    populate_annual_report_sheet(ws_ar, page_number=40, data=ar_data)

    wb.save(actual_path)
    return True



class JournalCoverApp(tk.Tk):
    """Главное окно с вкладками (Ввод параметров / Формирование Excel) и выпадающим списком истории."""

    def __init__(self):
        super().__init__()
        self.title(f"Журнал учёта работы педагога — Обложка А4 (v{APP_VERSION})")

        self.detect_crisp_fonts()
        self.configure_dpi_scaling()

        self.base_w = 1140
        self.base_h = 760
        self.minsize(920, 600)

        self._month_save_timers = {}
        self._events_save_timer = None
        self._creative_save_timer = None
        self._students_save_timer = None

        # Загрузка данных из config.json
        self.config_data = load_config()
        initial_year = self.config_data.get("last_academic_year", "2024 / 2025")
        initial_teacher = self.config_data.get("teacher_name", "")
        initial_org = self.config_data.get("org_name", "ГБУ ДО Республиканский детский образовательный технопарк")
        initial_title_yr = self.config_data.get("title_academic_year", initial_year)
        initial_sd = str(self.config_data.get("start_day_val", ""))
        initial_sm = self.config_data.get("start_month", "сентября")
        initial_sy = str(self.config_data.get("start_year_val", "2024"))
        if len(initial_sy) == 2:
            initial_sy = f"20{initial_sy}"
        initial_ed = str(self.config_data.get("end_day_val", ""))
        initial_em = self.config_data.get("end_month", "мая")
        initial_ey = str(self.config_data.get("end_year_val", "2025"))
        if len(initial_ey) == 2:
            initial_ey = f"20{initial_ey}"

        # Переменные Обложки, Титульного и Оборота
        self.year_var = tk.StringVar(value=initial_year)
        self.teacher_var = tk.StringVar(value=initial_teacher)
        self.org_var = tk.StringVar(value=initial_org)
        self.title_year_var = tk.StringVar(value=initial_title_yr)
        self.start_day_var = tk.StringVar(value=initial_sd)
        self.start_month_var = tk.StringVar(value=initial_sm)
        self.start_year_var = tk.StringVar(value=initial_sy)
        self.end_day_var = tk.StringVar(value=initial_ed)
        self.end_month_var = tk.StringVar(value=initial_em)
        self.end_year_var = tk.StringVar(value=initial_ey)

        # Переменные для вкладки «Основные данные» (стр. 3 журнала)
        initial_main_org = self.config_data.get("main_org_name", initial_org)
        initial_dept = self.config_data.get("department", "")
        initial_assoc = self.config_data.get("association", "")
        initial_group = self.config_data.get("group_name", "")
        initial_study_year = self.config_data.get("study_year", "1-й год")
        initial_ruk = self.config_data.get("rukovoditel", initial_teacher)
        initial_starosta = self.config_data.get("starosta", "")
        initial_acc = self.config_data.get("accompanist", "")
        initial_acc_sched = self.config_data.get("accompanist_schedule", "")
        initial_acc_chg = self.config_data.get("accompanist_changes", "")

        self.main_org_var = tk.StringVar(value=initial_main_org)
        self.department_var = tk.StringVar(value=initial_dept)
        self.association_var = tk.StringVar(value=initial_assoc)
        self.group_var = tk.StringVar(value=initial_group)
        self.study_year_var = tk.StringVar(value=initial_study_year)
        self.rukovoditel_var = tk.StringVar(value=initial_ruk)
        self.starosta_var = tk.StringVar(value=initial_starosta)
        self.accompanist_var = tk.StringVar(value=initial_acc)
        self.acc_schedule_var = tk.StringVar(value=initial_acc_sched)
        self.acc_changes_var = tk.StringVar(value=initial_acc_chg)

        # Таблица расписания занятий (6 строк и 2 столбца: дни недели и время)
        self.schedule_row_vars = []
        saved_sched = self.config_data.get("schedule", [])
        for i in range(6):
            d_val = saved_sched[i].get("day", "") if i < len(saved_sched) and isinstance(saved_sched[i], dict) else ""
            t_val = saved_sched[i].get("time", "") if i < len(saved_sched) and isinstance(saved_sched[i], dict) else ""
            dv = tk.StringVar(value=d_val)
            tv = tk.StringVar(value=t_val)
            dv.trace_add("write", lambda *a: self.auto_save_main_data())
            tv.trace_add("write", lambda *a: self.auto_save_main_data())
            self.schedule_row_vars.append((dv, tv))

        # Виджеты таблицы изменения расписания занятий
        self.schedule_changes_row_widgets = []
        self.saved_sched_changes = self.config_data.get("schedule_changes", [])

        # Автосохранение основных данных при изменении любого поля
        for v in [self.main_org_var, self.department_var, self.association_var, self.group_var,
                  self.study_year_var, self.rukovoditel_var, self.starosta_var,
                  self.accompanist_var, self.acc_schedule_var, self.acc_changes_var]:
            v.trace_add("write", lambda *a: self.auto_save_main_data())

        # =====================================================================
        # Данные всех 12 вкладок месяцев (Сентябрь — Август, стр. 4..27 журнала)
        # =====================================================================
        self.months_data = {}
        for m in MONTHS_CONFIG:
            m_key = m["key"]
            m_name = m["name"]

            # 1. Первая страница (стр. {p1}): 30 обучающихся, 15 дат занятий, отметки посещаемости (30x15)
            saved_students = self.config_data.get(f"{m_key}_students", [])
            student_vars = []
            for i in range(30):
                st_val = saved_students[i] if i < len(saved_students) and isinstance(saved_students[i], str) else ""
                sv = tk.StringVar(value=st_val)
                sv.trace_add("write", lambda *a, k=m_key: self.schedule_auto_save_month(k))
                student_vars.append(sv)

            saved_dates = self.config_data.get(f"{m_key}_dates", [])
            date_vars = []
            for i in range(15):
                d_val = saved_dates[i] if i < len(saved_dates) and isinstance(saved_dates[i], str) else ""
                dv = tk.StringVar(value=d_val)
                dv.trace_add("write", lambda *a, k=m_key: self.schedule_auto_save_month(k))
                date_vars.append(dv)

            saved_attendance = self.config_data.get(f"{m_key}_attendance", [])
            attendance_vars = []  # 30 строк x 15 колонок
            for r in range(30):
                row_vars = []
                saved_row = saved_attendance[r] if r < len(saved_attendance) and isinstance(saved_attendance[r], list) else []
                for c in range(15):
                    att_val = saved_row[c] if c < len(saved_row) and isinstance(saved_row[c], str) else ""
                    av = tk.StringVar(value=att_val)
                    av.trace_add("write", lambda *a, k=m_key: self.schedule_auto_save_month(k))
                    row_vars.append(av)
                attendance_vars.append(row_vars)

            # 2. Вторая страница (стр. {p2}): 16 строк и 6 столбцов
            saved_topics = self.config_data.get(f"{m_key}_topics", [])
            topic_row_vars = []
            for i in range(16):
                t_data = saved_topics[i] if i < len(saved_topics) and isinstance(saved_topics[i], dict) else {}
                v_date = tk.StringVar(value=t_data.get("date", ""))
                v_content = tk.StringVar(value=t_data.get("content", ""))
                v_ht = tk.StringVar(value=t_data.get("hours_teacher", ""))
                v_st = tk.StringVar(value=t_data.get("sign_teacher", ""))
                v_ha = tk.StringVar(value=t_data.get("hours_acc", ""))
                v_sa = tk.StringVar(value=t_data.get("sign_acc", ""))
                for v in [v_date, v_content, v_ht, v_st, v_ha, v_sa]:
                    v.trace_add("write", lambda *a, k=m_key: self.schedule_auto_save_month(k))
                topic_row_vars.append((v_date, v_content, v_ht, v_st, v_ha, v_sa))

            self.months_data[m_key] = {
                "config": m,
                "student_vars": student_vars,
                "date_vars": date_vars,
                "attendance_vars": attendance_vars,
                "topic_row_vars": topic_row_vars,
                "student_widgets": [],
                "topic_widgets": [],
                "notebook": None,
                "tab_p1": None,
                "tab_p2": None,
            }

        # Атрибуты обратной совместимости для Сентября
        self.september_student_vars = self.months_data["september"]["student_vars"]
        self.september_date_vars = self.months_data["september"]["date_vars"]
        self.september_attendance_vars = self.months_data["september"]["attendance_vars"]
        self.september_topic_row_vars = self.months_data["september"]["topic_row_vars"]
        self.september_student_widgets = self.months_data["september"]["student_widgets"]
        self.september_topic_widgets = self.months_data["september"]["topic_widgets"]

        # Инициализация данных для Учёта массовых мероприятий (стр. 30 и стр. 31, 26 строк на каждой странице)
        self.mass_events_p1_vars = []
        saved_ev_p1 = self.config_data.get("mass_events_p1", [])
        for i in range(26):
            item = saved_ev_p1[i] if i < len(saved_ev_p1) and isinstance(saved_ev_p1[i], dict) else {}
            v_d = tk.StringVar(value=item.get("date", ""))
            v_c = tk.StringVar(value=item.get("content", ""))
            v_cnt = tk.StringVar(value=item.get("count", ""))
            v_loc = tk.StringVar(value=item.get("location", ""))
            v_who = tk.StringVar(value=item.get("conducted_by", ""))
            for v in [v_d, v_c, v_cnt, v_loc, v_who]:
                v.trace_add("write", lambda *a: self.schedule_auto_save_mass_events())
            self.mass_events_p1_vars.append((v_d, v_c, v_cnt, v_loc, v_who))

        self.mass_events_p2_vars = []
        saved_ev_p2 = self.config_data.get("mass_events_p2", [])
        for i in range(26):
            item = saved_ev_p2[i] if i < len(saved_ev_p2) and isinstance(saved_ev_p2[i], dict) else {}
            v_d = tk.StringVar(value=item.get("date", ""))
            v_c = tk.StringVar(value=item.get("content", ""))
            v_cnt = tk.StringVar(value=item.get("count", ""))
            v_loc = tk.StringVar(value=item.get("location", ""))
            v_who = tk.StringVar(value=item.get("conducted_by", ""))
            for v in [v_d, v_c, v_cnt, v_loc, v_who]:
                v.trace_add("write", lambda *a: self.schedule_auto_save_mass_events())
            self.mass_events_p2_vars.append((v_d, v_c, v_cnt, v_loc, v_who))

        self.mass_events_p1_widgets = []
        self.mass_events_p2_widgets = []
        self.mass_events_notebook = None
        self.last_saved_events_p1_file = None
        self.last_saved_events_p2_file = None
        self.last_saved_events_spread_file = None

        # Инициализация данных для Творческих достижений (стр. 32 и стр. 33, 26 строк на каждой странице)
        self.creative_achievements_p1_vars = []
        saved_cr_p1 = self.config_data.get("creative_achievements_p1", [])
        for i in range(26):
            item = saved_cr_p1[i] if i < len(saved_cr_p1) and isinstance(saved_cr_p1[i], dict) else {}
            v_st = tk.StringVar(value=item.get("student", ""))
            v_ev = tk.StringVar(value=item.get("event", ""))
            for v in [v_st, v_ev]:
                v.trace_add("write", lambda *a: self.schedule_auto_save_creative())
            self.creative_achievements_p1_vars.append((v_st, v_ev))

        self.creative_achievements_p2_vars = []
        saved_cr_p2 = self.config_data.get("creative_achievements_p2", [])
        for i in range(26):
            item = saved_cr_p2[i] if i < len(saved_cr_p2) and isinstance(saved_cr_p2[i], dict) else {}
            v_res = tk.StringVar(value=item.get("results", ""))
            v_wrk = tk.StringVar(value=item.get("works", ""))
            for v in [v_res, v_wrk]:
                v.trace_add("write", lambda *a: self.schedule_auto_save_creative())
            self.creative_achievements_p2_vars.append((v_res, v_wrk))

        self.creative_achievements_p1_widgets = []
        self.creative_achievements_p2_widgets = []
        self.creative_achievements_notebook = None
        self.last_saved_creative_p1_file = None
        self.last_saved_creative_p2_file = None
        self.last_saved_creative_spread_file = None

        # Инициализация данных для Списка обучающихся (стр. 34-39 журнала, 6 страниц по 10 строк)
        # Страницы 1, 3, 5: 6 колонок (№, ФИО, Год рожд., Школа/класс, Район, Врач)
        # Страницы 2, 4, 6: 5 колонок (Адрес/тел, Родители/тел, Дата вступления, Выбыл, Примечания)
        self.students_list_pages_vars = {p: [] for p in range(1, 7)}
        for p in (1, 3, 5):
            k = f"students_list_p{p}"
            saved_items = self.config_data.get(k, [])
            for i in range(10):
                item = saved_items[i] if i < len(saved_items) and isinstance(saved_items[i], dict) else {}
                v_st = tk.StringVar(value=item.get("student", ""))
                v_by = tk.StringVar(value=item.get("birth_year", ""))
                v_sc = tk.StringVar(value=item.get("school_class", ""))
                v_ds = tk.StringVar(value=item.get("district", ""))
                v_dc = tk.StringVar(value=item.get("doctor_conclusion", ""))
                for v in (v_st, v_by, v_sc, v_ds, v_dc):
                    v.trace_add("write", lambda *a: self.schedule_auto_save_students_list())
                self.students_list_pages_vars[p].append((v_st, v_by, v_sc, v_ds, v_dc))

        for p in (2, 4, 6):
            k = f"students_list_p{p}"
            saved_items = self.config_data.get(k, [])
            for i in range(10):
                item = saved_items[i] if i < len(saved_items) and isinstance(saved_items[i], dict) else {}
                v_ap = tk.StringVar(value=item.get("address_phone", ""))
                v_pi = tk.StringVar(value=item.get("parents_info", ""))
                v_jd = tk.StringVar(value=item.get("join_date", ""))
                v_li = tk.StringVar(value=item.get("leave_info", ""))
                v_nt = tk.StringVar(value=item.get("notes", ""))
                for v in (v_ap, v_pi, v_jd, v_li, v_nt):
                    v.trace_add("write", lambda *a: self.schedule_auto_save_students_list())
                self.students_list_pages_vars[p].append((v_ap, v_pi, v_jd, v_li, v_nt))

        self.students_list_pages_widgets = {p: [] for p in range(1, 7)}
        self.students_list_notebook = None
        self.last_saved_students_page_files = {}
        self.last_saved_students_spread_files = {}
        self.last_saved_students_all_file = None

        # Инициализация данных для Инструктажа по технике безопасности (2 страницы по 28 строк)
        # 5 колонок: № п/п, ФИО обучающегося, Дата проведения, Краткое содержание, Подпись проводившего
        self.safety_briefing_p1_vars = []
        saved_sb_p1 = self.config_data.get("safety_briefing_p1", [])
        for i in range(28):
            item = saved_sb_p1[i] if i < len(saved_sb_p1) and isinstance(saved_sb_p1[i], dict) else {}
            v_st = tk.StringVar(value=item.get("student", ""))
            v_dt = tk.StringVar(value=item.get("date", ""))
            v_cnt = tk.StringVar(value=item.get("content", ""))
            v_sig = tk.StringVar(value=item.get("signature", ""))
            for v in (v_st, v_dt, v_cnt, v_sig):
                v.trace_add("write", lambda *a: self.schedule_auto_save_safety_briefing())
            self.safety_briefing_p1_vars.append((v_st, v_dt, v_cnt, v_sig))

        self.safety_briefing_p2_vars = []
        saved_sb_p2 = self.config_data.get("safety_briefing_p2", [])
        for i in range(28):
            item = saved_sb_p2[i] if i < len(saved_sb_p2) and isinstance(saved_sb_p2[i], dict) else {}
            v_st = tk.StringVar(value=item.get("student", ""))
            v_dt = tk.StringVar(value=item.get("date", ""))
            v_cnt = tk.StringVar(value=item.get("content", ""))
            v_sig = tk.StringVar(value=item.get("signature", ""))
            for v in (v_st, v_dt, v_cnt, v_sig):
                v.trace_add("write", lambda *a: self.schedule_auto_save_safety_briefing())
            self.safety_briefing_p2_vars.append((v_st, v_dt, v_cnt, v_sig))

        self.safety_briefing_p1_widgets = []
        self.safety_briefing_p2_widgets = []
        self.safety_briefing_notebook = None
        self.last_saved_safety_p1_file = None
        self.last_saved_safety_p2_file = None
        self.last_saved_safety_spread_file = None

        # Состояние вкладки 'Годовой цифровой отчёт' (стр. 40 журнала, 4 строки)
        # 18 колонок: period, total, boys, girls, classes (11), years_in_org (3)
        self.annual_report_vars = []
        saved_ar = self.config_data.get("annual_report", [])
        default_ar_periods = ["I полугодие", "II полугодие", "За год", ""]
        for r_idx in range(4):
            item = saved_ar[r_idx] if r_idx < len(saved_ar) and isinstance(saved_ar[r_idx], dict) else {}
            def_p = default_ar_periods[r_idx]
            v_per = tk.StringVar(value=item.get("period", def_p))
            v_tot = tk.StringVar(value=item.get("total", ""))
            v_boy = tk.StringVar(value=item.get("boys", ""))
            v_grl = tk.StringVar(value=item.get("girls", ""))

            # classes: 11 vars
            saved_cl = item.get("classes", [])
            cl_vars = []
            for c_i in range(11):
                val_c = saved_cl[c_i] if c_i < len(saved_cl) else ""
                cv = tk.StringVar(value=val_c)
                cl_vars.append(cv)

            # years_in_org: 3 vars (1, 2, 3 и более)
            saved_yr = item.get("years_in_org", [])
            yr_vars = []
            for y_i in range(3):
                val_y = saved_yr[y_i] if y_i < len(saved_yr) else ""
                yv = tk.StringVar(value=val_y)
                yr_vars.append(yv)

            all_row_vars = [v_per, v_tot, v_boy, v_grl] + cl_vars + yr_vars
            for v in all_row_vars:
                v.trace_add("write", lambda *a: self.schedule_auto_save_annual_report())

            self.annual_report_vars.append({
                "period": v_per,
                "total": v_tot,
                "boys": v_boy,
                "girls": v_grl,
                "classes": cl_vars,
                "years": yr_vars,
                "all_vars": all_row_vars
            })

        self.annual_report_widgets = []
        self.last_saved_annual_report_file = None

        # Состояние вкладки 'Отработанное время' (подсчёт часов, 12 месяцев)
        self.work_hours_vars = []
        saved_wh = self.config_data.get("work_hours_report", [])
        for i, m_cfg in enumerate(MONTHS_CONFIG):
            item = saved_wh[i] if i < len(saved_wh) and isinstance(saved_wh[i], dict) else {}
            v_les = tk.StringVar(value=str(item.get("lessons_count", "")))
            v_ht = tk.StringVar(value=str(item.get("hours_teacher", "")))
            v_ha = tk.StringVar(value=str(item.get("hours_acc", "")))
            v_tot = tk.StringVar(value=str(item.get("hours_total", "")))
            v_not = tk.StringVar(value=str(item.get("notes", "")))

            for v in (v_les, v_ht, v_ha, v_not):
                v.trace_add("write", lambda *a: self.schedule_auto_save_work_hours())

            self.work_hours_vars.append({
                "month_key": m_cfg["key"],
                "month_name": m_cfg["name"],
                "lessons_count": v_les,
                "hours_teacher": v_ht,
                "hours_acc": v_ha,
                "hours_total": v_tot,
                "notes": v_not
            })

        self.work_hours_total_lessons = tk.StringVar(value="0")
        self.work_hours_total_teacher = tk.StringVar(value="0")
        self.work_hours_total_acc = tk.StringVar(value="0")
        self.work_hours_total_all = tk.StringVar(value="0")
        self.last_saved_work_hours_file = None
        self.work_hours_widgets = []

        self.start_year = "2024"
        self.end_year = "2025"
        self.parse_years(initial_year)

        self.last_saved_file = None
        self.last_saved_title_file = None
        self.last_saved_inside_file = None
        self.last_saved_main_file = None
        self.last_saved_month_p1_file = {}
        self.last_saved_month_p2_file = {}
        self.last_saved_month_spread_file = {}
        self.last_saved_september_p1_file = None
        self.last_saved_september_p2_file = None
        self.last_saved_september_spread_file = None
        self.last_saved_full_file = None
        self.export_month_var = tk.StringVar(value="Сентябрь")
        self.status_var = tk.StringVar(value="Готов к работе. История загружена из config.json")

        self.year_var.trace_add("write", self.on_year_input_changed)
        self.teacher_var.trace_add("write", self.on_teacher_input_changed)
        self.org_var.trace_add("write", self.on_title_input_changed)
        self.title_year_var.trace_add("write", self.on_title_input_changed)
        self.title_year_var.trace_add("write", lambda *a: self.sync_years_from_academic_year(self.title_year_var.get()))
        self.start_day_var.trace_add("write", self.on_title_input_changed)
        self.start_month_var.trace_add("write", self.on_title_input_changed)
        self.start_year_var.trace_add("write", self.on_title_input_changed)
        self.end_day_var.trace_add("write", self.on_title_input_changed)
        self.end_month_var.trace_add("write", self.on_title_input_changed)
        self.end_year_var.trace_add("write", self.on_title_input_changed)

        self.protocol("WM_DELETE_WINDOW", self.on_window_close)

        self.setup_ui()
        # Глобальный перехват клика в любое место окна для сброса фокуса и автосохранения
        self.bind_all("<Button-1>", self.on_global_click, add="+")
        # Синхронизация годов при запуске
        self.sync_years_from_academic_year(self.title_year_var.get())

        # Центрирование полностью построенного окна по центру экрана
        self.center_window(self.base_w, self.base_h)

        # Сброс фокуса при запуске приложения, чтобы при открытии инпуты не были в фокусе
        self.clear_focus()
        self.after(30, self.clear_focus)
        self.after(100, self.clear_focus)
        self.after(250, self.clear_focus)

        # Фоновая проверка обновлений на GitHub при запуске
        self.after(1500, lambda: self.check_updates_background(silent=True))

    def detect_crisp_fonts(self):
        sys_os = platform.system()
        if sys_os == "Windows":
            self.font_sans = "Segoe UI"
            self.font_mono = "Consolas"
            self.font_serif = "Times New Roman"
        elif sys_os == "Darwin":
            self.font_sans = "SF Pro Text"
            self.font_mono = "Menlo"
            self.font_serif = "Times"
        else:
            self.font_sans = "Ubuntu"
            self.font_mono = "DejaVu Sans Mono"
            self.font_serif = "DejaVu Serif"

    def configure_dpi_scaling(self):
        try:
            dpi = self.winfo_fpixels('1i')
            self.scale_factor = max(1.0, dpi / 96.0)
            if platform.system() != "Windows":
                self.tk.call('tk', 'scaling', dpi / 72.0)
        except Exception:
            self.scale_factor = 1.0

    def center_window(self, w: int = 1140, h: int = 760):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        target_w = min(w, max(920, sw - 80))
        target_h = min(h, max(600, sh - 100))
        x = max(20, (sw - target_w) // 2)
        y = max(20, (sh - target_h) // 2)
        self.geometry(f"{target_w}x{target_h}+{x}+{y}")

    def parse_years(self, text: str):
        digits = re.findall(r'\d+', text)
        if len(digits) >= 2:
            self.start_year = digits[0]
            self.end_year = digits[1]
        elif len(digits) == 1 and len(digits[0]) == 4:
            self.start_year = digits[0]
            self.end_year = str(int(digits[0]) + 1)
        elif len(digits) == 1 and len(digits[0]) == 8:
            self.start_year = digits[0][:4]
            self.end_year = digits[0][4:]

    def add_to_history(self, val: str):
        val = val.strip()
        if not val:
            return
        history = self.config_data.get("history", [])
        if val in history:
            history.remove(val)
        history.insert(0, val)
        self.config_data["history"] = history[:25]
        self.config_data["last_academic_year"] = val
        save_config(self.config_data)
        if hasattr(self, 'combo_year') and isinstance(self.combo_year, ttk.Combobox):
            self.combo_year['values'] = self.config_data.get("history", [])

    def add_to_teacher_history(self, name: str):
        """Добавляет ФИО педагога в историю и сохраняет в config.json."""
        name = name.strip()
        if not name:
            return
        t_history = self.config_data.get("teacher_history", [])
        if name in t_history:
            t_history.remove(name)
        t_history.insert(0, name)
        self.config_data["teacher_history"] = t_history[:25]
        self.config_data["teacher_name"] = name
        save_config(self.config_data)
        if hasattr(self, 'combo_teacher') and isinstance(self.combo_teacher, ttk.Combobox):
            self.combo_teacher['values'] = self.config_data.get("teacher_history", [])

    def add_to_org_history(self, org: str):
        """Добавляет название организации в историю и сохраняет в config.json."""
        org = org.strip()
        if not org:
            return
        o_history = self.config_data.get("org_history", [])
        if org in o_history:
            o_history.remove(org)
        o_history.insert(0, org)
        self.config_data["org_history"] = o_history[:25]
        self.config_data["org_name"] = org
        save_config(self.config_data)
        if hasattr(self, 'combo_org') and isinstance(self.combo_org, ttk.Combobox):
            self.combo_org['values'] = self.config_data.get("org_history", [])

    def save_org_auto(self, event=None):
        """Автоматически сохраняет значение организации в config.json и историю
        при потере фокуса, выборе из списка или клике в любое место интерфейса."""
        if not hasattr(self, 'org_var'):
            return
        org = self.org_var.get().strip()
        if org:
            self.add_to_org_history(org)
            if hasattr(self, 'draw_title_page_preview'):
                self.draw_title_page_preview()
            if hasattr(self, 'update_summary_lbl'):
                self.update_summary_lbl()

    def clear_focus(self, event=None):
        """Сбрасывает фокус с активных полей ввода на главное окно и полностью снимает выделение текста."""
        try:
            for attr in ('combo_org', 'combo_year', 'combo_teacher', 'combo_sd', 'combo_sm', 'entry_sy', 'combo_ed', 'combo_em', 'entry_ey', 'combo_sy', 'combo_ey', 'entry_title_yr'):
                if hasattr(self, attr):
                    w = getattr(self, attr)
                    try:
                        w.selection_clear()
                    except Exception:
                        pass
            focused = self.focus_get()
            if focused is not None:
                try:
                    focused.selection_clear()
                except Exception:
                    pass
            self.focus_set()
        except Exception:
            pass

    def on_tab_changed(self, event=None):
        """При открытии или переключении вкладки:
        1. Если вкладка еще не была построена (отложенный рендеринг) — строит ее мгновенно.
        2. Сбрасывает фокус со всех полей ввода."""
        try:
            if hasattr(self, 'notebook'):
                selected_tab_id = self.notebook.select()
                if selected_tab_id:
                    tab_widget = self.nametowidget(selected_tab_id)
                    build_fn = getattr(tab_widget, "_build_func", None)
                    if build_fn and not getattr(tab_widget, "_is_built", False):
                        tab_widget._is_built = True
                        build_fn()
        except Exception:
            pass
        self.clear_focus()
        self.after(20, self.clear_focus)
        self.after(60, self.clear_focus)
        self.after(150, self.clear_focus)

    def on_global_click(self, event):
        """Обработчик глобального клика мышью в любое место интерфейса.
        Если клик совершен не по полю ввода (по фону, вкладке, кнопке, карточке, холсту и т.д.),
        фокус немедленно сбрасывается. Также автоматически сохраняются введенные данные."""
        widget = getattr(event, 'widget', None)
        is_input = False
        if widget is not None:
            try:
                if isinstance(widget, str):
                    widget = self.nametowidget(widget)
                cls = widget.winfo_class()
                # Поля ввода и интерактивные выпадающие списки
                if cls in ('Entry', 'TEntry', 'Combobox', 'TCombobox', 'Spinbox', 'TSpinbox', 'Text', 'Listbox', 'Scrollbar', 'TScrollbar'):
                    is_input = True
            except Exception:
                pass

        if not is_input:
            self.clear_focus()
        else:
            # Если кликнули в другое поле ввода, снимаем выделение с остальных комбобоксов
            if hasattr(self, 'combo_org') and widget != self.combo_org:
                try:
                    self.combo_org.selection_clear()
                except Exception:
                    pass

        # Автоматическое сохранение параметров при клике в любое место
        if hasattr(self, 'org_var'):
            org = self.org_var.get().strip()
            if org and org != self.config_data.get("org_name"):
                self.add_to_org_history(org)
        if hasattr(self, 'title_year_var'):
            y_val = self.title_year_var.get().strip()
            if y_val and y_val != self.config_data.get("title_academic_year"):
                self.config_data["title_academic_year"] = y_val
                self.config_data["last_academic_year"] = y_val
                save_config(self.config_data)
        if hasattr(self, 'teacher_var'):
            teacher = self.teacher_var.get().strip()
            if teacher and teacher != self.config_data.get("teacher_name"):
                self.config_data["teacher_name"] = teacher
                save_config(self.config_data)
        if hasattr(self, 'combo_year'):
            cov_val = self.combo_year.get().strip()
            if cov_val and cov_val != self.config_data.get("last_academic_year"):
                self.add_to_history(cov_val)

    def sync_years_from_academic_year(self, text: str):
        """Извлекает года из строки учебного года (гггг/гггг) и автоматически подставляет
        соответствующие значения в инпуты 'Начат' (начальный год) и 'Окончен' (конечный год)."""
        digits = re.findall(r'\d{4}', text)
        if len(digits) >= 2:
            y_start = digits[0]
            y_end = digits[1]
        elif len(digits) == 1:
            y_start = digits[0]
            y_end = str(int(digits[0]) + 1)
        else:
            all_d = re.findall(r'\d+', text)
            if len(all_d) >= 2:
                y_start = all_d[0]
                y_end = all_d[1]
            else:
                return

        sy_val = y_start
        ey_val = y_end

        self.start_year = y_start
        self.end_year = y_end

        if hasattr(self, 'year_var') and self.year_var.get() != text:
            self.year_var.set(text)

        if hasattr(self, 'start_year_var') and self.start_year_var.get() != sy_val:
            self.start_year_var.set(sy_val)
        if hasattr(self, 'end_year_var') and self.end_year_var.get() != ey_val:
            self.end_year_var.set(ey_val)

        self.config_data["start_year_val"] = sy_val
        self.config_data["end_year_val"] = ey_val

    def change_academic_year(self, delta: int):
        """Увеличивает или уменьшает учебный год формата гггг / гггг на указанную дельту."""
        val = self.title_year_var.get().strip() if hasattr(self, 'title_year_var') else "2024 / 2025"
        digits = re.findall(r'\d{4}', val)
        if len(digits) >= 2:
            y1 = int(digits[0]) + delta
            y2 = int(digits[1]) + delta
        elif len(digits) == 1:
            y1 = int(digits[0]) + delta
            y2 = y1 + 1
        else:
            y1 = 2024 + delta
            y2 = 2025 + delta

        new_val = f"{y1} / {y2}"
        if hasattr(self, 'title_year_var'):
            self.title_year_var.set(new_val)
        if hasattr(self, 'year_var'):
            self.year_var.set(new_val)

        self.sync_years_from_academic_year(new_val)
        self.config_data["title_academic_year"] = new_val
        self.config_data["last_academic_year"] = new_val
        save_config(self.config_data)

        if hasattr(self, 'draw_title_page_preview'):
            self.draw_title_page_preview()
        if hasattr(self, 'update_summary_lbl'):
            self.update_summary_lbl()
        if hasattr(self, 'status_var'):
            self.status_var.set(f"Учебный год установлен: {new_val}")

    def on_org_focus_out(self, event=None):
        """Сохраняет название организации и гарантированно снимает выделение текста."""
        self.save_org_auto()
        if hasattr(self, 'combo_org'):
            try:
                self.combo_org.selection_clear()
            except Exception:
                pass

    def on_month_selected(self, event=None):
        """Обработчик выбора месяца из выпадающего списка."""
        if hasattr(self, 'start_month_var'):
            self.config_data["start_month"] = self.start_month_var.get().strip()
        if hasattr(self, 'end_month_var'):
            self.config_data["end_month"] = self.end_month_var.get().strip()
        save_config(self.config_data)
        if hasattr(self, 'combo_sm'):
            try:
                self.combo_sm.selection_clear()
            except Exception:
                pass
        if hasattr(self, 'combo_em'):
            try:
                self.combo_em.selection_clear()
            except Exception:
                pass
        self.clear_focus()
        if hasattr(self, 'draw_title_page_preview'):
            self.draw_title_page_preview()
        if hasattr(self, 'update_summary_lbl'):
            self.update_summary_lbl()

    def on_day_selected(self, event=None):
        """Обработчик выбора дня начала или окончания."""
        if hasattr(self, 'start_day_var'):
            self.config_data["start_day_val"] = self.start_day_var.get().strip()
        if hasattr(self, 'end_day_var'):
            self.config_data["end_day_val"] = self.end_day_var.get().strip()
        save_config(self.config_data)
        for attr in ('combo_sd', 'combo_ed'):
            if hasattr(self, attr):
                try:
                    getattr(self, attr).selection_clear()
                except Exception:
                    pass
        self.clear_focus()
        if hasattr(self, 'draw_title_page_preview'):
            self.draw_title_page_preview()
        if hasattr(self, 'update_summary_lbl'):
            self.update_summary_lbl()

    def on_start_year_changed(self, event=None):
        """Обработчик ввода года начала в простое поле Entry."""
        sy = self.start_year_var.get().strip() if hasattr(self, 'start_year_var') else ""
        if len(sy) == 2 and sy.isdigit():
            sy = f"20{sy}"
            self.start_year_var.set(sy)
        self.config_data["start_year_val"] = sy
        save_config(self.config_data)
        for attr in ('entry_sy', 'combo_sy'):
            if hasattr(self, attr):
                try:
                    getattr(self, attr).selection_clear()
                except Exception:
                    pass
        self.clear_focus()
        if hasattr(self, 'draw_title_page_preview'):
            self.draw_title_page_preview()
        if hasattr(self, 'update_summary_lbl'):
            self.update_summary_lbl()

    on_start_year_selected = on_start_year_changed

    def on_end_year_changed(self, event=None):
        """Обработчик ввода года окончания в простое поле Entry."""
        ey = self.end_year_var.get().strip() if hasattr(self, 'end_year_var') else ""
        if len(ey) == 2 and ey.isdigit():
            ey = f"20{ey}"
            self.end_year_var.set(ey)
        self.config_data["end_year_val"] = ey
        save_config(self.config_data)
        for attr in ('entry_ey', 'combo_ey'):
            if hasattr(self, attr):
                try:
                    getattr(self, attr).selection_clear()
                except Exception:
                    pass
        self.clear_focus()
        if hasattr(self, 'draw_title_page_preview'):
            self.draw_title_page_preview()
        if hasattr(self, 'update_summary_lbl'):
            self.update_summary_lbl()

    on_end_year_selected = on_end_year_changed

    def copy_from_widget(self, widget):
        """Копирование выделенного текста в буфер обмена."""
        try:
            if widget.selection_present():
                text = widget.selection_get()
                self.clipboard_clear()
                self.clipboard_append(text)
                self.status_var.set("Текст скопирован в буфер обмена (Ctrl+C)")
            else:
                val = self.year_var.get()
                if val:
                    self.clipboard_clear()
                    self.clipboard_append(val)
                    self.status_var.set("Значение поля скопировано в буфер обмена (Ctrl+C)")
        except Exception:
            pass

    def paste_to_widget(self, widget):
        """Вставка текста из буфера обмена."""
        try:
            text = self.clipboard_get()
            if text:
                if widget.selection_present():
                    first = widget.index(tk.SEL_FIRST)
                    last = widget.index(tk.SEL_LAST)
                    widget.delete(first, last)
                    widget.insert(first, text)
                else:
                    widget.insert(tk.INSERT, text)
                self.status_var.set("Текст вставлен из буфера обмена (Ctrl+V)")
        except Exception:
            pass

    def cut_from_widget(self, widget):
        """Вырезание текста в буфер обмена."""
        try:
            if widget.selection_present():
                text = widget.selection_get()
                self.clipboard_clear()
                self.clipboard_append(text)
                first = widget.index(tk.SEL_FIRST)
                last = widget.index(tk.SEL_LAST)
                widget.delete(first, last)
                self.status_var.set("Текст вырезан в буфер обмена (Ctrl+X)")
        except Exception:
            pass

    def select_all_widget(self, widget):
        """Выделение всего текста в поле ввода."""
        try:
            widget.selection_range(0, tk.END)
            widget.icursor(tk.END)
        except Exception:
            pass

    def attach_context_menu(self, widget):
        """Контекстное меню мыши для поля ввода (правый клик)."""
        menu = tk.Menu(widget, tearoff=0, font=(self.font_sans, 9))
        menu.add_command(label="Копировать (Ctrl+C)", command=lambda: self.copy_from_widget(widget))
        menu.add_command(label="Вставить (Ctrl+V)", command=lambda: self.paste_to_widget(widget))
        menu.add_command(label="Вырезать (Ctrl+X)", command=lambda: self.cut_from_widget(widget))
        menu.add_separator()
        menu.add_command(label="Выделить всё (Ctrl+A)", command=lambda: self.select_all_widget(widget))
        menu.add_command(label="Очистить поле", command=lambda: self.year_var.set(""))

        def show_menu(event):
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)

    def handle_entry_shortcuts(self, event):
        """Универсальный перехват сочетаний клавиш с поддержкой русской раскладки (RU/EN)."""
        # Проверка зажатия Ctrl (Windows/Linux) или Command (macOS)
        state = getattr(event, 'state', 0)
        is_ctrl = bool(state & 0x4) or bool(state & 0x8) or bool(state & 0x40000)

        if not is_ctrl:
            return None

        key = (event.keysym or "").lower()
        code = getattr(event, 'keycode', 0)
        widget = event.widget

        # 1. Ctrl + A / Cmd + A (Выделить всё) - keycode 65 на Windows
        if key in ('a', 'cyrillic_ef') or code == 65:
            self.select_all_widget(widget)
            return "break"

        # 2. Ctrl + C / Cmd + C (Копировать) - keycode 67 на Windows
        elif key in ('c', 'cyrillic_es') or code == 67:
            self.copy_from_widget(widget)
            return "break"

        # 3. Ctrl + V / Cmd + V (Вставить) - keycode 86 на Windows
        elif key in ('v', 'cyrillic_em') or code == 86:
            self.paste_to_widget(widget)
            return "break"

        # 4. Ctrl + X / Cmd + X (Вырезать) - keycode 88 на Windows
        elif key in ('x', 'cyrillic_che') or code == 88:
            self.cut_from_widget(widget)
            return "break"

        # 5. Ctrl + Z / Cmd + Z (Отменить) - keycode 90 на Windows
        elif key in ('z', 'cyrillic_ya') or code == 90:
            last_saved = self.config_data.get("last_academic_year", "2024 / 2025")
            self.year_var.set(last_saved)
            self.status_var.set("Отменено: восстановлено предыдущее значение")
            return "break"

        return None

    def show_shortcuts_help(self):
        """Отображает окно со всеми доступными горячими клавишами."""
        help_text = (
            "Доступные горячие клавиши (RU / EN):\n\n"
            "  • Ctrl + C (Cmd+C) — Копировать\n"
            "  • Ctrl + V (Cmd+V) — Вставить\n"
            "  • Ctrl + A (Cmd+A) — Выделить всё\n"
            "  • Ctrl + X (Cmd+X) — Вырезать\n"
            "  • Ctrl + Z (Cmd+Z) — Отменить"
        )
        messagebox.showinfo("Горячие клавиши", help_text)

    def open_config_file(self):
        """
        Открывает сам файл конфигурации config.json в системной программе по умолчанию
        (Блокнот/Notepad на Windows, TextEdit/open на macOS, xdg-open на Linux).
        Перед открытием сохраняет актуальный учебный год и ФИО педагога и гарантирует существование файла.
        """
        try:
            # Актуализируем текущий год и педагога перед открытием файла пользователем
            current_year = self.year_var.get().strip()
            if current_year:
                self.config_data["last_academic_year"] = current_year
            current_teacher = self.teacher_var.get().strip()
            if current_teacher:
                self.config_data["teacher_name"] = current_teacher
            save_config(self.config_data)

            # Проверяем наличие файла на диске
            if not os.path.exists(CONFIG_FILE):
                save_config(self.config_data)

            # Открываем сам файл в зарегистрированном системном текстовом редакторе
            sys_os = platform.system()
            if sys_os == "Windows":
                # В Windows os.startfile открывает файл в программе, сопоставленной с .json (или Блокноте)
                os.startfile(CONFIG_FILE)
            elif sys_os == "Darwin":
                # В macOS используем команду open
                subprocess.run(["open", CONFIG_FILE], check=False)
            else:
                # В Linux используем xdg-open
                subprocess.run(["xdg-open", CONFIG_FILE], check=False)

            self.status_var.set(f"Файл config.json открыт: {CONFIG_FILE}")
        except Exception as e:
            # Если прямой запуск через систему дал сбой, сообщаем об ошибке и открываем встроенный просмотр
            messagebox.showerror(
                "Ошибка открытия файла",
                f"Не удалось открыть файл в системном редакторе:\n{e}\n\n"
                f"Путь к файлу: {CONFIG_FILE}"
            )
            self.show_config_dialog()

    def get_github_repo(self) -> str:
        """Возвращает репозиторий GitHub из config.json или значение по умолчанию."""
        repo = self.config_data.get("github_repo", "").strip()
        if not repo:
            repo = DEFAULT_GITHUB_REPO
        return repo

    def check_updates_background(self, silent: bool = True):
        """Запускает опрос GitHub API для проверки нового релиза в фоновом потоке."""
        t = threading.Thread(target=self._check_updates_worker, args=(silent,), daemon=True)
        t.start()

    def _check_updates_worker(self, silent: bool = True):
        """Фоновый воркер запроса к GitHub API releases/latest."""
        repo = self.get_github_repo()
        if not repo or "/" not in repo:
            if not silent:
                self.after(0, lambda: messagebox.showwarning(
                    "Проверка обновлений",
                    "Не указан репозиторий GitHub для проверки обновлений."
                ))
            return

        api_url = f"https://api.github.com/repos/{repo}/releases/latest"
        req = urllib.request.Request(
            api_url,
            headers={
                "User-Agent": f"JournalDOP/{APP_VERSION}",
                "Accept": "application/vnd.github.v3+json"
            }
        )

        try:
            with urllib.request.urlopen(req, timeout=4) as resp:
                if resp.status == 200:
                    payload = json.loads(resp.read().decode("utf-8"))
                    latest_tag = payload.get("tag_name", "").strip()
                    release_name = payload.get("name", "") or latest_tag
                    html_url = payload.get("html_url", f"https://github.com/{repo}/releases/latest")
                    release_body = payload.get("body", "")

                    current_v = parse_version_tuple(APP_VERSION)
                    latest_v = parse_version_tuple(latest_tag)

                    if latest_v > current_v:
                        self.after(0, lambda: self._prompt_update(latest_tag, release_name, html_url, release_body))
                    else:
                        if not silent:
                            self.after(0, lambda: messagebox.showinfo(
                                "Обновления",
                                f"У вас установлена самая актуальная версия программы:\n\n"
                                f"Текущая версия: v{APP_VERSION}\n"
                                f"Репозиторий: {repo}"
                            ))
        except Exception as e:
            if not silent:
                self.after(0, lambda: messagebox.showerror(
                    "Ошибка проверки обновлений",
                    f"Не удалось связаться с сервером GitHub:\n{e}\n\n"
                    f"Проверьте подключение к Интернету или репозиторий: {repo}"
                ))

    def _prompt_update(self, latest_tag: str, release_name: str, html_url: str, release_body: str = ""):
        """Показывает диалог с предложением перейти на GitHub и скачать свежий релиз."""
        body_snippet = f"\n\nОписание изменений:\n{release_body[:300]}..." if release_body else ""
        msg = (
            f"Доступна новая версия приложения!\n\n"
            f"Установленная версия: v{APP_VERSION}\n"
            f"Новая версия релиза: {latest_tag} ({release_name}){body_snippet}\n\n"
            f"Перейти на GitHub, чтобы скачать свежий релиз программы?"
        )
        if messagebox.askyesno("Доступно обновление", msg, icon="info"):
            try:
                webbrowser.open(html_url)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть браузер:\n{e}\n\nСсылка для скачивания:\n{html_url}")

    def show_config_dialog(self):
        """Открывает окно для просмотра и управления файлом конфигурации config.json."""
        self.config_data["last_academic_year"] = self.year_var.get().strip()
        cfg_content = json.dumps(self.config_data, ensure_ascii=False, indent=2)

        dlg = tk.Toplevel(self)
        dlg.title("Файл конфигурации: config.json")
        dlg.geometry("540x440")
        dlg.configure(bg="#f8fafc")
        dlg.transient(self)
        dlg.grab_set()

        # Верхняя плашка окна
        header_frame = tk.Frame(dlg, bg="#0f172a", padx=16, pady=12)
        header_frame.pack(fill=tk.X)

        tk.Label(
            header_frame,
            text="Файл конфигурации: config.json",
            font=(self.font_sans, 11, "bold"),
            fg="#ffffff",
            bg="#0f172a"
        ).pack(anchor="w")

        tk.Label(
            header_frame,
            text=f"Путь: {CONFIG_FILE}",
            font=(self.font_sans, 8),
            fg="#94a3b8",
            bg="#0f172a"
        ).pack(anchor="w", pady=(2, 0))

        # Текстовое поле с содержимым JSON
        txt_frame = tk.Frame(dlg, bg="#f8fafc", padx=14, pady=12)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        txt_box = tk.Text(
            txt_frame,
            font=(self.font_mono, 10),
            bg="#090d16",
            fg="#34d399",
            insertbackground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=10
        )
        txt_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        txt_box.insert("1.0", cfg_content)
        txt_box.config(state=tk.DISABLED)

        scroll = ttk.Scrollbar(txt_frame, orient=tk.VERTICAL, command=txt_box.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        txt_box.config(yscrollcommand=scroll.set)

        # Нижняя панель действий
        btn_panel = tk.Frame(dlg, bg="#f1f5f9", padx=14, pady=10, bd=1, relief=tk.SOLID)
        btn_panel.pack(fill=tk.X, side=tk.BOTTOM)

        def copy_to_clipboard():
            self.clipboard_clear()
            self.clipboard_append(cfg_content)
            self.status_var.set("Содержимое config.json скопировано в буфер обмена.")
            messagebox.showinfo("Буфер обмена", "Содержимое config.json успешно скопировано!")

        def open_in_system_editor():
            try:
                self.save_config()
                if platform.system() == "Windows":
                    os.startfile(CONFIG_FILE)
                elif platform.system() == "Darwin":
                    subprocess.run(["open", CONFIG_FILE])
                else:
                    subprocess.run(["xdg-open", CONFIG_FILE])
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")

        tk.Button(
            btn_panel,
            text="📋 Скопировать JSON",
            command=copy_to_clipboard,
            font=(self.font_sans, 9, "bold"),
            bg="#ffffff",
            fg="#1e293b",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT)

        tk.Button(
            btn_panel,
            text="📂 Открыть в редакторе",
            command=open_in_system_editor,
            font=(self.font_sans, 9),
            bg="#ffffff",
            fg="#1e293b",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(8, 0))

        tk.Button(
            btn_panel,
            text="Закрыть",
            command=dlg.destroy,
            font=(self.font_sans, 9, "bold"),
            bg="#0f172a",
            fg="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=14,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.RIGHT)

    def get_fitted_font_tuple(self, family, text, max_w, base_px, weight="normal", slant="roman"):
        """Рассчитывает кортеж шрифта Tkinter с гарантированным невыходом за max_w.

        В Tkinter отрицательный размер шрифта (-px) задает высоту строго в пикселях холста,
        полностью исключая искажения из-за системного DPI. Функция дополнительно измеряет
        фактическую ширину текста через tkinter.font.Font и, если текст шире допустимой границы,
        уменьшает размер шрифта до полного соответствия, предотвращая выход за поля листа.
        """
        px = max(6, int(base_px))
        try:
            import tkinter.font as tkfont
            f = tkfont.Font(family=family, size=-px, weight=weight, slant=slant)
            measured_w = f.measure(text)
            if measured_w > max_w and max_w > 10:
                ratio = max_w / float(measured_w)
                px = max(6, int(px * ratio * 0.96))
                f.configure(size=-px)
                while px > 6 and f.measure(text) > max_w:
                    px -= 1
                    f.configure(size=-px)
        except Exception:
            pass

        style = []
        if weight == "bold":
            style.append("bold")
        if slant == "italic":
            style.append("italic")
        if style:
            return (family, -px, " ".join(style))
        return (family, -px)

    def draw_cover_preview(self):
        pass

    def draw_title_page_preview(self):
        pass

    def draw_inside_cover_preview(self):
        pass

    def setup_ui(self):
        bg_main = "#f1f5f9"
        card_bg = "#ffffff"
        self.configure(bg=bg_main)

        # Верхняя панель (Header с заголовком и кнопками Горячие клавиши и config.json)
        header = tk.Frame(self, bg="#0f172a", padx=18, pady=12)
        header.pack(side=tk.TOP, fill=tk.X)

        header_left = tk.Frame(header, bg="#0f172a")
        header_left.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = tk.Label(
            header_left,
            text="Журнал учёта работы педагога: Обложка А4",
            font=(self.font_sans, 13, "bold"),
            fg="#ffffff",
            bg="#0f172a"
        )
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            header_left,
            text="Выпадающий список истории • Хранение в config.json • Левое поле 30 мм",
            font=(self.font_sans, 9),
            fg="#94a3b8",
            bg="#0f172a"
        )
        sub_lbl.pack(anchor="w", pady=(2, 0))

        # Кнопки в верхней части приложения: Открыть config.json и Горячие клавиши
        header_right = tk.Frame(header, bg="#0f172a")
        header_right.pack(side=tk.RIGHT, fill=tk.Y)

        btn_open_config = tk.Button(
            header_right,
            text="📂 config.json",
            command=self.open_config_file,
            font=(self.font_sans, 9, "bold"),
            bg="#064e3b",
            fg="#6ee7b7",
            activebackground="#065f46",
            activeforeground="#a7f3d0",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=5,
            cursor="hand2"
        )
        btn_open_config.pack(side=tk.LEFT, padx=(0, 8))

        btn_top_shortcuts = tk.Button(
            header_right,
            text="⌨ Горячие клавиши",
            command=self.show_shortcuts_help,
            font=(self.font_sans, 9, "bold"),
            bg="#1e293b",
            fg="#e2e8f0",
            activebackground="#334155",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=5,
            cursor="hand2"
        )
        btn_top_shortcuts.pack(side=tk.LEFT, padx=(0, 8))

        btn_check_updates = tk.Button(
            header_right,
            text=f"🔄 v{APP_VERSION}",
            command=lambda: self.check_updates_background(silent=False),
            font=(self.font_sans, 9, "bold"),
            bg="#1e1b4b",
            fg="#c7d2fe",
            activebackground="#312e81",
            activeforeground="#e0e7ff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=5,
            cursor="hand2"
        )
        btn_check_updates.pack(side=tk.LEFT)

        # Стили ttk для вкладок и выпадающего списка
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "TNotebook",
            background=bg_main,
            borderwidth=0
        )
        style.configure(
            "TNotebook.Tab",
            font=(self.font_sans, 8, "bold"),
            padding=[4, 3],
            background="#e2e8f0",
            foreground="#334155"
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#ffffff")],
            foreground=[("selected", "#0f172a")]
        )

        style.configure(
            "Large.TCombobox",
            font=(self.font_mono, 13, "bold"),
            padding=6
        )

        style.configure(
            "White.TCombobox",
            font=(self.font_sans, 10),
            fieldbackground="#ffffff",
            background="#ffffff",
            foreground="#0f172a",
            selectbackground="#ffffff",
            selectforeground="#0f172a",
            arrowcolor="#475569",
            lightcolor="#ffffff",
            darkcolor="#cbd5e1",
            bordercolor="#cbd5e1",
            padding=4
        )
        style.map(
            "White.TCombobox",
            fieldbackground=[
                ("readonly", "#ffffff"),
                ("focus", "#ffffff"),
                ("disabled", "#f8fafc"),
                ("!disabled", "#ffffff")
            ],
            foreground=[
                ("readonly", "#0f172a"),
                ("focus", "#0f172a"),
                ("!disabled", "#0f172a")
            ],
            selectbackground=[
                ("readonly", "#ffffff"),
                ("focus", "#ffffff"),
                ("!disabled", "#ffffff")
            ],
            selectforeground=[
                ("readonly", "#0f172a"),
                ("focus", "#0f172a"),
                ("!disabled", "#0f172a")
            ],
            background=[
                ("readonly", "#ffffff"),
                ("focus", "#ffffff"),
                ("!disabled", "#ffffff")
            ]
        )

        style.configure(
            "TCombobox",
            fieldbackground="#ffffff",
            background="#ffffff",
            foreground="#0f172a",
            selectbackground="#ffffff",
            selectforeground="#0f172a"
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#ffffff"), ("focus", "#ffffff")],
            selectbackground=[("readonly", "#ffffff"), ("focus", "#ffffff")],
            selectforeground=[("readonly", "#0f172a"), ("focus", "#0f172a")]
        )



        # Нижний статус-бар (пакуется до notebook, гарантируя постоянное закрепление внизу окна)
        status_bar = tk.Label(
            self,
            textvariable=self.status_var,
            font=(self.font_sans, 9),
            bg="#0f172a",
            fg="#e2e8f0",
            bd=0,
            anchor="w",
            padx=14,
            pady=6
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Контейнер с вкладками
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=16, pady=(8, 8))
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # =========================================================================
        # ВКЛАДКА 0: Обложка (параметры)
        # =========================================================================
        self.tab_input = tk.Frame(self.notebook, bg=card_bg, padx=24, pady=20)
        self.notebook.add(self.tab_input, text="Обложка")

        input_box = tk.LabelFrame(
            self.tab_input,
            text=" Учебный год ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=16,
            pady=14,
            relief=tk.SOLID,
            bd=1
        )
        input_box.pack(fill=tk.X, pady=(0, 16))

        tk.Label(
            input_box,
            text="Учебный год (например, 2024-2025):",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#1e293b"
        ).pack(anchor="w", pady=(0, 8))

        # Простое поле ввода без выпадающего списка (ttk.Entry)
        self.entry_year = ttk.Entry(
            input_box,
            textvariable=self.year_var,
            font=(self.font_mono, 12, "bold")
        )
        self.entry_year.pack(fill=tk.X, pady=(0, 10))
        self.combo_year = self.entry_year  # сохранение ссылки для совместимости focus_set

        # Привязка перехвата горячих клавиш (Ctrl+C, Ctrl+V, Ctrl+A, Ctrl+X, Ctrl+Z)
        self.entry_year.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        self.entry_year.bind("<Return>", lambda e: self.on_save_year_click())
        self.entry_year.bind("<FocusOut>", lambda e: self.on_save_year_click())
        self.attach_context_menu(self.entry_year)

        # Ряд кнопок для быстрого изменения учебного года ▼ −1 год и ▲ +1 год
        quick_btn_row = tk.Frame(input_box, bg=card_bg)
        quick_btn_row.pack(fill=tk.X, pady=(0, 10))

        btn_dec_cover_yr = tk.Button(
            quick_btn_row,
            text="▼  −1 год",
            command=lambda: self.change_academic_year(-1),
            font=(self.font_sans, 9, "bold"),
            bg="#f1f5f9",
            fg="#1e293b",
            activebackground="#e2e8f0",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=3,
            cursor="hand2"
        )
        btn_dec_cover_yr.pack(side=tk.LEFT, padx=(0, 6))

        btn_inc_cover_yr = tk.Button(
            quick_btn_row,
            text="▲  +1 год",
            command=lambda: self.change_academic_year(+1),
            font=(self.font_sans, 9, "bold"),
            bg="#f1f5f9",
            fg="#1e293b",
            activebackground="#e2e8f0",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=3,
            cursor="hand2"
        )
        btn_inc_cover_yr.pack(side=tk.LEFT)

        self.preview_lbl = None

        btn_next_to_title = tk.Button(
            self.tab_input,
            text="Далее: Титульный лист ➔",
            command=lambda: self.notebook.select(1),
            font=(self.font_sans, 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_next_to_title.pack(fill=tk.X, pady=(2, 10))

        # =========================================================================
        # ВКЛАДКА 1: Титульный лист (стр. 1 журнала)
        # =========================================================================
        self.tab_title_page = tk.Frame(self.notebook, bg=card_bg, padx=24, pady=20)
        self.notebook.add(self.tab_title_page, text="Титульный лист")

        title_card = tk.LabelFrame(
            self.tab_title_page,
            text=" Параметры титульного листа ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=16,
            pady=14,
            relief=tk.SOLID,
            bd=1
        )
        title_card.pack(fill=tk.X, pady=(0, 14))

        # 1. В самом верху: Название организации
        tk.Label(
            title_card,
            text="Название организации:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b"
        ).pack(anchor="w", pady=(0, 4))

        self.entry_org = ttk.Entry(
            title_card,
            textvariable=self.org_var,
            font=(self.font_sans, 10)
        )
        self.entry_org.pack(fill=tk.X, pady=(0, 12))
        self.combo_org = self.entry_org  # сохранение ссылки для совместимости
        self.entry_org.bind("<FocusOut>", self.on_org_focus_out)
        self.entry_org.bind("<Return>", self.on_org_focus_out)
        self.entry_org.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        self.attach_context_menu(self.entry_org)

        # 2. В середине: на гггг/гггг учебный год
        tk.Label(
            title_card,
            text="Учебный год (на титульный лист):",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b"
        ).pack(anchor="w", pady=(0, 4))

        year_row = tk.Frame(title_card, bg=card_bg)
        year_row.pack(fill=tk.X, pady=(0, 12))

        entry_title_yr = ttk.Entry(
            year_row,
            textvariable=self.title_year_var,
            font=(self.font_mono, 11, "bold")
        )
        entry_title_yr.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        entry_title_yr.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        entry_title_yr.bind("<Up>", lambda e: self.change_academic_year(+1))
        entry_title_yr.bind("<Down>", lambda e: self.change_academic_year(-1))
        entry_title_yr.bind("<FocusOut>", lambda e: self.sync_years_from_academic_year(self.title_year_var.get()))
        self.attach_context_menu(entry_title_yr)

        btn_dec_yr = tk.Button(
            year_row,
            text="▼  −1 год",
            command=lambda: self.change_academic_year(-1),
            font=(self.font_sans, 9, "bold"),
            bg="#f1f5f9",
            fg="#1e293b",
            activebackground="#e2e8f0",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=3,
            cursor="hand2"
        )
        btn_dec_yr.pack(side=tk.LEFT, padx=(0, 4))

        btn_inc_yr = tk.Button(
            year_row,
            text="▲  +1 год",
            command=lambda: self.change_academic_year(+1),
            font=(self.font_sans, 9, "bold"),
            bg="#f1f5f9",
            fg="#1e293b",
            activebackground="#e2e8f0",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=3,
            cursor="hand2"
        )
        btn_inc_yr.pack(side=tk.LEFT)

        # 3. Внизу: Сроки ведения журнала
        tk.Label(
            title_card,
            text="Сроки ведения журнала:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b"
        ).pack(anchor="w", pady=(0, 6))

        days_list = [""] + [f"{d:02d}" for d in range(1, 32)]
        months_list = [
            "января", "февраля", "марта", "апреля", "мая", "июня",
            "июля", "августа", "сентября", "октября", "ноября", "декабря"
        ]

        row_start = tk.Frame(title_card, bg=card_bg)
        row_start.pack(fill=tk.X, pady=(0, 8))

        tk.Label(row_start, text="Начат:", font=(self.font_sans, 9, "bold"), bg=card_bg, fg="#334155", width=8, anchor="w").pack(side=tk.LEFT)

        self.combo_sd = ttk.Combobox(
            row_start,
            textvariable=self.start_day_var,
            values=days_list,
            font=(self.font_sans, 10),
            style="White.TCombobox",
            width=4
        )
        self.combo_sd.pack(side=tk.LEFT, padx=(0, 6))
        self.combo_sd.bind("<<ComboboxSelected>>", self.on_day_selected)
        self.combo_sd.bind("<FocusOut>", self.on_day_selected)
        self.combo_sd.bind("<Return>", self.on_day_selected)
        self.combo_sd.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        self.attach_context_menu(self.combo_sd)

        self.combo_sm = ttk.Combobox(
            row_start,
            textvariable=self.start_month_var,
            values=months_list,
            font=(self.font_sans, 10),
            state="readonly",
            style="White.TCombobox",
            width=12
        )
        self.combo_sm.pack(side=tk.LEFT, padx=(0, 6))
        self.combo_sm.bind("<<ComboboxSelected>>", self.on_month_selected)

        self.entry_sy = ttk.Entry(
            row_start,
            textvariable=self.start_year_var,
            font=(self.font_mono, 10, "bold"),
            width=6
        )
        self.entry_sy.pack(side=tk.LEFT, padx=(0, 4))
        self.combo_sy = self.entry_sy
        self.entry_sy.bind("<FocusOut>", self.on_start_year_changed)
        self.entry_sy.bind("<Return>", self.on_start_year_changed)
        self.entry_sy.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        self.attach_context_menu(self.entry_sy)
        tk.Label(row_start, text="г.", font=(self.font_sans, 9, "bold"), bg=card_bg, fg="#334155").pack(side=tk.LEFT)

        row_end = tk.Frame(title_card, bg=card_bg)
        row_end.pack(fill=tk.X, pady=(0, 8))

        tk.Label(row_end, text="Окончен:", font=(self.font_sans, 9, "bold"), bg=card_bg, fg="#334155", width=8, anchor="w").pack(side=tk.LEFT)

        self.combo_ed = ttk.Combobox(
            row_end,
            textvariable=self.end_day_var,
            values=days_list,
            font=(self.font_sans, 10),
            style="White.TCombobox",
            width=4
        )
        self.combo_ed.pack(side=tk.LEFT, padx=(0, 6))
        self.combo_ed.bind("<<ComboboxSelected>>", self.on_day_selected)
        self.combo_ed.bind("<FocusOut>", self.on_day_selected)
        self.combo_ed.bind("<Return>", self.on_day_selected)
        self.combo_ed.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        self.attach_context_menu(self.combo_ed)

        self.combo_em = ttk.Combobox(
            row_end,
            textvariable=self.end_month_var,
            values=months_list,
            font=(self.font_sans, 10),
            state="readonly",
            style="White.TCombobox",
            width=12
        )
        self.combo_em.pack(side=tk.LEFT, padx=(0, 6))
        self.combo_em.bind("<<ComboboxSelected>>", self.on_month_selected)

        self.entry_ey = ttk.Entry(
            row_end,
            textvariable=self.end_year_var,
            font=(self.font_mono, 10, "bold"),
            width=6
        )
        self.entry_ey.pack(side=tk.LEFT, padx=(0, 4))
        self.combo_ey = self.entry_ey
        self.entry_ey.bind("<FocusOut>", self.on_end_year_changed)
        self.entry_ey.bind("<Return>", self.on_end_year_changed)
        self.entry_ey.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        self.attach_context_menu(self.entry_ey)
        tk.Label(row_end, text="г.", font=(self.font_sans, 9, "bold"), bg=card_bg, fg="#334155").pack(side=tk.LEFT)

        btn_next_to_inside = tk.Button(
            self.tab_title_page,
            text="Далее: Оборот титульного листа ➔",
            command=lambda: self.notebook.select(2),
            font=(self.font_sans, 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_next_to_inside.pack(fill=tk.X, pady=(4, 10))

        # =========================================================================
        # ВКЛАДКА 2: Оборот титульного листа (стр. 2 журнала)
        # =========================================================================
        self.tab_inside_cover = tk.Frame(self.notebook, bg=card_bg, padx=24, pady=20)
        self.notebook.add(self.tab_inside_cover, text="Оборот титула")

        teacher_card = tk.LabelFrame(
            self.tab_inside_cover,
            text=" Параметры оборота титульного листа ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=16,
            pady=14,
            relief=tk.SOLID,
            bd=1
        )
        teacher_card.pack(fill=tk.X, pady=(0, 14))

        lbl_teacher_title = tk.Label(
            teacher_card,
            text="ИНПУТ: ФИО ПЕДАГОГА",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#2563eb"
        )
        lbl_teacher_title.pack(anchor="w", pady=(0, 4))

        tk.Label(
            teacher_card,
            text="Фамилия, Имя, Отчество педагога (на оборот титульного листа):",
            font=(self.font_sans, 9),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 6))

        # Простое поле ввода без выпадающего списка (ttk.Entry)
        self.entry_teacher = ttk.Entry(
            teacher_card,
            textvariable=self.teacher_var,
            font=(self.font_sans, 11)
        )
        self.entry_teacher.pack(fill=tk.X, pady=(0, 8))
        self.combo_teacher = self.entry_teacher  # сохранение ссылки для совместимости
        self.entry_teacher.bind("<Return>", lambda e: self.on_save_teacher_click())
        self.entry_teacher.bind("<FocusOut>", lambda e: self.on_save_teacher_click())
        self.entry_teacher.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
        self.attach_context_menu(self.entry_teacher)

        tk.Label(
            teacher_card,
            text="💡 Значение автоматически сохраняется и подставляется в строку «Педагог» на обороте титульного листа.",
            font=(self.font_sans, 8),
            bg=card_bg,
            fg="#64748b",
            justify="left"
        ).pack(anchor="w", pady=(0, 4))

        btn_next_to_main = tk.Button(
            self.tab_inside_cover,
            text="Далее: Основные данные ➔",
            command=lambda: self.notebook.select(3),
            font=(self.font_sans, 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_next_to_main.pack(fill=tk.X, pady=(4, 12))

        # =========================================================================
        # ВКЛАДКА 3: Основные данные (стр. 3 журнала)
        # =========================================================================
        self.tab_main_data = tk.Frame(self.notebook, bg=card_bg)
        self.notebook.add(self.tab_main_data, text="Основные данные")

        canvas_main = tk.Canvas(self.tab_main_data, bg=card_bg, highlightthickness=0)
        vbar_main = ttk.Scrollbar(self.tab_main_data, orient="vertical", command=canvas_main.yview)
        scroll_content = tk.Frame(canvas_main, bg=card_bg, padx=18, pady=14)

        scroll_window_id = canvas_main.create_window((0, 0), window=scroll_content, anchor="nw")

        def _on_scroll_configure(event):
            canvas_main.configure(scrollregion=canvas_main.bbox("all"))

        def _on_canvas_configure(event):
            canvas_main.itemconfig(scroll_window_id, width=event.width)

        scroll_content.bind("<Configure>", _on_scroll_configure)
        canvas_main.bind("<Configure>", _on_canvas_configure)
        canvas_main.configure(yscrollcommand=vbar_main.set)

        canvas_main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vbar_main.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel(event):
            if canvas_main.winfo_exists():
                canvas_main.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas_main.bind("<MouseWheel>", _on_mousewheel)
        scroll_content.bind("<MouseWheel>", _on_mousewheel)

        # -------------------------------------------------------------
        # БЛОК 1: Реквизиты объединения и группы
        # -------------------------------------------------------------
        card_org_info = tk.LabelFrame(
            scroll_content,
            text=" 1. Реквизиты объединения и группы (стр. 3 журнала) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=14,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        card_org_info.pack(fill=tk.X, pady=(0, 12))

        # 1.1 Наименование организации
        tk.Label(
            card_org_info,
            text="Наименование организации:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))

        e_main_org = ttk.Entry(card_org_info, textvariable=self.main_org_var, font=(self.font_sans, 10))
        e_main_org.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_main_org)
        e_main_org.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # 1.2 Отдел
        tk.Label(
            card_org_info,
            text="Отдел:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_dept = ttk.Entry(card_org_info, textvariable=self.department_var, font=(self.font_sans, 10))
        e_dept.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_dept)
        e_dept.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # 1.3 Объединение
        tk.Label(
            card_org_info,
            text="Объединение (название):",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_assoc = ttk.Entry(card_org_info, textvariable=self.association_var, font=(self.font_sans, 10))
        e_assoc.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_assoc)
        e_assoc.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # 1.4 Группа и Год обучения (в одну строку)
        f_grp_yr = tk.Frame(card_org_info, bg=card_bg)
        f_grp_yr.pack(fill=tk.X, pady=(0, 4))

        col_grp = tk.Frame(f_grp_yr, bg=card_bg)
        col_grp.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Label(
            col_grp,
            text="Группа:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_grp = ttk.Entry(col_grp, textvariable=self.group_var, font=(self.font_sans, 10))
        e_grp.pack(fill=tk.X)
        self.attach_context_menu(e_grp)
        e_grp.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        col_yr = tk.Frame(f_grp_yr, bg=card_bg)
        col_yr.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
        tk.Label(
            col_yr,
            text="Год обучения (например, «1-й год»):",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_yr = ttk.Entry(col_yr, textvariable=self.study_year_var, font=(self.font_sans, 10))
        e_yr.pack(fill=tk.X)
        self.attach_context_menu(e_yr)
        e_yr.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # -------------------------------------------------------------
        # БЛОК 2: Расписание занятий (6 строк и 2 столбца)
        # -------------------------------------------------------------
        card_sched = tk.LabelFrame(
            scroll_content,
            text=" 2. РАСПИСАНИЕ ЗАНЯТИЙ (6 строк и 2 столбца: дни недели и время) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=14,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        card_sched.pack(fill=tk.X, pady=(0, 12))

        f_sched_hdr = tk.Frame(card_sched, bg="#f1f5f9", padx=6, pady=4, relief=tk.SOLID, bd=1)
        f_sched_hdr.pack(fill=tk.X, pady=(0, 4))
        lbl_h_day = tk.Label(f_sched_hdr, text="Дни недели", font=(self.font_sans, 9, "bold"), bg="#f1f5f9", fg="#1e293b", width=22, anchor="w")
        lbl_h_day.pack(side=tk.LEFT, padx=(0, 6))
        lbl_h_time = tk.Label(f_sched_hdr, text="Время (часы)", font=(self.font_sans, 9, "bold"), bg="#f1f5f9", fg="#1e293b", anchor="w")
        lbl_h_time.pack(side=tk.LEFT, fill=tk.X, expand=True)

        for idx, (dv, tv) in enumerate(self.schedule_row_vars, start=1):
            row_sched = tk.Frame(card_sched, bg=card_bg)
            row_sched.pack(fill=tk.X, pady=2)

            tk.Label(row_sched, text=f"{idx}.", font=(self.font_sans, 8, "bold"), bg=card_bg, fg="#64748b", width=2).pack(side=tk.LEFT, padx=(0, 4))

            e_d = ttk.Entry(row_sched, textvariable=dv, font=(self.font_sans, 9), width=24)
            e_d.pack(side=tk.LEFT, padx=(0, 6))
            self.attach_context_menu(e_d)
            e_d.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

            e_t = ttk.Entry(row_sched, textvariable=tv, font=(self.font_sans, 9))
            e_t.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.attach_context_menu(e_t)
            e_t.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # -------------------------------------------------------------
        # БЛОК 3: Изменение расписания занятий (динамическая таблица)
        # -------------------------------------------------------------
        card_changes = tk.LabelFrame(
            scroll_content,
            text=" 3. ИЗМЕНЕНИЕ РАСПИСАНИЯ ЗАНЯТИЙ (динамическая таблица) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=14,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        card_changes.pack(fill=tk.X, pady=(0, 12))

        f_chg_hdr = tk.Frame(card_changes, bg="#f1f5f9", padx=6, pady=4, relief=tk.SOLID, bd=1)
        f_chg_hdr.pack(fill=tk.X, pady=(0, 4))
        tk.Label(f_chg_hdr, text="Дни недели", font=(self.font_sans, 9, "bold"), bg="#f1f5f9", fg="#1e293b", width=22, anchor="w").pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(f_chg_hdr, text="Время (часы)", font=(self.font_sans, 9, "bold"), bg="#f1f5f9", fg="#1e293b", anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(f_chg_hdr, text="Удалить", font=(self.font_sans, 8), bg="#f1f5f9", fg="#64748b", width=6).pack(side=tk.RIGHT)

        self.changes_rows_frame = tk.Frame(card_changes, bg=card_bg)
        self.changes_rows_frame.pack(fill=tk.X, pady=(0, 6))

        # Заполнение строк изменений из сохраненной конфигурации
        for item in self.saved_sched_changes:
            if isinstance(item, dict):
                self.add_schedule_change_row(item.get("day", ""), item.get("time", ""), auto_save=False)

        btn_add_chg_row = tk.Button(
            card_changes,
            text="+ Добавить новую строку в изменение расписания",
            command=lambda: self.add_schedule_change_row("", "", auto_save=True),
            font=(self.font_sans, 9, "bold"),
            bg="#f0fdf4",
            fg="#166534",
            activebackground="#dcfce7",
            activeforeground="#14532d",
            relief=tk.SOLID,
            bd=1,
            pady=6,
            cursor="hand2"
        )
        btn_add_chg_row.pack(fill=tk.X, pady=(4, 2))

        # -------------------------------------------------------------
        # БЛОК 4: Руководитель, староста, концертмейстер
        # -------------------------------------------------------------
        card_staff = tk.LabelFrame(
            scroll_content,
            text=" 4. Руководитель, староста и аккомпаниатор (концертмейстер) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=14,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        card_staff.pack(fill=tk.X, pady=(0, 12))

        # 4.1 Руководитель
        tk.Label(
            card_staff,
            text="РУКОВОДИТЕЛЬ (фамилия, имя, отчество полностью):",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))

        e_ruk = ttk.Entry(card_staff, textvariable=self.rukovoditel_var, font=(self.font_sans, 10))
        e_ruk.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_ruk)
        e_ruk.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # 4.2 Староста
        tk.Label(
            card_staff,
            text="СТАРОСТА:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_star = ttk.Entry(card_staff, textvariable=self.starosta_var, font=(self.font_sans, 10))
        e_star.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_star)
        e_star.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # 4.3 Аккомпаниатор (Концертмейстер)
        tk.Label(
            card_staff,
            text="АККОМПАНИАТОР (КОНЦЕРТМЕЙСТЕР):",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_acc = ttk.Entry(card_staff, textvariable=self.accompanist_var, font=(self.font_sans, 10))
        e_acc.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_acc)
        e_acc.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # 4.4 Расписание работы аккомпаниатора
        tk.Label(
            card_staff,
            text="Расписание работы аккомпаниатора (концертмейстера):",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_acc_s = ttk.Entry(card_staff, textvariable=self.acc_schedule_var, font=(self.font_sans, 10))
        e_acc_s.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_acc_s)
        e_acc_s.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # 4.5 Изменение расписания работы аккомпаниатора
        tk.Label(
            card_staff,
            text="Изменение расписания работы аккомпаниатора (концертмейстера):",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(anchor="w", pady=(0, 2))
        e_acc_c = ttk.Entry(card_staff, textvariable=self.acc_changes_var, font=(self.font_sans, 10))
        e_acc_c.pack(fill=tk.X, pady=(0, 8))
        self.attach_context_menu(e_acc_c)
        e_acc_c.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        # -------------------------------------------------------------
        # Кнопки действий вкладки Основные данные
        # -------------------------------------------------------------
        btn_next_to_export = tk.Button(
            scroll_content,
            text="Далее: Сентябрь ➔",
            command=lambda: self.notebook.select(4),
            font=(self.font_sans, 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_next_to_export.pack(fill=tk.X, pady=(6, 12))

        # =========================================================================
        # ВКЛАДКИ 5..16: 12 Месяцев журнала (Сентябрь — Август, стр. 4..27 журнала)
        # =========================================================================
        for idx, m in enumerate(MONTHS_CONFIG):
            m_key = m["key"]
            m_name = m["name"]
            next_tab = (idx + 5) if idx < len(MONTHS_CONFIG) - 1 else 16

            tab_month = tk.Frame(self.notebook, bg=card_bg, padx=12, pady=10)
            self.notebook.add(tab_month, text=m_name)
            if m_key == "september":
                self.setup_month_tab(tab_month, card_bg, m, next_tab_idx=next_tab)
                self.tab_september = tab_month
                tab_month._is_built = True
            else:
                tab_month._is_built = False
                tab_month._build_func = (lambda f=tab_month, info=m, n=next_tab: self.setup_month_tab(f, card_bg, info, next_tab_idx=n))

        # =========================================================================
        # ВКЛАДКА 16: Учёт массовых мероприятий с обучающимися (стр. 30 и 31 журнала)
        # =========================================================================
        self.tab_mass_events = tk.Frame(self.notebook, bg=card_bg, padx=12, pady=10)
        self.notebook.add(self.tab_mass_events, text="Мероприятия")
        self.tab_mass_events._is_built = False
        self.tab_mass_events._build_func = lambda: self.setup_mass_events_tab(self.tab_mass_events, card_bg, next_tab_idx=17)

        # =========================================================================
        # ВКЛАДКА 17: Творческие достижения обучающихся (стр. 32 и 33 журнала)
        # =========================================================================
        self.tab_creative = tk.Frame(self.notebook, bg=card_bg, padx=12, pady=10)
        self.notebook.add(self.tab_creative, text="Творческие достижения")
        self.tab_creative._is_built = False
        self.tab_creative._build_func = lambda: self.setup_creative_achievements_tab(self.tab_creative, card_bg, next_tab_idx=18)

        # =========================================================================
        # ВКЛАДКА 18: Список обучающихся (стр. 34-39 журнала, 6 страниц)
        # =========================================================================
        self.tab_students_list = tk.Frame(self.notebook, bg=card_bg, padx=12, pady=10)
        self.notebook.add(self.tab_students_list, text="Список обучающихся")
        self.tab_students_list._is_built = False
        self.tab_students_list._build_func = lambda: self.setup_students_list_tab(self.tab_students_list, card_bg, next_tab_idx=19)

        # =========================================================================
        # ВКЛАДКА 19: Список обучающихся, прошедших инструктаж по ТБ (2 страницы)
        # =========================================================================
        self.tab_safety = tk.Frame(self.notebook, bg=card_bg, padx=12, pady=10)
        self.notebook.add(self.tab_safety, text="Инструктаж по ТБ")
        self.tab_safety._is_built = False
        self.tab_safety._build_func = lambda: self.setup_safety_briefing_tab(self.tab_safety, card_bg, next_tab_idx=20)

        # =========================================================================
        # ВКЛАДКА 20: Годовой цифровой отчёт (стр. 40 журнала, 1 страница)
        # =========================================================================
        self.tab_annual_report = tk.Frame(self.notebook, bg=card_bg)
        self.notebook.add(self.tab_annual_report, text="Годовой отчёт")
        self.setup_annual_report_tab(self.tab_annual_report, card_bg, next_tab_idx=21)

        # =========================================================================
        # ВКЛАДКА 21: Подсчёт отработанного времени (учёт часов за год)
        # =========================================================================
        self.tab_work_hours = tk.Frame(self.notebook, bg=card_bg)
        self.notebook.add(self.tab_work_hours, text="Отработанное время")
        self.tab_work_hours._is_built = False
        self.tab_work_hours._build_func = lambda: self.setup_work_hours_tab(self.tab_work_hours, card_bg, next_tab_idx=22)

        # =========================================================================
        # ВКЛАДКА 22: Формирование Excel файла (ВСЕГДА ПОСЛЕДНЯЯ ВКЛАДКА)
        # =========================================================================
        self.tab_export = tk.Frame(self.notebook, bg=card_bg)
        self.notebook.add(self.tab_export, text="Формирование Excel")

        canvas_export = tk.Canvas(self.tab_export, bg=card_bg, highlightthickness=0)
        vbar_export = ttk.Scrollbar(self.tab_export, orient="vertical", command=canvas_export.yview)
        scroll_export_content = tk.Frame(canvas_export, bg=card_bg, padx=18, pady=14)

        scroll_export_window = canvas_export.create_window((0, 0), window=scroll_export_content, anchor="nw")

        def _on_export_scroll_cfg(event):
            canvas_export.configure(scrollregion=canvas_export.bbox("all"))

        def _on_export_canvas_cfg(event):
            canvas_export.itemconfig(scroll_export_window, width=event.width)

        scroll_export_content.bind("<Configure>", _on_export_scroll_cfg)
        canvas_export.bind("<Configure>", _on_export_canvas_cfg)
        canvas_export.configure(yscrollcommand=vbar_export.set)

        canvas_export.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vbar_export.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_export_mousewheel(event):
            if canvas_export.winfo_exists():
                if sys.platform == "darwin":
                    canvas_export.yview_scroll(int(-1 * event.delta), "units")
                else:
                    canvas_export.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas_export.bind("<Enter>", lambda e: canvas_export.bind_all("<MouseWheel>", _on_export_mousewheel))
        canvas_export.bind("<Leave>", lambda e: canvas_export.unbind_all("<MouseWheel>"))

        self.summary_lbl = None

        # Блок экспорта в Excel
        export_box = tk.LabelFrame(
            scroll_export_content,
            text=" Экспорт в Excel (.xlsx) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=16,
            pady=14,
            relief=tk.SOLID,
            bd=1
        )
        export_box.pack(fill=tk.X, pady=(0, 12))

        # Сетка 1: Экспорт Обложки, Титульного листа, Оборота и Основных данных
        export_grid = tk.Frame(export_box, bg=card_bg)
        export_grid.pack(fill=tk.X, pady=(0, 8))

        # 1. Кнопка Обложки
        btn_col1 = tk.Frame(export_grid, bg=card_bg)
        btn_col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        btn_generate = tk.Button(
            btn_col1,
            text="💾 Обложка",
            command=self.on_generate_click,
            font=(self.font_sans, 9, "bold"),
            bg="#15803d",
            fg="#ffffff",
            activebackground="#166534",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_generate.pack(fill=tk.X, pady=(0, 4))

        self.btn_open = tk.Button(
            btn_col1,
            text="📂 Открыть Обложку",
            command=self.on_open_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open.pack(fill=tk.X)

        # 2. Кнопка Титульного листа
        btn_col2 = tk.Frame(export_grid, bg=card_bg)
        btn_col2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 2))

        btn_gen_title = tk.Button(
            btn_col2,
            text="💾 Титульный лист",
            command=self.on_generate_title_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_title.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_title = tk.Button(
            btn_col2,
            text="📂 Открыть Титульный",
            command=self.on_open_title_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_title.pack(fill=tk.X)

        # 3. Кнопка Оборота титульного листа
        btn_col3 = tk.Frame(export_grid, bg=card_bg)
        btn_col3.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 2))

        btn_gen_inside = tk.Button(
            btn_col3,
            text="💾 Оборот титульного",
            command=self.on_generate_inside_cover_click,
            font=(self.font_sans, 9, "bold"),
            bg="#1d4ed8",
            fg="#ffffff",
            activebackground="#1e40af",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_inside.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_inside = tk.Button(
            btn_col3,
            text="📂 Открыть Оборот",
            command=self.on_open_inside_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_inside.pack(fill=tk.X)

        # 4. Кнопка Основных данных (стр. 3)
        btn_col4 = tk.Frame(export_grid, bg=card_bg)
        btn_col4.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        btn_gen_main = tk.Button(
            btn_col4,
            text="💾 Основные данные",
            command=self.on_generate_main_data_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0d9488",
            fg="#ffffff",
            activebackground="#0f766e",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_main.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_main = tk.Button(
            btn_col4,
            text="📂 Открыть Данные",
            command=self.on_open_main_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_main.pack(fill=tk.X)

        # Сетка 2: Экспорт страниц любого месяца (выбор месяца, посещаемость, темы ДОП, разворот)
        month_export_frame = tk.LabelFrame(
            export_box,
            text=" Экспорт страниц месяца (посещаемость, темы ДОП, разворот) ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        month_export_frame.pack(fill=tk.X, pady=(0, 10))

        # Строка выбора месяца
        m_sel_row = tk.Frame(month_export_frame, bg=card_bg)
        m_sel_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            m_sel_row,
            text="Выберите месяц для формирования:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155"
        ).pack(side=tk.LEFT, padx=(0, 8))

        month_names = [m["name"] for m in MONTHS_CONFIG]
        self.combo_export_month = ttk.Combobox(
            m_sel_row,
            textvariable=self.export_month_var,
            values=month_names,
            state="readonly",
            width=18,
            font=(self.font_sans, 9, "bold")
        )
        self.combo_export_month.pack(side=tk.LEFT, padx=(0, 12))

        # 3 кнопки формирования для выбранного месяца
        export_month_grid = tk.Frame(month_export_frame, bg=card_bg)
        export_month_grid.pack(fill=tk.X, pady=(0, 4))

        # 1. Посещаемость (стр. 1 месяца)
        btn_col_m1 = tk.Frame(export_month_grid, bg=card_bg)
        btn_col_m1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        btn_gen_month_p1 = tk.Button(
            btn_col_m1,
            text="💾 Посещаемость (стр. 1)",
            command=lambda: self.on_generate_month_p1_click(),
            font=(self.font_sans, 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_month_p1.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_month_p1 = tk.Button(
            btn_col_m1,
            text="📂 Открыть Посещаемость",
            command=lambda: self.on_open_month_p1_file_click(),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_month_p1.pack(fill=tk.X)
        self.btn_open_sept_p1 = self.btn_open_month_p1

        # 2. Темы ДОП (стр. 2 месяца)
        btn_col_m2 = tk.Frame(export_month_grid, bg=card_bg)
        btn_col_m2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 2))

        btn_gen_month_p2 = tk.Button(
            btn_col_m2,
            text="💾 Темы ДОП (стр. 2)",
            command=lambda: self.on_generate_month_p2_click(),
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_month_p2.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_month_p2 = tk.Button(
            btn_col_m2,
            text="📂 Открыть Темы ДОП",
            command=lambda: self.on_open_month_p2_file_click(),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_month_p2.pack(fill=tk.X)
        self.btn_open_sept_p2 = self.btn_open_month_p2

        # 3. Разворот (стр. 1 и 2 месяца)
        btn_col_m_spread = tk.Frame(export_month_grid, bg=card_bg)
        btn_col_m_spread.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        btn_gen_month_spread = tk.Button(
            btn_col_m_spread,
            text="💾 Разворот месяца (2 стр.)",
            command=lambda: self.on_generate_month_spread_click(),
            font=(self.font_sans, 9, "bold"),
            bg="#4f46e5",
            fg="#ffffff",
            activebackground="#4338ca",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_month_spread.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_month_spread = tk.Button(
            btn_col_m_spread,
            text="📂 Открыть Разворот",
            command=lambda: self.on_open_month_spread_file_click(),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_month_spread.pack(fill=tk.X)
        self.btn_open_sept_spread = self.btn_open_month_spread

        # Сетка 3: Экспорт Учёта массовых мероприятий (стр. 30, стр. 31, разворот)
        events_export_frame = tk.LabelFrame(
            export_box,
            text=" Экспорт страниц «Учёт массовых мероприятий» (стр. 30 и 31) ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        events_export_frame.pack(fill=tk.X, pady=(0, 10))

        export_events_grid = tk.Frame(events_export_frame, bg=card_bg)
        export_events_grid.pack(fill=tk.X, pady=(0, 4))

        # 1. Мероприятия (стр. 1 - стр. 30 журнала)
        btn_col_ev1 = tk.Frame(export_events_grid, bg=card_bg)
        btn_col_ev1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        btn_gen_ev1 = tk.Button(
            btn_col_ev1,
            text="💾 Мероприятия (стр. 30)",
            command=self.on_generate_events_p1_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0891b2",
            fg="#ffffff",
            activebackground="#0e7490",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_ev1.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_events_p1 = tk.Button(
            btn_col_ev1,
            text="📂 Открыть стр. 30",
            command=self.on_open_events_p1_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_events_p1.pack(fill=tk.X)

        # 2. Мероприятия (стр. 2 - стр. 31 журнала)
        btn_col_ev2 = tk.Frame(export_events_grid, bg=card_bg)
        btn_col_ev2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 2))

        btn_gen_ev2 = tk.Button(
            btn_col_ev2,
            text="💾 Мероприятия (стр. 31)",
            command=self.on_generate_events_p2_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_ev2.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_events_p2 = tk.Button(
            btn_col_ev2,
            text="📂 Открыть стр. 31",
            command=self.on_open_events_p2_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_events_p2.pack(fill=tk.X)

        # 3. Разворот мероприятий (стр. 30 и 31 журнала)
        btn_col_ev_spread = tk.Frame(export_events_grid, bg=card_bg)
        btn_col_ev_spread.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        btn_gen_ev_spread = tk.Button(
            btn_col_ev_spread,
            text="💾 Разворот (стр. 30-31)",
            command=self.on_generate_events_spread_click,
            font=(self.font_sans, 9, "bold"),
            bg="#4f46e5",
            fg="#ffffff",
            activebackground="#4338ca",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_ev_spread.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_events_spread = tk.Button(
            btn_col_ev_spread,
            text="📂 Открыть Разворот",
            command=self.on_open_events_spread_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_events_spread.pack(fill=tk.X)

        # Сетка 4: Экспорт страниц «Творческие достижения» (стр. 32, стр. 33, разворот)
        creative_export_frame = tk.LabelFrame(
            export_box,
            text=" Экспорт страниц «Творческие достижения» (стр. 32 и 33) ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        creative_export_frame.pack(fill=tk.X, pady=(0, 10))

        export_creative_grid = tk.Frame(creative_export_frame, bg=card_bg)
        export_creative_grid.pack(fill=tk.X, pady=(0, 4))

        # 1. Творческие достижения (стр. 1 - стр. 32 журнала)
        btn_col_cr1 = tk.Frame(export_creative_grid, bg=card_bg)
        btn_col_cr1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        btn_gen_cr1 = tk.Button(
            btn_col_cr1,
            text="💾 Достижения (стр. 32)",
            command=self.on_generate_creative_p1_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0d9488",
            fg="#ffffff",
            activebackground="#0f766e",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_cr1.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_creative_p1 = tk.Button(
            btn_col_cr1,
            text="📂 Открыть стр. 32",
            command=self.on_open_creative_p1_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_creative_p1.pack(fill=tk.X)

        # 2. Обучающихся (стр. 2 - стр. 33 журнала)
        btn_col_cr2 = tk.Frame(export_creative_grid, bg=card_bg)
        btn_col_cr2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 2))

        btn_gen_cr2 = tk.Button(
            btn_col_cr2,
            text="💾 Достижения (стр. 33)",
            command=self.on_generate_creative_p2_click,
            font=(self.font_sans, 9, "bold"),
            bg="#14b8a6",
            fg="#ffffff",
            activebackground="#0d9488",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_cr2.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_creative_p2 = tk.Button(
            btn_col_cr2,
            text="📂 Открыть стр. 33",
            command=self.on_open_creative_p2_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_creative_p2.pack(fill=tk.X)

        # 3. Разворот достижений (стр. 32 и 33 журнала)
        btn_col_cr_spread = tk.Frame(export_creative_grid, bg=card_bg)
        btn_col_cr_spread.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        btn_gen_cr_spread = tk.Button(
            btn_col_cr_spread,
            text="💾 Разворот (стр. 32-33)",
            command=self.on_generate_creative_spread_click,
            font=(self.font_sans, 9, "bold"),
            bg="#4f46e5",
            fg="#ffffff",
            activebackground="#4338ca",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_cr_spread.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_creative_spread = tk.Button(
            btn_col_cr_spread,
            text="📂 Открыть Разворот",
            command=self.on_open_creative_spread_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_creative_spread.pack(fill=tk.X)

        # Сетка 5: Экспорт страниц «Список обучающихся» (стр. 34–39: 3 разворота, 6 страниц)
        students_export_frame = tk.LabelFrame(
            export_box,
            text=" Экспорт страниц «Список обучающихся» (стр. 34–39: 3 разворота, 6 страниц) ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        students_export_frame.pack(fill=tk.X, pady=(0, 10))

        export_students_grid = tk.Frame(students_export_frame, bg=card_bg)
        export_students_grid.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_students_spreads = {}

        # 1. Разворот 1 (стр. 34 и 35 журнала, ученики 1-10)
        btn_col_st_sp1 = tk.Frame(export_students_grid, bg=card_bg)
        btn_col_st_sp1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 3))

        btn_gen_st_sp1 = tk.Button(
            btn_col_st_sp1,
            text="💾 Разворот 1 (стр. 34-35)",
            command=lambda: self.on_generate_students_list_spread_click(1),
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_st_sp1.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_students_spreads[1] = tk.Button(
            btn_col_st_sp1,
            text="📂 Открыть Разворот 1",
            command=lambda: self.on_open_students_list_spread_file_click(1),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_students_spreads[1].pack(fill=tk.X)

        # 2. Разворот 2 (стр. 36 и 37 журнала, ученики 11-20)
        btn_col_st_sp2 = tk.Frame(export_students_grid, bg=card_bg)
        btn_col_st_sp2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 3))

        btn_gen_st_sp2 = tk.Button(
            btn_col_st_sp2,
            text="💾 Разворот 2 (стр. 36-37)",
            command=lambda: self.on_generate_students_list_spread_click(2),
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_st_sp2.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_students_spreads[2] = tk.Button(
            btn_col_st_sp2,
            text="📂 Открыть Разворот 2",
            command=lambda: self.on_open_students_list_spread_file_click(2),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_students_spreads[2].pack(fill=tk.X)

        # 3. Разворот 3 (стр. 38 и 39 журнала, ученики 21-30)
        btn_col_st_sp3 = tk.Frame(export_students_grid, bg=card_bg)
        btn_col_st_sp3.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 3))

        btn_gen_st_sp3 = tk.Button(
            btn_col_st_sp3,
            text="💾 Разворот 3 (стр. 38-39)",
            command=lambda: self.on_generate_students_list_spread_click(3),
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_st_sp3.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_students_spreads[3] = tk.Button(
            btn_col_st_sp3,
            text="📂 Открыть Разворот 3",
            command=lambda: self.on_open_students_list_spread_file_click(3),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_students_spreads[3].pack(fill=tk.X)

        # 4. Все 6 страниц списка обучающихся в один файл Excel
        btn_col_st_all = tk.Frame(export_students_grid, bg=card_bg)
        btn_col_st_all.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 0))

        btn_gen_st_all = tk.Button(
            btn_col_st_all,
            text="💾 Все 6 стр. списка (.xlsx)",
            command=self.on_generate_students_list_all_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0f766e",
            fg="#ffffff",
            activebackground="#115e59",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_st_all.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_students_all = tk.Button(
            btn_col_st_all,
            text="📂 Открыть 6 страниц",
            command=self.on_open_students_list_all_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_students_all.pack(fill=tk.X)

        # Сетка 6: Экспорт «Инструктаж по технике безопасности» (стр. 38 и 39 журнала: 2 страницы)
        safety_export_frame = tk.LabelFrame(
            export_box,
            text=" Экспорт «Инструктаж по технике безопасности» (стр. 38 и 39 журнала: 2 страницы) ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        safety_export_frame.pack(fill=tk.X, pady=(0, 10))

        export_safety_grid = tk.Frame(safety_export_frame, bg=card_bg)
        export_safety_grid.pack(fill=tk.X, pady=(0, 4))

        # 1. Страница 1 (стр. 38)
        btn_col_sb1 = tk.Frame(export_safety_grid, bg=card_bg)
        btn_col_sb1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 3))

        btn_gen_sb1 = tk.Button(
            btn_col_sb1,
            text="💾 Инструктаж (стр. 38)",
            command=lambda: self.on_generate_safety_briefing_page_click(1),
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_sb1.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_safety_p1 = tk.Button(
            btn_col_sb1,
            text="📂 Открыть стр. 38",
            command=lambda: self.on_open_safety_briefing_page_file_click(1),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_safety_p1.pack(fill=tk.X)

        # 2. Страница 2 (стр. 39)
        btn_col_sb2 = tk.Frame(export_safety_grid, bg=card_bg)
        btn_col_sb2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 3))

        btn_gen_sb2 = tk.Button(
            btn_col_sb2,
            text="💾 Инструктаж (стр. 39)",
            command=lambda: self.on_generate_safety_briefing_page_click(2),
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_sb2.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_safety_p2 = tk.Button(
            btn_col_sb2,
            text="📂 Открыть стр. 39",
            command=lambda: self.on_open_safety_briefing_page_file_click(2),
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_safety_p2.pack(fill=tk.X)

        # 3. Разворот ТБ (стр. 38 и 39 в один .xlsx)
        btn_col_sb_spread = tk.Frame(export_safety_grid, bg=card_bg)
        btn_col_sb_spread.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 0))

        btn_gen_sb_spread = tk.Button(
            btn_col_sb_spread,
            text="💾 Разворот ТБ (стр. 38-39)",
            command=self.on_generate_safety_briefing_spread_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0f766e",
            fg="#ffffff",
            activebackground="#115e59",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_sb_spread.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_safety_spread = tk.Button(
            btn_col_sb_spread,
            text="📂 Открыть Разворот ТБ",
            command=self.on_open_safety_briefing_spread_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_safety_spread.pack(fill=tk.X)

        # Сетка 7: Экспорт «Годовой цифровой отчёт» (стр. 40 журнала: 1 страница)
        ar_export_frame = tk.LabelFrame(
            export_box,
            text=" Экспорт «Годовой цифровой отчёт» (стр. 40 журнала: 1 страница) ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        ar_export_frame.pack(fill=tk.X, pady=(0, 10))

        export_ar_grid = tk.Frame(ar_export_frame, bg=card_bg)
        export_ar_grid.pack(fill=tk.X, pady=(0, 4))

        btn_col_ar = tk.Frame(export_ar_grid, bg=card_bg)
        btn_col_ar.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        btn_gen_ar = tk.Button(
            btn_col_ar,
            text="💾 Годовой цифровой отчёт (стр. 40)",
            command=self.on_generate_annual_report_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_ar.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_annual_report = tk.Button(
            btn_col_ar,
            text="📂 Открыть файл отчёта",
            command=self.on_open_annual_report_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_annual_report.pack(fill=tk.X)

        # Сетка 8: Экспорт «Отчёт об отработанном времени» (подсчёт часов за учебный год)
        wh_export_frame = tk.LabelFrame(
            export_box,
            text=" Экспорт «Отчёт об отработанном времени» (подсчёт часов за учебный год) ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        wh_export_frame.pack(fill=tk.X, pady=(0, 10))

        export_wh_grid = tk.Frame(wh_export_frame, bg=card_bg)
        export_wh_grid.pack(fill=tk.X, pady=(0, 4))

        btn_col_wh = tk.Frame(export_wh_grid, bg=card_bg)
        btn_col_wh.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        btn_gen_wh = tk.Button(
            btn_col_wh,
            text="💾 Отчёт об отработанном времени (подсчёт часов в .xlsx)",
            command=self.on_generate_work_hours_click,
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=8,
            cursor="hand2"
        )
        btn_gen_wh.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_work_hours = tk.Button(
            btn_col_wh,
            text="📂 Открыть отчёт об отработанном времени",
            command=self.on_open_work_hours_file_click,
            font=(self.font_sans, 8),
            bg="#e2e8f0",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_work_hours.pack(fill=tk.X)

        # Ряд для экспорта полного комплекта в один файл Excel (все страницы)
        full_export_box = tk.Frame(export_box, bg=card_bg, pady=4)
        full_export_box.pack(fill=tk.X)

        btn_gen_full = tk.Button(
            full_export_box,
            text="📦 Сформировать полный комплект журнала (все 42 страницы: обложка с чистым оборотом + титул + все 12 месяцев + мероприятия + достижения + список обучающихся + инструктаж по ТБ + годовой отчёт в один .xlsx)",
            command=self.on_generate_full_click,
            font=(self.font_sans, 10, "bold"),
            bg="#7c3aed",
            fg="#ffffff",
            activebackground="#6d28d9",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            pady=9,
            cursor="hand2"
        )
        btn_gen_full.pack(fill=tk.X, pady=(0, 4))

        self.btn_open_full = tk.Button(
            full_export_box,
            text="📂 Открыть файл полного комплекта",
            command=self.on_open_full_file_click,
            font=(self.font_sans, 9),
            bg="#f3e8ff",
            fg="#6b21a8",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_full.pack(fill=tk.X)

        # Глобальные горячие клавиши окна (Ctrl+S, Ctrl+0..5, F1)
        self.bind("<Control-s>", lambda e: self.on_generate_click())
        self.bind("<Control-S>", lambda e: self.on_generate_click())
        self.bind("<Command-s>", lambda e: self.on_generate_click())
        self.bind("<Control-Key-0>", lambda e: self.notebook.select(0))
        self.bind("<Control-Key-1>", lambda e: self.notebook.select(1))
        self.bind("<Control-Key-2>", lambda e: self.notebook.select(2))
        self.bind("<Control-Key-3>", lambda e: self.notebook.select(3))
        self.bind("<Control-Key-4>", lambda e: self.notebook.select(4))
        self.bind("<Control-Key-5>", lambda e: self.notebook.select(5))
        self.bind("<Command-Key-0>", lambda e: self.notebook.select(0))
        self.bind("<Command-Key-1>", lambda e: self.notebook.select(1))
        self.bind("<Command-Key-2>", lambda e: self.notebook.select(2))
        self.bind("<Command-Key-3>", lambda e: self.notebook.select(3))
        self.bind("<Command-Key-4>", lambda e: self.notebook.select(4))
        self.bind("<Command-Key-5>", lambda e: self.notebook.select(5))
        self.bind("<F1>", lambda e: self.show_shortcuts_help())
        if hasattr(self, 'combo_year') and isinstance(self.combo_year, ttk.Combobox):
            self.bind("<F4>", lambda e: self.combo_year.event_generate('<Down>'))
            self.bind("<Alt-Down>", lambda e: self.combo_year.event_generate('<Down>'))

        # Глобальная вставка Ctrl+V (EN/RU раскладки, Shift+Insert, Cmd+V)
        self.bind("<Control-v>", self.on_global_paste_shortcut)
        self.bind("<Control-V>", self.on_global_paste_shortcut)
        self.bind("<Command-v>", self.on_global_paste_shortcut)
        self.bind("<Command-V>", self.on_global_paste_shortcut)
        self.bind("<Shift-Insert>", self.on_global_paste_shortcut)
        self.bind("<KeyPress>", self.on_global_keypress, add="+")

    def setup_month_tab(self, parent_frame, card_bg, month_info: dict, next_tab_idx: int = None):
        """Создает интерфейс вкладки месяца с 2 страницами: 1. Учёт посещаемости; 2. Содержание занятий по ДОП."""
        m_key = month_info["key"]
        m_name = month_info["name"]
        tab_num = month_info["tab_num"]
        p1_no = month_info["p1"]
        p2_no = month_info["p2"]

        m_data = self.months_data[m_key]

        m_notebook = ttk.Notebook(parent_frame)
        m_notebook.pack(fill=tk.BOTH, expand=True)
        m_data["notebook"] = m_notebook
        if m_key == "september":
            self.sept_notebook = m_notebook

        m_data["student_widgets"] = []
        m_data["topic_widgets"] = []

        tab_m_p1 = tk.Frame(m_notebook, bg=card_bg, padx=10, pady=8)
        tab_m_p2 = tk.Frame(m_notebook, bg=card_bg, padx=10, pady=8)
        m_notebook.add(tab_m_p1, text=f"Страница 1: Учёт посещаемости ({m_name}, стр. {p1_no})")
        m_notebook.add(tab_m_p2, text=f"Страница 2: Содержание Занятий согласно ДОП ({m_name}, стр. {p2_no})")

        if m_key == "september":
            self.tab_sept_p1 = tab_m_p1
            self.tab_sept_p2 = tab_m_p2

        # =============================================================
        # СТРАНИЦА 1: Учёт посещаемости и выполнения (стр. p1_no журнала)
        # =============================================================
        p1_top = tk.Frame(tab_m_p1, bg=card_bg)
        p1_top.pack(fill=tk.X, pady=(0, 6))

        p1_info_bar = tk.Frame(p1_top, bg=card_bg)
        p1_info_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Label(
            p1_info_bar,
            text=f"Страница 1 (стр. {p1_no} журнала): Учёт посещаемости и выполнения — {m_name}",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a"
        ).pack(side=tk.LEFT)

        tk.Label(
            p1_info_bar,
            text="💡 Вставляйте скопированный список детей через Ctrl+V в любую строку таблицы",
            font=(self.font_sans, 8, "italic"),
            bg=card_bg,
            fg="#64748b"
        ).pack(side=tk.RIGHT)

        p1_btns = tk.Frame(p1_top, bg=card_bg)
        p1_btns.pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            p1_btns,
            text="📋 Вставить список детей (Ctrl+V)",
            command=lambda k=m_key: self.paste_students_from_clipboard(0, month_key=k),
            font=(self.font_sans, 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        if m_key != "september":
            tk.Button(
                p1_btns,
                text="👥 Скопировать детей из Сентября",
                command=lambda k=m_key: self.copy_students_from_month(k, "september"),
                font=(self.font_sans, 8, "bold"),
                bg="#f0fdf4",
                fg="#15803d",
                activebackground="#dcfce7",
                activeforeground="#166534",
                relief=tk.SOLID,
                bd=1,
                padx=8,
                pady=4,
                cursor="hand2"
            ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            p1_btns,
            text="🗑 Очистить список детей",
            command=lambda k=m_key: self.clear_month_students(k),
            font=(self.font_sans, 8),
            bg="#fee2e2",
            fg="#b91c1c",
            activebackground="#fecaca",
            activeforeground="#991b1b",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            p1_btns,
            text="Очистить отметки и даты",
            command=lambda k=m_key: self.clear_month_dates_and_attendance(k),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#475569",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT)

        tk.Button(
            p1_btns,
            text="К Странице 2 (Содержание ДОП) ➔",
            command=lambda nb=m_notebook: nb.select(1),
            font=(self.font_sans, 9, "bold"),
            bg="#0d9488",
            fg="#ffffff",
            activebackground="#0f766e",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.RIGHT)

        # Контейнер прокрутки таблицы Страницы 1
        table_container_p1 = tk.Frame(tab_m_p1, bg="#cbd5e1", bd=1, relief=tk.SOLID)
        table_container_p1.pack(fill=tk.BOTH, expand=True)

        canvas_p1 = tk.Canvas(table_container_p1, bg="#ffffff", highlightthickness=0)
        scroll_y_p1 = ttk.Scrollbar(table_container_p1, orient="vertical", command=canvas_p1.yview)
        scroll_x_p1 = ttk.Scrollbar(table_container_p1, orient="horizontal", command=canvas_p1.xview)
        canvas_p1.configure(xscrollcommand=scroll_x_p1.set, yscrollcommand=scroll_y_p1.set)

        scroll_y_p1.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x_p1.pack(side=tk.BOTTOM, fill=tk.X)
        canvas_p1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        grid_p1 = tk.Frame(canvas_p1, bg="#e2e8f0")
        canvas_p1.create_window((0, 0), window=grid_p1, anchor="nw")

        def _on_p1_configure(e):
            canvas_p1.configure(scrollregion=canvas_p1.bbox("all"))

        grid_p1.bind("<Configure>", _on_p1_configure)

        def _on_mousewheel_p1(e):
            if sys.platform == "darwin":
                canvas_p1.yview_scroll(int(-1 * e.delta), "units")
            else:
                canvas_p1.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas_p1.bind("<Enter>", lambda e: canvas_p1.bind_all("<MouseWheel>", _on_mousewheel_p1))
        canvas_p1.bind("<Leave>", lambda e: canvas_p1.unbind_all("<MouseWheel>"))

        for w in (tab_m_p1, canvas_p1, grid_p1):
            for seq in ("<Control-v>", "<Control-V>", "<Command-v>", "<Command-V>", "<Shift-Insert>"):
                w.bind(seq, lambda ev, k=m_key: self.on_global_paste_shortcut(ev, month_key=k))

        # Шапка таблицы Страницы 1
        tk.Label(
            grid_p1,
            text="№\nп/п",
            font=(self.font_sans, 8, "bold"),
            bg="#f1f5f9",
            fg="#0f172a",
            width=4,
            relief=tk.SOLID,
            bd=1
        ).grid(row=0, column=0, rowspan=2, sticky="nsew", padx=1, pady=1)

        tk.Label(
            grid_p1,
            text="Фамилия, имя обучающегося\n(Ctrl+V в строку для вставки списка)",
            font=(self.font_sans, 8, "bold"),
            bg="#f1f5f9",
            fg="#0f172a",
            width=28,
            relief=tk.SOLID,
            bd=1
        ).grid(row=0, column=1, rowspan=2, sticky="nsew", padx=1, pady=1)

        tk.Label(
            grid_p1,
            text=f"Д а т ы   з а н я т и й   ({m_name}, 15 колонок для журнала)",
            font=(self.font_sans, 8, "bold"),
            bg="#f1f5f9",
            fg="#0f172a",
            relief=tk.SOLID,
            bd=1
        ).grid(row=0, column=2, columnspan=15, sticky="nsew", padx=1, pady=1)

        # Ряд 1: поля ввода для 15 дат занятий
        for c in range(15):
            e_date = ttk.Entry(
                grid_p1,
                textvariable=m_data["date_vars"][c],
                width=4,
                font=(self.font_sans, 8, "bold"),
                justify="center"
            )
            e_date.grid(row=1, column=2 + c, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_date)
            e_date.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

        # Ряды 2..31: 30 строк для обучающихся
        for i in range(30):
            tk.Label(
                grid_p1,
                text=str(i + 1),
                font=(self.font_sans, 8, "bold"),
                bg="#f8fafc",
                fg="#334155",
                width=4,
                relief=tk.SOLID,
                bd=1
            ).grid(row=2 + i, column=0, sticky="nsew", padx=1, pady=1)

            e_student = ttk.Entry(
                grid_p1,
                textvariable=m_data["student_vars"][i],
                font=(self.font_sans, 9),
                width=28
            )
            e_student.grid(row=2 + i, column=1, sticky="nsew", padx=1, pady=1)
            m_data["student_widgets"].append(e_student)
            if m_key == "september":
                self.september_student_widgets.append(e_student)

            self.attach_student_context_menu(e_student, i, month_key=m_key)
            e_student.bind("<KeyPress>", lambda ev, idx=i, k=m_key: self.on_student_entry_keypress(ev, idx, month_key=k), add="+")
            e_student.bind("<<Paste>>", lambda ev, idx=i, k=m_key: self.on_student_entry_paste(ev, idx, month_key=k))
            e_student.bind("<Control-v>", lambda ev, idx=i, k=m_key: self.paste_students_from_clipboard(idx, ev, month_key=k))
            e_student.bind("<Control-V>", lambda ev, idx=i, k=m_key: self.paste_students_from_clipboard(idx, ev, month_key=k))
            e_student.bind("<Command-v>", lambda ev, idx=i, k=m_key: self.paste_students_from_clipboard(idx, ev, month_key=k))
            e_student.bind("<Shift-Insert>", lambda ev, idx=i, k=m_key: self.paste_students_from_clipboard(idx, ev, month_key=k))
            e_student.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

            for c in range(15):
                e_att = ttk.Entry(
                    grid_p1,
                    textvariable=m_data["attendance_vars"][i][c],
                    width=4,
                    font=(self.font_sans, 8),
                    justify="center"
                )
                e_att.grid(row=2 + i, column=2 + c, sticky="nsew", padx=1, pady=1)
                e_att.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

        # =============================================================
        # СТРАНИЦА 2: Содержание занятий согласно ДОП (стр. p2_no журнала)
        # =============================================================
        p2_top = tk.Frame(tab_m_p2, bg=card_bg)
        p2_top.pack(fill=tk.X, pady=(0, 6))

        p2_info_bar = tk.Frame(p2_top, bg=card_bg)
        p2_info_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Label(
            p2_info_bar,
            text=f"Страница 2 (стр. {p2_no} журнала): Содержание Занятий согласно ДОП (16 строк) — {m_name}",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a"
        ).pack(side=tk.LEFT)

        tk.Label(
            p2_info_bar,
            text="💡 Вставляйте список тем через Ctrl+V в столбец 'Содержание Занятий'",
            font=(self.font_sans, 8, "italic"),
            bg=card_bg,
            fg="#64748b"
        ).pack(side=tk.RIGHT)

        p2_btns = tk.Frame(p2_top, bg=card_bg)
        p2_btns.pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            p2_btns,
            text="📋 Вставить темы (Ctrl+V)",
            command=lambda k=m_key: self.paste_topics_from_clipboard(0, month_key=k),
            font=(self.font_sans, 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            p2_btns,
            text="📅 Скопировать даты со Стр. 1",
            command=lambda k=m_key: self.copy_dates_to_month_page2(k),
            font=(self.font_sans, 8),
            bg="#0d9488",
            fg="#ffffff",
            activebackground="#0f766e",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            p2_btns,
            text="⚡ Заполнить часы (2 ч)",
            command=lambda k=m_key: self.fill_default_hours_month(k),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#334155",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            p2_btns,
            text="🗑 Очистить таблицу",
            command=lambda k=m_key: self.clear_month_topics(k),
            font=(self.font_sans, 8),
            bg="#fee2e2",
            fg="#b91c1c",
            activebackground="#fecaca",
            activeforeground="#991b1b",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT)

        if next_tab_idx is not None:
            if next_tab_idx < 16:
                next_month_name = MONTHS_CONFIG[next_tab_idx - 4]["name"]
                next_text = f"Далее: {next_month_name} ➔"
            else:
                next_text = "Далее: Формирование Excel ➔"

            tk.Button(
                p2_btns,
                text=next_text,
                command=lambda idx=next_tab_idx: self.notebook.select(idx),
                font=(self.font_sans, 9, "bold"),
                bg="#2563eb",
                fg="#ffffff",
                activebackground="#1d4ed8",
                activeforeground="#ffffff",
                relief=tk.SOLID,
                bd=1,
                padx=10,
                pady=4,
                cursor="hand2"
            ).pack(side=tk.RIGHT)

        # Контейнер прокрутки таблицы Страницы 2
        table_container_p2 = tk.Frame(tab_m_p2, bg="#cbd5e1", bd=1, relief=tk.SOLID)
        table_container_p2.pack(fill=tk.BOTH, expand=True)

        canvas_p2 = tk.Canvas(table_container_p2, bg="#ffffff", highlightthickness=0)
        scroll_y_p2 = ttk.Scrollbar(table_container_p2, orient="vertical", command=canvas_p2.yview)
        scroll_x_p2 = ttk.Scrollbar(table_container_p2, orient="horizontal", command=canvas_p2.xview)
        canvas_p2.configure(xscrollcommand=scroll_x_p2.set, yscrollcommand=scroll_y_p2.set)

        scroll_y_p2.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x_p2.pack(side=tk.BOTTOM, fill=tk.X)
        canvas_p2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        grid_p2 = tk.Frame(canvas_p2, bg="#e2e8f0")
        canvas_p2.create_window((0, 0), window=grid_p2, anchor="nw")

        def _on_p2_configure(e):
            canvas_p2.configure(scrollregion=canvas_p2.bbox("all"))

        grid_p2.bind("<Configure>", _on_p2_configure)

        def _on_mousewheel_p2(e):
            if sys.platform == "darwin":
                canvas_p2.yview_scroll(int(-1 * e.delta), "units")
            else:
                canvas_p2.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas_p2.bind("<Enter>", lambda e: canvas_p2.bind_all("<MouseWheel>", _on_mousewheel_p2))
        canvas_p2.bind("<Leave>", lambda e: canvas_p2.unbind_all("<MouseWheel>"))

        for w in (tab_m_p2, canvas_p2, grid_p2):
            for seq in ("<Control-v>", "<Control-V>", "<Command-v>", "<Command-V>", "<Shift-Insert>"):
                w.bind(seq, lambda ev, k=m_key: self.on_global_paste_shortcut(ev, month_key=k))

        # Шапка таблицы Страницы 2 (6 столбцов)
        headers_p2 = [
            ("Даты занятий\nобъединения", 14),
            ("Содержание Занятий согласно ДОП\n(Ctrl+V для вставки списка тем)", 42),
            ("Часы", 5),
            ("Подпись\nпедагога", 12),
            ("Часы", 5),
            ("Подпись аккомпаниатора\n(концертмейстера)", 18)
        ]
        for col_idx, (h_title, h_width) in enumerate(headers_p2):
            tk.Label(
                grid_p2,
                text=h_title,
                font=(self.font_sans, 8, "bold"),
                bg="#f1f5f9",
                fg="#0f172a",
                width=h_width,
                relief=tk.SOLID,
                bd=1,
                pady=4
            ).grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)

        # 16 строк данных Страницы 2
        for i in range(16):
            # 1. Даты занятий объединения
            e_dt = ttk.Entry(
                grid_p2,
                textvariable=m_data["topic_row_vars"][i][0],
                width=14,
                font=(self.font_sans, 9),
                justify="center"
            )
            e_dt.grid(row=1 + i, column=0, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_dt)
            e_dt.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

            # 2. Содержание занятий согласно ДОП
            e_cnt = ttk.Entry(
                grid_p2,
                textvariable=m_data["topic_row_vars"][i][1],
                width=42,
                font=(self.font_sans, 9)
            )
            e_cnt.grid(row=1 + i, column=1, sticky="nsew", padx=1, pady=1)
            m_data["topic_widgets"].append(e_cnt)
            if m_key == "september":
                self.september_topic_widgets.append(e_cnt)

            self.attach_topic_context_menu(e_cnt, i, month_key=m_key)
            e_cnt.bind("<KeyPress>", lambda ev, idx=i, k=m_key: self.on_topic_entry_keypress(ev, idx, month_key=k), add="+")
            e_cnt.bind("<<Paste>>", lambda ev, idx=i, k=m_key: self.on_topic_entry_paste(ev, idx, month_key=k))
            e_cnt.bind("<Control-v>", lambda ev, idx=i, k=m_key: self.paste_topics_from_clipboard(idx, ev, month_key=k))
            e_cnt.bind("<Control-V>", lambda ev, idx=i, k=m_key: self.paste_topics_from_clipboard(idx, ev, month_key=k))
            e_cnt.bind("<Command-v>", lambda ev, idx=i, k=m_key: self.paste_topics_from_clipboard(idx, ev, month_key=k))
            e_cnt.bind("<Shift-Insert>", lambda ev, idx=i, k=m_key: self.paste_topics_from_clipboard(idx, ev, month_key=k))
            e_cnt.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

            # 3. Часы педагога
            e_ht = ttk.Entry(
                grid_p2,
                textvariable=m_data["topic_row_vars"][i][2],
                width=5,
                font=(self.font_sans, 9),
                justify="center"
            )
            e_ht.grid(row=1 + i, column=2, sticky="nsew", padx=1, pady=1)
            e_ht.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

            # 4. Подпись педагога
            e_st = ttk.Entry(
                grid_p2,
                textvariable=m_data["topic_row_vars"][i][3],
                width=12,
                font=(self.font_sans, 9),
                justify="center"
            )
            e_st.grid(row=1 + i, column=3, sticky="nsew", padx=1, pady=1)
            e_st.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

            # 5. Часы аккомпаниатора
            e_ha = ttk.Entry(
                grid_p2,
                textvariable=m_data["topic_row_vars"][i][4],
                width=5,
                font=(self.font_sans, 9),
                justify="center"
            )
            e_ha.grid(row=1 + i, column=4, sticky="nsew", padx=1, pady=1)
            e_ha.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

            # 6. Подпись аккомпаниатора
            e_sa = ttk.Entry(
                grid_p2,
                textvariable=m_data["topic_row_vars"][i][5],
                width=18,
                font=(self.font_sans, 9),
                justify="center"
            )
            e_sa.grid(row=1 + i, column=5, sticky="nsew", padx=1, pady=1)
            e_sa.bind("<FocusOut>", lambda ev, k=m_key: self.auto_save_month_data(k))

    def setup_september_tab(self, parent_frame, card_bg):
        """Совместимость: настройка вкладки 'Сентябрь'."""
        self.setup_month_tab(parent_frame, card_bg, MONTHS_CONFIG[0], next_tab_idx=5)

    def setup_mass_events_tab(self, parent_frame, card_bg, next_tab_idx: int = 17):
        """
        Настройка вкладки 'Учёт массовых мероприятий' с 2 одинаковыми страницами (стр. 30 и 31 журнала).
        На каждой странице таблица из 5 столбцов и 26 строк.
        """
        self.mass_events_notebook = ttk.Notebook(parent_frame)
        self.mass_events_notebook.pack(fill=tk.BOTH, expand=True)

        tab_p1 = tk.Frame(self.mass_events_notebook, bg=card_bg, padx=10, pady=8)
        self.mass_events_notebook.add(tab_p1, text="Страница 1 (стр. 30)")
        self.setup_single_mass_events_page(
            parent_tab=tab_p1,
            card_bg=card_bg,
            page_num=1,
            journal_page_no=30,
            var_list=self.mass_events_p1_vars,
            widget_list=self.mass_events_p1_widgets,
            next_action=lambda: self.mass_events_notebook.select(1),
            next_btn_text="К Странице 2 (стр. 31) ➔"
        )

        tab_p2 = tk.Frame(self.mass_events_notebook, bg=card_bg, padx=10, pady=8)
        self.mass_events_notebook.add(tab_p2, text="Страница 2 (стр. 31)")
        self.setup_single_mass_events_page(
            parent_tab=tab_p2,
            card_bg=card_bg,
            page_num=2,
            journal_page_no=31,
            var_list=self.mass_events_p2_vars,
            widget_list=self.mass_events_p2_widgets,
            next_action=lambda: self.notebook.select(next_tab_idx),
            next_btn_text="Далее: Формирование Excel ➔"
        )

    def setup_single_mass_events_page(
        self,
        parent_tab,
        card_bg,
        page_num: int,
        journal_page_no: int,
        var_list: list,
        widget_list: list,
        next_action,
        next_btn_text: str
    ):
        """Создает страницу учёта массовых мероприятий: 5 столбцов, 26 строк с прокруткой и действиями."""
        # Верхняя информационная панель
        top_frame = tk.Frame(parent_tab, bg=card_bg)
        top_frame.pack(fill=tk.X, pady=(0, 6))

        info_bar = tk.Frame(top_frame, bg=card_bg)
        info_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Label(
            info_bar,
            text=f"Страница {page_num} (стр. {journal_page_no} журнала): Учёт массовых мероприятий с обучающимися (26 строк)",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a"
        ).pack(side=tk.LEFT)

        tk.Label(
            info_bar,
            text="💡 Вставляйте строки через Ctrl+V, используйте Tab для перехода между ячейками",
            font=(self.font_sans, 8, "italic"),
            bg=card_bg,
            fg="#64748b"
        ).pack(side=tk.RIGHT)

        # Панель кнопок действий
        btns_bar = tk.Frame(top_frame, bg=card_bg)
        btns_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btns_bar,
            text="📋 Вставить из буфера (Ctrl+V)",
            command=lambda: self.paste_mass_events_from_clipboard(page_num=page_num, start_idx=0),
            font=(self.font_sans, 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btns_bar,
            text="🗑 Очистить страницу",
            command=lambda: self.clear_mass_events_page(page_num),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#ef4444",
            activebackground="#fee2e2",
            activeforeground="#b91c1c",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btns_bar,
            text=next_btn_text,
            command=next_action,
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7" if page_num == 1 else "#15803d",
            fg="#ffffff",
            activebackground="#0369a1" if page_num == 1 else "#166534",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.RIGHT)

        # Область таблицы с прокруткой (Canvas)
        table_container = tk.Frame(parent_tab, bg=card_bg, relief=tk.SOLID, bd=1)
        table_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(table_container, bg=card_bg, highlightthickness=0)
        v_scroll = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(table_container, orient="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        grid_frame = tk.Frame(canvas, bg=card_bg)
        canvas.create_window((0, 0), window=grid_frame, anchor="nw")

        def _on_cfg(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        grid_frame.bind("<Configure>", _on_cfg)

        def _on_mw(e):
            if sys.platform == "darwin":
                canvas.yview_scroll(int(-1 * e.delta), "units")
            else:
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mw))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        for w in (parent_tab, canvas, grid_frame):
            for seq in ("<Control-v>", "<Control-V>", "<Command-v>", "<Command-V>", "<Shift-Insert>"):
                w.bind(seq, lambda ev, p=page_num: self.paste_mass_events_from_clipboard(page_num=p, start_idx=0, event=ev))

        # Шапка таблицы (5 столбцов + порядковый номер)
        headers = [
            ("№", 4),
            ("Дата", 11),
            ("Название и краткое содержание проведенного мероприятия\n(Ctrl+V для построчной вставки)", 46),
            ("Количество\nучастников", 13),
            ("Место проведения\nмероприятия", 22),
            ("Кто проводил", 22)
        ]
        for col_idx, (h_title, h_width) in enumerate(headers):
            tk.Label(
                grid_frame,
                text=h_title,
                font=(self.font_sans, 8, "bold"),
                bg="#f1f5f9",
                fg="#0f172a",
                width=h_width,
                relief=tk.SOLID,
                bd=1,
                pady=5
            ).grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)

        # 26 строк данных
        for i in range(26):
            row_num = 1 + i

            # 0. Номер строки
            tk.Label(
                grid_frame,
                text=str(i + 1),
                font=(self.font_sans, 8, "bold"),
                bg="#f8fafc",
                fg="#64748b",
                width=4,
                relief=tk.SOLID,
                bd=1,
                pady=2
            ).grid(row=row_num, column=0, sticky="nsew", padx=1, pady=1)

            # 1. Дата
            e_date = ttk.Entry(
                grid_frame,
                textvariable=var_list[i][0],
                width=11,
                font=(self.font_sans, 9),
                justify="center"
            )
            e_date.grid(row=row_num, column=1, sticky="nsew", padx=1, pady=1)
            self.attach_mass_event_context_menu(e_date, page_num, i)
            e_date.bind("<KeyPress>", lambda ev, idx=i, p=page_num: self.on_mass_event_entry_keypress(ev, idx, p), add="+")
            e_date.bind("<FocusOut>", lambda ev: self.auto_save_mass_events_data())

            # 2. Название и краткое содержание
            e_content = ttk.Entry(
                grid_frame,
                textvariable=var_list[i][1],
                width=46,
                font=(self.font_sans, 9)
            )
            e_content.grid(row=row_num, column=2, sticky="nsew", padx=1, pady=1)
            widget_list.append(e_content)
            self.attach_mass_event_context_menu(e_content, page_num, i)
            e_content.bind("<KeyPress>", lambda ev, idx=i, p=page_num: self.on_mass_event_entry_keypress(ev, idx, p), add="+")
            e_content.bind("<<Paste>>", lambda ev, idx=i, p=page_num: self.on_mass_event_entry_paste(ev, idx, p))
            e_content.bind("<FocusOut>", lambda ev: self.auto_save_mass_events_data())

            # 3. Количество участников
            e_count = ttk.Entry(
                grid_frame,
                textvariable=var_list[i][2],
                width=13,
                font=(self.font_sans, 9),
                justify="center"
            )
            e_count.grid(row=row_num, column=3, sticky="nsew", padx=1, pady=1)
            self.attach_mass_event_context_menu(e_count, page_num, i)
            e_count.bind("<KeyPress>", lambda ev, idx=i, p=page_num: self.on_mass_event_entry_keypress(ev, idx, p), add="+")
            e_count.bind("<FocusOut>", lambda ev: self.auto_save_mass_events_data())

            # 4. Место проведения
            e_loc = ttk.Entry(
                grid_frame,
                textvariable=var_list[i][3],
                width=22,
                font=(self.font_sans, 9)
            )
            e_loc.grid(row=row_num, column=4, sticky="nsew", padx=1, pady=1)
            self.attach_mass_event_context_menu(e_loc, page_num, i)
            e_loc.bind("<KeyPress>", lambda ev, idx=i, p=page_num: self.on_mass_event_entry_keypress(ev, idx, p), add="+")
            e_loc.bind("<FocusOut>", lambda ev: self.auto_save_mass_events_data())

            # 5. Кто проводил
            e_who = ttk.Entry(
                grid_frame,
                textvariable=var_list[i][4],
                width=22,
                font=(self.font_sans, 9)
            )
            e_who.grid(row=row_num, column=5, sticky="nsew", padx=1, pady=1)
            self.attach_mass_event_context_menu(e_who, page_num, i)
            e_who.bind("<KeyPress>", lambda ev, idx=i, p=page_num: self.on_mass_event_entry_keypress(ev, idx, p), add="+")
            e_who.bind("<FocusOut>", lambda ev: self.auto_save_mass_events_data())

    def attach_mass_event_context_menu(self, widget, page_num: int, idx: int):
        """Контекстное меню для ячеек таблицы массовых мероприятий."""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(
            label="📋 Вставить строки из буфера (Ctrl+V)",
            command=lambda: self.paste_mass_events_from_clipboard(page_num=page_num, start_idx=idx)
        )
        menu.add_separator()
        menu.add_command(label="Вырезать", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Копировать", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Вставить текст", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_command(label="Выделить всё", command=lambda: widget.select_range(0, tk.END))

        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)

    def on_mass_event_entry_keypress(self, event, idx: int, page_num: int):
        """Обработка сочетаний клавиш в ячейках таблицы мероприятий."""
        if self.is_paste_event(event):
            self.paste_mass_events_from_clipboard(page_num=page_num, start_idx=idx, event=event)
            return "break"

        state = getattr(event, 'state', 0)
        is_ctrl = bool(state & 0x4) or bool(state & 0x8) or bool(state & 0x40000)
        code = getattr(event, 'keycode', 0)
        key = (getattr(event, 'keysym', '') or "").lower()

        if is_ctrl:
            if key in ('a', 'cyrillic_ef') or code == 65:
                self.select_all_widget(event.widget)
                return "break"
            elif key in ('c', 'cyrillic_es') or code == 67:
                self.copy_from_widget(event.widget)
                return "break"
            elif key in ('x', 'cyrillic_che') or code == 88:
                self.cut_from_widget(event.widget)
                return "break"

        return None

    def on_mass_event_entry_paste(self, event, idx: int, page_num: int):
        """Перехват события <<Paste>> для ячеек таблицы мероприятий."""
        self.paste_mass_events_from_clipboard(page_num=page_num, start_idx=idx, event=event)
        return "break"

    def clear_mass_events_page(self, page_num: int):
        """Очищает данные страницы мероприятий после подтверждения."""
        pg_title = f"Страница {page_num} (стр. {29 + page_num} журнала)"
        if not messagebox.askyesno("Очистить таблицу", f"Вы уверены, что хотите полностью очистить {pg_title}?"):
            return
        var_list = self.mass_events_p1_vars if page_num == 1 else self.mass_events_p2_vars
        for r in var_list:
            for v in r:
                v.set("")
        self.auto_save_mass_events_data()
        self.status_var.set(f"Таблица {pg_title} очищена.")

    def auto_save_mass_events_data(self, *args):
        """Автоматически сохраняет все данные вкладки 'Учёт массовых мероприятий' в config_data и на диск."""
        try:
            p1_data = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "count": r[2].get(),
                    "location": r[3].get(),
                    "conducted_by": r[4].get()
                }
                for r in self.mass_events_p1_vars
            ]
            p2_data = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "count": r[2].get(),
                    "location": r[3].get(),
                    "conducted_by": r[4].get()
                }
                for r in self.mass_events_p2_vars
            ]
            self.config_data["mass_events_p1"] = p1_data
            self.config_data["mass_events_p2"] = p2_data
            save_config(self.config_data)
        except Exception:
            pass

    def paste_mass_events_from_clipboard(self, page_num: int = 1, start_idx: int = 0, event=None):
        """Вставляет строки массовых мероприятий из буфера обмена, начиная со строки start_idx."""
        raw_text = self.get_clipboard_text()
        if not raw_text:
            self.status_var.set("Буфер обмена пуст или не содержит текст.")
            return "break"

        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
        if not lines:
            self.status_var.set("В буфере обмена нет строк для вставки.")
            return "break"

        var_list = self.mass_events_p1_vars if page_num == 1 else self.mass_events_p2_vars
        count = 0
        for offset, line in enumerate(lines):
            idx = start_idx + offset
            if idx >= 26:
                break
            parts = [p.strip() for p in line.split('\t')]
            if len(parts) >= 5:
                var_list[idx][0].set(parts[0])
                var_list[idx][1].set(parts[1])
                var_list[idx][2].set(parts[2])
                var_list[idx][3].set(parts[3])
                var_list[idx][4].set(parts[4])
            elif len(parts) == 4:
                var_list[idx][0].set(parts[0])
                var_list[idx][1].set(parts[1])
                var_list[idx][2].set(parts[2])
                var_list[idx][3].set(parts[3])
            elif len(parts) == 3:
                var_list[idx][0].set(parts[0])
                var_list[idx][1].set(parts[1])
                var_list[idx][2].set(parts[2])
            elif len(parts) == 2:
                if re.search(r'\d', parts[0]) and ('.' in parts[0] or '/' in parts[0]):
                    var_list[idx][0].set(parts[0])
                    var_list[idx][1].set(parts[1])
                else:
                    var_list[idx][1].set(parts[0])
                    var_list[idx][2].set(parts[1])
            else:
                clean_line = re.sub(r'^\d+[\.\)\s\-]+\s*', '', line)
                var_list[idx][1].set(clean_line or line)
            count += 1

        self.auto_save_mass_events_data()
        pg_str = f"стр. {29 + page_num}"
        self.status_var.set(f"Вставлено {count} записей мероприятий ({pg_str}, начиная со строки {start_idx + 1}).")
        return "break"

    def on_generate_events_p1_click(self):
        """Формирует и сохраняет отдельный Excel-файл Учёта массовых мероприятий (стр. 30 журнала)."""
        self.auto_save_mass_events_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Массовые_мероприятия_стр30_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Учёт массовых мероприятий (стр. 30) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            events = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "count": r[2].get(),
                    "location": r[3].get(),
                    "conducted_by": r[4].get()
                }
                for r in self.mass_events_p1_vars
            ]
            generate_excel_mass_events_p1(save_path, events=events)
            self.last_saved_events_p1_file = save_path
            if hasattr(self, 'btn_open_events_p1'):
                self.btn_open_events_p1.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Файл 'Учёт массовых мероприятий' (стр. 30) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении мероприятий (стр. 30).")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_events_p1_file_click(self):
        """Открывает сгенерированный файл мероприятий (стр. 30)."""
        if not hasattr(self, 'last_saved_events_p1_file') or not self.last_saved_events_p1_file or not os.path.exists(self.last_saved_events_p1_file):
            messagebox.showwarning("Файл не найден", "Созданный файл мероприятий (стр. 30) не найден на диске.")
            return
        self._open_file_system(self.last_saved_events_p1_file)

    def on_generate_events_p2_click(self):
        """Формирует и сохраняет отдельный Excel-файл Учёта массовых мероприятий (стр. 31 журнала)."""
        self.auto_save_mass_events_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Массовые_мероприятия_стр31_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Учёт массовых мероприятий (стр. 31) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            events = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "count": r[2].get(),
                    "location": r[3].get(),
                    "conducted_by": r[4].get()
                }
                for r in self.mass_events_p2_vars
            ]
            generate_excel_mass_events_p2(save_path, events=events)
            self.last_saved_events_p2_file = save_path
            if hasattr(self, 'btn_open_events_p2'):
                self.btn_open_events_p2.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Файл 'Учёт массовых мероприятий' (стр. 31) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении мероприятий (стр. 31).")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_events_p2_file_click(self):
        """Открывает сгенерированный файл мероприятий (стр. 31)."""
        if not hasattr(self, 'last_saved_events_p2_file') or not self.last_saved_events_p2_file or not os.path.exists(self.last_saved_events_p2_file):
            messagebox.showwarning("Файл не найден", "Созданный файл мероприятий (стр. 31) не найден на диске.")
            return
        self._open_file_system(self.last_saved_events_p2_file)

    def on_generate_events_spread_click(self):
        """Формирует и сохраняет разворот Учёта массовых мероприятий (стр. 30 и 31 журнала) в один файл."""
        self.auto_save_mass_events_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Массовые_мероприятия_Разворот_стр30-31_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Разворот мероприятий (стр. 30 и 31) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            ev_p1 = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "count": r[2].get(),
                    "location": r[3].get(),
                    "conducted_by": r[4].get()
                }
                for r in self.mass_events_p1_vars
            ]
            ev_p2 = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "count": r[2].get(),
                    "location": r[3].get(),
                    "conducted_by": r[4].get()
                }
                for r in self.mass_events_p2_vars
            ]
            generate_excel_mass_events_spread(save_path, events_p1=ev_p1, events_p2=ev_p2)
            self.last_saved_events_spread_file = save_path
            if hasattr(self, 'btn_open_events_spread'):
                self.btn_open_events_spread.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен разворот мероприятий: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Разворот мероприятий (стр. 30 и 31) успешно сохранен:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении разворота мероприятий.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_events_spread_file_click(self):
        """Открывает файл разворота мероприятий."""
        if not hasattr(self, 'last_saved_events_spread_file') or not self.last_saved_events_spread_file or not os.path.exists(self.last_saved_events_spread_file):
            messagebox.showwarning("Файл не найден", "Созданный файл разворота мероприятий не найден на диске.")
            return
        self._open_file_system(self.last_saved_events_spread_file)

    def setup_creative_achievements_tab(self, parent_tab, card_bg, next_tab_idx: int = 18):
        """Настройка вкладки 'Творческие достижения' с двумя вложенными страницами (стр. 32 и стр. 33)."""
        self.creative_achievements_notebook = ttk.Notebook(parent_tab)
        self.creative_achievements_notebook.pack(fill=tk.BOTH, expand=True)

        tab_p1 = tk.Frame(self.creative_achievements_notebook, bg=card_bg, padx=10, pady=8)
        self.creative_achievements_notebook.add(tab_p1, text="Страница 1 (стр. 32)")
        self.setup_single_creative_page(
            parent_tab=tab_p1,
            card_bg=card_bg,
            page_num=1,
            journal_page_no=32,
            var_list=self.creative_achievements_p1_vars,
            widget_list=self.creative_achievements_p1_widgets,
            next_action=lambda: self.creative_achievements_notebook.select(1),
            next_btn_text="К Странице 2 (стр. 33) ➔"
        )

        tab_p2 = tk.Frame(self.creative_achievements_notebook, bg=card_bg, padx=10, pady=8)
        self.creative_achievements_notebook.add(tab_p2, text="Страница 2 (стр. 33)")
        self.setup_single_creative_page(
            parent_tab=tab_p2,
            card_bg=card_bg,
            page_num=2,
            journal_page_no=33,
            var_list=self.creative_achievements_p2_vars,
            widget_list=self.creative_achievements_p2_widgets,
            next_action=lambda: self.notebook.select(next_tab_idx),
            next_btn_text="Далее: Формирование Excel ➔"
        )

    def setup_single_creative_page(
        self,
        parent_tab,
        card_bg,
        page_num: int,
        journal_page_no: int,
        var_list: list,
        widget_list: list,
        next_action,
        next_btn_text: str
    ):
        """Создает страницу творческих достижений (26 строк с прокруткой и действиями)."""
        top_frame = tk.Frame(parent_tab, bg=card_bg)
        top_frame.pack(fill=tk.X, pady=(0, 6))

        info_bar = tk.Frame(top_frame, bg=card_bg)
        info_bar.pack(fill=tk.X, pady=(0, 4))

        page_title = "Творческие достижения: участие в соревнованиях, смотрах, спектаклях" if page_num == 1 else "Творческие достижения: результаты и выполненные работы"
        tk.Label(
            info_bar,
            text=f"Страница {page_num} (стр. {journal_page_no} журнала): {page_title} (26 строк)",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a"
        ).pack(side=tk.LEFT)

        tk.Label(
            info_bar,
            text="💡 Вставляйте строки через Ctrl+V, используйте Tab для перехода между ячейками",
            font=(self.font_sans, 8, "italic"),
            bg=card_bg,
            fg="#64748b"
        ).pack(side=tk.RIGHT)

        # Панель кнопок действий
        btns_bar = tk.Frame(top_frame, bg=card_bg)
        btns_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btns_bar,
            text="📋 Вставить из буфера (Ctrl+V)",
            command=lambda: self.paste_creative_achievements_from_clipboard(page_num=page_num, start_idx=0),
            font=(self.font_sans, 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        if page_num == 1:
            tk.Button(
                btns_bar,
                text="👥 Заполнить обучающихся из группы",
                command=self.fill_creative_students_from_group,
                font=(self.font_sans, 8, "bold"),
                bg="#f0fdf4",
                fg="#166534",
                activebackground="#dcfce7",
                activeforeground="#15803d",
                relief=tk.SOLID,
                bd=1,
                padx=8,
                pady=4,
                cursor="hand2"
            ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btns_bar,
            text="🗑 Очистить страницу",
            command=lambda: self.clear_creative_achievements_page(page_num),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#ef4444",
            activebackground="#fee2e2",
            activeforeground="#b91c1c",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btns_bar,
            text=next_btn_text,
            command=next_action,
            font=(self.font_sans, 9, "bold"),
            bg="#0d9488" if page_num == 1 else "#15803d",
            fg="#ffffff",
            activebackground="#0f766e" if page_num == 1 else "#166534",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.RIGHT)

        # Область таблицы с прокруткой (Canvas)
        table_container = tk.Frame(parent_tab, bg=card_bg, relief=tk.SOLID, bd=1)
        table_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(table_container, bg=card_bg, highlightthickness=0)
        v_scroll = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(table_container, orient="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        grid_frame = tk.Frame(canvas, bg=card_bg)
        canvas.create_window((0, 0), window=grid_frame, anchor="nw")

        def _on_cfg(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        grid_frame.bind("<Configure>", _on_cfg)

        def _on_mw(e):
            if sys.platform == "darwin":
                canvas.yview_scroll(int(-1 * e.delta), "units")
            else:
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mw))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        for w in (parent_tab, canvas, grid_frame):
            for seq in ("<Control-v>", "<Control-V>", "<Command-v>", "<Command-V>", "<Shift-Insert>"):
                w.bind(seq, lambda ev, p=page_num: self.paste_creative_achievements_from_clipboard(page_num=p, start_idx=0, event=ev))

        # Шапка таблицы
        if page_num == 1:
            headers = [
                ("№\nп/п", 5),
                ("Фамилия, имя обучающегося\n(Ctrl+V для построчной вставки)", 34),
                ("В каких соревнованиях, смотрах, спектаклях и др. мероприятиях участвовал\n(Ctrl+V для построчной вставки)", 74)
            ]
        else:
            headers = [
                ("№\nп/п", 5),
                ("Результаты (полученное звание, разряд и другие результаты)\n(Ctrl+V для построчной вставки)", 56),
                ("Работы, выполненные объединением по заказам или инициативно\n(Ctrl+V для построчной вставки)", 56)
            ]

        for col_idx, (h_title, h_width) in enumerate(headers):
            tk.Label(
                grid_frame,
                text=h_title,
                font=(self.font_sans, 8, "bold"),
                bg="#f1f5f9",
                fg="#0f172a",
                width=h_width,
                relief=tk.SOLID,
                bd=1,
                pady=6
            ).grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)

        # 26 строк данных
        for i in range(26):
            row_num = 1 + i

            # 0. Номер строки
            tk.Label(
                grid_frame,
                text=str(i + 1),
                font=(self.font_sans, 8, "bold"),
                bg="#f8fafc",
                fg="#64748b",
                width=5,
                relief=tk.SOLID,
                bd=1,
                pady=2
            ).grid(row=row_num, column=0, sticky="nsew", padx=1, pady=1)

            # Столбец 1
            e_col1 = ttk.Entry(
                grid_frame,
                textvariable=var_list[i][0],
                width=headers[1][1],
                font=(self.font_sans, 9)
            )
            e_col1.grid(row=row_num, column=1, sticky="nsew", padx=1, pady=1)
            widget_list.append(e_col1)
            self.attach_creative_context_menu(e_col1, page_num, i, col_idx=0)
            e_col1.bind("<KeyPress>", lambda ev, idx=i, p=page_num: self.on_creative_entry_keypress(ev, idx, p, col_idx=0), add="+")
            e_col1.bind("<<Paste>>", lambda ev, idx=i, p=page_num: self.on_creative_entry_paste(ev, idx, p, col_idx=0))
            e_col1.bind("<FocusOut>", lambda ev: self.auto_save_creative_achievements_data())

            # Столбец 2
            e_col2 = ttk.Entry(
                grid_frame,
                textvariable=var_list[i][1],
                width=headers[2][1],
                font=(self.font_sans, 9)
            )
            e_col2.grid(row=row_num, column=2, sticky="nsew", padx=1, pady=1)
            widget_list.append(e_col2)
            self.attach_creative_context_menu(e_col2, page_num, i, col_idx=1)
            e_col2.bind("<KeyPress>", lambda ev, idx=i, p=page_num: self.on_creative_entry_keypress(ev, idx, p, col_idx=1), add="+")
            e_col2.bind("<<Paste>>", lambda ev, idx=i, p=page_num: self.on_creative_entry_paste(ev, idx, p, col_idx=1))
            e_col2.bind("<FocusOut>", lambda ev: self.auto_save_creative_achievements_data())

    def attach_creative_context_menu(self, widget, page_num: int, idx: int, col_idx: int):
        """Контекстное меню для ячеек таблицы творческих достижений."""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(
            label="📋 Вставить строки из буфера (Ctrl+V)",
            command=lambda: self.paste_creative_achievements_from_clipboard(page_num=page_num, start_idx=idx, col_idx=col_idx)
        )
        if page_num == 1 and col_idx == 0:
            menu.add_command(
                label="👥 Заполнить обучающихся из группы",
                command=self.fill_creative_students_from_group
            )
        menu.add_separator()
        menu.add_command(label="Вырезать", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Копировать", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Вставить текст", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_command(label="Выделить всё", command=lambda: widget.select_range(0, tk.END))

        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)

    def on_creative_entry_keypress(self, event, idx: int, page_num: int, col_idx: int):
        """Обработка сочетаний клавиш в ячейках таблицы достижений."""
        if self.is_paste_event(event):
            self.paste_creative_achievements_from_clipboard(page_num=page_num, start_idx=idx, col_idx=col_idx, event=event)
            return "break"

        state = getattr(event, 'state', 0)
        is_ctrl = bool(state & 0x4) or bool(state & 0x8) or bool(state & 0x40000)
        code = getattr(event, 'keycode', 0)
        key = (getattr(event, 'keysym', '') or "").lower()

        if is_ctrl:
            if key in ('a', 'cyrillic_ef') or code == 65:
                self.select_all_widget(event.widget)
                return "break"
            elif key in ('c', 'cyrillic_es') or code == 67:
                self.copy_from_widget(event.widget)
                return "break"
            elif key in ('x', 'cyrillic_che') or code == 88:
                self.cut_from_widget(event.widget)
                return "break"

        return None

    def on_creative_entry_paste(self, event, idx: int, page_num: int, col_idx: int):
        """Перехват события <<Paste>> для ячеек таблицы достижений."""
        self.paste_creative_achievements_from_clipboard(page_num=page_num, start_idx=idx, col_idx=col_idx, event=event)
        return "break"

    def clear_creative_achievements_page(self, page_num: int):
        """Очищает данные страницы творческих достижений после подтверждения."""
        pg_title = f"Страница {page_num} (стр. {31 + page_num} журнала)"
        if not messagebox.askyesno("Очистить таблицу", f"Вы уверены, что хотите полностью очистить {pg_title}?"):
            return
        var_list = self.creative_achievements_p1_vars if page_num == 1 else self.creative_achievements_p2_vars
        for r in var_list:
            for v in r:
                v.set("")
        self.auto_save_creative_achievements_data()
        self.status_var.set(f"Таблица {pg_title} очищена.")

    def fill_creative_students_from_group(self):
        """Заполняет список фамилий и имен обучающихся на стр. 32 из данных группы (сентябрь)."""
        st_vars = []
        if hasattr(self, 'months_data') and "september" in self.months_data:
            st_vars = self.months_data["september"]["student_vars"]
        elif hasattr(self, 'september_student_vars'):
            st_vars = self.september_student_vars

        if not st_vars:
            self.status_var.set("Список обучающихся группы пуст.")
            return

        count = 0
        for i, s_var in enumerate(st_vars):
            if i >= 26:
                break
            name = s_var.get().strip()
            if name:
                self.creative_achievements_p1_vars[i][0].set(name)
                count += 1

        self.auto_save_creative_achievements_data()
        self.status_var.set(f"Заполнено {count} фамилий обучающихся из группы.")

    def auto_save_creative_achievements_data(self, *args):
        """Автоматически сохраняет все данные вкладки 'Творческие достижения' в config_data и на диск."""
        try:
            p1_data = [
                {
                    "student": r[0].get(),
                    "event": r[1].get()
                }
                for r in self.creative_achievements_p1_vars
            ]
            p2_data = [
                {
                    "results": r[0].get(),
                    "works": r[1].get()
                }
                for r in self.creative_achievements_p2_vars
            ]
            self.config_data["creative_achievements_p1"] = p1_data
            self.config_data["creative_achievements_p2"] = p2_data
            save_config(self.config_data)
        except Exception:
            pass

    def paste_creative_achievements_from_clipboard(self, page_num: int = 1, start_idx: int = 0, col_idx: int = 0, event=None):
        """Вставляет строки творческих достижений из буфера обмена, начиная со строки start_idx."""
        raw_text = self.get_clipboard_text()
        if not raw_text:
            self.status_var.set("Буфер обмена пуст или не содержит текст.")
            return "break"

        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
        if not lines:
            self.status_var.set("В буфере обмена нет строк для вставки.")
            return "break"

        var_list = self.creative_achievements_p1_vars if page_num == 1 else self.creative_achievements_p2_vars
        count = 0
        for offset, line in enumerate(lines):
            idx = start_idx + offset
            if idx >= 26:
                break
            parts = [p.strip() for p in line.split('\t')]
            if len(parts) >= 2:
                # Если 2 или более колонок в строке буфера: заполняем столбец 1 и столбец 2
                var_list[idx][0].set(parts[0])
                var_list[idx][1].set(parts[1])
            else:
                # 1 колонка
                if col_idx == 1:
                    var_list[idx][1].set(line)
                else:
                    if page_num == 1:
                        # Очищаем от ведущей нумерации ("1. Иванов Иван" -> "Иванов Иван")
                        clean_line = re.sub(r'^\d+[\.\)\s\-]+\s*', '', line)
                        var_list[idx][0].set(clean_line or line)
                    else:
                        var_list[idx][0].set(line)
            count += 1

        self.auto_save_creative_achievements_data()
        pg_str = f"стр. {31 + page_num}"
        self.status_var.set(f"Вставлено {count} записей достижений ({pg_str}, начиная со строки {start_idx + 1}).")
        return "break"

    def on_generate_creative_p1_click(self):
        """Формирует и сохраняет отдельный Excel-файл Творческих достижений (стр. 32 журнала)."""
        self.auto_save_creative_achievements_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Творческие_достижения_стр32_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Творческие достижения (стр. 32) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            achievements = [
                {
                    "student": r[0].get(),
                    "event": r[1].get()
                }
                for r in self.creative_achievements_p1_vars
            ]
            generate_excel_creative_achievements_p1(save_path, achievements=achievements)
            self.last_saved_creative_p1_file = save_path
            if hasattr(self, 'btn_open_creative_p1'):
                self.btn_open_creative_p1.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Файл 'Творческие достижения' (стр. 32) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении достижений (стр. 32).")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_creative_p1_file_click(self):
        """Открывает сгенерированный файл достижений (стр. 32)."""
        if not hasattr(self, 'last_saved_creative_p1_file') or not self.last_saved_creative_p1_file or not os.path.exists(self.last_saved_creative_p1_file):
            messagebox.showwarning("Файл не найден", "Созданный файл достижений (стр. 32) не найден на диске.")
            return
        self._open_file_system(self.last_saved_creative_p1_file)

    def on_generate_creative_p2_click(self):
        """Формирует и сохраняет отдельный Excel-файл Творческих достижений (стр. 33 журнала)."""
        self.auto_save_creative_achievements_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Творческие_достижения_стр33_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Творческие достижения (стр. 33) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            achievements = [
                {
                    "results": r[0].get(),
                    "works": r[1].get()
                }
                for r in self.creative_achievements_p2_vars
            ]
            generate_excel_creative_achievements_p2(save_path, achievements=achievements)
            self.last_saved_creative_p2_file = save_path
            if hasattr(self, 'btn_open_creative_p2'):
                self.btn_open_creative_p2.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Файл 'Творческие достижения' (стр. 33) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении достижений (стр. 33).")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_creative_p2_file_click(self):
        """Открывает сгенерированный файл достижений (стр. 33)."""
        if not hasattr(self, 'last_saved_creative_p2_file') or not self.last_saved_creative_p2_file or not os.path.exists(self.last_saved_creative_p2_file):
            messagebox.showwarning("Файл не найден", "Созданный файл достижений (стр. 33) не найден на диске.")
            return
        self._open_file_system(self.last_saved_creative_p2_file)

    def on_generate_creative_spread_click(self):
        """Формирует и сохраняет разворот Творческих достижений (стр. 32 и 33 журнала) в один файл."""
        self.auto_save_creative_achievements_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Творческие_достижения_Разворот_стр32-33_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Разворот творческих достижений (стр. 32 и 33) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            cr_p1 = [
                {
                    "student": r[0].get(),
                    "event": r[1].get()
                }
                for r in self.creative_achievements_p1_vars
            ]
            cr_p2 = [
                {
                    "results": r[0].get(),
                    "works": r[1].get()
                }
                for r in self.creative_achievements_p2_vars
            ]
            generate_excel_creative_achievements_spread(save_path, achievements_p1=cr_p1, achievements_p2=cr_p2)
            self.last_saved_creative_spread_file = save_path
            if hasattr(self, 'btn_open_creative_spread'):
                self.btn_open_creative_spread.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен разворот достижений: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Разворот творческих достижений (стр. 32 и 33) успешно сохранен:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении разворота достижений.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_creative_spread_file_click(self):
        """Открывает файл разворота творческих достижений."""
        if not hasattr(self, 'last_saved_creative_spread_file') or not self.last_saved_creative_spread_file or not os.path.exists(self.last_saved_creative_spread_file):
            messagebox.showwarning("Файл не найден", "Созданный файл разворота достижений не найден на диске.")
            return
        self._open_file_system(self.last_saved_creative_spread_file)

    def setup_students_list_tab(self, parent_tab, card_bg, next_tab_idx: int = 19):
        """Создает вкладку 'Список обучающихся' из 6 страниц (3 разворота, стр. 34-39 журнала)."""
        self.students_list_notebook = ttk.Notebook(parent_tab)
        self.students_list_notebook.pack(fill=tk.BOTH, expand=True)

        pages_info = [
            (1, 34, 1, "Стр. 1 (стр. 34: Уч. 1–10)"),
            (2, 35, 1, "Стр. 2 (стр. 35: Родители 1–10)"),
            (3, 36, 11, "Стр. 3 (стр. 36: Уч. 11–20)"),
            (4, 37, 11, "Стр. 4 (стр. 37: Родители 11–20)"),
            (5, 38, 21, "Стр. 5 (стр. 38: Уч. 21–30)"),
            (6, 39, 21, "Стр. 6 (стр. 39: Родители 21–30)"),
        ]

        for p_idx, j_no, start_no, tab_title in pages_info:
            sub_tab = tk.Frame(self.students_list_notebook, bg=card_bg, padx=8, pady=8)
            self.students_list_notebook.add(sub_tab, text=tab_title)

            if p_idx < 6:
                next_action = lambda idx=p_idx: self.students_list_notebook.select(idx)
                next_btn_text = f"К Странице {p_idx + 1} (стр. {j_no + 1}) ➔"
            else:
                next_action = lambda: self.notebook.select(next_tab_idx)
                next_btn_text = "Далее: Формирование Excel ➔"

            self.setup_single_students_list_page(
                parent_tab=sub_tab,
                card_bg=card_bg,
                page_num=p_idx,
                journal_page_no=j_no,
                start_student_no=start_no,
                next_action=next_action,
                next_btn_text=next_btn_text
            )

    def setup_single_students_list_page(
        self,
        parent_tab,
        card_bg,
        page_num: int,
        journal_page_no: int,
        start_student_no: int,
        next_action,
        next_btn_text: str
    ):
        """Создает одну страницу списка обучающихся (10 строк с полями ввода и быстрыми действиями)."""
        top_frame = tk.Frame(parent_tab, bg=card_bg)
        top_frame.pack(fill=tk.X, pady=(0, 6))

        info_bar = tk.Frame(top_frame, bg=card_bg)
        info_bar.pack(fill=tk.X, pady=(0, 4))

        end_no = start_student_no + 9
        if page_num in (1, 3, 5):
            page_title = f"Список обучающихся в объединении (ученики {start_student_no}–{end_no})"
        else:
            page_title = f"Сведения о родителях (ученики {start_student_no}–{end_no})"

        tk.Label(
            info_bar,
            text=f"Страница {page_num} (стр. {journal_page_no} журнала): {page_title} (10 строк)",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a"
        ).pack(side=tk.LEFT)

        tk.Label(
            info_bar,
            text="💡 Вставляйте строки через Ctrl+V, используйте Tab для перехода",
            font=(self.font_sans, 8, "italic"),
            bg=card_bg,
            fg="#64748b"
        ).pack(side=tk.RIGHT)

        # Панель кнопок действий
        btns_bar = tk.Frame(top_frame, bg=card_bg)
        btns_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btns_bar,
            text="📋 Вставить из буфера (Ctrl+V)",
            command=lambda: self.paste_students_list_from_clipboard(page_num=page_num, start_idx=0),
            font=(self.font_sans, 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        if page_num in (1, 3, 5):
            tk.Button(
                btns_bar,
                text=f"👥 Заполнить имена ({start_student_no}–{end_no})",
                command=lambda: self.fill_students_list_from_group(page_num),
                font=(self.font_sans, 8, "bold"),
                bg="#f0fdf4",
                fg="#166534",
                activebackground="#dcfce7",
                activeforeground="#15803d",
                relief=tk.SOLID,
                bd=1,
                padx=8,
                pady=4,
                cursor="hand2"
            ).pack(side=tk.LEFT, padx=(0, 6))

            if page_num == 1:
                tk.Button(
                    btns_bar,
                    text="👥 Заполнить всех обучающихся группы (стр. 1, 3, 5)",
                    command=self.fill_all_students_from_group,
                    font=(self.font_sans, 8, "bold"),
                    bg="#ecfdf5",
                    fg="#047857",
                    activebackground="#d1fae5",
                    activeforeground="#065f46",
                    relief=tk.SOLID,
                    bd=1,
                    padx=8,
                    pady=4,
                    cursor="hand2"
                ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btns_bar,
            text="🗑 Очистить страницу",
            command=lambda: self.clear_students_list_page(page_num),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#ef4444",
            activebackground="#fee2e2",
            activeforeground="#b91c1c",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btns_bar,
            text=next_btn_text,
            command=next_action,
            font=(self.font_sans, 8, "bold"),
            bg="#3b82f6",
            fg="#ffffff",
            activebackground="#2563eb",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.RIGHT)

        # Контейнер с прокруткой для таблицы (стабильный без зацикливаний)
        table_container = tk.Frame(parent_tab, bg=card_bg, relief=tk.SOLID, bd=1)
        table_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(table_container, bg=card_bg, highlightthickness=0)
        v_scroll = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(table_container, orient="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        grid_frame = tk.Frame(canvas, bg=card_bg)
        canvas.create_window((0, 0), window=grid_frame, anchor="nw")

        def _on_cfg(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        grid_frame.bind("<Configure>", _on_cfg)

        def _on_mousewheel(e):
            if sys.platform == "darwin":
                canvas.yview_scroll(int(-1 * e.delta), "units")
            else:
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        for w in (parent_tab, canvas, grid_frame):
            for seq in ("<Control-v>", "<Control-V>", "<Command-v>", "<Command-V>", "<Shift-Insert>"):
                w.bind(seq, lambda ev, p=page_num: self.paste_students_list_from_clipboard(page_num=p, start_idx=0, col_idx=0, event=ev))

        header_bg = "#e0e7ff"
        header_fg = "#1e1b4b"

        var_list = self.students_list_pages_vars[page_num]
        widgets_list = []

        if page_num in (1, 3, 5):
            # Шапка нечетной страницы (6 столбцов):
            # 1) № п/п
            # 2) Фамилия, имя обучающегося
            # 3) Год рождения
            # 4) Школа, класс
            # 5) Район
            # 6) Заключение врача о допуске к занятиям
            h_cols = [
                ("№\nп/п", 5),
                ("Фамилия, имя обучающегося\n(Ctrl+V для вставки)", 30),
                ("Год\nрождения", 11),
                ("Школа, класс", 16),
                ("Район", 16),
                ("Заключение врача о допуске к занятиям", 30)
            ]
            for col_i, (col_name, col_w) in enumerate(h_cols):
                lbl = tk.Label(
                    grid_frame,
                    text=col_name,
                    font=(self.font_sans, 9, "bold"),
                    bg=header_bg,
                    fg=header_fg,
                    width=col_w,
                    relief=tk.SOLID,
                    bd=1,
                    pady=4
                )
                lbl.grid(row=0, column=col_i, sticky="nsew", padx=1, pady=1)

            # 10 строк таблицы
            for row_i in range(10):
                student_no = start_student_no + row_i
                v_st, v_by, v_sc, v_ds, v_dc = var_list[row_i]

                # Столбец 0: № п/п
                lbl_num = tk.Label(
                    grid_frame,
                    text=str(student_no),
                    font=(self.font_sans, 9, "bold"),
                    bg="#f1f5f9",
                    fg="#334155",
                    width=5,
                    relief=tk.SOLID,
                    bd=1
                )
                lbl_num.grid(row=row_i + 1, column=0, sticky="nsew", padx=1, pady=1)

                # Столбец 1: Фамилия, имя обучающегося
                e_st = ttk.Entry(grid_frame, textvariable=v_st, font=(self.font_sans, 9), width=30)
                e_st.grid(row=row_i + 1, column=1, sticky="nsew", padx=1, pady=1)

                # Столбец 2: Год рождения
                e_by = ttk.Entry(grid_frame, textvariable=v_by, font=(self.font_sans, 9), width=11, justify="center")
                e_by.grid(row=row_i + 1, column=2, sticky="nsew", padx=1, pady=1)

                # Столбец 3: Школа, класс
                e_sc = ttk.Entry(grid_frame, textvariable=v_sc, font=(self.font_sans, 9), width=16, justify="center")
                e_sc.grid(row=row_i + 1, column=3, sticky="nsew", padx=1, pady=1)

                # Столбец 4: Район
                e_ds = ttk.Entry(grid_frame, textvariable=v_ds, font=(self.font_sans, 9), width=16, justify="center")
                e_ds.grid(row=row_i + 1, column=4, sticky="nsew", padx=1, pady=1)

                # Столбец 5: Заключение врача о допуске к занятиям
                e_dc = ttk.Entry(grid_frame, textvariable=v_dc, font=(self.font_sans, 9), width=30, justify="center")
                e_dc.grid(row=row_i + 1, column=5, sticky="nsew", padx=1, pady=1)

                row_widgets = (e_st, e_by, e_sc, e_ds, e_dc)
                widgets_list.append(row_widgets)

                for col_idx, entry_w in enumerate(row_widgets):
                    self.attach_students_list_context_menu(entry_w, idx=row_i, page_num=page_num, col_idx=col_idx)
                    entry_w.bind("<KeyPress>", lambda ev, r=row_i, c=col_idx: self.on_students_list_entry_keypress(ev, r, page_num, c))
                    entry_w.bind("<<Paste>>", lambda ev, r=row_i, c=col_idx: self.on_students_list_entry_paste(ev, r, page_num, c))
                    entry_w.bind("<FocusOut>", lambda ev: self.auto_save_students_list_data())

        else:
            # Шапка четной страницы (5 столбцов + номер):
            # 1) Домашний адрес, телефон
            # 2) Фамилия, имя, отчество родителей, телефон
            # 3) Дата вступления в объединение
            # 4) Когда и почему выбыл
            # 5) Примечания
            h_cols = [
                ("№", 4),
                ("Домашний адрес, телефон", 28),
                ("Фамилия, имя, отчество родителей, телефон", 32),
                ("Дата\nвступления\nв объед.", 14),
                ("Когда и почему\nвыбыл", 16),
                ("Примечания", 18)
            ]
            for col_i, (col_name, col_w) in enumerate(h_cols):
                lbl = tk.Label(
                    grid_frame,
                    text=col_name,
                    font=(self.font_sans, 9, "bold"),
                    bg=header_bg,
                    fg=header_fg,
                    width=col_w,
                    relief=tk.SOLID,
                    bd=1,
                    pady=4
                )
                lbl.grid(row=0, column=col_i, sticky="nsew", padx=1, pady=1)

            # 10 строк таблицы
            for row_i in range(10):
                student_no = start_student_no + row_i
                v_ap, v_pi, v_jd, v_li, v_nt = var_list[row_i]

                # Столбец 0: № (для ориентации педагога)
                lbl_num = tk.Label(
                    grid_frame,
                    text=str(student_no),
                    font=(self.font_sans, 9, "bold"),
                    bg="#f1f5f9",
                    fg="#334155",
                    width=4,
                    relief=tk.SOLID,
                    bd=1
                )
                lbl_num.grid(row=row_i + 1, column=0, sticky="nsew", padx=1, pady=1)

                # Столбец 1: Домашний адрес, телефон
                e_ap = ttk.Entry(grid_frame, textvariable=v_ap, font=(self.font_sans, 9), width=28)
                e_ap.grid(row=row_i + 1, column=1, sticky="nsew", padx=1, pady=1)

                # Столбец 2: Фамилия, имя, отчество родителей, телефон
                e_pi = ttk.Entry(grid_frame, textvariable=v_pi, font=(self.font_sans, 9), width=32)
                e_pi.grid(row=row_i + 1, column=2, sticky="nsew", padx=1, pady=1)

                # Столбец 3: Дата вступления в объединение
                e_jd = ttk.Entry(grid_frame, textvariable=v_jd, font=(self.font_sans, 9), width=14, justify="center")
                e_jd.grid(row=row_i + 1, column=3, sticky="nsew", padx=1, pady=1)

                # Столбец 4: Когда и почему выбыл
                e_li = ttk.Entry(grid_frame, textvariable=v_li, font=(self.font_sans, 9), width=16)
                e_li.grid(row=row_i + 1, column=4, sticky="nsew", padx=1, pady=1)

                # Столбец 5: Примечания
                e_nt = ttk.Entry(grid_frame, textvariable=v_nt, font=(self.font_sans, 9), width=18)
                e_nt.grid(row=row_i + 1, column=5, sticky="nsew", padx=1, pady=1)

                row_widgets = (e_ap, e_pi, e_jd, e_li, e_nt)
                widgets_list.append(row_widgets)

                for col_idx, entry_w in enumerate(row_widgets):
                    self.attach_students_list_context_menu(entry_w, idx=row_i, page_num=page_num, col_idx=col_idx)
                    entry_w.bind("<KeyPress>", lambda ev, r=row_i, c=col_idx: self.on_students_list_entry_keypress(ev, r, page_num, c))
                    entry_w.bind("<<Paste>>", lambda ev, r=row_i, c=col_idx: self.on_students_list_entry_paste(ev, r, page_num, c))
                    entry_w.bind("<FocusOut>", lambda ev: self.auto_save_students_list_data())

        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

        self.students_list_pages_widgets[page_num] = widgets_list

    def attach_students_list_context_menu(self, widget, idx: int, page_num: int, col_idx: int):
        """Контекстное меню для ячеек таблицы списка обучающихся."""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(
            label="📋 Вставить строки из буфера (Ctrl+V)",
            command=lambda: self.paste_students_list_from_clipboard(page_num=page_num, start_idx=idx, col_idx=col_idx)
        )
        if page_num in (1, 3, 5) and col_idx == 0:
            menu.add_command(
                label="👥 Заполнить обучающихся из группы",
                command=lambda: self.fill_students_list_from_group(page_num)
            )
        menu.add_separator()
        menu.add_command(label="Вырезать", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Копировать", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Вставить текст", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_command(label="Выделить всё", command=lambda: widget.select_range(0, tk.END))

        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)

    def on_students_list_entry_keypress(self, event, idx: int, page_num: int, col_idx: int):
        """Обработка сочетаний клавиш в ячейках таблицы списка обучающихся."""
        if self.is_paste_event(event):
            self.paste_students_list_from_clipboard(page_num=page_num, start_idx=idx, col_idx=col_idx, event=event)
            return "break"

        state = getattr(event, 'state', 0)
        is_ctrl = bool(state & 0x4) or bool(state & 0x8) or bool(state & 0x40000)
        code = getattr(event, 'keycode', 0)
        key = (getattr(event, 'keysym', '') or "").lower()

        if is_ctrl:
            if key in ('a', 'cyrillic_ef') or code == 65:
                self.select_all_widget(event.widget)
                return "break"
            elif key in ('c', 'cyrillic_es') or code == 67:
                self.copy_from_widget(event.widget)
                return "break"
            elif key in ('x', 'cyrillic_che') or code == 88:
                self.cut_from_widget(event.widget)
                return "break"

        return None

    def on_students_list_entry_paste(self, event, idx: int, page_num: int, col_idx: int):
        """Перехват события <<Paste>> для ячеек таблицы списка обучающихся."""
        self.paste_students_list_from_clipboard(page_num=page_num, start_idx=idx, col_idx=col_idx, event=event)
        return "break"

    def clear_students_list_page(self, page_num: int):
        """Очищает данные страницы списка обучающихся после подтверждения."""
        pg_title = f"Страница {page_num} (стр. {33 + page_num} журнала)"
        if not messagebox.askyesno("Очистить таблицу", f"Вы уверены, что хотите полностью очистить {pg_title}?"):
            return
        var_list = self.students_list_pages_vars.get(page_num, [])
        for r in var_list:
            for v in r:
                v.set("")
        self.auto_save_students_list_data()
        self.status_var.set(f"Таблица {pg_title} очищена.")

    def fill_students_list_from_group(self, page_num: int):
        """Заполняет фамилии и имена обучающихся на странице (1, 3 или 5) из группы."""
        st_vars = []
        if hasattr(self, 'months_data') and "september" in self.months_data:
            st_vars = self.months_data["september"]["student_vars"]
        elif hasattr(self, 'september_student_vars'):
            st_vars = self.september_student_vars

        if not st_vars:
            self.status_var.set("Список обучающихся группы пуст.")
            messagebox.showwarning("Список пуст", "В группе пока нет заполненных обучающихся.")
            return

        offset = 0 if page_num == 1 else (10 if page_num == 3 else 20)
        var_list = self.students_list_pages_vars.get(page_num, [])

        count = 0
        for i in range(10):
            st_idx = offset + i
            if st_idx < len(st_vars):
                name = st_vars[st_idx].get().strip()
                if name:
                    var_list[i][0].set(name)
                    count += 1

        self.auto_save_students_list_data()
        self.status_var.set(f"Заполнено {count} фамилий обучающихся на стр. {page_num}.")

    def fill_all_students_from_group(self):
        """Заполняет все 3 страницы списка обучающихся (1, 3, 5) из списка группы (до 30 человек)."""
        st_vars = []
        if hasattr(self, 'months_data') and "september" in self.months_data:
            st_vars = self.months_data["september"]["student_vars"]
        elif hasattr(self, 'september_student_vars'):
            st_vars = self.september_student_vars

        if not st_vars:
            self.status_var.set("Список обучающихся группы пуст.")
            messagebox.showwarning("Список пуст", "В группе пока нет заполненных обучающихся.")
            return

        total_filled = 0
        for p_idx, offset in ((1, 0), (3, 10), (5, 20)):
            var_list = self.students_list_pages_vars.get(p_idx, [])
            for i in range(10):
                st_idx = offset + i
                if st_idx < len(st_vars):
                    name = st_vars[st_idx].get().strip()
                    if name:
                        var_list[i][0].set(name)
                        total_filled += 1

        self.auto_save_students_list_data()
        self.status_var.set(f"Заполнено {total_filled} обучающихся на страницах 1, 3 и 5.")

    def schedule_auto_save_students_list(self):
        """Отложенное сохранение данных списка обучающихся (debounced на 500 мс), устраняет лаги при наборе."""
        if hasattr(self, '_students_save_timer') and self._students_save_timer:
            try:
                self.after_cancel(self._students_save_timer)
            except Exception:
                pass
        self._students_save_timer = self.after(500, self.auto_save_students_list_data)

    def auto_save_students_list_data(self, *args):
        """Автоматически сохраняет все данные вкладки 'Список обучающихся' в config_data и на диск."""
        if hasattr(self, '_students_save_timer') and self._students_save_timer:
            try:
                self.after_cancel(self._students_save_timer)
            except Exception:
                pass
            self._students_save_timer = None
        try:
            for p in (1, 3, 5):
                p_data = [
                    {
                        "student": r[0].get(),
                        "birth_year": r[1].get(),
                        "school_class": r[2].get(),
                        "district": r[3].get(),
                        "doctor_conclusion": r[4].get()
                    }
                    for r in self.students_list_pages_vars[p]
                ]
                self.config_data[f"students_list_p{p}"] = p_data

            for p in (2, 4, 6):
                p_data = [
                    {
                        "address_phone": r[0].get(),
                        "parents_info": r[1].get(),
                        "join_date": r[2].get(),
                        "leave_info": r[3].get(),
                        "notes": r[4].get()
                    }
                    for r in self.students_list_pages_vars[p]
                ]
                self.config_data[f"students_list_p{p}"] = p_data

            save_config(self.config_data)
        except Exception:
            pass

    def paste_students_list_from_clipboard(self, page_num: int = 1, start_idx: int = 0, col_idx: int = 0, event=None):
        """Вставляет строки в таблицу списка обучающихся из буфера обмена."""
        raw_text = self.get_clipboard_text()
        if not raw_text:
            self.status_var.set("Буфер обмена пуст или не содержит текст.")
            return "break"

        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
        if not lines:
            self.status_var.set("В буфере обмена нет строк для вставки.")
            return "break"

        var_list = self.students_list_pages_vars.get(page_num, [])
        count = 0
        num_cols = 5  # У каждой страницы ровно 5 редактируемых столбцов

        for offset, line in enumerate(lines):
            idx = start_idx + offset
            if idx >= 10:
                break
            parts = [p.strip() for p in line.split('\t')]

            if len(parts) > 1:
                # Множество колонок (из Excel / таблицы)
                for c_offset, val in enumerate(parts):
                    target_c = col_idx + c_offset
                    if target_c < num_cols:
                        if page_num in (1, 3, 5) and target_c == 0:
                            clean_val = re.sub(r'^\d+[\.\)\s\-]+\s*', '', val)
                            var_list[idx][target_c].set(clean_val or val)
                        else:
                            var_list[idx][target_c].set(val)
            else:
                # Одиночная строка текста
                if page_num in (1, 3, 5) and col_idx == 0:
                    clean_line = re.sub(r'^\d+[\.\)\s\-]+\s*', '', line)
                    var_list[idx][0].set(clean_line or line)
                else:
                    var_list[idx][col_idx].set(line)

            count += 1

        self.auto_save_students_list_data()
        pg_str = f"стр. {33 + page_num}"
        self.status_var.set(f"Вставлено {count} записей ({pg_str}, начиная со строки {start_idx + 1}).")
        return "break"

    def on_generate_students_list_page_click(self, page_num: int):
        """Формирует и сохраняет отдельный лист страницы списка обучающихся в Excel."""
        self.auto_save_students_list_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        journal_p = 33 + page_num
        default_name = f"Список_обучающихся_стр{journal_p}_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title=f"Сохранить Страницу {page_num} (стр. {journal_p} журнала) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            var_list = self.students_list_pages_vars[page_num]
            if page_num in (1, 3, 5):
                items = [
                    {
                        "student": r[0].get(),
                        "birth_year": r[1].get(),
                        "school_class": r[2].get(),
                        "district": r[3].get(),
                        "doctor_conclusion": r[4].get()
                    }
                    for r in var_list
                ]
            else:
                items = [
                    {
                        "address_phone": r[0].get(),
                        "parents_info": r[1].get(),
                        "join_date": r[2].get(),
                        "leave_info": r[3].get(),
                        "notes": r[4].get()
                    }
                    for r in var_list
                ]
            generate_excel_students_list_page(save_path, page_num=page_num, items=items)
            self.last_saved_students_page_files[page_num] = save_path
            if hasattr(self, 'btn_open_students_pages') and page_num in self.btn_open_students_pages:
                self.btn_open_students_pages[page_num].config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Файл 'Список обучающихся' (стр. {journal_p}) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set(f"Ошибка при сохранении страницы {page_num}.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_students_list_page_file_click(self, page_num: int):
        """Открывает сгенерированный файл страницы списка обучающихся."""
        file_path = self.last_saved_students_page_files.get(page_num)
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("Файл не найден", f"Созданный файл страницы {page_num} не найден на диске.")
            return
        self._open_file_system(file_path)

    def on_generate_students_list_spread_click(self, spread_idx: int):
        """Формирует и сохраняет разворот списка обучающихся (разворот 1: стр. 34-35, 2: 36-37, 3: 38-39)."""
        self.auto_save_students_list_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        if spread_idx == 1:
            p_odd, p_even = 34, 35
            odd_key, even_key = 1, 2
        elif spread_idx == 2:
            p_odd, p_even = 36, 37
            odd_key, even_key = 3, 4
        else:
            p_odd, p_even = 38, 39
            odd_key, even_key = 5, 6

        default_name = f"Список_обучающихся_Разворот{spread_idx}_стр{p_odd}-{p_even}_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title=f"Сохранить Разворот {spread_idx} (стр. {p_odd} и {p_even}) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            items_odd = [
                {
                    "student": r[0].get(),
                    "birth_year": r[1].get(),
                    "school_class": r[2].get(),
                    "district": r[3].get(),
                    "doctor_conclusion": r[4].get()
                }
                for r in self.students_list_pages_vars[odd_key]
            ]
            items_even = [
                {
                    "address_phone": r[0].get(),
                    "parents_info": r[1].get(),
                    "join_date": r[2].get(),
                    "leave_info": r[3].get(),
                    "notes": r[4].get()
                }
                for r in self.students_list_pages_vars[even_key]
            ]
            generate_excel_students_list_spread(save_path, spread_idx=spread_idx, items_odd=items_odd, items_even=items_even)
            self.last_saved_students_spread_files[spread_idx] = save_path
            if hasattr(self, 'btn_open_students_spreads') and spread_idx in self.btn_open_students_spreads:
                self.btn_open_students_spreads[spread_idx].config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен разворот: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Разворот {spread_idx} (стр. {p_odd}-{p_even}) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set(f"Ошибка при сохранении разворота {spread_idx}.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_students_list_spread_file_click(self, spread_idx: int):
        """Открывает файл сохраненного разворота списка обучающихся."""
        file_path = self.last_saved_students_spread_files.get(spread_idx)
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("Файл не найден", f"Созданный файл разворота {spread_idx} не найден на диске.")
            return
        self._open_file_system(file_path)

    def on_generate_students_list_all_click(self):
        """Формирует и сохраняет все 6 страниц списка обучающихся (стр. 34-39 журнала) в один файл Excel."""
        self.auto_save_students_list_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Список_обучающихся_Все_6_страниц_стр34-39_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Все 6 страниц 'Список обучающихся' (стр. 34-39) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            all_pages_data = {}
            for p in (1, 3, 5):
                all_pages_data[f"students_list_p{p}"] = [
                    {
                        "student": r[0].get(),
                        "birth_year": r[1].get(),
                        "school_class": r[2].get(),
                        "district": r[3].get(),
                        "doctor_conclusion": r[4].get()
                    }
                    for r in self.students_list_pages_vars[p]
                ]
            for p in (2, 4, 6):
                all_pages_data[f"students_list_p{p}"] = [
                    {
                        "address_phone": r[0].get(),
                        "parents_info": r[1].get(),
                        "join_date": r[2].get(),
                        "leave_info": r[3].get(),
                        "notes": r[4].get()
                    }
                    for r in self.students_list_pages_vars[p]
                ]

            generate_excel_students_list_all(save_path, all_pages=all_pages_data)
            self.last_saved_students_all_file = save_path
            if hasattr(self, 'btn_open_students_all'):
                self.btn_open_students_all.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен полный список обучающихся: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Все 6 страниц 'Список обучающихся' (стр. 34-39) успешно сохранены:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении полного списка обучающихся.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_students_list_all_file_click(self):
        """Открывает файл всех 6 страниц списка обучающихся."""
        if not hasattr(self, 'last_saved_students_all_file') or not self.last_saved_students_all_file or not os.path.exists(self.last_saved_students_all_file):
            messagebox.showwarning("Файл не найден", "Созданный файл всех страниц списка обучающихся не найден на диске.")
            return
        self._open_file_system(self.last_saved_students_all_file)

    def setup_safety_briefing_tab(self, parent_tab, card_bg, next_tab_idx: int = 20):
        """Создает вкладку 'Инструктаж по ТБ' из 2 страниц (разворот: стр. 38 и 39 журнала)."""
        self.safety_briefing_notebook = ttk.Notebook(parent_tab)
        self.safety_briefing_notebook.pack(fill=tk.BOTH, expand=True)

        pages_info = [
            (1, 38, "Страница 1 (стр. 38 журнала)"),
            (2, 39, "Страница 2 (стр. 39 журнала)"),
        ]

        for p_idx, j_no, tab_title in pages_info:
            sub_tab = tk.Frame(self.safety_briefing_notebook, bg=card_bg, padx=8, pady=8)
            self.safety_briefing_notebook.add(sub_tab, text=tab_title)

            if p_idx == 1:
                next_action = lambda: self.safety_briefing_notebook.select(1)
                next_btn_text = "К Странице 2 (стр. 39) ➔"
            else:
                next_action = lambda: self.notebook.select(next_tab_idx)
                next_btn_text = "Далее: Формирование Excel ➔"

            var_list = self.safety_briefing_p1_vars if p_idx == 1 else self.safety_briefing_p2_vars
            widget_list = self.safety_briefing_p1_widgets if p_idx == 1 else self.safety_briefing_p2_widgets

            self.setup_single_safety_briefing_page(
                parent_tab=sub_tab,
                card_bg=card_bg,
                page_num=p_idx,
                journal_page_no=j_no,
                var_list=var_list,
                widget_list=widget_list,
                next_action=next_action,
                next_btn_text=next_btn_text
            )

    def setup_single_safety_briefing_page(
        self,
        parent_tab,
        card_bg,
        page_num: int,
        journal_page_no: int,
        var_list: list,
        widget_list: list,
        next_action,
        next_btn_text: str
    ):
        """Создает одну страницу инструктажа по технике безопасности: 5 колонок, 28 строк с прокруткой и действиями."""
        top_frame = tk.Frame(parent_tab, bg=card_bg)
        top_frame.pack(fill=tk.X, pady=(0, 6))

        # Информационная строка
        info_bar = tk.Frame(top_frame, bg=card_bg)
        info_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Label(
            info_bar,
            text=f"Страница {page_num} (стр. {journal_page_no} журнала): Список обучающихся, прошедших инструктаж по ТБ (28 строк)",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a"
        ).pack(side=tk.LEFT)

        tk.Label(
            info_bar,
            text="💡 Поддерживается Ctrl+V из Excel, кнопки быстрого автозаполнения, Tab для перехода",
            font=(self.font_sans, 8, "italic"),
            bg=card_bg,
            fg="#64748b"
        ).pack(side=tk.RIGHT)

        # Панель быстрых действий
        btns_bar = tk.Frame(top_frame, bg=card_bg)
        btns_bar.pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btns_bar,
            text="📋 Вставить из буфера (Ctrl+V)",
            command=lambda: self.paste_safety_briefing_from_clipboard(page_num=page_num, start_idx=0, col_idx=0),
            font=(self.font_sans, 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            btns_bar,
            text="👥 ФИО из списка обучающихся",
            command=lambda: self.fill_safety_students_from_list(page_num),
            font=(self.font_sans, 8, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            btns_bar,
            text="📅 Дата всем",
            command=lambda: self.fill_safety_date_for_all(page_num),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#0f172a",
            activebackground="#e2e8f0",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            btns_bar,
            text="📝 Содержание всем",
            command=lambda: self.fill_safety_content_for_all(page_num),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#0f172a",
            activebackground="#e2e8f0",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            btns_bar,
            text="✍ Подпись педагога всем",
            command=lambda: self.fill_safety_signature_for_all(page_num),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#0f172a",
            activebackground="#e2e8f0",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            btns_bar,
            text="🗑 Очистить",
            command=lambda: self.clear_safety_briefing_page(page_num),
            font=(self.font_sans, 8),
            bg="#f1f5f9",
            fg="#ef4444",
            activebackground="#fee2e2",
            activeforeground="#b91c1c",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            btns_bar,
            text=next_btn_text,
            command=next_action,
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7" if page_num == 1 else "#15803d",
            fg="#ffffff",
            activebackground="#0369a1" if page_num == 1 else "#166534",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=4,
            cursor="hand2"
        ).pack(side=tk.RIGHT)

        # Контейнер таблицы с прокруткой (Canvas)
        table_container = tk.Frame(parent_tab, bg=card_bg, relief=tk.SOLID, bd=1)
        table_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(table_container, bg=card_bg, highlightthickness=0)
        v_scroll = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(table_container, orient="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        grid_frame = tk.Frame(canvas, bg=card_bg)
        canvas.create_window((0, 0), window=grid_frame, anchor="nw")

        def _on_config(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        grid_frame.bind("<Configure>", _on_config)

        def _on_mousewheel(event):
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")
        canvas.bind_all("<Button-4>", _on_mousewheel, add="+")
        canvas.bind_all("<Button-5>", _on_mousewheel, add="+")

        # Шапка таблицы (5 столбцов)
        headers = [
            ("№ п/п", 5),
            ("Фамилия, имя обучающегося", 28),
            ("Дата проведения\nинструктажа", 16),
            ("Краткое содержание инструктажа", 36),
            ("Подпись проводившего\nинструктаж (разборчиво)", 22)
        ]

        for c_idx, (h_title, w_char) in enumerate(headers):
            h_lbl = tk.Label(
                grid_frame,
                text=h_title,
                font=(self.font_sans, 8, "bold"),
                bg="#f1f5f9",
                fg="#1e293b",
                relief=tk.RIDGE,
                bd=1,
                padx=4,
                pady=6,
                width=w_char,
                anchor="center",
                justify=tk.CENTER
            )
            h_lbl.grid(row=0, column=c_idx, sticky="nsew", padx=1, pady=1)

        # 28 строк таблицы
        widget_list.clear()
        for r_idx in range(28):
            row_widgets = []

            # 1. № п/п
            lbl_num = tk.Label(
                grid_frame,
                text=str(r_idx + 1),
                font=(self.font_sans, 9, "bold"),
                bg="#f8fafc" if r_idx % 2 == 0 else "#ffffff",
                fg="#64748b",
                relief=tk.SOLID,
                bd=1,
                padx=2,
                pady=3,
                width=5,
                anchor="center"
            )
            lbl_num.grid(row=r_idx + 1, column=0, sticky="nsew", padx=1, pady=1)

            # 4 поля ввода:
            col_widths = [28, 14, 36, 20]
            for c_offset in range(4):
                var = var_list[r_idx][c_offset]
                e = ttk.Entry(grid_frame, textvariable=var, font=(self.font_sans, 9), width=col_widths[c_offset])
                e.grid(row=r_idx + 1, column=c_offset + 1, sticky="nsew", padx=1, pady=1)

                self.attach_context_menu(e)
                e.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
                e.bind("<Control-v>", lambda ev, r=r_idx, c=c_offset: self.paste_safety_briefing_from_clipboard(page_num=page_num, start_idx=r, col_idx=c, event=ev))
                e.bind("<Control-V>", lambda ev, r=r_idx, c=c_offset: self.paste_safety_briefing_from_clipboard(page_num=page_num, start_idx=r, col_idx=c, event=ev))

                row_widgets.append(e)

            widget_list.append(row_widgets)

        # Навигация Tab/Shift+Tab и стрелки вверх/вниз
        for r_idx, row in enumerate(widget_list):
            for c_offset, w in enumerate(row):
                if r_idx < 27:
                    next_down = widget_list[r_idx + 1][c_offset]
                    w.bind("<Down>", lambda e, target=next_down: target.focus_set())
                if r_idx > 0:
                    prev_up = widget_list[r_idx - 1][c_offset]
                    w.bind("<Up>", lambda e, target=prev_up: target.focus_set())

        grid_frame.columnconfigure(0, weight=0)
        grid_frame.columnconfigure(1, weight=1)
        grid_frame.columnconfigure(2, weight=0)
        grid_frame.columnconfigure(3, weight=2)
        grid_frame.columnconfigure(4, weight=1)

    def fill_safety_students_from_list(self, page_num: int):
        """Заполняет фамилии и имена обучающихся из вкладки 'Список обучающихся' или 'Сентябрь'."""
        students = []
        # Проверяем список обучающихся
        if hasattr(self, 'students_list_pages_vars'):
            for p in (1, 3, 5):
                for row in self.students_list_pages_vars.get(p, []):
                    st_val = row[0].get().strip()
                    if st_val:
                        students.append(st_val)

        # Если в списке обучающихся пусто, берем из Сентября
        if not students and hasattr(self, 'months_data') and "september" in self.months_data:
            for sv in self.months_data["september"]["student_vars"]:
                st_val = sv.get().strip()
                if st_val:
                    students.append(st_val)

        if not students:
            messagebox.showinfo("Нет данных", "Список обучающихся не найден ни в Сентябре, ни во вкладке 'Список обучающихся'. Сначала введите фамилии обучающихся.")
            return

        var_list = self.safety_briefing_p1_vars if page_num == 1 else self.safety_briefing_p2_vars
        count = 0
        for i, st_name in enumerate(students):
            if i >= 28:
                break
            var_list[i][0].set(st_name)
            count += 1

        self.schedule_auto_save_safety_briefing()
        self.status_var.set(f"Вставлено {count} обучающихся на страницу {page_num} инструктажа по ТБ.")
        messagebox.showinfo("Успешно", f"Вставлено {count} обучающихся на страницу {page_num}.")

    def fill_safety_date_for_all(self, page_num: int):
        """Проставляет дату инструктажа во все строки страницы."""
        import tkinter.simpledialog as sd
        def_date = ""
        if hasattr(self, 'months_data') and "september" in self.months_data:
            d_vars = self.months_data["september"]["date_vars"]
            if d_vars and d_vars[0].get().strip():
                def_date = d_vars[0].get().strip()
        if not def_date:
            def_date = "02.09.2024"

        ans = sd.askstring("Дата инструктажа", f"Введите дату проведения инструктажа для страницы {page_num}:", initialvalue=def_date, parent=self)
        if ans is None:
            return
        ans = ans.strip()
        if not ans:
            return

        var_list = self.safety_briefing_p1_vars if page_num == 1 else self.safety_briefing_p2_vars
        count = 0
        for r in var_list:
            if r[0].get().strip():
                r[1].set(ans)
                count += 1

        if count == 0:
            for r in var_list:
                r[1].set(ans)
            count = 28

        self.schedule_auto_save_safety_briefing()
        self.status_var.set(f"Дата '{ans}' проставлена в {count} строк на странице {page_num}.")

    def fill_safety_content_for_all(self, page_num: int):
        """Проставляет краткое содержание инструктажа во все строки страницы."""
        import tkinter.simpledialog as sd
        def_theme = "Вводный инструктаж по ТБ и правилам поведения"
        ans = sd.askstring("Тема инструктажа", f"Введите краткое содержание инструктажа для страницы {page_num}:", initialvalue=def_theme, parent=self)
        if ans is None:
            return
        ans = ans.strip()
        if not ans:
            return

        var_list = self.safety_briefing_p1_vars if page_num == 1 else self.safety_briefing_p2_vars
        count = 0
        for r in var_list:
            if r[0].get().strip():
                r[2].set(ans)
                count += 1

        if count == 0:
            for r in var_list:
                r[2].set(ans)
            count = 28

        self.schedule_auto_save_safety_briefing()
        self.status_var.set(f"Содержание инструктажа проставлено в {count} строк на странице {page_num}.")

    def fill_safety_signature_for_all(self, page_num: int):
        """Проставляет подпись (ФИО педагога) во все строки страницы."""
        teacher = ""
        if hasattr(self, 'teacher_var') and self.teacher_var.get().strip():
            teacher = self.teacher_var.get().strip()
        elif hasattr(self, 'rukovoditel_var') and self.rukovoditel_var.get().strip():
            teacher = self.rukovoditel_var.get().strip()

        if not teacher:
            import tkinter.simpledialog as sd
            teacher = sd.askstring("Подпись педагога", "Введите ФИО проводящего инструктаж (разборчиво):", parent=self)
            if not teacher:
                return
            teacher = teacher.strip()

        var_list = self.safety_briefing_p1_vars if page_num == 1 else self.safety_briefing_p2_vars
        count = 0
        for r in var_list:
            if r[0].get().strip():
                r[3].set(teacher)
                count += 1

        if count == 0:
            for r in var_list:
                r[3].set(teacher)
            count = 28

        self.schedule_auto_save_safety_briefing()
        self.status_var.set(f"Подпись педагога '{teacher}' проставлена в {count} строк на странице {page_num}.")

    def clear_safety_briefing_page(self, page_num: int):
        """Очищает страницу инструктажа по технике безопасности после подтверждения."""
        if not messagebox.askyesno("Очистить страницу", f"Вы уверены, что хотите полностью очистить все 28 строк на странице {page_num}?"):
            return
        var_list = self.safety_briefing_p1_vars if page_num == 1 else self.safety_briefing_p2_vars
        for r in var_list:
            for v in r:
                v.set("")
        self.auto_save_safety_briefing_data()
        self.status_var.set(f"Страница {page_num} инструктажа по ТБ очищена.")

    def paste_safety_briefing_from_clipboard(self, page_num: int = 1, start_idx: int = 0, col_idx: int = 0, event=None):
        """Вставляет строки в таблицу инструктажа по ТБ из буфера обмена."""
        raw_text = self.get_clipboard_text()
        if not raw_text:
            self.status_var.set("Буфер обмена пуст или не содержит текст.")
            return "break"

        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
        if not lines:
            self.status_var.set("В буфере обмена нет строк для вставки.")
            return "break"

        var_list = self.safety_briefing_p1_vars if page_num == 1 else self.safety_briefing_p2_vars
        num_cols = 4

        for offset, line in enumerate(lines):
            idx = start_idx + offset
            if idx >= 28:
                break
            parts = [p.strip() for p in line.split('\t')]

            if len(parts) > 1:
                if col_idx == 0 and len(parts) >= 5 and re.match(r'^\d+$', parts[0]):
                    parts = parts[1:]

                for c_offset, val in enumerate(parts):
                    target_c = col_idx + c_offset
                    if target_c < num_cols:
                        if target_c == 0:
                            clean_val = re.sub(r'^\d+[\.\)\s\-]+\s*', '', val)
                            var_list[idx][target_c].set(clean_val or val)
                        else:
                            var_list[idx][target_c].set(val)
            else:
                if col_idx == 0:
                    clean_val = re.sub(r'^\d+[\.\)\s\-]+\s*', '', line)
                    var_list[idx][0].set(clean_val or line)
                elif col_idx < num_cols:
                    var_list[idx][col_idx].set(line)

        self.schedule_auto_save_safety_briefing()
        self.status_var.set(f"Вставлено {len(lines)} строк на страницу {page_num} инструктажа по ТБ.")
        return "break"

    def schedule_auto_save_safety_briefing(self):
        """Отложенное сохранение данных инструктажа по ТБ (debounced на 500 мс)."""
        if hasattr(self, '_safety_save_timer') and self._safety_save_timer:
            try:
                self.after_cancel(self._safety_save_timer)
            except Exception:
                pass
        self._safety_save_timer = self.after(500, self.auto_save_safety_briefing_data)

    def auto_save_safety_briefing_data(self, *args):
        """Автоматически сохраняет все данные вкладки 'Инструктаж по ТБ' в config_data и на диск."""
        try:
            p1_data = [
                {
                    "student": r[0].get(),
                    "date": r[1].get(),
                    "content": r[2].get(),
                    "signature": r[3].get()
                }
                for r in self.safety_briefing_p1_vars
            ]
            p2_data = [
                {
                    "student": r[0].get(),
                    "date": r[1].get(),
                    "content": r[2].get(),
                    "signature": r[3].get()
                }
                for r in self.safety_briefing_p2_vars
            ]
            self.config_data["safety_briefing_p1"] = p1_data
            self.config_data["safety_briefing_p2"] = p2_data
            save_config(self.config_data)
        except Exception:
            pass

    def on_generate_safety_briefing_page_click(self, page_num: int):
        """Формирует и сохраняет отдельную страницу инструктажа по ТБ (стр. 38 или 39) в Excel."""
        self.auto_save_safety_briefing_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        journal_p = 38 if page_num == 1 else 39
        default_name = f"Инструктаж_по_ТБ_стр{journal_p}_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title=f"Сохранить Страницу {page_num} (стр. {journal_p} журнала) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            var_list = self.safety_briefing_p1_vars if page_num == 1 else self.safety_briefing_p2_vars
            items = [
                {
                    "student": r[0].get(),
                    "date": r[1].get(),
                    "content": r[2].get(),
                    "signature": r[3].get()
                }
                for r in var_list
            ]
            if page_num == 1:
                generate_excel_safety_briefing_p1(save_path, items=items)
                self.last_saved_safety_p1_file = save_path
                if hasattr(self, 'btn_open_safety_p1'):
                    self.btn_open_safety_p1.config(state=tk.NORMAL)
            else:
                generate_excel_safety_briefing_p2(save_path, items=items)
                self.last_saved_safety_p2_file = save_path
                if hasattr(self, 'btn_open_safety_p2'):
                    self.btn_open_safety_p2.config(state=tk.NORMAL)

            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Файл 'Инструктаж по технике безопасности' (стр. {journal_p}) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set(f"Ошибка при сохранении страницы {page_num} ТБ.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_safety_briefing_page_file_click(self, page_num: int):
        """Открывает сгенерированный файл страницы инструктажа по ТБ."""
        file_path = self.last_saved_safety_p1_file if page_num == 1 else self.last_saved_safety_p2_file
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("Файл не найден", f"Созданный файл страницы {page_num} не найден на диске.")
            return
        self._open_file_system(file_path)

    def on_generate_safety_briefing_spread_click(self):
        """Формирует и сохраняет разворот инструктажа по ТБ (стр. 38 и 39) в один файл Excel."""
        self.auto_save_safety_briefing_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Инструктаж_по_ТБ_Разворот_стр38-39_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Разворот 'Инструктаж по технике безопасности' (стр. 38 и 39) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            items_p1 = [
                {
                    "student": r[0].get(),
                    "date": r[1].get(),
                    "content": r[2].get(),
                    "signature": r[3].get()
                }
                for r in self.safety_briefing_p1_vars
            ]
            items_p2 = [
                {
                    "student": r[0].get(),
                    "date": r[1].get(),
                    "content": r[2].get(),
                    "signature": r[3].get()
                }
                for r in self.safety_briefing_p2_vars
            ]
            generate_excel_safety_briefing_spread(save_path, items_p1=items_p1, items_p2=items_p2)
            self.last_saved_safety_spread_file = save_path
            if hasattr(self, 'btn_open_safety_spread'):
                self.btn_open_safety_spread.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен разворот ТБ: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Разворот 'Инструктаж по технике безопасности' (стр. 38 и 39) успешно сформирован в одном файле:\n{save_path}\n\n"
                f"В книге созданы 2 листа:\n"
                f"• Инструктаж по ТБ (стр. 38) — левая страница (отступ под корешок 30 мм СПРАВА)\n"
                f"• Инструктаж по ТБ (стр. 39) — правая страница (отступ под корешок 30 мм СЛЕВА)"
            )
        except Exception as e:
            self.status_var.set("Ошибка при сохранении разворота ТБ.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_safety_briefing_spread_file_click(self):
        """Открывает сгенерированный файл разворота инструктажа по ТБ."""
        if not hasattr(self, 'last_saved_safety_spread_file') or not self.last_saved_safety_spread_file or not os.path.exists(self.last_saved_safety_spread_file):
            messagebox.showwarning("Файл не найден", "Созданный файл разворота инструктажа по ТБ не найден на диске.")
            return
        self._open_file_system(self.last_saved_safety_spread_file)

    # =========================================================================
    # ВКЛАДКА 20: Годовой цифровой отчёт (стр. 40 журнала)
    # =========================================================================
    def setup_annual_report_tab(self, parent_tab, card_bg, next_tab_idx: int = 21):
        """Сборка интерфейса вкладки 'Годовой цифровой отчёт' (стр. 40 журнала)."""
        # Скроллируемый контейнер Canvas для поддержки любых разрешений экрана
        canvas = tk.Canvas(parent_tab, bg=card_bg, highlightthickness=0)
        v_scroll = ttk.Scrollbar(parent_tab, orient="vertical", command=canvas.yview)
        scroll_content = tk.Frame(canvas, bg=card_bg, padx=14, pady=12)

        window_id = canvas.create_window((0, 0), window=scroll_content, anchor="nw")

        def _on_scroll_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            new_w = max(event.width, 920)
            canvas.itemconfig(window_id, width=new_w)

        scroll_content.bind("<Configure>", _on_scroll_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=v_scroll.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel(event):
            if canvas.winfo_exists():
                if sys.platform == "darwin":
                    canvas.yview_scroll(int(-1 * event.delta), "units")
                else:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # 1. Верхняя панель заголовка и быстрых действий
        header_card = tk.LabelFrame(
            scroll_content,
            text=" Годовой цифровой отчёт (стр. 40 журнала) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=14,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        header_card.pack(fill=tk.X, pady=(0, 10))

        info_lbl = tk.Label(
            header_card,
            text="Таблица из 6 основных столбцов (18 полей в строке) и 4 строк: Учебный период, Всего, Мальчиков, Девочек, "
                 "Количество по классам (I–XI), Сколько лет посещает объединение (1, 2, 3 и более). "
                 "Ниже таблицы приведены обязательные требования по охране труда и ТБ.",
            font=(self.font_sans, 9),
            bg=card_bg,
            fg="#475569",
            justify=tk.LEFT,
            wraplength=850
        )
        info_lbl.pack(anchor="w", pady=(0, 8))

        btn_bar = tk.Frame(header_card, bg=card_bg)
        btn_bar.pack(fill=tk.X)

        btn_calc = tk.Button(
            btn_bar,
            text="🔢 Рассчитать по списку обучающихся",
            command=self.calc_annual_report_from_students,
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        )
        btn_calc.pack(side=tk.LEFT, padx=(0, 6))

        btn_paste = tk.Button(
            btn_bar,
            text="📋 Вставить строку (Ctrl+V)",
            command=lambda: self.paste_annual_report_from_clipboard(0, 0),
            font=(self.font_sans, 9),
            bg="#f1f5f9",
            fg="#1e293b",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        )
        btn_paste.pack(side=tk.LEFT, padx=(0, 6))

        btn_clear = tk.Button(
            btn_bar,
            text="🗑 Очистить таблицу",
            command=self.clear_annual_report_table,
            font=(self.font_sans, 9),
            bg="#fee2e2",
            fg="#991b1b",
            activebackground="#fecaca",
            activeforeground="#7f1d1d",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        )
        btn_clear.pack(side=tk.LEFT, padx=(0, 6))

        btn_next = tk.Button(
            btn_bar,
            text="Далее: Формирование Excel ➔",
            command=lambda: self.notebook.select(next_tab_idx),
            font=(self.font_sans, 9, "bold"),
            bg="#3b82f6",
            fg="#ffffff",
            activebackground="#2563eb",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=4,
            cursor="hand2"
        )
        btn_next.pack(side=tk.RIGHT)

        # 2. Карточка таблицы (4 строки, 18 колонок)
        table_card = tk.LabelFrame(
            scroll_content,
            text=" Таблица отчёта (4 строки) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=10,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        table_card.pack(fill=tk.X, pady=(0, 12))

        # Горизонтальный контейнер для таблицы
        h_table_frame = tk.Frame(table_card, bg="#e2e8f0", relief=tk.SOLID, bd=1)
        h_table_frame.pack(fill=tk.X, pady=(0, 4))

        # Двухуровневая шапка таблицы:
        # Строка 0 шапки:
        # Col 0: Учебный период (rowspan 2)
        # Col 1: Всего в объединении (rowspan 2)
        # Col 2: Мальчиков (rowspan 2)
        # Col 3: Девочек (rowspan 2)
        # Col 4..14: Количество обучающихся по классам (colspan 11)
        # Col 15..17: Сколько лет посещает объединение (colspan 3)

        th_bg = "#f1f5f9"
        th_fg = "#1e293b"
        th_font = (self.font_sans, 8, "bold")

        lbl_p = tk.Label(h_table_frame, text="Учебный\nпериод", font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, width=12, pady=4)
        lbl_p.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=1, pady=1)

        lbl_tot = tk.Label(h_table_frame, text="Всего в\nобъед.", font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, width=7, pady=4)
        lbl_tot.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=1, pady=1)

        lbl_boy = tk.Label(h_table_frame, text="Маль-\nчиков", font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, width=6, pady=4)
        lbl_boy.grid(row=0, column=2, rowspan=2, sticky="nsew", padx=1, pady=1)

        lbl_grl = tk.Label(h_table_frame, text="Дево-\nчек", font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, width=6, pady=4)
        lbl_grl.grid(row=0, column=3, rowspan=2, sticky="nsew", padx=1, pady=1)

        lbl_cls = tk.Label(h_table_frame, text="Количество обучающихся по классам", font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, pady=3)
        lbl_cls.grid(row=0, column=4, columnspan=11, sticky="nsew", padx=1, pady=1)

        lbl_yrs = tk.Label(h_table_frame, text="Сколько лет посещает", font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, pady=3)
        lbl_yrs.grid(row=0, column=15, columnspan=3, sticky="nsew", padx=1, pady=1)

        # Строка 1 шапки (подшапка):
        # Классы: I..XI
        roman_headers = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI"]
        for c_idx, r_txt in enumerate(roman_headers):
            h_c = tk.Label(h_table_frame, text=r_txt, font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, width=3, pady=2)
            h_c.grid(row=1, column=4 + c_idx, sticky="nsew", padx=1, pady=1)

        # Года: 1, 2, 3 и более
        years_headers = [("1", 3), ("2", 3), ("3+", 5)]
        for y_idx, (y_txt, y_w) in enumerate(years_headers):
            h_y = tk.Label(h_table_frame, text=y_txt, font=th_font, bg=th_bg, fg=th_fg, relief=tk.RIDGE, bd=1, width=y_w, pady=2)
            h_y.grid(row=1, column=15 + y_idx, sticky="nsew", padx=1, pady=1)

        # 4 строки данных
        self.annual_report_widgets.clear()
        for r_idx in range(4):
            row_widgets = []
            r_data = self.annual_report_vars[r_idx]

            # 1. Период (Col 0)
            e_per = ttk.Entry(h_table_frame, textvariable=r_data["period"], font=(self.font_sans, 9, "bold"), width=14, justify="center")
            e_per.grid(row=2 + r_idx, column=0, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_per)
            e_per.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
            row_widgets.append(e_per)

            # 2. Всего (Col 1)
            e_tot = ttk.Entry(h_table_frame, textvariable=r_data["total"], font=(self.font_sans, 9, "bold"), width=7, justify="center")
            e_tot.grid(row=2 + r_idx, column=1, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_tot)
            e_tot.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
            row_widgets.append(e_tot)

            # 3. Мальчиков (Col 2)
            e_boy = ttk.Entry(h_table_frame, textvariable=r_data["boys"], font=(self.font_sans, 9), width=6, justify="center")
            e_boy.grid(row=2 + r_idx, column=2, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_boy)
            e_boy.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
            row_widgets.append(e_boy)

            # 4. Девочек (Col 3)
            e_grl = ttk.Entry(h_table_frame, textvariable=r_data["girls"], font=(self.font_sans, 9), width=6, justify="center")
            e_grl.grid(row=2 + r_idx, column=3, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_grl)
            e_grl.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
            row_widgets.append(e_grl)

            # 5..15. Классы (Col 4..14)
            for c_i in range(11):
                e_cl = ttk.Entry(h_table_frame, textvariable=r_data["classes"][c_i], font=(self.font_sans, 9), width=3, justify="center")
                e_cl.grid(row=2 + r_idx, column=4 + c_i, sticky="nsew", padx=1, pady=1)
                self.attach_context_menu(e_cl)
                e_cl.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
                row_widgets.append(e_cl)

            # 16..18. Сколько лет (Col 15..17)
            for y_i in range(3):
                e_yr = ttk.Entry(h_table_frame, textvariable=r_data["years"][y_i], font=(self.font_sans, 9), width=3 if y_i < 2 else 5, justify="center")
                e_yr.grid(row=2 + r_idx, column=15 + y_i, sticky="nsew", padx=1, pady=1)
                self.attach_context_menu(e_yr)
                e_yr.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")
                row_widgets.append(e_yr)

            # Привязка горячих клавиш вставки
            for col_idx, w in enumerate(row_widgets):
                w.bind("<Control-v>", lambda ev, r=r_idx, c=col_idx: self.paste_annual_report_from_clipboard(r, c, ev))
                w.bind("<Control-V>", lambda ev, r=r_idx, c=col_idx: self.paste_annual_report_from_clipboard(r, c, ev))

            self.annual_report_widgets.append(row_widgets)

        # Навигация стрелками Вверх/Вниз
        for r_idx, row_w in enumerate(self.annual_report_widgets):
            for col_idx, w in enumerate(row_w):
                if r_idx < 3:
                    down_target = self.annual_report_widgets[r_idx + 1][col_idx]
                    w.bind("<Down>", lambda e, t=down_target: t.focus_set())
                if r_idx > 0:
                    up_target = self.annual_report_widgets[r_idx - 1][col_idx]
                    w.bind("<Up>", lambda e, t=up_target: t.focus_set())

        # 3. Карточка текста: Требования по охране труда и технике безопасности
        req_card = tk.LabelFrame(
            scroll_content,
            text=" ТРЕБОВАНИЯ ПО ОХРАНЕ ТРУДА, ТЕХНИКЕ БЕЗОПАСНОСТИ И ПРОИЗВОДСТВЕННОЙ САНИТАРИИ ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=14,
            pady=12,
            relief=tk.SOLID,
            bd=1
        )
        req_card.pack(fill=tk.X, pady=(0, 10))

        req_title_1 = tk.Label(
            req_card,
            text="ТРЕБОВАНИЯ\nК РУКОВОДИТЕЛЯМ ОБЪЕДИНЕНИЙ И ПОДРАЗДЕЛЕНИЙ ОРГАНИЗАЦИЙ ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ ДЕТЕЙ "
                 "ПО ОХРАНЕ ТРУДА, ТЕХНИКЕ БЕЗОПАСНОСТИ И ПРОИЗВОДСТВЕННОЙ САНИТАРИИ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            justify=tk.CENTER
        )
        req_title_1.pack(fill=tk.X, pady=(0, 8))

        req_lead = tk.Label(
            req_card,
            text="РУКОВОДИТЕЛЬ ОБЪЕДИНЕНИЯ при непосредственном участии и помощи заведующего лабораторией, кабинетом, мастерской:",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#334155",
            anchor="w",
            justify=tk.LEFT
        )
        req_lead.pack(fill=tk.X, pady=(0, 6))

        points = [
            "1. Принимает необходимые меры для создания здоровых и безопасных условий проведения занятий.",
            "2. Обеспечивает выполнение действующих правил и инструкций по технике безопасности и производственной санитарии.",
            "3. Проводит занятия и работы при наличии соответствующего оборудования и других условий, предусмотренных правилами и нормами по технике безопасности.",
            "4. Обеспечивает безопасное состояние рабочих мест, оборудования, приборов, инструментов и санитарное состояние помещений.",
            "5. Проводит инструктаж обучающихся в объединении по технике безопасности с соответствующим оформлением инструктажа в журнале (см. «Список обучающихся в объединении, прошедших инструктаж по ТБ»).",
            "6. Разрабатывает мероприятия по технике безопасности для включения их в план и соглашение по охране труда.",
            "7. Не допускает обучающихся в объединении к проведению работы или занятий без предусмотренной спецодежды и защитных приспособлений.",
            "8. Приостанавливает проведение работы и занятий, сопряжённых с опасностью для жизни, и докладывает об этом руководителю организации.",
            "9. Немедленно извещает руководителя организации о каждом несчастном случае.",
            "10. Несет ответственность за несчастные случаи, происшедшие в результате невыполнения им обязанностей, возложенных настоящими требованиями, ФЗ № 273 от 29.12.2012 «Об образовании в РФ» (п. 4 ч. 4 ст. 41), Приказом Минобрнауки РФ № 602 от 27.06.2017 «Об утверждении Порядка расследования и учета несчастных случаев с обучающимися во время пребывания в организации, осуществляющей образовательную деятельность»."
        ]

        for p_text in points:
            lbl_pt = tk.Label(
                req_card,
                text=p_text,
                font=(self.font_sans, 8),
                bg=card_bg,
                fg="#334155",
                anchor="w",
                justify=tk.LEFT,
                wraplength=850
            )
            lbl_pt.pack(fill=tk.X, pady=2)

        try:
            scroll_content.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        except Exception:
            pass

    def calc_annual_report_from_students(self):
        """Автоматический подсчёт показателей отчёта из имеющихся данных списка обучающихся."""
        # Собираем список детей и их данные (класс, дата зачисления, отчество/пол)
        student_records = []

        # Сначала проверяем вкладку 'Список обучающихся' (страницы 1, 3, 5)
        if hasattr(self, 'students_list_pages_vars'):
            for odd_p, even_p in ((1, 2), (3, 4), (5, 6)):
                odd_rows = self.students_list_pages_vars.get(odd_p, [])
                even_rows = self.students_list_pages_vars.get(even_p, [])
                for i in range(len(odd_rows)):
                    st_name = odd_rows[i][0].get().strip()
                    if st_name:
                        st_class = odd_rows[i][2].get().strip()
                        join_date = even_rows[i][2].get().strip() if i < len(even_rows) else ""
                        student_records.append({
                            "name": st_name,
                            "class": st_class,
                            "join_date": join_date
                        })

        # Если список обучающихся пуст, берем из месяцев (Сентябрь)
        if not student_records and hasattr(self, 'months_data') and "september" in self.months_data:
            st_vars = self.months_data["september"]["student_vars"]
            for sv in st_vars:
                name = sv.get().strip()
                if name:
                    student_records.append({"name": name, "class": "", "join_date": ""})

        if not student_records:
            messagebox.showinfo("Нет данных", "В журнале пока нет заполненных списков обучающихся (вкладка 'Список обучающихся' или 'Сентябрь').")
            return

        total_cnt = len(student_records)
        boys_cnt = 0
        girls_cnt = 0

        # Определение пола по отчеству / окончанию фамилии или имени
        for rec in student_records:
            full_name = rec["name"].lower()
            parts = full_name.split()
            gender = None
            for p in parts:
                if p.endswith("вич") or p.endswith("оглы") or p.endswith("улы"):
                    gender = "boy"
                    break
                elif p.endswith("вна") or p.endswith("кызы") or p.endswith("гызы"):
                    gender = "girl"
                    break

            if not gender:
                # Если отчества нет, проверяем имя или фамилию
                if len(parts) >= 2:
                    first_name = parts[1]
                    if first_name.endswith("а") or first_name.endswith("я") or first_name in ("любовь",):
                        gender = "girl"
                    else:
                        gender = "boy"
                else:
                    gender = "boy"

            if gender == "boy":
                boys_cnt += 1
            else:
                girls_cnt += 1

        # Распределение по классам I..XI
        class_counts = [0] * 11
        for rec in student_records:
            cl_str = rec["class"]
            # Ищем цифру от 1 до 11
            m = re.search(r'\b(1[0-1]|[1-9])\b', cl_str)
            if m:
                num = int(m.group(1))
                if 1 <= num <= 11:
                    class_counts[num - 1] += 1

        # Распределение по стажу (1 год, 2 года, 3 и более)
        current_study_year = 2024
        try:
            current_study_year = int(self.start_year.strip() or "2024")
        except Exception:
            pass

        years_counts = [0, 0, 0]  # 1, 2, 3+
        for rec in student_records:
            jd = rec["join_date"]
            years_in = 1
            if jd:
                m_yr = re.search(r'\b(19\d\d|20\d\d)\b', jd)
                if m_yr:
                    join_yr = int(m_yr.group(1))
                    diff = current_study_year - join_yr + 1
                    years_in = max(1, diff)

            if years_in == 1:
                years_counts[0] += 1
            elif years_in == 2:
                years_counts[1] += 1
            else:
                years_counts[2] += 1

        # Заполняем строки отчёта:
        # Для строк 0 (I полугодие), 1 (II полугодие), 2 (За год)
        for r_idx in range(3):
            r = self.annual_report_vars[r_idx]
            r["total"].set(str(total_cnt))
            r["boys"].set(str(boys_cnt))
            r["girls"].set(str(girls_cnt))

            for c_i in range(11):
                cnt = class_counts[c_i]
                r["classes"][c_i].set(str(cnt) if cnt > 0 else "")

            for y_i in range(3):
                cnt_y = years_counts[y_i]
                r["years"][y_i].set(str(cnt_y) if cnt_y > 0 else "")

        self.schedule_auto_save_annual_report()
        self.status_var.set(f"Годовой цифровой отчёт рассчитан: всего {total_cnt} чел. (мальчиков: {boys_cnt}, девочек: {girls_cnt}).")
        messagebox.showinfo(
            "Расчёт выполнен!",
            f"Показатели успешно рассчитаны по данным {total_cnt} обучающихся:\n\n"
            f"• Всего в объединении: {total_cnt}\n"
            f"• Мальчиков: {boys_cnt}\n"
            f"• Девочек: {girls_cnt}\n"
            f"• Распределено по классам и стажу посещения.\n\n"
            f"Данные внесены в строки 'I полугодие', 'II полугодие' и 'За год'. При необходимости вы можете скорректировать любое значение."
        )

    def clear_annual_report_table(self):
        """Очищает данные полей таблицы Годового цифрового отчёта."""
        if not messagebox.askyesno("Очистить таблицу", "Вы действительно хотите очистить все числовые поля Годового цифрового отчёта?"):
            return
        for r in self.annual_report_vars:
            r["total"].set("")
            r["boys"].set("")
            r["girls"].set("")
            for cv in r["classes"]:
                cv.set("")
            for yv in r["years"]:
                yv.set("")
        self.auto_save_annual_report_data()
        self.status_var.set("Таблица Годового цифрового отчёта очищена.")

    def paste_annual_report_from_clipboard(self, start_row_idx: int = 0, col_idx: int = 0, event=None):
        """Вставляет значения в таблицу Годового отчёта из буфера обмена (разделитель табуляция или перевод строки)."""
        raw_text = self.get_clipboard_text()
        if not raw_text:
            self.status_var.set("Буфер обмена пуст или не содержит текст.")
            return "break"

        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
        if not lines:
            self.status_var.set("В буфере обмена нет строк для вставки.")
            return "break"

        for offset, line in enumerate(lines):
            r_i = start_row_idx + offset
            if r_i >= 4:
                break
            parts = [p.strip() for p in line.split('\t')]
            row_vars = self.annual_report_vars[r_i]["all_vars"]
            for c_offset, val in enumerate(parts):
                target_col = col_idx + c_offset
                if target_col < len(row_vars):
                    row_vars[target_col].set(val)

        self.schedule_auto_save_annual_report()
        self.status_var.set(f"Вставлено {len(lines)} строк в Годовой цифровой отчёт.")
        return "break"

    def schedule_auto_save_annual_report(self):
        """Отложенное сохранение данных Годового отчёта (debounced 500 мс)."""
        if hasattr(self, '_annual_report_save_timer') and self._annual_report_save_timer:
            try:
                self.after_cancel(self._annual_report_save_timer)
            except Exception:
                pass
        self._annual_report_save_timer = self.after(500, self.auto_save_annual_report_data)

    def auto_save_annual_report_data(self, *args):
        """Автоматически сохраняет все данные вкладки 'Годовой цифровой отчёт' в config_data и на диск."""
        if not hasattr(self, 'annual_report_vars'):
            return
        try:
            ar_list = []
            for r in self.annual_report_vars:
                ar_list.append({
                    "period": r["period"].get(),
                    "total": r["total"].get(),
                    "boys": r["boys"].get(),
                    "girls": r["girls"].get(),
                    "classes": [cv.get() for cv in r["classes"]],
                    "years_in_org": [yv.get() for yv in r["years"]]
                })
            self.config_data["annual_report"] = ar_list
            save_config(self.config_data)
        except Exception:
            pass

    def on_generate_annual_report_click(self):
        """Формирует и сохраняет файл 'Годовой цифровой отчёт' (стр. 40 журнала) в Excel."""
        self.auto_save_annual_report_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Годовой_цифровой_отчёт_стр40_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Годовой цифровой отчёт (стр. 40) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            data = []
            for r in self.annual_report_vars:
                data.append({
                    "period": r["period"].get(),
                    "total": r["total"].get(),
                    "boys": r["boys"].get(),
                    "girls": r["girls"].get(),
                    "classes": [cv.get() for cv in r["classes"]],
                    "years_in_org": [yv.get() for yv in r["years"]]
                })
            generate_excel_annual_report(save_path, data=data)
            self.last_saved_annual_report_file = save_path
            if hasattr(self, 'btn_open_annual_report'):
                self.btn_open_annual_report.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo("Успешно!", f"Файл 'Годовой цифровой отчёт' (стр. 40) успешно создан:\n{save_path}")
        except Exception as e:
            self.status_var.set("Ошибка при сохранении Годового отчёта.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_annual_report_file_click(self):
        """Открывает файл Годового цифрового отчёта в системной программе."""
        if not hasattr(self, 'last_saved_annual_report_file') or not self.last_saved_annual_report_file or not os.path.exists(self.last_saved_annual_report_file):
            messagebox.showwarning("Файл не найден", "Созданный файл Годового цифрового отчёта не найден на диске.")
            return
        self._open_file_system(self.last_saved_annual_report_file)

    # =========================================================================
    # ВКЛАДКА 21: Подсчёт отработанного времени (учёт часов за год)
    # =========================================================================
    def setup_work_hours_tab(self, parent_tab, card_bg, next_tab_idx: int = 22):
        """Сборка интерфейса вкладки 'Подсчёт отработанного времени'."""
        canvas = tk.Canvas(parent_tab, bg=card_bg, highlightthickness=0)
        v_scroll = ttk.Scrollbar(parent_tab, orient="vertical", command=canvas.yview)
        scroll_content = tk.Frame(canvas, bg=card_bg, padx=14, pady=12)

        window_id = canvas.create_window((0, 0), window=scroll_content, anchor="nw")

        def _on_scroll_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            new_w = max(event.width, 920)
            canvas.itemconfig(window_id, width=new_w)

        scroll_content.bind("<Configure>", _on_scroll_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=v_scroll.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel(event):
            if canvas.winfo_exists():
                if sys.platform == "darwin":
                    canvas.yview_scroll(int(-1 * event.delta), "units")
                else:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # 1. Верхняя панель заголовка и быстрых действий
        header_card = tk.LabelFrame(
            scroll_content,
            text=" Подсчёт отработанного времени (учёт часов за учебный год) ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=14,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        header_card.pack(fill=tk.X, pady=(0, 10))

        info_lbl = tk.Label(
            header_card,
            text="Автоматический подсчёт проведённых занятий и отработанных педагогических часов по всем месяцам учебного года "
                 "(Сентябрь – Август) на основе журнала. Позволяет сформировать отдельный документ Excel в виде таблицы "
                 "для печати и сдачи отчётности.",
            font=(self.font_sans, 9),
            bg=card_bg,
            fg="#475569",
            justify=tk.LEFT,
            wraplength=880
        )
        info_lbl.pack(anchor="w", pady=(0, 8))

        btn_bar = tk.Frame(header_card, bg=card_bg)
        btn_bar.pack(fill=tk.X)

        btn_calc = tk.Button(
            btn_bar,
            text="🔄 Обновить расчёт из журнала",
            command=lambda: self.calc_work_hours_from_months(notify=True),
            font=(self.font_sans, 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        )
        btn_calc.pack(side=tk.LEFT, padx=(0, 6))

        btn_gen = tk.Button(
            btn_bar,
            text="💾 Экспорт отчёта в Excel (.xlsx)",
            command=self.on_generate_work_hours_click,
            font=(self.font_sans, 9, "bold"),
            bg="#16a34a",
            fg="#ffffff",
            activebackground="#15803d",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2"
        )
        btn_gen.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_open_work_hours_tab = tk.Button(
            btn_bar,
            text="📂 Открыть файл Excel",
            command=self.on_open_work_hours_file_click,
            font=(self.font_sans, 9),
            bg="#f1f5f9",
            fg="#1e293b",
            state=tk.DISABLED,
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        )
        self.btn_open_work_hours_tab.pack(side=tk.LEFT, padx=(0, 6))

        btn_clear = tk.Button(
            btn_bar,
            text="🗑 Очистить таблицу",
            command=self.clear_work_hours_table,
            font=(self.font_sans, 9),
            bg="#fee2e2",
            fg="#991b1b",
            activebackground="#fecaca",
            activeforeground="#7f1d1d",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2"
        )
        btn_clear.pack(side=tk.LEFT, padx=(0, 6))

        btn_next = tk.Button(
            btn_bar,
            text="Далее: Формирование Excel ➔",
            command=lambda: self.notebook.select(next_tab_idx),
            font=(self.font_sans, 9, "bold"),
            bg="#3b82f6",
            fg="#ffffff",
            activebackground="#2563eb",
            activeforeground="#ffffff",
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=4,
            cursor="hand2"
        )
        btn_next.pack(side=tk.RIGHT)

        # 2. Карточка сводных показателей (4 KPI-блока)
        kpi_card = tk.Frame(scroll_content, bg=card_bg)
        kpi_card.pack(fill=tk.X, pady=(0, 12))

        kpi_data = [
            ("📅 Всего занятий за год", self.work_hours_total_lessons, "", "#0284c7", "#f0f9ff"),
            ("👨‍🏫 Отработано часов (педагог)", self.work_hours_total_teacher, " ч.", "#0f766e", "#f0fdfa"),
            ("🎹 Отработано часов (аккомп.)", self.work_hours_total_acc, " ч.", "#7c3aed", "#f5f3ff"),
            ("⏱ ИТОГОВАЯ СУММА ЧАСОВ", self.work_hours_total_all, " ч.", "#b91c1c", "#fef2f2"),
        ]

        for title, var, suffix, color, bg_col in kpi_data:
            box = tk.Frame(kpi_card, bg=bg_col, relief=tk.SOLID, bd=1, padx=12, pady=8)
            box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=3)

            lbl_t = tk.Label(box, text=title, font=(self.font_sans, 8, "bold"), bg=bg_col, fg="#475569")
            lbl_t.pack(anchor="w")

            val_frame = tk.Frame(box, bg=bg_col)
            val_frame.pack(anchor="w", pady=(4, 0))

            lbl_v = tk.Label(val_frame, textvariable=var, font=(self.font_sans, 14, "bold"), bg=bg_col, fg=color)
            lbl_v.pack(side=tk.LEFT)

            if suffix:
                lbl_s = tk.Label(val_frame, text=suffix, font=(self.font_sans, 10, "bold"), bg=bg_col, fg=color)
                lbl_s.pack(side=tk.LEFT, padx=(2, 0), pady=(3, 0))

        # 3. Карточка таблицы (12 месяцев + Итого)
        table_card = tk.LabelFrame(
            scroll_content,
            text=" Ведомость учёта фактически отработанного времени по месяцам ",
            font=(self.font_sans, 10, "bold"),
            bg=card_bg,
            fg="#0f172a",
            padx=10,
            pady=10,
            relief=tk.SOLID,
            bd=1
        )
        table_card.pack(fill=tk.X, pady=(0, 12))

        tbl_grid = tk.Frame(table_card, bg="#cbd5e1", relief=tk.SOLID, bd=1)
        tbl_grid.pack(fill=tk.X, pady=(0, 4))

        th_bg = "#f1f5f9"
        th_fg = "#1e293b"
        th_font = (self.font_sans, 8, "bold")

        # Шапка таблицы
        headers_cfg = [
            ("№", 4, "center"),
            ("Учебный месяц", 15, "w"),
            ("Кол-во занятий", 13, "center"),
            ("Часы (педагог)", 13, "center"),
            ("Часы (аккомп.)", 14, "center"),
            ("Всего часов", 13, "center"),
            ("Примечание / подпись", 24, "w"),
        ]

        for col_i, (h_title, h_width, h_anchor) in enumerate(headers_cfg):
            lbl_th = tk.Label(
                tbl_grid,
                text=h_title,
                font=th_font,
                bg=th_bg,
                fg=th_fg,
                relief=tk.RIDGE,
                bd=1,
                width=h_width,
                pady=5
            )
            lbl_th.grid(row=0, column=col_i, sticky="nsew", padx=1, pady=1)

        # 12 строк месяцев
        self.work_hours_widgets.clear()
        for i in range(12):
            m_item = self.work_hours_vars[i]
            r_idx = 1 + i
            row_bg = "#ffffff" if i % 2 == 0 else "#f8fafc"

            # Col 0: Номер п/п
            lbl_num = tk.Label(
                tbl_grid,
                text=str(i + 1),
                font=(self.font_sans, 9),
                bg=row_bg,
                fg="#334155",
                relief=tk.RIDGE,
                bd=1,
                width=4,
                pady=3
            )
            lbl_num.grid(row=r_idx, column=0, sticky="nsew", padx=1, pady=1)

            # Col 1: Название месяца
            lbl_m = tk.Label(
                tbl_grid,
                text=m_item["month_name"],
                font=(self.font_sans, 9, "bold"),
                bg=row_bg,
                fg="#0f172a",
                relief=tk.RIDGE,
                bd=1,
                width=15,
                anchor="w",
                padx=6,
                pady=3
            )
            lbl_m.grid(row=r_idx, column=1, sticky="nsew", padx=1, pady=1)

            # Col 2: Количество занятий
            e_les = ttk.Entry(tbl_grid, textvariable=m_item["lessons_count"], font=(self.font_sans, 9), width=13, justify="center")
            e_les.grid(row=r_idx, column=2, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_les)
            e_les.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

            # Col 3: Часы педагога
            e_ht = ttk.Entry(tbl_grid, textvariable=m_item["hours_teacher"], font=(self.font_sans, 9), width=13, justify="center")
            e_ht.grid(row=r_idx, column=3, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_ht)
            e_ht.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

            # Col 4: Часы аккомпаниатора
            e_ha = ttk.Entry(tbl_grid, textvariable=m_item["hours_acc"], font=(self.font_sans, 9), width=14, justify="center")
            e_ha.grid(row=r_idx, column=4, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_ha)
            e_ha.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

            # Col 5: Всего за месяц
            e_tot = ttk.Entry(tbl_grid, textvariable=m_item["hours_total"], font=(self.font_sans, 9, "bold"), width=13, justify="center")
            e_tot.grid(row=r_idx, column=5, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_tot)
            e_tot.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

            # Col 6: Примечание
            e_not = ttk.Entry(tbl_grid, textvariable=m_item["notes"], font=(self.font_sans, 9), width=24, justify="left")
            e_not.grid(row=r_idx, column=6, sticky="nsew", padx=1, pady=1)
            self.attach_context_menu(e_not)
            e_not.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

            self.work_hours_widgets.append((e_les, e_ht, e_ha, e_tot, e_not))

        # Строка 13: ИТОГО ЗА ГОД
        tot_row_bg = "#e2e8f0"
        tot_font = (self.font_sans, 9, "bold")

        lbl_tot_title = tk.Label(
            tbl_grid,
            text="ИТОГО ЗА УЧЕБНЫЙ ГОД:",
            font=tot_font,
            bg=tot_row_bg,
            fg="#0f172a",
            relief=tk.RIDGE,
            bd=1,
            anchor="e",
            padx=8,
            pady=5
        )
        lbl_tot_title.grid(row=13, column=0, columnspan=2, sticky="nsew", padx=1, pady=1)

        lbl_tot_les = tk.Label(
            tbl_grid,
            textvariable=self.work_hours_total_lessons,
            font=tot_font,
            bg=tot_row_bg,
            fg="#0f172a",
            relief=tk.RIDGE,
            bd=1,
            justify="center",
            pady=5
        )
        lbl_tot_les.grid(row=13, column=2, sticky="nsew", padx=1, pady=1)

        lbl_tot_ht = tk.Label(
            tbl_grid,
            textvariable=self.work_hours_total_teacher,
            font=tot_font,
            bg=tot_row_bg,
            fg="#0f766e",
            relief=tk.RIDGE,
            bd=1,
            justify="center",
            pady=5
        )
        lbl_tot_ht.grid(row=13, column=3, sticky="nsew", padx=1, pady=1)

        lbl_tot_ha = tk.Label(
            tbl_grid,
            textvariable=self.work_hours_total_acc,
            font=tot_font,
            bg=tot_row_bg,
            fg="#7c3aed",
            relief=tk.RIDGE,
            bd=1,
            justify="center",
            pady=5
        )
        lbl_tot_ha.grid(row=13, column=4, sticky="nsew", padx=1, pady=1)

        lbl_tot_all = tk.Label(
            tbl_grid,
            textvariable=self.work_hours_total_all,
            font=(self.font_sans, 10, "bold"),
            bg=tot_row_bg,
            fg="#b91c1c",
            relief=tk.RIDGE,
            bd=1,
            justify="center",
            pady=5
        )
        lbl_tot_all.grid(row=13, column=5, sticky="nsew", padx=1, pady=1)

        lbl_tot_empty = tk.Label(
            tbl_grid,
            text="",
            bg=tot_row_bg,
            relief=tk.RIDGE,
            bd=1
        )
        lbl_tot_empty.grid(row=13, column=6, sticky="nsew", padx=1, pady=1)

        # 4. Карточка реквизитов (для контроля перед печатью)
        meta_card = tk.LabelFrame(
            scroll_content,
            text=" Сведения для формирования печатного листа ",
            font=(self.font_sans, 9, "bold"),
            bg=card_bg,
            fg="#1e293b",
            padx=12,
            pady=8,
            relief=tk.SOLID,
            bd=1
        )
        meta_card.pack(fill=tk.X, pady=(0, 6))

        info_meta = tk.Label(
            meta_card,
            text="В печатный документ автоматически подставляются: наименование организации, учебный год, "
                 "название объединения, группа, год обучения, ФИО педагога и аккомпаниатора из титульного листа и вкладки "
                 "«Основные данные». Внизу страницы формируются поля для подписей педагога и руководителя организации.",
            font=(self.font_sans, 8),
            bg=card_bg,
            fg="#64748b",
            justify=tk.LEFT,
            wraplength=880
        )
        info_meta.pack(anchor="w")

        # Первоначальный расчёт, если часы ещё не были загружены
        if all(v["hours_total"].get() == "" for v in self.work_hours_vars):
            self.calc_work_hours_from_months(notify=False)
        else:
            self.recalc_work_hours_totals()

        try:
            scroll_content.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        except Exception:
            pass

    def calc_work_hours_from_months(self, notify: bool = True):
        """Автоматический расчёт проведённых занятий и отработанных часов из всех 12 месяцев журнала."""
        updated_any = False
        for i, m_cfg in enumerate(MONTHS_CONFIG):
            m_key = m_cfg["key"]
            topic_rows = []
            if hasattr(self, 'months_data') and m_key in self.months_data:
                topic_rows = self.months_data[m_key].get("topic_row_vars", [])

            les_count = 0
            ht_sum = 0.0
            ha_sum = 0.0

            if topic_rows:
                for r in topic_rows:
                    d_val = r[0].get().strip() if len(r) > 0 else ""
                    c_val = r[1].get().strip() if len(r) > 1 else ""
                    ht_raw = r[2].get().strip() if len(r) > 2 else ""
                    ha_raw = r[4].get().strip() if len(r) > 4 else ""

                    if d_val or c_val or ht_raw or ha_raw:
                        les_count += 1

                    if ht_raw:
                        try:
                            ht_sum += float(ht_raw.replace(',', '.'))
                        except ValueError:
                            pass

                    if ha_raw:
                        try:
                            ha_sum += float(ha_raw.replace(',', '.'))
                        except ValueError:
                            pass
            else:
                saved_topics = self.config_data.get(f"{m_key}_topics", [])
                for t in saved_topics:
                    if not isinstance(t, dict):
                        continue
                    d_val = str(t.get("date", "")).strip()
                    c_val = str(t.get("content", "")).strip()
                    ht_raw = str(t.get("hours_teacher", "")).strip()
                    ha_raw = str(t.get("hours_acc", "")).strip()

                    if d_val or c_val or ht_raw or ha_raw:
                        les_count += 1

                    if ht_raw:
                        try:
                            ht_sum += float(ht_raw.replace(',', '.'))
                        except ValueError:
                            pass

                    if ha_raw:
                        try:
                            ha_sum += float(ha_raw.replace(',', '.'))
                        except ValueError:
                            pass

            tot_sum = ht_sum + ha_sum
            v_item = self.work_hours_vars[i]

            les_str = str(les_count) if les_count > 0 else ""
            ht_str = f"{int(ht_sum)}" if ht_sum == int(ht_sum) else f"{ht_sum:.1f}"
            if ht_sum == 0.0:
                ht_str = ""
            ha_str = f"{int(ha_sum)}" if ha_sum == int(ha_sum) else f"{ha_sum:.1f}"
            if ha_sum == 0.0:
                ha_str = ""
            tot_str = f"{int(tot_sum)}" if tot_sum == int(tot_sum) else f"{tot_sum:.1f}"
            if tot_sum == 0.0:
                tot_str = ""

            v_item["lessons_count"].set(les_str)
            v_item["hours_teacher"].set(ht_str)
            v_item["hours_acc"].set(ha_str)
            v_item["hours_total"].set(tot_str)
            if les_count > 0 or tot_sum > 0:
                updated_any = True

        self.recalc_work_hours_totals()
        self.auto_save_work_hours_data()

        if notify:
            tot_hours = self.work_hours_total_all.get()
            tot_les = self.work_hours_total_lessons.get()
            messagebox.showinfo(
                "Подсчёт завершён",
                f"Данные об отработанном времени успешно обновлены из журнала!\n\n"
                f"• Всего проведённых занятий: {tot_les}\n"
                f"• Итого часов за учебный год: {tot_hours} ч.\n"
                f"(Педагог: {self.work_hours_total_teacher.get()} ч., "
                f"Аккомпаниатор: {self.work_hours_total_acc.get()} ч.)"
            )

    def recalc_work_hours_totals(self, *args):
        """Пересчитывает суммарные часы и количество занятий по всем месяцам."""
        tot_les = 0
        tot_ht = 0.0
        tot_ha = 0.0
        tot_all = 0.0

        for v_item in self.work_hours_vars:
            les_val = v_item["lessons_count"].get().strip()
            ht_val = v_item["hours_teacher"].get().strip()
            ha_val = v_item["hours_acc"].get().strip()
            tot_m_val = v_item["hours_total"].get().strip()

            if les_val:
                try:
                    tot_les += int(les_val)
                except ValueError:
                    pass

            ht_f = 0.0
            if ht_val:
                try:
                    ht_f = float(ht_val.replace(',', '.'))
                    tot_ht += ht_f
                except ValueError:
                    pass

            ha_f = 0.0
            if ha_val:
                try:
                    ha_f = float(ha_val.replace(',', '.'))
                    tot_ha += ha_f
                except ValueError:
                    pass

            if ht_f > 0 or ha_f > 0:
                calc_m = ht_f + ha_f
                calc_str = f"{int(calc_m)}" if calc_m == int(calc_m) else f"{calc_m:.1f}"
                if not tot_m_val or tot_m_val != calc_str:
                    v_item["hours_total"].set(calc_str)
                tot_all += calc_m
            elif tot_m_val:
                try:
                    tot_all += float(tot_m_val.replace(',', '.'))
                except ValueError:
                    pass

        self.work_hours_total_lessons.set(str(tot_les))
        self.work_hours_total_teacher.set(f"{int(tot_ht)}" if tot_ht == int(tot_ht) else f"{tot_ht:.1f}")
        self.work_hours_total_acc.set(f"{int(tot_ha)}" if tot_ha == int(tot_ha) else f"{tot_ha:.1f}")
        self.work_hours_total_all.set(f"{int(tot_all)}" if tot_all == int(tot_all) else f"{tot_all:.1f}")

    def schedule_auto_save_work_hours(self):
        """Отложенное сохранение отчёта об отработанном времени."""
        self.recalc_work_hours_totals()
        if hasattr(self, '_work_hours_save_timer') and self._work_hours_save_timer:
            try:
                self.after_cancel(self._work_hours_save_timer)
            except Exception:
                pass
        self._work_hours_save_timer = self.after(500, self.auto_save_work_hours_data)

    def auto_save_work_hours_data(self, *args):
        """Автоматически сохраняет все данные вкладки 'Отработанное время' в config_data и на диск."""
        if not hasattr(self, 'work_hours_vars'):
            return
        try:
            wh_list = []
            for item in self.work_hours_vars:
                wh_list.append({
                    "month_key": item["month_key"],
                    "month_name": item["month_name"],
                    "lessons_count": item["lessons_count"].get(),
                    "hours_teacher": item["hours_teacher"].get(),
                    "hours_acc": item["hours_acc"].get(),
                    "hours_total": item["hours_total"].get(),
                    "notes": item["notes"].get()
                })
            self.config_data["work_hours_report"] = wh_list
            save_config(self.config_data)
        except Exception:
            pass

    def clear_work_hours_table(self):
        """Очищает поля таблицы отработанного времени."""
        if not messagebox.askyesno("Подтверждение", "Очистить все введённые данные в таблице отработанного времени?"):
            return
        for item in self.work_hours_vars:
            item["lessons_count"].set("")
            item["hours_teacher"].set("")
            item["hours_acc"].set("")
            item["hours_total"].set("")
            item["notes"].set("")
        self.recalc_work_hours_totals()
        self.auto_save_work_hours_data()

    def on_generate_work_hours_click(self):
        """Формирует и сохраняет отдельный Excel файл 'Отчёт об отработанном времени'."""
        self.auto_save_work_hours_data()
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"Отчёт_об_отработанном_времени_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Отчёт об отработанном времени в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            items = []
            for item in self.work_hours_vars:
                items.append({
                    "month_key": item["month_key"],
                    "month_name": item["month_name"],
                    "lessons_count": item["lessons_count"].get(),
                    "hours_teacher": item["hours_teacher"].get(),
                    "hours_acc": item["hours_acc"].get(),
                    "hours_total": item["hours_total"].get(),
                    "notes": item["notes"].get()
                })

            org_name = (
                (hasattr(self, 'main_org_var') and self.main_org_var.get().strip())
                or (hasattr(self, 'org_var') and self.org_var.get().strip())
                or self.config_data.get("main_org_name", "")
                or self.config_data.get("org_name", "")
            )
            academic_year = (
                (hasattr(self, 'year_var') and self.year_var.get().strip())
                or self.config_data.get("title_academic_year", "")
                or self.config_data.get("last_academic_year", "")
            )
            assoc_name = (
                (hasattr(self, 'association_var') and self.association_var.get().strip())
                or self.config_data.get("association", "")
            )
            group_name = (
                (hasattr(self, 'group_var') and self.group_var.get().strip())
                or (hasattr(self, 'group_name_var') and self.group_name_var.get().strip())
                or self.config_data.get("group_name", "")
            )
            study_year = (
                (hasattr(self, 'study_year_var') and self.study_year_var.get().strip())
                or self.config_data.get("study_year", "")
            )
            teacher_name = (
                (hasattr(self, 'rukovoditel_var') and self.rukovoditel_var.get().strip())
                or (hasattr(self, 'teacher_var') and self.teacher_var.get().strip())
                or self.config_data.get("leader", "")
                or self.config_data.get("teacher_name", "")
            )
            accompanist_name = (
                (hasattr(self, 'accompanist_var') and self.accompanist_var.get().strip())
                or self.config_data.get("accompanist", "")
            )

            generate_excel_work_hours(
                save_path=save_path,
                data=items,
                org_name=org_name,
                academic_year=academic_year,
                association=assoc_name,
                group_name=group_name,
                study_year=study_year,
                teacher_name=teacher_name,
                accompanist_name=accompanist_name
            )

            self.last_saved_work_hours_file = save_path
            if hasattr(self, 'btn_open_work_hours') and self.btn_open_work_hours:
                self.btn_open_work_hours.config(state=tk.NORMAL)
            if hasattr(self, 'btn_open_work_hours_tab') and self.btn_open_work_hours_tab:
                self.btn_open_work_hours_tab.config(state=tk.NORMAL)

            self.status_var.set(f"Успешно сохранён отчёт: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Отчёт об отработанном времени успешно сформирован в Excel:\n{save_path}\n\n"
                f"Документ оформлен в виде готовой таблицы для печати на А4 с суммами часов за учебный год и реквизитами для подписи."
            )
        except Exception as e:
            self.status_var.set("Ошибка при сохранении отчёта об отработанном времени.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_work_hours_file_click(self):
        """Открывает сгенерированный файл отчёта об отработанном времени в Excel."""
        if not hasattr(self, 'last_saved_work_hours_file') or not self.last_saved_work_hours_file or not os.path.exists(self.last_saved_work_hours_file):
            messagebox.showwarning("Файл не найден", "Созданный файл отчёта об отработанном времени не найден на диске.")
            return
        self._open_file_system(self.last_saved_work_hours_file)

    def attach_student_context_menu(self, widget, idx: int, month_key: str = "september"):
        """Контекстное меню для поля ввода обучающегося с пунктом быстрой вставки списка."""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="📋 Вставить список детей (Ctrl+V)", command=lambda: self.paste_students_from_clipboard(idx, month_key=month_key))
        menu.add_separator()
        menu.add_command(label="Вырезать", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Копировать", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Вставить текст", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_command(label="Выделить всё", command=lambda: widget.select_range(0, tk.END))

        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)

    def attach_topic_context_menu(self, widget, idx: int, month_key: str = "september"):
        """Контекстное меню для поля ввода темы ДОП с пунктом вставки списка тем."""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="📋 Вставить список тем (Ctrl+V)", command=lambda: self.paste_topics_from_clipboard(idx, month_key=month_key))
        menu.add_separator()
        menu.add_command(label="Вырезать", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Копировать", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Вставить текст", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_command(label="Выделить всё", command=lambda: widget.select_range(0, tk.END))

        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)

    def on_save_year_click(self):
        """Применяет изменения введенного учебного года и сохраняет его в историю config.json."""
        val = self.year_var.get().strip()
        if not val:
            self.status_var.set("Пожалуйста, введите учебный год для сохранения.")
            return
        # Разбор года для обложки журнала
        self.parse_years(val)
        # Сохранение в историю и файл config.json
        self.add_to_history(val)
        # Обновление элементов выпадающего списка
        if hasattr(self, 'combo_year') and isinstance(self.combo_year, ttk.Combobox):
            self.combo_year['values'] = self.config_data.get("history", [])
        if hasattr(self, 'title_year_var'):
            self.title_year_var.set(val)
        self.sync_years_from_academic_year(val)
        # Обновление полноценного предпросмотра листа
        self.draw_cover_preview()
        if hasattr(self, 'draw_title_page_preview'):
            self.draw_title_page_preview()
        self.update_summary_lbl()
        self.status_var.set(f"Значение «{val}» сохранено в истории config.json и применено.")

    def on_combo_select(self, event=None):
        val = self.combo_year.get().strip() if hasattr(self, 'combo_year') else ""
        if val:
            self.add_to_history(val)
            if hasattr(self, 'title_year_var'):
                self.title_year_var.set(val)
            self.sync_years_from_academic_year(val)
            self.draw_cover_preview()
            if hasattr(self, 'draw_title_page_preview'):
                self.draw_title_page_preview()
            self.update_summary_lbl()
            self.status_var.set(f"Выбрано из истории: {val}")

    def update_summary_lbl(self):
        st = self.start_year or "____"
        en = self.end_year or "____"
        teacher = self.teacher_var.get().strip() if hasattr(self, 'teacher_var') else ""
        teacher_display = teacher if teacher else "не указан (будет пустая линия для подписи)"
        org = self.org_var.get().strip() if hasattr(self, 'org_var') and self.org_var.get().strip() else "ГБУ ДО Республиканский детский образовательный технопарк"
        title_yr = self.title_year_var.get().strip() if hasattr(self, 'title_year_var') and self.title_year_var.get().strip() else f"{st} / {en}"
        sd = self.start_day_var.get().strip() if hasattr(self, 'start_day_var') else ""
        sm = self.start_month_var.get().strip() if hasattr(self, 'start_month_var') and self.start_month_var.get().strip() else "сентября"
        sy = self.start_year_var.get().strip() if hasattr(self, 'start_year_var') and self.start_year_var.get().strip() else "2024"
        ed = self.end_day_var.get().strip() if hasattr(self, 'end_day_var') else ""
        em = self.end_month_var.get().strip() if hasattr(self, 'end_month_var') and self.end_month_var.get().strip() else "мая"
        ey = self.end_year_var.get().strip() if hasattr(self, 'end_year_var') and self.end_year_var.get().strip() else "2025"

        sy_disp = f"20{sy} г." if len(sy) == 2 else (f"{sy} г." if sy and not sy.endswith("г.") else (sy or "20___ г."))
        ey_disp = f"20{ey} г." if len(ey) == 2 else (f"{ey} г." if ey and not ey.endswith("г.") else (ey or "20___ г."))
        sd_disp = f"« {sd} »" if sd else "« ___ »"
        ed_disp = f"« {ed} »" if ed else "« ___ »"

        # Основные данные (стр. 3)
        main_org = self.main_org_var.get().strip() if hasattr(self, 'main_org_var') else ""
        dept = self.department_var.get().strip() if hasattr(self, 'department_var') else ""
        assoc = self.association_var.get().strip() if hasattr(self, 'association_var') else ""
        grp = self.group_var.get().strip() if hasattr(self, 'group_var') else ""
        yr = self.study_year_var.get().strip() if hasattr(self, 'study_year_var') else ""
        ruk = self.rukovoditel_var.get().strip() if hasattr(self, 'rukovoditel_var') else ""

        if hasattr(self, 'summary_lbl') and self.summary_lbl:
            self.summary_lbl.config(
                text=f"1. Обложка журнала (Лист 1):\n"
                     f"   • Учебный год: «на {st} / {en} учебный год» (отступ 30 мм СЛЕВА)\n\n"
                     f"2. Титульный лист (Лист 2, стр. 1 журнала):\n"
                     f"   • Организация: {org}\n"
                     f"   • Учебный год: «на {title_yr} учебный год»\n"
                     f"   • Сроки: Начат {sd_disp} {sm} {sy_disp} — Окончен {ed_disp} {em} {ey_disp} (отступ 30 мм СЛЕВА)\n\n"
                     f"3. Оборот титульного листа (Лист 3, стр. 2 журнала):\n"
                     f"   • Педагог: {teacher_display}\n"
                     f"   • Правила ведения (10 пунктов) и ФЗ-273 (отступ 30 мм СПРАВА под корешок)\n\n"
                     f"4. Основные данные (Лист 4, стр. 3 журнала):\n"
                     f"   • Организация: {main_org or '—'} | Отдел: {dept or '—'} | Объединение: {assoc or '—'}\n"
                     f"   • Группа: {grp or '—'} | Год: {yr or '—'} | Руководитель: {ruk or '—'}"
            )

    def on_save_org_click(self):
        """Сохраняет введенное название организации в историю config.json."""
        org = self.org_var.get().strip()
        if not org:
            self.status_var.set("Пожалуйста, введите название организации.")
            return
        self.add_to_org_history(org)
        if hasattr(self, 'combo_org') and isinstance(self.combo_org, ttk.Combobox):
            self.combo_org['values'] = self.config_data.get("org_history", [])
            try:
                self.combo_org.selection_clear()
            except Exception:
                pass
        self.draw_title_page_preview()
        self.status_var.set(f"Организация «{org}» сохранена в config.json.")

    def on_org_combo_select(self, event=None):
        val = self.combo_org.get().strip() if hasattr(self, 'combo_org') else ""
        if val:
            self.org_var.set(val)
            self.add_to_org_history(val)
            self.draw_title_page_preview()
            self.status_var.set(f"Организация выбрана из истории: {val}")
        if hasattr(self, 'combo_org'):
            try:
                self.combo_org.selection_clear()
            except Exception:
                pass
        self.clear_focus()

    def on_title_input_changed(self, *args):
        """Мгновенно обновляет предпросмотр листа 'Титульный лист' при вводе любого поля."""
        if hasattr(self, 'org_var'):
            self.config_data["org_name"] = self.org_var.get().strip()
        if hasattr(self, 'title_year_var'):
            self.config_data["title_academic_year"] = self.title_year_var.get().strip()
        if hasattr(self, 'start_day_var'):
            self.config_data["start_day_val"] = self.start_day_var.get().strip()
        if hasattr(self, 'start_month_var'):
            self.config_data["start_month"] = self.start_month_var.get().strip()
        if hasattr(self, 'start_year_var'):
            self.config_data["start_year_val"] = self.start_year_var.get().strip()
        if hasattr(self, 'end_day_var'):
            self.config_data["end_day_val"] = self.end_day_var.get().strip()
        if hasattr(self, 'end_month_var'):
            self.config_data["end_month"] = self.end_month_var.get().strip()
        if hasattr(self, 'end_year_var'):
            self.config_data["end_year_val"] = self.end_year_var.get().strip()

        save_config(self.config_data)
        self.draw_title_page_preview()
        self.update_summary_lbl()

    def on_year_input_changed(self, *args):
        val = self.year_var.get()
        self.parse_years(val)
        st = self.start_year or "____"
        en = self.end_year or "____"
        if hasattr(self, 'preview_lbl') and self.preview_lbl:
            self.preview_lbl.config(text=f"«на {st} / {en} учебный год»")
        # Мгновенная перерисовка листа А4 при каждом изменении года
        self.draw_cover_preview()
        self.update_summary_lbl()
        cur_yr = val.strip()
        if cur_yr:
            self.config_data["last_academic_year"] = cur_yr
            save_config(self.config_data)

    def on_teacher_input_changed(self, *args):
        """Мгновенно обновляет предпросмотр листа 'Оборот титульного листа' при вводе ФИО."""
        name = self.teacher_var.get().strip()
        self.config_data["teacher_name"] = name
        save_config(self.config_data)
        self.draw_inside_cover_preview()
        self.update_summary_lbl()

    def on_enter_pressed(self):
        val = self.year_var.get().strip()
        if val:
            self.add_to_history(val)
            self.notebook.select(1)

    def on_go_to_export_click(self):
        val = self.year_var.get().strip()
        if val:
            self.add_to_history(val)
        self.notebook.select(1)

    def on_generate_click(self):
        st = self.start_year.strip()
        en = self.end_year.strip()

        if not (st.isdigit() and len(st) == 4 and en.isdigit() and len(en) == 4):
            messagebox.showwarning(
                "Проверка данных",
                "Пожалуйста, укажите корректные 4-значные года (например, 2024 / 2025)."
            )
            self.notebook.select(0)
            self.combo_year.focus_set()
            return

        cur_input = self.year_var.get().strip()
        self.add_to_history(cur_input)

        default_name = f"Обложка_журнала_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить обложку журнала в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )

        if not save_path:
            return

        try:
            generate_excel_cover(st, en, save_path)
            self.last_saved_file = save_path
            self.btn_open.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Файл Excel успешно сформирован:\n{save_path}\n\n"
                f"Учебный год: {st} / {en}\n"
                f"Параметры: А4, левое поле 30 мм под подшивку.\n"
                f"Запись сохранена в config.json"
            )
        except Exception as e:
            self.status_var.set("Ошибка при сохранении файла.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_file_click(self):
        if not self.last_saved_file or not os.path.exists(self.last_saved_file):
            messagebox.showwarning("Файл не найден", "Созданный файл не найден на диске.")
            return

        filepath = os.path.abspath(self.last_saved_file)
        sys_os = platform.system()

        try:
            if sys_os == "Windows":
                os.startfile(filepath)
            elif sys_os == "Darwin":
                subprocess.run(["open", filepath], check=True)
            else:
                subprocess.run(["xdg-open", filepath], check=True)
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{e}")

    def on_teacher_combo_select(self, event=None):
        """Срабатывает при выборе ФИО педагога из выпадающего списка."""
        name = self.combo_teacher.get().strip()
        if name:
            self.teacher_var.set(name)
            self.add_to_teacher_history(name)
            self.draw_inside_cover_preview()
            self.status_var.set(f"Педагог выбран из истории: {name}")

    def on_save_teacher_click(self):
        """Сохраняет введенное ФИО педагога в историю config.json."""
        name = self.teacher_var.get().strip()
        if not name:
            self.status_var.set("Пожалуйста, введите ФИО педагога перед сохранением.")
            return
        self.add_to_teacher_history(name)
        if hasattr(self, 'combo_teacher') and isinstance(self.combo_teacher, ttk.Combobox):
            self.combo_teacher['values'] = self.config_data.get("teacher_history", [])
        self.draw_inside_cover_preview()
        self.status_var.set(f"ФИО педагога «{name}» сохранено в config.json и применено.")

    def on_generate_inside_cover_click(self):
        """Формирует и сохраняет в Excel страницу «Оборот обложки» (стр. 2)."""
        teacher = self.teacher_var.get().strip()
        if teacher:
            self.add_to_teacher_history(teacher)

        default_name = "Оборот_обложки_журнала_стр2.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить оборот обложки журнала в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )

        if not save_path:
            return

        try:
            generate_excel_inside_cover(teacher, save_path)
            self.last_saved_inside_file = save_path
            if hasattr(self, 'btn_open_inside'):
                self.btn_open_inside.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Файл «Оборот обложки» успешно сформирован:\n{save_path}\n\n"
                f"Педагог: {teacher if teacher else '(не указан, оставлена пустая линия)'}\n"
                f"Параметры: А4, отступ под корешок 30 мм СПРАВА (левая страница разворота).\n"
                f"Запись сохранена в config.json"
            )
        except Exception as e:
            self.status_var.set("Ошибка при сохранении оборота обложки.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_inside_file_click(self):
        """Открывает сгенерированный файл оборота обложки в системной программе."""
        if not hasattr(self, 'last_saved_inside_file') or not self.last_saved_inside_file or not os.path.exists(self.last_saved_inside_file):
            messagebox.showwarning("Файл не найден", "Созданный файл оборота обложки не найден на диске.")
            return

        filepath = os.path.abspath(self.last_saved_inside_file)
        sys_os = platform.system()

        try:
            if sys_os == "Windows":
                os.startfile(filepath)
            elif sys_os == "Darwin":
                subprocess.run(["open", filepath], check=True)
            else:
                subprocess.run(["xdg-open", filepath], check=True)
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{e}")

    def on_generate_title_click(self):
        """Формирует и сохраняет в Excel страницу «Титульный лист» (стр. 1)."""
        org = self.org_var.get().strip()
        if org:
            self.add_to_org_history(org)
        title_yr = self.title_year_var.get().strip()
        sd = self.start_day_var.get().strip()
        sm = self.start_month_var.get().strip()
        sy = self.start_year_var.get().strip()
        ed = self.end_day_var.get().strip()
        em = self.end_month_var.get().strip()
        ey = self.end_year_var.get().strip()

        default_name = "Титульный_лист_журнала_стр1.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить титульный лист журнала в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )

        if not save_path:
            return

        try:
            generate_excel_title_page(
                org_name=org,
                academic_year=title_yr,
                start_month=sm,
                start_year=sy,
                end_month=em,
                end_year=ey,
                filename=save_path,
                start_day=sd,
                end_day=ed
            )
            self.last_saved_title_file = save_path
            if hasattr(self, 'btn_open_title'):
                self.btn_open_title.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Файл «Титульный лист» успешно сформирован:\n{save_path}\n\n"
                f"Организация: {org}\n"
                f"Учебный год: {title_yr}\n"
                f"Параметры: А4, отступ 30 мм СЛЕВА под скоросшиватель.\n"
                f"Запись сохранена в config.json"
            )
        except Exception as e:
            self.status_var.set("Ошибка при сохранении титульного листа.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_title_file_click(self):
        """Открывает сгенерированный файл титульного листа в системной программе."""
        if not hasattr(self, 'last_saved_title_file') or not self.last_saved_title_file or not os.path.exists(self.last_saved_title_file):
            messagebox.showwarning("Файл не найден", "Созданный файл титульного листа не найден на диске.")
            return

        filepath = os.path.abspath(self.last_saved_title_file)
        sys_os = platform.system()

        try:
            if sys_os == "Windows":
                os.startfile(filepath)
            elif sys_os == "Darwin":
                subprocess.run(["open", filepath], check=True)
            else:
                subprocess.run(["xdg-open", filepath], check=True)
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{e}")

    def schedule_auto_save_month(self, month_key: str):
        """Отложенное сохранение данных месяца (debounced на 500 мс)."""
        if not hasattr(self, '_month_save_timers'):
            self._month_save_timers = {}
        timer = self._month_save_timers.get(month_key)
        if timer:
            try:
                self.after_cancel(timer)
            except Exception:
                pass
        self._month_save_timers[month_key] = self.after(500, lambda k=month_key: self.auto_save_month_data(k))

    def schedule_auto_save_mass_events(self):
        """Отложенное сохранение мероприятий (debounced на 500 мс)."""
        if hasattr(self, '_events_save_timer') and self._events_save_timer:
            try:
                self.after_cancel(self._events_save_timer)
            except Exception:
                pass
        self._events_save_timer = self.after(500, self.auto_save_mass_events_data)

    def schedule_auto_save_creative(self):
        """Отложенное сохранение достижений (debounced на 500 мс)."""
        if hasattr(self, '_creative_save_timer') and self._creative_save_timer:
            try:
                self.after_cancel(self._creative_save_timer)
            except Exception:
                pass
        self._creative_save_timer = self.after(500, self.auto_save_creative_achievements_data)

    def save_all_pending_data(self):
        """Мгновенно сохраняет все измененные данные всех вкладок в config.json."""
        try:
            # Обложка
            if hasattr(self, 'year_var'):
                val_yr = self.year_var.get().strip()
                if val_yr:
                    self.config_data["last_academic_year"] = val_yr
            # Титульный лист
            if hasattr(self, 'org_var'):
                self.config_data["org_name"] = self.org_var.get().strip()
            if hasattr(self, 'title_year_var'):
                self.config_data["title_academic_year"] = self.title_year_var.get().strip()
            if hasattr(self, 'start_day_var'):
                self.config_data["start_day_val"] = self.start_day_var.get().strip()
            if hasattr(self, 'start_month_var'):
                self.config_data["start_month"] = self.start_month_var.get().strip()
            if hasattr(self, 'start_year_var'):
                self.config_data["start_year_val"] = self.start_year_var.get().strip()
            if hasattr(self, 'end_day_var'):
                self.config_data["end_day_val"] = self.end_day_var.get().strip()
            if hasattr(self, 'end_month_var'):
                self.config_data["end_month"] = self.end_month_var.get().strip()
            if hasattr(self, 'end_year_var'):
                self.config_data["end_year_val"] = self.end_year_var.get().strip()
            # Оборот титульного
            if hasattr(self, 'teacher_var'):
                self.config_data["teacher_name"] = self.teacher_var.get().strip()
            # Основные данные (стр. 3)
            self.auto_save_main_data()
            if hasattr(self, 'months_data'):
                for m in MONTHS_CONFIG:
                    self.auto_save_month_data(m["key"])
            if hasattr(self, 'mass_events_p1_vars'):
                self.auto_save_mass_events_data()
            if hasattr(self, 'creative_achievements_p1_vars'):
                self.auto_save_creative_achievements_data()
            if hasattr(self, 'students_list_pages_vars'):
                self.auto_save_students_list_data()
            if hasattr(self, 'safety_briefing_p1_vars'):
                self.auto_save_safety_briefing_data()
            if hasattr(self, 'annual_report_vars'):
                self.auto_save_annual_report_data()
            if hasattr(self, 'work_hours_vars'):
                self.auto_save_work_hours_data()
            save_config(self.config_data)
        except Exception as e:
            print(f"Предупреждение: Не удалось сохранить все данные: {e}")

    def on_window_close(self):
        """Гарантирует сохранение всех данных перед закрытием окна."""
        try:
            self.save_all_pending_data()
        except Exception:
            pass
        self.destroy()

    def auto_save_main_data(self, *args):
        """Сохраняет состояние полей вкладки 'Основные данные' в config.json."""
        if not hasattr(self, 'main_org_var'):
            return
        self.config_data["main_org_name"] = self.main_org_var.get().strip()
        self.config_data["department"] = self.department_var.get().strip()
        self.config_data["association"] = self.association_var.get().strip()
        self.config_data["group_name"] = self.group_var.get().strip()
        self.config_data["study_year"] = self.study_year_var.get().strip()
        self.config_data["rukovoditel"] = self.rukovoditel_var.get().strip()
        self.config_data["starosta"] = self.starosta_var.get().strip()
        self.config_data["accompanist"] = self.accompanist_var.get().strip()
        self.config_data["accompanist_schedule"] = self.acc_schedule_var.get().strip()
        self.config_data["accompanist_changes"] = self.acc_changes_var.get().strip()

        # Расписание 6 строк
        sch_list = []
        for dv, tv in self.schedule_row_vars:
            sch_list.append({"day": dv.get().strip(), "time": tv.get().strip()})
        self.config_data["schedule"] = sch_list

        # Изменения расписания
        chg_list = []
        for item in self.schedule_changes_row_widgets:
            d_val = item["day_var"].get().strip()
            t_val = item["time_var"].get().strip()
            if d_val or t_val:
                chg_list.append({"day": d_val, "time": t_val})
        self.config_data["schedule_changes"] = chg_list

        save_config(self.config_data)
        self.update_summary_lbl()

    def add_schedule_change_row(self, day="", time="", auto_save=True):
        """Добавляет строку в таблицу 'Изменение расписания занятий'."""
        if not hasattr(self, 'changes_rows_frame') or not self.changes_rows_frame.winfo_exists():
            return
        row_f = tk.Frame(self.changes_rows_frame, bg=self.changes_rows_frame.cget("bg"))
        row_f.pack(fill=tk.X, pady=2)

        day_var = tk.StringVar(value=day)
        time_var = tk.StringVar(value=time)

        e_day = ttk.Entry(row_f, textvariable=day_var, font=(self.font_sans, 9), width=24)
        e_day.pack(side=tk.LEFT, padx=(0, 6))
        self.attach_context_menu(e_day)
        e_day.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        e_time = ttk.Entry(row_f, textvariable=time_var, font=(self.font_sans, 9))
        e_time.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.attach_context_menu(e_time)
        e_time.bind("<KeyPress>", self.handle_entry_shortcuts, add="+")

        item_data = {"day_var": day_var, "time_var": time_var, "frame": row_f}
        self.schedule_changes_row_widgets.append(item_data)

        def on_delete():
            if item_data in self.schedule_changes_row_widgets:
                self.schedule_changes_row_widgets.remove(item_data)
            row_f.destroy()
            self.auto_save_main_data()

        btn_del = tk.Button(
            row_f,
            text="✕",
            command=on_delete,
            font=(self.font_sans, 8, "bold"),
            bg="#fee2e2",
            fg="#dc2626",
            activebackground="#fecaca",
            activeforeground="#b91c1c",
            relief=tk.SOLID,
            bd=1,
            padx=6,
            pady=1,
            cursor="hand2"
        )
        btn_del.pack(side=tk.RIGHT)

        day_var.trace_add("write", lambda *a: self.auto_save_main_data())
        time_var.trace_add("write", lambda *a: self.auto_save_main_data())

        if auto_save:
            self.auto_save_main_data()
            e_day.focus_set()

    def on_save_main_data_manual(self):
        """Ручное сохранение основных данных по кнопке."""
        self.auto_save_main_data()
        self.status_var.set("Основные данные успешно сохранены в config.json.")
        messagebox.showinfo("Сохранено", "Основные данные успешно сохранены в config.json и не удалятся после перезапуска программы.")

    def auto_save_month_data(self, month_key: str = "september", *args):
        """Автоматически сохраняет все данные вкладки месяца в self.config_data и на диск."""
        if not hasattr(self, 'months_data') or month_key not in self.months_data:
            return
        try:
            m_data = self.months_data[month_key]
            st_list = [v.get() for v in m_data["student_vars"]]
            dt_list = [v.get() for v in m_data["date_vars"]]
            att_list = [[cell.get() for cell in row] for row in m_data["attendance_vars"]]
            top_list = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "hours_teacher": r[2].get(),
                    "sign_teacher": r[3].get(),
                    "hours_acc": r[4].get(),
                    "sign_acc": r[5].get()
                }
                for r in m_data["topic_row_vars"]
            ]
            self.config_data[f"{month_key}_students"] = st_list
            self.config_data[f"{month_key}_dates"] = dt_list
            self.config_data[f"{month_key}_attendance"] = att_list
            self.config_data[f"{month_key}_topics"] = top_list

            if month_key == "september":
                self.config_data["september_students"] = st_list
                self.config_data["september_dates"] = dt_list
                self.config_data["september_attendance"] = att_list
                self.config_data["september_topics"] = top_list

            save_config(self.config_data)
        except Exception:
            pass

    def auto_save_september_data(self, *args):
        """Автоматически сохраняет все данные вкладки 'Сентябрь' (для обратной совместимости)."""
        self.auto_save_month_data("september", *args)

    def get_clipboard_text(self):
        """Надежно извлекает текстовые данные из системного буфера обмена (Tkinter / Win32)."""
        for _ in range(3):
            try:
                text = self.clipboard_get()
                if text:
                    return text
            except Exception:
                import time
                time.sleep(0.04)

        if sys.platform.startswith("win"):
            try:
                import ctypes
                CF_UNICODETEXT = 13
                u32 = ctypes.windll.user32
                k32 = ctypes.windll.kernel32
                if u32.OpenClipboard(None):
                    try:
                        h_data = u32.GetClipboardData(CF_UNICODETEXT)
                        if h_data:
                            k32.GlobalLock.restype = ctypes.c_void_p
                            ptr = k32.GlobalLock(h_data)
                            if ptr:
                                try:
                                    text = ctypes.wstring_at(ptr)
                                    if text:
                                        return text
                                finally:
                                    k32.GlobalUnlock(h_data)
                    finally:
                        u32.CloseClipboard()
            except Exception:
                pass

        return ""

    def is_paste_event(self, event):
        """Определяет, нажато ли сочетание вставки (Ctrl+V / Cmd+V / Shift+Insert) на любой раскладке клавиатуры (RU / EN)."""
        state = getattr(event, 'state', 0)
        is_ctrl = bool(state & 0x4) or bool(state & 0x8) or bool(state & 0x40000) or ('control' in str(getattr(event, 'keysym', '')).lower())
        code = getattr(event, 'keycode', 0)
        key = (getattr(event, 'keysym', '') or "").lower()
        char = getattr(event, 'char', '')

        if is_ctrl:
            if code in (86, 54):
                return True
            if key in ('v', 'cyrillic_em', 'cyrillic_m'):
                return True
            if char in ('\x16', 'v', 'V', 'м', 'М'):
                return True

        if (state & 0x1) and (code == 45 or key == 'insert'):
            return True

        return False

    def on_student_entry_keypress(self, event, idx: int, month_key: str = "september"):
        """Обработка нажатий клавиш в строке обучающегося (Ctrl+V, Ctrl+A, Ctrl+C, Ctrl+X)."""
        if self.is_paste_event(event):
            self.paste_students_from_clipboard(start_idx=idx, event=event, month_key=month_key)
            return "break"

        state = getattr(event, 'state', 0)
        is_ctrl = bool(state & 0x4) or bool(state & 0x8) or bool(state & 0x40000)
        code = getattr(event, 'keycode', 0)
        key = (getattr(event, 'keysym', '') or "").lower()

        if is_ctrl:
            if key in ('a', 'cyrillic_ef') or code == 65:
                self.select_all_widget(event.widget)
                return "break"
            elif key in ('c', 'cyrillic_es') or code == 67:
                self.copy_from_widget(event.widget)
                return "break"
            elif key in ('x', 'cyrillic_che') or code == 88:
                self.cut_from_widget(event.widget)
                return "break"

        return None

    def on_student_entry_paste(self, event, idx: int, month_key: str = "september"):
        """Перехват события <<Paste>> для строки обучающегося."""
        self.paste_students_from_clipboard(start_idx=idx, event=event, month_key=month_key)
        return "break"

    def on_topic_entry_keypress(self, event, idx: int, month_key: str = "september"):
        """Обработка нажатий клавиш в строке темы ДОП (Ctrl+V, Ctrl+A, Ctrl+C, Ctrl+X)."""
        if self.is_paste_event(event):
            self.paste_topics_from_clipboard(start_idx=idx, event=event, month_key=month_key)
            return "break"

        state = getattr(event, 'state', 0)
        is_ctrl = bool(state & 0x4) or bool(state & 0x8) or bool(state & 0x40000)
        code = getattr(event, 'keycode', 0)
        key = (getattr(event, 'keysym', '') or "").lower()

        if is_ctrl:
            if key in ('a', 'cyrillic_ef') or code == 65:
                self.select_all_widget(event.widget)
                return "break"
            elif key in ('c', 'cyrillic_es') or code == 67:
                self.copy_from_widget(event.widget)
                return "break"
            elif key in ('x', 'cyrillic_che') or code == 88:
                self.cut_from_widget(event.widget)
                return "break"

        return None

    def on_topic_entry_paste(self, event, idx: int, month_key: str = "september"):
        """Перехват события <<Paste>> для строки темы ДОП."""
        self.paste_topics_from_clipboard(start_idx=idx, event=event, month_key=month_key)
        return "break"

    def on_global_keypress(self, event):
        """Глобальный перехватчик клавиш окна для вставки Ctrl+V на любой раскладке клавиатуры."""
        if self.is_paste_event(event):
            res = self.on_global_paste_shortcut(event)
            if res == "break":
                return "break"
        return None

    def on_global_paste_shortcut(self, event=None, month_key: str = None):
        """Глобальный перехватчик сочетания Ctrl+V для активной вкладки месяца."""
        if month_key is None:
            try:
                cur_tab = self.notebook.index(self.notebook.select())
            except Exception:
                return None

            # Вкладки с 4 по 15 соответствуют 12 месяцам (Сентябрь — Август)
            if 4 <= cur_tab < 4 + len(MONTHS_CONFIG):
                month_key = MONTHS_CONFIG[cur_tab - 4]["key"]
            elif cur_tab == 16:
                # Вкладка 16: Учёт массовых мероприятий
                if hasattr(self, 'mass_events_notebook') and self.mass_events_notebook:
                    try:
                        cur_page = self.mass_events_notebook.index(self.mass_events_notebook.select()) + 1
                    except Exception:
                        cur_page = 1
                    return self.paste_mass_events_from_clipboard(page_num=cur_page, start_idx=0, event=event)
                return None
            elif cur_tab == 17:
                # Вкладка 17: Творческие достижения
                if hasattr(self, 'creative_achievements_notebook') and self.creative_achievements_notebook:
                    try:
                        cur_page = self.creative_achievements_notebook.index(self.creative_achievements_notebook.select()) + 1
                    except Exception:
                        cur_page = 1
                    return self.paste_creative_achievements_from_clipboard(page_num=cur_page, start_idx=0, event=event)
                return None
            elif cur_tab == 18:
                # Вкладка 18: Список обучающихся (6 страниц)
                if hasattr(self, 'students_list_notebook') and self.students_list_notebook:
                    try:
                        cur_page = self.students_list_notebook.index(self.students_list_notebook.select()) + 1
                    except Exception:
                        cur_page = 1
                    return self.paste_students_list_from_clipboard(page_num=cur_page, start_idx=0, col_idx=0, event=event)
                return None
            else:
                return None

        if not hasattr(self, 'months_data') or month_key not in self.months_data:
            return None

        m_data = self.months_data[month_key]
        m_notebook = m_data.get("notebook")
        if not m_notebook:
            return None

        try:
            cur_page = m_notebook.index(m_notebook.select())
        except Exception:
            cur_page = 0

        focused = self.focus_get()

        if cur_page == 0:
            start_idx = 0
            for i, w in enumerate(m_data.get("student_widgets", [])):
                if w == focused:
                    start_idx = i
                    break
            self.paste_students_from_clipboard(start_idx=start_idx, event=event, month_key=month_key)
            return "break"

        elif cur_page == 1:
            start_idx = 0
            for i, w in enumerate(m_data.get("topic_widgets", [])):
                if w == focused:
                    start_idx = i
                    break
            self.paste_topics_from_clipboard(start_idx=start_idx, event=event, month_key=month_key)
            return "break"

        return None

    def paste_students_from_clipboard(self, start_idx=0, event=None, month_key: str = "september"):
        """Вставляет скопированный из буфера список учеников построчно, начиная со строки start_idx."""
        raw_text = self.get_clipboard_text()
        if not raw_text:
            self.status_var.set("Буфер обмена пуст или не содержит текст.")
            return "break"

        raw_lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
        if not raw_lines:
            self.status_var.set("В буфере обмена нет строк для вставки.")
            return "break"

        cleaned_names = []
        for line in raw_lines:
            parts = line.split('\t')
            if len(parts) >= 2 and parts[0].strip().isdigit():
                name = parts[1].strip()
            else:
                name = parts[0].strip()
            name = re.sub(r'^\d+[\.\)\s\-]+\s*', '', name)
            if name:
                cleaned_names.append(name)
            elif line:
                cleaned_names.append(line)

        m_data = self.months_data.get(month_key, self.months_data.get("september"))
        if not m_data:
            return "break"

        count = 0
        for offset, name in enumerate(cleaned_names):
            idx = start_idx + offset
            if idx < 30:
                m_data["student_vars"][idx].set(name)
                count += 1

        self.auto_save_month_data(month_key)
        m_name = m_data["config"]["name"]
        self.status_var.set(f"Вставлено {count} обучающихся из буфера обмена ({m_name}, со строки {start_idx + 1}).")
        return "break"

    def paste_topics_from_clipboard(self, start_idx=0, event=None, month_key: str = "september"):
        """Вставляет скопированный список тем занятий ДОП построчно, начиная со строки start_idx."""
        raw_text = self.get_clipboard_text()
        if not raw_text:
            self.status_var.set("Буфер обмена пуст или не содержит текст.")
            return "break"

        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
        if not lines:
            self.status_var.set("В буфере обмена нет строк для вставки.")
            return "break"

        m_data = self.months_data.get(month_key, self.months_data.get("september"))
        if not m_data:
            return "break"

        count = 0
        for offset, line in enumerate(lines):
            idx = start_idx + offset
            if idx < 16:
                parts = line.split('\t')
                if len(parts) >= 3:
                    m_data["topic_row_vars"][idx][0].set(parts[0].strip())
                    m_data["topic_row_vars"][idx][1].set(parts[1].strip())
                    m_data["topic_row_vars"][idx][2].set(parts[2].strip())
                elif len(parts) == 2 and ('.' in parts[0] or '/' in parts[0]):
                    m_data["topic_row_vars"][idx][0].set(parts[0].strip())
                    m_data["topic_row_vars"][idx][1].set(parts[1].strip())
                else:
                    clean_theme = re.sub(r'^(?:Тема\s*\d+[\.\:\s]*|\d+[\.\)\s\-]+)\s*', '', line)
                    m_data["topic_row_vars"][idx][1].set(clean_theme or line)
                count += 1

        self.auto_save_month_data(month_key)
        m_name = m_data["config"]["name"]
        self.status_var.set(f"Вставлено {count} тем занятий из буфера обмена ({m_name}, со строки {start_idx + 1}).")
        return "break"

    def copy_students_from_month(self, target_month_key: str, source_month_key: str = "september"):
        """Копирует список обучающихся из одного месяца в другой."""
        if target_month_key not in self.months_data or source_month_key not in self.months_data:
            return
        src_names = [v.get() for v in self.months_data[source_month_key]["student_vars"]]
        non_empty = [n for n in src_names if n.strip()]
        if not non_empty:
            messagebox.showinfo("Список пуст", f"В месяце «{self.months_data[source_month_key]['config']['name']}» нет заполненных имен детей.")
            return
        for i in range(30):
            self.months_data[target_month_key]["student_vars"][i].set(src_names[i])
        self.auto_save_month_data(target_month_key)
        tgt_name = self.months_data[target_month_key]["config"]["name"]
        src_name = self.months_data[source_month_key]["config"]["name"]
        self.status_var.set(f"Скопировано {len(non_empty)} обучающихся из «{src_name}» в «{tgt_name}».")

    def clear_month_students(self, month_key: str):
        """Очищает имена всех 30 обучающихся указанного месяца."""
        if month_key not in self.months_data:
            return
        m_name = self.months_data[month_key]["config"]["name"]
        if not messagebox.askyesno("Подтверждение", f"Очистить список всех 30 обучающихся ({m_name}, страница 1)?"):
            return
        for v in self.months_data[month_key]["student_vars"]:
            v.set("")
        self.auto_save_month_data(month_key)
        self.status_var.set(f"Список обучающихся ({m_name}) очищен.")

    def clear_september_students(self):
        """Совместимость: очищает имена 30 обучающихся Сентября."""
        self.clear_month_students("september")

    def clear_month_dates_and_attendance(self, month_key: str):
        """Очищает даты и все отметки посещаемости указанного месяца."""
        if month_key not in self.months_data:
            return
        m_name = self.months_data[month_key]["config"]["name"]
        if not messagebox.askyesno("Подтверждение", f"Очистить все 15 дат занятий и отметки посещаемости ({m_name})?"):
            return
        for dv in self.months_data[month_key]["date_vars"]:
            dv.set("")
        for row in self.months_data[month_key]["attendance_vars"]:
            for cv in row:
                cv.set("")
        self.auto_save_month_data(month_key)
        self.status_var.set(f"Даты и отметки посещаемости ({m_name}) очищены.")

    def clear_september_dates_and_attendance(self):
        """Совместимость: очищает даты и отметки посещаемости Сентября."""
        self.clear_month_dates_and_attendance("september")

    def clear_month_topics(self, month_key: str):
        """Очищает таблицу тем и занятий (16 строк) указанного месяца."""
        if month_key not in self.months_data:
            return
        m_name = self.months_data[month_key]["config"]["name"]
        if not messagebox.askyesno("Подтверждение", f"Очистить все 16 строк содержания занятий ДОП ({m_name}, страница 2)?"):
            return
        for row in self.months_data[month_key]["topic_row_vars"]:
            for var in row:
                var.set("")
        self.auto_save_month_data(month_key)
        self.status_var.set(f"Таблица содержания занятий ({m_name}) очищена.")

    def clear_september_topics(self):
        """Совместимость: очищает содержание занятий Сентября."""
        self.clear_month_topics("september")

    def copy_dates_to_month_page2(self, month_key: str):
        """Копирует 15 дат со страницы 1 в первые строки страницы 2 указанного месяца."""
        if month_key not in self.months_data:
            return
        dates = [v.get().strip() for v in self.months_data[month_key]["date_vars"]]
        copied_count = 0
        for i, d in enumerate(dates):
            if i < 16 and d:
                self.months_data[month_key]["topic_row_vars"][i][0].set(d)
                copied_count += 1
        self.auto_save_month_data(month_key)
        m_name = self.months_data[month_key]["config"]["name"]
        self.status_var.set(f"Скопировано {copied_count} дат со Страницы 1 на Страницу 2 ({m_name}).")

    def copy_dates_to_september_page2(self):
        """Совместимость: копирует даты Сентября со страницы 1 на страницу 2."""
        self.copy_dates_to_month_page2("september")

    def fill_default_hours_month(self, month_key: str):
        """Заполняет часы (2 ч) для всех строк, где указана дата или тема в указанном месяце."""
        if month_key not in self.months_data:
            return
        filled = 0
        for row in self.months_data[month_key]["topic_row_vars"]:
            d = row[0].get().strip()
            c = row[1].get().strip()
            if d or c:
                if not row[2].get().strip():
                    row[2].set("2")
                    filled += 1
        self.auto_save_month_data(month_key)
        m_name = self.months_data[month_key]["config"]["name"]
        self.status_var.set(f"Установлено 2 часа для {filled} занятий ({m_name}).")

    def fill_default_hours_september(self):
        """Совместимость: заполняет часы занятий Сентября."""
        self.fill_default_hours_month("september")

    def _get_current_export_month(self):
        """Возвращает конфигурацию выбранного для экспорта месяца."""
        m_name = self.export_month_var.get().strip() if hasattr(self, 'export_month_var') else "Сентябрь"
        for m in MONTHS_CONFIG:
            if m["name"] == m_name or m["key"] == m_name:
                return m
        return MONTHS_CONFIG[0]

    def on_generate_month_p1_click(self, month_key: str = None):
        """Формирует и сохраняет отдельный Excel-файл листа 'Посещаемость' выбранного месяца."""
        if month_key:
            m_cfg = next((m for m in MONTHS_CONFIG if m["key"] == month_key), MONTHS_CONFIG[0])
        else:
            m_cfg = self._get_current_export_month()

        m_key = m_cfg["key"]
        m_name = m_cfg["name"]
        p1_no = m_cfg["p1"]
        self.auto_save_month_data(m_key)
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"{m_name}_Посещаемость_стр{p1_no}_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title=f"Сохранить {m_name}: Посещаемость (стр. {p1_no}) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            m_data = self.months_data[m_key]
            students = [v.get() for v in m_data["student_vars"]]
            dates = [v.get() for v in m_data["date_vars"]]
            attendance = [[c.get() for c in row] for row in m_data["attendance_vars"]]
            generate_excel_month_p1(save_path, month_name=m_name, page_number=p1_no, students=students, dates=dates, attendance=attendance)
            self.last_saved_month_p1_file[m_key] = save_path
            self.last_saved_september_p1_file = save_path
            if hasattr(self, 'btn_open_month_p1'):
                self.btn_open_month_p1.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Файл '{m_name}: Учёт посещаемости' (стр. {p1_no}) успешно создан:\n{save_path}"
            )
        except Exception as e:
            self.status_var.set(f"Ошибка при сохранении страницы посещаемости ({m_name}).")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_month_p1_file_click(self, month_key: str = None):
        """Открывает сгенерированный файл Посещаемости выбранного месяца."""
        if month_key:
            m_cfg = next((m for m in MONTHS_CONFIG if m["key"] == month_key), MONTHS_CONFIG[0])
        else:
            m_cfg = self._get_current_export_month()
        m_key = m_cfg["key"]
        fpath = self.last_saved_month_p1_file.get(m_key) or self.last_saved_september_p1_file
        if not fpath or not os.path.exists(fpath):
            messagebox.showwarning("Файл не найден", f"Созданный файл для месяца «{m_cfg['name']}» не найден на диске.")
            return
        self._open_file_system(fpath)

    def on_generate_september_p1_click(self):
        """Совместимость: формирует Excel-файл Сентябрь: Посещаемость."""
        self.on_generate_month_p1_click("september")

    def on_open_september_p1_file_click(self):
        """Совместимость: открывает файл Сентябрь: Посещаемость."""
        self.on_open_month_p1_file_click("september")

    def on_generate_month_p2_click(self, month_key: str = None):
        """Формирует и сохраняет отдельный Excel-файл листа 'Выполнение ДОП' выбранного месяца."""
        if month_key:
            m_cfg = next((m for m in MONTHS_CONFIG if m["key"] == month_key), MONTHS_CONFIG[0])
        else:
            m_cfg = self._get_current_export_month()

        m_key = m_cfg["key"]
        m_name = m_cfg["name"]
        p2_no = m_cfg["p2"]
        self.auto_save_month_data(m_key)
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"{m_name}_Темы_ДОП_стр{p2_no}_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title=f"Сохранить {m_name}: Темы ДОП (стр. {p2_no}) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            m_data = self.months_data[m_key]
            topics = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "hours_teacher": r[2].get(),
                    "sign_teacher": r[3].get(),
                    "hours_acc": r[4].get(),
                    "sign_acc": r[5].get()
                }
                for r in m_data["topic_row_vars"]
            ]
            generate_excel_month_p2(save_path, month_name=m_name, page_number=p2_no, topics=topics)
            self.last_saved_month_p2_file[m_key] = save_path
            self.last_saved_september_p2_file = save_path
            if hasattr(self, 'btn_open_month_p2'):
                self.btn_open_month_p2.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Файл '{m_name}: Содержание занятий по ДОП' (стр. {p2_no}) успешно создан:\n{save_path}"
            )
        except Exception as e:
            self.status_var.set(f"Ошибка при сохранении страницы содержания занятий ({m_name}).")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_month_p2_file_click(self, month_key: str = None):
        """Открывает сгенерированный файл Темы ДОП выбранного месяца."""
        if month_key:
            m_cfg = next((m for m in MONTHS_CONFIG if m["key"] == month_key), MONTHS_CONFIG[0])
        else:
            m_cfg = self._get_current_export_month()
        m_key = m_cfg["key"]
        fpath = self.last_saved_month_p2_file.get(m_key) or self.last_saved_september_p2_file
        if not fpath or not os.path.exists(fpath):
            messagebox.showwarning("Файл не найден", f"Созданный файл для месяца «{m_cfg['name']}» не найден на диске.")
            return
        self._open_file_system(fpath)

    def on_generate_september_p2_click(self):
        """Совместимость: формирует Excel-файл Сентябрь: Темы ДОП."""
        self.on_generate_month_p2_click("september")

    def on_open_september_p2_file_click(self):
        """Совместимость: открывает файл Сентябрь: Темы ДОП."""
        self.on_open_month_p2_file_click("september")

    def on_generate_month_spread_click(self, month_key: str = None):
        """Формирует и сохраняет разворот выбранного месяца (обе страницы) в один файл Excel."""
        if month_key:
            m_cfg = next((m for m in MONTHS_CONFIG if m["key"] == month_key), MONTHS_CONFIG[0])
        else:
            m_cfg = self._get_current_export_month()

        m_key = m_cfg["key"]
        m_name = m_cfg["name"]
        p1_no = m_cfg["p1"]
        p2_no = m_cfg["p2"]
        self.auto_save_month_data(m_key)
        st = self.start_year.strip()
        en = self.end_year.strip()
        default_name = f"{m_name}_Разворот_стр{p1_no}-{p2_no}_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title=f"Сохранить Разворот {m_name} (стр. {p1_no} и {p2_no}) в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return
        try:
            m_data = self.months_data[m_key]
            students = [v.get() for v in m_data["student_vars"]]
            dates = [v.get() for v in m_data["date_vars"]]
            attendance = [[c.get() for c in row] for row in m_data["attendance_vars"]]
            topics = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "hours_teacher": r[2].get(),
                    "sign_teacher": r[3].get(),
                    "hours_acc": r[4].get(),
                    "sign_acc": r[5].get()
                }
                for r in m_data["topic_row_vars"]
            ]
            generate_excel_month_spread(
                save_path,
                month_name=m_name,
                p1_no=p1_no,
                p2_no=p2_no,
                students=students,
                dates=dates,
                attendance=attendance,
                topics=topics
            )
            self.last_saved_month_spread_file[m_key] = save_path
            self.last_saved_september_spread_file = save_path
            if hasattr(self, 'btn_open_month_spread'):
                self.btn_open_month_spread.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен разворот: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Разворот месяца {m_name} (стр. {p1_no} и стр. {p2_no}) успешно сохранен в файле:\n{save_path}"
            )
        except Exception as e:
            self.status_var.set(f"Ошибка при сохранении разворота месяца {m_name}.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_month_spread_file_click(self, month_key: str = None):
        """Открывает сгенерированный файл разворота выбранного месяца."""
        if month_key:
            m_cfg = next((m for m in MONTHS_CONFIG if m["key"] == month_key), MONTHS_CONFIG[0])
        else:
            m_cfg = self._get_current_export_month()
        m_key = m_cfg["key"]
        fpath = self.last_saved_month_spread_file.get(m_key) or self.last_saved_september_spread_file
        if not fpath or not os.path.exists(fpath):
            messagebox.showwarning("Файл не найден", f"Созданный файл разворота для месяца «{m_cfg['name']}» не найден на диске.")
            return
        self._open_file_system(fpath)

    def on_generate_september_spread_click(self):
        """Совместимость: формирует разворот Сентября."""
        self.on_generate_month_spread_click("september")

    def on_open_september_spread_file_click(self):
        """Совместимость: открывает файл разворота Сентября."""
        self.on_open_month_spread_file_click("september")

    def _open_file_system(self, filepath: str):
        """Открывает файл системным приложением."""
        actual_fp = os.path.abspath(filepath)
        sys_os = platform.system()
        try:
            if sys_os == "Windows":
                os.startfile(actual_fp)
            elif sys_os == "Darwin":
                subprocess.run(["open", actual_fp], check=True)
            else:
                subprocess.run(["xdg-open", actual_fp], check=True)
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{e}")

    def on_generate_main_data_click(self):
        """Формирует и сохраняет отдельный Excel-файл листа 'Основные данные' (стр. 3 журнала)."""
        self.auto_save_main_data()
        org = self.main_org_var.get().strip()
        dept = self.department_var.get().strip()
        assoc = self.association_var.get().strip()
        grp = self.group_var.get().strip()
        yr = self.study_year_var.get().strip()
        ruk = self.rukovoditel_var.get().strip()
        star = self.starosta_var.get().strip()
        acc = self.accompanist_var.get().strip()
        acc_s = self.acc_schedule_var.get().strip()
        acc_c = self.acc_changes_var.get().strip()

        sch_list = []
        for dv, tv in self.schedule_row_vars:
            sch_list.append({"day": dv.get().strip(), "time": tv.get().strip()})

        chg_list = []
        for item in self.schedule_changes_row_widgets:
            d_val = item["day_var"].get().strip()
            t_val = item["time_var"].get().strip()
            if d_val or t_val:
                chg_list.append({"day": d_val, "time": t_val})

        default_name = "Основные_данные_журнала_стр3.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Основные данные в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not save_path:
            return

        try:
            generate_excel_main_data(
                org_name=org,
                department=dept,
                association=assoc,
                group_name=grp,
                study_year=yr,
                schedule=sch_list,
                schedule_changes=chg_list,
                rukovoditel=ruk,
                starosta=star,
                accompanist=acc,
                accompanist_schedule=acc_s,
                accompanist_changes=acc_c,
                filename=save_path
            )
            self.last_saved_main_file = save_path
            if hasattr(self, 'btn_open_main'):
                self.btn_open_main.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен лист Основные данные: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Файл 'Основные данные' успешно сформирован:\n{save_path}\n\n"
                f"Организация: {org}\n"
                f"Объединение: {assoc}, Группа: {grp}\n"
                f"Параметры: А4, отступ 30 мм СЛЕВА под скоросшиватель.\n"
                f"Запись сохранена в config.json"
            )
        except Exception as e:
            self.status_var.set("Ошибка при сохранении листа Основные данные.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_main_file_click(self):
        """Открывает сгенерированный файл листа 'Основные данные' в системной программе."""
        if not hasattr(self, 'last_saved_main_file') or not self.last_saved_main_file or not os.path.exists(self.last_saved_main_file):
            messagebox.showwarning("Файл не найден", "Созданный файл листа Основные данные не найден на диске.")
            return

        filepath = os.path.abspath(self.last_saved_main_file)
        sys_os = platform.system()

        try:
            if sys_os == "Windows":
                os.startfile(filepath)
            elif sys_os == "Darwin":
                subprocess.run(["open", filepath], check=True)
            else:
                subprocess.run(["xdg-open", filepath], check=True)
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{e}")

    def on_generate_full_click(self):
        """Формирует и сохраняет в один Excel-файл полный комплект журнала (все 4 страницы)."""
        st = self.start_year.strip()
        en = self.end_year.strip()
        cur_input = self.year_var.get().strip()
        if cur_input:
            self.add_to_history(cur_input)

        org = self.org_var.get().strip()
        if org:
            self.add_to_org_history(org)

        teacher = self.teacher_var.get().strip()
        if teacher:
            self.add_to_teacher_history(teacher)

        title_yr = self.title_year_var.get().strip()
        sd = self.start_day_var.get().strip()
        sm = self.start_month_var.get().strip()
        sy = self.start_year_var.get().strip()
        ed = self.end_day_var.get().strip()
        em = self.end_month_var.get().strip()
        ey = self.end_year_var.get().strip()

        # Основные данные (стр. 3)
        self.auto_save_main_data()
        main_org = self.main_org_var.get().strip()
        dept = self.department_var.get().strip()
        assoc = self.association_var.get().strip()
        grp = self.group_var.get().strip()
        yr = self.study_year_var.get().strip()
        ruk = self.rukovoditel_var.get().strip()
        star = self.starosta_var.get().strip()
        acc = self.accompanist_var.get().strip()
        acc_s = self.acc_schedule_var.get().strip()
        acc_c = self.acc_changes_var.get().strip()

        sch_list = []
        for dv, tv in self.schedule_row_vars:
            sch_list.append({"day": dv.get().strip(), "time": tv.get().strip()})

        chg_list = []
        for item in self.schedule_changes_row_widgets:
            d_val = item["day_var"].get().strip()
            t_val = item["time_var"].get().strip()
            if d_val or t_val:
                chg_list.append({"day": d_val, "time": t_val})

        # Все 12 месяцев учебного года (стр. 4–27 журнала)
        month_kwargs = {}
        for m in MONTHS_CONFIG:
            k = m["key"]
            self.auto_save_month_data(k)
            m_data = self.months_data[k]
            month_kwargs[f"{k}_students"] = [v.get() for v in m_data["student_vars"]]
            month_kwargs[f"{k}_dates"] = [v.get() for v in m_data["date_vars"]]
            month_kwargs[f"{k}_attendance"] = [[c.get() for c in row] for row in m_data["attendance_vars"]]
            month_kwargs[f"{k}_topics"] = [
                {
                    "date": r[0].get(),
                    "content": r[1].get(),
                    "hours_teacher": r[2].get(),
                    "sign_teacher": r[3].get(),
                    "hours_acc": r[4].get(),
                    "sign_acc": r[5].get()
                }
                for r in m_data["topic_row_vars"]
            ]

        # Учёт массовых мероприятий (стр. 30 и 31)
        self.auto_save_mass_events_data()
        ev_p1 = [
            {
                "date": r[0].get(),
                "content": r[1].get(),
                "count": r[2].get(),
                "location": r[3].get(),
                "conducted_by": r[4].get()
            }
            for r in self.mass_events_p1_vars
        ]
        ev_p2 = [
            {
                "date": r[0].get(),
                "content": r[1].get(),
                "count": r[2].get(),
                "location": r[3].get(),
                "conducted_by": r[4].get()
            }
            for r in self.mass_events_p2_vars
        ]

        # Творческие достижения (стр. 32 и 33)
        self.auto_save_creative_achievements_data()
        cr_p1 = [
            {
                "student": r[0].get(),
                "event": r[1].get()
            }
            for r in self.creative_achievements_p1_vars
        ]
        cr_p2 = [
            {
                "results": r[0].get(),
                "works": r[1].get()
            }
            for r in self.creative_achievements_p2_vars
        ]

        # Список обучающихся (стр. 34-39 журнала, 6 страниц / 3 разворота)
        self.auto_save_students_list_data()
        st_p1 = [
            {
                "student": r[0].get(),
                "birth_year": r[1].get(),
                "school_class": r[2].get(),
                "district": r[3].get(),
                "doctor_conclusion": r[4].get()
            }
            for r in self.students_list_pages_vars[1]
        ]
        st_p2 = [
            {
                "address_phone": r[0].get(),
                "parents_info": r[1].get(),
                "join_date": r[2].get(),
                "leave_info": r[3].get(),
                "notes": r[4].get()
            }
            for r in self.students_list_pages_vars[2]
        ]
        st_p3 = [
            {
                "student": r[0].get(),
                "birth_year": r[1].get(),
                "school_class": r[2].get(),
                "district": r[3].get(),
                "doctor_conclusion": r[4].get()
            }
            for r in self.students_list_pages_vars[3]
        ]
        st_p4 = [
            {
                "address_phone": r[0].get(),
                "parents_info": r[1].get(),
                "join_date": r[2].get(),
                "leave_info": r[3].get(),
                "notes": r[4].get()
            }
            for r in self.students_list_pages_vars[4]
        ]
        st_p5 = [
            {
                "student": r[0].get(),
                "birth_year": r[1].get(),
                "school_class": r[2].get(),
                "district": r[3].get(),
                "doctor_conclusion": r[4].get()
            }
            for r in self.students_list_pages_vars[5]
        ]
        st_p6 = [
            {
                "address_phone": r[0].get(),
                "parents_info": r[1].get(),
                "join_date": r[2].get(),
                "leave_info": r[3].get(),
                "notes": r[4].get()
            }
            for r in self.students_list_pages_vars[6]
        ]

        # Инструктаж по технике безопасности (стр. 38-39 журнала, 2 страницы по 28 строк)
        self.auto_save_safety_briefing_data()
        sb_p1 = [
            {
                "student": r[0].get(),
                "date": r[1].get(),
                "content": r[2].get(),
                "signature": r[3].get()
            }
            for r in self.safety_briefing_p1_vars
        ]
        sb_p2 = [
            {
                "student": r[0].get(),
                "date": r[1].get(),
                "content": r[2].get(),
                "signature": r[3].get()
            }
            for r in self.safety_briefing_p2_vars
        ]

        # Годовой цифровой отчёт (стр. 40 журнала)
        self.auto_save_annual_report_data()
        ar_data = [
            {
                "period": r["period"].get(),
                "total": r["total"].get(),
                "boys": r["boys"].get(),
                "girls": r["girls"].get(),
                "classes": [cv.get() for cv in r["classes"]],
                "years_in_org": [yv.get() for yv in r["years"]]
            }
            for r in self.annual_report_vars
        ]

        default_name = f"Полный_комплект_журнала_{st}-{en}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить полный комплект журнала в Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )

        if not save_path:
            return

        try:
            generate_excel_full_journal(
                start_year=st,
                end_year=en,
                org_name=org,
                title_academic_year=title_yr,
                start_month=sm,
                start_year_val=sy,
                end_month=em,
                end_year_val=ey,
                teacher_name=teacher,
                filename=save_path,
                start_day=sd,
                end_day=ed,
                main_org_name=main_org,
                department=dept,
                association=assoc,
                group_name=grp,
                study_year=yr,
                schedule=sch_list,
                schedule_changes=chg_list,
                rukovoditel=ruk,
                starosta=star,
                accompanist=acc,
                accompanist_schedule=acc_s,
                accompanist_changes=acc_c,
                mass_events_p1=ev_p1,
                mass_events_p2=ev_p2,
                creative_achievements_p1=cr_p1,
                creative_achievements_p2=cr_p2,
                students_list_p1=st_p1,
                students_list_p2=st_p2,
                students_list_p3=st_p3,
                students_list_p4=st_p4,
                students_list_p5=st_p5,
                students_list_p6=st_p6,
                safety_briefing_p1=sb_p1,
                safety_briefing_p2=sb_p2,
                annual_report=ar_data,
                **month_kwargs
            )
            self.last_saved_full_file = save_path
            if hasattr(self, 'btn_open_full'):
                self.btn_open_full.config(state=tk.NORMAL)
            self.status_var.set(f"Успешно сохранен полный комплект: {os.path.basename(save_path)}")
            messagebox.showinfo(
                "Успешно!",
                f"Полный комплект журнала успешно сформирован в одном файле:\n{save_path}\n\n"
                f"В книге создано 42 листа (с чистым оборотом обложки для двусторонней печати):\n"
                f"1. Обложка (внешняя, нечётная сторона, переплет слева)\n"
                f"2. Оборот обложки (чистый лист, чётная сторона, переплет справа)\n"
                f"3. Титульный лист (стр. 1, нечётная сторона, переплет слева)\n"
                f"4. Оборот титульного (стр. 2, чётная сторона, переплет справа)\n"
                f"5. Основные данные (стр. 3, нечётная сторона, переплет слева)\n"
                f"6–29. Учебные месяцы: Сентябрь — Август (стр. 4–27 с чередованием полей под переплет)\n"
                f"30–31. Учёт массовых мероприятий с обучающимися (стр. 30–31)\n"
                f"32–33. Творческие достижения обучающихся (стр. 32–33)\n"
                f"34–39. Список обучающихся: 3 разворота по 10 чел. (стр. 34–39)\n"
                f"40–41. Список обучающихся, прошедших инструктаж по ТБ (стр. 38–39)\n"
                f"42. Годовой цифровой отчёт (стр. 40, чётная сторона, переплет справа)\n\n"
                f"Все данные и история сохранены в config.json"
            )
        except Exception as e:
            self.status_var.set("Ошибка при сохранении полного комплекта журнала.")
            messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{e}")

    def on_open_full_file_click(self):
        """Открывает сгенерированный файл полного комплекта журнала в системной программе."""
        if not hasattr(self, 'last_saved_full_file') or not self.last_saved_full_file or not os.path.exists(self.last_saved_full_file):
            messagebox.showwarning("Файл не найден", "Созданный файл полного комплекта журнала не найден на диске.")
            return

        filepath = os.path.abspath(self.last_saved_full_file)
        sys_os = platform.system()

        try:
            if sys_os == "Windows":
                os.startfile(filepath)
            elif sys_os == "Darwin":
                subprocess.run(["open", filepath], check=True)
            else:
                subprocess.run(["xdg-open", filepath], check=True)
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{e}")


def main():
    app = JournalCoverApp()
    app.mainloop()


if __name__ == "__main__":
    main()
