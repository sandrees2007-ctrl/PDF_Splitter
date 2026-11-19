"""
Flask-приложение: серверная версия PDFSplitterApp (main.py).
Вся бизнес-логика перенесена без изменений; UI-слой заменён
на REST API + SSE для стриминга логов и прогресса.
"""

import os
import json
import sys
import re
import shutil
import threading
import uuid
import time
import base64
import datetime
import difflib
import zipfile
import math
import queue
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import (
    Flask, request, jsonify, Response, send_from_directory, send_file, stream_with_context, render_template
)

try:
    import pymupdf as fitz
except ImportError:
    import fitz

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests as http_requests
from requests.adapters import HTTPAdapter

# ---------------------------------------------------------------------------
# Константы (из main.py без изменений)
# ---------------------------------------------------------------------------
DEFAULT_PROXYAPI_KEY = "sk-W02LrjSK9PapzakiwgjkP0IBkFN7p3rE"
DEFAULT_AI_BASE_URL = "https://api.proxyapi.ru/v1"
DEFAULT_AI_MODEL = "google/gemma-4-26b-a4b-it"

# ---------------------------------------------------------------------------
# Вспомогательные чистые функции (из main.py без изменений)
# ---------------------------------------------------------------------------

def find_registry_columns(headers):
    """Определяет индексы колонок (1-based) в реестре по заголовкам."""
    cols = {
        'date': 2, 'bldg': 3, 'building': 3, 'struct': 4,
        'nz7': 5, 'nz28': 7, 'cub7': 9, 'cub28': 11,
        '7_nz': 5, '28_nz': 7, '7_cub': 9, '28_cub': 11
    }
    for idx, h in enumerate(headers, start=1):
        v = str(h or "").lower()
        if "дата бетонирования" in v: cols['date'] = idx
        elif "корпус" in v: cols['bldg'] = cols['building'] = idx
        elif "наименование" in v: cols['struct'] = idx
        elif "неразрушающ" in v and "7" in v: cols['nz7'] = cols['7_nz'] = idx
        elif "неразрушающ" in v and "28" in v: cols['nz28'] = cols['28_nz'] = idx
        elif ("кубик" in v or "куб" in v) and "7" in v: cols['cub7'] = cols['7_cub'] = idx
        elif ("кубик" in v or "куб" in v) and "28" in v: cols['cub28'] = cols['28_cub'] = idx
    return cols


def get_target_col(cols, age, is_cube):
    """Вычисляет целевой столбец реестра по возрасту и признаку кубиков."""
    is_7 = ("7" in str(age)) and ("28" not in str(age))
    if is_cube:
        return cols.get('cub7', cols.get('7_cub', 9)) if is_7 else cols.get('cub28', cols.get('28_cub', 11))
    return cols.get('nz7', cols.get('7_nz', 5)) if is_7 else cols.get('nz28', cols.get('28_nz', 7))


def dates_match(d1, d2):
    """Сравнивает даты с учётом 2/4-значных годов и очистки от пробелов."""
    if not d1 and not d2: return True
    if not d1 or not d2: return False
    if d1 == d2: return True
    p1, p2 = str(d1).replace(' ', '').split('.'), str(d2).replace(' ', '').split('.')
    if len(p1) == 3 and len(p2) == 3 and p1[0] == p2[0] and p1[1] == p2[1]:
        return p1[2][-2:] == p2[2][-2:]
    c1, c2 = str(d1).replace('.', '').replace(' ', ''), str(d2).replace('.', '').replace(' ', '')
    return (c1 in c2) or (c2 in c1)


def buildings_match(b1, b2, sheet_name=""):
    """Проверяет соответствие корпусов/зданий по номерам и подстрокам."""
    if not b1 or not b2: return True
    n1 = re.findall(r'\d+\.?\d*', str(b1))
    n2 = re.findall(r'\d+\.?\d*', str(b2))
    ns = re.findall(r'\d+\.?\d*', str(sheet_name)) if sheet_name else []
    if n1 and n2 and any(n in n2 for n in n1): return True
    if n1 and ns and any(n in ns for n in n1): return True
    return (b1.lower() in b2.lower()) or (b2.lower() in b1.lower())


def structures_match(s1, s2):
    """Нечёткое сравнение названий конструкций."""
    if not s1 or not s2: return False
    c1, c2 = str(s1).lower().replace(" ", ""), str(s2).lower().replace(" ", "")
    if (c1[:12] in c2) or (c2[:12] in c1): return True
    return bool(difflib.get_close_matches(str(s1), [str(s2)], n=1, cutoff=0.4))


def compute_match_score(p_bldg, p_cdate, p_strc, e_bldg, e_cdate, e_strc, sheet_name="", target_has_val=False):
    """Подсчитывает балл соответствия протокола существующей строке таблицы."""
    if not buildings_match(p_bldg, e_bldg, sheet_name):
        return -1, False, False
    d_match = dates_match(p_cdate, e_cdate)
    s_match = structures_match(p_strc, e_strc)
    if not (d_match or s_match):
        return -1, False, False
    score = (10 if (d_match and p_cdate) else 0) + (5 if s_match else 0) - (3 if target_has_val else 0)
    return score, d_match, s_match


# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
    static_url_path="/static"
)

# Глобальный реестр фоновых задач: task_id -> TaskState
_tasks: dict = {}
_tasks_lock = threading.Lock()


class TaskState:
    """Состояние одной фоновой задачи обработки."""
    def __init__(self, task_id: str):
        self.task_id = task_id
        self.queue: queue.Queue = queue.Queue()
        self.done = False
        self.records: list = []
        self.files: list = []

    # --- адаптеры, заменяющие self.log() / self.progress_bar.set() ---
    def log(self, message: str):
        self.queue.put({"type": "log", "message": message})

    def set_progress(self, value: float):
        self.queue.put({"type": "progress", "value": round(value, 4)})

    def finish(self, records: list, files: list):
        self.records = records
        self.files = files
        self.done = True
        self.queue.put({"type": "done", "records": records, "files": files})


# ---------------------------------------------------------------------------
# Загрузка / сохранение настроек (из main.py)
# ---------------------------------------------------------------------------

def load_settings() -> dict:
    defaults = {
        "theme": "Liquid Cyan",
        "ai_api_key": DEFAULT_PROXYAPI_KEY,
        "ai_base_url": DEFAULT_AI_BASE_URL,
        "ai_model": DEFAULT_AI_MODEL,
        "ai_workers": 10,
        "registry_mode": "local",
        "registry_path": None,
        "google_creds_path": None,
        "google_url": "",
    }
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                s = json.load(f)
            theme_map = {
                "Dark Orange": "Liquid Sunset",
                "Dark Teal": "Liquid Cyan",
                "Light Blue": "Liquid Cyan",
                "Light Green": "Liquid Emerald"
            }
            raw_theme = s.get("theme", "Liquid Cyan")
            s["theme"] = theme_map.get(raw_theme, raw_theme)
            defaults.update(s)
        except Exception:
            pass

    # Авто-поиск credentials.json рядом
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    auto_creds = os.path.join(base_path, "credentials.json")
    if os.path.exists(auto_creds) and not defaults.get("google_creds_path"):
        defaults["google_creds_path"] = auto_creds

    return defaults


def save_settings(data: dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Бизнес-логика (перенос методов PDFSplitterApp)
# ---------------------------------------------------------------------------

def _make_http_session():
    session = http_requests.Session()
    adapter = HTTPAdapter(pool_connections=40, pool_maxsize=40, max_retries=1)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def prepare_page_data(page, page_num=0, total_pages=1):
    """Быстрое извлечение текста и ультра-быстрый сжатый рендеринг изображения в MuPDF."""
    text_layer = page.get_text("text").strip()
    has_good_text = len(text_layer) > 120 and any(k in text_layer.lower() for k in ["протокол", "испытан", "бетон"])
    use_vision = not has_good_text

    b64_img = None
    if use_vision:
        rect = page.rect
        max_dim = max(rect.width, rect.height)
        target_dim = 1600.0
        scale = min(2.5, target_dim / max_dim) if max_dim > 0 else 1.5
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        jpeg_bytes = pix.tobytes("jpeg", jpg_quality=85)
        b64_img = base64.b64encode(jpeg_bytes).decode("ascii")

    return {
        "page_num": page_num,
        "text_layer": text_layer[:3000] if text_layer else "",
        "use_vision": use_vision,
        "b64_img": b64_img
    }


def query_ai_for_page(page_data, total_pages, settings, http_session, task: TaskState):
    """Отправка подготовленной страницы в ИИ через постоянную HTTP-сессию."""
    page_num = page_data["page_num"]
    text_layer = page_data["text_layer"]
    use_vision = page_data["use_vision"]
    b64_img = page_data["b64_img"]

    api_key = (settings.get("ai_api_key") or DEFAULT_PROXYAPI_KEY).strip()
    base_url = (settings.get("ai_base_url") or DEFAULT_AI_BASE_URL).rstrip('/')
    model = settings.get("ai_model") or DEFAULT_AI_MODEL

    prompt_system = (
        "Ты эксперт по анализу строительных протоколов испытания бетона лаборатории ООО «СтройЛабГрупп».\n"
        "Извлеки данные страницы строго в формате JSON:\n"
        "{\n"
        '  "is_protocol": true,\n'
        '  "protocol_number": "СВБ-28-0107_26-2",\n'
        '  "protocol_date": "29.07.2026",\n'
        '  "building": "корпус 4",\n'
        '  "structure": "наименование конструкции/элемента",\n'
        '  "concreting_date": "01.07.2026",\n'
        '  "age_days": "28",\n'
        '  "is_cube": true,\n'
        '  "is_continuation": false\n'
        "}\n\n"
        "КРИТИЧЕСКИ ВАЖНЫЕ ПРАВИЛА И СТРУКТУРА НОМЕРА:\n"
        "1. СТРУКТУРА НОМЕРА ПРОТОКОЛА: СВБ-{префикс}-{ДДММ}_{ГГ}-{серия} (все слэши замени на _).\n"
        "   - Префиксы:\n"
        "     * 28: образцы-кубы, проектный возраст 28 суток (is_cube: true).\n"
        "     * 7: образцы-кубы, промежуточный возраст 7 суток (is_cube: true).\n"
        "     * 10: протоколы на конструкцию / неразрушающий контроль (is_cube: false). ВНИМАНИЕ: протоколы на конструкцию обозначаются префиксом 10 КАК НА 7 СУТОК, ТАК И НА 28 СУТОК! Точный возраст (7 или 28) определяй по графе «Возраст, сут.», заключению или по разнице дат (дата испытания минус дата бетонирования).\n"
        "     * 12: редкое исключение для протоколов на конструкцию (неразрушающий контроль, 28 суток, is_cube: false).\n"
        "   - {ДДММ}: 4 цифры — это СТРОГО ДЕНЬ И МЕСЯЦ ДАТЫ БЕТОНИРОВАНИЯ / ИЗГОТОВЛЕНИЯ ОБРАЗЦОВ! (ДД.ММ).\n"
        "   - {ГГ}: 2 последние цифры года (например, 2026 -> 26).\n"
        "   - {серия}: номер серии/партии (например, -1, -2, -3).\n"
        "2. ВНИМАНИЕ: СИНЯЯ КРУГЛАЯ ПЕЧАТЬ В ШАПКЕ ПЕРЕКРЫВАЕТ НОМЕР!\n"
        "   Круглая синяя печать лаборатории («СтройЛабГрупп») часто наложена прямо на номер протокола в шапке, из-за чего цифры искажаются: например, «0107» ошибочно кажется как «8187», «0307» — как «03706», «2607» — как «2697».\n"
        "   НИКОГДА не пытайся читать размытые цифры только из шапки под печатью!\n"
        "   ОБЯЗАТЕЛЬНО найди в таблице результатов испытаний (раздел 11) колонку «Дата изготовления» (или строку «Дата бетонирования»).\n"
        "   Четыре цифры {ДДММ} в номере протокола ОБЯЗАНЫ в точности совпадать с днем и месяцем изготовления/бетонирования!\n"
        "   Если дата изготовления 01.07.2026 — в номере протокола ОБЯЗАНО быть 0107 (а НЕ 8187)! Если 26.07.2026 — 2607, если 03.07.2026 — 0307, если 18.07.2026 — 1807.\n"
        "3. МАТЕМАТИЧЕСКАЯ СВЯЗЬ ДАТ:\n"
        "   Дата протокола (проведения испытаний) = Дата изготовления + Возраст (суток).\n"
        "   - При возрасте 28 суток: 01.07.2026 + 28 сут = 29.07.2026; 03.07.2026 + 28 сут = 31.07.2026; 13.07.2026 + 28 сут = 10.08.2026.\n"
        "   - При возрасте 7 суток: 01.07.2026 + 7 сут = 08.07.2026; 26.07.2026 + 7 сут = 02.08.2026.\n"
        "   Используй эту формулу для взаимной проверки и восстановления дат!\n"
        "4. is_cube: true если в таблице размер образцов 10х10х10 см / 100х100х100 мм (ГОСТ 10180, кубы), иначе false (ГОСТ 22690 / неразрушающий контроль конструкции).\n"
        "5. is_continuation: true если это 2-я страница протокола (нет шапки с заказчиком, таблица продолжается).\n"
        "Ответь ИСКЛЮЧИТЕЛЬНО валидным JSON без каких-либо markdown блоков."
    )

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    if use_vision and b64_img:
        payload = {
            "model": model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt_system},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{b64_img}",
                                "detail": "high"
                            }
                        }
                    ]
                }
            ],
            "temperature": 0.1,
            "max_tokens": 500,
            "response_format": {"type": "json_object"}
        }
        mode_desc = f"Vision ({model})"
    else:
        payload = {
            "model": model,
            "messages": [
                {"role": "user", "content": f"{prompt_system}\n\nТекст страницы:\n{text_layer}"}
            ],
            "temperature": 0.1,
            "max_tokens": 500,
            "response_format": {"type": "json_object"}
        }
        mode_desc = f"Text ({model})"

    for attempt in range(2):
        try:
            task.log(f"  [>] Стр. {page_num}/{total_pages}: отправка в ИИ [{mode_desc}]...")
            resp = http_session.post(
                f"{base_url}/chat/completions",
                headers=headers,
                json=payload,
                timeout=(5, 30)
            )
            if resp.status_code == 200:
                raw_content = resp.json()["choices"][0]["message"]["content"].strip()
                if raw_content.startswith("```"):
                    lines = raw_content.split("\n")
                    if lines[0].startswith("```"):
                        lines = lines[1:]
                    if lines and lines[-1].startswith("```"):
                        lines = lines[:-1]
                    raw_content = "\n".join(lines).strip()
                return json.loads(raw_content)
            elif resp.status_code == 429:
                time.sleep(1.0 * (attempt + 1))
            else:
                task.log(f"  [!] Ответ ProxyAPI ({resp.status_code}) на стр. {page_num}: {resp.text[:150]}")
                time.sleep(0.5)
        except Exception as ex:
            task.log(f"  [!] Ошибка запроса к ИИ (стр. {page_num}, попытка {attempt+1}/2): {str(ex)}")
            time.sleep(0.5)

    return None


def clean_protocol_number(num_str, concreting_date=None, protocol_date=None, age=None):
    if not num_str:
        return None
    num_str = str(num_str).strip()

    core_str = re.sub(r'^[^0-9]*[СC][ВB][БB5]\s*-?',
                      '', num_str, flags=re.IGNORECASE)

    replacements = {
        "O": "0", "o": "0", "О": "0", "о": "0",
        "l": "1", "I": "1", "І": "1", "|": "1", "!": "1",
        "З": "3", "б": "6", "В": "8", "B": "8"
    }
    for k, v in replacements.items():
        core_str = core_str.replace(k, v)

    core_str = re.sub(r'^[^0-9]+', '', core_str)
    core_str = core_str.replace("/", "_").replace("\\", "_")

    core_str = re.sub(r'([0-9]+)46\-1', r'\1_26-1', core_str)
    core_str = re.sub(r'([0-9]+)46\-', r'\1_26-', core_str)

    if core_str.startswith("707") and "26" in core_str:
        core_str = "10-1" + core_str

    prefix = None
    match = re.search(r'(10|12|28|7)[\-_]', core_str)
    if match:
        prefix = match.group(1)
        core_str = core_str[match.start():]
    else:
        match_broken = re.search(r'^(0|2|8)[\-_]', core_str)
        if match_broken:
            digit = match_broken.group(1)
            prefix = '10' if digit == '0' else ('12' if digit == '2' else '28')
            core_str = prefix + core_str[1:]
        elif age:
            age_str = str(age).strip()
            prefix = '28' if '28' in age_str else ('7' if '7' in age_str else '10')
            core_str = f'{prefix}_{core_str}'
        else:
            match_any = re.search(r'(10|12|28|7)', core_str)
            if match_any:
                prefix = match_any.group(1)
                core_str = core_str[match_any.start():]

    target_ddmm = None
    target_yy = "26"

    def _parse_d(d_val):
        if not d_val:
            return None
        for fmt in ("%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.datetime.strptime(str(d_val).strip(), fmt)
            except Exception:
                pass
        return None

    dt_c = _parse_d(concreting_date)
    dt_p = _parse_d(protocol_date)

    if not dt_c and dt_p:
        if age and str(age).isdigit():
            dt_c = dt_p - datetime.timedelta(days=int(age))
        elif prefix in ('28', '12'):
            dt_c = dt_p - datetime.timedelta(days=28)
        elif prefix == '7':
            dt_c = dt_p - datetime.timedelta(days=7)
        else:
            m_num_dm = re.search(r'[\-_](\d{2})(\d{2})[\-_]', core_str)
            if m_num_dm:
                try:
                    cand_d = datetime.datetime(dt_p.year, int(m_num_dm.group(2)), int(m_num_dm.group(1)))
                    diff = (dt_p - cand_d).days
                    if (4 <= diff <= 14) or (20 <= diff <= 35):
                        dt_c = cand_d
                except Exception:
                    pass
            if not dt_c:
                dt_c = dt_p - datetime.timedelta(days=28)

    if dt_c:
        target_ddmm = f"{dt_c.day:02d}{dt_c.month:02d}"
        target_yy = f"{dt_c.year % 100:02d}"

    m_full = re.search(r'^(10|12|28|7)[\-_](\d{3,5})[\-_](\d{2})[\-_](\d+)', core_str)
    if m_full:
        p_val, raw_ddmm, raw_yy, ser = m_full.groups()
        ddmm_to_use = target_ddmm if target_ddmm else raw_ddmm
        yy_to_use = target_yy if target_yy else raw_yy
        core_str = f"{p_val}-{ddmm_to_use}_{yy_to_use}-{ser}"
    else:
        m_flex = re.search(r'^(10|12|28|7)[\-_](\d{3,5})[\-_/]?.*?-(\d+)', core_str)
        if m_flex and target_ddmm:
            p_val, raw_ddmm, ser = m_flex.groups()
            core_str = f"{p_val}-{target_ddmm}_{target_yy}-{ser}"
        elif target_ddmm and prefix:
            m_ser = re.search(r'-(\d+)$', core_str)
            ser = m_ser.group(1) if m_ser else '1'
            core_str = f"{prefix}-{target_ddmm}_{target_yy}-{ser}"

    core_str = core_str.replace("/", "_").replace("\\", "_")
    return f"СВБ-{core_str}"


def create_formatted_registry(all_records, output_filepath):
    """Создает промежуточный реестр в точном формате итоговой таблицы."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Реестр протоколов"

    headers = [
        "№",
        "Дата бетонирования",
        "Корпус",
        "Наименование конструкции",
        "Протокол неразрушающего контроля на 7 суток (конструкция)",
        "В бумаге",
        "Протокол неразрушающего контроля на 28 суток (конструкция)",
        "В бумаге",
        "Протокол на 7 сутки (кубики)",
        "В бумаге",
        "Протокол на 28сутки (кубики)",
        "В бумаге",
        "Имя файла"
    ]

    ws.append(headers)

    font_header = Font(name="Calibri", size=11, bold=True)
    fill_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    font_link = Font(name="Calibri", size=11, color="0563C1", underline="single")
    font_cell = Font(name="Calibri", size=11)

    ws.row_dimensions[1].height = 45
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all

    grouped_rows = []
    group_map = {}

    for rec in all_records:
        p_bldg = str(rec.get("Корпус") or "").strip()
        p_cdate = str(rec.get("Дата бетонирования") or "").strip()
        p_strc = str(rec.get("Наименование конструкции") or "").strip()
        p_num = str(rec.get("Номер протокола") or "").strip()
        p_date = str(rec.get("Дата протокола") or "").strip()
        p_age = str(rec.get("Возраст суток") or "").strip()
        p_cube = rec.get("Кубики", False)
        filename = rec.get("Имя файла", "")
        p_prefix = rec.get("Префикс_Колонка", "")

        is_7 = ("7" in p_age) and ("28" not in p_age)
        if p_cube:
            target_col = 9 if is_7 else 11
        else:
            target_col = 5 if is_7 else 7

        entry = f"№{p_num} от {p_date}" if (p_num and p_date and p_date != "Не найдена") else f"№{p_num}"
        rel_path = f"Корпус {p_bldg}\\{p_prefix}\\{filename}" if p_bldg else f"Корпус_Неизвестен\\{p_prefix}\\{filename}"

        key = None
        if p_strc and (p_cdate or p_bldg):
            key = (p_bldg.lower(), p_cdate, p_strc.lower()[:30])

        if key and key in group_map and target_col not in group_map[key]["cols"]:
            row_data = group_map[key]
            row_data["cols"][target_col] = {
                "entry": entry,
                "rel_path": rel_path
            }
            if filename and filename not in row_data["filenames"]:
                row_data["filenames"].append(filename)
        else:
            row_data = {
                "cdate": p_cdate,
                "building": p_bldg,
                "structure": p_strc,
                "cols": {
                    target_col: {
                        "entry": entry,
                        "rel_path": rel_path
                    }
                },
                "filenames": [filename] if filename else []
            }
            grouped_rows.append(row_data)
            if key:
                group_map[key] = row_data

    for r_idx, rdata in enumerate(grouped_rows, start=2):
        ws.row_dimensions[r_idx].height = 28
        c_num = ws.cell(row=r_idx, column=1, value=r_idx - 1)
        c_num.alignment = align_center
        c_num.font = font_cell
        c_num.border = border_all

        c_date = ws.cell(row=r_idx, column=2, value=rdata["cdate"])
        c_date.alignment = align_center
        c_date.font = font_cell
        c_date.border = border_all

        c_bldg = ws.cell(row=r_idx, column=3, value=rdata["building"])
        c_bldg.alignment = align_center
        c_bldg.font = font_cell
        c_bldg.border = border_all

        c_strc = ws.cell(row=r_idx, column=4, value=rdata["structure"])
        c_strc.alignment = align_left
        c_strc.font = font_cell
        c_strc.border = border_all

        for col_idx in range(5, 13):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.border = border_all
            cell.alignment = align_center
            cell.font = font_cell
            if col_idx in rdata["cols"]:
                item = rdata["cols"][col_idx]
                cell.value = item["entry"]
                cell.fill = fill_green
                if item["rel_path"]:
                    cell.hyperlink = item["rel_path"]
                    cell.font = font_link

        c_file = ws.cell(row=r_idx, column=13, value=", ".join(rdata["filenames"]))
        c_file.alignment = align_left
        c_file.font = font_cell
        c_file.border = border_all

    col_widths = {
        1: 6, 2: 15, 3: 15, 4: 42, 5: 30, 6: 10,
        7: 30, 8: 10, 9: 30, 10: 10, 11: 30, 12: 10, 13: 35
    }
    for c_idx, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_idx)].width = w

    dir_name = os.path.dirname(output_filepath)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
    wb.save(output_filepath)


def perform_integration(all_records, project_output_dir, settings, task: TaskState):
    """Интеграция протоколов в существующий реестр (локальный Excel или Google Sheets)."""
    mode_val = settings.get("registry_mode", "local")
    unmatched_records = all_records.copy()

    registry_path = settings.get("registry_path")
    google_creds_path = settings.get("google_creds_path")
    google_url = settings.get("google_url", "")

    if mode_val == "local" and registry_path and os.path.exists(registry_path):
        task.log("\n[*] Начинаем интеграцию в локальный реестр Excel...")
        try:
            wb = openpyxl.load_workbook(registry_path)
            sheets_to_process = [sheet for sheet in wb.sheetnames if sheet != "Нераспределенные"]

            for sheet_name in sheets_to_process:
                ws = wb[sheet_name]
                task.log(f"  Проверка листа: {sheet_name}")

                hdr = [str(ws.cell(row=1, column=c).value or "").lower() for c in range(1, ws.max_column + 1)]
                cols = find_registry_columns(hdr)
                task.log(f"Найдены столбцы: дата={cols['date']}, здание={cols['bldg']}, констр={cols['struct']}, 7={cols['nz7']}, 28={cols['nz28']}, 7к={cols['cub7']}, 28к={cols['cub28']}")

                still_unmatched = []

                for rec in unmatched_records:
                    if rec["Номер протокола"] == "СВБ-НЕ_РАСПОЗНАН":
                        still_unmatched.append(rec)
                        continue

                    p_num = rec["Номер протокола"]
                    p_date = rec["Дата протокола"]
                    p_bldg = str(rec["Корпус"]).strip()
                    p_strc = str(rec.get("Наименование конструкции") or "").strip()
                    p_cdate = str(rec.get("Дата бетонирования") or "").strip()
                    p_age = str(rec.get("Возраст суток") or "").strip()
                    p_cube = rec.get("Кубики", False)
                    p_prefix = rec.get("Префикс_Колонка", "")

                    target_col = get_target_col(cols, p_age, p_cube)
                    best_row, best_score = None, 0

                    for r in range(2, ws.max_row + 1):
                        e_bldg = str(ws.cell(row=r, column=cols['bldg']).value or "").strip()
                        e_cdate = ws.cell(row=r, column=cols['date']).value
                        e_cdate = e_cdate.strftime("%d.%m.%Y") if isinstance(e_cdate, datetime.datetime) else str(e_cdate or "").strip()
                        e_strc = str(ws.cell(row=r, column=cols['struct']).value or "").strip()

                        if not e_bldg and not e_cdate and not e_strc:
                            continue

                        has_val = bool(str(ws.cell(row=r, column=target_col).value or "").strip())
                        score, d_match, s_match = compute_match_score(
                            p_bldg, p_cdate, p_strc, e_bldg, e_cdate, e_strc, sheet_name, has_val
                        )
                        if score > best_score and (d_match or s_match):
                            best_score = score
                            best_row = r

                    if best_row and target_col:
                        cell = ws.cell(row=best_row, column=target_col)
                        cell.value = f"№{p_num} от {p_date}"
                        filename = rec["Имя файла"]
                        cell.hyperlink = f"Корпус {p_bldg}\\{p_prefix}\\{filename}" if p_bldg else f"Корпус_Неизвестен\\{p_prefix}\\{filename}"
                        cell.style = "Hyperlink"
                        cell.fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
                    else:
                        still_unmatched.append(rec)

                unmatched_records = still_unmatched

            # Добавление новых строк в подходящий лист
            if unmatched_records:
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

                still_unmatched = []
                for rec in unmatched_records:
                    p_bldg = str(rec["Корпус"]).strip()
                    target_ws = None
                    for s_name in sheets_to_process:
                        s_clean = s_name.replace('Корпус', '').replace('корпус', '').strip()
                        if (p_bldg in s_name) or (s_clean and s_clean in p_bldg) or (p_bldg in s_clean):
                            target_ws = wb[s_name]
                            break

                    if target_ws and p_bldg:
                        new_row = target_ws.max_row + 1
                        t_headers = [str(target_ws.cell(row=1, column=c).value or "").lower() for c in range(1, target_ws.max_column + 1)]
                        t_cols = find_registry_columns(t_headers)

                        target_ws.cell(row=new_row, column=1).value = new_row - 1
                        target_ws.cell(row=new_row, column=t_cols['date']).value = rec.get("Дата бетонирования", "")
                        target_ws.cell(row=new_row, column=t_cols['bldg']).value = rec.get("Корпус", "")
                        target_ws.cell(row=new_row, column=t_cols['struct']).value = rec.get("Наименование конструкции", "")

                        p_num = rec["Номер протокола"]
                        p_date = rec["Дата протокола"]
                        p_age = str(rec.get("Возраст суток", "")).strip()
                        p_cube = rec.get("Кубики", False)
                        p_prefix = rec.get("Префикс_Колонка", "")

                        tc = get_target_col(t_cols, p_age, p_cube)
                        if tc:
                            cell = target_ws.cell(row=new_row, column=tc)
                            cell.value = f"№{p_num} от {p_date}"
                            cell.hyperlink = f"Корпус {p_bldg}\\{p_prefix}\\{rec['Имя файла']}"
                            cell.style = "Hyperlink"

                        for c in range(1, max(12, target_ws.max_column + 1)):
                            cell = target_ws.cell(row=new_row, column=c)
                            cell.border = thin_border
                            cell.alignment = center_align
                            cell.fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

                        task.log(f"    [+] Создана новая строка в {target_ws.title} для протокола №{p_num}")
                    else:
                        still_unmatched.append(rec)

                unmatched_records = still_unmatched

            if unmatched_records:
                if "Нераспределенные" in wb.sheetnames:
                    ws_un = wb["Нераспределенные"]
                else:
                    ws_un = wb.create_sheet("Нераспределенные")
                    ws_un.append(["Номер протокола", "Дата протокола", "Корпус", "Наименование конструкции",
                                  "Дата бетонирования", "Имя файла", "Ожидаемый столбец"])

                for rec in unmatched_records:
                    ws_un.append([
                        rec["Номер протокола"], rec["Дата протокола"], rec["Корпус"],
                        rec.get("Наименование конструкции", ""), rec.get("Дата бетонирования", ""),
                        rec["Имя файла"], rec["Префикс_Колонка"]
                    ])

            out_registry = os.path.join(project_output_dir, "Реестр_обновлен.xlsx")
            wb.save(out_registry)
            task.log(f"[✓] Локальный реестр успешно обновлен: {out_registry}")
        except Exception as e:
            task.log(f"[ОШИБКА] Не удалось обновить реестр: {str(e)}")

    elif mode_val == "google" and google_creds_path and google_url:
        task.log("\n[*] Начинаем интеграцию в Google Таблицу...")
        try:
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name(google_creds_path, scope)
            client = gspread.authorize(creds)
            url = google_url.strip()
            sh = client.open_by_url(url)

            for ws in sh.worksheets():
                if ws.title == "Нераспределенные":
                    continue
                task.log(f"  Проверка листа: {ws.title}")

                data = ws.get_all_values()
                if not data:
                    continue

                cols = find_registry_columns(data[0])
                task.log(
                    f"GS Найдены столбцы: дата={cols['date']}, здание={cols['building']}, констр={cols['struct']}, 7={cols['7_nz']}, 28={cols['28_nz']}, 7к={cols['7_cub']}, 28к={cols['28_cub']}")
                still_unmatched = []
                updates = []
                format_updates = []

                for rec in unmatched_records:
                    if rec["Номер протокола"] == "СВБ-НЕ_РАСПОЗНАН":
                        still_unmatched.append(rec)
                        continue

                    p_num = rec["Номер протокола"]
                    p_date = rec["Дата протокола"]
                    p_bldg = str(rec["Корпус"]).strip()
                    p_strc = str(rec["Наименование конструкции"]).strip() if rec["Наименование конструкции"] else ""
                    p_cdate = str(rec["Дата бетонирования"]).strip() if rec["Дата бетонирования"] else ""
                    p_age = str(rec["Возраст суток"]).strip() if rec["Возраст суток"] else ""
                    p_cube = rec.get("Кубики", False)

                    target_col = get_target_col(cols, p_age, p_cube)

                    best_row = None
                    best_score = 0

                    for r_idx, row in enumerate(data):
                        if r_idx == 0:
                            continue
                        e_bldg = str(row[cols['bldg'] - 1] if cols['bldg'] - 1 < len(row) else "").strip()
                        e_cdate = str(row[cols['date'] - 1] if cols['date'] - 1 < len(row) else "").strip()
                        e_strc = str(row[cols['struct'] - 1] if cols['struct'] - 1 < len(row) else "").strip()
                        if not e_bldg and not e_cdate and not e_strc:
                            continue

                        has_val = bool(target_col and target_col - 1 < len(row) and str(row[target_col - 1]).strip())
                        score, d_match, s_match = compute_match_score(
                            p_bldg, p_cdate, p_strc, e_bldg, e_cdate, e_strc, ws.title, has_val
                        )
                        if score > best_score and (d_match or s_match):
                            best_score = score
                            best_row = r_idx + 1

                    if best_row and target_col:
                        entry = f"№{p_num} от {p_date}"
                        updates.append(
                            {'range': f'{gspread.utils.rowcol_to_a1(best_row, target_col)}', 'values': [[entry]]})
                        task.log(
                            f'      [+] Совпадение! Протокол №{p_num} -> строка {best_row}, колонка {target_col}')

                        format_updates.append({
                            "repeatCell": {
                                "range": {
                                    "sheetId": ws.id,
                                    "startRowIndex": best_row - 1,
                                    "endRowIndex": best_row,
                                    "startColumnIndex": target_col - 1,
                                    "endColumnIndex": target_col
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "backgroundColor": {"red": 0.8, "green": 1.0, "blue": 0.8}
                                    }
                                },
                                "fields": "userEnteredFormat.backgroundColor"
                            }
                        })
                    else:
                        still_unmatched.append(rec)

                unmatched_records = still_unmatched
                if updates:
                    ws.batch_update(updates, value_input_option='USER_ENTERED')
                    time.sleep(1)
                if format_updates:
                    sh.batch_update({'requests': format_updates})
                    time.sleep(1)

            # Пакетная вставка новых строк
            if unmatched_records:
                still_unmatched = []
                grouped = {}
                for rec in unmatched_records:
                    k = (str(rec.get("Корпус", "")).strip(), str(rec.get("Дата бетонирования", "")).strip(), str(rec.get("Наименование конструкции", "")).strip())
                    if k not in grouped: grouped[k] = []
                    grouped[k].append(rec)

                all_worksheets = [w for w in sh.worksheets() if w.title != "Сводная"]
                ws_to_groups = {}

                for (p_bldg, p_cdate, p_strc), recs in grouped.items():
                    target_ws = None
                    for ws_check in all_worksheets:
                        if p_bldg and (p_bldg in ws_check.title):
                            target_ws = ws_check
                            break
                    if target_ws and p_bldg:
                        if target_ws not in ws_to_groups:
                            ws_to_groups[target_ws] = []
                        ws_to_groups[target_ws].append((p_bldg, p_cdate, p_strc, recs))
                    else:
                        for r in recs:
                            still_unmatched.append(r)

                for target_ws, items in ws_to_groups.items():
                    try:
                        data_target = target_ws.get_all_values()
                        current_rows = len(data_target)
                        t_cols = find_registry_columns(data_target[0]) if data_target else cols
                        batch_rows = []
                        format_requests = []

                        for i, (p_bldg, p_cdate, p_strc, recs) in enumerate(items):
                            row_idx = current_rows + 1 + i
                            row_data = ["" for _ in range(max(12, len(data_target[0]) if data_target else 12))]
                            row_data[0] = str(row_idx - 1)
                            row_data[t_cols['date'] - 1] = p_cdate
                            row_data[t_cols['building'] - 1] = p_bldg
                            row_data[t_cols['struct'] - 1] = p_strc

                            for rec in recs:
                                tc = get_target_col(t_cols, rec.get("Возраст суток", ""), rec.get("Кубики", False))
                                if tc:
                                    row_data[tc - 1] = f"№{rec['Номер протокола']} от {rec['Дата протокола']}"

                            batch_rows.append(row_data)

                            format_requests.append({
                                "repeatCell": {
                                    "range": {
                                        "sheetId": target_ws.id,
                                        "startRowIndex": row_idx - 1,
                                        "endRowIndex": row_idx,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": len(row_data)
                                    },
                                    "cell": {
                                        "userEnteredFormat": {
                                            "backgroundColor": {"red": 0.8, "green": 1.0, "blue": 0.8},
                                            "horizontalAlignment": "CENTER",
                                            "verticalAlignment": "MIDDLE",
                                            "wrapStrategy": "WRAP",
                                            "borders": {
                                                "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                                                "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                                                "left": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                                                "right": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                                            }
                                        }
                                    },
                                    "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,borders)"
                                }
                            })

                        if batch_rows:
                            target_ws.append_rows(batch_rows, value_input_option='USER_ENTERED')
                            time.sleep(0.5)
                            if format_requests:
                                sh.batch_update({'requests': format_requests})
                                time.sleep(0.5)
                            task.log(f"    [+] Пакетно добавлено {len(batch_rows)} новых строк в {target_ws.title}")
                    except Exception as ex_b:
                        task.log(f"    [!] Ошибка добавления строк в {target_ws.title}: {str(ex_b)}")
                        for _, _, _, recs in items:
                            for r in recs:
                                still_unmatched.append(r)

                unmatched_records = still_unmatched

            if unmatched_records:
                try:
                    ws_un = sh.worksheet("Нераспределенные")
                except gspread.exceptions.WorksheetNotFound:
                    ws_un = sh.add_worksheet(title="Нераспределенные", rows="100", cols="10")
                    ws_un.append_row(["Номер протокола", "Дата протокола", "Корпус",
                                      "Наименование конструкции", "Дата бетонирования", "Имя файла", "Ожидаемый столбец"])

                new_rows = []
                task.log(f"      [!] Отправлено в Нераспределенные: {len(unmatched_records)} шт.")
                for rec in unmatched_records:
                    new_rows.append([
                        rec["Номер протокола"], rec["Дата протокола"], rec["Корпус"],
                        rec.get("Наименование конструкции", ""), rec.get("Дата бетонирования", ""),
                        rec["Имя файла"], rec["Префикс_Колонка"]
                    ])
                ws_un.append_rows(new_rows)

            task.log(f"[✓] Google Таблица успешно обновлена!")
        except Exception as e:
            task.log(f"[ОШИБКА] Не удалось обновить Google Таблицу: {str(e)}")


# ---------------------------------------------------------------------------
# Основная обработка PDF (process_pdfs)
# ---------------------------------------------------------------------------

def process_pdfs(task: TaskState, file_paths: list, project_output_dir: str, settings: dict):
    """Полная обработка PDF-файлов (перенос PDFSplitterApp.process_pdfs)."""
    temp_dir = None
    http_session = _make_http_session()
    try:
        temp_dir = os.path.join(project_output_dir, "Сборка_протоколов_temp")
        os.makedirs(temp_dir, exist_ok=True)

        all_records = []
        files_to_zip = []
        total_files = len(file_paths)
        task.set_progress(0)

        pattern = re.compile(
            r"ПРОТОКОЛ\s*(?:№|N|Ne|N2)?\s*([A-Za-zА-Яа-я0-9\-_/\\]+)(?:\s*от\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}))?", re.IGNORECASE)
        pattern_building = re.compile(
            r"(корпус|кран|секци[яи]|парковка|паркинг|лс[- ]?)\s*([0-9]+(?:\.[0-9]+)*)", re.IGNORECASE)
        pattern_cdate = re.compile(
            r"(?:Дата\s*бетонирования|забетонирован\w*|Дата\s*изготовления)[\s\S]{0,200}?([0-9]{2}\.[0-9]{2}\.[0-9]{2,4})", re.IGNORECASE)
        pattern_struct = re.compile(
            r"(?:Наименование\s+объекта\s+испытаний|Наименование\s+конструкции|Конструкция|Конструкции|Элемент)[\s\:\.]+(.{5,150}?)(?:Дата|Объем|Класс|Требуемая|Поставщик|Фактическая|Сведения|Организация|Марка|Заказчик|Подрядчик|Возраст|Методики)", re.IGNORECASE)
        pattern_age = re.compile(
            r"(?:в\s+возрасте|возраст\s*бетона|проектный\s+возраст|проектном\s+возрасте|возраст)\s*[\:\–\-]?\s*(\d+)\s*сут", re.IGNORECASE)

        def calculate_age_from_dates(p_date, p_cdate):
            if not p_date or not p_cdate:
                return None
            try:
                d1 = datetime.datetime.strptime(p_date, "%d.%m.%Y")
                d2 = datetime.datetime.strptime(p_cdate, "%d.%m.%y" if len(p_cdate.split('.')[-1]) == 2 else "%d.%m.%Y")
                diff = (d1 - d2).days
                if 4 <= diff <= 16:
                    return "7"
                elif 20 <= diff <= 45:
                    return "28"
                return None
            except Exception:
                return None

        for file_idx, file_path in enumerate(file_paths):
            try:
                task.log(f"\n[{file_idx+1}/{total_files}] Открытие: {os.path.basename(file_path)}")
                doc = fitz.open(file_path)
                num_pages = len(doc)

                task.log(f"[*] Подготовка {num_pages} стр. к параллельному анализу...")
                pages_data = []
                for p_i in range(num_pages):
                    p_obj = doc.load_page(p_i)
                    p_data = prepare_page_data(p_obj, page_num=p_i + 1, total_pages=num_pages)
                    pages_data.append(p_data)

                workers = int(settings.get("ai_workers", 10))
                task.log(f"[*] Запуск параллельного распознавания (потоков: {workers})...")
                ai_results = {}
                completed_pages = 0

                with ThreadPoolExecutor(max_workers=workers) as executor:
                    future_to_idx = {
                        executor.submit(query_ai_for_page, p_data, num_pages, settings, http_session, task): p_data["page_num"] - 1
                        for p_data in pages_data
                    }
                    for future in as_completed(future_to_idx):
                        p_idx = future_to_idx[future]
                        completed_pages += 1
                        overall_progress = (file_idx + (completed_pages / num_pages)) / total_files
                        task.set_progress(overall_progress)
                        try:
                            res = future.result()
                            ai_results[p_idx] = res
                        except Exception as e:
                            task.log(f"  [!] Ошибка анализа стр. {p_idx + 1}: {str(e)}")
                            ai_results[p_idx] = None

                protocols = []
                current_protocol = None

                for page_num in range(num_pages):
                    page = doc.load_page(page_num)

                    p_num, p_date, p_building = None, None, None
                    p_structure, p_concreting_date, p_age, p_cube = None, None, None, False
                    is_continuation = False

                    if current_protocol and (page_num - current_protocol["start_page"] == 2):
                        protocols.append(current_protocol)
                        current_protocol = None
                        task.log(f"[*] Стр. {page_num}: принудительное завершение предыдущего протокола (лимит 2 стр.)")

                    ai_data = ai_results.get(page_num)

                    if ai_data and isinstance(ai_data, dict):
                        raw_num = ai_data.get("protocol_number")
                        p_date = ai_data.get("protocol_date")
                        p_building = ai_data.get("building")
                        if p_building:
                            p_building = str(p_building).strip()
                            m_b = pattern_building.search(p_building)
                            if m_b:
                                b_type = m_b.group(1).lower().replace('и', 'я')
                                p_building = f"{b_type} {m_b.group(2)}".strip()
                            elif p_building.isdigit():
                                p_building = f"корпус {p_building}"

                        p_structure = ai_data.get("structure")
                        p_concreting_date = ai_data.get("concreting_date")
                        if ai_data.get("age_days") is not None:
                            p_age = str(ai_data.get("age_days")).strip()
                        p_cube = bool(ai_data.get("is_cube", False))
                        is_continuation = bool(ai_data.get("is_continuation", False))

                        def _parse_d(d_val):
                            if not d_val:
                                return None
                            for fmt in ("%d.%m.%Y", "%d.%m.%y"):
                                try:
                                    return datetime.datetime.strptime(str(d_val).strip(), fmt)
                                except Exception:
                                    pass
                            return None

                        dt_p = _parse_d(p_date)
                        dt_c = _parse_d(p_concreting_date)

                        m_pref = re.search(r'(?:СВБ-)?(28|7|10|12)', str(raw_num))
                        pref_val = m_pref.group(1) if m_pref else None

                        if not p_age or p_age not in ['7', '28']:
                            if pref_val in ['7', '28']:
                                p_age = pref_val
                            elif dt_p and dt_c:
                                diff = (dt_p - dt_c).days
                                if 4 <= diff <= 14:
                                    p_age = "7"
                                elif 20 <= diff <= 35:
                                    p_age = "28"

                        days_num = int(p_age) if (p_age and p_age.isdigit()) else (28 if pref_val in ['28', '12'] else (7 if pref_val in ['7', '10'] else None))

                        if not dt_c and dt_p and days_num:
                            dt_c = dt_p - datetime.timedelta(days=days_num)
                            p_concreting_date = dt_c.strftime("%d.%m.%Y")
                        if not dt_p and dt_c and days_num:
                            dt_p = dt_c + datetime.timedelta(days=days_num)
                            p_date = dt_p.strftime("%d.%m.%Y")

                        if raw_num:
                            p_num = clean_protocol_number(
                                raw_num,
                                concreting_date=p_concreting_date,
                                protocol_date=p_date,
                                age=p_age
                            )

                        if p_num:
                            task.log(f"  [✓] Стр. {page_num + 1}: Распознан {p_num} (от {p_date or '?'}), Здание: {p_building or '?'}")
                    else:
                        text = page.get_text("text").replace('\n', ' ')
                        match = pattern.search(text)
                        match_building = pattern_building.search(text)
                        if match_building:
                            p_building = match_building.group(1).strip()
                        match_c_date = pattern_cdate.search(text)
                        if match_c_date:
                            p_concreting_date = match_c_date.group(1).strip()
                        match_struct = pattern_struct.search(text)
                        if match_struct:
                            p_structure = match_struct.group(1).strip()
                        match_age = pattern_age.search(text)
                        if match_age:
                            p_age = match_age.group(1).strip()
                        if match:
                            pfx = text[max(0, match.start() - 30):match.start()].lower()
                            if "продолжени" not in pfx:
                                p_date = match.group(2)
                                p_num = clean_protocol_number(
                                    match.group(1),
                                    concreting_date=p_concreting_date,
                                    protocol_date=p_date,
                                    age=p_age
                                )

                    if is_continuation and current_protocol:
                        current_protocol["end_page"] = page_num
                        if p_building and not current_protocol.get("building"):
                            current_protocol["building"] = p_building
                        if p_structure and not current_protocol.get("structure"):
                            current_protocol["structure"] = p_structure
                        if p_concreting_date and not current_protocol.get("concreting_date"):
                            current_protocol["concreting_date"] = p_concreting_date
                        if p_age and not current_protocol.get("age"):
                            current_protocol["age"] = p_age
                        if p_cube is not None and current_protocol.get("is_cube") is None:
                            current_protocol["is_cube"] = p_cube
                        task.log(f"  [+] Стр. {page_num + 1}: продолжение протокола {current_protocol['number']}")
                    elif p_num:
                        if current_protocol:
                            if current_protocol["number"] == p_num and not current_protocol["uncertain"]:
                                current_protocol["end_page"] = page_num
                                if p_building and not current_protocol.get("building"):
                                    current_protocol["building"] = p_building
                                if p_structure and not current_protocol.get("structure"):
                                    current_protocol["structure"] = p_structure
                                if p_concreting_date and not current_protocol.get("concreting_date"):
                                    current_protocol["concreting_date"] = p_concreting_date
                                if p_age and not current_protocol.get("age"):
                                    current_protocol["age"] = p_age
                                if p_cube is not None and current_protocol.get("is_cube") is None:
                                    current_protocol["is_cube"] = p_cube
                            else:
                                protocols.append(current_protocol)
                                current_protocol = {
                                    "number": p_num, "date": p_date, "building": p_building,
                                    "structure": p_structure, "concreting_date": p_concreting_date, "age": p_age, "is_cube": p_cube,
                                    "start_page": page_num, "end_page": page_num,
                                    "uncertain": p_date is None
                                }
                        else:
                            current_protocol = {
                                "number": p_num, "date": p_date, "building": p_building,
                                "structure": p_structure, "concreting_date": p_concreting_date, "age": p_age, "is_cube": p_cube,
                                "start_page": page_num, "end_page": page_num,
                                "uncertain": p_date is None
                            }
                    else:
                        if current_protocol:
                            current_protocol["end_page"] = page_num
                        else:
                            current_protocol = {
                                "number": "СВБ-НЕ_РАСПОЗНАН", "date": None, "building": None,
                                "structure": None, "concreting_date": None, "age": None,
                                "start_page": page_num, "end_page": page_num,
                                "uncertain": True
                            }

                if current_protocol:
                    protocols.append(current_protocol)

                if not protocols:
                    task.log(f"[!] Не найдено ни одного протокола в файле {os.path.basename(file_path)}")
                    doc.close()
                    continue

                # --- Слияние двухстраничных протоколов ---
                merged_protocols = []
                for prot in protocols:
                    if not merged_protocols:
                        merged_protocols.append(prot)
                        continue

                    prev_prot = merged_protocols[-1]
                    prev_len = prev_prot["end_page"] - prev_prot["start_page"] + 1

                    can_merge = (prev_len == 1) and (prot["start_page"] == prev_prot["end_page"] + 1)

                    if can_merge:
                        if prot["uncertain"]:
                            prev_prot["end_page"] = prot["end_page"]
                            for k in ["building", "structure", "concreting_date", "age"]:
                                if prot.get(k) and not prev_prot.get(k):
                                    prev_prot[k] = prot[k]
                            if prot.get("is_cube") is not None and prev_prot.get("is_cube") is None:
                                prev_prot["is_cube"] = prot["is_cube"]
                            continue

                        if prev_prot["uncertain"] and not prot["uncertain"]:
                            prev_prot["number"] = prot["number"]
                            prev_prot["date"] = prot["date"]
                            prev_prot["end_page"] = prot["end_page"]
                            prev_prot["uncertain"] = False
                            for k in ["building", "structure", "concreting_date", "age"]:
                                if prot.get(k) and not prev_prot.get(k):
                                    prev_prot[k] = prot[k]
                            if prot.get("is_cube") is not None and prev_prot.get("is_cube") is None:
                                prev_prot["is_cube"] = prot["is_cube"]
                            continue

                    merged_protocols.append(prot)

                protocols = merged_protocols

                # Формирование файлов
                for idx, prot in enumerate(protocols, start=1):
                    clean_num = prot["number"].replace("/", "_").replace("\\", "_")
                    date_str = prot["date"] if prot["date"] else "Без_даты"
                    clean_date = date_str.replace("/", "_").replace("\\", "_")
                    building_str = prot.get("building")
                    building_folder_name = f"Корпус {building_str}" if building_str else "Корпус_Неизвестен"

                    start_p = prot["start_page"]
                    end_p = prot["end_page"]
                    num_p = end_p - start_p + 1

                    page_range_str = f"{start_p+1}" if num_p == 1 else f"{start_p+1}-{end_p+1}"

                    m_pref = re.search(r'СВБ-(10|12|28|7)', prot["number"])
                    p_val = m_pref.group(1) if m_pref else "Неизвестно"
                    prefix_folder = "7_28" if p_val in ["7", "28"] else p_val

                    filename = f"Протокол_{clean_num}_от_{clean_date}.pdf"

                    target_dir = os.path.join(project_output_dir, building_folder_name, prefix_folder)
                    os.makedirs(target_dir, exist_ok=True)

                    target_filepath = os.path.join(target_dir, filename)

                    counter = 1
                    while os.path.exists(target_filepath):
                        filename_base = f"Протокол_{clean_num}_от_{clean_date}_{counter}.pdf"
                        target_filepath = os.path.join(target_dir, filename_base)
                        filename = filename_base
                        counter += 1

                    new_doc = fitz.open()
                    new_doc.insert_pdf(doc, from_page=start_p, to_page=end_p)
                    os.makedirs(os.path.dirname(target_filepath), exist_ok=True)
                    new_doc.save(target_filepath, garbage=3, deflate=True)
                    new_doc.close()
                    files_to_zip.append(target_filepath)

                    note = "уверенно" if not prot["uncertain"] else "требует ручной проверки"

                    final_age = str(prot.get("age")).strip() if prot.get("age") else None
                    if final_age not in ["7", "28"]:
                        if final_age and "28" in final_age:
                            final_age = "28"
                        elif final_age and "7" in final_age:
                            final_age = "7"
                        else:
                            final_age = None

                    if not final_age:
                        final_age = calculate_age_from_dates(prot["date"], prot.get("concreting_date"))

                    if not final_age and prot.get("date") and prot.get("number"):
                        m_num_date = re.search(r'[-_](\d{2})(\d{2})_(\d{2})', prot["number"])
                        if m_num_date:
                            dd, mm, yy = m_num_date.groups()
                            final_age = calculate_age_from_dates(prot["date"], f"{dd}.{mm}.20{yy}")

                    protocol_prefix_match = re.search(r'СВБ-(10|12|28|7)[\-_\\]', prot["number"])
                    if protocol_prefix_match:
                        p_prefix_val = protocol_prefix_match.group(1)
                        if p_prefix_val in ("10", "12"):
                            final_is_cube = False
                            if p_prefix_val == "12" and not final_age:
                                final_age = "28"
                        else:
                            final_is_cube = True
                            if p_prefix_val == "28":
                                final_age = "28"
                            elif p_prefix_val == "7":
                                final_age = "7"
                    else:
                        final_is_cube = bool(prot.get("is_cube", False))

                    if not final_age:
                        final_age = "28"

                    all_records.append({
                        "Исходный файл": os.path.basename(file_path),
                        "Номер протокола": prot["number"],
                        "Дата протокола": prot["date"] if prot["date"] else "Не найдена",
                        "Корпус": building_str if building_str else "Не найден",
                        "Наименование конструкции": prot.get("structure"),
                        "Дата бетонирования": prot.get("concreting_date"),
                        "Возраст суток": final_age,
                        "Страницы в исходном PDF": page_range_str,
                        "Количество страниц": num_p,
                        "Имя файла": filename,
                        "Примечание": note,
                        "Префикс_Колонка": prefix_folder,
                        "Кубики": final_is_cube
                    })

                doc.close()

            except Exception as ex:
                task.log(f"[ОШИБКА] Ошибка при обработке {os.path.basename(file_path)}: {str(ex)}")

        # --- Завершение ---
        result_files = []
        if all_records:
            final_excel_path = os.path.join(project_output_dir, "Реестр_разбивки.xlsx")
            create_formatted_registry(all_records, final_excel_path)
            files_to_zip.append(final_excel_path)
            result_files.append(os.path.basename(final_excel_path))

            perform_integration(all_records, project_output_dir, settings, task)

            zip_filename = os.path.join(project_output_dir, "Все_протоколы_архив.zip")
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file in files_to_zip:
                    zipf.write(file, os.path.basename(file))
            result_files.append(os.path.basename(zip_filename))

            task.set_progress(1.0)
            task.log(
                f"\n[✓] Готово! \nВсего обработано файлов: {total_files}\nФайлы распакованы в папки по корпусам: {project_output_dir}\nТакже создан ZIP-архив: {zip_filename}\nПромежуточный реестр сохранен по шаблону: {final_excel_path}")
        else:
            task.log("\n[!] Нет данных для сохранения (ни один протокол не найден).")

        task.finish(all_records, result_files)

    except Exception as e:
        task.log(f"\n[КРИТИЧЕСКАЯ ОШИБКА] {str(e)}")
        task.finish([], [])
    finally:
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Интеграция-only обработка
# ---------------------------------------------------------------------------

def integrate_only_worker(task: TaskState, registry_xlsx_path: str, settings: dict):
    """Запуск интеграции из готового Реестр_разбивки.xlsx."""
    try:
        df = pd.read_excel(registry_xlsx_path)
        df = df.fillna("")

        col_names = [str(c).lower() for c in df.columns]
        has_template_cols = any("неразрушающ" in c or "кубик" in c or "куб" in c for c in col_names)

        if has_template_cols:
            all_records = []
            for _, row in df.iterrows():
                p_cdate = str(row.get("Дата бетонирования", "")).strip()
                if p_cdate.lower() == "nan": p_cdate = ""
                p_bldg = str(row.get("Корпус", "")).strip()
                if p_bldg.lower() == "nan": p_bldg = ""
                p_strc = str(row.get("Наименование конструкции", "")).strip()
                if p_strc.lower() == "nan": p_strc = ""
                filename = str(row.get("Имя файла", "")).strip()
                if filename.lower() == "nan": filename = ""

                for col_name in df.columns:
                    c_lower = str(col_name).lower()
                    val = str(row.get(col_name, "")).strip()
                    if not val or val.lower() == "nan":
                        continue

                    is_7_nz = ("неразрушающ" in c_lower and "7" in c_lower)
                    is_28_nz = ("неразрушающ" in c_lower and "28" in c_lower)
                    is_7_cub = (("кубик" in c_lower or "куб" in c_lower) and "7" in c_lower)
                    is_28_cub = (("кубик" in c_lower or "куб" in c_lower) and "28" in c_lower)

                    if is_7_nz or is_28_nz or is_7_cub or is_28_cub:
                        m = re.search(r'№?\s*([A-Za-zА-Яа-я0-9\-_/\\]+)(?:\s*от\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}))?', val)
                        if m:
                            p_num = m.group(1).strip()
                            p_date = m.group(2) if m.group(2) else ""
                        else:
                            p_num = val.replace("№", "").strip()
                            p_date = ""

                        age = "7" if (is_7_nz or is_7_cub) else "28"
                        is_cube = (is_7_cub or is_28_cub)

                        all_records.append({
                            "Номер протокола": p_num,
                            "Дата протокола": p_date,
                            "Корпус": p_bldg,
                            "Наименование конструкции": p_strc,
                            "Дата бетонирования": p_cdate,
                            "Возраст суток": age,
                            "Имя файла": filename,
                            "Префикс_Колонка": "7_28" if age in ["7", "28"] else "10",
                            "Кубики": is_cube,
                            "uncertain": False
                        })
        else:
            all_records = df.to_dict('records')

        project_output_dir = os.path.dirname(registry_xlsx_path)
        task.log(f"=== Начало интеграции из готового файла ({len(all_records)} протоколов) ===")
        perform_integration(all_records, project_output_dir, settings, task)
        task.log("=== Интеграция завершена ===")
        task.finish(all_records, [])
    except Exception as e:
        task.log(f"[ОШИБКА] Не удалось прочитать файл: {str(e)}")
        task.finish([], [])


# ===========================================================================
# API ENDPOINTS
# ===========================================================================

@app.route("/")
def index():
    """Отдаёт index.html из папки templates."""
    return render_template("index.html")


@app.route("/api/settings", methods=["GET"])
def api_get_settings():
    """Возвращает текущие настройки."""
    s = load_settings()
    s["output_dir"] = OUTPUT_DIR
    return jsonify(s)


@app.route("/api/settings", methods=["POST"])
def api_save_settings():
    """Сохраняет настройки."""
    data = request.get_json(force=True)
    current = load_settings()
    for key in ["theme", "ai_api_key", "ai_base_url", "ai_model", "ai_workers",
                 "registry_mode", "google_url", "google_creds_path", "registry_path"]:
        if key in data:
            current[key] = data[key]
    save_settings(current)
    return jsonify({"ok": True})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    """Загрузка PDF-файлов (multipart). Сохраняет в uploads/, возвращает список имён."""
    if "files" not in request.files:
        return jsonify({"error": "No files part"}), 400

    uploaded = request.files.getlist("files")
    names = []
    for f in uploaded:
        if f.filename:
            safe_name = f.filename
            dest = os.path.join(UPLOAD_DIR, safe_name)
            f.save(dest)
            names.append(safe_name)
    return jsonify({"files": names})


@app.route("/api/upload-registry", methods=["POST"])
def api_upload_registry():
    """Загрузка Excel-реестра для интеграции."""
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400
    dest = os.path.join(UPLOAD_DIR, f.filename)
    f.save(dest)
    # Сохраняем путь в настройки
    settings = load_settings()
    settings["registry_path"] = dest
    save_settings(settings)
    return jsonify({"ok": True, "path": dest, "filename": f.filename})


@app.route("/api/upload-creds", methods=["POST"])
def api_upload_creds():
    """Загрузка credentials.json для Google Sheets."""
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400
    dest = os.path.join(BASE_DIR, "credentials.json")
    f.save(dest)
    settings = load_settings()
    settings["google_creds_path"] = dest
    save_settings(settings)
    return jsonify({"ok": True, "path": dest})


@app.route("/api/process", methods=["POST"])
def api_process():
    """Запускает обработку загруженных PDF в фоновом потоке, возвращает {task_id}."""
    data = request.get_json(force=True) if request.is_json else {}
    file_names = data.get("files", [])

    # Если явный список не передан — берём всё из uploads/
    if not file_names:
        file_names = [f for f in os.listdir(UPLOAD_DIR) if f.lower().endswith(".pdf")]

    if not file_names:
        return jsonify({"error": "No PDF files to process"}), 400

    file_paths = [os.path.join(UPLOAD_DIR, fn) for fn in file_names]
    missing = [fp for fp in file_paths if not os.path.exists(fp)]
    if missing:
        return jsonify({"error": f"Files not found: {missing}"}), 404

    settings = load_settings()
    if "registry_mode" in data and data["registry_mode"]:
        settings["registry_mode"] = data["registry_mode"]
    if "google_url" in data and data["google_url"]:
        settings["google_url"] = data["google_url"]

    task_id = str(uuid.uuid4())
    task = TaskState(task_id)

    with _tasks_lock:
        _tasks[task_id] = task

    thread = threading.Thread(target=process_pdfs, args=(task, file_paths, OUTPUT_DIR, settings), daemon=True)
    thread.start()

    return jsonify({"task_id": task_id})


@app.route("/api/integrate", methods=["POST"])
def api_integrate():
    """Загружает Excel-файл Реестр_разбивки.xlsx и запускает только интеграцию."""
    # Файл может быть передан как multipart или уже лежать в output/
    registry_path = None

    if "file" in request.files:
        f = request.files["file"]
        if f.filename:
            dest = os.path.join(OUTPUT_DIR, f.filename)
            f.save(dest)
            registry_path = dest

    if not registry_path:
        # Пробуем найти в output/
        candidate = os.path.join(OUTPUT_DIR, "Реестр_разбивки.xlsx")
        if os.path.exists(candidate):
            registry_path = candidate

    if not registry_path or not os.path.exists(registry_path):
        return jsonify({"error": "Registry file not found. Upload it or run /api/process first."}), 400

    settings = load_settings()
    if "registry_mode" in data and data["registry_mode"]:
        settings["registry_mode"] = data["registry_mode"]
    if "google_url" in data and data["google_url"]:
        settings["google_url"] = data["google_url"]

    task_id = str(uuid.uuid4())
    task = TaskState(task_id)

    with _tasks_lock:
        _tasks[task_id] = task

    thread = threading.Thread(target=integrate_only_worker, args=(task, registry_path, settings), daemon=True)
    thread.start()

    return jsonify({"task_id": task_id})


@app.route("/api/status/<task_id>")
def api_status(task_id):
    """SSE стрим логов и прогресса."""
    with _tasks_lock:
        task = _tasks.get(task_id)
    if not task:
        return jsonify({"error": "Task not found"}), 404

    def generate():
        while True:
            try:
                msg = task.queue.get(timeout=30)
                yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
                if msg.get("type") == "done":
                    break
            except queue.Empty:
                # keep-alive
                yield f"data: {json.dumps({'type': 'ping'})}\n\n"
                if task.done:
                    break

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        }
    )


@app.route("/api/download/<filename>")
def api_download(filename):
    """Скачивание результата (ZIP-архив, Excel) из output/."""
    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404
    return send_file(filepath, as_attachment=True)


# ===========================================================================
# Запуск
# ===========================================================================

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
