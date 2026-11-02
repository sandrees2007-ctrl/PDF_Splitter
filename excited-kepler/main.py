import os
import json
import sys
import subprocess
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import zipfile
import re
import shutil
import difflib
import time
import base64
import datetime
import math
from tkinter import Canvas, messagebox, filedialog

try:
    import customtkinter as ctk
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
    import customtkinter as ctk

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
import requests
from requests.adapters import HTTPAdapter

DEFAULT_PROXYAPI_KEY = "sk-W02LrjSK9PapzakiwgjkP0IBkFN7p3rE"
DEFAULT_AI_BASE_URL = "https://api.proxyapi.ru/v1"
DEFAULT_AI_MODEL = "google/gemma-4-26b-a4b-it"


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


class LiquidGlassProgressBar(ctk.CTkFrame):
    """Стеклянный прогресс-бар с эффектом светящейся жидкости и бегущей световой волны."""
    def __init__(self, master, height=20, **kwargs):
        super().__init__(master, height=height, corner_radius=12,
                         fg_color="#080D18", border_width=1, border_color="#1C273C", **kwargs)
        self.height = height
        self.target_progress = 0.0
        self.curr_progress = 0.0
        self.anim_tick = 0
        self.accent_rgb1 = (0, 229, 255)   # #00E5FF
        self.accent_rgb2 = (79, 172, 254)  # #4FACFE
        self.glow_border = "#00E5FF"
        self.pct_label = None

        self.canvas = Canvas(self, height=self.height, bg="#080D18", highlightthickness=0)
        self.canvas.pack(expand=True, fill="both", padx=2, pady=2)
        self.canvas.bind("<Configure>", lambda e: self._render_liquid())
        self._animate()

    def set(self, value):
        self.target_progress = max(0.0, min(1.0, float(value)))
        if self.pct_label:
            pct = int(self.target_progress * 100)
            try:
                self.after(0, lambda: self.pct_label.configure(text=f"{pct}%"))
            except Exception:
                pass

    def update_theme(self, rgb1, rgb2, border_col):
        self.accent_rgb1 = rgb1
        self.accent_rgb2 = rgb2
        self.glow_border = border_col
        self._render_liquid()

    def _animate(self):
        self.anim_tick = (self.anim_tick + 1) % 720
        diff = self.target_progress - self.curr_progress
        if abs(diff) > 0.002:
            self.curr_progress += diff * 0.18
        else:
            self.curr_progress = self.target_progress
        self._render_liquid()
        self.after(40, self._animate)

    def _render_liquid(self):
        self.canvas.delete("all")
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 4 or h <= 4:
            return

        # 1. Стеклянный фон трубки (Glass groove)
        self.canvas.create_rectangle(0, 0, w, h, fill="#070B14", width=0)
        self.canvas.create_line(0, 0, w, 0, fill="#182335")
        self.canvas.create_line(0, h-1, w, h-1, fill="#0C131F")

        fill_w = int(w * self.curr_progress)
        shimmer_pos = (self.anim_tick * 8) % (w + 140) - 70

        if fill_w > 0:
            for x in range(fill_w):
                t = x / max(1, w)
                r = int(self.accent_rgb1[0] + (self.accent_rgb2[0] - self.accent_rgb1[0]) * t)
                g = int(self.accent_rgb1[1] + (self.accent_rgb2[1] - self.accent_rgb1[1]) * t)
                b = int(self.accent_rgb1[2] + (self.accent_rgb2[2] - self.accent_rgb1[2]) * t)

                # Бегущий световой блик (Liquid glare shimmer)
                dist = abs(x - shimmer_pos)
                if dist < 45:
                    boost = (1.0 - (dist / 45.0)) * 0.75
                    r = min(255, int(r + (255 - r) * boost))
                    g = min(255, int(g + (255 - g) * boost))
                    b = min(255, int(b + (255 - b) * boost))

                color = f"#{r:02x}{g:02x}{b:02x}"
                self.canvas.create_line(x, 1, x, h-1, fill=color)

            # Верхний зеркальный блик стекла (Specular meniscus highlight)
            self.canvas.create_line(0, 1, fill_w, 1, fill="#E0FAFF")
            # Правая светящаяся головка жидкости
            if fill_w < w:
                self.canvas.create_line(fill_w-1, 1, fill_w-1, h-1, fill="#FFFFFF")

        # Мягкое фоновое дыхание незаполненного стекла
        if fill_w < w:
            pulse_a = (math.sin(self.anim_tick * 0.08) + 1.0) * 0.5
            glow_c = f"#{int(16 + 10*pulse_a):02x}{int(24 + 14*pulse_a):02x}{int(36 + 22*pulse_a):02x}"
            self.canvas.create_line(max(fill_w, 0), 1, w, 1, fill=glow_c)


class LiquidGlassHeader(ctk.CTkFrame):
    """Стеклянный хедер с эффектом преломления света и плавающими жидкими неоновыми сферами."""
    def __init__(self, master, height=84, on_settings_click=None, **kwargs):
        super().__init__(master, height=height, corner_radius=16,
                         fg_color="#0E1422", border_width=1, border_color="#1E2A3E", **kwargs)
        self.height = height
        self.anim_tick = 0
        self.accent_hex = "#00E5FF"
        self.status_text = "СИСТЕМА ГОТОВА К РАБОТЕ"
        self.status_active = False
        self.on_settings_click = on_settings_click

        self.canvas = Canvas(self, height=self.height, bg="#0E1422", highlightthickness=0)
        self.canvas.pack(expand=True, fill="both", padx=2, pady=2)
        self.canvas.bind("<Configure>", lambda e: self._render_glass())

        # Кнопка настроек в правом углу хедера (всегда поверх канваса)
        self.btn_settings = ctk.CTkButton(
            self, text="⚙ Настройки", command=self.on_settings_click,
            fg_color="#151F33", hover_color="#202D45",
            border_width=1, border_color="#2A3B56",
            text_color="#F0F6FC",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            corner_radius=10, width=110, height=32
        )
        self.btn_settings.place(relx=0.98, rely=0.5, anchor="e")
        self.btn_settings.lift()

        self._animate()

    def set_status(self, text, active=False):
        self.status_text = text
        self.status_active = active
        self._render_glass()

    def update_theme(self, accent_hex, border_hex):
        self.accent_hex = accent_hex
        self.configure(border_color=border_hex)
        self._render_glass()

    def _animate(self):
        self.anim_tick = (self.anim_tick + 1) % 720
        self._render_glass()
        self.after(50, self._animate)

    def _render_glass(self):
        self.canvas.delete("glass_bg")
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 10 or h <= 10:
            return

        # 1. Плавающие неоновые сферы жидкого стекла (Ambient liquid glow orbs)
        t = self.anim_tick * 0.035
        # Левая сфера (акцентный неон)
        ox1 = w * 0.22 + math.sin(t) * 40
        oy1 = h * 0.5 + math.cos(t * 0.7) * 14
        for r_step, col in [(95, "#101B2C"), (65, "#132338"), (40, "#162C47"), (18, "#1A375A")]:
            self.canvas.create_oval(ox1 - r_step, oy1 - r_step * 0.6, ox1 + r_step, oy1 + r_step * 0.6,
                                   fill=col, width=0, tags="glass_bg")

        # Правая сфера (лазурно-индиговый свет)
        ox2 = w * 0.68 + math.cos(t * 0.6) * 45
        oy2 = h * 0.5 + math.sin(t * 0.8) * 16
        for r_step, col in [(100, "#181832"), (70, "#1E1E42"), (45, "#252452"), (20, "#2F2C65")]:
            self.canvas.create_oval(ox2 - r_step, oy2 - r_step * 0.6, ox2 + r_step, oy2 + r_step * 0.6,
                                   fill=col, width=0, tags="glass_bg")

        # 2. Верхний зеркальный срез стекла (Specular reflection line)
        self.canvas.create_line(15, 2, w - 15, 2, fill="#2C3D57", width=1, tags="glass_bg")
        self.canvas.create_line(w * 0.12, 2, w * 0.45, 2, fill="#6497D2", width=1, tags="glass_bg")

        # 3. Неоновая стеклянная эмблема (Shield badge)
        gx, gy = 36, h // 2
        self.canvas.create_polygon(gx, gy-17, gx+15, gy-8, gx+15, gy+8, gx, gy+17, gx-15, gy+8, gx-15, gy-8,
                                   fill="#12243A", outline=self.accent_hex, width=2, tags="glass_bg")
        self.canvas.create_text(gx, gy, text="🧪", font=("Segoe UI", 12), tags="glass_bg")

        # 4. Типографика с эффектом стеклянного свечения
        self.canvas.create_text(66, gy - 12, text="STROYLAB INTELLIGENCE SUITE • LIQUID GLASS v3.0",
                                anchor="w", fill=self.accent_hex, font=("Segoe UI", 9, "bold"), tags="glass_bg")
        self.canvas.create_text(66, gy + 10, text="Умная разбивка & Реестр протоколов",
                                anchor="w", fill="#F0F6FC", font=("Segoe UI", 16, "bold"), tags="glass_bg")

        # 5. Статусный индикатор с пульсирующим неоновым диодом
        pw, ph = 240, 32
        px = w - 130 - pw - 20
        py = gy - ph // 2
        if px > 380:
            self.canvas.create_rectangle(px, py, px + pw, py + ph, fill="#0B121E", outline="#1E2C42", width=1, tags="glass_bg")
            self.canvas.create_line(px + 4, py + 1, px + pw - 4, py + 1, fill="#2E4464", tags="glass_bg")

            pulse = (math.sin(self.anim_tick * 0.15) + 1.0) * 0.5
            dot_col = "#00F5A0" if not self.status_active else "#FFB800"
            halo_col = "#06462D" if not self.status_active else "#4F3600"
            dx, dy = px + 16, gy
            self.canvas.create_oval(dx - 6*pulse, dy - 6*pulse, dx + 6*pulse, dy + 6*pulse, fill=halo_col, width=0, tags="glass_bg")
            self.canvas.create_oval(dx - 4, dy - 4, dx + 4, dy + 4, fill=dot_col, width=0, tags="glass_bg")

            self.canvas.create_text(px + 28, gy, text=self.status_text, anchor="w",
                                    fill="#E2E8F0", font=("Segoe UI", 9, "bold"), tags="glass_bg")


class PDFSplitterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Настройка современного стиля Liquid Glass
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Умная разбивка PDF-протоколов • Liquid Glass Suite")
        self.geometry("960x860")
        self.minsize(900, 800)

        # Палитра Liquid Glass
        self.themes = {
            "Liquid Cyan": {
                "bg_root": "#080C14", "bg_frame": "#101726",
                "border_glass": "#1C273C", "border_glow": "#00E5FF",
                "accent_col": "#00E5FF", "accent_hover": "#33EEFF",
                "accent_rgb1": (0, 229, 255), "accent_rgb2": (79, 172, 254),
                "text_col": "#F0F6FC", "text_muted": "#8B9BB4",
                "btn_bg": "#151F33", "btn_hover": "#1E2B44", "btn_border": "#283955"
            },
            "Liquid Emerald": {
                "bg_root": "#060F0C", "bg_frame": "#0D1D17",
                "border_glass": "#18372C", "border_glow": "#00F5A0",
                "accent_col": "#00F5A0", "accent_hover": "#33F8B5",
                "accent_rgb1": (0, 245, 160), "accent_rgb2": (0, 210, 255),
                "text_col": "#F0FDF4", "text_muted": "#80A495",
                "btn_bg": "#122A21", "btn_hover": "#1B3B2F", "btn_border": "#225241"
            },
            "Liquid Amethyst": {
                "bg_root": "#0B0917", "bg_frame": "#151226",
                "border_glass": "#282046", "border_glow": "#B388FF",
                "accent_col": "#B388FF", "accent_hover": "#C7A6FF",
                "accent_rgb1": (179, 136, 255), "accent_rgb2": (124, 77, 255),
                "text_col": "#FAF7FF", "text_muted": "#9A8EB8",
                "btn_bg": "#1F1A38", "btn_hover": "#2B244D", "btn_border": "#3E346B"
            },
            "Liquid Sunset": {
                "bg_root": "#0F0A07", "bg_frame": "#1C130E",
                "border_glass": "#392319", "border_glow": "#FF8C00",
                "accent_col": "#FF8C00", "accent_hover": "#FFA333",
                "accent_rgb1": (255, 140, 0), "accent_rgb2": (255, 61, 0),
                "text_col": "#FFF9F5", "text_muted": "#A89083",
                "btn_bg": "#2B1C14", "btn_hover": "#3D281D", "btn_border": "#5A3828"
            }
        }
        self.current_theme_name = "Liquid Cyan"

        self.file_paths = []
        self.output_dir = None
        self.registry_path = None
        self.google_creds_path = None
        self.is_processing = False
        self.saved_google_url = ""
        self.registry_mode_val = "local"
        self.ai_workers = 10

        self.http_session = requests.Session()
        adapter = HTTPAdapter(pool_connections=40, pool_maxsize=40, max_retries=1)
        self.http_session.mount("https://", adapter)
        self.http_session.mount("http://", adapter)

        self.load_settings()
        self.apply_theme_colors()
        self.configure(fg_color=self.bg_root)

        self.setup_ui()
        self.apply_saved_settings_to_ui()

    def setup_ui(self):
        # 1. АНИМИРОВАННЫЙ СТЕКЛЯННЫЙ ХЕДЕР С ПЛАВАЮЩИМИ НЕОНОВЫМИ СФЕРАМИ
        self.header = LiquidGlassHeader(self, height=84, on_settings_click=self.open_settings)
        self.header.pack(fill="x", padx=30, pady=(20, 15))

        # 2. СТЕКЛЯННАЯ КАРТОЧКА 1: Выбор файлов
        self.frame_input = ctk.CTkFrame(
            self, fg_color=self.bg_frame, border_width=1, border_color=self.border_glass, corner_radius=14)
        self.frame_input.pack(fill="x", padx=30, pady=(0, 12))

        self.btn_select_file = ctk.CTkButton(
            self.frame_input, text="📁  1. Выбрать PDF файлы", command=self.select_file,
            fg_color=self.btn_bg, hover_color=self.btn_hover,
            border_width=1, border_color=self.btn_border,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            corner_radius=10, width=210, height=42
        )
        self.btn_select_file.pack(side="left", padx=16, pady=16)

        info_input = ctk.CTkFrame(self.frame_input, fg_color="transparent")
        info_input.pack(side="left", fill="both", expand=True, padx=(0, 16), pady=12)

        self.lbl_file_badge = ctk.CTkLabel(
            info_input, text="● Ожидание выбора файлов",
            text_color=self.text_muted, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.lbl_file_badge.pack(anchor="w")

        self.lbl_file = ctk.CTkLabel(
            info_input, text="Перетащите или укажите PDF файлы сканов строительных протоколов",
            text_color="#64748B", font=ctk.CTkFont(family="Segoe UI", size=12),
            wraplength=600, justify="left"
        )
        self.lbl_file.pack(anchor="w", pady=(2, 0))

        # 3. СТЕКЛЯННАЯ КАРТОЧКА 2: Папка сохранения
        self.frame_output = ctk.CTkFrame(
            self, fg_color=self.bg_frame, border_width=1, border_color=self.border_glass, corner_radius=14)
        self.frame_output.pack(fill="x", padx=30, pady=(0, 12))

        self.btn_select_dir = ctk.CTkButton(
            self.frame_output, text="📂  2. Папка сохранения", command=self.select_output_dir,
            fg_color=self.btn_bg, hover_color=self.btn_hover,
            border_width=1, border_color=self.btn_border,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            corner_radius=10, width=210, height=42
        )
        self.btn_select_dir.pack(side="left", padx=16, pady=16)

        info_output = ctk.CTkFrame(self.frame_output, fg_color="transparent")
        info_output.pack(side="left", fill="both", expand=True, padx=(0, 16), pady=12)

        self.lbl_dir_badge = ctk.CTkLabel(
            info_output, text="● Структурированный каталог вывода",
            text_color=self.text_muted, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.lbl_dir_badge.pack(anchor="w")

        self.lbl_dir = ctk.CTkLabel(
            info_output, text="Рядом с исходными файлами (автоматическая разбивка по корпусам и префиксам)",
            text_color="#64748B", font=ctk.CTkFont(family="Segoe UI", size=12),
            wraplength=600, justify="left"
        )
        self.lbl_dir.pack(anchor="w", pady=(2, 0))

        # 4. СТЕКЛЯННАЯ КАРТОЧКА 3: Выбор реестра (Локальный Excel / Google)
        self.frame_registry = ctk.CTkFrame(
            self, fg_color=self.bg_frame, border_width=1, border_color=self.border_glass, corner_radius=14)
        self.frame_registry.pack(fill="x", padx=30, pady=(0, 15))

        top_reg = ctk.CTkFrame(self.frame_registry, fg_color="transparent")
        top_reg.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(
            top_reg, text="📊  3. Целевой реестр бетонирования:",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.text_col
        ).pack(side="left")

        self.registry_mode = ctk.StringVar(value="local")

        self.rb_local = ctk.CTkRadioButton(
            top_reg, text="Локальный Excel (.xlsx)",
            variable=self.registry_mode, value="local", command=self.toggle_registry_mode,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=self.accent_col, hover_color=self.accent_hover
        )
        self.rb_local.pack(side="left", padx=(25, 15))

        self.rb_google = ctk.CTkRadioButton(
            top_reg, text="Google Таблица (Облако)",
            variable=self.registry_mode, value="google", command=self.toggle_registry_mode,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=self.accent_col, hover_color=self.accent_hover
        )
        self.rb_google.pack(side="left", padx=10)

        # Контейнер для локального файла Excel
        self.frame_local = ctk.CTkFrame(self.frame_registry, fg_color="transparent")
        self.frame_local.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_select_registry = ctk.CTkButton(
            self.frame_local, text="Выбрать Excel файл", command=self.select_registry,
            fg_color=self.btn_bg, hover_color=self.btn_hover,
            border_width=1, border_color=self.btn_border,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            corner_radius=8, width=170, height=36
        )
        self.btn_select_registry.pack(side="left", padx=(0, 12))

        self.lbl_registry = ctk.CTkLabel(
            self.frame_local, text="Файл реестра не выбран (запись в существующий реестр отключена)",
            text_color="#64748B", font=ctk.CTkFont(family="Segoe UI", size=12)
        )
        self.lbl_registry.pack(side="left", fill="x", expand=True)

        # Контейнер для Google Sheets
        self.frame_google = ctk.CTkFrame(self.frame_registry, fg_color="transparent")

        self.entry_google_url = ctk.CTkEntry(
            self.frame_google, placeholder_text="Вставьте ссылку на Google Таблицу...",
            fg_color="#0A101D", border_color=self.border_glass, text_color=self.text_col,
            corner_radius=8, height=36
        )
        self.entry_google_url.pack(side="left", padx=(0, 10), fill="x", expand=True)

        self.btn_creds = ctk.CTkButton(
            self.frame_google, text="Ключ credentials.json", command=self.select_creds,
            fg_color=self.btn_bg, hover_color=self.btn_hover,
            border_width=1, border_color=self.btn_border,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            corner_radius=8, width=190, height=36
        )
        self.btn_creds.pack(side="left")

        # 5. КНОПКИ ДЕЙСТВИЯ (Liquid Glass Action Buttons)
        btns_row = ctk.CTkFrame(self, fg_color="transparent")
        btns_row.pack(fill="x", padx=30, pady=(0, 10))

        self.btn_start = ctk.CTkButton(
            btns_row, text="✨  НАЧАТЬ ОБРАБОТКУ  (PDF ➔ Excel ➔ Google Таблица)",
            command=self.start_processing,
            fg_color=self.btn_bg, hover_color=self.btn_hover,
            text_color=self.text_muted,
            border_width=1, border_color=self.border_glass,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            corner_radius=12, height=50, state="disabled"
        )
        self.btn_start.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_integrate_only = ctk.CTkButton(
            btns_row, text="⚡  ТОЛЬКО ИНТЕГРАЦИЯ",
            command=self.integrate_only,
            fg_color="#121D2F", hover_color="#1A2A44",
            text_color=self.accent_col,
            border_width=1, border_color=self.border_glass,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            corner_radius=12, width=220, height=50
        )
        self.btn_integrate_only.pack(side="right")

        # 6. ЖИДКИЙ ПРОГРЕСС-БАР С НЕОНОВОЙ ВОЛНОЙ
        prog_header = ctk.CTkFrame(self, fg_color="transparent")
        prog_header.pack(fill="x", padx=30, pady=(4, 4))

        self.lbl_progress_status = ctk.CTkLabel(
            prog_header, text="Прогресс анализа и классификации:",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.text_muted
        )
        self.lbl_progress_status.pack(side="left")

        self.lbl_progress_pct = ctk.CTkLabel(
            prog_header, text="0%",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=self.accent_col
        )
        self.lbl_progress_pct.pack(side="right")

        self.progress_bar = LiquidGlassProgressBar(self, height=18)
        self.progress_bar.pct_label = self.lbl_progress_pct
        self.progress_bar.update_theme(
            self.themes[self.current_theme_name]["accent_rgb1"],
            self.themes[self.current_theme_name]["accent_rgb2"],
            self.accent_col
        )
        self.progress_bar.pack(fill="x", padx=30, pady=(0, 14))

        # 7. ТЕРМИНАЛ ИИ-ЛОГОВ (Liquid Glass HUD Console)
        self.frame_terminal = ctk.CTkFrame(
            self, fg_color=self.bg_frame, border_width=1, border_color=self.border_glass, corner_radius=14)
        self.frame_terminal.pack(fill="both", expand=True, padx=30, pady=(0, 24))

        # Верхняя панель терминала с цветными кнопками окна
        term_bar = ctk.CTkFrame(self.frame_terminal, fg_color="#090E18", height=32, corner_radius=10)
        term_bar.pack(fill="x", padx=2, pady=2)

        dots_box = ctk.CTkFrame(term_bar, fg_color="transparent")
        dots_box.pack(side="left", padx=12)

        for col in ["#FF5F56", "#FFBD2E", "#27C93F"]:
            d = ctk.CTkFrame(dots_box, width=10, height=10, corner_radius=5, fg_color=col)
            d.pack(side="left", padx=3)

        ctk.CTkLabel(
            term_bar, text="✦ ТЕРМИНАЛ ИИ-РАСПОЗНАВАНИЯ & ВАЛИДАЦИИ",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.accent_col
        ).pack(side="left", padx=10)

        self.lbl_term_info = ctk.CTkLabel(
            term_bar, text=f"ProxyAPI • {self.ai_model} • 150 DPI",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color="#64748B"
        )
        self.lbl_term_info.pack(side="right", padx=14)

        self.log_area = ctk.CTkTextbox(
            self.frame_terminal, fg_color="#060A12", text_color=self.text_col,
            font=ctk.CTkFont(family="Consolas", size=12), corner_radius=10,
            border_width=0
        )
        self.log_area.pack(fill="both", expand=True, padx=8, pady=(4, 8))

    def select_file(self):
        filepaths = filedialog.askopenfilenames(
            title="Выберите PDF файлы", filetypes=[("PDF Files", "*.pdf")])
        if filepaths:
            self.file_paths = list(filepaths)
            self.lbl_file_badge.configure(
                text=f"✨ Готово к обработке: {len(self.file_paths)} PDF", text_color=self.accent_col)
            first_name = os.path.basename(self.file_paths[0])
            extra = f" (+ еще {len(self.file_paths)-1} файлов)" if len(self.file_paths) > 1 else ""
            self.lbl_file.configure(
                text=f"{first_name}{extra}", text_color=self.text_col)

            self.btn_start.configure(
                state="normal", fg_color=self.accent_col, hover_color=self.accent_hover,
                text_color="#000000" if self.accent_col in ("#00E5FF", "#00F5A0") else "#FFFFFF")

            if not self.output_dir:
                self.output_dir = os.path.dirname(self.file_paths[0])
                self.lbl_dir.configure(
                    text=self.output_dir, text_color=self.text_col)

            self.log(f"[*] Выбрано файлов для обработки: {len(self.file_paths)}")

    def select_output_dir(self):
        directory = filedialog.askdirectory(
            title="Выберите папку для сохранения архивов")
        if directory:
            self.output_dir = directory
            self.lbl_dir_badge.configure(text="● Выбрана пользовательская папка", text_color=self.accent_col)
            self.lbl_dir.configure(text=self.output_dir, text_color=self.text_col)
            self.log(f"[*] Выбрана папка сохранения: {self.output_dir}")

    def toggle_registry_mode(self):
        if self.registry_mode.get() == "local":
            self.frame_google.pack_forget()
            self.frame_local.pack(fill="x", padx=16, pady=(0, 14))
        else:
            self.frame_local.pack_forget()
            self.frame_google.pack(fill="x", padx=16, pady=(0, 14))

    def select_creds(self):
        filepath = filedialog.askopenfilename(
            title="Выберите credentials.json", filetypes=[("JSON Files", "*.json")])
        if filepath:
            self.google_creds_path = filepath
            self.btn_creds.configure(
                text="Ключ выбран ✓", fg_color="#10B981", hover_color="#059669")
            self.log(f"[*] Выбран ключ Google API: {os.path.basename(filepath)}")

    def apply_theme_colors(self):
        theme = self.themes.get(self.current_theme_name,
                                self.themes["Liquid Cyan"])
        self.bg_root = theme["bg_root"]
        self.bg_frame = theme["bg_frame"]
        self.border_glass = theme.get("border_glass", "#1C273C")
        self.border_glow = theme.get("border_glow", "#00E5FF")
        self.accent_col = theme["accent_col"]
        self.accent_hover = theme["accent_hover"]
        self.text_col = theme["text_col"]
        self.text_muted = theme.get("text_muted", "#8B9BB4")
        self.btn_bg = theme["btn_bg"]
        self.btn_hover = theme["btn_hover"]
        self.btn_border = theme.get("btn_border", "#283955")

    def load_settings(self):
        self.settings_file = "settings.json"
        self.ai_api_key = DEFAULT_PROXYAPI_KEY
        self.ai_base_url = DEFAULT_AI_BASE_URL
        self.ai_model = DEFAULT_AI_MODEL
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    s = json.load(f)
                    raw_theme = s.get("theme", "Liquid Cyan")
                    # Совместимость со старыми названиями тем
                    theme_map = {
                        "Dark Orange": "Liquid Sunset",
                        "Dark Teal": "Liquid Cyan",
                        "Light Blue": "Liquid Cyan",
                        "Light Green": "Liquid Emerald"
                    }
                    self.current_theme_name = theme_map.get(raw_theme, raw_theme)
                    if self.current_theme_name not in self.themes:
                        self.current_theme_name = "Liquid Cyan"

                    self.registry_path = s.get("registry_path")
                    self.google_creds_path = s.get("google_creds_path")
                    self.saved_google_url = s.get("google_url", "")
                    self.registry_mode_val = s.get("registry_mode", "local")
                    self.ai_api_key = s.get("ai_api_key", DEFAULT_PROXYAPI_KEY)
                    self.ai_base_url = s.get("ai_base_url", DEFAULT_AI_BASE_URL)
                    self.ai_model = s.get("ai_model", DEFAULT_AI_MODEL)
                    self.ai_workers = int(s.get("ai_workers", 10))
            except Exception:
                pass

        # Авто-поиск credentials.json рядом с exe или в MEIPASS
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        auto_creds = os.path.join(base_path, "credentials.json")
        if os.path.exists(auto_creds):
            self.google_creds_path = auto_creds

    def save_settings(self):
        try:
            s = {
                "theme": self.current_theme_name,
                "registry_path": self.registry_path,
                "google_creds_path": self.google_creds_path,
                "google_url": self.entry_google_url.get() if hasattr(self, 'entry_google_url') else self.saved_google_url,
                "registry_mode": self.registry_mode.get() if hasattr(self, 'registry_mode') else self.registry_mode_val,
                "ai_api_key": getattr(self, 'ai_api_key', DEFAULT_PROXYAPI_KEY),
                "ai_base_url": getattr(self, 'ai_base_url', DEFAULT_AI_BASE_URL),
                "ai_model": getattr(self, 'ai_model', DEFAULT_AI_MODEL),
                "ai_workers": getattr(self, 'ai_workers', 10)
            }
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(s, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    def apply_saved_settings_to_ui(self):
        self.registry_mode.set(self.registry_mode_val)
        self.toggle_registry_mode()
        if self.registry_path:
            self.lbl_registry.configure(
                text=os.path.basename(self.registry_path), text_color=self.text_col)
        if self.google_creds_path:
            self.btn_creds.configure(
                text="Ключ выбран ✓", fg_color="#10B981", hover_color="#059669")
        if self.saved_google_url:
            self.entry_google_url.insert(0, self.saved_google_url)

    def open_settings(self):
        settings_win = ctk.CTkToplevel(self)
        settings_win.title("Параметры • Liquid Glass Suite")
        settings_win.geometry("560x700")
        settings_win.transient(self)
        settings_win.configure(fg_color=self.bg_root)

        # Хедер окна настроек
        hdr = ctk.CTkFrame(settings_win, fg_color=self.bg_frame, border_width=1, border_color=self.border_glass, corner_radius=12)
        hdr.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(
            hdr, text="⚙  ПАРАМЕТРЫ СИСТЕМЫ И ИИ-РОУТЕРА",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=self.accent_col
        ).pack(pady=12)

        frame = ctk.CTkScrollableFrame(
            settings_win, fg_color=self.bg_frame, border_width=1, border_color=self.border_glass, corner_radius=12)
        frame.pack(fill="both", expand=True, padx=20, pady=5)

        # Тема жидкого стекла
        ctk.CTkLabel(frame, text="🎨 Тема жидкого стекла:", text_color=self.text_col, font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", padx=15, pady=(12, 4))
        theme_var = ctk.StringVar(value=self.current_theme_name)

        def on_theme_change(choice):
            self.current_theme_name = choice
            self.apply_theme_colors()
            self.update_all_widgets_colors()
            self.save_settings()
            settings_win.configure(fg_color=self.bg_root)
            hdr.configure(fg_color=self.bg_frame, border_color=self.border_glass)
            frame.configure(fg_color=self.bg_frame, border_color=self.border_glass)

        theme_menu = ctk.CTkOptionMenu(
            frame, values=list(self.themes.keys()), variable=theme_var, command=on_theme_change,
            fg_color=self.btn_bg, button_color=self.accent_col, button_hover_color=self.accent_hover,
            corner_radius=8, height=34
        )
        theme_menu.pack(fill="x", padx=15, pady=(2, 12))

        # Настройки ИИ (ProxyAPI)
        ctk.CTkLabel(frame, text="⚡ ИИ-Роутер (ProxyAPI):", text_color=self.accent_col, font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", padx=15, pady=(10, 4))

        ctk.CTkLabel(frame, text="ProxyAPI API Key:", text_color=self.text_muted, font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=15, pady=(4, 2))
        entry_api_key = ctk.CTkEntry(frame, width=400, show="*", fg_color="#0A101D", border_color=self.border_glass)
        entry_api_key.insert(0, getattr(self, 'ai_api_key', DEFAULT_PROXYAPI_KEY))
        entry_api_key.pack(fill="x", padx=15, pady=2)

        ctk.CTkLabel(frame, text="Base URL:", text_color=self.text_muted, font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=15, pady=(8, 2))
        entry_base_url = ctk.CTkEntry(frame, width=400, fg_color="#0A101D", border_color=self.border_glass)
        entry_base_url.insert(0, getattr(self, 'ai_base_url', DEFAULT_AI_BASE_URL))
        entry_base_url.pack(fill="x", padx=15, pady=2)

        ctk.CTkLabel(frame, text="Модель распознавания:", text_color=self.text_muted, font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=15, pady=(8, 2))
        model_options = [
            "google/gemma-4-26b-a4b-it",
            "gpt-4o-mini",
            "google/gemini-2.0-flash-lite",
            "google/gemini-2.5-flash"
        ]
        current_model = getattr(self, 'ai_model', DEFAULT_AI_MODEL)
        if current_model not in model_options:
            model_options.insert(0, current_model)
        model_var = ctk.StringVar(value=current_model)
        opt_model = ctk.CTkOptionMenu(
            frame, values=model_options, variable=model_var,
            fg_color=self.btn_bg, button_color=self.accent_col, button_hover_color=self.accent_hover,
            corner_radius=8, height=34
        )
        opt_model.pack(fill="x", padx=15, pady=(2, 12))

        # Скорость: параллельные потоки
        ctk.CTkLabel(frame, text="Параллельных потоков (Workers):", text_color=self.text_muted, font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=15, pady=(6, 2))
        workers_var = ctk.StringVar(value=str(getattr(self, 'ai_workers', 10)))
        opt_workers = ctk.CTkOptionMenu(
            frame, values=["5", "8", "10", "12", "15", "20"], variable=workers_var,
            fg_color=self.btn_bg, button_color=self.accent_col, button_hover_color=self.accent_hover,
            corner_radius=8, height=34
        )
        opt_workers.pack(fill="x", padx=15, pady=(2, 16))

        # Локальный реестр
        ctk.CTkLabel(frame, text="📄 Локальный файл реестра:", text_color=self.text_col, font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", padx=15, pady=(8, 4))
        btn_reg = ctk.CTkButton(
            frame, text="Выбрать Excel файл (.xlsx)", command=self.select_registry,
            fg_color=self.btn_bg, hover_color=self.btn_hover, border_width=1, border_color=self.btn_border,
            corner_radius=8, height=36
        )
        btn_reg.pack(fill="x", padx=15, pady=(2, 12))

        # Ключ Google
        ctk.CTkLabel(frame, text="☁ Google Service Account Credentials:", text_color=self.text_col, font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", padx=15, pady=(8, 4))
        btn_cred = ctk.CTkButton(
            frame, text="Выбрать credentials.json", command=self.select_creds,
            fg_color=self.btn_bg, hover_color=self.btn_hover, border_width=1, border_color=self.btn_border,
            corner_radius=8, height=36
        )
        btn_cred.pack(fill="x", padx=15, pady=(2, 16))

        def on_close():
            self.ai_api_key = entry_api_key.get().strip()
            self.ai_base_url = entry_base_url.get().strip()
            self.ai_model = model_var.get().strip()
            try:
                self.ai_workers = int(workers_var.get())
            except Exception:
                self.ai_workers = 10
            self.lbl_term_info.configure(text=f"ProxyAPI • {self.ai_model} • 150 DPI")
            self.save_settings()
            settings_win.destroy()

        btn_ok = ctk.CTkButton(
            settings_win, text="Сохранить настройки", command=on_close,
            fg_color=self.accent_col, hover_color=self.accent_hover,
            text_color="#000000" if self.accent_col in ("#00E5FF", "#00F5A0") else "#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            corner_radius=10, height=42
        )
        btn_ok.pack(pady=15)

    def update_all_widgets_colors(self):
        self.configure(fg_color=self.bg_root)

        if hasattr(self, 'header'):
            self.header.configure(fg_color=self.bg_frame, border_color=self.border_glass)
            self.header.update_theme(self.accent_col, self.border_glass)
            self.header.btn_settings.configure(
                fg_color=self.btn_bg, hover_color=self.btn_hover,
                border_color=self.btn_border, text_color=self.text_col
            )

        for fr in [self.frame_input, self.frame_output, self.frame_registry, getattr(self, 'frame_terminal', None)]:
            if fr:
                fr.configure(fg_color=self.bg_frame, border_color=self.border_glass)

        for btn in [self.btn_select_file, self.btn_select_dir, self.btn_select_registry, self.btn_creds]:
            if btn:
                btn.configure(fg_color=self.btn_bg, hover_color=self.btn_hover,
                              border_color=self.btn_border, text_color=self.text_col)

        self.rb_local.configure(text_color=self.text_col, fg_color=self.accent_col, hover_color=self.accent_hover)
        self.rb_google.configure(text_color=self.text_col, fg_color=self.accent_col, hover_color=self.accent_hover)

        if hasattr(self, 'progress_bar'):
            t = self.themes.get(self.current_theme_name, self.themes["Liquid Cyan"])
            self.progress_bar.update_theme(t["accent_rgb1"], t["accent_rgb2"], self.accent_col)
            self.lbl_progress_pct.configure(text_color=self.accent_col)

        if hasattr(self, 'btn_integrate_only'):
            self.btn_integrate_only.configure(
                border_color=self.border_glass, text_color=self.accent_col)

        if self.btn_start.cget("state") == "normal":
            self.btn_start.configure(
                fg_color=self.accent_col, hover_color=self.accent_hover,
                text_color="#000000" if self.accent_col in ("#00E5FF", "#00F5A0") else "#FFFFFF")
        else:
            self.btn_start.configure(
                fg_color=self.btn_bg, hover_color=self.btn_hover, text_color=self.text_muted)

    def select_registry(self):

        filepath = filedialog.askopenfilename(title="Выберите Excel реестр", filetypes=[
                                              ("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if filepath:
            self.registry_path = filepath
            self.lbl_registry.configure(text=os.path.basename(
                filepath), text_color=self.text_col)
            self.log(f"[*] Выбран Excel-реестр: {self.registry_path}")

    def log(self, message):

        self.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        try:
            with open("debug_log.txt", "a", encoding="utf-8") as f:
                f.write(message + "\n")
        except:
            pass

    def set_buttons_state(self, state):
        self.btn_select_file.configure(state=state)
        self.btn_select_dir.configure(state=state)
        self.btn_select_registry.configure(state=state)
        self.btn_start.configure(state=state)
        self.btn_integrate_only.configure(state=state)

        if state == "disabled":
            self.btn_start.configure(fg_color=self.btn_bg)
            self.btn_integrate_only.configure(fg_color=self.btn_bg)
        else:
            if self.file_paths:
                self.btn_start.configure(fg_color=self.accent_col)
            self.btn_integrate_only.configure(fg_color=self.accent_col)

    def prepare_page_data(self, page, page_num=0, total_pages=1):
        """Быстрое извлечение текста и ультра-быстрый сжатый рендеринг изображения в MuPDF."""
        text_layer = page.get_text("text").strip()
        has_good_text = len(text_layer) > 120 and any(k in text_layer.lower() for k in ["протокол", "испытан", "бетон"])
        use_vision = not has_good_text

        b64_img = None
        if use_vision:
            rect = page.rect
            max_dim = max(rect.width, rect.height)
            # Для четкого распознавания мелкого шрифта и цифр под синей печатью
            # используем масштабирование до ~150-200 DPI (длинная сторона ~1600px)
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

    def query_ai_for_page(self, page_data, total_pages=1):
        """Отправка подготовленной страницы в ИИ через постоянную HTTP-сессию с минимальным временем ответа."""
        page_num = page_data["page_num"]
        text_layer = page_data["text_layer"]
        use_vision = page_data["use_vision"]
        b64_img = page_data["b64_img"]

        api_key = getattr(self, 'ai_api_key', DEFAULT_PROXYAPI_KEY).strip()
        base_url = getattr(self, 'ai_base_url', DEFAULT_AI_BASE_URL).rstrip('/')
        model = getattr(self, 'ai_model', DEFAULT_AI_MODEL)

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
                self.log(f"  [>] Стр. {page_num}/{total_pages}: отправка в ИИ [{mode_desc}]...")
                resp = self.http_session.post(
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
                    self.log(f"  [!] Ответ ProxyAPI ({resp.status_code}) на стр. {page_num}: {resp.text[:150]}")
                    time.sleep(0.5)
            except Exception as ex:
                self.log(f"  [!] Ошибка запроса к ИИ (стр. {page_num}, попытка {attempt+1}/2): {str(ex)}")
                time.sleep(0.5)

        return None

    def extract_protocol_data_ai(self, page, page_num=0, total_pages=1):
        """Извлечение данных протокола через ИИ-роутер ProxyAPI."""
        p_data = self.prepare_page_data(page, page_num=page_num, total_pages=total_pages)
        return self.query_ai_for_page(p_data, total_pages=total_pages)

    def clean_protocol_number(self, num_str, concreting_date=None, protocol_date=None, age=None):
        if not num_str:
            return None
        num_str = str(num_str).strip()

        # 1. Отсекаем все буквенные префиксы
        core_str = re.sub(r'^[^0-9]*[СC][ВB][БB5]\s*-?',
                          '', num_str, flags=re.IGNORECASE)

        # 2. Исправляем буквы, визуально похожие на цифры
        replacements = {
            "O": "0", "o": "0", "О": "0", "о": "0",
            "l": "1", "I": "1", "І": "1", "|": "1", "!": "1",
            "З": "3", "б": "6", "В": "8", "B": "8"
        }
        for k, v in replacements.items():
            core_str = core_str.replace(k, v)

        # 3. Очищаем весь оставшийся буквенный мусор спереди
        core_str = re.sub(r'^[^0-9]+', '', core_str)

        # Нормализуем разделители
        core_str = core_str.replace("/", "_").replace("\\", "_")

        # 4. Восстанавливаем типичные слияния символов
        core_str = re.sub(r'([0-9]+)46\-1', r'\1_26-1', core_str)
        core_str = re.sub(r'([0-9]+)46\-', r'\1_26-', core_str)

        if core_str.startswith("707") and "26" in core_str:
            core_str = "10-1" + core_str

        # 5. Жесткое правило: протоколы начинаются ТОЛЬКО с 10, 12, 28 или 7.
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

        # 6. КОРРЕКЦИЯ ИСКАЖЕНИЙ ПОД ПЕЧАТЬЮ ПО ДАТЕ БЕТОНИРОВАНИЯ/ИЗГОТОВЛЕНИЯ:
        # Номер протокола ВСЕГДА содержит дату заливки: СВБ-{префикс}-{ДДММ}_{ГГ}-{серия}
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

        # Если даты бетонирования нет, но есть дата протокола и возраст / номер:
        if not dt_c and dt_p:
            if age and str(age).isdigit():
                dt_c = dt_p - datetime.timedelta(days=int(age))
            elif prefix in ('28', '12'):
                dt_c = dt_p - datetime.timedelta(days=28)
            elif prefix == '7':
                dt_c = dt_p - datetime.timedelta(days=7)
            else:
                # Для префикса 10 (конструкция может быть как 7, так и 28 суток):
                # Пробуем проверить день и месяц из самого номера
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
                    # По умолчанию для неразрушающего контроля чаще всего проектный возраст 28
                    dt_c = dt_p - datetime.timedelta(days=28)

        if dt_c:
            target_ddmm = f"{dt_c.day:02d}{dt_c.month:02d}"
            target_yy = f"{dt_c.year % 100:02d}"

        # Проверяем структуру core_str и заменяем искаженные под печатью цифры
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

    def start_processing(self):
        if not self.file_paths:
            return

        if self.is_processing:
            return

        api_key = getattr(self, 'ai_api_key', DEFAULT_PROXYAPI_KEY).strip()
        if not api_key:
            messagebox.showwarning("API Ключ", "Не указан API ключ ProxyAPI в настройках!")
            return

        self.is_processing = True
        self.set_buttons_state("disabled")
        if hasattr(self, 'header'):
            self.header.set_status("ИИ-АНАЛИЗ И СБОРКА ПРОТОКОЛОВ...", active=True)
        self.log_area.delete("0.0", "end")
        self.log("=== Начало обработки файлов ===")

        thread = threading.Thread(target=self.process_pdfs, daemon=True)
        thread.start()

    def process_pdfs(self):
        temp_dir = None
        try:
            project_output_dir = self.output_dir
            temp_dir = os.path.join(
                project_output_dir, "Сборка_протоколов_temp")
            os.makedirs(temp_dir, exist_ok=True)

            all_records = []
            files_to_zip = []

            total_files = len(self.file_paths)
            self.progress_bar.set(0)

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

            for file_idx, file_path in enumerate(self.file_paths):
                try:
                    self.log(
                        f"\n[{file_idx+1}/{total_files}] Открытие: {os.path.basename(file_path)}")
                    doc = fitz.open(file_path)
                    num_pages = len(doc)

                    # 1. Быстрая предварительная подготовка всех страниц в основном потоке (<1 сек)
                    self.log(f"[*] Подготовка {num_pages} стр. к параллельному анализу...")
                    pages_data = []
                    for p_i in range(num_pages):
                        p_obj = doc.load_page(p_i)
                        p_data = self.prepare_page_data(
                            p_obj, page_num=p_i + 1, total_pages=num_pages
                        )
                        pages_data.append(p_data)

                    # 2. Параллельный опрос ИИ через пул потоков (ускорение в 5-8 раз)
                    workers = int(getattr(self, 'ai_workers', 5))
                    self.log(f"[*] Запуск параллельного распознавания (потоков: {workers})...")
                    ai_results = {}
                    completed_pages = 0

                    with ThreadPoolExecutor(max_workers=workers) as executor:
                        future_to_idx = {
                            executor.submit(self.query_ai_for_page, p_data, num_pages): p_data["page_num"] - 1
                            for p_data in pages_data
                        }
                        for future in as_completed(future_to_idx):
                            p_idx = future_to_idx[future]
                            completed_pages += 1
                            overall_progress = (file_idx + (completed_pages / num_pages)) / total_files
                            self.progress_bar.set(overall_progress)
                            try:
                                res = future.result()
                                ai_results[p_idx] = res
                            except Exception as e:
                                self.log(f"  [!] Ошибка анализа стр. {p_idx + 1}: {str(e)}")
                                ai_results[p_idx] = None

                    protocols = []
                    current_protocol = None

                    # 3. Мгновенная последовательная сборка и группировка протоколов
                    for page_num in range(num_pages):
                        page = doc.load_page(page_num)

                        p_num, p_date, p_building = None, None, None
                        p_structure, p_concreting_date, p_age, p_cube = None, None, None, False
                        is_continuation = False

                        # ЖЕСТКОЕ ОГРАНИЧЕНИЕ: Максимум 2 страницы на протокол
                        if current_protocol and (page_num - current_protocol["start_page"] == 2):
                            protocols.append(current_protocol)
                            current_protocol = None
                            self.log(
                                f"[*] Стр. {page_num}: принудительное завершение предыдущего протокола (лимит 2 стр.)")

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

                            # Взаимная математическая проверка и восстановление дат:
                            # Дата проведения испытаний = Дата изготовления + Возраст суток
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

                            # Определяем возраст по префиксу номера (СВБ-28 -> 28, СВБ-7 -> 7, СВБ-10 -> 7/разница, СВБ-12 -> 28)
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

                            # Восстановление отсутствующей даты
                            if not dt_c and dt_p and days_num:
                                dt_c = dt_p - datetime.timedelta(days=days_num)
                                p_concreting_date = dt_c.strftime("%d.%m.%Y")
                            if not dt_p and dt_c and days_num:
                                dt_p = dt_c + datetime.timedelta(days=days_num)
                                p_date = dt_p.strftime("%d.%m.%Y")

                            if raw_num:
                                p_num = self.clean_protocol_number(
                                    raw_num,
                                    concreting_date=p_concreting_date,
                                    protocol_date=p_date,
                                    age=p_age
                                )

                            if p_num:
                                self.log(f"  [✓] Стр. {page_num + 1}: Распознан {p_num} (от {p_date or '?'}), Здание: {p_building or '?'}")
                        else:
                            # Резервный разбор по текстовому слою (fallback при сетевой ошибке)
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
                                prefix = text[max(0, match.start() - 30):match.start()].lower()
                                if "продолжени" not in prefix:
                                    p_date = match.group(2)
                                    p_num = self.clean_protocol_number(
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
                            self.log(f"  [+] Стр. {page_num + 1}: продолжение протокола {current_protocol['number']}")
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
                        self.log(
                            f"[!] Не найдено ни одного протокола в файле {os.path.basename(file_path)}")
                        doc.close()
                        continue

                    # --- ИДЕАЛЬНОЕ СЛИЯНИЕ ДВУХСТРАНИЧНЫХ ПРОТОКОЛОВ ---
                    merged_protocols = []
                    for prot in protocols:
                        if not merged_protocols:
                            merged_protocols.append(prot)
                            continue

                        prev_prot = merged_protocols[-1]
                        prev_len = prev_prot["end_page"] - \
                            prev_prot["start_page"] + 1

                        # Мы можем объединить страницу, только если предыдущая часть - ровно 1 страница
                        can_merge = (prev_len == 1) and (
                            prot["start_page"] == prev_prot["end_page"] + 1)

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

                    # Формирование файлов для текущего PDF
                    for idx, prot in enumerate(protocols, start=1):
                        clean_num = prot["number"].replace(
                            "/", "_").replace("\\", "_")
                        date_str = prot["date"] if prot["date"] else "Без_даты"
                        clean_date = date_str.replace(
                            "/", "_").replace("\\", "_")
                        building_str = prot.get("building")
                        building_folder_name = f"Корпус {building_str}" if building_str else "Корпус_Неизвестен"

                        start_p = prot["start_page"]
                        end_p = prot["end_page"]
                        num_p = end_p - start_p + 1

                        page_range_str = f"{start_p+1}" if num_p == 1 else f"{start_p+1}-{end_p+1}"

                        # Определяем папку по префиксу (10, 12, 7_28)
                        m_pref = re.search(r'СВБ-(10|12|28|7)', prot["number"])
                        p_val = m_pref.group(1) if m_pref else "Неизвестно"
                        prefix_folder = "7_28" if p_val in ["7", "28"] else p_val

                        # ФОРМАТ: Протокол_СВБ-10-1807_26-1_от_25.07.2026.pdf
                        # Замечание: убрали _стр_X-Y по просьбе пользователя
                        filename = f"Протокол_{clean_num}_от_{clean_date}.pdf"

                        # Сохраняем в выбранную папку со структурой Корпус -> Префикс
                        target_dir = os.path.join(
                            project_output_dir, building_folder_name, prefix_folder)
                        os.makedirs(target_dir, exist_ok=True)

                        target_filepath = os.path.join(target_dir, filename)

                        # Для предотвращения перезаписи при одинаковых именах из разных файлов
                        counter = 1
                        while os.path.exists(target_filepath):
                            filename_base = f"Протокол_{clean_num}_от_{clean_date}_{counter}.pdf"
                            target_filepath = os.path.join(
                                target_dir, filename_base)
                            filename = filename_base
                            counter += 1

                        new_doc = fitz.open()
                        new_doc.insert_pdf(doc, from_page=start_p, to_page=end_p)
                        os.makedirs(os.path.dirname(target_filepath), exist_ok=True)
                        new_doc.save(target_filepath, garbage=3, deflate=True)
                        new_doc.close()
                        files_to_zip.append(target_filepath)

                        note = "уверенно" if not prot["uncertain"] else "требует ручной проверки"

                        # ПРИОРИТЕТ 1: Возраст, прямо указанный в протоколе (распознанный ИИ или текстом)
                        final_age = str(prot.get("age")).strip() if prot.get("age") else None
                        if final_age not in ["7", "28"]:
                            if final_age and "28" in final_age:
                                final_age = "28"
                            elif final_age and "7" in final_age:
                                final_age = "7"
                            else:
                                final_age = None

                        # ПРИОРИТЕТ 2: Если возраст не указан явно, рассчитываем по разнице дат испытания и бетонирования
                        if not final_age:
                            final_age = calculate_age_from_dates(
                                prot["date"], prot.get("concreting_date"))

                        # ПРИОРИТЕТ 2.1: Если дата бетонирования не извлечена отдельно, берем её прямо из номера протокола СВБ-...-ДДММ_ГГ
                        if not final_age and prot.get("date") and prot.get("number"):
                            m_num_date = re.search(r'[-_](\d{2})(\d{2})_(\d{2})', prot["number"])
                            if m_num_date:
                                dd, mm, yy = m_num_date.groups()
                                final_age = calculate_age_from_dates(prot["date"], f"{dd}.{mm}.20{yy}")

                        # ПРИОРИТЕТ 3: Определяем is_cube и возраст ПО НОМЕРУ ПРОТОКОЛА:
                        # СВБ-10-... = протокол на конструкцию / неразрушающий контроль (НЕ кубики, может быть 7 или 28 суток)
                        # СВБ-12-... = протокол на конструкцию / неразрушающий контроль (НЕ кубики, 28 суток, редкое исключение)
                        # СВБ-7-... = кубики, 7 суток
                        # СВБ-28-... = кубики, 28 суток
                        protocol_prefix_match = re.search(r'СВБ-(10|12|28|7)[\-_\\]', prot["number"])
                        if protocol_prefix_match:
                            p_prefix_val = protocol_prefix_match.group(1)
                            if p_prefix_val in ("10", "12"):
                                final_is_cube = False
                                if p_prefix_val == "12" and not final_age:
                                    final_age = "28"
                            else:
                                final_is_cube = True
                                # Для кубиков 28 и 7 возраст строго определяется префиксом
                                if p_prefix_val == "28":
                                    final_age = "28"
                                elif p_prefix_val == "7":
                                    final_age = "7"
                        else:
                            final_is_cube = bool(prot.get("is_cube", False))

                        # Если возраст для неразрушающего контроля так и не определился, по умолчанию 28
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
                    self.log(
                        f"[ОШИБКА] Ошибка при обработке {os.path.basename(file_path)}: {str(ex)}")

            # --- Завершение обработки всех файлов ---
            if all_records:
                final_excel_path = os.path.join(project_output_dir, "Реестр_разбивки.xlsx")
                self.create_formatted_registry(all_records, final_excel_path)
                files_to_zip.append(final_excel_path)

                self.perform_integration(all_records, project_output_dir)

                zip_filename = os.path.join(
                    project_output_dir, "Все_протоколы_архив.zip")
                with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for file in files_to_zip:
                        zipf.write(file, os.path.basename(file))

                self.progress_bar.set(1.0)
                self.log(
                    f"\n[✓] Готово! \nВсего обработано файлов: {total_files}\nФайлы распакованы в папки по корпусам: {project_output_dir}\nТакже создан ZIP-архив: {zip_filename}\nПромежуточный реестр сохранен по шаблону: {final_excel_path}")
            else:
                self.log(
                    "\n[!] Нет данных для сохранения (ни один протокол не найден).")

        except Exception as e:
            self.log(f"\n[КРИТИЧЕСКАЯ ОШИБКА] {str(e)}")
        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.is_processing = False
            self.after(0, lambda: self.set_buttons_state("normal"))
            if hasattr(self, 'header'):
                self.after(0, lambda: self.header.set_status("СИСТЕМА ГОТОВА К РАБОТЕ", active=False))

    def create_formatted_registry(self, all_records, output_filepath):
        """Создает промежуточный реестр в точном формате итоговой таблицы (с колонками 7, 28 суток и кубиков)."""
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

            # Целевой столбец (1-based index)
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

    def perform_integration(self, all_records, project_output_dir):
        # --- ИНТЕГРАЦИЯ В СУЩЕСТВУЮЩИЙ РЕЕСТР ---
        mode = getattr(self, 'registry_mode', None)
        if mode:
            mode_val = mode.get()
            unmatched_records = all_records.copy()

            if mode_val == "local" and getattr(self, 'registry_path', None) and os.path.exists(self.registry_path):
                self.log(
                    "\n[*] Начинаем интеграцию в локальный реестр Excel...")
                try:
                    wb = openpyxl.load_workbook(self.registry_path)

                    # Перебираем все листы, кроме 'Нераспределенные'
                    sheets_to_process = [
                        sheet for sheet in wb.sheetnames if sheet != "Нераспределенные"]

                    for sheet_name in sheets_to_process:
                        ws = wb[sheet_name]
                        self.log(f"  Проверка листа: {sheet_name}")

                        headers = [str(ws.cell(row=1, column=c).value or "").lower() for c in range(1, ws.max_column + 1)]
                        cols = find_registry_columns(headers)
                        self.log(f"Найдены столбцы: дата={cols['date']}, здание={cols['bldg']}, констр={cols['struct']}, 7={cols['nz7']}, 28={cols['nz28']}, 7к={cols['cub7']}, 28к={cols['cub28']}")

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

                                target_col = get_target_col(t_cols, p_age, p_cube)
                                if target_col:
                                    cell = target_ws.cell(row=new_row, column=target_col)
                                    cell.value = f"№{p_num} от {p_date}"
                                    cell.hyperlink = f"Корпус {p_bldg}\\{p_prefix}\\{rec['Имя файла']}"
                                    cell.style = "Hyperlink"

                                for c in range(1, max(12, target_ws.max_column + 1)):
                                    cell = target_ws.cell(row=new_row, column=c)
                                    cell.border = thin_border
                                    cell.alignment = center_align
                                    cell.fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

                                self.log(f"    [+] Создана новая строка в {target_ws.title} для протокола №{p_num}")
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
                                rec.get("Наименование конструкции", ""), rec.get(
                                    "Дата бетонирования", ""),
                                rec["Имя файла"], rec["Префикс_Колонка"]
                            ])

                    out_registry = os.path.join(
                        project_output_dir, "Реестр_обновлен.xlsx")
                    wb.save(out_registry)
                    self.log(
                        f"[✓] Локальный реестр успешно обновлен: {out_registry}")
                except Exception as e:
                    self.log(f"[ОШИБКА] Не удалось обновить реестр: {str(e)}")

            elif mode_val == "google" and getattr(self, 'google_creds_path', None) and self.entry_google_url.get():
                self.log("\n[*] Начинаем интеграцию в Google Таблицу...")
                try:
                    scope = ['https://spreadsheets.google.com/feeds',
                             'https://www.googleapis.com/auth/drive']
                    creds = ServiceAccountCredentials.from_json_keyfile_name(
                        self.google_creds_path, scope)
                    client = gspread.authorize(creds)

                    url = self.entry_google_url.get().strip()
                    sh = client.open_by_url(url)

                    for ws in sh.worksheets():
                        if ws.title == "Нераспределенные":
                            continue
                        self.log(f"  Проверка листа: {ws.title}")

                        data = ws.get_all_values()
                        if not data:
                            continue

                        cols = find_registry_columns(data[0])
                        self.log(
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
                                self.log(
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
                            ws.batch_update(
                                updates, value_input_option='USER_ENTERED')
                            time.sleep(1)  # Пауза для предотвращения 429 ошибки квоты
                        if format_updates:
                            sh.batch_update({'requests': format_updates})
                            time.sleep(1)

                    if unmatched_records:
                        still_unmatched = []
                        grouped = {}
                        for rec in unmatched_records:
                            k = (str(rec.get("Корпус", "")).strip(), str(rec.get("Дата бетонирования", "")).strip(), str(rec.get("Наименование конструкции", "")).strip())
                            if k not in grouped: grouped[k] = []
                            grouped[k].append(rec)

                        # Кэшируем доступные листы один раз, чтобы не делать повторные сетевые запросы
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

                        # Пакетная вставка новых строк для каждого листа (вместо десятков отдельных запросов)
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
                                        target_col = get_target_col(t_cols, rec.get("Возраст суток", ""), rec.get("Кубики", False))
                                        if target_col:
                                            row_data[target_col - 1] = f"№{rec['Номер протокола']} от {rec['Дата протокола']}"

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
                                    self.log(f"    [+] Пакетно добавлено {len(batch_rows)} новых строк в {target_ws.title}")
                            except Exception as ex_b:
                                self.log(f"    [!] Ошибка добавления строк в {target_ws.title}: {str(ex_b)}")
                                for _, _, _, recs in items:
                                    for r in recs:
                                        still_unmatched.append(r)

                        unmatched_records = still_unmatched

                    if unmatched_records:
                        try:
                            ws_un = sh.worksheet("Нераспределенные")
                        except gspread.exceptions.WorksheetNotFound:
                            ws_un = sh.add_worksheet(
                                title="Нераспределенные", rows="100", cols="10")
                            ws_un.append_row(["Номер протокола", "Дата протокола", "Корпус",
                                             "Наименование конструкции", "Дата бетонирования", "Имя файла", "Ожидаемый столбец"])

                        new_rows = []
                        self.log(
                            f"      [!] Отправлено в Нераспределенные: {len(unmatched_records)} шт.")
                        for rec in unmatched_records:
                            new_rows.append([
                                rec["Номер протокола"], rec["Дата протокола"], rec["Корпус"],
                                rec.get("Наименование конструкции", ""), rec.get(
                                    "Дата бетонирования", ""),
                                rec["Имя файла"], rec["Префикс_Колонка"]
                            ])
                        ws_un.append_rows(new_rows)

                    self.log(f"[✓] Google Таблица успешно обновлена!")
                except Exception as e:
                    self.log(
                        f"[ОШИБКА] Не удалось обновить Google Таблицу: {str(e)}")

    def integrate_only(self):
        filepath = filedialog.askopenfilename(
            title="Выберите Реестр_разбивки.xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not filepath:
            return

        try:
            df = pd.read_excel(filepath)
            df = df.fillna("")

            cols = [str(c).lower() for c in df.columns]
            has_template_cols = any("неразрушающ" in c or "кубик" in c or "куб" in c for c in cols)

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

            project_output_dir = os.path.dirname(filepath)

            def run():
                self.set_buttons_state("disabled")
                if hasattr(self, 'header'):
                    self.after(0, lambda: self.header.set_status("ИНТЕГРАЦИЯ В РЕЕСТР...", active=True))
                self.log_area.delete("0.0", "end")
                self.log(f"=== Начало интеграции из готового файла ({len(all_records)} протоколов) ===")
                self.perform_integration(all_records, project_output_dir)
                self.log("=== Интеграция завершена ===")
                self.set_buttons_state("normal")
                if hasattr(self, 'header'):
                    self.after(0, lambda: self.header.set_status("СИСТЕМА ГОТОВА К РАБОТЕ", active=False))
            threading.Thread(target=run, daemon=True).start()
        except Exception as e:
            messagebox.showerror(
                "Ошибка", f"Не удалось прочитать файл: {str(e)}")


if __name__ == "__main__":
    app = PDFSplitterApp()
    app.mainloop()
