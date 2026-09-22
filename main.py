import os
import sys
import subprocess
import threading
import zipfile
import re
import shutil
import difflib

# Автоматическая установка customtkinter, если его нет
try:
    import customtkinter as ctk
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
    import customtkinter as ctk

from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image, ImageEnhance, ImageOps

class PDFSplitterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Настройка современного дизайна
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue") # Базовая тема, но цвета мы зададим свои
        
        self.title("Умная разбивка PDF-протоколов")
        self.geometry("900x750")
        
        # Фирменные цвета
        self.bg_root = "#181818"
        self.bg_frame = "#242424"
        self.accent_col = "#FF7B00"
        self.accent_hover = "#FF9522"
        self.text_col = "#EAEAEA"
        self.btn_bg = "#3A3A3A"
        self.btn_hover = "#4A4A4A"
        
        self.configure(fg_color=self.bg_root)
        
        self.file_path = None
        self.output_dir = None
        self.is_processing = False
        
        self.setup_ui()
        
    def setup_ui(self):
        # Заголовок
        self.title_lbl = ctk.CTkLabel(
            self, text="Умная разбивка PDF-протоколов", 
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"), 
            text_color=self.accent_col
        )
        self.title_lbl.pack(pady=(30, 20))

        # ФРЕЙМ 1: Выбор файла
        self.frame_input = ctk.CTkFrame(self, fg_color=self.bg_frame, corner_radius=10)
        self.frame_input.pack(fill="x", padx=40, pady=(0, 15))
        
        self.btn_select_file = ctk.CTkButton(
            self.frame_input, text="1. Выбрать PDF", command=self.select_file,
            fg_color=self.btn_bg, hover_color=self.btn_hover,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            corner_radius=8, width=160, height=40
        )
        self.btn_select_file.pack(side="left", padx=20, pady=20)
        
        self.lbl_file = ctk.CTkLabel(
            self.frame_input, text="Файл не выбран", 
            text_color="#999999", font=ctk.CTkFont(family="Segoe UI", size=13),
            wraplength=550, justify="left"
        )
        self.lbl_file.pack(side="left", fill="x", expand=True, padx=(0, 20))

        # ФРЕЙМ 2: Выбор папки
        self.frame_output = ctk.CTkFrame(self, fg_color=self.bg_frame, corner_radius=10)
        self.frame_output.pack(fill="x", padx=40, pady=(0, 15))
        
        self.btn_select_dir = ctk.CTkButton(
            self.frame_output, text="2. Папка сохранения", command=self.select_output_dir,
            fg_color=self.btn_bg, hover_color=self.btn_hover,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            corner_radius=8, width=160, height=40
        )
        self.btn_select_dir.pack(side="left", padx=20, pady=20)
        
        self.lbl_dir = ctk.CTkLabel(
            self.frame_output, text="Сохранить рядом с исходным файлом (по умолчанию)", 
            text_color="#999999", font=ctk.CTkFont(family="Segoe UI", size=13),
            wraplength=550, justify="left"
        )
        self.lbl_dir.pack(side="left", fill="x", expand=True, padx=(0, 20))

        # ФРЕЙМ 3: Опции
        self.frame_options = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_options.pack(fill="x", padx=40, pady=(5, 10))
        
        self.chk_force_ocr = ctk.CTkCheckBox(
            self.frame_options, text="Глубокое OCR-распознавание (игнорирует текстовый слой, читает как картинку)", 
            fg_color=self.accent_col, hover_color=self.accent_hover,
            font=ctk.CTkFont(family="Segoe UI", size=13), text_color=self.text_col
        )
        self.chk_force_ocr.pack(side="left", padx=10)

        # КНОПКА ЗАПУСКА
        self.btn_start = ctk.CTkButton(
            self, text="Начать обработку", command=self.start_processing,
            fg_color=self.btn_bg, hover_color=self.btn_hover, # Будет оранжевой при активации
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            corner_radius=8, height=50, state="disabled"
        )
        self.btn_start.pack(fill="x", padx=40, pady=(15, 20))

        # ЛОГИ
        self.log_label = ctk.CTkLabel(
            self, text="Журнал работы:", 
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"), 
            text_color=self.accent_col
        )
        self.log_label.pack(anchor="w", padx=40, pady=(0, 5))
        
        self.log_area = ctk.CTkTextbox(
            self, fg_color="#121212", text_color=self.text_col,
            font=ctk.CTkFont(family="Consolas", size=13), corner_radius=10,
            border_width=1, border_color="#333333"
        )
        self.log_area.pack(fill="both", expand=True, padx=40, pady=(0, 30))
        
    def select_file(self):
        filepath = filedialog.askopenfilename(title="Выберите PDF файл", filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.file_path = filepath
            self.lbl_file.configure(text=self.file_path, text_color=self.text_col)
            
            self.btn_start.configure(state="normal", fg_color=self.accent_col, hover_color=self.accent_hover)
            
            if not self.output_dir:
                self.output_dir = os.path.dirname(self.file_path)
                self.lbl_dir.configure(text=self.output_dir, text_color=self.text_col)
                
            self.log(f"[*] Выбран файл: {self.file_path}")

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Выберите папку для сохранения архивов")
        if directory:
            self.output_dir = directory
            self.lbl_dir.configure(text=self.output_dir, text_color=self.text_col)
            self.log(f"[*] Выбрана папка сохранения: {self.output_dir}")
            
    def log(self, message):
        self.after(0, self._append_log, message)
        
    def _append_log(self, message):
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        
    def set_buttons_state(self, state):
        self.btn_select_file.configure(state=state)
        self.btn_select_dir.configure(state=state)
        self.btn_start.configure(state=state)
        self.chk_force_ocr.configure(state=state)
        
        if state == "disabled":
            self.btn_start.configure(fg_color=self.btn_bg)
        else:
            self.btn_start.configure(fg_color=self.accent_col)
            
    def find_tesseract(self):
        try:
            pytesseract.get_tesseract_version()
            return True
        except:
            pass
            
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Tesseract-OCR\tesseract.exe")
        ]
        for path in tesseract_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return True
        return False
        
    def clean_protocol_number(self, num_str):
        num_str = num_str.strip()
        
        # 1. Отсекаем все буквенные префиксы (любые варианты СВБ, CBB и прочий мусор)
        core_str = re.sub(r'^[^0-9]*[СC][ВB][БB5]\s*-?', '', num_str, flags=re.IGNORECASE)
        
        # 2. Исправляем буквы, которые OCR принял за цифры (например lO-l707 -> 10-1707)
        replacements = {
            "O": "0", "o": "0", "О": "0", "о": "0",
            "l": "1", "I": "1", "І": "1", "|": "1", "!": "1",
            "З": "3", "б": "6", "В": "8", "B": "8"
        }
        for k, v in replacements.items():
            core_str = core_str.replace(k, v)
            
        # 3. Очищаем весь оставшийся буквенный мусор спереди
        core_str = re.sub(r'^[^0-9]+', '', core_str)
        
        # 4. Восстанавливаем типичные слияния символов в OCR
        # Слэш и двойка часто сливаются в цифру 4 (например /26-1 читается как 46-1)
        core_str = re.sub(r'([0-9]+)46\-1', r'\1_26-1', core_str)
        core_str = re.sub(r'([0-9]+)46\-', r'\1_26-', core_str)
        
        # Специфичное восстановление потерянного префикса, если OCR "съел" начало
        # (например 707_26-1 вместо 10-1707_26-1)
        if core_str.startswith("707") and "26" in core_str:
            core_str = "10-1" + core_str
        
        # 5. Жесткое правило: протоколы начинаются ТОЛЬКО с 10, 12, 28 или 7.
        # Ищем это обязательное начало, за которым идет разделитель (- или _)
        match = re.search(r'(10|12|28|7)[\-_\\]', core_str)
        if match:
            # Отрезаем любой цифровой мусор, который OCR приписал в начале
            core_str = core_str[match.start():]
        else:
            # Если правильного префикса нет, возможно OCR потерял первую цифру
            # Например, 0- вместо 10-, 2- вместо 12-, 8- вместо 28-
            match_broken = re.search(r'(0|2|8)([\-_\\])', core_str)
            if match_broken:
                digit = match_broken.group(1)
                sep = match_broken.group(2)
                prefix = '10' if digit == '0' else ('12' if digit == '2' else '28')
                core_str = prefix + sep + core_str[match_broken.end():]
            else:
                # Если разделителя почему-то нет, ищем только двузначные префиксы. 
                # Одиночную 7 без дефиса не берем, чтобы не обрезать номера вроде 1707 до 707.
                match_any = re.search(r'(10|12|28)', core_str)
                if match_any:
                    core_str = core_str[match_any.start():]
                
        core_str = core_str.replace("/", "_").replace("\\", "_")
                    
        return f"СВБ-{core_str}"

    def start_processing(self):
        if not self.file_path:
            return
            
        if self.is_processing:
            return
            
        self.use_ocr = self.chk_force_ocr.get()
        
        if not self.use_ocr:
            try:
                doc = fitz.open(self.file_path)
                has_text = any(doc.load_page(i).get_text("text").strip() for i in range(min(5, len(doc))))
                doc.close()
                
                if not has_text:
                    ans = messagebox.askyesno(
                        "Требуется OCR", 
                        "В PDF-файле не найден распознанный текст.\n"
                        "Включить автоматическое распознавание (OCR)?"
                    )
                    if not ans:
                        return
                    self.use_ocr = True
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось прочитать файл: {str(e)}")
                return
                
        if self.use_ocr:
            if not self.find_tesseract():
                messagebox.showinfo(
                    "Tesseract не найден", 
                    "Программа не смогла найти Tesseract OCR.\nУкажите путь к tesseract.exe"
                )
                tess_path = filedialog.askopenfilename(
                    title="Укажите путь к tesseract.exe",
                    filetypes=[("Executable", "tesseract.exe *.exe")]
                )
                if tess_path:
                    pytesseract.pytesseract.tesseract_cmd = tess_path
                else:
                    self.log("[!] Отмена обработки: Tesseract OCR не настроен.")
                    return
        
        self.is_processing = True
        self.set_buttons_state("disabled")
        self.log_area.delete("0.0", "end")
        self.log("=== Начало глубокой обработки ===")
        
        thread = threading.Thread(target=self.process_pdf, daemon=True)
        thread.start()
        
    def process_pdf(self):
        temp_dir = None
        try:
            doc = fitz.open(self.file_path)
            num_pages = len(doc)
            self.log(f"Документ открыт. Всего страниц: {num_pages}")
                
            protocols = []
            current_protocol = None
            
            pattern = re.compile(r"ПРОТОКОЛ\s*(?:№|N|Ne|N2)?\s*([A-Za-zА-Яа-я0-9\-_/\\]+)(?:\s*от\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}))?", re.IGNORECASE)
            
            for page_num in range(num_pages):
                page = doc.load_page(page_num)
                
                best_match = None
                p_num, p_date = None, None
                
                # ЖЕСТКОЕ ОГРАНИЧЕНИЕ: Максимум 2 страницы на протокол
                if current_protocol and (page_num - current_protocol["start_page"] == 2):
                    protocols.append(current_protocol)
                    current_protocol = None
                    self.log(f"[*] Стр. {page_num}: принудительное завершение предыдущего протокола (лимит 2 стр.)")
                
                if self.use_ocr:
                    pix = page.get_pixmap(dpi=400)
                    base_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    
                    self.log(f"Анализ стр. {page_num + 1}/{num_pages}...")
                    
                    for rotation in [0, 180]:
                        if rotation == 180 and best_match and best_match[1]:
                            break
                            
                        if rotation == 180:
                            self.log(f"[-] Стр. {page_num + 1}: Данные неполные. Переворот на 180 градусов и повтор...")
                            rotated_img = base_img.rotate(180, expand=True)
                        else:
                            rotated_img = base_img

                        for attempt in range(5):
                            if attempt == 0:
                                img, config = rotated_img, '--psm 3'
                            elif attempt == 1:
                                img, config = ImageEnhance.Contrast(ImageOps.grayscale(rotated_img)).enhance(2.0), '--psm 3'
                            elif attempt == 2:
                                img, config = rotated_img.convert('L').point(lambda x: 0 if x < 130 else 255, '1'), '--psm 6'
                            elif attempt == 3:
                                img, config = rotated_img.convert('L').point(lambda x: 0 if x < 180 else 255, '1'), '--psm 6'
                            else:
                                w, h = rotated_img.size
                                img = rotated_img.crop((0, 0, w, h // 3))
                                config = '--psm 6'
                                
                            text = pytesseract.image_to_string(img, lang="rus+eng", config=config)
                            header_text = text[:3000].replace('\n', ' ')
                            
                            match = pattern.search(header_text)
                            if match:
                                prefix = header_text[max(0, match.start() - 30):match.start()].lower()
                                if "продолжени" in prefix:
                                    continue
                                    
                                temp_num = self.clean_protocol_number(match.group(1))
                                temp_date = match.group(2)
                                
                                if temp_date:
                                    best_match = (temp_num, temp_date)
                                    break
                                elif best_match is None:
                                    best_match = (temp_num, None)
                                    
                        if best_match and best_match[1]:
                            break
                            
                    if best_match:
                        p_num, p_date = best_match
                else:
                    text = page.get_text("text").replace('\n', ' ')
                    match = pattern.search(text[:3000])
                    if match:
                        prefix = text[max(0, match.start() - 30):match.start()].lower()
                        if "продолжени" not in prefix:
                            p_num = self.clean_protocol_number(match.group(1))
                            p_date = match.group(2)
                    
                if p_num:
                    if current_protocol:
                        if current_protocol["number"] == p_num and not current_protocol["uncertain"]:
                            current_protocol["end_page"] = page_num
                        else:
                            protocols.append(current_protocol)
                            current_protocol = {
                                "number": p_num, "date": p_date,
                                "start_page": page_num, "end_page": page_num,
                                "uncertain": p_date is None
                            }
                    else:
                        current_protocol = {
                            "number": p_num, "date": p_date,
                            "start_page": page_num, "end_page": page_num,
                            "uncertain": p_date is None
                        }
                else:
                    if current_protocol:
                        current_protocol["end_page"] = page_num
                    else:
                        current_protocol = {
                            "number": "СВБ-НЕ_РАСПОЗНАН", "date": None,
                            "start_page": page_num, "end_page": page_num,
                            "uncertain": True
                        }
            
            if current_protocol:
                protocols.append(current_protocol)
                
            if not protocols:
                raise ValueError("Не найдено ни одного протокола.")
                
            self.log(f"\nПервичный парсинг завершен. Выполняется анализ целостности...")
            
            # --- ИДЕАЛЬНОЕ СЛИЯНИЕ ДВУХСТРАНИЧНЫХ ПРОТОКОЛОВ ---
            merged_protocols = []
            for prot in protocols:
                if not merged_protocols:
                    merged_protocols.append(prot)
                    continue
                    
                prev_prot = merged_protocols[-1]
                prev_len = prev_prot["end_page"] - prev_prot["start_page"] + 1
                
                # Мы можем объединить страницу, только если предыдущая часть - ровно 1 страница
                can_merge = (prev_len == 1) and (prot["start_page"] == prev_prot["end_page"] + 1)
                
                if can_merge:
                    # 1. Если текущая страница вообще не распозналась (нет даты = uncertain),
                    # это 100% вторая страница предыдущего протокола. Объединяем без вопросов.
                    if prot["uncertain"]:
                        self.log(f"[*] Слияние: стр. {prot['start_page']+1} (без даты/шапки) присоединена к протоколу {prev_prot['number']}")
                        prev_prot["end_page"] = prot["end_page"]
                        continue
                    
                    # 2. Если на текущей странице ЕСТЬ шапка (уверенная), но на предыдущей НЕ БЫЛО шапки (была неопределенная),
                    # мы берем все данные из текущей страницы, но присоединяем к ней предыдущую.
                    if prev_prot["uncertain"] and not prot["uncertain"]:
                        self.log(f"[*] Слияние: стр. {prev_prot['start_page']+1} была без шапки. Данные обновлены на: {prot['number']}")
                        prev_prot["number"] = prot["number"]
                        prev_prot["date"] = prot["date"]
                        prev_prot["end_page"] = prot["end_page"]
                        prev_prot["uncertain"] = False
                        continue
                            
                merged_protocols.append(prot)
                
            protocols = merged_protocols
            self.log(f"Протоколов после слияния: {len(protocols)}. Формирование файлов...")
            
            # Файловая система
            base_name = os.path.splitext(os.path.basename(self.file_path))[0]
            
            # Сохраняем архивы напрямую в выбранную папку
            project_output_dir = self.output_dir
            
            temp_dir = os.path.join(project_output_dir, f"{base_name}_temp")
            os.makedirs(temp_dir, exist_ok=True)
            
            records = []
            files_to_zip = []
            
            for idx, prot in enumerate(protocols, start=1):
                clean_num = prot["number"].replace("/", "_").replace("\\", "_")
                date_str = prot["date"] if prot["date"] else "Без_даты"
                clean_date = date_str.replace("/", "_").replace("\\", "_")
                
                start_p = prot["start_page"]
                end_p = prot["end_page"]
                num_p = end_p - start_p + 1
                
                page_range_str = f"{start_p+1}" if num_p == 1 else f"{start_p+1}-{end_p+1}"
                
                # ФОРМАТ: Протокол_СВБ-10-1807_26-1_от_25.07.2026_стр_27-28.pdf
                filename = f"Протокол_{clean_num}_от_{clean_date}_стр_{page_range_str}.pdf"
                filepath = os.path.join(temp_dir, filename)
                
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=start_p, to_page=end_p)
                new_doc.save(filepath)
                new_doc.close()
                files_to_zip.append(filepath)
                
                note = "уверенно" if not prot["uncertain"] else "требует ручной проверки"
                
                records.append({
                    "Номер протокола": prot["number"],
                    "Дата протокола": prot["date"] if prot["date"] else "Не найдена",
                    "Страницы в исходном PDF": page_range_str,
                    "Количество страниц": num_p,
                    "Имя файла": filename,
                    "Примечание": note
                })
                
            doc.close()
            
            df = pd.DataFrame(records)
            excel_path = os.path.join(temp_dir, "Реестр_разбивки.xlsx")
            df.to_excel(excel_path, index=False)
            files_to_zip.append(excel_path)
            
            zip_filename = os.path.join(project_output_dir, f"{base_name}_результат.zip")
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file in files_to_zip:
                    zipf.write(file, os.path.basename(file))
                    
            self.log(f"\n[✓] Готово! Архив сохранен:\n{zip_filename}")
            
        except Exception as e:
            self.log(f"\n[ОШИБКА] {str(e)}")
        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.is_processing = False
            self.after(0, lambda: self.set_buttons_state("normal"))

if __name__ == "__main__":
    app = PDFSplitterApp()
    app.mainloop()
