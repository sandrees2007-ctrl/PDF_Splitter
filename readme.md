# StroiLab Intelligence Suite — AI PDF Protocol Splitter & Registry Integrator

[🇷🇺 Русский](#-русский) | [🇬🇧 English](#-english)

---

## 🇷🇺 Русский

**StroiLab Intelligence Suite (PDF Splitter)** — это веб-приложение на базе Flask для автоматического разделения многостраничных PDF-файлов с протоколами испытаний бетона (ООО «СтройЛабГрупп»), интеллектуального распознавания реквизитов с помощью мультимодальных ИИ-моделей (Vision + Text) и автоматической интеграции данных в реестры Excel или Google Таблицы.

### ✨ Основные возможности

* **Гибридный ИИ-анализ (Text + Vision):** Автоматически извлекает текстовый слой PDF (через PyMuPDF) или рендерит страницы в JPEG для анализа через LLM/Vision API (например, `google/gemma-4-26b-a4b-it` через ProxyAPI).
* **Умная коррекция ошибок OCR и печатей:** Восстанавливает номера протоколов формата `СВБ-{префикс}-{ДДММ}_{ГГ}-{серия}`, исправляя искажения от синих печатей путем сверки и математического пересчета даты бетонирования, возраста бетона (7 или 28 суток) и даты проведения испытаний.
* **Автоматическое слияние и сортировка:** Объединяет двухстраничные протоколы, переименовывает файлы по шаблону `Протокол_СВБ-..._от_ДД.ММ.ГГГГ.pdf`, раскладывает их по папкам корпусов и типов испытаний (`7_28`, `10`, `12`) и упаковывает в единый ZIP-архив.
* **Интеграция с реестрами:**
  * Формирует промежуточный форматированный реестр `Реестр_разбивки.xlsx`.
  * Поддерживает интеллектуальное сопоставление (fuzzy matching по дате, корпусу и конструкции) и автозаполнение **локальных Excel-файлов** (`.xlsx`) или **Google Таблиц** (через Google Sheets API) с добавлением гиперссылок и цветовой индикацией.
* **Асинхронная обработка и стриминг логов:** Многопоточный пул воркеров (`ThreadPoolExecutor`) для параллельной отправки страниц в ИИ и трансляция прогресса в реальном времени через Server-Sent Events (SSE).

### 🛠 Технологический стек

* **Backend:** Python 3.10+, Flask 3.1
* **PDF & Изображения:** PyMuPDF (`fitz`), Pillow
* **Данные и Таблицы:** Pandas, OpenPyXL, `gspread`, `oauth2client`
* **ИИ и Сеть:** REST API (совместимый с OpenAI / ProxyAPI), `requests` (с пулом соединений), SSE

### 📂 Структура проекта

```text
PDF_Splitter/
├── app.py                  # Основной сервер Flask, REST API и бизнес-логика обработки
├── requirements.txt        # Список зависимостей Python
├── settings.example.json   # Шаблон файла конфигурации
├── templates/
│   └── index.html          # Веб-интерфейс приложения
├── static/                 # Статические ресурсы (CSS/JS)
├── uploads/                # Директория для загружаемых PDF и реестров
└── output/                 # Результаты: папки по корпусам, Excel-реестры и ZIP-архив
```

### 🚀 Установка и запуск

1. **Клонируйте репозиторий:**
   ```bash
   git clone https://github.com/sandrees2007-ctrl/PDF_Splitter.git
   cd PDF_Splitter
   ```

2. **Установите зависимости:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Настройте конфигурацию:**
   Скопируйте `settings.example.json` в `settings.json` (или настройте параметры позже прямо в веб-интерфейсе):
   ```bash
   cp settings.example.json settings.json
   ```
   *Для работы с Google Таблицами поместите файл сервисного аккаунта `credentials.json` в корень проекта или загрузите его через интерфейс.*

4. **Запустите сервер:**
   ```bash
   python app.py
   ```
   После запуска откройте в браузере: `http://127.0.0.1:5000`

> ⚠️ **Примечание по безопасности:** Не публикуйте реальные API-ключи (`ai_api_key` в `settings.json`) и ключи сервисного аккаунта (`credentials.json`) в публичных репозиториях. Убедитесь, что оба файла добавлены в `.gitignore`.

---

## 🇬🇧 English

**StroiLab Intelligence Suite (PDF Splitter)** is a Flask-based web application designed to automatically split multi-page concrete testing protocol PDFs, extract metadata using multimodal AI models (Vision + Text), and seamlessly integrate the parsed records into local Excel registries or Google Sheets.

### ✨ Key Features

* **Hybrid AI Recognition (Text + Vision):** Extracts native PDF text layers via PyMuPDF or renders scanned pages into compressed JPEGs for LLM Vision processing (e.g., `google/gemma-4-26b-a4b-it` via ProxyAPI).
* **Smart Stamp & OCR Error Correction:** Normalizes protocol numbers (`СВБ-{prefix}-{DDMM}_{YY}-{series}`) and resolves digits obscured by laboratory stamps by cross-referencing and mathematically validating concreting dates, concrete curing age (7 or 28 days), and testing dates.
* **Automatic Merging & File Organization:** Detects and merges multi-page (2-page) protocols, renames files systematically, organizes them into structured directories by building and test category, and bundles everything into a downloadable ZIP archive.
* **Registry Integration (Excel & Google Sheets):**
  * Generates a formatted summary spreadsheet (`Реестр_разбивки.xlsx`).
  * Performs fuzzy matching (by date, building, and structure name) to update existing rows or append new entries into **local `.xlsx` workbooks** or **Google Sheets** via API, complete with hyperlinks and cell highlighting.
* **Concurrent Processing & Real-Time Logs:** Uses `ThreadPoolExecutor` for parallel page analysis and streams live progress bars and execution logs to the browser via Server-Sent Events (SSE).

### 🛠 Tech Stack

* **Backend:** Python 3.10+, Flask 3.1
* **PDF & Imaging:** PyMuPDF (`fitz`), Pillow
* **Data & Spreadsheets:** Pandas, OpenPyXL, `gspread`, `oauth2client`
* **AI & Networking:** OpenAI-compatible REST API (ProxyAPI), `requests` with connection pooling, SSE

### 📂 Project Structure

```text
PDF_Splitter/
├── app.py                  # Flask server, REST API endpoints, and core processing logic
├── requirements.txt        # Python dependencies
├── settings.example.json   # Configuration template
├── templates/
│   └── index.html          # Frontend web UI
├── static/                 # Static assets (CSS/JS)
├── uploads/                # Temporary storage for uploaded PDFs and Excel templates
└── output/                 # Generated folders, Excel registries, and ZIP archives
```

### 🚀 Installation & Quick Start

1. **Clone the repository:**
   ```bash
   git clone https://github.com/sandrees2007-ctrl/PDF_Splitter.git
   cd PDF_Splitter
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure settings:**
   Copy `settings.example.json` to `settings.json` (or configure everything directly through the web UI):
   ```bash
   cp settings.example.json settings.json
   ```
   *To enable Google Sheets integration, place your Google Service Account `credentials.json` in the project root or upload it via the UI.*

4. **Run the application:**
   ```bash
   python app.py
   ```
   Navigate to `http://127.0.0.1:5000` in your web browser.

> ⚠️ **Security Note:** Never commit sensitive credentials (`settings.json` containing your API key or `credentials.json`) to version control. Ensure both files remain listed in `.gitignore`.