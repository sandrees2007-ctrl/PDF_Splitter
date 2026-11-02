import requests
import base64
import json
import io
from PIL import Image, ImageDraw

def test():
    img = Image.new('RGB', (800, 350), color='white')
    draw = ImageDraw.Draw(img)
    text_lines = [
        "ПРОТОКОЛ ИСПЫТАНИЙ № 10-1807/26-1 от 25.07.2026",
        "Объект: Жилой комплекс, Корпус 3",
        "Конструкция: Плита перекрытия на отм. +12.000 в осях 1-5/А-Г",
        "Дата бетонирования: 27.06.2026",
        "Возраст бетона: в возрасте 28 сут",
        "Метод испытания: ГОСТ 22690 (неразрушающий контроль)"
    ]
    y = 30
    for line in text_lines:
        draw.text((30, y), line, fill='black')
        y += 45

    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=90)
    b64 = base64.b64encode(buf.getvalue()).decode('utf-8')

    prompt = """Ты специализированный эксперт по анализу строительных протоколов испытания бетона (ООО «СтройЛабГрупп»).
Извлеки данные строго в формате JSON:
{
  "is_protocol": true,
  "protocol_number": "СВБ-28-0107_26-2",
  "protocol_date": "29.07.2026",
  "building": "корпус 4",
  "structure": "наименование конструкции/элемента",
  "concreting_date": "01.07.2026",
  "age_days": "28",
  "is_cube": true,
  "is_continuation": false
}

КРИТИЧЕСКИ ВАЖНЫЕ ПРАВИЛА И СТРУКТУРА:
1. СТРУКТУРА НОМЕРА: СВБ-{возраст}-{ДДММ}_{ГГ}-{серия} (все слэши замени на _).
   - Префикс: 28 или 7 (для кубиков), 10 или 12 (для неразрушающего контроля).
   - {ДДММ}: 4 цифры — это СТРОГО день и месяц даты бетонирования/изготовления (concreting_date)!
   - {ГГ}: 2 последние цифры года (2026 -> 26).
   - {серия}: номер партии (-1, -2, -3).
2. ВНИМАНИЕ: СИНЯЯ КРУГЛАЯ ПЕЧАТЬ В ШАПКЕ ПЕРЕКРЫВАЕТ НОМЕР!
   Круглая печать лаборатории часто наложена прямо на номер протокола, из-за чего цифры под печатью искажаются: например «0107» кажется «8187», «0307» кажется «03706», «2607» кажется «2697».
   НИКОГДА не пытайся читать размытые цифры только из шапки!
   ОБЯЗАТЕЛЬНО найди в таблице результатов испытаний (раздел 11) колонку «Дата изготовления» (или строку «Дата бетонирования»).
   Если дата изготовления 01.07.2026 — в номере протокола ОБЯЗАНО быть 0107 (а не 8187)! Если 26.07.2026 — 2607, если 03.07.2026 — 0307.
3. МАТЕМАТИЧЕСКАЯ СВЯЗЬ ДАТ:
   Дата протокола (проведения испытаний) = Дата изготовления + Возраст (суток).
   - 01.07.2026 + 28 сут = 29.07.2026.
   - 01.07.2026 + 7 сут = 08.07.2026.
   - 26.07.2026 + 7 сут = 02.08.2026.
   Используй это равенство для проверки и самокоррекции дат!
4. is_cube: true если в таблице размер образцов 10х10х10 см / 100х100х100 мм (ГОСТ 10180), иначе false (ГОСТ 22690 / неразрушающий).
5. Ответь ИСКЛЮЧИТЕЛЬНО валидным JSON без каких-либо markdown блоков."""

    api_key = "sk-W02LrjSK9PapzakiwgjkP0IBkFN7p3rE"
    url = "https://api.proxyapi.ru/v1/chat/completions"

    payload = {
        "model": "google/gemma-4-26b-a4b-it",
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": "data:image/jpeg;base64," + b64}}
                ]
            }
        ],
        "temperature": 0.1
    }

    print("Sending request to ProxyAPI...")
    r = requests.post(url, headers={"Authorization": "Bearer " + api_key}, json=payload, timeout=30)
    print("HTTP Status:", r.status_code)
    if r.status_code != 200:
        print("Error response:", r.text)
        return

    content = r.json()["choices"][0]["message"]["content"].strip()
    print("Raw Content:\n", content)
    if content.startswith("```"):
        lines = content.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        content = "\n".join(lines).strip()

    data = json.loads(content)
    print("\nParsed Result:")
    for k, v in data.items():
        print(f"  {k}: {v}")

if __name__ == "__main__":
    test()
