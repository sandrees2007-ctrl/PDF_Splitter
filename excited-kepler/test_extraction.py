import os
import re
import sys

# Import pdfplumber using the same path/env the user has
try:
    import pdfplumber
except ImportError:
    print("pdfplumber not found, using generic mock test.")
    sys.exit(0)

from main import PDFDataExtractor

examples_dir = os.path.join(os.path.dirname(__file__), "examples")
if not os.path.exists(examples_dir):
    print("No examples directory.")
    sys.exit(0)

pdfs = [os.path.join(examples_dir, f) for f in os.listdir(examples_dir) if f.endswith(".pdf")]

print(f"Found {len(pdfs)} PDFs in examples.")

extractor = PDFDataExtractor(pdfs, lambda x: None, lambda x: None)
records = extractor.process_all()

print("\n--- RESULTS ---")
for r in records:
    print(f"File: {r.get('Имя файла')}")
    print(f"  Протокол: {r.get('Номер протокола')}")
    print(f"  Дата бетонирования: {r.get('Дата бетонирования')}")
    print(f"  Корпус: {r.get('Корпус')}")
    print(f"  Конструкция: {r.get('Наименование конструкции')}")
    print("-" * 40)
