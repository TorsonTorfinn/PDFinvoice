import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Находим все файлы Excel в папке "invoices"
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Читаем данные из первого листа Excel файла
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Создаем PDF-документ
    pdf = FPDF(orientation="portrait", unit="mm", format='A4')
    pdf.add_page()

    # Извлекаем номер счета из имени файла
    filename = Path(filepath).stem
    invoice_number = filename.split('-')[0]
    invoice_date = filename.split('-')[1]

    # Добавляем текст в PDF
    pdf.set_font(family='Times', size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_number}", ln=True)

    pdf.set_font(family='Times', size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}")

    # Сохраняем PDF-файл
    pdf.output(f"PDFs/{filename}.pdf")
