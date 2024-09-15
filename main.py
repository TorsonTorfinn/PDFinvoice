import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Находим все файлы Excel в папке "invoices"
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Создаем PDF-документ
    pdf = FPDF(orientation="portrait", unit="mm", format='A4')
    pdf.add_page()

    # Извлекаем номер счета из имени файла
    filename = Path(filepath).stem
    invoice_number = filename.split('-')[0]
    invoice_date = filename.split('-')[1]

    # Добавляем текст в PDF
    pdf.set_font(family='Courier', size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_number}", ln=True)

    # добавляем дату в PDF
    pdf.set_font(family='Courier', size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=True)

    # Читаем данные из первого листа Excel файла
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # дОбавляем заголовок 
    columns_df = df.columns
    columns_df = [
        str(item).replace('_', ' ').upper() for item in columns_df
    ]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=columns_df[0], border=True)
    pdf.cell(w=70, h=8, txt=columns_df[1], border=True)
    pdf.cell(w=30, h=8, txt=columns_df[2], border=True)
    pdf.cell(w=30, h=8, txt=columns_df[3], border=True)
    pdf.cell(w=30, h=8, txt=columns_df[4], border=True, ln=True)

    # add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family='Courier', size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=True)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=True)
        pdf.cell(w=30, h=8, txt=str(row['purchased']), border=True)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=True)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=True, ln=1)

    # total sum row
    total_sum = df['total_price'].sum()
    pdf.set_font(family='Courier', size=12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt="", border=True)
    pdf.cell(w=70, h=8, txt="", border=True)
    pdf.cell(w=30, h=8, txt="", border=True)
    pdf.cell(w=30, h=8, txt="", border=True)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=True, ln=1)
    
    # total sum sentence
    pdf.set_font(family='Courier', size=12, style='B')
    pdf.cell(w=30, h=8, txt=f"The total price is ${total_sum} in USD", ln=1)

    # name and logo just for beauty
    pdf.set_font(family='Courier', size=10, style='B')
    pdf.cell(w=30, h=8, txt=f"MadeWithPython")
    pdf.image("images/pythonhow.png", w=15)

    # Сохраняем PDF-файл
    pdf.output(f"PDFs/{filename}.pdf")
