from docx import Document
import pandas as pd
from docx2pdf import convert
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import tkinter as tk
from tkinter import simpledialog

tk.Tk().withdraw()
location = simpledialog.askstring("Letter Encryption", "Location of the staff folder")

# location = input('Location of the staff folder: ')

# Load the Excel file with staff data
df = pd.read_excel(f'{location}/staff.xlsx')

# Loop through each row of the sheet (grabbing the fullname, bonus, salary and code cell)
for index, row in df.iterrows():
    # Each cell's data
    fullname = row['fullname']
    password = str(row['code'])
    bonus = str(row['bonus'])
    salary = str(row['salary'])

    # Fullname appended to the file format required
    file = f'{fullname}.docx'

    # Word template
    doc = Document(f'{location}/year_end.docx')

    # Go through paragraph in word template
    for paragraph in doc.paragraphs:
        # Replace templates with variables
        paragraph.text = paragraph.text.replace('{fullname}', fullname)
        paragraph.text = paragraph.text.replace('{bonus}', bonus)
        paragraph.text = paragraph.text.replace('{salary}', salary)

    # Create new Word file
    doc.save(f'{location}/{file}')

    # Convert the Word file to PDF
    convert(f'{location}/{file}')

    # Now we're going to encrypt the PDF
    pdf_file = file.replace('.docx', '.pdf')
    output_pdf_file = file.replace('.docx', '_protected.pdf')

    pdf_reader = PdfFileReader(f'{location}/{pdf_file}')

    # Read the PDF content
    pdf_writer = PdfFileWriter()
    for page_num in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page_num)
        pdf_writer.addPage(page)

    with open(f'{location}/{output_pdf_file}', 'wb') as output_pdf:
        pdf_writer.encrypt(password)
        pdf_writer.write(output_pdf)

    # Delete the Word and non-encrypted PDF file once encrypted file is created
    if os.path.exists(f'{location}/{fullname}.pdf') and os.path.exists(f'{location}/{file}'):
        os.remove(f'{location}/{fullname}.pdf')
        os.remove(f'{location}/{file}')
