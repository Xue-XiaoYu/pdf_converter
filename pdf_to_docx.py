import os
import win32com.client
import img2pdf
import fitz
from PIL import Image
from pdf2docx import extract_tables as tbl
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger


# Word Document to PDF
def docx_to_pdf(input, output):
    wdFormatPDF = 17

    in_file = os.path.abspath(input)
    out_file = os.path.abspath(output)

    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


# Powerpoint to PDF
def pptx_to_pdf(input, output):
    wdFormatPDF = 32

    in_file = os.path.abspath(input)
    out_file = os.path.abspath(output)

    word = win32com.client.Dispatch('Powerpoint.Application')
    doc = word.Presentations.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


# Image to PDF
def img_to_pdf(input, output):
    image = Image.open(input)
    pdf_bytes = img2pdf.convert(image.filename)
    file = open(output, "wb")
    file.write(pdf_bytes)
    image.close()
    file.close()


# PDF to Image
def pdf_to_img(input, output_file):
    pdf_file = input
    doc = fitz.open(pdf_file)
    page = doc.loadPage(0)  # number of page
    pix = page.getPixmap()
    output = output_file
    pix.writePNG(output)


# Extract Tables from PDF
def extract_pdf_tables(input_file):
    pdf_file = input_file

    tables = tbl(pdf_file, start=0, end=1)
    for table in tables:
        print(table)


# Merge PDF
def merge_pdf(input_dir, output_file):
    merger = PdfFileMerger()
    for file in os.listdir(input_dir):
        if file.endswith('.pdf'):
            merger.append(PdfFileReader(file, 'rb'))

    merger.write(output_file)


# Split PDF files into single pages
def split_pdf(input_file):
    file_name = os.path.basename(input_file).replace('.pdf', '')
    file_path = os.path.dirname(input_file)
    with open(input_file, 'rb') as org_file:
        pdf_reader = PdfFileReader(org_file, strict=False)
        num_of_pages = pdf_reader.getNumPages()
        for i in range(num_of_pages):
            writer = PdfFileWriter()
            writer.addPage(pdf_reader.getPage(i))
            with open(f'{file_path}/{file_name}_p{i + 1}.pdf', 'wb') as outfile:
                writer.write(outfile)


# Encrypt PDF with password
def encrypt_pdf(input_file, password):

    if not input_file.endswith('.pdf'):
        print(f'This file {input_file} is not supported for encryption')
        return

    out = PdfFileWriter()
    file = PdfFileReader(input_file)
    file_name = os.path.basename(input_file).split('.')[0]
    num = file.numPages
    for idx in range(num):
        page = file.getPage(idx)
        out.addPage(page)
    out.encrypt(password)
    with open(f"{file_name}_encrypted.pdf", "wb") as f:
        out.write(f)


if __name__ == '__main__':
    input_file = 'data/SOB_ALLIANZ_IP_ocr.pdf'
    encrypt_pdf(input_file, 'meow')

