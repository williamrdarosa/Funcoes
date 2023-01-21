import os
from win32com import client
import fitz

dir_path = os.path.dirname(os.path.realpath('__file__'))

def imprimir(n, nome):
    excel = client.Dispatch("Excel.Application")
    sheets = excel.Workbooks.Open(os.path.join(dir_path, f'{nome}.xlsx'))
    work_sheets = sheets.Worksheets[n]
    work_sheets.ExportAsFixedFormat(0, os.path.join(dir_path, f'{nome}.pdf'))
    sheets.Close(True)
    doc = fitz.open(os.path.join(dir_path, f'{nome}.pdf'))
    page = doc.load_page(0)
    pix = page.get_pixmap()
    pix.save(os.path.join(dir_path, f'{nome}.png'))