import numpy as np
import pandas as pd
import PyPDF4 as pypdf
import tabula as tbl
import re

# Path to PDF file
pdf_path = r'C:\Users\Navneeth\Documents\WORK\BankSME\Programming\banksme\OCK.pdf'
#ock_bs = tbl.read_pdf(pdf_path, pages="12")
pdf_obj = pypdf.PdfFileReader(pdf_path)
num_pages = pdf_obj.getNumPages()

bs_page = np.nan
is_page = np.nan
cs_page = np.nan

balance_sheet = tbl.read_pdf(pdf_path, pages="12")
print(balance_sheet)

#Search where tables are
# for i in range(0, num_pages):
#     curr_page = pdf_obj.getPage(i)
#     text = curr_page.extractText()
#     text = text.casefold()
#     match = 'balance sheet' in text
#     if match:
#         print("found balance sheet page")
#         balance_sheet = tbl.read_pdf(pdf_path, pages=str(i))
#         print(balance_sheet)
#         break
#     else:
#         print("not balance sheet page, {}".format(str(i)))