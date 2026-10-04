# Поднимает видимые Excel (2 черновые книги) и Word (1 документ) для stdio-теста сервера.
import sys

import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
scratch = excel.Workbooks.Add()
deliverable = excel.Workbooks.Add()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True
doc = word.Documents.Add()
print("EXCEL_SCRATCH:" + scratch.Name)
print("EXCEL_DELIV:" + deliverable.Name)
print("WORD_SCRATCH:" + doc.Name)
sys.stdout.flush()
