# Уборка после stdio-теста: закрыть Word-документ scratch и выйти из Word;
# в Excel закрыть scratch-книги, но ОСТАВИТЬ открытой ZCode_MCP_Test.xlsx.
import pythoncom
import win32com.client

pythoncom.CoInitialize()
try:
    try:
        word = win32com.client.GetActiveObject("Word.Application")
        for i in range(word.Documents.Count, 0, -1):
            d = word.Documents(i)
            if d.Name.startswith("scratch"):
                d.Close(SaveChanges=0)
        word.Quit()
        print("word: closed+quit")
    except pythoncom.com_error:
        print("word: not running")
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        for i in range(excel.Workbooks.Count, 0, -1):
            wb = excel.Workbooks(i)
            if wb.Name.startswith("scratch") or wb.Name.startswith("Книга"):
                wb.Close(SaveChanges=0)
        names = [excel.Workbooks(j).Name for j in range(1, excel.Workbooks.Count + 1)]
        print("excel: cleaned, open now:", names)
    except pythoncom.com_error:
        print("excel: not running")
finally:
    pythoncom.CoUninitialize()
