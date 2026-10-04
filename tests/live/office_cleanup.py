"""Уборка после живых тестов: закрывает Word только если он пуст (ни одного документа) — документы пользователя не трогаем."""
import pythoncom, win32com.client

def quit_word_if_idle():
    pythoncom.CoInitialize()
    try:
        try:
            w = win32com.client.GetActiveObject("Word.Application")
        except pythoncom.com_error:
            return "word not running"
        if w.Documents.Count == 0:
            w.Quit(0)
            return "word quit (was idle)"
        return f"word left running ({w.Documents.Count} documents open)"
    finally:
        pythoncom.CoUninitialize()
