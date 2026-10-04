"""Уборка после живых тестов.

Приложения пользователя не трогаем: Word закрывается ТОЛЬКО если его запустили сами тесты (до прогона он не работал)
и в нём не осталось ни одного документа. Если Word был запущен до тестов — он остаётся как есть, даже пустой.
"""

import pythoncom
import win32com.client


def word_running() -> bool:
    """Запущен ли Word сейчас (не запуская его)."""
    pythoncom.CoInitialize()
    try:
        try:
            win32com.client.GetActiveObject("Word.Application")
            return True
        except pythoncom.com_error:
            return False
    finally:
        pythoncom.CoUninitialize()


def quit_word_if_idle(started_by_tests: bool = False) -> str:
    if not started_by_tests:
        return "word not started by the tests - left alone"
    pythoncom.CoInitialize()
    try:
        try:
            w = win32com.client.GetActiveObject("Word.Application")
        except pythoncom.com_error:
            return "word not running"
        if w.Documents.Count == 0:
            w.Quit(0)
            return "word quit (started by the tests and idle)"
        return f"word left running ({w.Documents.Count} documents open)"
    finally:
        pythoncom.CoUninitialize()
