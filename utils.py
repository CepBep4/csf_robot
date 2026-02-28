import openpyxl
from pyautogui import hotkey

try:
    import pyperclip
except ImportError:
    pyperclip = None


def load_from_xlsx(path: str) -> list[dict]:
    """
    Загружает xlsx-файл и возвращает список словарей.
    Первая строка — названия столбцов (ключи), каждая следующая строка — один словарь.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return []

    headers = [str(h) if h is not None else "" for h in rows[0]]
    data = []
    for row in rows[1:]:
        data.append(dict(zip(headers, (v if v is not None else "" for v in row))))

    return data

def addToBuffer(text: str) -> None:
    """
    Копирует переданный текст в буфер обмена.
    Требует установленного pyperclip и доступного системного буфера обмена.
    """
    if pyperclip is None:
        raise RuntimeError("pyperclip не установлен. Добавьте его в зависимости и установите.")

    pyperclip.copy(str(text))
    
def getFromBuffer() -> str:
    """
    Возвращает текст из буфера обмена.
    Требует установленного pyperclip и доступного системного буфера обмена.
    """
    if pyperclip is None:
        raise RuntimeError("pyperclip не установлен. Добавьте его в зависимости и установите.")

    return pyperclip.paste()

def selectAll() -> None:
    hotkey("ctrl", "a")
    
def copy() -> None:
    """
    Нажимает Ctrl+C в активном окне для копирования выделенного текста.
    """
    hotkey("ctrl", "c")

def paste() -> None:
    """
    Нажимает Ctrl+V в активном окне для вставки текста из буфера обмена.
    """
    hotkey("ctrl", "v")
    