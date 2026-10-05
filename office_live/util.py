"""Чистые функции без COM: адреса A1, цвета, нормализация значений, разбор текста."""

import datetime
import decimal
import re

from .errors import ToolError

# ------------------------------------------------------------------ адреса A1

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_COL_RE = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW_RE = re.compile(r"^\$?(\d+)$")
MAX_ROWS = 1048576
MAX_COLS = 16384


def col_letter(n: int) -> str:
    if n < 1 or n > MAX_COLS:
        raise ValueError(f"column {n} out of range")
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        if not "A" <= ch <= "Z":
            raise ValueError(f"bad column {letters!r}")
        n = n * 26 + ord(ch) - 64
    return n


def a1_cell(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


def a1_range(r1: int, c1: int, r2: int, c2: int) -> str:
    if (r1, c1) == (r2, c2):
        return a1_cell(r1, c1)
    return f"{a1_cell(r1, c1)}:{a1_cell(r2, c2)}"


def parse_a1(addr: str) -> tuple[int, int, int, int]:
    """'B2', 'B2:D5', 'A:C', '3:7' -> (r1, c1, r2, c2). Одна область, без имён листов."""
    a = addr.replace("$", "").strip()
    if ":" in a:
        left, right = a.split(":", 1)
    else:
        left = right = a
    m1, m2 = _CELL_RE.match(left), _CELL_RE.match(right)
    if m1 and m2:
        r1, c1 = int(m1.group(2)), col_number(m1.group(1))
        r2, c2 = int(m2.group(2)), col_number(m2.group(1))
    elif _COL_RE.match(left) and _COL_RE.match(right):
        r1, r2 = 1, MAX_ROWS
        c1, c2 = col_number(left), col_number(right)
    elif _ROW_RE.match(left) and _ROW_RE.match(right):
        r1, r2 = int(left), int(right)
        c1, c2 = 1, MAX_COLS
    else:
        raise ValueError(f"cannot parse address {addr!r}")
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def split_sheet_ref(ref: str) -> tuple[str | None, str]:
    """"'Мой лист'!A1:B2" -> ("Мой лист", "A1:B2"); "A1" -> (None, "A1")."""
    ref = ref.strip()
    if "!" not in ref:
        return None, ref
    sheet, addr = ref.rsplit("!", 1)
    sheet = sheet.strip()
    if len(sheet) >= 2 and sheet[0] == "'" and sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, addr.strip()


def quote_sheet(name: str) -> str:
    if re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def parse_row_spec(spec: str, limit: int | None = None) -> list[int]:
    """'3-5,8' -> [3,4,5,8] (1-based)."""
    out: list[int] = []
    for part in spec.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            lo, hi = int(a), int(b)
            if hi < lo:
                lo, hi = hi, lo
            out.extend(range(lo, hi + 1))
        else:
            out.append(int(part))
        if limit and len(out) > limit:
            raise ToolError(f"Too many items in '{spec}' (limit {limit}).")
    return out


_PH = set("0#?")  # заменители цифр в коде формата
_SPACES = ("\u00a0", "\u202f", " ")


def _split_literals(fmt: str):
    """Код формата -> список (текст, is_literal): кавычки, [скобки] и экранированные символы (\\x) не трогаем."""
    parts, buf, i = [], "", 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            j = fmt.find('"', i + 1)
            j = len(fmt) - 1 if j == -1 else j
            if buf:
                parts.append((buf, False))
                buf = ""
            parts.append((fmt[i:j + 1], True))
            i = j + 1
        elif ch == "[":
            j = fmt.find("]", i + 1)
            j = len(fmt) - 1 if j == -1 else j
            if buf:
                parts.append((buf, False))
                buf = ""
            parts.append((fmt[i:j + 1], True))
            i = j + 1
        elif ch == "\\" and i + 1 < len(fmt):
            if buf:
                parts.append((buf, False))
                buf = ""
            parts.append((fmt[i:i + 2], True))
            i += 2
        else:
            buf += ch
            i += 1
    if buf:
        parts.append((buf, False))
    return parts


def localize_number_format(fmt: str, decimal: str, thousands: str) -> str:
    """Инвариантный код формата ('#,##0.00') -> запись региональных настроек ('# ##0,00').

    Нужен потому, что через COM Excel принимает NumberFormat в ЛОКАЛЬНОЙ записи. Точки в масках дат ('dd.mm.yyyy')
    не трогаем: десятичной считается точка между заменителями цифр. Пробел-литерал после цифр в локальной записи
    стал бы разделителем тысяч (деление на 1000), поэтому экранируем его.
    """
    if decimal == "." and thousands == ",":
        return fmt
    out = []
    for text, literal in _split_literals(fmt):
        if literal:
            out.append(text)
            continue
        chars = list(text)
        for i, ch in enumerate(chars):
            prev = chars[i - 1] if i else ""
            nxt = chars[i + 1] if i + 1 < len(chars) else ""
            if ch == "." and prev in _PH and nxt in _PH:
                out.append(decimal)
            elif ch == "," and prev in _PH:
                out.append(thousands)  # между цифрами — разделитель групп, после цифр — масштаб на 1000
            elif ch == " " and prev in _PH:
                out.append("\\ ")  # литерал-пробел после цифр
            else:
                out.append(ch)
    return "".join(out)


def delocalize_number_format(fmt: str, decimal: str, thousands: str) -> str:
    """Обратное преобразование: локальная запись ('# ##0,00') -> инвариантная ('#,##0.00')."""
    if decimal == "." and thousands == ",":
        return fmt
    out = []
    for text, literal in _split_literals(fmt):
        if literal:
            out.append(" " if text == "\\ " else text)
            continue
        chars = list(text)
        for i, ch in enumerate(chars):
            prev = chars[i - 1] if i else ""
            nxt = chars[i + 1] if i + 1 < len(chars) else ""
            if ch == decimal and (prev in _PH or nxt in _PH):
                out.append(".")
            elif (ch == thousands or ch in _SPACES) and prev in _PH:
                out.append(",")
            else:
                out.append(ch)
    return "".join(out)


def cm_to_points(cm: float) -> float:
    """Application.CentimetersToPoints в late-binding Word падает (0x80004005) — считаем сами."""
    return float(cm) * 72.0 / 2.54


# ------------------------------------------------------------------ цвета

NAMED_COLORS = {
    "black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "green": "#00B050",
    "blue": "#0070C0", "yellow": "#FFFF00", "orange": "#FFC000", "purple": "#7030A0",
    "pink": "#FF99CC", "gray": "#808080", "grey": "#808080", "lightgray": "#D9D9D9",
    "lightgrey": "#D9D9D9", "darkgray": "#404040", "lightblue": "#DDEBF7", "lightgreen": "#E2EFDA",
    "lightyellow": "#FFF2CC", "lightred": "#FFC7CE", "lightorange": "#FCE4D6", "lightpurple": "#E4DFEC",
    "darkblue": "#1F4E78", "darkgreen": "#375623", "darkred": "#C00000", "brown": "#833C0B",
    "teal": "#008080", "navy": "#000080", "gold": "#FFD700", "cyan": "#00B0F0",
}


def parse_color(value: str) -> int:
    """'#RRGGBB' | 'RRGGBB' | '#RGB' | имя -> COLORREF (0x00BBGGRR) для Excel/Word."""
    if value is None:
        raise ToolError("Color is empty.")
    s = str(value).strip().lower()
    s = NAMED_COLORS.get(s, s).lstrip("#").lower()
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    if len(s) != 6 or not re.fullmatch(r"[0-9a-f]{6}", s):
        raise ToolError(f"Color must be '#RRGGBB' or a known name ({sorted(NAMED_COLORS)[:8]}...), got: {value!r}")
    r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    return b * 65536 + g * 256 + r


def color_to_hex(colorref) -> str | None:
    if colorref is None or isinstance(colorref, bool):
        return None
    try:
        v = int(colorref)
    except (TypeError, ValueError):
        return None
    if v < 0:
        return None
    return f"#{v & 0xFF:02X}{(v >> 8) & 0xFF:02X}{(v >> 16) & 0xFF:02X}"


# ------------------------------------------------------------------ значения Excel


def _error_map() -> dict[int, str]:
    names = {
        2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?",
        2036: "#NUM!", 2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!", 2046: "#CONNECT!",
        2047: "#BLOCKED!", 2048: "#UNKNOWN!", 2049: "#FIELD!", 2050: "#CALC!",
    }
    out = {}
    for code, text in names.items():
        scode = 0x800A0000 + code
        out[((scode ^ 0x80000000) - 0x80000000)] = text
    return out


EXCEL_ERRORS = _error_map()


def norm_value(v):
    """Значение ячейки -> JSON-совместимое."""
    if v is None or isinstance(v, (bool, str)):
        return v
    if isinstance(v, datetime.datetime):
        if v.year == 1899 and v.month == 12 and v.day == 30:  # «чистое» время Excel
            return v.strftime("%H:%M:%S")
        if v.hour == v.minute == v.second == 0 and v.microsecond == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, decimal.Decimal):
        v = float(v)
    if isinstance(v, int):
        return EXCEL_ERRORS.get(v, v)
    if isinstance(v, float):
        if v != v or v in (float("inf"), float("-inf")):
            return None
        if v == int(v) and abs(v) < 1e15:
            return int(v)
        return v
    if isinstance(v, (tuple, list)):
        return [norm_value(x) for x in v]
    return str(v)


def to_grid(value) -> list[list]:
    """Range.Value (скаляр | кортеж | кортеж кортежей) -> list[list]."""
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [[norm_value(c) for c in row] for row in value]
        return [[norm_value(c) for c in value]]
    return [[norm_value(value)]]


def to_com_grid(values, pad=None) -> tuple:
    """list[list] | list | скаляр -> tuple[tuple] для Range.Value. Короткие строки дополняются."""
    if not isinstance(values, (list, tuple)):
        values = [[values]]
    elif not values:
        raise ToolError("'values' is empty.")
    elif not isinstance(values[0], (list, tuple)):
        values = [list(values)]
    width = max(len(r) if isinstance(r, (list, tuple)) else 1 for r in values)
    rows = []
    for r in values:
        r = list(r) if isinstance(r, (list, tuple)) else [r]
        r += [pad] * (width - len(r))
        rows.append(tuple(r))
    return tuple(rows)


# ------------------------------------------------------------------ текст/числа

_NUM_PLAIN = re.compile(r"^[+-]?\d+(\.\d+)?$")


_SPACE_GROUPED = re.compile(r"([1-9]\d{0,2}(?: \d{3})+)(?:[.,](\d+))?")
_DOT_GROUPED = re.compile(r"[1-9]\d{0,2}(?:\.\d{3})+")
_COMMA_GROUPED = re.compile(r"[1-9]\d{0,2}(?:,\d{3})+")


def _plain_digits(core: str, decimal_sep: str) -> str | None:
    """Тело числа без знака -> строка вида '1234.5' или None, если запись не похожа на число.

    Допустимы только осмысленные записи: целое/десятичное; тысячи группами строго по 3 цифры (пробел, запятая, точка);
    смесь '.' и ',' — только в виде 1,234.56 / 1.234,56. Всё остальное (12.03.2026, 1.2.3, 1,23,4, 12 34) — не число.
    """
    if " " in core:
        m = _SPACE_GROUPED.fullmatch(core)
        return None if not m else m.group(1).replace(" ", "") + ("." + m.group(2) if m.group(2) else "")
    has_dot, has_comma = "." in core, "," in core
    if has_dot and has_comma:
        dec = "." if core.rfind(".") > core.rfind(",") else ","
        head, _, frac = core.rpartition(dec)
        grouped = _COMMA_GROUPED if dec == "." else _DOT_GROUPED
        if not frac.isdigit() or not grouped.fullmatch(head):
            return None
        return head.replace("," if dec == "." else ".", "") + "." + frac
    if has_dot or has_comma:
        sep = "," if has_comma else "."
        grouped = _COMMA_GROUPED if sep == "," else _DOT_GROUPED
        if core.count(sep) > 1:
            return core.replace(sep, "") if grouped.fullmatch(core) else None  # 1,234,567 — тысячи; 1.2.3 — не число
        if sep != decimal_sep and grouped.fullmatch(core):
            return core.replace(sep, "")  # 1,234 при десятичной точке -> тысячи
        head, _, frac = core.partition(sep)
        return head + "." + frac if head.isdigit() and frac.isdigit() else None
    return core if core.isdigit() else None


def smart_number(text: str, decimal_sep: str = ",") -> float | int | None:
    """Строку вида '1 234,56' | '1,234.56' | '12%' -> число; иначе None. Идентификаторы, даты и номера разделов остаются текстом."""
    if text is None:
        return None
    s = str(text).strip().replace("\u00a0", " ").replace("\u202f", " ")
    if not s:
        return None
    if s.startswith("+") and " " in s:
        return None  # «+7 999 123 45 67» — телефон, а не число с разделителями тысяч
    pct = s.endswith("%")
    if pct:
        s = s[:-1].strip()
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()
    sign = ""
    if s and s[0] in "+-":
        sign, s = s[0], s[1:].strip()
    if not s or not re.fullmatch(r"[\d., ]+", s):
        return None
    if re.match(r"^0\d", s):
        return None  # 007, 0123, 00,5 — коды, не числа ('0,5' и '0.5' допустимы)
    plain = _plain_digits(s, decimal_sep)
    if plain is None:
        return None
    if len(plain.replace(".", "")) > 15:
        return None  # точность double исчерпана — оставляем текстом
    num = float(plain)
    if sign == "-":
        num = -num
    if neg:
        num = -num
    if pct:
        return num / 100
    return int(num) if "." not in plain else num


DATE_FORMATS = {"DD.MM.YYYY": "%d.%m.%Y", "DD/MM/YYYY": "%d/%m/%Y", "MM/DD/YYYY": "%m/%d/%Y",
                "YYYY-MM-DD": "%Y-%m-%d", "DD-MM-YYYY": "%d-%m-%Y", "YYYY.MM.DD": "%Y.%m.%d"}
DATE_FORMATS.update({base + suffix: fmt + code for base, fmt in list(DATE_FORMATS.items())
                     for suffix, code in ((" HH:MM", " %H:%M"), (" HH:MM:SS", " %H:%M:%S"))})


def excel_date_serial(text, date_format, date1904=False):
    """Strict fixed-width date, with Excel's 1900 leap-year bug and 1904 epoch."""
    pattern = re.sub(r"YYYY|DD|MM|HH|SS", lambda m: r"[0-9]{4}" if m[0] == "YYYY" else r"[0-9]{2}", re.escape(date_format))
    if not re.fullmatch(pattern, text):
        return None
    try:
        value = datetime.datetime.strptime(text, DATE_FORMATS[date_format])
    except ValueError:
        return None
    epoch = datetime.datetime(1904, 1, 1) if date1904 else datetime.datetime(1899, 12, 31)
    if value < (epoch if date1904 else datetime.datetime(1900, 1, 1)):
        return None
    days = (value - epoch).days + int(not date1904 and value >= datetime.datetime(1900, 3, 1))
    return days + (value.hour * 3600 + value.minute * 60 + value.second) / 86400


def cell_kind(v) -> str:
    """Тип значения ячейки для отчётов: empty | bool | number | date | time | text | error."""
    if v is None or v == "":
        return "empty"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, (int, float)):
        return "number"
    if isinstance(v, str):
        if v in EXCEL_ERRORS.values():
            return "error"
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?", v):
            return "date"
        if re.fullmatch(r"\d{2}:\d{2}:\d{2}", v):
            return "time"
        return "text"
    return "other"


def clean_word_text(text: str) -> str:
    """Текст ячейки/абзаца Word -> читаемая строка (без маркеров конца ячейки)."""
    if text is None:
        return ""
    t = str(text)
    if t.endswith("\r\x07"):
        t = t[:-2]
    elif t.endswith("\x07") or t.endswith("\r"):
        t = t[:-1]
    return t.replace("\r", "\n").replace("\x0b", "\n").replace("\x0c", "\n").replace("\x07", "")


def to_word_text(text: str) -> str:
    """Обычные переводы строк -> абзацы Word (\\r)."""
    return str(text).replace("\r\n", "\r").replace("\n", "\r")


def truncate(text: str, limit: int) -> tuple[str, bool]:
    if limit <= 0 or len(text) <= limit:
        return text, False
    return text[:limit], True
