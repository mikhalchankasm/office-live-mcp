"""Общие помощники Word: выбор документа, стили, абзацы, диапазоны.

Правила COM те же (см. com.py): только позиционные аргументы; необязательные Variant — None/Missing по месту.
"""

import os

import pythoncom
import pywintypes

from . import com, config, safety
from .errors import ToolError
from .util import clean_word_text

MISSING = pythoncom.Missing
WRITE_KINDS = {"write", "destructive", "save"}

# имя -> WdBuiltinStyle (работает в любой языковой версии Word)
BUILTIN_STYLES = {
    "normal": -1, "heading 1": -2, "heading 2": -3, "heading 3": -4, "heading 4": -5, "heading 5": -6,
    "heading 6": -7, "heading 7": -8, "heading 8": -9, "heading 9": -10, "toc 1": -20, "toc 2": -21,
    "toc 3": -22, "footnote text": -30, "comment text": -31, "header": -32, "footer": -33, "caption": -35,
    "list": -48, "list bullet": -49, "list number": -50, "list bullet 2": -55, "list bullet 3": -56,
    "list number 2": -59, "list number 3": -60, "title": -63, "closing": -64, "signature": -65,
    "default paragraph font": -66, "body text": -67, "body text indent": -68, "subtitle": -75,
    "hyperlink": -86, "strong": -88, "emphasis": -89, "plain text": -91, "no spacing": -158,
    "list paragraph": -180, "quote": -181, "intense quote": -182, "subtle emphasis": -261,
    "intense emphasis": -262, "subtle reference": -263, "intense reference": -264, "book title": -265,
}
TABLE_STYLE_ALIASES = {"table grid": ["Table Grid", "Сетка таблицы", "Tabellenraster", "Grille du tableau"]}

ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3}
HIGHLIGHTS = {
    "none": 0, "black": 1, "blue": 2, "turquoise": 3, "cyan": 3, "green": 4, "bright_green": 4, "pink": 5,
    "red": 6, "yellow": 7, "white": 8, "dark_blue": 9, "teal": 10, "dark_green": 11, "violet": 12,
    "dark_red": 13, "dark_yellow": 14, "gray": 15, "gray_50": 15, "light_gray": 16, "gray_25": 16,
}


# ------------------------------------------------------------------ документы


def _strict_write() -> bool:
    return bool(config.SETTINGS.strict_target and com.current_kind() in WRITE_KINDS)


def _need_explicit(name: str):
    if _strict_write() and not (name or "").strip():
        raise ToolError("Document name is required for write operations (OFFICE_LIVE_STRICT_TARGET is on). Call word_list_documents first.")


def _same_path(a: str, b: str) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def document_path(doc) -> str:
    try:
        return doc.FullName if doc.Path else ""
    except pywintypes.com_error:
        return ""


def document_allowed(doc) -> bool:
    return safety.doc_allowed(document_path(doc))


def _guard_document(app, doc, allow_autosave: bool = False):
    com.note_target(f"document:{document_path(doc) or doc.Name}")
    if com.current_kind() in WRITE_KINDS and not allow_autosave and config.SETTINGS.autosave == "block":
        try:
            auto = bool(doc.AutoSaveOn)
        except (pywintypes.com_error, AttributeError):
            auto = False
        if auto:
            raise ToolError(
                f"AutoSave is ON for {doc.Name}: every change would be saved to the cloud file immediately, before the user can review it. "
                "Ask the user to turn AutoSave off for this document (or start the server with OFFICE_LIVE_AUTOSAVE=allow)."
            )
    from .undo import selected

    selected("document", app, doc)
    return app, doc


def all_documents(launch: bool = False) -> list:
    pairs = []
    for app in com.apps("word", launch):
        for i in range(1, app.Documents.Count + 1):
            d = app.Documents(i)
            if document_allowed(d):
                pairs.append((app, d))
    return pairs


def active_document(app):
    """ActiveDocument без исключения, когда документов нет."""
    if app.Documents.Count == 0:
        return None
    try:
        return app.ActiveDocument
    except pywintypes.com_error:
        return None


def pick_document(name: str = "", launch: bool = False, allow_autosave: bool = False, *, exact_only: bool = False):
    """(app, document): точное имя / полный путь / уникальная подстрока / '' = активный.

    В строгом режиме для пишущих инструментов — только точное имя или полный путь.
    """
    _need_explicit(name)
    strict = _strict_write() or exact_only
    name = (name or "").strip()
    apps = com.apps("word", launch)
    pairs, outside = [], 0
    for app in apps:
        for i in range(1, app.Documents.Count + 1):
            d = app.Documents(i)
            if document_allowed(d):
                pairs.append((app, d))
            else:
                outside += 1
    if not pairs:
        note = f" ({outside} open document(s) are outside OFFICE_LIVE_ALLOWED_DIRS)" if outside else ""
        raise ToolError("Word has no open documents" + note + ". Use word_open_document or word_new_document.")
    if not name:
        for app in apps:
            d = active_document(app)
            if d is not None:
                if not document_allowed(d):
                    raise ToolError("The active document is outside OFFICE_LIVE_ALLOWED_DIRS; pass the name of an allowed document.")
                return _guard_document(app, d, allow_autosave)
        return _guard_document(*pairs[0], allow_autosave)
    low = name.lower()
    has_sep = "\\" in name or "/" in name
    infos = [(a, d, d.Name) for a, d in pairs]
    exact = [(a, d) for a, d, n in infos if n.lower() == low]
    if not exact and has_sep:
        exact = [(a, d) for a, d, n in infos if d.Path and _same_path(d.FullName, name)]
    if len(exact) == 1:
        return _guard_document(*exact[0], allow_autosave)
    if len(exact) > 1:
        raise ToolError(f"Document name '{name}' exists in several Word instances; pass the full path. Candidates: {[d.FullName for _, d in exact]}")
    names = [n for _, _, n in infos]
    if not strict:
        subs = [(a, d) for a, d, n in infos if low in n.lower()]
        if len(subs) == 1:
            return _guard_document(*subs[0], allow_autosave)
        if subs:
            raise ToolError(f"Document name '{name}' is ambiguous. Open documents: {names}")
    raise ToolError(f"Document '{name}' not found. Open documents: {names}")


# ------------------------------------------------------------------ стили


def list_style_names(doc, limit: int = 60) -> list[str]:
    names = []
    try:
        for i in range(1, int(doc.Styles.Count) + 1):
            st = doc.Styles(i)
            if bool(st.InUse):
                names.append(st.NameLocal)
                if len(names) >= limit:
                    break
    except pywintypes.com_error:
        pass
    return names


def resolve_style(doc, name: str):
    """Имя стиля -> значение для Range.Style: встроенные английские имена работают в любой локали."""
    key = name.strip().lower()
    if key in BUILTIN_STYLES:
        return BUILTIN_STYLES[key]
    try:
        doc.Styles(name)
        return name
    except pywintypes.com_error:
        pass
    raise ToolError(f"Style '{name}' not found. Built-in English names work everywhere (Normal, Heading 1..9, Title, Subtitle, List Bullet, List Number, Quote, Caption, No Spacing). Styles in use: {list_style_names(doc)}")


def apply_table_style(doc, tbl, name: str) -> str:
    """Стиль таблицы по имени (в т.ч. 'Table Grid' в любой локали); при неудаче — обычные границы."""
    candidates = TABLE_STYLE_ALIASES.get(name.strip().lower(), [name])
    for cand in candidates:
        try:
            tbl.Style = cand
            return cand
        except pywintypes.com_error:
            continue
    if name.strip().lower() == "table grid":
        tbl.Borders.Enable = True
        return "borders"
    raise ToolError(f"Table style '{name}' not found (style names depend on the Word UI language). Use word_list_styles(kind='table').")


# ------------------------------------------------------------------ абзацы и диапазоны


def para_count(doc) -> int:
    return int(doc.Paragraphs.Count)


def check_paragraph(doc, index: int) -> int:
    n = para_count(doc)
    if not 1 <= int(index) <= n:
        raise ToolError(f"Paragraph {index} does not exist; the document has {n} paragraphs.")
    return int(index)


def paragraphs_range(doc, start: int, end: int | None = None):
    """Диапазон абзацев start..end (1-based, включительно)."""
    end = start if end is None else end
    start, end = check_paragraph(doc, start), check_paragraph(doc, end)
    if end < start:
        start, end = end, start
    return doc.Range(doc.Paragraphs(start).Range.Start, doc.Paragraphs(end).Range.End)


def paragraph_index_at(doc, position: int) -> int:
    """Номер абзаца, содержащего символ в позиции position."""
    end = min(int(position) + 1, int(doc.Content.End))
    return int(doc.Range(0, max(end, 1)).Paragraphs.Count)


def para_text(rng_or_para) -> str:
    return clean_word_text(rng_or_para.Range.Text if hasattr(rng_or_para, "Range") else rng_or_para.Text)


def story_ranges(doc, scope: str):
    """Диапазоны для поиска/замены: main | all (колонтитулы, сноски, надписи) | selection."""
    s = (scope or "main").lower()
    if s == "main":
        return [doc.Content]
    if s == "selection":
        return [doc.ActiveWindow.Selection.Range]
    if s == "all":
        out = []
        for story in doc.StoryRanges:
            r = story
            guard = 0
            while r is not None and guard < 200:
                out.append(r)
                r = r.NextStoryRange
                guard += 1
        return out
    raise ToolError("scope must be 'main', 'all' or 'selection'.")


def find_all(rng_story, text: str, match_case=False, whole_word=False, wildcards=False, limit=100000):
    """Генератор найденных диапазонов (копий) в story; текст уже подготовлен (экранирован) вызывающим."""
    rng = rng_story.Duplicate
    end = int(rng_story.End)
    count = 0
    while count < limit:
        f = rng.Find
        f.ClearFormatting()
        # Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap=wdFindStop)
        ok = f.Execute(text, bool(match_case), bool(whole_word), bool(wildcards), False, False, True, 0)
        if not ok:
            return
        hit_start, hit_end = int(rng.Start), int(rng.End)
        yield rng.Duplicate
        count += 1
        if hit_end >= end or hit_end <= hit_start and hit_start >= end:
            return
        nxt = hit_end if hit_end > hit_start else hit_end + 1
        rng.SetRange(nxt, end)


def escape_find(text: str) -> str:
    """Спецсимвол '^' в строках поиска/замены Word означает код (^p, ^t...). Буквальный '^' — '^^'."""
    return text.replace("^", "^^")


def _in_table(doc, position: int) -> bool:
    return bool(doc.Range(int(position), int(position)).Information(12))  # wdWithInTable


def _paragraph_above_table(doc, tbl) -> int:
    """Создаёт пустой абзац прямо над таблицей и возвращает его начало. Table.Split() здесь не работает ('Неверный параметр')."""
    start_t = int(tbl.Range.Start)
    if start_t == 0 or _in_table(doc, start_t - 1):
        # таблица в самом начале документа: единственный путь — Selection.SplitTable (вставляет абзац над строкой)
        app = doc.Application
        doc.Activate()
        sel = app.Selection
        saved = (int(sel.Start), int(sel.End))
        try:
            tbl.Rows(1).Range.Select()
            app.Selection.SplitTable()
        finally:
            try:
                doc.Range(saved[0], saved[1]).Select()
            except Exception:  # noqa: BLE001 — позиция могла сдвинуться
                pass
        return 0 if start_t == 0 else int(tbl.Range.Start) - 1
    doc.Range(start_t - 1, start_t - 1).InsertAfter("\r")  # конец предыдущего абзаца: разрез надвое
    return start_t


def ensure_free_anchor(doc, position: str, paragraph: int = 0, separate_from_table: bool = False, table_index: int = 0):
    """Создаёт новый пустой абзац в нужном месте (стиль Normal) и возвращает свёрнутый диапазон в нём.

    Сюда вставляются таблицы, оглавление, картинки. position: end | start | after_paragraph | before_paragraph |
    after_table | before_table (с table_index — не нужно считать абзацы: ячейки таблиц тоже абзацы).
    Вставка «перед таблицей» делается разрезом абзаца, а не символом в начале ячейки (иначе абзац попал бы внутрь таблицы).
    """
    pos = (position or "end").lower()
    n = para_count(doc)
    if pos in ("after_paragraph", "before_paragraph"):
        check_paragraph(doc, paragraph)
        if pos == "after_paragraph" and int(paragraph) == n:
            pos = "end"
        elif _in_table(doc, int(doc.Paragraphs(int(paragraph)).Range.Start)):
            raise ToolError(f"Paragraph {paragraph} is inside a table; pick a paragraph outside it (see word_read_document mode='paragraphs').")
    if pos in ("after_table", "before_table"):
        total = int(doc.Tables.Count)
        if not 1 <= int(table_index) <= total:
            raise ToolError(f"table_index {table_index} is invalid; the document has {total} table(s).")
    if pos == "end":
        last = doc.Paragraphs.Last.Range
        if clean_word_text(last.Text) != "":
            doc.Content.InsertParagraphAfter()
            last = doc.Paragraphs.Last.Range
        start = int(last.Start)
    elif pos == "start":
        if _in_table(doc, 0):
            _paragraph_above_table(doc, doc.Tables(1))  # абзац над таблицей, с которой начинается документ
        else:
            doc.Range(0, 0).InsertBefore("\r")
        start = 0
    elif pos == "after_paragraph":
        end = int(doc.Paragraphs(int(paragraph)).Range.End) - 1  # перед маркером абзаца: разрезаем абзац надвое
        doc.Range(end, end).InsertAfter("\r")
        start = end + 1
    elif pos == "before_paragraph":
        start = int(doc.Paragraphs(int(paragraph)).Range.Start)
        doc.Range(start, start).InsertBefore("\r")
    elif pos == "after_table":
        start = int(doc.Tables(int(table_index)).Range.End)  # начало следующего абзаца — уже вне таблицы
        doc.Range(start, start).InsertBefore("\r")
    elif pos == "before_table":
        start = _paragraph_above_table(doc, doc.Tables(int(table_index)))  # пустой абзац над таблицей
    else:
        raise ToolError("position must be 'end', 'start', 'after_paragraph', 'before_paragraph', 'after_table' or 'before_table'.")
    anchor = doc.Range(start, start)
    anchor.Style = -1  # Normal: пустой абзац не должен унаследовать заголовок соседа
    if separate_from_table:
        # новая таблица вплотную к существующей СЛИВАЕТСЯ с ней — оставляем между ними пустой абзац-разделитель
        if start > 0 and _in_table(doc, start - 1):  # прямо после таблицы: разделитель ПЕРЕД новым абзацем
            doc.Range(start, start).InsertBefore("\r")
            start += 1
            doc.Range(start, start).Style = -1
        nxt = int(doc.Range(start, start).Paragraphs(1).Range.End)
        if nxt < int(doc.Content.End) and _in_table(doc, nxt):
            doc.Range(start, start).InsertParagraphAfter()
    return doc.Range(start, start)


def collapsed_at(doc, position: str, paragraph: int = 0):
    """Свёрнутый диапазон для вставки «на месте» (разрывы и т. п.): end | start | after_paragraph | before_paragraph | selection."""
    pos = (position or "end").lower()
    if pos == "end":
        e = int(doc.Content.End) - 1
        return doc.Range(e, e)
    if pos == "start":
        return doc.Range(0, 0)
    if pos == "selection":
        r = doc.ActiveWindow.Selection.Range
        r.Collapse(1)
        return r
    if pos in ("after_paragraph", "before_paragraph"):
        check_paragraph(doc, paragraph)
        p = doc.Paragraphs(int(paragraph)).Range
        if pos == "before_paragraph":
            return doc.Range(int(p.Start), int(p.Start))
        e = int(p.End) - 1  # конец текста абзаца, до маркера
        return doc.Range(e, e)
    raise ToolError("position must be 'end', 'start', 'after_paragraph', 'before_paragraph' or 'selection'.")
