"""Word: документы, структура, чтение, поиск/замена, вставка, форматирование, выделение."""

import contextlib
import math
import os
import re
import time

import pywintypes

from . import com, config, safety, undo
from . import wd_common as wd
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .safety import WORD_EXTS, check_path
from .util import clean_word_text, cm_to_points, color_to_hex, parse_color, to_word_text, truncate
from .wd_common import (
    ALIGNMENTS,
    HIGHLIGHTS,
    active_document,
    all_documents,
    check_paragraph,
    document_allowed,
    ensure_free_anchor,
    escape_find,
    find_all,
    para_count,
    paragraph_index_at,
    paragraphs_range,
    pick_document,
    resolve_style,
    story_ranges,
)

_SAVE_FORMATS = {"docx": 16, "doc": 0, "docm": 13, "rtf": 6, "txt": 7, "html": 8, "htm": 8, "odt": 23, "xml": 19, "dotx": 14, "dotm": 15}
_STORY_NAMES = {1: "main", 2: "footnotes", 3: "endnotes", 4: "comments", 5: "textbox", 6: "even_header", 7: "primary_header", 8: "even_footer", 9: "primary_footer", 10: "first_header", 11: "first_footer"}


def _doc_pid_free(app) -> int | None:
    try:
        import win32process

        return win32process.GetWindowThreadProcessId(int(app.ActiveWindow.Hwnd))[1]
    except Exception:
        return None


def _active_window(app):
    try:
        return app.ActiveWindow
    except pywintypes.com_error:  # ни одного открытого документа
        return None


def _open_temporary_copy(app, path, previous):
    """Только чтение, макросы отключены, без записи в недавние файлы. Скрытое открытие (Open с Visible=False) Word тут же
    дописывает АКТИВНОМУ документу пользователя разделители сносок и стили колонтитулов (живая проверка: отпечаток
    меняется, Saved остаётся True). Поэтому открываем видимым при выключенной перерисовке экрана, сразу прячем окно
    и возвращаем прежнее активное окно. Не удалось спрятать — копию сразу закрываем без сохранения."""
    updating = app.ScreenUpdating
    app.ScreenUpdating = False
    try:
        with com.macros_disabled(app):
            # Open: Visible — 12-й, NoEncodingDialog — 15-й; пароли-пустышки исключают диалоги.
            doc = app.Documents.Open(path, False, True, False, "__office_live_no_password__",
                                     "__office_live_no_password__", False, "__office_live_no_password__",
                                     "__office_live_no_password__", 0, None, True, False, None, True)
        try:
            doc.Windows(1).Visible = False
        except (pywintypes.com_error, ToolError):
            with contextlib.suppress(pywintypes.com_error, ToolError):
                doc.Close(0)
            raise
        if previous is not None:
            # вернуть окно пользователя — желательно, но не ценой сравнения: копия уже спрятана
            with contextlib.suppress(pywintypes.com_error, ToolError):
                previous.Activate()
        return doc
    finally:
        app.ScreenUpdating = updating


def _word_running() -> bool:
    try:
        com.apps("word")
        return True
    except ToolError:
        return False


# ================================================================== документы


@office_tool("word_core", "write", title="Compare documents", file_args=("original", "revised"))
def word_compare_documents(
    original: str,
    revised: str,
    granularity: str = "word",
    compare_formatting: bool = True,
    compare_case: bool = True,
    compare_whitespace: bool = True,
    compare_tables: bool = True,
    compare_headers: bool = True,
    compare_footnotes: bool = True,
    compare_comments: bool = True,
    compare_moves: bool = True,
    author: str = "Office Live",
) -> dict:
    """Compare two versions into a new unsaved Word document with revisions, leaving both sources unchanged. To undo, close the result without saving; this tool has no office_undo entry.

    Args:
        original, revised: exact open document names or file paths inside allowed directories; both must use the same Word instance.
        granularity: word (default) or character.
        compare_formatting, compare_case, compare_whitespace, compare_tables, compare_headers, compare_footnotes, compare_comments, compare_moves: comparison options, all true by default. Textboxes and fields are always compared.
        author: revision author for the result, default Office Live.

    Closed files open temporarily and read-only, with macros disabled and no recent-file entry; their window is hidden at once (screen updating off) and your active window restored, and they close without saving even on failure.
    Word treats existing source revisions as accepted for comparison; warnings report this. Modal comparison warnings are disabled.
    """
    if granularity not in {"word", "character"}:
        raise ToolError("granularity must be word or character.")
    if not original.strip() or not revised.strip():
        raise ToolError("original and revised are required; pass exact names or file paths.")
    if not author.strip():
        raise ToolError("author must not be empty.")
    # Пути обоих источников проверяем до открытия первого временного документа.
    sources = []
    for name in (original, revised):
        is_path = os.path.isabs(name) or "/" in name or "\\" in name
        path = check_path(name, "read") if is_path else None
        sources.append({"name": name, "path": path, "app": None, "doc": None})
    for source in sources:
        try:
            source["app"], source["doc"] = pick_document(source["path"] or source["name"])
        except ToolError as exc:
            if not source["path"] or not any(s in str(exc) for s in ("not found. Open documents", "has no open documents")):
                raise
            if not os.path.isfile(source["path"]):
                raise ToolError(f"File not found: {source['path']}") from None
    opened_apps = [s["app"] for s in sources if s["app"] is not None]
    app = opened_apps[0] if opened_apps else com.apps("word")[0]
    if any(com.raw(a) is not com.raw(app) and not com._same_app(com.raw(a), com.raw(app)) for a in opened_apps):
        raise ToolError("Both documents must be in the same Word instance; close and reopen one in the other's instance.")
    temporary, warnings, result_name = [], [], ""
    previous = _active_window(app)  # окно пользователя: после каждого временного открытия возвращаем его
    try:
        for source in sources:
            if source["doc"] is None:
                existing = next((d for d in temporary if os.path.normcase(d.FullName) == os.path.normcase(source["path"])), None)
                if existing is not None:
                    source["doc"] = existing
            if source["doc"] is None:
                temporary.append(_open_temporary_copy(app, source["path"], previous))
                _, source["doc"] = pick_document(source["path"])
            if int(source["doc"].Revisions.Count):
                warnings.append(f"{source['doc'].Name}: existing revisions are treated as accepted for comparison; the source is unchanged.")
        result = app.CompareDocuments(sources[0]["doc"], sources[1]["doc"], 2, 1 if granularity == "word" else 0,
                                      bool(compare_formatting), bool(compare_case), bool(compare_whitespace), bool(compare_tables),
                                      bool(compare_headers), bool(compare_footnotes), True, True, bool(compare_comments),
                                      bool(compare_moves), author, True)
        result_name = result.Name
        counts = dict.fromkeys(("insertion", "deletion", "format", "move", "other"), 0)
        revisions = []
        total = int(result.Revisions.Count)
        for i in range(1, total + 1):
            revision = result.Revisions(i)
            code = int(revision.Type)
            kind = ("insertion" if code in {1, 16, 19} else "deletion" if code in {2, 17, 20} else "move" if code in {14, 15}
                    else "format" if code in {3, 4, 8, 10, 11, 12, 13} else "other")
            counts[kind] += 1
            if len(revisions) < 30:
                text, truncated = truncate(clean_word_text(revision.Range.Text), 200)
                revisions.append({"type": kind, "type_code": code, "author": revision.Author, "text": text, "text_truncated": truncated})
        return {"document": result_name, "saved": bool(result.Saved), "revision_count": total, "counts": counts,
                "revisions": revisions, "truncated": total > 30, "warnings": warnings,
                "undo_note": "Close the result without saving to discard the comparison; office_undo does not record this tool."}
    finally:
        failures = []
        for doc in reversed(temporary):
            try:
                doc.Close(0)
            except (pywintypes.com_error, ToolError) as exc:
                failures.append(com.translate(exc))
        if failures:
            raise ToolError(f"Comparison result: {result_name or 'not created'}. Could not close temporary read-only sources: {failures}")


@office_tool("word_core", "read", title="List open documents")
def word_list_documents() -> dict:
    """List every open Word document across ALL running Word instances: name, path, saved flag, page/paragraph counts, and which one is active. Call this first. Returns an empty list (not an error) when Word has no documents."""
    apps = com.apps("word")
    docs = []
    active = None
    for idx, app in enumerate(apps):
        act = active_document(app)
        act_name = act.Name if act is not None and document_allowed(act) else None  # запрещённый документ и как «активный» не называем
        if idx == 0:
            active = act_name
        for i in range(1, app.Documents.Count + 1):
            d = app.Documents(i)
            if not document_allowed(d):
                continue  # вне OFFICE_LIVE_ALLOWED_DIRS: даже имя не показываем
            docs.append({
                "name": d.Name, "path": d.Path or None, "saved": bool(d.Saved), "read_only": bool(d.ReadOnly),
                "paragraphs": para_count(d), "tables": int(d.Tables.Count), "active": idx == 0 and d.Name == act_name,
                "instance": idx,
            })
    return {"word_running": True, "instances": len(apps), "documents": docs, "active_document": active}


@office_tool("word_core", "read", title="Document structure")
def word_get_structure(document: str = "", max_headings: int = 200, per_match_links: bool = False) -> dict:
    """Overview of a document: counts (pages, words, paragraphs, tables, images, comments, tracked changes, bookmarks, fields), the heading outline with paragraph numbers, a summary of every table, and sections. Use it to navigate a long document before reading parts of it.

    Args:
        document: exact name, unique part of the name, full path, or '' for the active document.
        max_headings: cap on listed headings.
        per_match_links: true returns links for individual headings (up to 20); default returns one document link.
    """
    app, doc = pick_document(document)
    n = para_count(doc)
    headings = []
    scan = min(n, 3000)
    if scan:
        for i, p in enumerate(doc.Paragraphs, start=1):
            if i > scan:
                break
            lvl = int(p.OutlineLevel)
            if lvl < 10:
                txt = clean_word_text(p.Range.Text).strip()
                if txt:
                    headings.append({"paragraph": i, "level": lvl, "text": truncate(txt, 120)[0]})
                    if len(headings) >= max_headings:
                        break
    tables = []
    for i in range(1, int(doc.Tables.Count) + 1):
        t = doc.Tables(i)
        try:
            rows, cols = int(t.Rows.Count), int(t.Columns.Count)
        except pywintypes.com_error:
            rows, cols = int(t.Rows.Count), None  # неоднородная таблица
        first = clean_word_text(t.Cell(1, 1).Range.Text)
        tables.append({"index": i, "rows": rows, "columns": cols, "first_cell": truncate(first, 40)[0], "start_paragraph": paragraph_index_at(doc, int(t.Range.Start))})
    return {
        "document": doc.Name, "path": doc.FullName if doc.Path else None, "saved": bool(doc.Saved), "read_only": bool(doc.ReadOnly),
        "pages": int(doc.ComputeStatistics(2)), "words": int(doc.ComputeStatistics(0)), "characters": int(doc.ComputeStatistics(3)),
        "paragraphs": n, "sections": int(doc.Sections.Count), "tables": tables,
        "inline_images": int(doc.InlineShapes.Count), "floating_shapes": int(doc.Shapes.Count),
        "comments": int(doc.Comments.Count), "tracked_changes": int(doc.Revisions.Count), "track_changes_on": bool(doc.TrackRevisions),
        "bookmarks": int(doc.Bookmarks.Count), "fields": int(doc.Fields.Count), "footnotes": int(doc.Footnotes.Count),
        "headings": headings, "headings_truncated": len(headings) >= max_headings or scan < n,
        "protection": int(doc.ProtectionType),
    }


@office_tool("word_core", "open", title="Open document")
def word_open_document(path: str, read_only: bool = False) -> dict:
    """Open an existing document file in Word (starts Word if needed). Macros are disabled while opening. If the file is already open it is returned instead of opened twice.

    Args:
        path: full path to a .docx/.doc/.rtf/.txt/.html/.odt file.
        read_only: open read-only (always forced when the server runs in readonly mode).
    """
    full = check_path(path, "read")
    if not os.path.isfile(full):
        raise ToolError(f"File not found: {full}")
    if _word_running():
        for _, d in all_documents():
            if d.Path and os.path.normcase(os.path.abspath(d.FullName)) == os.path.normcase(full):
                return {"ok": True, "already_open": True, "document": d.Name, "path": d.FullName}
    app = com.primary_app("word", launch=True)
    try:
        with com.macros_disabled(app):
            # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument): пароль-пустышка не даёт зависнуть на диалоге
            doc = app.Documents.Open(full, False, bool(read_only or config.SETTINGS.readonly), False, "__office_live_no_password__")
    except pywintypes.com_error as exc:
        raise ToolError(f"Word could not open {full}: {com.com_error_text(exc)} (password-protected or corrupt file?)") from None
    return {"ok": True, "already_open": False, "document": doc.Name, "path": doc.FullName, "read_only": bool(doc.ReadOnly), "paragraphs": para_count(doc)}


@office_tool("word_core", "destructive", title="Reload document")
def word_reload_document(document: str, discard_changes: bool = False) -> dict:
    """Re-read an open document from disk after an external edit. Refuses unsaved changes unless discard_changes=true.

    Closes without saving and reopens in the same Word instance, preserving read-only mode and window placement.
    This is an undo barrier: history for the old document instance is dropped. Allowed folders and AutoSave guards apply.
    """
    if not document.strip():
        raise ToolError("'document' is required for reloading.")
    app, doc = wd.pick_document(document)
    if not doc.Path:
        raise ToolError(f"{doc.Name} has never been saved; use word_save_as first.")
    full = check_path(doc.FullName, "read")
    if not os.path.isfile(full):
        raise ToolError(f"File not found: {full}")
    if not doc.Saved and not discard_changes:
        raise ToolError(f"{doc.Name} has unsaved changes. Pass discard_changes=true to drop them and reload from disk.")
    read_only = bool(doc.ReadOnly)
    window = doc.ActiveWindow
    state = int(window.WindowState)
    bounds = (window.Left, window.Top, window.Width, window.Height)
    visible = bool(window.Visible)
    key = undo.stack_key("document", app, doc)
    doc.Close(0)  # wdDoNotSaveChanges
    undo.forget_closed_target(key, app)
    reason = "Reload closes the old document instance; its undo history has been dropped."
    try:
        with com.macros_disabled(app):
            # Same positional Open contract as temporary comparison sources: no password/conversion dialogs.
            reopened = app.Documents.Open(full, False, read_only, False, "__office_live_no_password__",
                                          "__office_live_no_password__", False, "__office_live_no_password__",
                                          "__office_live_no_password__", 0, None, True, False, None, True)
        new_window = reopened.ActiveWindow
        new_window.WindowState = 0  # wdWindowStateNormal
        new_window.Left, new_window.Top, new_window.Width, new_window.Height = bounds
        new_window.WindowState = state
        new_window.Visible = visible
        from .navigation import saved_as

        return {"ok": True, "document": reopened.Name, "path": reopened.FullName, "read_only": bool(reopened.ReadOnly),
                "undo": "not available: " + reason, **saved_as("word", reopened, (full,))}
    except Exception as exc:  # noqa: BLE001 — the old instance is already closed
        raise PartialChangeError(f"Reload of {full} interrupted after closing without saving. Open {full} manually if it is not open. "
                                 f"{reason} {com.translate(exc)}") from None


@office_tool("word_core", "write", title="New document")
def word_new_document(text: str = "") -> dict:
    """Create a new blank document in Word (starts Word if needed). The document is unsaved until word_save_as.

    Args:
        text: optional initial text (newlines become paragraphs).
    """
    app = com.primary_app("word", launch=True)
    doc = app.Documents.Add()
    if text:
        doc.Content.InsertAfter(to_word_text(text))
    return {"ok": True, "document": doc.Name, "paragraphs": para_count(doc)}


@office_tool("word_core", "destructive", title="Close document")
def word_close_document(document: str, save: bool = False, discard: bool = False) -> dict:
    """Close an open document. Refuses when it has unsaved changes unless you choose: save=true (saves first, needs an existing path) or discard=true. Never quits Word itself.

    Args:
        document: exact document name or full path (required).
        save: save before closing.
        discard: close without saving unsaved changes.
    """
    if not document:
        raise ToolError("'document' is required for closing.")
    app, doc = pick_document(document, allow_autosave=True)
    name = doc.Name
    saved_first = False
    if not doc.Saved:
        if save:
            if not doc.Path:
                raise ToolError(f"{name} has never been saved; use word_save_as first or pass discard=true.")
            check_path(doc.FullName, "write")  # зона доступа и запретные каталоги — и для сохранения уже открытого документа
            doc.Save()
            saved_first = True
        elif not discard:
            raise ToolError(f"{name} has unsaved changes. Pass save=true to save them or discard=true to drop them.")
    doc.Close(0)  # wdDoNotSaveChanges
    remaining = [d.Name for a, d in all_documents()] if _word_running() else []
    return {"ok": True, "closed": name, "saved_before_close": saved_first, "open_documents": remaining}


@office_tool("word_core", "save", title="Save document")
def word_save(document: str = "") -> dict:
    """Save a document to its current file (Ctrl+S). A document that was never saved needs word_save_as instead.

    Args:
        document: exact name or '' for the active document.
    """
    app, doc = pick_document(document, allow_autosave=True)
    if not doc.Path:
        raise ToolError(f"{doc.Name} has never been saved - use word_save_as with a path.")
    if doc.ReadOnly:
        raise ToolError(f"{doc.Name} is open read-only; use word_save_as to write a copy.")
    check_path(doc.FullName, "write")  # зона доступа и запретные каталоги
    doc.Save()
    return {"ok": True, "document": doc.Name, "path": doc.FullName, "saved": bool(doc.Saved)}


@office_tool("word_core", "save", title="Save document as")
def word_save_as(document: str, path: str, overwrite: bool = False) -> dict:
    """Save a document under a new name/format. Format follows the extension: .docx .doc .docm .rtf .txt .html .odt (use word_export_pdf for PDF). Refuses to replace an existing file unless overwrite=true. Returns links with the NEW file name; old link names resolve only inside this server session, not the separate browser link handler.

    Args:
        document: exact name of the open document (required).
        path: destination full path with extension.
        overwrite: allow replacing an existing file.
    """
    app, doc = pick_document(document, allow_autosave=True)
    full = check_path(path, "write", WORD_EXTS - {"pdf"})
    if os.path.exists(full) and not overwrite:
        raise ToolError(f"File already exists: {full}. Pass overwrite=true to replace it.")
    ext = full.rsplit(".", 1)[-1].lower()
    fmt = _SAVE_FORMATS.get(ext)
    if fmt is None:
        raise ToolError(f"Unsupported extension .{ext}. Use one of {sorted(_SAVE_FORMATS)}; for PDF use word_export_pdf.")
    os.makedirs(os.path.dirname(full), exist_ok=True)
    before = app.DisplayAlerts
    app.DisplayAlerts = 0
    try:
        old_names = (str(doc.Name), str(doc.FullName))
        doc.SaveAs2(full, fmt)
        from .navigation import saved_as

        link_result = saved_as("word", doc, old_names)
    finally:
        try:
            app.DisplayAlerts = before
        except pywintypes.com_error:
            pass
    return {"ok": True, "document": doc.Name, "path": doc.FullName, "overwrite": overwrite,
            "note": "The document is now named after the new file. Old links resolve only within this server session; use the new links in the chat.", **link_result}


@office_tool("word_core", "save", title="Export to PDF")
def word_export_pdf(document: str, path: str, overwrite: bool = False) -> dict:
    """Export a document to a PDF file (the open document keeps its name and format).

    Args:
        document: exact document name (required).
        path: destination .pdf path.
        overwrite: allow replacing an existing file.
    """
    app, doc = pick_document(document, allow_autosave=True)
    if not doc.Path:
        # Word обнуляет имя НЕсохранённого документа после экспорта в PDF — поэтому требуем сначала сохранить
        raise ToolError(f"{doc.Name} has never been saved. Save it with word_save_as first, then export to PDF.")
    full = check_path(path, "write", {"pdf"})
    if os.path.exists(full) and not overwrite:
        raise ToolError(f"File already exists: {full}. Pass overwrite=true to replace it.")
    os.makedirs(os.path.dirname(full), exist_ok=True)
    # ExportAsFixedFormat(OutputFileName, ExportFormat=wdExportFormatPDF(17), OpenAfterExport, OptimizeFor, Range=all, From, To, Item)
    doc.ExportAsFixedFormat(full, 17, False, 0, 0, 0, 0, 0)
    return {"ok": True, "path": full, "size_bytes": os.path.getsize(full) if os.path.exists(full) else None}


# ================================================================== чтение и поиск


@office_tool("word_core", "read", title="Read document")
def word_read_document(
    document: str = "",
    mode: str = "text",
    start_paragraph: int = 1,
    end_paragraph: int = 0,
    offset: int = 0,
    max_chars: int = 8000,
) -> dict:
    """Read document content. mode='text' returns plain text in pages of max_chars (continue with `offset` = previous offset + max_chars; the response says when the end is reached). mode='paragraphs' returns numbered paragraphs with their style, heading level and whether they sit in a table - use it to learn paragraph numbers for editing.

    Args:
        document: exact name or '' for the active document.
        mode: 'text' or 'paragraphs'.
        start_paragraph, end_paragraph: for mode='paragraphs' (end 0 = start+49; at most 200 per call).
        offset, max_chars: paging for mode='text'.
    """
    app, doc = pick_document(document)
    m = mode.lower()
    if m == "text":
        text = clean_word_text(doc.Content.Text) if False else doc.Content.Text.replace("\r", "\n").replace("\x07", "\t").replace("\x0b", "\n")
        total = len(text)
        off = max(0, int(offset))
        limit = max(1, int(max_chars))
        chunk = text[off:off + limit]
        return {
            "document": doc.Name, "total_chars": total, "offset": off, "returned_chars": len(chunk),
            "truncated": off + limit < total, "next_offset": off + limit if off + limit < total else None, "text": chunk,
        }
    if m != "paragraphs":
        raise ToolError("mode must be 'text' or 'paragraphs'.")
    n = para_count(doc)
    start = check_paragraph(doc, start_paragraph)
    end = int(end_paragraph) if end_paragraph else start + 49
    end = min(end, n, start + 199)
    items = []
    for i in range(start, end + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        item = {"n": i, "text": clean_word_text(rng.Text), "style": p.Style.NameLocal}
        lvl = int(p.OutlineLevel)
        if lvl < 10:
            item["heading_level"] = lvl
        if bool(rng.Information(12)):  # wdWithInTable
            item["in_table"] = True
        items.append(item)
    return {"document": doc.Name, "total_paragraphs": n, "from": start, "to": end, "has_more": end < n, "paragraphs": items}


FIND_TIME_BUDGET = 60.0  # с; дальше word_find отдаёт частичный результат вместо зависания клиента (#55)


@office_tool("word_core", "read", title="Find text")
def word_find(
    query: str,
    document: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    wildcards: bool = False,
    scope: str = "main",
    max_results: int = 100,
    context_chars: int = 60,
    per_match_links: bool = False,
) -> dict:
    """Find every occurrence of a text and report where it is (paragraph number, position) with surrounding context.

    Args:
        query: text to find (literal unless wildcards=true; Word wildcard syntax like 'cat?', '[0-9]{3}').
        document: exact name or '' for the active document.
        match_case, whole_word: matching options.
        wildcards: interpret query with Word wildcard syntax; counts like {1,2} work in any Word language.
        scope: 'main' (body), 'all' (also headers, footers, footnotes, text boxes) or 'selection'.
        max_results: stop collecting after this many hits.
        context_chars: characters of context on each side.
        per_match_links: true returns links for individual matching paragraphs (up to 20); default returns one document link.

    A search that runs longer than a minute stops and returns partial=true with the stories searched so far.
    """
    if not query:
        raise ToolError("'query' is empty.")
    if len(query) > 255:
        raise ToolError("Word cannot search for text longer than 255 characters.")
    app, doc = pick_document(document)
    text = wd.localize_wildcards(query, wd.list_separator(app)) if wildcards else escape_find(query)
    hits, total = [], 0
    cap = max(1, int(max_results))
    stories = story_ranges(doc, scope)
    deadline = time.monotonic() + FIND_TIME_BUDGET
    searched, partial = 0, False
    for story in stories:
        if time.monotonic() >= deadline:
            partial = True
            break
        for r in find_all(story, text, match_case, whole_word, wildcards, limit=100000):
            if time.monotonic() >= deadline:
                partial = True
                break
            total += 1
            if len(hits) >= cap:
                continue
            kind = int(r.StoryType)
            item = {"story": _STORY_NAMES.get(kind, f"story_{kind}"), "start": int(r.Start), "end": int(r.End), "text": truncate(r.Text, 200)[0]}
            if kind == 1:
                item["paragraph"] = paragraph_index_at(doc, int(r.Start))
                para = r.Paragraphs(1).Range
                ptxt = clean_word_text(para.Text)
                rel = int(r.Start) - int(para.Start)
                c = max(0, int(context_chars))
                item["context"] = ptxt[max(0, rel - c): rel + (int(r.End) - int(r.Start)) + c]
            hits.append(item)
        if partial:
            break
        searched += 1
    result = {"document": doc.Name, "query": query, "total_matches": total, "returned": len(hits), "matches": hits}
    if partial:
        result.update(partial=True, stories_searched=searched, stories_total=len(stories),
                      note=f"Search stopped after {FIND_TIME_BUDGET:.0f} s; counts cover the stories searched so far.")
    return result


# ================================================================== замена и вставка


def replace_in_doc(doc, find: str, replace: str, match_case=False, whole_word=False, wildcards=False, special_codes=False, scope="main") -> int:
    """Замена во всём документе (общий код word_replace_text, word_fill_placeholders и мостов). Возвращает число замен."""
    if not find:
        raise ToolError("'find' is empty.")
    if len(find) > 255:
        raise ToolError("Word cannot search for text longer than 255 characters.")
    ftext = find if (wildcards or special_codes) else escape_find(find)
    if wildcards:
        ftext = wd.localize_wildcards(ftext, wd.list_separator(doc.Application))
    rtext = to_word_text(replace)
    if not (wildcards or special_codes):
        rtext = escape_find(rtext)
    stories = story_ranges(doc, scope)
    total = 0
    for story in stories:
        matches = list(find_all(story, ftext, match_case, whole_word, wildcards))
        if not matches:
            continue
        # Find.Execute(Replace=all) «подстраивает» регистр замены под найденный текст, трактует обратный слэш+цифру как группу
        # даже без wildcards и не принимает строки >255 символов. Поэтому буквальную замену делаем через Range.Text
        # (точно, с конца, чтобы не сбить позиции); быстрый путь — только для wildcards/special_codes.
        fast = (wildcards or special_codes) and len(rtext) <= 255
        if fast:
            f = story.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            # Execute(FindText, MatchCase, WholeWord, Wildcards, SoundsLike, AllWordForms, Forward, Wrap=stop, Format, ReplaceWith, Replace=all)
            f.Execute(ftext, bool(match_case), bool(whole_word), bool(wildcards), False, False, True, 0, False, rtext, 2)
            total += len(matches)
        else:
            if wildcards or special_codes:
                raise ToolError("A replacement longer than 255 characters cannot be combined with wildcards/special codes.")
            literal = to_word_text(replace)
            for r in reversed(matches):
                r.Text = literal
                total += 1
    return total


def fill_placeholders(doc, values: dict, left: str = "{{", right: str = "}}", scope: str = "all") -> dict:
    """Подставляет значения вместо {{ключ}} (и {{ ключ }}) во всём документе; возвращает счётчики и оставшиеся метки."""
    replaced = {}
    for key, val in values.items():
        text = "" if val is None else str(val)
        n = 0
        for token in (f"{left}{key}{right}", f"{left} {key} {right}"):
            n += replace_in_doc(doc, token, text, match_case=True, scope=scope)
        replaced[str(key)] = n
    leftover = []
    pattern = re.compile(re.escape(left) + r"\s*([^{}]{1,60}?)\s*" + re.escape(right))
    for story in story_ranges(doc, scope):
        for m in pattern.finditer(story.Text or ""):
            if m.group(1) not in leftover:
                leftover.append(m.group(1))
    return {"replaced": replaced, "unfilled_placeholders": leftover}


@office_tool("word_core", "write", title="Fill placeholders")
def word_fill_placeholders(document: str, values: dict, left_delimiter: str = "{{", right_delimiter: str = "}}", scope: str = "all") -> dict:
    """Fill a template: replace every {{key}} (also {{ key }}) with its value, in the body, headers, footers and text boxes. Reports placeholders left unfilled.

    Args:
        document: exact document name (required).
        values: mapping like {"Client": "ACME Ltd", "Date": "04.10.2026", "Amount": "1 250,00"}.
        left_delimiter, right_delimiter: placeholder brackets (default '{{' and '}}').
        scope: 'all' (default), 'main' or 'selection'.
    """
    if not isinstance(values, dict) or not values:
        raise ToolError("'values' must be a non-empty object.")
    app, doc = pick_document(document)
    out = fill_placeholders(doc, values, left_delimiter, right_delimiter, scope)
    return {"ok": True, "document": doc.Name, **out}


@office_tool("word_core", "write", title="Find and replace")
def word_replace_text(
    find: str | None = None,
    replace: str | None = None,
    document: str | None = None,
    match_case: bool = False,
    whole_word: bool = False,
    wildcards: bool = False,
    special_codes: bool = False,
    scope: str = "main",
    pairs: list[dict[str, str]] | None = None,
) -> dict:
    """Replace every occurrence of a text (formatting of the replaced text is kept). Literal by default: the characters '^' and '\\' are NOT special. Returns the exact number of replacements. With 'Track changes' on, replacements are recorded as revisions.

    Args:
        find: text to look for (max 255 characters).
        replace: replacement ('' deletes the text). Newlines become paragraph marks.
        document: exact document name (required).
        match_case, whole_word: matching options.
        wildcards: Word wildcard syntax in `find` (counts like {1,2} work in any Word language); the replacement may then use \\1 \\2 for groups.
        special_codes: let '^p' (paragraph), '^t' (tab), '^l' (line break) etc. work in find/replace.
        scope: 'main' (body, default), 'all' (also headers, footers, footnotes, text boxes) or 'selection'.
        pairs: 1–200 {find, replace} objects instead of find/replace. Applied in order
            (later pairs see earlier replacements), with shared flags and one undo.
    """
    if document is None:
        raise ToolError("'document' is required.")
    batch = pairs is not None
    if batch:
        if find is not None or replace is not None:
            raise ToolError("'pairs' is mutually exclusive with 'find'/'replace'.")
        if not isinstance(pairs, list) or not 1 <= len(pairs) <= 200:
            raise ToolError("'pairs' must contain 1–200 find/replace objects.")
    else:
        pairs = [{"find": find, "replace": replace}]
    for pair in pairs:
        if not isinstance(pair, dict) or set(pair) != {"find", "replace"}:
            raise ToolError("Each pair must contain exactly 'find' and 'replace'.")
        if not isinstance(pair["find"], str) or not pair["find"]:
            raise ToolError("'find' must be a non-empty string.")
        if len(pair["find"]) > 255:
            raise ToolError("Word cannot search for text longer than 255 characters.")
        if not isinstance(pair["replace"], str):
            raise ToolError("'replace' must be a string ('' deletes the text).")
        if (wildcards or special_codes) and len(to_word_text(pair["replace"])) > 255:
            raise ToolError("A replacement longer than 255 characters cannot be combined with wildcards/special codes.")
        if wildcards:
            # Count unescaped groups outside character sets; Word checks the pattern syntax below.
            groups, in_set = 0, False
            for token in re.findall(r"\\.|.", pair["find"]):
                if token == "[":
                    in_set = True
                elif token == "]":
                    in_set = False
                elif token == "(" and not in_set:
                    groups += 1
            for token in re.findall(r"\\.|.", pair["replace"]):
                if len(token) == 2 and token[0] == "\\" and token[1] in "123456789" and int(token[1]) > groups:
                    raise ToolError(f"Replacement refers to missing wildcard group {token}.")
    app, doc = pick_document(document)
    stories = story_ranges(doc, scope)  # validate scope before preparing undo or changing text
    if wildcards or special_codes:
        for pair in pairs:
            text = wd.localize_wildcards(pair["find"], wd.list_separator(app)) if wildcards else pair["find"]
            # Let Word validate every pattern, including pairs which currently have no hits.
            for story in stories or [doc.Content]:
                list(find_all(story, text, match_case, whole_word, wildcards, limit=1))
    name, tracking = doc.Name, bool(doc.TrackRevisions)
    undo.require_undo("document", app, doc)
    counts = []
    try:
        for pair in pairs:
            count = replace_in_doc(doc, pair["find"], pair["replace"], match_case, whole_word, wildcards, special_codes, scope)
            counts.append({**pair, "count": count})
    except Exception as exc:
        raise PartialChangeError(f"Replacement interrupted after {len(counts)} completed pairs; text may have changed. {exc}") from None
    result = {
        "ok": True, "document": name, "replacements": sum(p["count"] for p in counts), "scope": scope,
        "track_changes_on": tracking,
    }
    if batch:
        result["pairs"] = counts
    return result


def _ensure_empty_last_paragraph(doc):
    """Гарантирует пустой последний абзац (для добавления в конец) и возвращает его диапазон."""
    last = doc.Paragraphs.Last.Range
    if clean_word_text(last.Text) != "":
        doc.Content.InsertParagraphAfter()
        last = doc.Paragraphs.Last.Range
    return last


@office_tool("word_core", "write", title="Insert text")
def word_insert_text(
    document: str,
    text: str,
    position: str = "end",
    paragraph: int = 0,
    bookmark: str = "",
    new_paragraph: bool = True,
    style: str = "Normal",
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    table_index: int = 0,
) -> dict:
    """Insert text. By default each insertion becomes its own paragraph(s) with the given style (newlines in `text` make several paragraphs). Returns the paragraph numbers of the inserted text.

    Args:
        document: exact document name (required).
        text: the text to insert.
        position: 'end' (append, default), 'start', 'after_paragraph', 'before_paragraph', 'after_table' / 'before_table' (with `table_index` - the easy way to put text next to a table), 'replace_paragraph' (swap the text of paragraph N, keeping its mark), 'bookmark' (replace the bookmark's content), 'selection' (replace the user's selection / insert at the cursor).
        paragraph: paragraph number (1-based) for the *_paragraph positions.
        bookmark: bookmark name for position='bookmark'.
        new_paragraph: true = insert as new paragraph(s); false = insert inline without adding a paragraph mark (end -> appended to the last paragraph; after_paragraph -> end of that paragraph).
        style: style for new paragraphs ('Normal', 'Heading 1', 'Title', 'List Bullet', 'Quote', ... or a custom name); '' keeps whatever Word inherits.
        bold, italic, font_size, font_color: direct formatting of the inserted text.
    """
    app, doc = pick_document(document)
    body = to_word_text(text)
    pos = position.lower()
    if pos not in ("end", "start", "after_paragraph", "before_paragraph", "after_table", "before_table", "replace_paragraph", "bookmark", "selection"):
        raise ToolError("position must be end, start, after_paragraph, before_paragraph, after_table, before_table, replace_paragraph, bookmark or selection.")
    style_val = resolve_style(doc, style) if style else None
    n = para_count(doc)
    if pos in ("after_paragraph", "before_paragraph", "replace_paragraph"):
        check_paragraph(doc, paragraph)
    if pos == "after_paragraph" and paragraph == n and new_paragraph:
        pos = "end"  # после последнего абзаца = в конец

    if pos == "end":
        if new_paragraph:
            last = _ensure_empty_last_paragraph(doc)
            r = doc.Range(int(last.Start), int(last.Start))
            r.InsertAfter(body)
        else:
            end = int(doc.Content.End) - 1
            r = doc.Range(end, end)
            r.InsertAfter(body)
    elif pos == "start":
        r = doc.Range(0, 0)
        r.InsertBefore(body + ("\r" if new_paragraph else ""))
    elif pos == "after_paragraph":
        p = doc.Paragraphs(paragraph).Range
        if new_paragraph:
            r = doc.Range(int(p.End), int(p.End))
            r.InsertBefore(body + "\r")
        else:
            r = doc.Range(int(p.End) - 1, int(p.End) - 1)
            r.InsertAfter(body)
    elif pos == "before_paragraph":
        p = doc.Paragraphs(paragraph).Range
        r = doc.Range(int(p.Start), int(p.Start))
        r.InsertBefore(body + ("\r" if new_paragraph else ""))
    elif pos in ("after_table", "before_table"):
        r = ensure_free_anchor(doc, pos, 0, table_index=table_index)
        r.InsertAfter(body)
    elif pos == "replace_paragraph":
        p = doc.Paragraphs(paragraph).Range
        r = doc.Range(int(p.Start), int(p.End) - 1)
        r.Text = body
        style_val = resolve_style(doc, style) if style and style != "Normal" else None  # стиль заменяемого абзаца по умолчанию сохраняем
    elif pos == "bookmark":
        if not bookmark:
            raise ToolError("'bookmark' is required for position='bookmark'.")
        if not doc.Bookmarks.Exists(bookmark):
            names = [doc.Bookmarks(i).Name for i in range(1, int(doc.Bookmarks.Count) + 1)]
            raise ToolError(f"Bookmark '{bookmark}' not found. Bookmarks: {names}")
        r = doc.Bookmarks(bookmark).Range
        r.Text = body
        doc.Bookmarks.Add(bookmark, r)  # замена текста удаляет закладку — восстанавливаем
        style_val = None
    else:  # selection
        r = doc.ActiveWindow.Selection.Range
        r.Text = body
        style_val = None
    if style_val is not None and new_paragraph:
        r.Style = style_val
    if bold is not None:
        r.Font.Bold = bool(bold)
    if italic is not None:
        r.Font.Italic = bool(italic)
    if font_size is not None:
        r.Font.Size = float(font_size)
    if font_color:
        r.Font.Color = parse_color(font_color)
    start_par = paragraph_index_at(doc, int(r.Start))
    end_par = paragraph_index_at(doc, max(int(r.End) - 1, int(r.Start)))
    return {"ok": True, "document": doc.Name, "inserted_chars": len(text), "paragraphs": [start_par, end_par], "total_paragraphs": para_count(doc)}


@office_tool("word_core", "destructive", title="Delete paragraphs")
def word_delete_paragraphs(document: str, start_paragraph: int, end_paragraph: int = 0) -> dict:
    """Delete one paragraph or a run of paragraphs (with their text).

    Args:
        document: exact document name (required).
        start_paragraph: first paragraph to delete (1-based).
        end_paragraph: last paragraph to delete (0 = only the first).
    """
    app, doc = pick_document(document)
    end = int(end_paragraph) or int(start_paragraph)
    rng = paragraphs_range(doc, int(start_paragraph), end)
    preview = truncate(clean_word_text(rng.Text), 120)[0]
    rng.Delete()
    return {"ok": True, "document": doc.Name, "deleted_paragraphs": [int(start_paragraph), end], "deleted_text_preview": preview, "total_paragraphs": para_count(doc)}


# ================================================================== форматирование


def _targets(doc, app, scope, start_paragraph, end_paragraph, find_text, match_case, table_index=0):
    s = scope.lower()
    if s == "paragraphs":
        if not start_paragraph:
            raise ToolError("'start_paragraph' is required for scope='paragraphs'.")
        return [paragraphs_range(doc, int(start_paragraph), int(end_paragraph) or int(start_paragraph))]
    if s == "find":
        if not find_text:
            raise ToolError("'find_text' is required for scope='find'.")
        hits = []
        for r in find_all(doc.Content, escape_find(find_text), match_case, False, False, limit=500):
            hits.append(r)
        if not hits:
            raise ToolError(f"Text '{find_text}' not found.")
        return hits
    if s == "selection":
        return [doc.ActiveWindow.Selection.Range]
    if s == "document":
        return [doc.Content]
    raise ToolError("scope must be 'paragraphs', 'find', 'selection' or 'document'.")


@office_tool("word_core", "write", title="Format text")
def word_format_text(
    document: str,
    scope: str = "paragraphs",
    start_paragraph: int = 0,
    end_paragraph: int = 0,
    find_text: str = "",
    match_case: bool = False,
    style: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_color: str | None = None,
    highlight: str | None = None,
    shading_color: str | None = None,
    superscript: bool | None = None,
    subscript: bool | None = None,
    all_caps: bool | None = None,
    alignment: str | None = None,
    line_spacing: float | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    left_indent_cm: float | None = None,
    first_line_indent_cm: float | None = None,
    keep_with_next: bool | None = None,
    page_break_before: bool | None = None,
    paragraph_shading: str | None = None,
    paragraph_border: str | None = None,
    border_color: str | None = None,
    border_width_pt: float | None = None,
    keep_together: bool | None = None,
) -> dict:
    """Apply a style and/or direct formatting. Pick the target with `scope`: paragraphs (start..end paragraph numbers), find (every occurrence of find_text), selection (what the user selected), or document (everything). Only the properties you pass are changed. Colors: '#RRGGBB' or names.

    Args:
        document: exact document name (required).
        scope: 'paragraphs' | 'find' | 'selection' | 'document'.
        start_paragraph, end_paragraph: for scope='paragraphs' (end 0 = just start).
        find_text, match_case: for scope='find'.
        style: paragraph/character style ('Heading 1', 'Title', 'Quote', 'List Bullet', 'Strong', or a custom style).
        font_name, font_size, bold, italic, underline, strikethrough, font_color, superscript, subscript, all_caps: character formatting.
        highlight: yellow|green|cyan|pink|blue|red|dark_blue|teal|dark_green|violet|dark_red|dark_yellow|gray|light_gray|black|none (marker-pen colors). shading_color: any '#RRGGBB' background behind the text.
        alignment: left|center|right|justify.
        line_spacing: multiple of single spacing (1, 1.15, 1.5, 2). space_before, space_after: points.
        left_indent_cm, first_line_indent_cm: indents in centimetres (first_line negative = hanging).
        keep_with_next, page_break_before: pagination flags.
        paragraph_shading: paragraph background color or 'none' to clear (all touched paragraphs).
        paragraph_border: none|box|left|top|bottom; replaces the four paragraph borders.
        border_color: border color (default black); requires a visible paragraph_border.
        border_width_pt: positive points (default 0.5), rounded to the nearest Word width; actual value returned.
        keep_together: keep all lines of each touched paragraph on one page.
    """
    app, doc = pick_document(document)
    targets = _targets(doc, app, scope, start_paragraph, end_paragraph, find_text, match_case)
    style_val = resolve_style(doc, style) if style else None
    if highlight is not None and highlight.lower() not in HIGHLIGHTS:
        raise ToolError(f"highlight must be one of {sorted(HIGHLIGHTS)}")
    if alignment is not None and alignment.lower() not in ALIGNMENTS:
        raise ToolError("alignment must be left, center, right or justify.")
    border_sides = {"none": (), "box": (-1, -2, -3, -4), "left": (-2,), "top": (-1,), "bottom": (-3,)}
    # wdBorderTop=-1, wdBorderLeft=-2, wdBorderBottom=-3, wdBorderRight=-4
    border = paragraph_border.lower() if paragraph_border is not None else None
    if border is not None and border not in border_sides:
        raise ToolError("paragraph_border must be none, box, left, top or bottom.")
    if (border_color is not None or border_width_pt is not None) and border in (None, "none"):
        raise ToolError("border_color/border_width_pt require a visible paragraph_border.")
    if keep_together is not None and not isinstance(keep_together, bool):
        raise ToolError("keep_together must be a boolean.")
    shade = None if paragraph_shading is None else (
        -16777216 if paragraph_shading.lower() == "none" else parse_color(paragraph_shading))  # wdColorAutomatic
    border_rgb = parse_color(border_color if border_color is not None else "black")
    width = 4  # wdLineWidth050pt, in eighths of a point
    if border_width_pt is not None:
        if isinstance(border_width_pt, bool) or not math.isfinite(border_width_pt) or border_width_pt <= 0:
            raise ToolError("border_width_pt must be finite and positive.")
        # WdLineWidth: 025pt, 050pt, 075pt, 100pt, 150pt, 225pt, 300pt, 450pt, 600pt.
        width = min((2, 4, 6, 8, 12, 18, 24, 36, 48), key=lambda w: abs(w / 8 - border_width_pt))
    # Parse existing colors before any style/font changes too.
    font_rgb = parse_color(font_color) if font_color else None
    text_shade = parse_color(shading_color) if shading_color else None
    paragraph_targets = []
    if paragraph_shading is not None or border is not None or keep_together is not None:
        seen = set()
        for rng in targets:
            for i in range(1, int(rng.Paragraphs.Count) + 1):
                p = rng.Paragraphs(i).Range
                key = (int(p.StoryType), int(p.Start), int(p.End))
                if key not in seen:
                    seen.add(key)
                    paragraph_targets.append(p)
        undo.require_undo("document", app, doc)  # refuse before the first change if undo cannot be prepared
    applied = []
    try:
        for rng in targets:
            if style_val is not None:
                rng.Style = style_val
            font = rng.Font
            if font_name:
                font.Name = font_name
            if font_size is not None:
                font.Size = float(font_size)
            if bold is not None:
                font.Bold = bool(bold)
            if italic is not None:
                font.Italic = bool(italic)
            if underline is not None:
                font.Underline = 1 if underline else 0
            if strikethrough is not None:
                font.StrikeThrough = bool(strikethrough)
            if font_color:
                font.Color = font_rgb
            if superscript is not None:
                font.Superscript = bool(superscript)
            if subscript is not None:
                font.Subscript = bool(subscript)
            if all_caps is not None:
                font.AllCaps = bool(all_caps)
            if highlight is not None:
                rng.HighlightColorIndex = HIGHLIGHTS[highlight.lower()]
            if shading_color:
                font.Shading.BackgroundPatternColor = text_shade
            pf = rng.ParagraphFormat
            if alignment is not None:
                pf.Alignment = ALIGNMENTS[alignment.lower()]
            if line_spacing is not None:
                pf.LineSpacingRule = 5  # wdLineSpaceMultiple
                pf.LineSpacing = 12.0 * float(line_spacing)  # 12 пт = одинарный интервал
            if space_before is not None:
                pf.SpaceBefore = float(space_before)
            if space_after is not None:
                pf.SpaceAfter = float(space_after)
            if left_indent_cm is not None:
                pf.LeftIndent = cm_to_points(left_indent_cm)
            if first_line_indent_cm is not None:
                pf.FirstLineIndent = cm_to_points(first_line_indent_cm)
            if keep_with_next is not None:
                pf.KeepWithNext = bool(keep_with_next)
            if page_break_before is not None:
                pf.PageBreakBefore = bool(page_break_before)
        for rng in paragraph_targets:
            pf = rng.ParagraphFormat
            if shade is not None:
                pf.Shading.Texture = 0  # wdTextureNone
                pf.Shading.BackgroundPatternColor = shade
            if border is not None:
                for side in (-1, -2, -3, -4):  # wdBorderTop/Left/Bottom/Right
                    edge = pf.Borders(side)
                    edge.LineStyle = 1 if side in border_sides[border] else 0  # wdLineStyleSingle / wdLineStyleNone
                    if side in border_sides[border]:
                        edge.LineWidth = width
                        edge.Color = border_rgb
            if keep_together is not None:
                pf.KeepTogether = keep_together
    except Exception as exc:
        raise PartialChangeError(f"Paragraph formatting interrupted; use office_undo. {com.translate(exc)}") from None
    for k, v in (("style", style), ("font_name", font_name), ("font_size", font_size), ("bold", bold), ("italic", italic), ("underline", underline),
                 ("strikethrough", strikethrough), ("font_color", font_color), ("highlight", highlight), ("shading_color", shading_color),
                 ("superscript", superscript), ("subscript", subscript), ("all_caps", all_caps), ("alignment", alignment),
                 ("line_spacing", line_spacing), ("space_before", space_before), ("space_after", space_after),
                 ("left_indent_cm", left_indent_cm), ("first_line_indent_cm", first_line_indent_cm),
                 ("keep_with_next", keep_with_next), ("page_break_before", page_break_before),
                 ("paragraph_shading", paragraph_shading), ("paragraph_border", paragraph_border),
                 ("border_color", border_color), ("keep_together", keep_together),
                 ("border_width_pt", width / 8 if border not in (None, "none") else None)):
        if v is not None and v != "":
            applied.append(f"{k}={v}")
    if not applied:
        raise ToolError("Nothing to apply: pass a style or at least one formatting parameter.")
    result = {"ok": True, "document": doc.Name, "scope": scope, "ranges_formatted": len(targets), "applied": applied}
    if border not in (None, "none"):
        result["border_width_pt"] = width / 8
    return result


@office_tool("word_core", "read", title="List styles")
def word_list_styles(document: str = "", kind: str = "paragraph", in_use_only: bool = True) -> dict:
    """List the styles available in a document (names are what word_format_text/word_insert_text accept).

    Args:
        document: exact name or '' for the active document.
        kind: 'paragraph', 'character', 'table', 'list' or 'all'.
        in_use_only: only styles used/defined in this document (default); false lists every built-in style too.
    """
    app, doc = pick_document(document)
    kinds = {"paragraph": 1, "character": 2, "table": 3, "list": 4}
    want = None if kind.lower() == "all" else kinds.get(kind.lower())
    if kind.lower() != "all" and want is None:
        raise ToolError("kind must be paragraph, character, table, list or all.")
    names = []
    for i in range(1, int(doc.Styles.Count) + 1):
        st = doc.Styles(i)
        if want is not None and int(st.Type) != want:
            continue
        if in_use_only and not bool(st.InUse):
            continue
        names.append(st.NameLocal)
        if len(names) >= 300:
            break
    return {"document": doc.Name, "kind": kind, "styles": names}


@office_tool("word_core", "write", title="Bulleted / numbered list")
def word_make_list(document: str, start_paragraph: int, end_paragraph: int = 0, list_type: str = "bullet") -> dict:
    """Turn paragraphs into a bulleted or numbered list, or remove list formatting.

    Args:
        document: exact document name (required).
        start_paragraph: first paragraph (1-based).
        end_paragraph: last paragraph (0 = only the first).
        list_type: 'bullet', 'number' or 'none'.
    """
    app, doc = pick_document(document)
    rng = paragraphs_range(doc, int(start_paragraph), int(end_paragraph) or int(start_paragraph))
    t = list_type.lower()
    if t == "bullet":
        rng.ListFormat.ApplyBulletDefault()
    elif t == "number":
        rng.ListFormat.ApplyNumberDefault()
    elif t == "none":
        rng.ListFormat.RemoveNumbers()
    else:
        raise ToolError("list_type must be 'bullet', 'number' or 'none'.")
    return {"ok": True, "document": doc.Name, "paragraphs": [int(start_paragraph), int(end_paragraph) or int(start_paragraph)], "list_type": t}


# ================================================================== выделение пользователя


def _selection(app, doc):
    try:
        return doc.ActiveWindow.Selection
    except pywintypes.com_error:
        return app.Selection


@office_tool("word_core", "read", title="Get user selection")
def word_get_selection(document: str = "") -> dict:
    """What the user currently has selected (or where the cursor is) in a Word document: text, paragraph numbers, style, whether it is inside a table. Use it for requests like 'rewrite the paragraph I selected'.

    Args:
        document: exact name or '' for the active document.
    """
    app, doc = pick_document(document)
    sel = _selection(app, doc)
    start, end = int(sel.Start), int(sel.End)
    text, truncated = truncate(clean_word_text(sel.Text) if start != end else "", 3000)
    out = {
        "document": doc.Name, "start": start, "end": end, "is_cursor_only": start == end,
        "text": text, "truncated": truncated,
        "first_paragraph": paragraph_index_at(doc, start),
        "last_paragraph": paragraph_index_at(doc, max(end - 1, start)),
        "in_table": bool(sel.Information(12)),
    }
    try:
        out["style"] = sel.Paragraphs(1).Style.NameLocal
        f = sel.Font
        size = f.Size
        out["font"] = {
            "name": f.Name or "mixed", "size": "mixed" if size == 9999999 else size,
            "bold": "mixed" if f.Bold == 9999999 else bool(f.Bold), "color": color_to_hex(f.Color),
        }
    except pywintypes.com_error:
        pass
    return out


@office_tool("word_core", "ui", title="Select / go to")
def word_select(document: str, target: str = "", start_paragraph: int = 0, end_paragraph: int = 0, find_text: str = "", table_index: int = 0) -> dict:
    """Move the user's cursor/selection and scroll there, to point at a result.

    Args:
        document: exact document name.
        target: 'paragraphs' (start..end), 'find' (first occurrence of find_text), 'table' (table_index), 'start' or 'end' of the document.
            '' (default) follows the arguments: find_text -> 'find', table_index -> 'table', otherwise 'paragraphs'.
        start_paragraph, end_paragraph: for target='paragraphs'.
        find_text: for target='find'.
        table_index: for target='table'.
    """
    t = _select_target(target, start_paragraph, end_paragraph, find_text, table_index)
    app, doc = pick_document(document)
    if t == "paragraphs":
        rng = paragraphs_range(doc, int(start_paragraph) or 1, int(end_paragraph) or int(start_paragraph) or 1)
    elif t == "find":
        hits = list(find_all(doc.Content, escape_find(find_text), False, False, False, limit=1))
        if not hits:
            raise ToolError(f"Text '{find_text}' not found.")
        rng = hits[0]
    elif t == "table":
        if not 1 <= int(table_index) <= int(doc.Tables.Count):
            raise ToolError(f"Table {table_index} does not exist; the document has {int(doc.Tables.Count)}.")
        rng = doc.Tables(int(table_index)).Range
    elif t == "start":
        rng = doc.Range(0, 0)
    elif t == "end":
        e = int(doc.Content.End) - 1
        rng = doc.Range(e, e)
    else:
        raise ToolError("target must be 'paragraphs', 'find', 'table', 'start' or 'end'.")
    return select_resolved_range(app, doc, rng)


def _select_target(target, start_paragraph, end_paragraph, find_text, table_index) -> str:
    """Цель word_select. Раньше find_text без target='find' молча выделял абзац 1 (демо 2026-10-05):
    пустая цель выводится из аргументов, а аргументы, которые выбранная цель проигнорировала бы, — ошибка."""
    given = {"find": "find_text", "table": "table_index", "paragraphs": "start_paragraph/end_paragraph"}
    passed = {"find": bool(find_text), "table": bool(table_index), "paragraphs": bool(start_paragraph or end_paragraph)}
    t = (target or "").strip().lower()
    if t and t not in {"paragraphs", "find", "table", "start", "end"}:
        raise ToolError("target must be 'paragraphs', 'find', 'table', 'start' or 'end'.")
    if not t:
        inferred = [k for k, v in passed.items() if v]
        if len(inferred) > 1:
            raise ToolError(f"Pass only one of {', '.join(given[k] for k in inferred)}, or set target.")
        return inferred[0] if inferred else "paragraphs"
    if t == "find" and not find_text:
        raise ToolError("target='find' needs find_text.")
    ignored = [given[k] for k, v in passed.items() if v and k != t]
    if ignored:
        raise ToolError(f"target='{t}' does not use {', '.join(ignored)}. Use target='find' for find_text, 'table' for table_index, "
                        "'paragraphs' for start_paragraph/end_paragraph (or leave target empty).")
    return t


def select_resolved_range(app, doc, rng):
    """Shared navigation, including scrolling; never edits the document."""
    try:
        doc.Activate()
        rng.Select()
        app.ActiveWindow.ScrollIntoView(rng, True)
    except pywintypes.com_error as exc:
        raise ToolError("Could not move the selection (the document window may be hidden): " + com.com_error_text(exc)) from None
    return {"ok": True, "document": doc.Name, "selected": [int(rng.Start), int(rng.End)]}


@office_tool("word_core", "write", title="Insert document", file_args=("file_path",))
def word_insert_document(document: str, file_path: str, position: str = "end", paragraph: int = 0,
                         bookmark: str = "", source_bookmark: str = "") -> dict:
    """Insert a saved document into a Range without changing Selection. Uses one Word UndoRecord.

    Args:
        document: exact target document.
        file_path: existing docx/doc/dotx/rtf/txt/odt within allowed folders; macro-enabled formats refused.
        position: end/start/after_paragraph/after_bookmark.
        paragraph, bookmark: insertion anchor for the corresponding position.
        source_bookmark: optional named range in the source file.

    Unsaved edits in an open source are not inserted: the saved disk version is used and a warning is returned.
    """
    if position not in {"end", "start", "after_paragraph", "after_bookmark"}:
        raise ToolError("Invalid insertion position.")
    path = safety.check_path(file_path, "read")
    if os.path.splitext(path)[1].lower() not in {".docx", ".doc", ".dotx", ".rtf", ".txt", ".odt"}:
        raise ToolError("Unsupported source extension (macro-enabled files are refused).")
    if not os.path.isfile(path):
        raise ToolError("Source file was not found.")
    app, doc = wd.pick_document(document)
    if wd.document_path(doc) and safety._real(path) == safety._real(wd.document_path(doc)):
        raise ToolError("Cannot insert the target document into itself.")
    warnings = []
    for _, source in wd.all_documents():
        if wd.document_path(source) and safety._real(wd.document_path(source)) == safety._real(path) and not bool(source.Saved):
            warnings.append("Вставлена сохранённая на диске версия; исходный документ содержит несохранённые изменения.")
    if position == "start":
        anchor = int(doc.Content.Start)
    elif position == "end":
        anchor = max(int(doc.Content.Start), int(doc.Content.End) - 1)
    elif position == "after_paragraph":
        if type(paragraph) is not int or not 1 <= paragraph <= int(doc.Paragraphs.Count):
            raise ToolError("paragraph is outside the document.")
        anchor = min(int(doc.Paragraphs(paragraph).Range.End), int(doc.Content.End) - 1)
    else:
        if not bookmark or not doc.Bookmarks.Exists(bookmark):
            raise ToolError("bookmark was not found.")
        anchor = min(int(doc.Bookmarks.Item(bookmark).Range.End), int(doc.Content.End) - 1)
    if wd._in_table(doc, anchor):
        raise ToolError("Insertion anchor is inside a table; choose a free paragraph.")
    rng = doc.Range(anchor, anchor)
    before_end, before_paragraphs = int(doc.Content.End), int(doc.Paragraphs.Count)
    undo.require_undo("document", app, doc)
    try:
        with com.macros_disabled(app):  # legacy .doc can contain macros too
            rng.InsertFile(path, source_bookmark, False, False, False)
        inserted = int(doc.Content.End) - before_end
        return {"document": doc.Name, "start": anchor, "end": anchor + inserted, "characters_inserted": inserted,
                "paragraphs_before": before_paragraphs, "paragraphs_after": int(doc.Paragraphs.Count), "warnings": warnings}
    except pywintypes.com_error as exc:
        try:
            applied = int(doc.Content.End) - before_end
        except pywintypes.com_error:
            applied = "unknown"
        raise PartialChangeError(f"InsertFile interrupted: applied={applied} characters. {com.translate(exc)}") from None
