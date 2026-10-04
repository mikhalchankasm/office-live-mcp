"""Word: макет страницы, колонтитулы, оглавление, картинки, разрывы, комментарии, исправления, закладки, ссылки, сноски, свойства."""

import io
import json
import os
import re

import pythoncom
import pywintypes
from mcp.server.mcpserver import Image

from . import com
from .errors import ToolError
from .registry import office_tool
from .safety import IMAGE_EXTS, check_path
from .util import clean_word_text, cm_to_points, truncate
from .wd_common import (
    ALIGNMENTS, active_document, collapsed_at, ensure_free_anchor, escape_find, find_all, paragraph_index_at, paragraphs_range,
    pick_document,
)

MISSING = pythoncom.Missing

PAPER = {"a4": 7, "a3": 6, "a5": 9, "letter": 2, "legal": 5}
HF_TYPES = {"primary": 1, "first_page": 2, "even": 3}


def _sections(doc, section: int):
    n = int(doc.Sections.Count)
    if int(section) == 0:
        return [doc.Sections(i) for i in range(1, n + 1)]
    if not 1 <= int(section) <= n:
        raise ToolError(f"Section {section} does not exist; the document has {n}.")
    return [doc.Sections(int(section))]


# ================================================================== страница


@office_tool("word_layout", "write", title="Page setup")
def word_page_setup(
    document: str,
    section: int = 0,
    orientation: str | None = None,
    paper: str | None = None,
    margins_cm: float | None = None,
    top_cm: float | None = None,
    bottom_cm: float | None = None,
    left_cm: float | None = None,
    right_cm: float | None = None,
    columns: int | None = None,
    different_first_page: bool | None = None,
) -> dict:
    """Page layout: orientation, paper size, margins, text columns.

    Args:
        document: exact document name (required).
        section: section number (1-based) or 0 for every section.
        orientation: 'portrait' or 'landscape'.
        paper: A4, A3, A5, Letter or Legal.
        margins_cm: all four margins at once (centimetres); top_cm/bottom_cm/left_cm/right_cm override individual ones.
        columns: number of text columns (1-6).
        different_first_page: use a different header/footer on the first page.
    """
    app, doc = pick_document(document)
    applied = []
    for sec in _sections(doc, section):
        ps = sec.PageSetup
        if orientation:
            o = orientation.lower()
            if o not in ("portrait", "landscape"):
                raise ToolError("orientation must be 'portrait' or 'landscape'.")
            ps.Orientation = 0 if o == "portrait" else 1
        if paper:
            if paper.lower() not in PAPER:
                raise ToolError(f"paper must be one of {sorted(PAPER)}")
            ps.PaperSize = PAPER[paper.lower()]
        if margins_cm is not None:
            ps.TopMargin = ps.BottomMargin = ps.LeftMargin = ps.RightMargin = cm_to_points(margins_cm)
        if top_cm is not None:
            ps.TopMargin = cm_to_points(top_cm)
        if bottom_cm is not None:
            ps.BottomMargin = cm_to_points(bottom_cm)
        if left_cm is not None:
            ps.LeftMargin = cm_to_points(left_cm)
        if right_cm is not None:
            ps.RightMargin = cm_to_points(right_cm)
        if columns is not None:
            ps.TextColumns.SetCount(max(1, min(6, int(columns))))
        if different_first_page is not None:
            ps.DifferentFirstPageHeaderFooter = bool(different_first_page)
    for k, v in (("orientation", orientation), ("paper", paper), ("margins_cm", margins_cm), ("top_cm", top_cm), ("bottom_cm", bottom_cm),
                 ("left_cm", left_cm), ("right_cm", right_cm), ("columns", columns), ("different_first_page", different_first_page)):
        if v is not None:
            applied.append(f"{k}={v}")
    if not applied:
        raise ToolError("Nothing to change.")
    ps = doc.Sections(1).PageSetup
    return {
        "ok": True, "document": doc.Name, "applied": applied,
        "page_cm": [round(float(ps.PageWidth) * 2.54 / 72, 2), round(float(ps.PageHeight) * 2.54 / 72, 2)],
        "pages": int(doc.ComputeStatistics(2)),
    }


# ================================================================== колонтитулы


def _hf(sec, kind: str, typ: int):
    return sec.Headers(typ) if kind == "header" else sec.Footers(typ)


def _append_text(hf, text: str):
    r = hf.Range
    e = int(r.End) - 1
    rr = hf.Range
    rr.SetRange(e, e)
    rr.InsertAfter(text)


def _append_field(hf, code: int):
    r = hf.Range
    e = int(r.End) - 1
    rr = hf.Range
    rr.SetRange(e, e)
    hf.Range.Fields.Add(rr, code)


@office_tool("word_layout", "write", title="Headers and footers")
def word_headers_footers(
    document: str,
    action: str = "get",
    kind: str = "footer",
    section: int = 1,
    which: str = "primary",
    text: str = "",
    alignment: str = "center",
    page_numbers: str = "none",
    page_label: str = "Page ",
    of_label: str = " of ",
    font_size: float | None = None,
    different_first_page: bool | None = None,
) -> dict:
    """Read or set page headers/footers: text and automatic page numbers.

    Args:
        document: exact document name (required).
        action: 'get' (all header/footer texts), 'set' (replace content), 'clear'.
        kind: 'header' or 'footer'.
        section: section number (1-based) or 0 for every section.
        which: 'primary' (normal pages), 'first_page' or 'even'.
        text: text for 'set' (may be empty when only page numbers are wanted).
        alignment: left|center|right|justify.
        page_numbers: 'none', 'page' (7) or 'page_of_total' (Page 7 of 12) appended after the text.
        page_label, of_label: wording around the numbers ('Стр. ' / ' из ' for Russian).
        font_size: points.
        different_first_page: enable a separate first-page header/footer.
    """
    app, doc = pick_document(document)
    act = action.lower()
    if act == "get":
        out = []
        for i, sec in enumerate(_sections(doc, section), start=1 if int(section) == 0 else int(section)):
            item = {"section": i, "different_first_page": bool(sec.PageSetup.DifferentFirstPageHeaderFooter)}
            for k in ("header", "footer"):
                for name, typ in HF_TYPES.items():
                    t = clean_word_text(_hf(sec, k, typ).Range.Text).strip()
                    if t:
                        item[f"{k}_{name}"] = t
            out.append(item)
        return {"document": doc.Name, "sections": out}
    if kind not in ("header", "footer"):
        raise ToolError("kind must be 'header' or 'footer'.")
    if which not in HF_TYPES:
        raise ToolError("which must be 'primary', 'first_page' or 'even'.")
    if alignment.lower() not in ALIGNMENTS:
        raise ToolError("alignment must be left, center, right or justify.")
    for sec in _sections(doc, section):
        if different_first_page is not None:
            sec.PageSetup.DifferentFirstPageHeaderFooter = bool(different_first_page)
        hf = _hf(sec, kind, HF_TYPES[which])
        if act == "clear":
            hf.Range.Text = ""
            continue
        if act != "set":
            raise ToolError("action must be 'get', 'set' or 'clear'.")
        hf.Range.Text = ""
        pn = page_numbers.lower()
        if text:
            _append_text(hf, text + ("  " if pn != "none" else ""))
        if pn == "page":
            _append_field(hf, 33)  # wdFieldPage
        elif pn == "page_of_total":
            _append_text(hf, page_label)
            _append_field(hf, 33)
            _append_text(hf, of_label)
            _append_field(hf, 26)  # wdFieldNumPages
        elif pn != "none":
            raise ToolError("page_numbers must be 'none', 'page' or 'page_of_total'.")
        hf.Range.ParagraphFormat.Alignment = ALIGNMENTS[alignment.lower()]
        if font_size is not None:
            hf.Range.Font.Size = float(font_size)
    return {"ok": True, "document": doc.Name, "action": act, "kind": kind, "section": section, "which": which}


# ================================================================== оглавление и поля


@office_tool("word_layout", "write", title="Table of contents")
def word_manage_toc(
    document: str,
    action: str = "insert",
    position: str = "start",
    paragraph: int = 0,
    levels_from: int = 1,
    levels_to: int = 3,
    page_numbers: bool = True,
    hyperlinks: bool = True,
    table_index: int = 0,
) -> dict:
    """Insert, refresh or remove a table of contents built from the Heading styles.

    Args:
        document: exact document name (required).
        action: 'insert' | 'update' (refresh entries and page numbers) | 'delete' | 'list'.
        position: for insert - 'start', 'end', 'after_paragraph' / 'before_paragraph' (with `paragraph`) or 'after_table' / 'before_table' (with `table_index` - no need to count paragraphs).
        paragraph: paragraph number for the *_paragraph positions.
        levels_from, levels_to: heading levels included (1-9).
        page_numbers, hyperlinks: TOC options.
    """
    app, doc = pick_document(document)
    act = action.lower()
    tocs = doc.TablesOfContents
    count = int(tocs.Count)
    if act == "list":
        return {"document": doc.Name, "tocs": count}
    if act == "update":
        for i in range(1, count + 1):
            tocs(i).Update()
        return {"ok": True, "document": doc.Name, "updated": count}
    if act == "delete":
        for i in range(count, 0, -1):
            tocs(i).Delete()
        return {"ok": True, "document": doc.Name, "deleted": count}
    if act != "insert":
        raise ToolError("action must be 'insert', 'update', 'delete' or 'list'.")
    anchor = ensure_free_anchor(doc, position, paragraph, table_index=table_index)
    # Add(Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel)
    toc = tocs.Add(anchor, True, int(levels_from), int(levels_to))
    toc.IncludePageNumbers = bool(page_numbers)
    toc.UseHyperlinks = bool(hyperlinks)
    toc.Update()
    return {"ok": True, "document": doc.Name, "tocs": int(doc.TablesOfContents.Count)}


@office_tool("word_layout", "write", title="Update fields")
def word_update_fields(document: str) -> dict:
    """Refresh every field of the document: page numbers, total pages, table of contents, cross-references, dates."""
    app, doc = pick_document(document)
    n = int(doc.Fields.Count)
    doc.Fields.Update()
    for i in range(1, int(doc.TablesOfContents.Count) + 1):
        doc.TablesOfContents(i).Update()
    doc.Repaginate()
    return {"ok": True, "document": doc.Name, "fields_updated": n, "pages": int(doc.ComputeStatistics(2))}


# ================================================================== картинки и разрывы


def insert_picture(doc, path, position="end", paragraph=0, width_cm=None, alt_text="", alignment="center", caption="", trusted=False, table_index=0) -> dict:
    """Вставляет файл-картинку отдельным абзацем (общий код word_insert_image и мостов Excel->Word)."""
    full = os.path.abspath(path) if trusted else check_path(path, "read")  # trusted — наш собственный временный файл
    if not os.path.isfile(full):
        raise ToolError(f"Image file not found: {full}")
    if full.rsplit(".", 1)[-1].lower() not in IMAGE_EXTS:
        raise ToolError(f"Unsupported image type; allowed: {sorted(IMAGE_EXTS)}")
    if alignment.lower() not in ALIGNMENTS:
        raise ToolError("alignment must be left, center or right.")
    anchor = ensure_free_anchor(doc, position, paragraph, table_index=table_index)
    # AddPicture(FileName, LinkToFile, SaveWithDocument, Range)
    shape = doc.InlineShapes.AddPicture(full, False, True, anchor)
    shape.LockAspectRatio = True
    page = doc.Sections(1).PageSetup
    avail = float(page.PageWidth) - float(page.LeftMargin) - float(page.RightMargin)
    target = cm_to_points(width_cm) if width_cm else min(float(shape.Width), avail)
    ratio = float(shape.Height) / float(shape.Width) if float(shape.Width) else 1
    shape.Width = target
    shape.Height = target * ratio
    if alt_text:
        shape.AlternativeText = alt_text
    shape.Range.ParagraphFormat.Alignment = ALIGNMENTS[alignment.lower()]
    if caption:
        e = int(shape.Range.Paragraphs(1).Range.End) - 1
        doc.Range(e, e).InsertAfter("\r" + caption)
        doc.Range(e + 1, e + 1).Paragraphs(1).Style = -35  # Caption
    return {"ok": True, "document": doc.Name, "width_cm": round(float(shape.Width) * 2.54 / 72, 2), "height_cm": round(float(shape.Height) * 2.54 / 72, 2), "paragraph": paragraph_index_at(doc, int(shape.Range.Start))}


@office_tool("word_layout", "write", title="Insert picture")
def word_insert_image(
    document: str,
    path: str,
    position: str = "end",
    paragraph: int = 0,
    width_cm: float | None = None,
    alt_text: str = "",
    alignment: str = "center",
    caption: str = "",
    table_index: int = 0,
) -> dict:
    """Insert a picture file (png/jpg/gif/bmp/emf) as its own paragraph.

    Args:
        document: exact document name (required).
        path: full path of the image.
        position: 'end', 'start', 'after_paragraph' / 'before_paragraph' (with `paragraph`) or 'after_table' / 'before_table' (with `table_index` - no need to count paragraphs).
        paragraph: paragraph number for the *_paragraph positions.
        width_cm: picture width in centimetres (aspect ratio kept); default - native size scaled to fit the page width.
        alt_text: accessibility description.
        alignment: left|center|right.
        caption: optional caption paragraph under the picture (style 'Caption').
    """
    app, doc = pick_document(document)
    return insert_picture(doc, path, position, paragraph, width_cm, alt_text, alignment, caption, False, table_index)


@office_tool("word_layout", "write", title="Insert break")
def word_insert_break(document: str, kind: str = "page", position: str = "end", paragraph: int = 0) -> dict:
    """Insert a page, section or column break.

    Args:
        document: exact document name (required).
        kind: 'page', 'section_next_page', 'section_continuous', 'column' or 'line'.
        position: 'end', 'start', 'after_paragraph', 'before_paragraph' (with `paragraph`) or 'selection'.
        paragraph: paragraph number for the *_paragraph positions.
    """
    app, doc = pick_document(document)
    codes = {"page": 7, "section_next_page": 2, "section_continuous": 3, "column": 8, "line": 6}
    if kind.lower() not in codes:
        raise ToolError(f"kind must be one of {sorted(codes)}")
    r = collapsed_at(doc, position, paragraph)
    r.InsertBreak(codes[kind.lower()])
    return {"ok": True, "document": doc.Name, "kind": kind, "pages": int(doc.ComputeStatistics(2)), "sections": int(doc.Sections.Count)}


# ================================================================== комментарии, исправления, закладки


def _target_range(doc, paragraph: int, find_text: str, occurrence: int = 1):
    if find_text:
        hits = list(find_all(doc.Content, escape_find(find_text), False, False, False, limit=max(1, int(occurrence))))
        if len(hits) < int(occurrence):
            raise ToolError(f"Text '{find_text}' (occurrence {occurrence}) not found.")
        return hits[int(occurrence) - 1]
    if paragraph:
        r = paragraphs_range(doc, int(paragraph))
        r.MoveEnd(1, -1)  # без маркера абзаца
        return r
    raise ToolError("Pass `paragraph` or `find_text` to say where.")


@office_tool("word_layout", "write", title="Comments")
def word_manage_comments(
    document: str,
    action: str = "list",
    text: str = "",
    paragraph: int = 0,
    find_text: str = "",
    occurrence: int = 1,
    index: int = 0,
    author: str = "",
) -> dict:
    """Review comments: list, add (on a paragraph or on found text), reply, resolve, delete.

    Args:
        document: exact document name (required for changes).
        action: 'list' | 'add' | 'reply' (to comment `index`) | 'resolve' (mark comment `index` done) | 'delete' (comment `index`, or all when index=0).
        text: comment text for add/reply.
        paragraph / find_text (+ occurrence): where to attach a new comment.
        index: comment number from 'list'.
        author: optional author name for a new comment.
    """
    app, doc = pick_document(document)
    act = action.lower()
    n = int(doc.Comments.Count)
    if act == "list":
        items = []
        for i in range(1, min(n, 300) + 1):
            c = doc.Comments(i)
            items.append({
                "index": i, "author": c.Author, "anchored_text": truncate(clean_word_text(c.Scope.Text), 120)[0],
                "comment": clean_word_text(c.Range.Text), "resolved": bool(c.Done), "paragraph": paragraph_index_at(doc, int(c.Scope.Start)),
            })
        return {"document": doc.Name, "count": n, "comments": items}
    if act == "add":
        if not text:
            raise ToolError("'text' is required.")
        c = doc.Comments.Add(_target_range(doc, paragraph, find_text, occurrence), text)
        if author:
            c.Author = author
        return {"ok": True, "document": doc.Name, "comments": int(doc.Comments.Count)}
    if not 0 <= int(index) <= n or (act in ("reply", "resolve") and not index):
        raise ToolError(f"'index' must be between 1 and {n}.")
    if act == "reply":
        if not text:
            raise ToolError("'text' is required.")
        c = doc.Comments(int(index))
        c.Replies.Add(c.Scope, text)
        return {"ok": True, "document": doc.Name, "replied_to": int(index)}
    if act == "resolve":
        try:
            doc.Comments(int(index)).Done = True
        except pywintypes.com_error:
            raise ToolError("This Word build does not allow marking comments as resolved through automation. Reply to the comment or delete it instead.") from None
        return {"ok": True, "document": doc.Name, "resolved": int(index)}
    if act == "delete":
        if index:
            doc.Comments(int(index)).Delete()
        else:
            for i in range(n, 0, -1):
                doc.Comments(i).Delete()
        return {"ok": True, "document": doc.Name, "comments_left": int(doc.Comments.Count)}
    raise ToolError("action must be 'list', 'add', 'reply', 'resolve' or 'delete'.")


@office_tool("word_layout", "write", title="Track changes")
def word_track_changes(document: str, action: str = "status", index: int = 0, author: str = "") -> dict:
    """Control 'Track Changes' and review tracked revisions.

    Args:
        document: exact document name (required for changes).
        action: 'status' | 'enable' | 'disable' | 'list' | 'accept' (revision `index`) | 'reject' (revision `index`) | 'accept_all' | 'reject_all'.
        index: revision number from 'list'.
        author: for accept_all/reject_all - only revisions by this author.
    """
    app, doc = pick_document(document)
    act = action.lower()
    n = int(doc.Revisions.Count)
    kinds = {1: "insert", 2: "delete", 3: "property", 4: "paragraph_number", 5: "display_field", 6: "reconcile", 7: "conflict", 8: "style", 9: "replace", 10: "paragraph_property", 11: "table_property", 12: "section_property", 13: "style_definition", 14: "moved_from", 15: "moved_to", 16: "cell_insertion", 17: "cell_deletion", 18: "cell_merge"}
    if act == "status":
        return {"document": doc.Name, "track_changes_on": bool(doc.TrackRevisions), "revisions": n}
    if act in ("enable", "disable"):
        doc.TrackRevisions = act == "enable"
        return {"ok": True, "document": doc.Name, "track_changes_on": bool(doc.TrackRevisions)}
    if act == "list":
        items = []
        for i in range(1, min(n, 300) + 1):
            r = doc.Revisions(i)
            items.append({"index": i, "type": kinds.get(int(r.Type), int(r.Type)), "author": r.Author, "text": truncate(clean_word_text(r.Range.Text), 120)[0]})
        return {"document": doc.Name, "revisions": n, "items": items}
    if act in ("accept", "reject"):
        if not 1 <= int(index) <= n:
            raise ToolError(f"'index' must be between 1 and {n}.")
        rev = doc.Revisions(int(index))
        rev.Accept() if act == "accept" else rev.Reject()
        return {"ok": True, "document": doc.Name, "revisions_left": int(doc.Revisions.Count)}
    if act in ("accept_all", "reject_all"):
        if author:
            for i in range(n, 0, -1):
                r = doc.Revisions(i)
                if r.Author.lower() == author.lower():
                    r.Accept() if act == "accept_all" else r.Reject()
        elif act == "accept_all":
            doc.Revisions.AcceptAll()
        else:
            doc.Revisions.RejectAll()
        return {"ok": True, "document": doc.Name, "revisions_left": int(doc.Revisions.Count)}
    raise ToolError("action must be 'status', 'enable', 'disable', 'list', 'accept', 'reject', 'accept_all' or 'reject_all'.")


@office_tool("word_layout", "write", title="Bookmarks")
def word_manage_bookmarks(document: str, action: str = "list", name: str = "", paragraph: int = 0, find_text: str = "", occurrence: int = 1) -> dict:
    """List, add or delete bookmarks (named places you can later fill with word_insert_text(position='bookmark')).

    Args:
        document: exact document name (required for changes).
        action: 'list', 'add' (on a paragraph or on found text) or 'delete'.
        name: bookmark name (letters, digits, underscore; starts with a letter).
        paragraph / find_text (+ occurrence): what the new bookmark covers.
    """
    app, doc = pick_document(document)
    act = action.lower()
    if act == "list":
        items = []
        for i in range(1, int(doc.Bookmarks.Count) + 1):
            b = doc.Bookmarks(i)
            items.append({"name": b.Name, "text": truncate(clean_word_text(b.Range.Text), 80)[0], "paragraph": paragraph_index_at(doc, int(b.Range.Start))})
        return {"document": doc.Name, "bookmarks": items}
    if not name:
        raise ToolError("'name' is required.")
    if act == "add":
        if not re.fullmatch(r"[A-Za-zА-Яа-яЁё][A-Za-z0-9_А-Яа-яЁё]{0,39}", name):
            raise ToolError("Bookmark names start with a letter and contain only letters, digits and underscores (max 40).")
        doc.Bookmarks.Add(name, _target_range(doc, paragraph, find_text, occurrence))
        return {"ok": True, "document": doc.Name, "bookmark": name}
    if act == "delete":
        if not doc.Bookmarks.Exists(name):
            raise ToolError(f"Bookmark '{name}' not found.")
        doc.Bookmarks(name).Delete()
        return {"ok": True, "document": doc.Name, "deleted": name}
    raise ToolError("action must be 'list', 'add' or 'delete'.")


@office_tool("word_layout", "write", title="Insert hyperlink")
def word_insert_hyperlink(document: str, url: str, find_text: str = "", paragraph: int = 0, text: str = "", occurrence: int = 1, tooltip: str = "") -> dict:
    """Make text a hyperlink: either link existing text (find_text) or append new linked `text` to the end of a paragraph.

    Args:
        document: exact document name (required).
        url: 'https://...' or 'mailto:...'.
        find_text: existing text to turn into the link (with `occurrence`).
        paragraph + text: alternatively, append `text` as a link at the end of this paragraph.
        tooltip: hover text.
    """
    if not re.match(r"^(https?://|mailto:|ftp://)", url, re.IGNORECASE):
        raise ToolError("Only http(s), ftp and mailto links are allowed.")
    app, doc = pick_document(document)
    if find_text:
        rng = _target_range(doc, 0, find_text, occurrence)
        shown = MISSING
    elif paragraph and text:
        p = paragraphs_range(doc, int(paragraph))
        e = int(p.End) - 1
        rng = doc.Range(e, e)
        shown = text
    else:
        raise ToolError("Pass `find_text`, or both `paragraph` and `text`.")
    # Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
    doc.Hyperlinks.Add(rng, url, MISSING, tooltip or MISSING, shown)
    return {"ok": True, "document": doc.Name, "link": url, "hyperlinks": int(doc.Hyperlinks.Count)}


@office_tool("word_layout", "write", title="Footnotes")
def word_manage_footnotes(document: str, action: str = "list", text: str = "", find_text: str = "", paragraph: int = 0, occurrence: int = 1, index: int = 0) -> dict:
    """List, add or delete footnotes.

    Args:
        document: exact document name (required for changes).
        action: 'list', 'add' (footnote mark after the found text / at the end of a paragraph) or 'delete' (footnote `index`).
        text: footnote text for add.
        find_text / paragraph (+ occurrence): where the mark goes.
        index: footnote number for delete.
    """
    app, doc = pick_document(document)
    act = action.lower()
    n = int(doc.Footnotes.Count)
    if act == "list":
        return {"document": doc.Name, "footnotes": [{"index": i, "text": clean_word_text(doc.Footnotes(i).Range.Text), "paragraph": paragraph_index_at(doc, int(doc.Footnotes(i).Reference.Start))} for i in range(1, min(n, 200) + 1)]}
    if act == "add":
        if not text:
            raise ToolError("'text' is required.")
        rng = _target_range(doc, paragraph, find_text, occurrence)
        rng.Collapse(0)  # метка сноски — после найденного текста
        note = doc.Footnotes.Add(rng)  # текст задаём отдельно: Missing в середине списка аргументов обнуляет остальные
        note.Range.Text = text
        return {"ok": True, "document": doc.Name, "footnotes": int(doc.Footnotes.Count)}
    if act == "delete":
        if not 1 <= int(index) <= n:
            raise ToolError(f"'index' must be between 1 and {n}.")
        doc.Footnotes(int(index)).Delete()
        return {"ok": True, "document": doc.Name, "footnotes": int(doc.Footnotes.Count)}
    raise ToolError("action must be 'list', 'add' or 'delete'.")


@office_tool("word_layout", "write", title="Document properties")
def word_document_properties(document: str, action: str = "get", properties: dict | None = None) -> dict:
    """Read or set built-in document properties (File > Info): Title, Subject, Author, Keywords, Comments, Category, Company, Manager.

    Args:
        document: exact name or '' for the active document (required for set).
        action: 'get' or 'set'.
        properties: for set - e.g. {"Title": "Quarterly report", "Author": "Finance team"}.
    """
    app, doc = pick_document(document)
    names = ["Title", "Subject", "Author", "Keywords", "Comments", "Category", "Company", "Manager"]
    props = doc.BuiltInDocumentProperties
    act = action.lower()
    if act == "set":
        if not properties:
            raise ToolError("'properties' is required.")
        for k, v in properties.items():
            key = next((n for n in names if n.lower() == str(k).lower()), None)
            if key is None:
                raise ToolError(f"Unsupported property '{k}'. Allowed: {names}")
            props(key).Value = str(v)
    elif act != "get":
        raise ToolError("action must be 'get' or 'set'.")
    out = {}
    for n in names:
        try:
            v = props(n).Value
            if v:
                out[n] = str(v)
        except pywintypes.com_error:
            continue
    return {"document": doc.Name, "properties": out}



@office_tool("word_layout", "read", title="Render a page as image", unstructured=True)
def word_render_page_image(document: str = "", page: int = 1, max_width_px: int = 1100) -> list:
    """Render one page of a Word document exactly as laid out (fonts, tables, pictures, headers/footers) and return it as a PNG image - use it to visually verify a document. Needs the Pillow package. The user's document window is switched to Print Layout briefly and restored.

    Args:
        document: exact name or '' for the active document.
        page: 1-based page number (see word_get_structure for the page count).
        max_width_px: width of the returned image (default 1100).
    """
    try:
        from PIL import Image as PILImage
    except ImportError:
        raise ToolError("Rendering a page needs the Pillow package: pip install pillow") from None
    app, doc = pick_document(document)
    prev = active_document(app)
    try:
        doc.Activate()
        win = doc.ActiveWindow
        old_view = win.View.Type
        win.View.Type = 3  # wdPrintView: страницы существуют только в этом режиме
        try:
            pages = win.ActivePane.Pages
            total = int(pages.Count)
            if not 1 <= int(page) <= total:
                raise ToolError(f"Page {page} does not exist; the document has {total} page(s).")
            emf = bytes(pages(int(page)).EnhMetaFileBits)
        finally:
            try:
                win.View.Type = old_view
            except pywintypes.com_error:
                pass
    except pywintypes.com_error as exc:
        raise ToolError("Could not render the page (the document window must be visible): " + com.com_error_text(exc)) from None
    finally:
        if prev is not None and prev.Name != doc.Name:
            try:
                prev.Activate()
            except pywintypes.com_error:
                pass
    im = PILImage.open(io.BytesIO(emf))
    im.load()
    im = im.convert("RGB")
    width = max(300, min(int(max_width_px), 3000))
    if im.width > width:
        im = im.resize((width, round(im.height * width / im.width)), PILImage.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "PNG")
    info = {"ok": True, "document": doc.Name, "page": int(page), "pages": total, "png_bytes": buf.tell(), "size_px": [im.width, im.height]}
    return [json.dumps(info, ensure_ascii=False), Image(data=buf.getvalue(), format="png")]
