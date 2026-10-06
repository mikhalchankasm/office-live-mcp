"""Exact, tracked fragment edits in Word's main story. No LLM, Selection or file writes."""

import hashlib
import json
import unicodedata

import pywintypes

from . import com, config, undo, wd_common as wd
from .errors import PartialChangeError, ToolError
from .registry import office_tool

STALE = "Фрагмент изменился после предпросмотра; перечитайте документ"
FONT_PROPERTIES = ("Name", "Size", "Bold", "Italic", "Underline", "Color")


def _units(text):
    return len(text.encode("utf-16-le")) // 2


def _items(collection):
    return [collection.Item(i) for i in range(1, int(collection.Count) + 1)]


def _overlap(a, b):
    return int(a.Start) < int(b.End) and int(b.Start) < int(a.End)


def _style_name(rng):
    style = rng.Style
    return str(style) if isinstance(style, (str, int)) else str(style.NameLocal)


def _formatting(rng):
    values = {p: getattr(rng.Font, p) for p in FONT_PROPERTIES}
    values["Highlight"] = rng.HighlightColorIndex
    mixed = [p for p, value in values.items() if value == 9999999]
    return {"kind": "mixed" if mixed else "uniform", "mixed_properties": mixed, "properties": values}


def _limitations(doc, paragraph_range, rng, expected_text, new_text):
    reasons = []
    if not expected_text or any(c in expected_text for c in "\r\x07\x0c\x0e\x01"):
        reasons.append("Фрагмент пуст или содержит структурный маркер/объект.")
    if len(new_text) > 10000 or any(unicodedata.category(c) == "Cc" and c not in "\t\v" for c in new_text):
        reasons.append("new_text: максимум 10000 символов; управляющие разрешены только tab и разрыв строки (\\v).")
    if int(rng.StoryType) != 1:
        reasons.append("Правка разрешена только в основном тексте документа.")
    for member in ("Footnotes", "Endnotes", "Fields", "Hyperlinks", "ContentControls", "InlineShapes", "Revisions"):
        if int(getattr(rng, member).Count):
            reasons.append(f"Фрагмент содержит {member}.")
    # Range.Fields may miss a selection wholly or partially inside a field's Code/Result.
    for field in _items(paragraph_range.Fields) + _items(doc.Fields):
        if int(field.Code.StoryType) != 1:
            continue
        span = doc.Range(int(field.Code.Start) - 1, int(field.Result.End) + 1)
        if _overlap(rng, span):
            reasons.append("Фрагмент пересекает поле (включая код/результат).")
    for member in ("Hyperlinks", "ContentControls"):
        for obj in _items(getattr(doc, member)):
            span = obj.Range
            if int(span.StoryType) == 1 and _overlap(rng, span):
                reasons.append(f"Фрагмент пересекает {member}.")
    for bookmark in _items(doc.Bookmarks):
        span = bookmark.Range
        if int(span.StoryType) != 1:
            continue
        a, b, start, end = int(span.Start), int(span.End), int(rng.Start), int(rng.End)
        if start < a < end or start < b < end or start <= a <= b <= end:
            reasons.append("Фрагмент удалит закладку или пересечёт её границу.")
    for comment in _items(doc.Comments):
        span = comment.Scope
        if int(span.StoryType) == 1 and _overlap(rng, span) and not (int(rng.Start) <= int(span.Start) <= int(span.End) <= int(rng.End)):
            reasons.append("Фрагмент частично пересекает привязку примечания.")
    if int(doc.ProtectionType) not in {-1, 0}:
        reasons.append("Документ защищён: снимите ограничение word_restrict_editing action=unprotect.")
    try:
        if config.SETTINGS.autosave == "block" and bool(doc.AutoSaveOn):
            reasons.append("AutoSave включён: отключите его перед правкой.")
    except (AttributeError, pywintypes.com_error):
        pass
    return list(dict.fromkeys(reasons))


def _prepare(document, paragraph, expected_text, occurrence, new_text):
    if config.SETTINGS.strict_target and not document.strip():
        raise ToolError("Document name is required (OFFICE_LIVE_STRICT_TARGET is on).")
    if type(paragraph) is not int or paragraph < 1 or type(occurrence) is not int or occurrence < 1:
        raise ToolError("paragraph and occurrence must be positive integers.")
    app, doc = wd.pick_document(document, exact_only=True)
    if paragraph > int(doc.Paragraphs.Count):
        raise ToolError(STALE)
    para = doc.Paragraphs(paragraph).Range
    text = str(para.Text).removesuffix("\x07").removesuffix("\r")
    offset, cursor = -1, 0
    for _ in range(occurrence):
        offset = text.find(expected_text, cursor)
        if offset < 0:
            raise ToolError(STALE)
        cursor = offset + max(1, len(expected_text))
    start = int(para.Start) + _units(text[:offset])
    rng = doc.Range(start, start + _units(expected_text))
    fingerprint = hashlib.sha256(json.dumps([wd.document_path(doc) or doc.Name, text, start, expected_text],
                                           ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest()
    try:
        limitations = _limitations(doc, para, rng, expected_text, new_text)
    except (AttributeError, pywintypes.com_error) as exc:
        if com.is_busy(exc) or com.is_dead(exc):
            raise
        limitations = ["Не удалось проверить структурные ограничения фрагмента; правка отклонена."]
    formatting = _formatting(rng)
    warnings = ["Смешанное форматирование внутри фрагмента не сохраняется; новый текст наследует первый символ."] if formatting["kind"] == "mixed" else []
    link = wd.range_link(doc, paragraph, rng)
    result = {"document": doc.Name, "location": wd.range_location(doc, rng, paragraph), "original": expected_text,
              "new_text": new_text, "context_before": text[max(0, offset - 80):offset],
              "context_after": text[offset + len(expected_text):offset + len(expected_text) + 80],
              "fingerprint": fingerprint, "link": link, "links": [{"label": f"{doc.Name}: paragraph {paragraph}", "uri": link}],
              "formatting": formatting["kind"], "formatting_details": formatting,
              "limitations": limitations, "will_track_changes": True, "blocked": bool(limitations),
              "blocked_reason": "; ".join(limitations) or None, "warnings": warnings}
    return app, doc, para, rng, result


@office_tool("word_core", "read", title="Preview tracked fragment edit")
def word_edit_preview(document: str, paragraph: int, expected_text: str, new_text: str, occurrence: int = 1) -> dict:
    """Preview an exact fragment edit in a main-story paragraph without changes, journal or undo.

    Args:
        document: exact open document name; required in strict mode even for preview.
        paragraph: 1-based main-story paragraph; tables support plain text fragments only.
        expected_text: exact current fragment, case/whitespace sensitive; required, nonempty.
        new_text: proposed text (empty deletes), max 10000 characters, tab and vertical-tab allowed.
        occurrence: 1-based non-overlapping occurrence within this paragraph.

    Return fingerprint to word_edit_apply after user agreement. A preview is not permission to apply.
    Structural objects, existing revisions and partial comment anchors block edits; mixed formatting inherits the first character.
    """
    return _prepare(document, paragraph, expected_text, occurrence, new_text)[-1]


@office_tool("word_core", "write", title="Apply tracked fragment edit")
def word_edit_apply(document: str, paragraph: int, expected_text: str, new_text: str, occurrence: int = 1, fingerprint: str = "") -> dict:
    """Apply an agreed exact fragment replacement as tracked changes; restores TrackRevisions even on failure.

    Args:
        document, paragraph, expected_text, new_text, occurrence: same target/text as word_edit_preview.
        fingerprint: pass the preview fingerprint to reject stale paragraphs/shifted occurrences before writing.

    Uses one office_undo record; no Accept/Reject, saving, Selection or window changes. New text inherits the first
    character's font and paragraph style/numbering; mixed inline formatting is lost. Re-read after a stale target error.
    """
    app, doc, para, rng, result = _prepare(document, paragraph, expected_text, occurrence, new_text)
    if fingerprint and fingerprint != result["fingerprint"]:
        raise ToolError(STALE)
    if result["blocked"]:
        raise ToolError(result["blocked_reason"])
    old_tracking = bool(doc.TrackRevisions)
    old_style, old_list = _style_name(para), str(para.ListFormat.ListString)
    paragraphs, revisions = int(doc.Paragraphs.Count), int(doc.Revisions.Count)
    undo.require_undo("document", app, doc)
    try:
        try:
            doc.TrackRevisions = True
            # Range.Text inherits the first character's formatting; re-applying Font here would also format the
            # tracked deletion and add formatting revisions.
            rng.Text = new_text
        finally:
            doc.TrackRevisions = old_tracking
    except pywintypes.com_error as exc:
        raise PartialChangeError(f"Tracked fragment edit interrupted; {com.translate(exc)}") from None
    try:
        after = doc.Paragraphs(paragraph).Range
        for changed, warning in ((_style_name(after) != old_style, "Стиль абзаца изменился."),
                                 (str(after.ListFormat.ListString) != old_list, "Нумерация абзаца изменилась."),
                                 (int(doc.Paragraphs.Count) != paragraphs, "Число абзацев изменилось."),
                                 (bool(new_text) and new_text not in str(after.Text), "Контрольное чтение не нашло новый текст.")):
            if changed:
                result["warnings"].append(warning)
        result.update(ok=True, revisions_added=int(doc.Revisions.Count) - revisions,
                      location=wd.range_location(doc, rng, paragraph), link=wd.range_link(doc, paragraph, rng))
        result["links"][0]["uri"] = result["link"]
    except pywintypes.com_error as exc:
        raise PartialChangeError(f"Tracked edit completed but read-back failed; {com.translate(exc)}") from None
    return result
