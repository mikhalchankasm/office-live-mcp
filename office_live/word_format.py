"""Word text formatting, styles and lists."""

import math
from . import com, undo
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .util import cm_to_points, parse_color
from .wd_common import (
    ALIGNMENTS,
    HIGHLIGHTS,
    escape_find,
    find_all,
    paragraphs_range,
    pick_document,
    resolve_style,
)


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
