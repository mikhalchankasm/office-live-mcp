"""Word: разделы, водяной знак, подписи и перекрёстные ссылки (регистрирует word_layout в прежнем порядке каталога)."""

import posixpath
import re
import uuid
from xml.etree import ElementTree as ET

import pythoncom

from . import com, undo
from . import wd_common as wd
from .errors import PartialChangeError, ToolError
from .util import clean_word_text, parse_color
from .wd_common import paragraph_indexes, paragraphs_range


HF_TYPES = {"primary": 1, "first_page": 2, "even": 3}

# wdSectionBreakNextPage / Continuous / EvenPage / OddPage
SECTION_BREAKS = {"next_page": 2, "continuous": 3, "even_page": 4, "odd_page": 5}
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_PAGE_DIMS = ("PageWidth", "PageHeight", "TopMargin", "BottomMargin", "LeftMargin", "RightMargin")


def _section_xml(doc):
    """Только XML: даже проверка HeaderFooter.Exists меняет документ."""
    root = ET.fromstring(doc.Content.WordOpenXML)
    pkg = "{http://schemas.microsoft.com/office/2006/xmlPackage}"
    rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    parts = {p.get(pkg + "name"): p for p in root.iter(pkg + "part")}
    relationships = parts.get("/word/_rels/document.xml.rels")
    targets = {} if relationships is None else {r.get("Id"): posixpath.normpath(posixpath.join("/word", r.get("Target", "")))
                                               for r in relationships.iter(rel + "Relationship")}
    anchored = {path for path, part in parts.items() if any(
        (n.tag.rsplit("}", 1)[-1] in {"positionH", "positionV"} and n.get("relativeFrom") == "page")
        or re.search(r"mso-position-(?:horizontal|vertical)-relative\s*:\s*page", n.get("style", "")) for n in part.iter())}
    body = next(root.iter(_W + "body"), None)
    if body is None:
        raise ToolError("Cannot read section properties from WordOpenXML.")
    def walk(node):
        if node.tag in {_W + "txbxContent", _W + "sectPrChange"}:
            return
        if node.tag == _W + "sectPr":
            yield node
        else:
            for child in node:
                yield from walk(child)
    properties = list(walk(body))
    if len(properties) != int(doc.Sections.Count):
        raise ToolError("XML and Word section counts disagree; no changes made.")
    result, inherited = [], {}
    for index, prop in enumerate(properties, 1):
        refs = {(k, {"default": 1, "first": 2, "even": 3}[r.get(_W + "type")]): r.get(rid)
                for k in ("header", "footer") for r in prop.findall(_W + k + "Reference")}
        inherited |= refs
        number = prop.find(_W + "pgNumType")
        start = number.get(_W + "start") if number is not None else None
        result.append({"linked_to_previous": {f"{k}_{n}": index > 1 and (k, t) not in refs
                                              for k in ("header", "footer") for n, t in HF_TYPES.items()},
                       "restart_numbering_at_section": start is not None,
                       "starting_number": int(start) if start is not None else None,
                       "_existing": sorted(inherited),
                       "_header_page_shapes": any(targets.get(r) in anchored for r in inherited.values())})
    return result


def _section_list(doc):
    result, sections = [], _section_xml(doc)
    spans = [(int(s.Range.Start), int(s.Range.End) - 1) for s in _items(doc.Sections)]
    paragraphs = paragraph_indexes(doc, [p for span in spans for p in span])
    for index, meta in enumerate(sections, 1):
        ps, (start, last) = doc.Sections(index).PageSetup, spans[index - 1]
        result.append({"section": index, "start_paragraph": paragraphs[start], "end_paragraph": paragraphs[last],
                       "orientation": "landscape" if int(ps.Orientation) == 1 else "portrait",  # wdOrientLandscape
                       "paper_size_id": int(ps.PaperSize),
                       "page_cm": [round(float(getattr(ps, n)) * 2.54 / 72, 3) for n in _PAGE_DIMS[:2]],
                       "margins_cm": {n: round(float(getattr(ps, n)) * 2.54 / 72, 3) for n in _PAGE_DIMS[2:]},
                       "different_first_page": bool(ps.DifferentFirstPageHeaderFooter),
                       **{k: v for k, v in meta.items() if not k.startswith("_")}})
    return result


def word_sections(document: str, action: str = "list", paragraph: int = 0, kind: str = "next_page",
                  start_paragraph: int = 0, end_paragraph: int = 0, orientation: str = "landscape") -> dict:
    """List sections, insert a section break before a paragraph, or orient an inclusive paragraph range.

    action: list (readonly), insert_break, orientation. kind: next_page, continuous, even_page, odd_page.
    orientation: landscape or portrait; inserts only missing next-page boundaries. Whole tables may be enclosed,
    but boundaries inside tables are refused. Rotates paper and margin pairs; preserves existing section links
    and restarts, while newly split continuations inherit headers/footers and continue numbering.
    List uses XML without touching headers; starting_number is null when numbering continues.
    Returns inserted breaks (original paragraph positions), resulting sections and page-anchored shape warnings.
    """
    action, kind, orientation = action.lower(), kind.lower(), orientation.lower()
    if action not in {"list", "insert_break", "orientation"}:
        raise ToolError("action must be list, insert_break or orientation.")
    app, doc = wd.pick_document(document)
    if action == "list":
        return {"document": doc.Name, "sections": _section_list(doc)}
    if kind not in SECTION_BREAKS or orientation not in {"landscape", "portrait"}:
        raise ToolError("Invalid section break kind or orientation.")
    if action == "orientation" and start_paragraph > end_paragraph:
        raise ToolError("start_paragraph must not exceed end_paragraph.")
    rng = paragraphs_range(doc, paragraph if action == "insert_break" else start_paragraph,
                           paragraph if action == "insert_break" else end_paragraph)
    start, end = int(rng.Start), int(rng.End)
    meta = _section_xml(doc)
    sections = [(i, int(s.Range.Start), int(s.Range.End), s.PageSetup)
                for i, s in enumerate(_items(doc.Sections), 1)]
    boundaries = {s for _, s, _, _ in sections} | {int(doc.Content.End)}
    breaks = [(start, paragraph, kind)] if action == "insert_break" else [
        (p, n, "next_page") for p, n in ((start, start_paragraph), (end, end_paragraph + 1)) if p not in boundaries]
    for table in _items(doc.Tables):
        a, b = int(table.Range.Start), int(table.Range.End)
        # A break inserted at a table's first position lands in cell (1,1): refuse it like a boundary inside the table.
        if a <= start < b or (action == "orientation" and a < end < b) or any(a <= p < b for p, _, _ in breaks):
            raise ToolError("Section boundaries inside or partially covering a table are unsupported.")
    target = 1 if orientation == "landscape" else 0  # wdOrientLandscape / wdOrientPortrait
    pages = {i: tuple(float(getattr(ps, n)) for n in _PAGE_DIMS)
             for i, a, b, ps in sections if a < end and b > start and int(ps.Orientation) != target}
    warnings = [f"Page-anchored floating shape in header/footer of section {i} may move."
                for i, a, b, _ in sections if a < end and b > start and meta[i - 1]["_header_page_shapes"]]
    for shape in _items(doc.Shapes):
        anchor = shape.Anchor
        if int(anchor.StoryType) == 1 and any(a <= int(anchor.Start) < b for _, a, b, _ in sections if a < end and b > start):  # wdMainTextStory
            if int(shape.RelativeHorizontalPosition) == 1 or int(shape.RelativeVerticalPosition) == 1:  # wdRelativeHorizontal/VerticalPositionPage
                warnings.append("Page-anchored floating shape may move: " + str(shape.Name))
    name, inserted, changed, shifts = str(doc.Name), [], [], {}
    if action == "orientation" and not breaks and not pages:
        return {"document": name, "inserted_breaks": [], "sections": _section_list(doc), "warnings": warnings,
                "range_sections": [i for i, a, b, _ in sections if a < end and b > start]}
    undo.require_undo("document", app, doc)
    stage = "section break insertion attempted; completion unknown"
    try:
        # С конца: исходные позиции ещё не сдвинуты. Word сам копирует свойства раздела.
        for pos, para, typ in sorted(breaks, reverse=True):
            before_end = int(doc.Content.End)
            doc.Range(pos, pos).InsertBreak(SECTION_BREAKS[typ])
            # The break ends a new empty paragraph that inherits the next paragraph's formatting; a heading with
            # "page break before" or numbering would add a blank page and a phantom number/TOC entry.
            holder = doc.Range(pos, pos).Paragraphs(1).Range
            if str(holder.Text).strip("\x0c\r") == "":
                holder.Style = -1  # wdStyleNormal
                holder.ParagraphFormat.PageBreakBefore = False
            inserted.append({"before_paragraph": para, "kind": typ})
            stage = f"inserted breaks={inserted}; next operation attempted"
            shifts[pos] = int(doc.Content.End) - before_end
        moved = {p: p + sum(delta for q, delta in shifts.items() if q <= p) for p in shifts}
        # Новые разделы справа от разрыва должны продолжать нумерацию исходного раздела.
        for pos, _, _ in breaks:
            shifted = moved[pos]
            index = next(i for i in range(1, int(doc.Sections.Count) + 1) if int(doc.Sections(i).Range.Start) == shifted)
            source = next(i for i, a, b, _ in sections if a <= pos < b)
            sec = doc.Sections(index)
            existing = meta[source - 1]["_existing"]
            # Если номеров ещё нет в колонтитулах, сбрасываем restart через основной footer.
            if not existing and meta[source - 1]["restart_numbering_at_section"]:
                existing = [("footer", 1)]  # wdHeaderFooterPrimary; только изменяющее действие
            for k, typ in existing:
                hf = _hf(sec, k, typ)
                hf.LinkToPrevious = True
                hf.PageNumbers.RestartNumberingAtSection = False
        if action == "orientation":
            lo = start + sum(delta for p, delta in shifts.items() if p <= start)
            hi = end + sum(delta for p, delta in shifts.items() if p < end)
            for i in range(1, int(doc.Sections.Count) + 1):
                sec = doc.Sections(i)
                a, b = int(sec.Range.Start), int(sec.Range.End)
                if a >= hi or b <= lo:
                    continue
                original = a - sum(shifts[p] for p, shifted in moved.items() if shifted <= a)
                source = next(j for j, x, y, _ in sections if x <= original < y)
                if source in pages:
                    width, height, top, bottom, left, right = pages[source]
                    stage = f"inserted breaks={inserted}; changed sections={changed}; setting section {i} attempted"
                    sec.PageSetup.Orientation = target
                    for key, value in zip(_PAGE_DIMS, (height, width, left, right, top, bottom), strict=True):
                        setattr(sec.PageSetup, key, value)
                    changed.append(i)
        doc.Repaginate()  # расчёт страниц до итогового отпечатка отмены
        result = _section_list(doc)
        return {"ok": True, "document": name, "inserted_breaks": sorted(inserted, key=lambda b: b["before_paragraph"]),
                "sections": result, "range_sections": [r["section"] for r in result if action == "orientation"
                    and int(doc.Sections(r["section"]).Range.Start) < hi and int(doc.Sections(r["section"]).Range.End) > lo],
                "warnings": warnings}
    except Exception as exc:
        raise PartialChangeError(f"Section operation interrupted: {stage}. Use office_undo to undo this call. {exc}") from None


def _hf(sec, kind: str, typ: int):
    return sec.Headers(typ) if kind == "header" else sec.Footers(typ)


# ================================================================== водяной знак

WATERMARK_PREFIX = "PowerPlusWaterMarkObject"
HEADER_STORIES = {1: 7, 2: 10, 3: 6}  # WdHeaderFooterIndex -> wdPrimary/FirstPage/EvenPagesHeaderStory


def _header_shape_names(doc):
    """Names of all shapes in existing header/footer parts, from WordOpenXML (no HeaderFooter COM access)."""
    names = set()
    root = ET.fromstring(doc.Content.WordOpenXML)
    for part in root.findall("pkg:part", _WM_NS):
        if re.match(r"/word/(header|footer)\d*\.xml$", part.get("{" + _WM_NS["pkg"] + "}name", "")):
            names.update(node.get("id", "") for node in part.iter("{" + _WM_NS["v"] + "}shape"))
            names.update(node.get("name", "") for node in part.iter("{" + _WM_NS["wp"] + "}docPr"))
    return names - {""}


def _in_header(shape, owner, typ):
    """HeaderFooter.Shapes returns the shapes of ALL headers and footers of the document (Word quirk): keep only the
    ones anchored in this section's header of this type. Unreadable anchor -> not ours (never delete blindly)."""
    try:
        anchor = shape.Anchor
        return int(anchor.StoryType) == HEADER_STORIES[typ] and int(anchor.Sections(1).Index) == owner
    except Exception:  # noqa: BLE001
        return False
_WM_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "pkg": "http://schemas.microsoft.com/office/2006/xmlPackage",
          "v": "urn:schemas-microsoft-com:vml",
          "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}


def _watermark_headers(doc):
    """Resolve existing and inherited headers from Flat OPC, without touching HeaderFooter COM objects."""
    def attr(node, name):
        prefix, local = name.split(":")
        return node.get("{" + _WM_NS[prefix] + "}" + local)

    def on(node):
        return node is not None and attr(node, "w:val") not in {"0", "false", "off"}

    root = ET.fromstring(doc.Content.WordOpenXML)
    parts = {attr(p, "pkg:name"): p for p in root.findall("pkg:part", _WM_NS)}
    main = parts.get("/word/document.xml")
    if main is None:
        raise ToolError("Cannot resolve document headers from WordOpenXML; no changes made.")
    rel_part = parts.get("/word/_rels/document.xml.rels")
    rels = {r.get("Id"): posixpath.normpath(posixpath.join("/word", r.get("Target", "")))
            for r in rel_part.iter() if r.tag.endswith("}Relationship") and r.get("TargetMode") != "External"} if rel_part is not None else {}
    settings = parts.get("/word/settings.xml")
    even = settings is not None and on(settings.find(".//w:evenAndOddHeaders", _WM_NS))
    body = main.find(".//w:body", _WM_NS)
    specs = body.findall("w:p/w:pPr/w:sectPr", _WM_NS) + body.findall("w:sectPr", _WM_NS) if body is not None else []
    if len(specs) != int(doc.Sections.Count):
        raise ToolError("Cannot resolve all sections from WordOpenXML; no changes made.")
    current, groups = {}, {}
    for index, spec in enumerate(specs, 1):
        refs = {attr(r, "w:type"): attr(r, "r:id") for r in spec.findall("w:headerReference", _WM_NS)}
        for kind, typ in (("default", 1), ("first", 2), ("even", 3)):  # WdHeaderFooterIndex
            if kind in refs:
                path = rels.get(refs[kind])
                if path not in parts:
                    raise ToolError("Cannot resolve a header relationship; no changes made.")
                current[kind] = (path, index)
            path, owner = current.get(kind, (None, 1))
            key = (path or owner, typ)
            if key not in groups:
                marks = {}
                if path:
                    header = parts[path]
                    for node in header.findall(".//v:shape", _WM_NS) + header.findall(".//wp:anchor", _WM_NS):
                        prop = node.find("wp:docPr", _WM_NS)
                        name = node.get("id", "") if prop is None else prop.get("name", "")
                        if name.startswith(WATERMARK_PREFIX):
                            textpath = node.find("v:textpath", _WM_NS)
                            text = textpath.get("string", "") if textpath is not None else "".join(n.text or "" for n in node.iter() if n.tag.endswith("}t"))
                            marks[name] = {"name": name, "text": text}
                groups[key] = {"owner": owner, "type": typ, "header": kind, "sections": [], "enabled": [], "marks": list(marks.values())}
            group = groups[key]
            group["sections"].append(index)
            if kind == "default" or (kind == "first" and on(spec.find("w:titlePg", _WM_NS))) or (kind == "even" and even):
                group["enabled"].append(index)
    return list(groups.values()), len(specs)


def word_watermark(document: str, action: str = "status", text: str = "", font: str = "Calibri",
                   color: str = "gray", transparency: float = 0.5, layout: str = "diagonal",
                   sections: list[int] | None = None) -> dict:
    """Inspect, add or remove Word-recognised text watermarks in section headers.

    Args:
        document: exact document name/path.
        action: status (readonly, XML only), add, remove. Remove deletes only PowerPlusWaterMarkObject* shapes.
        text: 1–50 printable characters for add; existing watermarks must be removed first.
        font: WordArt font (default Calibri).
        color: #RRGGBB or named colour (default gray).
        transparency: 0 (opaque) to 1 (invisible), default 0.5.
        layout: diagonal or horizontal.
        sections: 1-based numbers, omitted for all. Include the entire linked header chain; links are never broken.
    """
    action = action.lower()
    if action not in {"status", "add", "remove"}:
        raise ToolError("action must be status, add or remove.")
    rgb = None
    if action == "add":
        if not text.strip() or len(text) > 50 or not text.isprintable():
            raise ToolError("text must contain 1–50 printable characters.")
        if not font.strip() or not font.isprintable():
            raise ToolError("font must be a non-empty printable font name.")
        if layout not in {"diagonal", "horizontal"} or not 0 <= transparency <= 1:
            raise ToolError("layout must be diagonal/horizontal; transparency must be 0–1.")
        rgb = parse_color(color)
    app, doc = wd.pick_document(document)
    groups, count = _watermark_headers(doc)
    if sections is not None and (not sections or any(type(i) is not int or not 1 <= i <= count for i in sections)):
        raise ToolError(f"sections must be a non-empty list of section numbers 1–{count}.")
    chosen = set(range(1, count + 1) if sections is None else sections)
    marks = [{**m, "sections": sorted(chosen.intersection(g["sections"])), "header": g["header"]}
             for g in groups if chosen.intersection(g["sections"]) for m in g["marks"]]
    if action == "status":
        return {"document": doc.Name, "present": bool(marks), "watermarks": marks}
    if int(doc.ProtectionType) != -1:  # wdNoProtection; tracked-only protection cannot edit header shapes
        raise ToolError("Document is protected; unprotect it before changing watermarks.")
    if bool(doc.ReadOnly):
        raise ToolError("Document is read-only; no changes made.")
    if action == "add" and marks:
        raise ToolError("A watermark already exists in the selected sections; remove it first.")
    targets = [g for g in groups if chosen.intersection(g["enabled"] if action == "add" else g["sections"])
               and (action == "add" or g["marks"])]
    for g in targets:
        if not set(g["sections"]) <= chosen:
            raise ToolError(f"Linked header affects sections {g['sections']}; include the whole chain in sections.")
        if action == "add":
            pages = [doc.Sections(i).PageSetup for i in g["enabled"]]
            g["max_width"] = .8 * min(min(float(p.PageWidth), float(p.PageHeight)) for p in pages)
    if not targets:
        return {"ok": True, "document": doc.Name, "action": action, "changed": []}
    name = doc.Name
    plan = {"_watermark_add": [(g["owner"], g["type"], WATERMARK_PREFIX + str(g["owner"] * 10 + g["type"]))
                               for g in targets],
            "_header_shapes_before": sorted(_header_shape_names(doc))} if action == "add" else None
    undo.require_undo("document", app, doc, plan=plan)
    changed, pending = [], None
    try:
        for g in targets:
            pending = {"sections": g["sections"], "header": g["header"]}
            # Even obtaining a missing header may create it: everything below is inside the mutation boundary.
            hf = doc.Sections(g["owner"]).Headers(g["type"])
            if action == "add":
                shape = hf.Shapes.AddTextEffect(0, text, font, 36, 0, 0, 0, 0, hf.Range)  # msoTextEffect1, msoFalse
                pending["name"] = str(shape.Name)
                shape.Name = WATERMARK_PREFIX + str(g["owner"] * 10 + g["type"])
                pending["name"] = str(shape.Name)
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = rgb
                shape.Fill.Transparency = transparency
                shape.Line.Visible = 0  # msoFalse
                shape.LockAspectRatio = -1  # msoTrue
                if float(shape.Width) > g["max_width"]:
                    shape.Width = g["max_width"]
                shape.Rotation = 315 if layout == "diagonal" else 0
                shape.WrapFormat.Type = 5  # wdWrapBehind
                shape.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionPage
                shape.RelativeVerticalPosition = 1  # wdRelativeVerticalPositionPage
                shape.Left = shape.Top = -999995  # wdShapeCenter
                changed.append(dict(pending))
            else:
                for i in range(int(hf.Shapes.Count), 0, -1):
                    shape = hf.Shapes(i)
                    if str(shape.Name).startswith(WATERMARK_PREFIX) and _in_header(shape, g["owner"], g["type"]):
                        pending["name"] = str(shape.Name)
                        shape.Delete()
                        changed.append(dict(pending))
    except Exception as exc:
        raise PartialChangeError(f"Watermark {action} interrupted; completed={changed}; possibly changed={pending}. "
                                 f"Use office_undo. {com.translate(exc)}") from None
    if action == "remove":
        # Never report success while the XML still shows a watermark in the chosen sections (anchor filter missed it).
        left = [m["name"] for g in _watermark_headers(doc)[0] if chosen.intersection(g["sections"]) for m in g["marks"]]
        if left:
            raise PartialChangeError(f"Watermark remove incomplete: still present {sorted(set(left))}; removed={changed}. "
                                     "Use office_undo or remove them in Word (Design > Watermark > Remove).")
    return {"ok": True, "document": name, "action": action, "changed": changed}



# ================================================================== подписи и перекрёстные ссылки

LABELS = {"table": -2, "figure": -1, "equation": -3}  # wdCaptionTable/Figure/Equation
KINDS = {"label_and_number": 3, "text": 2, "page": 7}  # wdOnlyLabelAndNumber, wdEntireCaption, wdPageNumber


def _items(collection):
    return [collection(i) for i in range(1, int(collection.Count) + 1)]


def _text(text):
    return clean_word_text(str(text)).strip()


def _inventory(app, doc):
    labels = {str(app.CaptionLabels(v).Name).casefold(): v for v in LABELS.values()}
    objects = []
    for kind, collection in (("table", doc.Tables), ("picture", doc.InlineShapes)):
        for i, obj in enumerate(_items(collection), 1):
            r = obj.Range
            if kind == "picture":
                if int(obj.Type) not in (3, 4):  # wdInlineShapePicture, wdInlineShapeLinkedPicture
                    continue
                r = r.Paragraphs(1).Range
            objects.append({"target": kind, "index": i, "start": int(r.Start), "end": int(r.End)})
    captions, counts = [], {}
    for field in _items(doc.Content.Fields):
        if int(field.Type) != 12:  # wdFieldSequence
            continue
        match = re.match(r'\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))', str(field.Code.Text), re.I)
        if not match:
            raise ToolError("Cannot parse a SEQ label; caption inventory is incomplete.")
        label = match[1] or match[2]
        key = label.casefold()
        counts[key] = counts.get(key, 0) + 1
        r = field.Result.Paragraphs(1).Range
        near = [{"target": o["target"], "index": o["index"]} for o in objects
                if o["end"] == int(r.Start) or o["start"] == int(r.End)]
        captions.append({"caption": counts[key], "label_id": labels.get(key), "label": label,
                         "number": _text(field.Result.Text), "text": _text(r.Text), "paragraph": int(r.Start),
                         "objects": near, "orphan": not near, "link": wd.range_link(doc, 0, r),
                         "_field": field})
    paragraphs = wd.paragraph_indexes(doc, [c["paragraph"] for c in captions])
    for c in captions:  # one sweep instead of counting from the document start for every caption
        c["paragraph"] = paragraphs[c["paragraph"]]
    missing = {kind: [o["index"] for o in objects if o["target"] == kind and not any(
        {"target": kind, "index": o["index"]} in c["objects"] for c in captions)] for kind in ("table", "picture")}
    return captions, missing


def _indexed(collection, index, what):
    if not 1 <= index <= int(collection.Count):
        raise ToolError(f"{what} index must be between 1 and {collection.Count}.")
    return collection(index)


def _reference(doc, captions, label_name, caption, heading, bookmark, reference):
    number_field = None
    if caption in ("heading", "bookmark"):
        typ = 1 if caption == "heading" else 2  # wdRefTypeHeading, wdRefTypeBookmark
        items = list(doc.GetCrossReferenceItems(typ) or ())
        if typ == 1:
            if not 1 <= heading <= len(items):
                raise ToolError("heading must be an index in GetCrossReferenceItems(heading).")
            item = heading
        else:
            matches = [s for s in items if str(s).casefold() == bookmark.casefold()]
            if len(matches) != 1:
                raise ToolError("Bookmark is missing or ambiguous.")
            item = matches[0]  # Word requires a bookmark name, not an index.
        kind = 7 if reference == "page" else -3 if reference == "number" else -1  # wdPageNumber/NumberNoContext/ContentText
    else:
        chosen = [c for c in captions if c["label"].casefold() == label_name.casefold() and str(c["caption"]) == str(caption)]
        if len(chosen) != 1:
            raise ToolError("caption must be the per-label index returned by list.")
        selected = chosen[0]
        typ = label_name
        items = list(doc.GetCrossReferenceItems(typ) or ())
        matches = [i for i, s in enumerate(items, 1) if _text(s) == selected["text"]]
        if len(matches) != 1 or sum(c["text"] == selected["text"] for c in captions if c["label"] == selected["label"]) != 1:
            raise ToolError("Cross-reference caption text is missing or ambiguous.")
        item, kind = matches[0], KINDS.get(reference, 3)  # wdOnlyLabelAndNumber
        if reference == "number":
            number_field = selected["_field"]
    return typ, kind, item, number_field


def word_captions(document: str, action: str = "list", target: str = "", table_index: int = 0,
                  picture_index: int = 0, paragraph: int = 0, label: str = "", text: str = "",
                  position: str = "", caption: int | str = 0, heading: int = 0, bookmark: str = "",
                  reference: str = "label_and_number", at: str = "paragraph_end", find_text: str = "",
                  occurrence: int = 1, hyperlink: bool = True) -> dict:
    """List SEQ captions, add an automatic caption or insert a cross-reference without selecting anything.

    action: list (readonly), add, cross_reference. Main story only; list includes orphan captions and uncaptioned objects.
    add: target=table/picture/paragraph with table_index/picture_index/paragraph (1-based). Pictures must be inline.
    label: table/figure/equation (built-in IDs); defaults to table for tables/references, figure for pictures;
        required for paragraph targets. text is passed unchanged as InsertCaption's Title (no invented separator).
    position: above/below; defaults to above for tables, below otherwise.
    cross_reference: caption is a per-label list index, or 'heading' with heading index, or 'bookmark' with bookmark name.
    reference: label_and_number, number (SEQ result only), text (entire caption), page.
        For headings/bookmarks, label_and_number and text both mean content text; number means paragraph number.
    at: after_text with literal find_text and occurrence, or paragraph_end with paragraph. hyperlink defaults to true.
    Each write uses one UndoRecord; other captions are renumbered by word_update_fields, not by this tool.
    """
    if action not in ("list", "add", "cross_reference"):
        raise ToolError("action must be list, add or cross_reference.")
    app, doc = wd.pick_document(document)
    captions, missing = _inventory(app, doc)
    if action == "list":
        public = [{k: v for k, v in c.items() if k != "_field"} for c in captions]
        return {"document": doc.Name, "captions": public, "orphans": [c for c in public if c["orphan"]],
                "tables_without_caption": missing["table"], "pictures_without_caption": missing["picture"]}
    label = label or ("figure" if target == "picture" else "table" if target != "paragraph" else "")
    if label not in LABELS:
        raise ToolError("label must be table, figure or equation; paragraph targets require an explicit label.")
    label_id, label_name = LABELS[label], str(app.CaptionLabels(LABELS[label]).Name)
    number_range = None
    if action == "add":
        if target not in ("table", "picture", "paragraph") or position not in ("", "above", "below"):
            raise ToolError("target must be table/picture/paragraph; position must be above/below.")
        if any(c in text for c in "\r\n\x07"):
            raise ToolError("Caption text must be a single paragraph.")
        if target == "table":
            obj = _indexed(doc.Tables, table_index, "Table")
            if int(obj.NestingLevel) != 1:
                raise ToolError("Nested tables cannot receive captions.")
            rng = obj.Range
        elif target == "picture":
            obj = _indexed(doc.InlineShapes, picture_index, "Inline picture (floating Shapes are unsupported; convert to inline)")
            if int(obj.Type) not in (3, 4):  # wdInlineShapePicture, wdInlineShapeLinkedPicture
                raise ToolError("Target is not an inline picture; floating Shapes are unsupported.")
            rng = obj.Range
            if int(rng.Paragraphs(1).Range.InlineShapes.Count) != 1:
                raise ToolError("Multiple inline objects share the paragraph; caption placement is ambiguous.")
        else:
            rng = wd.paragraphs_range(doc, paragraph)
        if target != "table" and bool(rng.Information(12)):  # wdWithInTable
            raise ToolError("Picture/paragraph captions inside tables are unsupported; choose an outside paragraph.")
        bounds = rng if target == "table" else rng.Paragraphs(1).Range
        if any(c["label_id"] == label_id and (
            int(c["_field"].Result.Paragraphs(1).Range.End) == int(bounds.Start) or
            int(c["_field"].Result.Paragraphs(1).Range.Start) == int(bounds.End)) for c in captions):
            raise ToolError("An adjacent caption already uses this label.")
        position = position or ("above" if target == "table" else "below")
        start = int(bounds.Start if position == "above" else bounds.End)
    else:
        if reference not in (*KINDS, "number"):
            raise ToolError("reference must be label_and_number, number, text or page.")
        typ, kind, item, number_field = _reference(doc, captions, label_name, caption, heading, bookmark, reference)
        if caption in ("heading", "bookmark"):
            label_id, label_name = None, None
        if at == "after_text" and find_text and occurrence > 0:
            hits = list(wd.find_all(doc.Content, wd.escape_find(find_text), False, False, False, limit=occurrence))
            if len(hits) < occurrence:
                raise ToolError("find_text occurrence was not found.")
            start = int(hits[occurrence - 1].End)
        elif at == "paragraph_end":
            start = int(wd.paragraphs_range(doc, paragraph).End) - 1
            if str(doc.Range(start, start + 1).Text) == "\x07":
                start -= 1  # insert before the cell's paragraph marker
        else:
            raise ToolError("at must be paragraph_end, or after_text with find_text and positive occurrence.")
        if any(int(f.Code.Start) - 1 < start < int(f.Result.End) + 1 for f in _items(doc.Content.Fields)):
            raise ToolError("Cannot insert a cross-reference inside an existing field.")
        rng = doc.Range(start, start)
        if number_field is not None:
            number_range = doc.Range(int(number_field.Code.Start) - 1, int(number_field.Result.End) + 1)
            item = "_OfficeLiveCaption" + uuid.uuid4().hex[:20]
            if doc.Bookmarks.Exists(item):
                raise ToolError("Generated bookmark already exists; retry.")
            typ, kind = 2, -1  # wdRefTypeBookmark, wdContentText: preserve any SEQ number format
    before_end = int(doc.Content.End)
    name = str(doc.Name)
    undo.require_undo("document", app, doc)
    stage = "insertion attempted; completion unknown"
    try:
        if action == "add":
            # TitleAutoText "" not Missing: Missing in the middle of a Word argument list nulls the rest (Position).
            rng.InsertCaption(label_id, text, "", 0 if position == "above" else 1, False)  # wdCaptionPositionAbove/Below
        else:
            if number_range is not None:
                doc.Bookmarks.Add(item, number_range)
                stage = "number bookmark created; reference insertion attempted"
            rng.InsertCrossReference(typ, kind, item, bool(hyperlink), False, False, pythoncom.Missing)
        stage = "caption inserted" if action == "add" else "cross-reference inserted"
        delta = int(doc.Content.End) - before_end
        if delta <= 0:
            raise RuntimeError("Cannot identify the inserted range.")
        touched = doc.Range(start, start + delta)
        if int(touched.Fields.Update()):
            raise RuntimeError("Inserted fields could not be updated.")
        return {"ok": True, "document": name, "action": action, "label_id": label_id, "label": label_name,
                "text": _text(touched.Text), "link": wd.range_link(doc, 0, touched),
                "note": "Other captions are not renumbered here. After inserting in the middle, use word_update_fields."}
    except Exception as exc:
        raise PartialChangeError(f"Caption operation interrupted: {stage}. Use office_undo to undo this call. {exc}") from None
