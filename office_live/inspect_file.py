"""Разбор файлов Office БЕЗ запуска приложения: xlsx/xlsm (листы, диапазоны, объединения, условное форматирование,
макросы) и docx/docm (абзацы, таблицы, картинки, колонтитулы, стили). Файл открывается как zip и только читается."""

import re
import zipfile
from xml.etree import ElementTree as ET

from .vba import read_vba_project

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _unescape(s: str) -> str:
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    out = []
    for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(f"{NS_MAIN}si"):
        out.append("".join(t.text or "" for t in si.iter(f"{NS_MAIN}t")))
    return out


def _sheet_files(z: zipfile.ZipFile) -> list[dict]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")).findall(f"{NS_REL}Relationship")}
    out = []
    for sh in wb.find(f"{NS_MAIN}sheets"):
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rels.get(rid, "")
        target = target.lstrip("/")
        path = target if target.startswith("xl/") else "xl/" + target
        out.append({"name": sh.get("name"), "state": sh.get("state", "visible"), "path": path})
    return out


def _peek(sheet_xml: str, strings: list[str], rows: int) -> list[dict]:
    out = []
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.S):
        if len(out) >= rows:
            break
        cells = []
        for c in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', m.group(2), re.S):
            body = c.group(3) or ""
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
            if v is None and f is None:
                continue
            val = v.group(1) if v else None
            if val is not None and 't="s"' in c.group(2):
                val = strings[int(val)] if int(val) < len(strings) else val
            elif val is not None and 't="inlineStr"' in c.group(2):
                val = val
            item = {"cell": c.group(1), "value": _unescape(val)[:80] if val is not None else None}
            if f:
                item["formula"] = _unescape(f.group(1))[:100]
            cells.append(item)
        if cells:
            out.append({"row": int(m.group(1)), "cells": cells[:25]})
    return out


def inspect_xlsx(path: str, peek_rows: int = 6, include_vba_source: bool = False, vba_module: str = "") -> dict:
    z = zipfile.ZipFile(path)
    names = z.namelist()
    info: dict = {"kind": "excel", "path": path, "parts": len(names), "has_vba": "xl/vbaProject.bin" in names}
    strings = _shared_strings(z)
    wb_xml = z.read("xl/workbook.xml").decode("utf-8")
    info["defined_names"] = [
        {"name": m.group(2), "refers_to": _unescape(m.group(3))[:100], "hidden": 'hidden="1"' in m.group(1)}
        for m in re.finditer(r"<definedName([^>]*?)name=\"([^\"]+)\"[^>]*>(.*?)</definedName>", wb_xml, re.S)
    ][:60]
    sheets = []
    for s in _sheet_files(z):
        if s["path"] not in names:
            continue
        xml = z.read(s["path"]).decode("utf-8", "replace")
        item = {"name": s["name"], "state": s["state"], "xml_kb": round(len(xml) / 1024)}
        dim = re.search(r'<dimension ref="([^"]+)"', xml)
        item["dimension"] = dim.group(1) if dim else None
        item["rows_with_cells"] = len(re.findall(r"<row ", xml))
        item["formula_cells"] = len(re.findall(r"<f[ >/]", xml))
        merges = re.findall(r'<mergeCell ref="([^"]+)"', xml)
        item["merged_ranges"] = {"count": len(merges), "examples": merges[:20]}
        pane = re.search(r"<pane ([^>]*)/>", xml)
        if pane:
            attrs = dict(re.findall(r'(\w+)="([^"]*)"', pane.group(1)))
            item["freeze"] = {k: attrs.get(k) for k in ("xSplit", "ySplit", "topLeftCell", "state") if k in attrs}
        zoom = re.search(r'zoomScale="(\d+)"', xml)
        if zoom:
            item["zoom"] = int(zoom.group(1))
        cols = re.findall(r"<col ([^>]*)/>", xml)
        item["columns_defined"] = len(cols)
        item["hidden_columns"] = [f"{a.get('min')}-{a.get('max')}" for a in (dict(re.findall(r'(\w+)="([^"]*)"', c)) for c in cols) if a.get("hidden") == "1"]
        cfs = []
        for m in re.finditer(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>', xml, re.S):
            for r in re.finditer(r'<cfRule ([^>]*?)(?:/>|>(.*?)</cfRule>)', m.group(2), re.S):
                a = dict(re.findall(r'(\w+)="([^"]*)"', r.group(1)))
                f = re.search(r"<formula>(.*?)</formula>", r.group(2) or "", re.S)
                cfs.append({"applies_to": m.group(1), "type": a.get("type"), "priority": a.get("priority"), "formula": _unescape(f.group(1)) if f else None})
        item["conditional_formats"] = {"count": len(cfs), "rules": cfs[:30]}
        item["data_validations"] = len(re.findall(r"<dataValidation ", xml))
        item["hyperlinks"] = len(re.findall(r"<hyperlink ", xml))
        item["has_autofilter"] = "<autoFilter" in xml
        item["has_drawing"] = "<drawing " in xml
        item["has_legacy_comments"] = "<legacyDrawing" in xml
        item["tables"] = len(re.findall(r"<tablePart ", xml))
        if peek_rows:
            item["peek"] = _peek(xml, strings, min(int(peek_rows), 20))
        sheets.append(item)
    info["sheets"] = sheets
    info["parts_of_interest"] = {
        "charts": len([n for n in names if n.startswith("xl/charts/chart")]),
        "images": len([n for n in names if n.startswith("xl/media/")]),
        "pivot_tables": len([n for n in names if n.startswith("xl/pivotTables/pivotTable")]),
        "tables": len([n for n in names if n.startswith("xl/tables/table")]),
        "slicers": len([n for n in names if n.startswith("xl/slicers/slicer")]),
        "external_links": len([n for n in names if n.startswith("xl/externalLinks/externalLink")]),
    }
    if "xl/styles.xml" in names:
        st = z.read("xl/styles.xml").decode("utf-8", "replace")
        info["styles"] = {
            "cell_formats": len(re.findall(r"<xf ", st.split("<cellXfs")[1].split("</cellXfs>")[0])) if "<cellXfs" in st else 0,
            "fonts": len(re.findall(r"<font>|<font ", st)), "fills": len(re.findall(r"<fill>", st)), "dxfs": len(re.findall(r"<dxf>", st)),
        }
    if info["has_vba"]:
        info["vba"] = read_vba_project(z.read("xl/vbaProject.bin"), include_vba_source, vba_module)
    return info


def inspect_docx(path: str, peek_paragraphs: int = 8, include_vba_source: bool = False, vba_module: str = "") -> dict:
    z = zipfile.ZipFile(path)
    names = z.namelist()
    info: dict = {"kind": "word", "path": path, "parts": len(names), "has_vba": "word/vbaProject.bin" in names}
    doc = z.read("word/document.xml").decode("utf-8", "replace")
    paragraphs = re.findall(r"<w:p[ >].*?</w:p>", doc, re.S)
    info["paragraphs"] = len(paragraphs)
    info["tables"] = len(re.findall(r"<w:tbl>", doc))
    info["images"] = len([n for n in names if n.startswith("word/media/")])
    info["sections"] = len(re.findall(r"<w:sectPr", doc))
    info["fields"] = len(re.findall(r"<w:fldSimple|<w:instrText", doc))
    info["tracked_changes"] = {"insertions": len(re.findall(r"<w:ins ", doc)), "deletions": len(re.findall(r"<w:del ", doc))}
    info["headers"] = len([n for n in names if n.startswith("word/header")])
    info["footers"] = len([n for n in names if n.startswith("word/footer")])
    info["comments"] = len(re.findall(r"<w:comment ", z.read("word/comments.xml").decode("utf-8", "replace"))) if "word/comments.xml" in names else 0
    styles: dict[str, int] = {}
    for s in re.findall(r'<w:pStyle w:val="([^"]+)"', doc):
        styles[s] = styles.get(s, 0) + 1
    info["paragraph_styles_used"] = dict(sorted(styles.items(), key=lambda kv: -kv[1])[:20])
    peek = []
    for p in paragraphs:
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, re.S))
        if text.strip():
            sty = re.search(r'<w:pStyle w:val="([^"]+)"', p)
            peek.append({"style": sty.group(1) if sty else None, "text": _unescape(text)[:120]})
            if len(peek) >= peek_paragraphs:
                break
    info["peek"] = peek
    placeholders = sorted({m for m in re.findall(r"\{\{\s*([^{}<]{1,60}?)\s*\}\}", re.sub(r"<[^>]+>", "", doc))})
    if placeholders:
        info["placeholders"] = placeholders
    if info["has_vba"]:
        info["vba"] = read_vba_project(z.read("word/vbaProject.bin"), include_vba_source, vba_module)
    return info
