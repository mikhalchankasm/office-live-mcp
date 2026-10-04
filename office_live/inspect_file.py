"""Разбор файлов Office БЕЗ запуска приложения: xlsx/xlsm (листы, диапазоны, объединения, условное форматирование,
макросы) и docx/docm (абзацы, таблицы, картинки, колонтитулы, стили). Файл открывается как zip и только читается.

XML листов разбирается потоково (ElementTree.iterparse), а не регулярными выражениями: порядок атрибутов, inlineStr,
пространства имён и экранирование не ломают разбор. Размер каждой читаемой части ограничен — архив-«бомба» не съест память."""

import re
import zipfile
from xml.etree import ElementTree as ET

from .errors import ToolError
from .vba import read_vba_project

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
NS_DOC_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

MAX_PART_BYTES = 256 * 1024 * 1024  # потолок распакованного размера одной читаемой части
MAX_SHARED_STRINGS = 300_000  # сколько общих строк держим в памяти (остальные — по индексу не покажем)
MAX_CELL_CHARS = 80


def _human(n: int) -> str:
    return f"{n / (1024 * 1024):.0f} MB" if n >= 1024 * 1024 else f"{n / 1024:.1f} KB" if n >= 1024 else f"{n} bytes"


def _check_size(z: zipfile.ZipFile, name: str) -> None:
    size = z.getinfo(name).file_size
    if size > MAX_PART_BYTES:
        raise ToolError(f"Part '{name}' is {_human(size)} unpacked, above the safety limit of {_human(MAX_PART_BYTES)}; open the file in Excel/Word instead.")


def _read_part(z: zipfile.ZipFile, name: str) -> bytes:
    _check_size(z, name)
    return z.read(name)


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    _check_size(z, "xl/sharedStrings.xml")
    out: list[str] = []
    with z.open("xl/sharedStrings.xml") as f:
        for _, el in ET.iterparse(f, events=("end",)):
            if el.tag == f"{NS_MAIN}si":
                if len(out) < MAX_SHARED_STRINGS:
                    out.append("".join(t.text or "" for t in el.iter(f"{NS_MAIN}t"))[:MAX_CELL_CHARS * 2])
                el.clear()
    return out


def _sheet_files(z: zipfile.ZipFile) -> list[dict]:
    wb = ET.fromstring(_read_part(z, "xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(_read_part(z, "xl/_rels/workbook.xml.rels")).findall(f"{NS_REL}Relationship")}
    out = []
    sheets = wb.find(f"{NS_MAIN}sheets")
    for sh in sheets if sheets is not None else []:
        target = (rels.get(sh.get(f"{NS_DOC_REL}id"), "") or "").lstrip("/")
        path = target if target.startswith("xl/") else "xl/" + target
        out.append({"name": sh.get("name"), "state": sh.get("state", "visible"), "path": path})
    return out


def _defined_names(z: zipfile.ZipFile) -> list[dict]:
    wb = ET.fromstring(_read_part(z, "xl/workbook.xml"))
    names = wb.find(f"{NS_MAIN}definedNames")
    out = []
    for dn in names if names is not None else []:
        out.append({"name": dn.get("name"), "refers_to": (dn.text or "")[:100], "hidden": dn.get("hidden") in ("1", "true")})
        if len(out) >= 60:
            break
    return out


def _cell_value(cell: ET.Element, strings: list[str]):
    """(значение, формула) ячейки <c>: общие и встроенные строки, формульные результаты, числа."""
    kind = cell.get("t", "n")
    v = cell.find(f"{NS_MAIN}v")
    f = cell.find(f"{NS_MAIN}f")
    value = None
    if kind == "inlineStr":
        is_el = cell.find(f"{NS_MAIN}is")
        value = "".join(t.text or "" for t in is_el.iter(f"{NS_MAIN}t")) if is_el is not None else None
    elif v is not None and v.text is not None:
        value = v.text
        if kind == "s":
            try:
                idx = int(value)
                value = strings[idx] if idx < len(strings) else f"<shared string #{idx}>"
            except ValueError:
                pass
        elif kind == "b":
            value = "TRUE" if value == "1" else "FALSE"
    formula = (f.text or "") if f is not None and (f.text or f.get("t") == "shared") else None
    return value, formula, f is not None


def _scan_sheet(z: zipfile.ZipFile, path: str, strings: list[str], peek_rows: int) -> dict:
    """Один потоковый проход по листу: размеры, объединения, закрепление, столбцы, условное форматирование, предпросмотр."""
    _check_size(z, path)
    item: dict = {"xml_kb": round(z.getinfo(path).file_size / 1024)}
    merges: list[str] = []
    merge_count = 0
    cf_rules: list[dict] = []
    cf_count = 0
    hidden_cols: list[str] = []
    cols_defined = rows = formulas = validations = hyperlinks = tables = 0
    flags = {"autofilter": False, "drawing": False, "legacy": False}
    peek: list[dict] = []
    current_cf: str | None = None
    with z.open(path) as f:
        for event, el in ET.iterparse(f, events=("start", "end")):
            name = _local(el.tag)
            if event == "start":
                if name == "dimension":
                    item["dimension"] = el.get("ref")
                elif name == "pane":
                    item["freeze"] = {k: el.get(k) for k in ("xSplit", "ySplit", "topLeftCell", "state") if el.get(k) is not None}
                elif name == "sheetView" and el.get("zoomScale") and "zoom" not in item:
                    item["zoom"] = int(el.get("zoomScale"))
                elif name == "conditionalFormatting":
                    current_cf = el.get("sqref")
                elif name == "autoFilter":
                    flags["autofilter"] = True
                elif name == "drawing":
                    flags["drawing"] = True
                elif name == "legacyDrawing":
                    flags["legacy"] = True
                continue
            # event == "end"
            if name == "c":
                continue  # обрабатывается при закрытии строки, пока дочерние элементы целы
            if name == "f":
                formulas += 1
            elif name == "row":
                rows += 1
                if len(peek) < peek_rows:
                    cells = []
                    for c in el.findall(f"{NS_MAIN}c"):
                        value, formula, has_formula = _cell_value(c, strings)
                        if value is None and not has_formula:
                            continue
                        entry = {"cell": c.get("r"), "value": value[:MAX_CELL_CHARS] if isinstance(value, str) else value}
                        if has_formula:
                            entry["formula"] = (formula or "")[:100]
                        cells.append(entry)
                        if len(cells) >= 25:
                            break
                    if cells:
                        peek.append({"row": int(el.get("r") or rows), "cells": cells})
                el.clear()
            elif name == "mergeCell":
                merge_count += 1
                if len(merges) < 20:
                    merges.append(el.get("ref"))
            elif name == "col":
                cols_defined += 1
                if el.get("hidden") in ("1", "true"):
                    hidden_cols.append(f"{el.get('min')}-{el.get('max')}")
            elif name == "cfRule":
                cf_count += 1
                if len(cf_rules) < 30:
                    formula_el = el.find(f"{NS_MAIN}formula")
                    cf_rules.append({
                        "applies_to": current_cf, "type": el.get("type"), "priority": el.get("priority"),
                        "formula": formula_el.text if formula_el is not None else None,
                    })
            elif name == "conditionalFormatting":
                current_cf = None
            elif name == "dataValidation":
                validations += 1
            elif name == "hyperlink":
                hyperlinks += 1
            elif name == "tablePart":
                tables += 1
    item.setdefault("dimension", None)
    item.update({
        "rows_with_cells": rows, "formula_cells": formulas, "merged_ranges": {"count": merge_count, "examples": merges},
        "columns_defined": cols_defined, "hidden_columns": hidden_cols, "conditional_formats": {"count": cf_count, "rules": cf_rules},
        "data_validations": validations, "hyperlinks": hyperlinks, "has_autofilter": flags["autofilter"], "has_drawing": flags["drawing"],
        "has_legacy_comments": flags["legacy"], "tables": tables,
    })
    if peek_rows:
        item["peek"] = peek
    return item


def inspect_xlsx(path: str, peek_rows: int = 6, include_vba_source: bool = False, vba_module: str = "") -> dict:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        info: dict = {"kind": "excel", "path": path, "parts": len(names), "has_vba": "xl/vbaProject.bin" in names}
        strings = _shared_strings(z)
        info["defined_names"] = _defined_names(z)
        sheets = []
        for s in _sheet_files(z):
            if s["path"] not in names:
                continue
            item = {"name": s["name"], "state": s["state"]}
            try:
                item.update(_scan_sheet(z, s["path"], strings, min(int(peek_rows), 20)))
            except ToolError as exc:
                item["skipped"] = str(exc)
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
            st = ET.fromstring(_read_part(z, "xl/styles.xml"))

            def count(parent: str, child: str) -> int:
                node = st.find(f"{NS_MAIN}{parent}")
                return len(node.findall(f"{NS_MAIN}{child}")) if node is not None else 0

            info["styles"] = {"cell_formats": count("cellXfs", "xf"), "fonts": count("fonts", "font"), "fills": count("fills", "fill"), "dxfs": count("dxfs", "dxf")}
        if info["has_vba"]:
            info["vba"] = read_vba_project(_read_part(z, "xl/vbaProject.bin"), include_vba_source, vba_module)
    return info


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _wval(el: ET.Element) -> str | None:
    """Значение атрибута w:val при любом префиксе пространства имён."""
    return el.get(f"{{{W_NS}}}val")


def _scan_document(z: zipfile.ZipFile, peek_paragraphs: int) -> dict:
    """Один потоковый проход по word/document.xml: считаем по локальным именам {пространство имён}p/tbl/…, а не по префиксу `w:`."""
    _check_size(z, "word/document.xml")
    counts = {"p": 0, "tbl": 0, "sectPr": 0, "field": 0, "ins": 0, "del": 0}
    styles: dict[str, int] = {}
    peek: list[dict] = []
    texts: list[str] = []
    with z.open("word/document.xml") as f:
        for _, el in ET.iterparse(f, events=("end",)):
            if el.tag.startswith(f"{{{W_NS}}}"):
                name = el.tag.rsplit("}", 1)[-1]
                if name == "p":
                    counts["p"] += 1
                    text = "".join(t.text or "" for t in el.iter(f"{{{W_NS}}}t"))
                    if text:
                        texts.append(text)
                    ps = el.find(f"{{{W_NS}}}pPr/{{{W_NS}}}pStyle")
                    style = _wval(ps) if ps is not None else None
                    if style:
                        styles[style] = styles.get(style, 0) + 1
                    if text.strip() and len(peek) < peek_paragraphs:
                        peek.append({"style": style, "text": text[:120]})
                    el.clear()  # абзац обработан — не держим дерево большого документа в памяти
                elif name in ("tbl", "sectPr", "ins", "del"):
                    counts[name] += 1
                elif name in ("fldSimple", "instrText"):
                    counts["field"] += 1
    return {"counts": counts, "styles": styles, "peek": peek, "text": "\n".join(texts)}


def inspect_docx(path: str, peek_paragraphs: int = 8, include_vba_source: bool = False, vba_module: str = "") -> dict:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        info: dict = {"kind": "word", "path": path, "parts": len(names), "has_vba": "word/vbaProject.bin" in names}
        scan = _scan_document(z, peek_paragraphs)
        counts = scan["counts"]
        info["paragraphs"] = counts["p"]
        info["tables"] = counts["tbl"]
        info["images"] = len([n for n in names if n.startswith("word/media/")])
        info["sections"] = counts["sectPr"]
        info["fields"] = counts["field"]
        info["tracked_changes"] = {"insertions": counts["ins"], "deletions": counts["del"]}
        info["headers"] = len([n for n in names if n.startswith("word/header")])
        info["footers"] = len([n for n in names if n.startswith("word/footer")])
        comments = 0
        if "word/comments.xml" in names:
            _check_size(z, "word/comments.xml")
            with z.open("word/comments.xml") as f:
                comments = sum(1 for _, el in ET.iterparse(f, events=("end",)) if el.tag == f"{{{W_NS}}}comment")
        info["comments"] = comments
        info["paragraph_styles_used"] = dict(sorted(scan["styles"].items(), key=lambda kv: -kv[1])[:20])
        info["peek"] = scan["peek"]
        placeholders = sorted(set(re.findall(r"\{\{\s*([^{}<]{1,60}?)\s*\}\}", scan["text"])))
        if placeholders:
            info["placeholders"] = placeholders
        if info["has_vba"]:
            info["vba"] = read_vba_project(_read_part(z, "word/vbaProject.bin"), include_vba_source, vba_module)
    return info
