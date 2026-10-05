"""Resolve link destinations and derive review links from actual tool targets/results."""

from . import links, wd_common as wd, xl_common as xl
from .errors import ToolError
from .registry import office_tool
from .util import split_sheet_ref


def identity(app, obj):
    return (xl.workbook_path(obj) if app == "excel" else wd.document_path(obj)) or str(obj.Name)


def word_range(doc, params):
    if "paragraph" in params:
        return wd.paragraphs_range(doc, int(params["paragraph"]))
    if "paragraphs" in params:
        start, end = map(int, params["paragraphs"].split("-"))
        return wd.paragraphs_range(doc, start, end)
    if "table" in params:
        index = int(params["table"])
        if index > int(doc.Tables.Count):
            raise ToolError("The linked table no longer exists.")
        return doc.Tables(index).Range
    if "bookmark" in params:
        if not doc.Bookmarks.Exists(params["bookmark"]):
            raise ToolError("The linked bookmark no longer exists.")
        return doc.Bookmarks(params["bookmark"]).Range
    return None


def resolve(app, params):
    """Exact matching across existing instances, through the standard folder guards."""
    links.validate(app, params)
    if app == "excel":
        application, obj = xl.pick_workbook(params["book"], exact_only=True)
        sheet = xl.pick_sheet(obj, params["sheet"], exact_only=True) if "sheet" in params else None
        rng = sheet.Range(links.a1(params["range"])) if "range" in params else None
    else:
        application, obj = wd.pick_document(params["doc"], exact_only=True)
        sheet, rng = None, word_range(obj, params)
    # A short name must not bypass the prohibition on open UNC/cloud/device targets.
    links.file_target(identity(app, obj))
    return application, obj, sheet, rng


def item(app, obj, **place):
    name = identity(app, obj)
    if app == "excel":
        uri = links.build(app, book=name, **place)
        label = place.get("sheet", str(obj.Name)) + ("!" + place["range"] if place.get("range") else "")
    else:
        uri = links.build(app, doc=name, **place)
        suffix = next((f"{k} {v}" for k, v in place.items() if v), "")
        label = str(obj.Name) + (": " + suffix if suffix else "")
    return {"label": label, "uri": uri}


@office_tool("history", "read", title="Link to an open Office location")
def office_link(workbook: str = "", sheet: str = "", cells: str = "", document: str = "",
                paragraph: int = 0, end_paragraph: int = 0, table: int = 0, bookmark: str = "") -> dict:
    """Build a clickable officelive:// link for a plan or result. Checks an exact, already open local target and existing location without selecting or editing it. Clicking only navigates. A range requires an explicit sheet; Word accepts one paragraph interval, table or bookmark.

    Args:
        workbook, document: exactly one open file name or full local path.
        sheet, cells: Excel sheet and optional A1 range (no defined names or external references).
        paragraph, end_paragraph: Word paragraph or inclusive interval (1..1000000).
        table, bookmark: alternative Word location, mutually exclusive with paragraphs.
    """
    if bool(workbook) == bool(document):
        raise ToolError("Pass exactly one workbook or document.")
    if any(type(v) is not int or v < 0 for v in (paragraph, end_paragraph, table)):
        raise ToolError("Paragraph/table indexes must be nonnegative integers.")
    if workbook:
        if paragraph or end_paragraph or table or bookmark:
            raise ToolError("Word selectors cannot be used with a workbook.")
        app, params = "excel", {"book": workbook, **({"sheet": sheet} if sheet else {}), **({"range": cells} if cells else {})}
    else:
        if sheet or cells or (end_paragraph and not paragraph):
            raise ToolError("Invalid Word location.")
        app, params = "word", {"doc": document}
        if paragraph:
            params["paragraphs" if end_paragraph else "paragraph"] = f"{paragraph}-{end_paragraph}" if end_paragraph else str(paragraph)
        if table:
            params["table"] = str(table)
        if bookmark:
            params["bookmark"] = bookmark
    _, obj, ws, _ = resolve(app, params)
    place = {k: v for k, v in params.items() if k not in {"book", "doc"}}
    if ws is not None:
        place["sheet"] = str(ws.Name)
    return {"links": [item(app, obj, **place)], "links_truncated": False}


def result_links(recording, result):
    """A metadata failure must never discard a completed edit's undo snapshot."""
    try:
        _result_links(recording, result)
    except Exception:  # noqa: BLE001 — even reading a target's identity can fail after a completed write
        pass


def _result_links(recording, result):
    if not isinstance(result, dict) or "links" in result:
        return
    tool, args = recording.tool, recording.args
    found, seen = [], set()

    def add(app, obj, **place):
        if len(found) > links.MAX_LINKS:
            return
        try:
            if app == "excel":
                if not place.get("sheet") or not place.get("range"):
                    return
                ws = xl.pick_sheet(obj, place["sheet"], exact_only=True)
                if str(ws.Name) == "Лог":
                    return
                place = {"sheet": str(ws.Name), "range": links.a1(place["range"])}
            else:
                word_range(obj, {k: str(v) for k, v in place.items()})
            link = item(app, obj, **place)
            if link["uri"] not in seen:
                seen.add(link["uri"])
                found.append(link)
        except Exception:  # noqa: BLE001 — navigation metadata cannot invalidate a completed edit
            pass

    for key, (_, obj) in recording.targets.items():
        app = "excel" if key[0] == "workbook" else "word"
        if app == "excel":
            if tool == "excel_find":
                for hit in result.get("matches", []):
                    add(app, obj, sheet=hit["sheet"], range=hit["cell"])
            elif tool == "excel_compare_ranges":
                for side in ("a", "b"):
                    target = result.get("target_" + side, {})
                    if target.get("workbook") != identity(app, obj):
                        continue
                    for diff in result.get("differences", []):
                        add(app, obj, sheet=target["sheet"], range=diff["cell_" + side])
                    for area in result.get("only_in_" + side, []):
                        add(app, obj, sheet=target["sheet"], range=area)
            elif tool == "excel_find_issues":
                for issue in result.get("issues", []):
                    for example in issue.get("examples", []):
                        address = example["cell"]
                        # Issue locations may denote a whole column.
                        if address.isalpha():
                            address += ":" + address
                        add(app, obj, sheet=result["sheet"], range=address)
            elif tool == "excel_trace_formula":
                for node in result.get("nodes", []):
                    sheet, address = split_sheet_ref(node["address"])
                    add(app, obj, sheet=sheet, range=address)
            elif recording.kind in {"write", "destructive"} and tool not in {"office_journal", "office_undo"}:
                entry = recording.pending.get(key)
                areas = (entry.areas or [op for op in entry.ops if "sheet" in op and "address" in op]) if entry else []
                if areas:
                    for area in areas:
                        add(app, obj, sheet=area["sheet"], range=area["address"])
                elif tool == "excel_copy_range":
                    for field, owner in (("to", "dest_workbook"), ("from", "source_workbook")):
                        if (field == "to" or result.get("moved")) and result.get(owner) == identity(app, obj):
                            sheet, address = split_sheet_ref(result[field])
                            add(app, obj, sheet=sheet, range=address)
                elif tool == "excel_hide_rows_columns":
                    for address in result.get("affected", []):
                        add(app, obj, sheet=result["sheet"], range=address)
                elif tool in {"excel_insert_rows_columns", "excel_delete_rows_columns", "excel_group_rows_columns"}:
                    from .excel_core import _line_spec
                    from .util import col_letter

                    spans = _line_spec(args["lines"], args["axis"]) if args.get("lines") else [(args["start"], args["start"] + args.get("count", 1) - 1)] if "start" in args else []
                    for start, end in spans:
                        address = f"{start}:{end}" if args["axis"] == "rows" else f"{col_letter(start)}:{col_letter(end)}"
                        add(app, obj, sheet=result.get("sheet") or args.get("sheet"), range=address)
                else:
                    # Result is authoritative (expanded writes, split column, etc.).
                    address = next((result[k] for k in ("range", "written", "output_range", "cells", "sorted", "filtered", "block", "cleared_range", "range_cleared") if isinstance(result.get(k), str)), "")
                    address = address or next((args[k] for k in ("cells", "cell") if args.get(k)), "")
                    sheet, address = split_sheet_ref(address)
                    add(app, obj, sheet=sheet or result.get("sheet") or args.get("sheet"), range=address)
        elif tool in {"word_find", "word_get_structure"}:
            for hit in result.get("matches" if tool == "word_find" else "headings", []):
                if hit.get("paragraph"):
                    add(app, obj, paragraph=hit["paragraph"])
        elif recording.kind in {"write", "destructive"} and tool.startswith(("word_", "bridge_")):
            if any(k in result for k in ("deleted_paragraphs", "deleted_table", "converted_table", "deleted")):
                continue
            if result.get("bookmark"):
                add(app, obj, bookmark=result["bookmark"])
            elif type(result.get("table")) is int:
                add(app, obj, table=result["table"])
            elif type(result.get("paragraph")) is int:
                add(app, obj, paragraph=result["paragraph"])
            elif isinstance(result.get("paragraphs"), list) and len(result["paragraphs"]) == 2:
                add(app, obj, paragraphs="-".join(map(str, result["paragraphs"])))
            elif tool == "word_format_text" and args.get("start_paragraph") and args.get("scope") == "paragraphs":
                add(app, obj, paragraphs=f"{args['start_paragraph']}-{args.get('end_paragraph') or args['start_paragraph']}")
    if found:
        result.update(links=found[:links.MAX_LINKS], links_truncated=len(found) > links.MAX_LINKS)


def undo_links(app, obj, entry):
    """Only surviving/restored snapshot destinations can be offered after an undo."""
    from types import SimpleNamespace

    if entry.kind != "workbook":
        return {}
    key = ("workbook", identity("excel", obj), 0)
    result = {}
    result_links(SimpleNamespace(tool="excel_undo_result", kind="write", args={}, targets={key: (app, obj)}, pending={key: entry}), result)
    return result
