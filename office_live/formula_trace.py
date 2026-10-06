"""Bounded read-only reference graph. Never evaluates formulas or activates worksheets."""

import re
from collections import deque

import pywintypes

from .errors import ToolError
from .util import MAX_COLS, MAX_ROWS, cell_kind, norm_value, parse_a1, quote_sheet, to_grid
from .xl_common import addr_of, bounds

SCAN_LIMIT = 500000
_STRING = re.compile(r'"(?:[^"]|"")*"')
_IDENT = r"(?:[^\W\d]|\\)[\w.\\]*"
_QUALIFIER = re.compile(r"(?:'((?:[^']|'')+)'|([\w.\[\]:\\]+))!")
_CELL = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*"
_REFERENCE = re.compile(rf"(?:{_CELL}(?::{_CELL})?|\$?[A-Za-z]{{1,3}}:\$?[A-Za-z]{{1,3}}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)(?![\w.])")
_NAME = re.compile(_IDENT)
_NUMBER = re.compile(r"(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?")


class Trace:
    def __init__(self, app, wb, direction, depth, max_nodes):
        self.app, self.wb = app, wb
        self.direction, self.depth, self.max_nodes = direction, depth, max_nodes
        self.reasons, self.nodes, self.edges = set(), {}, set()
        self.sheets = {wb.Worksheets(i).Name.casefold(): wb.Worksheets(i) for i in range(1, int(wb.Worksheets.Count) + 1)}
        self.scanned = None
        self.truncated = False
        self.reference_checks = 0
        self.function_names = {}

    def limited(self, reason):
        self.reasons.add(reason)
        self.truncated = True

    def areas(self, rng):
        if rng.Worksheet.Parent.FullName.casefold() != self.wb.FullName.casefold():
            self.reasons.add("external_reference")
            return []
        return [rng.Areas(i) for i in range(1, int(rng.Areas.Count) + 1)]

    def named(self, ws, name):
        for names in (ws.Names, self.wb.Names):
            try:
                obj = names(name)
            except (pywintypes.com_error, KeyError):
                continue
            definition = _STRING.sub("", str(obj.RefersTo))
            if re.search(r"\[[^\]]+\][^!]*!", definition):
                self.reasons.add("external_reference")
                return []
            if re.search(r"[^!]+:[^!]+!", definition):
                self.reasons.add("3d_reference")
                return []
            if re.search(r"[\w.]\s*\(", definition):
                self.reasons.add("name_not_range")
                return []
            try:
                return self.areas(obj.RefersToRange)
            except (pywintypes.com_error, AttributeError):
                self.reasons.add("name_not_range")
                return []
        self.reasons.add("unresolved_name")
        return []

    def structured(self, table_name, spec):
        # Only a whole data column is exact without row/total/header context.
        if not re.fullmatch(r"\[[^\[\]#@':]+\]", spec):
            self.reasons.add("unsupported_structured_reference")
            return []
        for ws in self.sheets.values():
            for i in range(1, int(ws.ListObjects.Count) + 1):
                table = ws.ListObjects(i)
                if table.Name.casefold() == table_name.casefold():
                    try:
                        rng = table.ListColumns(spec[1:-1]).DataBodyRange
                        return [] if rng is None else self.areas(rng)
                    except (pywintypes.com_error, KeyError):
                        self.reasons.add("unresolved_table_column")
                        return []
        self.reasons.add("unresolved_table")
        return []

    def named_function(self, ws, name):
        key = (ws.Name.casefold(), name.casefold())
        if key not in self.function_names:
            self.function_names[key] = False
            for names in (ws.Names, self.wb.Names):
                try:
                    names(name)
                    self.function_names[key] = True
                    break
                except (pywintypes.com_error, KeyError):
                    continue
        if self.function_names[key]:
            # A called defined name can be a LAMBDA/LET closure with hidden dependencies.
            self.reasons.add("named_function")

    def references(self, ws, formula):
        text = _STRING.sub(lambda m: " " * len(m[0]), formula[1:])
        if re.search(r"[\w\]#]\s+(?:\$?[A-Za-z]+\$?\d|')", text):
            self.reasons.add("possible_intersection_operator")
        refs, i = [], 0
        while i < len(text):
            if text[i].isspace() or text[i] in "+-*/^&=<>%,;(){}":
                i += 1
                continue
            start, target, qualified, external = i, ws, False, False
            qualifier = _QUALIFIER.match(text, i)
            if qualifier:
                qualified = True
                sheet = (qualifier[1] or qualifier[2]).replace("''", "'")
                i = qualifier.end()
                if "[" in sheet or "]" in sheet:
                    self.reasons.add("external_reference")
                    external = True
                elif ":" in sheet:
                    self.reasons.add("3d_reference")
                    external = True
                else:
                    target = self.sheets.get(sheet.casefold())
                    if target is None:
                        self.reasons.add("unresolved_sheet")
                        external = True
            name = _NAME.match(text, i)
            if name and text[name.end():].lstrip().startswith("("):
                function = name[0].upper().removeprefix("_XLFN.").removeprefix("_XLWS.")
                if function in {"INDIRECT", "OFFSET", "INDEX", "LET", "LAMBDA"}:
                    self.reasons.add(function.lower())
                if not external:
                    self.named_function(target, name[0])
                i = name.end()
                continue
            if name and name.end() < len(text) and text[name.end()] == "[":
                j, level = name.end(), 0
                while j < len(text):
                    level += (text[j] == "[") - (text[j] == "]")
                    j += 1
                    if level == 0:
                        break
                if not external:
                    refs.extend(self.structured(name[0], text[name.end():j]))
                i = j
                continue
            ref = _REFERENCE.match(text, i)
            if ref:
                i = ref.end()
                if i < len(text) and text[i] == "#":
                    self.reasons.add("spill_reference")
                    i += 1
                if not external:
                    r, c, end, right = parse_a1(ref[0])
                    if end > MAX_ROWS or right > MAX_COLS:
                        self.reasons.add("unsupported_reference")
                    else:
                        try:
                            refs.extend(self.areas(target.Range(ref[0])))
                        except pywintypes.com_error:
                            self.reasons.add("unresolved_reference")
                continue
            if name:
                i = name.end()
                if not external and name[0].upper() not in {"TRUE", "FALSE"}:
                    refs.extend(self.named(target, name[0]))
                continue
            number = _NUMBER.match(text, i)
            if number and not qualified:
                i = number.end()
                continue
            # Unsupported tokens can hide references; never certify this formula as complete.
            self.reasons.add("unsupported_formula_syntax")
            i = max(start + 1, i + 1)
        return refs

    def direct(self, cell):
        ws = cell.Worksheet
        active_book, active_sheet = self.app.ActiveWorkbook, self.app.ActiveSheet
        if (active_book is None or active_sheet is None or active_book.FullName.casefold() != self.wb.FullName.casefold()
                or active_sheet.Name.casefold() != ws.Name.casefold()):
            self.reasons.add("inactive_sheet_com_unavailable")
            return []
        try:
            return self.areas(getattr(cell, "DirectPrecedents" if self.direction == "precedents" else "DirectDependents"))
        except pywintypes.com_error as exc:
            # 1004 also covers failures other than 'No cells found'. Treat as empty, flag uncertainty.
            code = exc.excepinfo[5] if exc.excepinfo and len(exc.excepinfo) > 5 else exc.hresult
            if code in (-2146827284, 1004):
                self.reasons.add("direct_references_unavailable_or_empty")
            else:
                self.reasons.add("direct_references_com_error")
            return []

    def scan(self):
        if self.scanned is not None:
            return self.scanned
        self.scanned, remaining = [], SCAN_LIMIT
        for ws in self.sheets.values():
            used = ws.UsedRange
            r, c, end, right = bounds(used)
            width = right - c + 1
            size = (end - r + 1) * width
            if size > remaining:
                self.limited("scan_limit")
                continue  # never request an over-budget COM array
            remaining -= size
            grid = to_grid(used.Formula)
            for i, row in enumerate(grid):
                for j, formula in enumerate(row):
                    if isinstance(formula, str) and formula.startswith("="):
                        cell = ws.Cells(r + i, c + j)
                        if bool(cell.HasFormula):
                            if bool(cell.HasSpill):
                                self.reasons.add("dynamic_array")
                            self.scanned.append((cell, self.references(ws, formula)))
        return self.scanned

    def neighbors(self, cell):
        refs = self.direct(cell)
        if self.direction == "precedents":
            if bool(cell.HasFormula):
                refs += self.references(cell.Worksheet, str(cell.Formula))
            if bool(cell.HasSpill):
                self.reasons.add("dynamic_array")
        else:
            for candidate, ranges in self.scan():
                for rng in ranges:
                    self.reference_checks += 1
                    if self.reference_checks > SCAN_LIMIT * 4:
                        self.limited("reference_check_limit")
                        break
                    r, c, end, right = bounds(rng)
                    if (rng.Worksheet.Name.casefold() == cell.Worksheet.Name.casefold()
                            and r <= int(cell.Row) <= end and c <= int(cell.Column) <= right):
                        refs.append(candidate)
                        break
                if self.reference_checks > SCAN_LIMIT * 4:
                    break
        seen = set()
        for rng in refs:
            r, c, end, right = bounds(rng)
            for row in range(r, end + 1):
                for col in range(c, right + 1):
                    key = (rng.Worksheet.Name.casefold(), row, col)
                    if key in seen:
                        continue
                    if len(seen) >= self.max_nodes:
                        self.limited("node_limit")
                        return
                    seen.add(key)
                    yield rng.Worksheet.Cells(row, col)

    @staticmethod
    def label(cell):
        return quote_sheet(cell.Worksheet.Name) + "!" + addr_of(cell)

    def node(self, cell):
        label = self.label(cell)
        if label not in self.nodes:
            formula = str(cell.Formula) if bool(cell.HasFormula) else None
            value = norm_value(cell.Value2)
            self.nodes[label] = {"address": label, "formula": formula, "value": value,
                                 "type": "formula" if formula is not None else cell_kind(value)}
        return label

    def run(self, root):
        queue = deque([(root, 0)])
        self.node(root)
        expanded = set()
        while queue:
            cell, level = queue.popleft()
            label = self.label(cell)
            if label in expanded:
                continue
            expanded.add(label)
            for other in self.neighbors(cell):
                other_label = self.label(other)
                edge = (other_label, label) if self.direction == "precedents" else (label, other_label)
                if level >= self.depth:
                    if edge not in self.edges:
                        self.limited("depth_limit")
                    continue
                if other_label not in self.nodes and len(self.nodes) >= self.max_nodes:
                    self.limited("node_limit")
                    break
                self.node(other)
                self.edges.add(edge)
                if other_label not in expanded:
                    queue.append((other, level + 1))
        return {"nodes": list(self.nodes.values()), "edges": [{"from": a, "to": b} for a, b in sorted(self.edges)],
                "complete": not self.reasons, "incomplete_reasons": sorted(self.reasons), "truncated": self.truncated}


def validate_trace(direction, depth, max_nodes):
    if direction not in {"precedents", "dependents"}:
        raise ToolError("direction must be precedents or dependents.")
    if type(depth) is not int or not 1 <= depth <= 5:
        raise ToolError("depth must be an integer from 1 to 5.")
    if type(max_nodes) is not int or not 1 <= max_nodes <= 2000:
        raise ToolError("max_nodes must be an integer from 1 to 2000.")


def same_sheet_dependency(app, wb, target, changing, max_nodes=2000):
    """Parse a bounded transitive graph without DirectPrecedents, activation or formula evaluation."""
    trace = Trace(app, wb, "precedents", max_nodes, max_nodes)
    queue, seen, found = deque([target]), set(), False
    ws = target.Worksheet
    wanted = (int(changing.Row), int(changing.Column))
    while queue:
        cell = queue.popleft()
        key = (int(cell.Row), int(cell.Column))
        if key in seen:
            continue
        if len(seen) >= max_nodes:
            trace.limited("node_limit")
            break
        seen.add(key)
        if not bool(cell.HasFormula):
            continue
        for rng in trace.references(ws, str(cell.Formula)):
            if rng.Worksheet.Name.casefold() != ws.Name.casefold():
                trace.reasons.add("cross_sheet_dependency")
                continue
            r, c, end, right = bounds(rng)
            if r <= wanted[0] <= end and c <= wanted[1] <= right:
                found = True
            for row in range(r, end + 1):
                for col in range(c, right + 1):
                    if (row, col) in seen:
                        continue
                    if len(seen) + len(queue) >= max_nodes:
                        trace.limited("node_limit")
                        break
                    queue.append(ws.Cells(row, col))
                if "node_limit" in trace.reasons:
                    break
    return found, not trace.reasons, sorted(trace.reasons)
