"""Минимальный Office в памяти: никакого Dispatch, ROT или живых приложений."""

import copy
from types import SimpleNamespace as NS

import pytest
import pywintypes

from office_live.util import MAX_COLS, MAX_ROWS, a1_range, parse_a1, to_grid


class Collection:
    def __init__(self, items=()):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, key):
        if isinstance(key, int):
            return self.items[key - 1]
        for item in self.items:
            if item.Name.lower() == key.lower():
                return item
        raise KeyError(key)

    def Item(self, key, /):
        return self(key)

    def __iter__(self):
        return iter(self.items)


class Names(Collection):
    def Add(self, name, formula, visible=True, /):
        try:
            self(name).Delete()
        except KeyError:
            pass
        item = NS(Name=name, RefersTo=formula, Visible=visible)
        item.Delete = lambda: self.items.remove(item)
        self.items.append(item)
        return item


class RangeConditions(Collection):
    """Range.FormatConditions: правила, пересекающие диапазон; Add добавляет правило листа с наименьшим приоритетом."""

    def __init__(self, rng, items):
        super().__init__(items)
        self.rng = rng

    def Add(self, kind, operator=None, formula1=None, formula2=None, /):
        rules = self.rng.Worksheet.conditions.items
        rule = NS(Type=kind, AppliesTo=self.rng, Formula1=formula1, Formula2=formula2, Operator=operator,
                  Interior=NS(Color=16777215, Pattern=1, PatternColor=0), Font=NS(Bold=False, Color=0))
        rule.Delete = lambda: rules.remove(rule)
        rules.append(rule)
        return rule

    def Delete(self):
        for rule in self.items:
            self.rng.Worksheet.conditions.items.remove(rule)


class Lines:
    def __init__(self, ws, axis, count):
        self.ws, self.axis, self.Count = ws, axis, count

    def __call__(self, i):
        values = self.ws.rows if self.axis == "rows" else self.ws.columns
        return values.setdefault(i, NS(ColumnWidth=self.ws.StandardWidth, RowHeight=15, Hidden=False))


class RangeFormat:
    """Свойства диапазона агрегируются по ячейкам; неизвестных членов у COM-интерфейса нет."""

    __slots__ = ("rng", "prefix")
    defaults = {"NumberFormat": "General", "Font.Name": "Calibri", "Font.Size": 11, "Font.Bold": False,
                "Font.Italic": False, "Font.Underline": -4142, "Font.Color": 0, "Interior.Color": 16777215,
                "Interior.Pattern": -4142, "HorizontalAlignment": -4131, "VerticalAlignment": -4107,
                "WrapText": False, "MergeCells": False}
    defaults.update({"Font.Strikethrough": False, "Font.Subscript": False, "Font.Superscript": False,
                     "Font.OutlineFont": False, "Font.Shadow": False, "Font.TintAndShade": 0,
                     "Interior.PatternColor": 0, "Interior.TintAndShade": 0, "Interior.PatternTintAndShade": 0,
                     "ShrinkToFit": False, "IndentLevel": 0, "Orientation": 0, "ReadingOrder": -5002,
                     "AddIndent": False, "Locked": True, "FormulaHidden": False})
    defaults.update({f"Borders{i}.{prop}": value for i in range(5, 13) for prop, value in (("LineStyle", -4142), ("Weight", 2), ("Color", 0))})
    defaults.update({"Font.ThemeColor": 1, "Interior.ThemeColor": 1, "Interior.PatternThemeColor": 1, "Style": "Normal"})
    defaults.update({f"Borders{i}.{prop}": value for i in range(5, 13) for prop, value in (("ThemeColor", 1), ("TintAndShade", 0))})

    def __init__(self, rng, prefix):
        object.__setattr__(self, "rng", rng)
        object.__setattr__(self, "prefix", prefix)

    def __getattr__(self, name):
        prop = self.prefix + name
        if prop not in self.defaults:
            raise AttributeError(prop)
        ws = self.rng.Worksheet
        ws.format_reads.append((self.rng.Address, prop))
        values = {ws.formats.get(rc, {}).get(prop, self.defaults[prop]) for rc in self.rng.positions()}
        return values.pop() if len(values) == 1 else None

    def __setattr__(self, name, value):
        prop = self.prefix + name
        if prop not in self.defaults:
            raise AttributeError(prop)
        for rc in self.rng.positions():
            self.rng.Worksheet.formats.setdefault(rc, {})[prop] = value


def format_property(name):
    return property(lambda rng: getattr(RangeFormat(rng, ""), name),
                    lambda rng, value: setattr(RangeFormat(rng, ""), name, value))


class Range:
    NumberFormat = format_property("NumberFormat")
    HorizontalAlignment = format_property("HorizontalAlignment")
    VerticalAlignment = format_property("VerticalAlignment")
    WrapText = format_property("WrapText")
    MergeCells = format_property("MergeCells")
    ShrinkToFit = format_property("ShrinkToFit")
    IndentLevel = format_property("IndentLevel")
    Orientation = format_property("Orientation")
    ReadingOrder = format_property("ReadingOrder")
    AddIndent = format_property("AddIndent")
    Locked = format_property("Locked")
    FormulaHidden = format_property("FormulaHidden")
    Style = format_property("Style")

    def __init__(self, ws, address):
        self.Worksheet = ws
        self.Row, self.Column, self.r2, self.c2 = parse_a1(address)
        self.Rows = NS(Count=self.r2 - self.Row + 1)
        self.Columns = NS(Count=self.c2 - self.Column + 1)
        self.Address = a1_range(self.Row, self.Column, self.r2, self.c2)
        self.Areas = Collection([self])

    @property
    def HasFormula(self):
        return any(isinstance(self.Worksheet.data.get(rc), str) and self.Worksheet.data[rc].startswith("=")
                   and self.Worksheet.prefixes.get(rc) != "'" for rc in self.positions())

    @property
    def HasSpill(self):
        return any(rc in self.Worksheet.spills for rc in self.positions()) if hasattr(self.Worksheet, "spills") else False

    def _direct(self, direction):
        app = self.Worksheet.Parent.Application
        assert app.ActiveSheet is self.Worksheet, "Direct references need an already active sheet"
        ranges = self.Worksheet.direct.get((direction, self.Address), [])
        return NS(Worksheet=self.Worksheet, Areas=Collection(self.Worksheet.Range(a) for a in ranges))

    DirectPrecedents = property(lambda self: self._direct("precedents"))
    DirectDependents = property(lambda self: self._direct("dependents"))

    @property
    def PrefixCharacter(self):
        values = {self.Worksheet.prefixes.get(rc, "") for rc in self.positions()}
        return values.pop() if len(values) == 1 else None

    @property
    def MergeArea(self):
        for area in self.Worksheet.merges:
            rng = self.Worksheet.Range(area)
            if rng.Row <= self.Row <= rng.r2 and rng.Column <= self.Column <= rng.c2:
                return rng
        return self

    def positions(self):
        return ((r, c) for r in range(self.Row, self.r2 + 1) for c in range(self.Column, self.c2 + 1))

    @property
    def Formula(self):
        values = tuple(tuple(self.Worksheet.data.get((r, c)) for c in range(self.Column, self.c2 + 1)) for r in range(self.Row, self.r2 + 1))
        return values[0][0] if self.Rows.Count == self.Columns.Count == 1 else values

    @property
    def FormulaLocal(self):  # русская запись: разделитель ';' (кавычки в фейковых формулах не встречаются)
        value = self.Formula
        return value.replace(",", ";") if isinstance(value, str) else value

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.Formula = value.replace(";", ",")

    @Formula.setter
    def Formula(self, value):
        self.Worksheet.formula_writes.append((self.Address, value))
        self._assign(value, formula=True)

    @property
    def Value(self):
        values = tuple(tuple(self.Worksheet.results.get((r, c), self.Worksheet.data.get((r, c))) for c in range(self.Column, self.c2 + 1)) for r in range(self.Row, self.r2 + 1))
        return values[0][0] if self.Rows.Count == self.Columns.Count == 1 else values

    @Value.setter
    def Value(self, value):
        self._assign(value)

    Formula2 = Formula
    Value2 = Value

    def _assign(self, value, formula=False):
        grid = to_grid(value)
        for r, c in self.positions():
            item = grid[min(r - self.Row, len(grid) - 1)][min(c - self.Column, len(grid[0]) - 1)]
            if formula and isinstance(item, str) and item.isdigit():
                item = int(item)  # ловушка Excel: '007' через Formula становится числом
            if isinstance(item, str) and item.startswith("'"):
                item = item[1:]
                self.Worksheet.prefixes[r, c] = "'"
            else:
                self.Worksheet.prefixes.pop((r, c), None)
            self.Worksheet.data[r, c] = item

    @property
    def Font(self):
        return RangeFormat(self, "Font.")

    @property
    def Interior(self):
        return RangeFormat(self, "Interior.")

    @property
    def FormatConditions(self):
        rules = []
        for rule in self.Worksheet.conditions:
            rng = rule.AppliesTo
            if self.Row <= rng.r2 and rng.Row <= self.r2 and self.Column <= rng.c2 and rng.Column <= self.c2:
                rules.append(rule)
        return RangeConditions(self, rules)

    @property
    def Hyperlinks(self):
        return Collection(h for h in self.Worksheet.Hyperlinks if h.Range.Address == self.Address)

    def SpecialCells(self, cell_type, /):
        assert cell_type in (-4144, -4174)
        raise pywintypes.com_error(-2147352567, "No cells found", None, None)

    def Borders(self, index, /):
        assert 5 <= index <= 12
        return RangeFormat(self, f"Borders{index}.")

    @property
    def ColumnWidth(self):
        return self.Worksheet.Columns(self.Column).ColumnWidth

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        for i in range(self.Column, self.c2 + 1):
            self.Worksheet.Columns(i).ColumnWidth = value

    @property
    def RowHeight(self):
        return self.Worksheet.Rows(self.Row).RowHeight

    @RowHeight.setter
    def RowHeight(self, value):
        for i in range(self.Row, self.r2 + 1):
            self.Worksheet.Rows(i).RowHeight = value

    def Copy(self, destination=None, /):
        assert not self.Worksheet.Parent.Application.EnableEvents
        if destination is None:
            self.Worksheet.Parent.Application.clipboard = self
            return
        self._copy_to(destination, formats=True)

    def _copy_to(self, destination, *, formats):
        raw = self.Formula
        grid = raw if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple) else ((raw,),)
        for i, row in enumerate(grid):
            for j, value in enumerate(row):
                if (self.Worksheet.Parent is not destination.Worksheet.Parent and isinstance(value, str) and value.startswith("=")
                        and self.Worksheet.prefixes.get((self.Row + i, self.Column + j)) != "'"):
                    value = "=[other.xlsx]" + value[1:]
                destination.Worksheet.data[destination.Row + i, destination.Column + j] = value
                rc = (destination.Row + i, destination.Column + j)
                prefix = self.Worksheet.prefixes.get((self.Row + i, self.Column + j))
                if prefix:
                    destination.Worksheet.prefixes[rc] = prefix
                else:
                    destination.Worksheet.prefixes.pop(rc, None)
                if formats:
                    source_format = self.Worksheet.formats.get((self.Row + i, self.Column + j))
                    if source_format:
                        destination.Worksheet.formats[rc] = copy.deepcopy(source_format)
                    else:
                        destination.Worksheet.formats.pop(rc, None)

    def PasteSpecial(self, paste, operation, skip_blanks, transpose, /):
        assert (paste, operation, skip_blanks, transpose) == (-4123, -4142, False, False)
        self.Worksheet.Parent.Application.clipboard._copy_to(self, formats=False)

    def Cut(self, destination):
        self.Copy(destination)
        self.ClearContents()

    def ClearContents(self):
        for rc in self.positions():
            self.Worksheet.data.pop(rc, None)

    def UnMerge(self):
        pass

    def Insert(self):
        self._shift(1)

    def Delete(self):
        self._shift(-1)

    def _shift(self, direction):
        axis = 0 if self.Columns.Count == self.Worksheet.Columns.Count else 1
        start = self.Row if axis == 0 else self.Column
        count = self.Rows.Count if axis == 0 else self.Columns.Count
        out = {}
        for rc, value in self.Worksheet.data.items():
            p = rc[axis]
            if direction < 0 and start <= p < start + count:
                continue
            p += direction * count if p >= start else 0
            out[(p, rc[1]) if axis == 0 else (rc[0], p)] = value
        self.Worksheet.data = out

    def End(self, direction):
        assert direction == -4162
        rows = [r for (r, c), v in self.Worksheet.data.items() if c == self.Column and v is not None]
        return NS(Row=max(rows, default=1))

    def Select(self):
        self.Worksheet.Parent.Application.Selection = self


class Cells:
    def __init__(self, ws):
        self.ws = ws
        self.Row = self.Column = 1
        self.Rows, self.Columns = NS(Count=MAX_ROWS), NS(Count=MAX_COLS)
        self.Worksheet = ws
        self.Address = "A1:XFD1048576"

    def __call__(self, row, col):
        return self.ws.Range(a1_range(row, col, row, col))

    @property
    def FormatConditions(self):
        return self.ws.conditions

    def SpecialCells(self, cell_type, /):
        assert cell_type == -4174  # xlCellTypeAllValidation
        if not self.ws.validations:
            raise pywintypes.com_error(-2147352567, "No cells found", None, None)
        return Collection(self.ws.Cells(r, c) for r, c in self.ws.validations)

    def _dimension(self, lines, name, default, count):
        self.ws.dimension_reads.append(name)
        values = {getattr(line, name) for line in lines.values()}
        if len(lines) < count:
            values.add(default)
        return values.pop() if len(values) == 1 else None

    @property
    def ColumnWidth(self):
        return self._dimension(self.ws.columns, "ColumnWidth", self.ws.StandardWidth, MAX_COLS)

    @property
    def RowHeight(self):
        return self._dimension(self.ws.rows, "RowHeight", 15, MAX_ROWS)


class Sheets(Collection):
    def __init__(self, wb):
        super().__init__()
        self.wb = wb

    def Add(self, before=None, after=None):
        name = f"Sheet{len(self.items) + 1}"
        while any(s.Name == name for s in self.items):
            name += "x"
        sheet = Sheet(self.wb, name)
        index = self.items.index(before) if before is not None else self.items.index(after) + 1 if after is not None else 0
        self.items.insert(index, sheet)
        self.wb.ActiveSheet = sheet
        return sheet


class Hyperlinks(Collection):
    def Add(self, anchor, address, subaddress, screentip, text, /):
        assert address == "" and subaddress.startswith("'")
        anchor.Value = text
        self.items.append(NS(Range=anchor, Address=address, SubAddress=subaddress, ScreenTip=screentip, TextToDisplay=text))


class Sheet:
    def __init__(self, wb, name):
        self.Parent, self.Name = wb, name
        self.Visible = -1
        self.Tab = NS(ColorIndex=-4142, Color=0)
        self.data, self.columns, self.rows, self.formats = {}, {}, {}, {}
        self.prefixes, self.results, self.merges = {}, {}, []
        self.ProtectContents = False
        self.ProtectDrawingObjects = self.ProtectScenarios = self.ProtectionMode = False
        self.EnableSelection = 0
        self.secret = ""
        self.protection_calls = []
        self.direct = {}
        self.Protection = NS(**dict.fromkeys(("AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns",
                                             "AllowInsertingRows", "AllowInsertingHyperlinks", "AllowDeletingColumns", "AllowDeletingRows",
                                             "AllowSorting", "AllowFiltering", "AllowUsingPivotTables"), False))
        self.StandardWidth = 8.43
        self.dimension_reads = []
        self.format_reads = []
        self.formula_writes = []
        self.Rows, self.Columns = Lines(self, "rows", MAX_ROWS), Lines(self, "columns", MAX_COLS)
        self.Cells = Cells(self)
        self.Shapes, self.ListObjects = Collection(), Collection()
        self.pivots, self.Comments, self.CommentsThreaded = Collection(), Collection(), Collection()
        self.validations = set()
        self.Names = Names()
        self.Hyperlinks = Hyperlinks()
        self.conditions = Collection()
        self.view = NS(SplitRow=0, SplitColumn=0, FreezePanes=False, Zoom=100, DisplayGridlines=True, DisplayHeadings=True, ScrollRow=1, ScrollColumn=1)

    @property
    def Index(self):
        return self.Parent.Sheets.items.index(self) + 1

    @property
    def UsedRange(self):
        points = [(r, c) for (r, c), v in self.data.items() if v is not None]
        if not points:
            return self.Range("A1")
        return self.Range(a1_range(min(r for r, _ in points), min(c for _, c in points), max(r for r, _ in points), max(c for _, c in points)))

    def Range(self, address):
        return Range(self, address)

    def Protect(self, password, objects, contents, scenarios, ui, format_cells, format_columns, format_rows,
                insert_columns, insert_rows, insert_hyperlinks, delete_columns, delete_rows, sort, filtering, pivot, /):
        self.protection_calls.append("protect")
        self.secret = password
        self.ProtectDrawingObjects, self.ProtectContents, self.ProtectScenarios, self.ProtectionMode = objects, contents, scenarios, ui
        for name, value in zip(vars(self.Protection), (format_cells, format_columns, format_rows, insert_columns, insert_rows,
                                                     insert_hyperlinks, delete_columns, delete_rows, sort, filtering, pivot), strict=True):
            setattr(self.Protection, name, value)

    def Unprotect(self, password, /):
        self.protection_calls.append("unprotect")
        if password != self.secret:
            raise pywintypes.com_error(-2147352567, "Wrong password: " + password, None, None)
        self.ProtectContents = self.ProtectDrawingObjects = self.ProtectScenarios = self.ProtectionMode = False
        for name in vars(self.Protection):
            setattr(self.Protection, name, False)

    def PivotTables(self, /):
        return self.pivots

    def Activate(self):
        self.Parent.Activate()
        self.Parent.ActiveSheet = self

    def Delete(self):
        self.Parent.Sheets.items.remove(self)
        self.Parent.ActiveSheet = self.Parent.Sheets(1)

    def Copy(self, before=None, after=None):
        target = before.Parent if before is not None else after.Parent
        if not target.Windows(1).Visible:  # как настоящий Excel: «Метод Copy из класса Worksheet завершён неверно»
            raise pywintypes.com_error(-2147352567, "Exception occurred.", (0, "Microsoft Excel", "Copy method of Worksheet class failed", None, 0, -2146827284), None)
        sheet = target.Sheets.Add(before, after)
        source_book = self.Parent
        if source_book is not target:  # как настоящий Excel: имена уровня книги переезжают вместе с листом
            for item in list(source_book.Names):
                if not any(n.Name == item.Name for n in target.Names):
                    target.Names.Add(item.Name, item.RefersTo, item.Visible)
        sheet.data, sheet.columns, sheet.rows, sheet.formats = copy.deepcopy((self.data, self.columns, self.rows, self.formats))
        sheet.StandardWidth = self.StandardWidth
        sheet.Shapes, sheet.ListObjects, sheet.conditions = copy.deepcopy((self.Shapes, self.ListObjects, self.conditions), {id(self): sheet})
        sheet.pivots, sheet.Comments, sheet.CommentsThreaded, sheet.validations = copy.deepcopy(
            (self.pivots, self.Comments, self.CommentsThreaded, self.validations), {id(self): sheet})
        name = self.Name
        if any(s.Name == name for s in target.Sheets if s is not sheet):
            name += " (2)"
        sheet.Name = name
        sheet.Visible = self.Visible
        target.ActiveSheet = sheet

    def Move(self, before=None, after=None):
        coll = self.Parent.Sheets.items
        coll.remove(self)
        coll.insert(coll.index(before) if before is not None else coll.index(after) + 1, self)


class Books(Collection):
    def __init__(self, app):
        super().__init__()
        self.app = app

    def Add(self, template=None):
        wb = Workbook(self.app, f"Book{len(self.items) + 1}")
        self.items.append(wb)
        self.app.ActiveWorkbook = wb
        return wb


class Workbook:
    def __init__(self, app, name):
        self.Application, self.Name, self.FullName, self.Path = app, name, name, ""
        self.AutoSaveOn = self.ReadOnly = False
        self.Date1904 = self.ProtectStructure = self.ProtectWindows = False
        self.secret = ""
        self.Saved = True
        self.Names, self.SlicerCaches = Names(), Collection()
        self.Sheets = self.Worksheets = Sheets(self)
        self.ActiveSheet = self.Worksheets.Add()
        self.ActiveSheet.Name = "Data"
        self.Windows = Collection([NS(Visible=True, Hwnd=app.Hwnd)])

    def Activate(self):
        self.Application.ActiveWorkbook = self

    def Close(self, save):
        assert not save
        self.Application.Workbooks.items.remove(self)

    def Protect(self, password, structure, windows, /):
        self.secret, self.ProtectStructure, self.ProtectWindows = password, structure, windows

    def Unprotect(self, password, /):
        if password != self.secret:
            raise pywintypes.com_error(-2147352567, "Wrong password: " + password, None, None)
        self.ProtectStructure = self.ProtectWindows = False


class Excel:
    def __init__(self):
        self.Hwnd = 42
        self.Ready = self.Interactive = True
        self.EnableEvents = self.DisplayAlerts = self.ScreenUpdating = True
        self.Calculation = -4105
        self.Workbooks = Books(self)
        self.ActiveWorkbook = None
        self.Selection = None
        self.International = {2: ",", 3: " ", 4: ";"}
        self.WorksheetFunction = NS(CountA=lambda rng: sum(v is not None for row in to_grid(rng.Value) for v in row))

    @property
    def ActiveSheet(self):
        return self.ActiveWorkbook.ActiveSheet if self.ActiveWorkbook else None

    @property
    def ActiveWindow(self):
        return self.ActiveSheet.view

    def Calculate(self):
        pass

    def Goto(self, rng, scroll, /):
        assert scroll is True
        rng.Worksheet.Activate()
        self.Selection = rng
        self.ActiveWindow.ScrollRow, self.ActiveWindow.ScrollColumn = rng.Row, rng.Column

    def Intersect(self, a, b, /):
        if a.Worksheet is not b.Worksheet:
            return None
        r, c, end, right = max(a.Row, b.Row), max(a.Column, b.Column), min(a.r2, b.r2), min(a.c2, b.c2)
        return a.Worksheet.Range(a1_range(r, c, end, right)) if r <= end and c <= right else None


class HeaderFooter:
    def __init__(self):
        self.Exists = False
        self._range = NS(Text="")

    @property
    def Range(self):
        assert self.Exists, "Do not read a missing header/footer"
        return self._range


class Section:
    def __init__(self):
        self.Headers = Collection(HeaderFooter() for _ in range(3))
        self.Footers = Collection(HeaderFooter() for _ in range(3))


class Content:
    """Range документа. WordOpenXML собирается из состояния документа, как у настоящего Word: колонтитулы, примечания и
    сноски — отдельные части пакета, а идентификаторы абзацев и сеансов правки при каждом чтении новые (шум)."""

    _reads = 0

    def __init__(self, doc, text):
        self._doc, self.Text = doc, text

    def __deepcopy__(self, memo):
        return Content(self._doc, self.Text)

    @property
    def WordOpenXML(self):  # noqa: N802
        Content._reads += 1
        noise = f'w14:paraId="{Content._reads:08X}" w14:textId="77777777" w:rsidR="{Content._reads:08X}"'
        doc = self._doc
        parts = [("/word/document.xml", f"<w:body><w:p {noise}><w:t>{self.Text}</w:t></w:p></w:body>")]
        if any(hasattr(table, "data") for table in doc.Tables):
            parts[0] = (parts[0][0], parts[0][1] + repr([table.data for table in doc.Tables]))
        number = 0
        for section in doc.__dict__.get("_xml_sections") or doc.Sections:  # тест может спрятать Sections от кода
            for kind in ("Headers", "Footers"):
                for part in getattr(section, kind):
                    number += 1
                    if part.Exists:
                        parts.append((f"/word/{kind[:-1].lower()}{number}.xml", f"<w:p {noise}><w:t>{part._range.Text}</w:t></w:p>"))
        if doc.Comments.Count:
            parts.append(("/word/comments.xml", "".join(f"<w:comment><w:p {noise}><w:t>{c.Range.Text}</w:t></w:p></w:comment>" for c in doc.Comments)))
        parts.append(("/word/settings.xml", f'<w:rsids><w:rsid w:val="{Content._reads:08X}"/></w:rsids>'))
        return "<pkg:package>" + "".join(f'<pkg:part pkg:name="{n}">{b}</pkg:part>' for n, b in parts) + "</pkg:package>"


class Document:
    native_fields = ("Content", "Paragraphs", "Tables", "InlineShapes", "Sections", "Comments", "Footnotes", "Endnotes", "Shapes")

    def __init__(self, app):
        self.Application = app
        self.ActiveWindow = NS(Hwnd=84)
        self.Name, self.FullName, self.Path = "Draft.docx", "Draft.docx", ""
        self.AutoSaveOn = False
        self.Saved, self.ReadOnly, self.TrackRevisions, self.ProtectionType = True, False, False, -1
        self.Revisions = Collection()
        self.Content = Content(self, "original")
        self.Paragraphs, self.Tables, self.InlineShapes = Collection([NS()]), Collection(), Collection()
        self.Sections, self.Comments = Collection([Section()]), Collection()
        self.Footnotes, self.Endnotes, self.Shapes = Collection(), Collection(), Collection()
        self.undo_calls, self.states = [], []
        self.undo_ok = True

    def save_state(self):
        self.states.append(copy.deepcopy({name: getattr(self, name) for name in self.native_fields}))

    def Undo(self, count, /):
        assert count == 1
        self.undo_calls.append(count)
        if not self.undo_ok or not self.states:
            return False
        self.__dict__.update(self.states.pop())
        return True

    def Close(self, save, /):
        assert save == 0
        self.Application.Documents.items.remove(self)


class Word:
    def __init__(self):  # как у настоящего Word.Application: свойства Hwnd нет
        self.Documents = Collection([Document(self)])
        self.ActiveDocument = self.Documents(1)
        self.starts, self.ends = [], 0
        self.AutomationSecurity = 1
        self.UndoRecord = NS(IsRecordingCustomRecord=False, StartCustomRecord=self.start, EndCustomRecord=self.end)

    def start(self, label):
        self.starts.append(label)
        self.ActiveDocument.save_state()
        self.UndoRecord.IsRecordingCustomRecord = True

    def end(self):
        self.ends += 1
        self.UndoRecord.IsRecordingCustomRecord = False


@pytest.fixture(name="office")
def fake_office(monkeypatch, tmp_path):
    from office_live import bridge, com, config, excel_analysis, excel_core, excel_format, excel_pivot, journal, navigation, registry, templates, undo, wd_common, window, word_core, word_tables, word_layout, word_edits

    excel, word = Excel(), Word()
    wb = excel.Workbooks.Add()
    doc = word.ActiveDocument
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_JOURNAL_DIR": str(tmp_path / "journal")}))
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [excel if kind == "excel" else word])
    monkeypatch.setattr(undo, "STACKS", {})
    monkeypatch.setattr(undo, "_active", None)
    monkeypatch.setattr(navigation, "RENAMES", {})
    monkeypatch.setattr(window, "_LAST", [])
    saved_catalog = dict(registry.CATALOG)
    registered = {}
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda fn, name, **kw: registered.__setitem__(name, fn)))
    modules = (excel_core, excel_format, excel_analysis, excel_pivot, templates, bridge, journal, navigation, undo, window, word_core, word_tables, word_layout, word_edits)

    def call(tool_name, **kwargs):
        fn = next(getattr(m, tool_name) for m in modules if hasattr(m, tool_name))
        info = registry.CATALOG[tool_name]
        registry.office_tool(info.group, info.kind, read_actions=info.read_actions, file_args=info.file_args, preview_arg=info.preview_arg)(fn)
        return registered[tool_name](**kwargs)

    def word_write(text, fail=False, tool="word_insert_text"):
        def write(document, text):
            _, target = wd_common.pick_document(document)
            target.Content.Text = text
            if fail:
                raise ValueError("simulated failure")
            return {"ok": True}

        write.__name__ = tool
        registry.office_tool("word_core", "write")(write)
        return registered[tool](document=doc.Name, text=text)

    yield NS(excel=excel, word=word, wb=wb, ws=wb.ActiveSheet, doc=doc, call=call, word_write=word_write, tmp=tmp_path)
    registry.CATALOG.clear()
    registry.CATALOG.update(saved_catalog)
