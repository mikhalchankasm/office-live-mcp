"""Office operations COM contracts in memory. No Office objects or native windows."""

import copy
from types import SimpleNamespace as NS

import pywintypes

from office_live import excel_format as formatting, xl_common as xl
from tests.history_fakes import Collection, Range


def error():
    return pywintypes.com_error(-2147352567, "simulated COM failure", None, None)


def shift(ws, start, count):
    for rng in ws.range_refs:
        if rng.Row >= start:
            rng.Row += count
        if rng.r2 >= start:
            rng.r2 += count
        rng.Rows.Count = rng.r2 - rng.Row + 1
    for name in ("data", "formats", "prefixes", "results"):
        values = getattr(ws, name)
        setattr(ws, name, {(r + count if r >= start else r, c): v for (r, c), v in values.items()})


class AnalysisRange(Range):
    @property
    def Count(self):
        return int(self.Rows.Count) * int(self.Columns.Count)

    @property
    def Precedents(self):
        if self.Worksheet.no_precedents:
            raise error()
        return self.Worksheet.Range("A1")

    def GoalSeek(self, goal, changing, /):
        self.Worksheet.calls.append(("goal", goal, changing.Address))
        changing.Value2 = goal / 2
        self.Worksheet.results[self.Row, self.Column] = goal
        if self.Worksheet.fail == "goal":
            raise error()
        return self.Worksheet.converged

    def Sort(self, key, order, key2, kind, order2, key3, order3, header, custom, case, orientation, method, option1, option2, option3, subfield, /):
        assert (custom, case, orientation, method, option1, option2, option3, subfield) == (1, False, 1, 1, 0, 0, 0, None)
        self.Worksheet.calls.append(("sort", key.Address, order, key2, kind, order2, key3, order3, header))
        grid = list(self.Value2)
        col = key.Column - self.Column
        self.Value2 = tuple([grid[0]] + sorted(grid[1:], key=lambda row: row[col]))

    def Subtotal(self, key, function, totals, replace, breaks, below, /):
        ws = self.Worksheet
        original_end = self.r2
        ws.calls.append(("subtotal", key, function, totals, replace, breaks, below))
        data = self.Value2[1:]
        boundaries = [i + self.Row + 1 for i in range(len(data)) if i == len(data) - 1 or data[i][key - 1] != data[i + 1][key - 1]]
        # Whole row insertions preserve neighbouring and trailing cells independently of the block snapshot.
        for row in reversed(boundaries):
            shift(ws, row + 1, 1)
            ws.data[row + 1, self.Column] = "Subtotal"
            ws.data[row + 1, self.Column + totals[0] - 1] = "=SUBTOTAL(9,B2:B3)"
        end = original_end + len(boundaries) + 1
        shift(ws, end, 1)
        ws.data[end, self.Column] = "Grand total"
        ws.data[end, self.Column + totals[0] - 1] = "=SUBTOTAL(9,B2:B6)"
        ws.subtotal_rows = [r for (r, c), v in ws.data.items() if c == self.Column and v in {"Subtotal", "Grand total"}]
        for r in range(self.Row + 1, end + 1):
            ws.Rows(r).OutlineLevel = 2 if r not in ws.subtotal_rows else 1
        if ws.fail == "subtotal":
            raise error()

    def RemoveSubtotal(self, /):
        ws = self.Worksheet
        ws.calls.append(("remove",))
        for row in sorted([r for r in ws.subtotal_rows if self.Row <= r <= self.r2], reverse=True):
            for rng in ws.range_refs:
                if rng.Row > row:
                    rng.Row -= 1
                if rng.r2 > row:
                    rng.r2 -= 1
                rng.Rows.Count = rng.r2 - rng.Row + 1
            for name in ("data", "formats", "prefixes", "results"):
                values = getattr(ws, name)
                setattr(ws, name, {(r - 1 if r > row else r, c): v for (r, c), v in values.items() if r != row})
        ws.subtotal_rows = []
        for r in range(self.Row, self.r2 + 1):
            ws.Rows(r).OutlineLevel = 1

    @property
    def SparklineGroups(self):
        return SparkGroups(self)


def color():
    return NS(Visible=False, Color=NS(Color=0, TintAndShade=0, ThemeColor=0))


class Spark:
    def __init__(self, location, kind, source):
        self.Location, self.Type, self.SourceData = location, kind, source
        self.SeriesColor = NS(Color=0, TintAndShade=0, ThemeColor=0)  # live: SeriesColor is a FormatColor itself
        self.Points = NS(**{p: color() for p in formatting.SPARK_POINTS.values()})
        self.Axes = NS(Horizontal=NS(Axis=color(), RightToLeftPlotOrder=False),
                       Vertical=NS(MinScaleType=1, MaxScaleType=1, CustomMinScaleValue=0, CustomMaxScaleValue=0))
        self.DateRange, self.DisplayBlanksAs, self.DisplayHidden, self.LineWeight, self.PlotBy = "", 1, False, 0.75, 0

    def Delete(self, /):
        ws = self.Location.Worksheet
        if ws.fail == "delete_spark":
            raise error()
        ws.sparks.remove(self)


class SparkGroups:
    def __init__(self, rng):
        self.rng = rng

    def items(self):
        return [g for g in self.rng.Worksheet.sparks if xl.bounds(self.rng)[0] <= g.Location.Row <= self.rng.r2
                and self.rng.Column <= g.Location.Column <= self.rng.c2]

    @property
    def Count(self):
        return len(self.items())

    def Item(self, i, /):
        return self.items()[i - 1]

    def Add(self, kind, source, /):
        group = Spark(self.rng, kind, source)
        self.rng.Worksheet.sparks.append(group)
        if self.rng.Worksheet.fail == "add_spark":
            raise error()
        return group


def excel(o, monkeypatch):
    ws = o.ws
    ws.calls, ws.fail, ws.no_precedents, ws.converged = [], "", False, True
    ws.AutoFilterMode, ws.subtotal_rows, ws.sparks, ws.range_refs = False, [], [], []
    o.excel.MaxIterations, o.excel.MaxChange = 100, 0.001
    old_range = ws.Range

    def get_range(address):
        rng = AnalysisRange(ws, address)
        ws.range_refs.append(rng)
        for r in range(rng.Row, min(rng.r2, 200) + 1):
            if not hasattr(ws.Rows(r), "OutlineLevel"):
                ws.Rows(r).OutlineLevel = 1
        return rng

    monkeypatch.setattr(ws, "Range", get_range)
    return old_range


class WordRange:
    def __init__(self, doc, start, end):
        self.doc, self.Start, self.End = doc, start, end
        self.Cells, self.Tables, self.Editors = Collection(), Collection(), Collection()

    @property
    def Text(self):
        return self.doc.Content.Text[self.Start:self.End]

    @property
    def Paragraphs(self):
        return NS(Count=max(1, len(self.Text.rstrip("\r").split("\r"))))

    def Information(self, kind, /):
        assert kind == 12
        return bool(self.doc.in_table)

    def InsertFile(self, path, bookmark, conversions, link, attachment, /):
        assert self.doc.Application.AutomationSecurity == 3
        self.doc.calls.append(("insert", path, bookmark, conversions, link, attachment))
        text = self.doc.inserted_text
        old = self.doc.Content.Text
        self.doc.Content.Text = old[:self.Start] + text + old[self.Start:]
        refresh_word(self.doc)
        if self.doc.fail == "insert":
            raise error()

    def ConvertToTable(self, separator, rows, columns, /):
        self.doc.calls.append(("to_table", separator, rows, columns))
        sep = {1: "\t", 2: ","}.get(separator, separator)
        data = [r.split(sep) for r in self.Text.rstrip("\r").split("\r")]
        tbl = WordTable(self.doc, data, self.Start, self.End, columns)
        self.doc.Tables.items.append(tbl)
        if self.doc.fail == "to_table":
            raise error()
        return tbl


class WordTable:
    def __init__(self, doc, data, start=0, end=4, columns=None):
        self.doc, self.data, self.Style = doc, copy.deepcopy(data), ""
        self.Uniform, self.NestingLevel = True, 1
        self.Tables = Collection()
        self.Rows = Collection([NS(HeadingFormat=False) for _ in data])
        self.Columns = NS(Count=columns or len(data[0]))
        self.Range = WordRange(doc, start, end)
        self.Range.Cells = Collection([NS() for _ in range(len(data) * self.Columns.Count)])

    def ConvertToText(self, sep, nested, /):
        self.doc.calls.append(("to_text", sep, nested))
        self.doc.Tables.items.remove(self)
        delimiter = {0: "\r", 1: "\t", 2: ","}.get(sep, sep)
        self.doc.Content.Text = "\r".join(delimiter.join(r) for r in self.data) + "\r"
        refresh_word(self.doc)
        return self.doc.Range(0, self.doc.Content.End)


def refresh_word(doc):
    doc.Content.Start, doc.Content.End = 0, len(doc.Content.Text)
    paragraphs, start = [], 0
    for row in doc.Content.Text.splitlines(keepends=True):
        paragraphs.append(NS(Range=WordRange(doc, start, start + len(row))))
        start += len(row)
    doc.Paragraphs = Collection(paragraphs)


def word(o):
    doc = o.doc
    doc.calls, doc.fail, doc.in_table, doc.inserted_text = [], "", False, "inserted\r"
    doc.Range = lambda start, end: WordRange(doc, start, end)
    doc.Content.Text = "a\tb\rc\td\r"
    doc.Content.Editors = Collection()
    doc.Bookmarks = NS(Exists=lambda name: name == "anchor", Item=lambda name: NS(Range=NS(End=4)))
    doc.Styles = lambda name: NS(Type=3) if name == "Grid" else NS(Type=1)
    doc.secret = ""

    def protect(mode, no_reset, password, irm, lock, /):
        doc.calls.append(("protect", mode, no_reset, password, irm, lock))
        doc.ProtectionType, doc.secret = mode, password
        if mode == 0:
            doc.TrackRevisions = True

    def unprotect(password, /):
        doc.calls.append(("unprotect", password))
        if password != doc.secret:
            raise error()
        doc.ProtectionType = -1

    doc.Protect, doc.Unprotect = protect, unprotect
    refresh_word(doc)
    return doc
