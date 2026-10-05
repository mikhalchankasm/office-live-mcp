"""COM members used by the journal/undo code exist on the real Excel/Word interfaces, with matching argument order and
writable properties. Reads the registered type libraries only (LoadRegTypeLib never starts Office); skipped where Office
is not installed, e.g. on CI runners."""

import pythoncom
import pytest
import pywintypes


EXCEL = {
    "_Application": "Hwnd ScreenUpdating ActiveWorkbook ActiveSheet Workbooks DisplayAlerts EnableEvents Calculation ActiveWindow CutCopyMode Calculate Selection Intersect",
    "Workbooks": "Count Item Add",
    "_Workbook": "Name FullName Path Names Worksheets Sheets Windows Saved Application Activate Close SlicerCaches",
    "_Worksheet": "Names Name Parent Rows Columns Cells Range UsedRange Shapes ListObjects PivotTables Comments CommentsThreaded Index Visible Tab Activate Delete Copy Move StandardWidth",
    "Sheets": "Count Item Add",
    "Names": "Count Item Add",
    "Name": "Name RefersTo Visible Delete",
    "Range": "Address Row Column Worksheet Rows Columns Cells Count Hidden Value Formula Formula2 Font Interior NumberFormat HorizontalAlignment VerticalAlignment WrapText MergeCells ColumnWidth RowHeight Copy PasteSpecial UnMerge Insert Delete End FormatConditions Borders ShrinkToFit IndentLevel Orientation ReadingOrder AddIndent Locked FormulaHidden Hyperlinks SpecialCells Comment Validation ClearComments Style Select",
    "Font": "Name Size Bold Italic Underline Color Strikethrough Subscript Superscript OutlineFont Shadow TintAndShade ThemeColor",
    "Interior": "Color Pattern PatternColor TintAndShade PatternTintAndShade ThemeColor PatternThemeColor",
    "Style": "Name",
    "Border": "LineStyle Weight Color ThemeColor TintAndShade",
    "Borders": "Item",
    "FormatConditions": "Count Item",
    "FormatCondition": "Type AppliesTo Priority StopIfTrue Formula1 Formula2 Operator Text TextOperator DateOperator",
    "Top10": "Rank Percent TopBottom",
    "AboveAverage": "AboveBelow NumStdDev",
    "UniqueValues": "DupeUnique",
    "Validation": "Type AlertStyle Operator Formula1 Formula2 IgnoreBlank InCellDropdown ShowError ErrorTitle ErrorMessage ShowInput InputTitle InputMessage Delete",
    "Comment": "Text Author Visible",
    "Comments": "Count",
    "CommentsThreaded": "Count",
    "PivotTables": "Count",
    "Hyperlinks": "Count Item",
    "Hyperlink": "Range Address SubAddress ScreenTip TextToDisplay",
    "Shapes": "Count Item",
    "ListObjects": "Count Item",
    "ListObject": "Unlist Name Range TableStyle ShowTotals ShowHeaders",
    "TableStyle": "Name",
    "Shape": "Delete Name Type Left Top Width Height Rotation Chart",
    "_Chart": "ChartType",
    "SlicerCaches": "Item",
    "SlicerCache": "Delete",
    "Tab": "Color ColorIndex",
    "Window": "Visible SplitRow SplitColumn FreezePanes Zoom DisplayGridlines DisplayHeadings ScrollRow ScrollColumn",
    "Windows": "Count Item",
}
WORD = {
    "_Application": "UndoRecord Documents ActiveDocument",
    "_Document": "Name FullName Path Content Paragraphs Tables InlineShapes Sections Comments Footnotes Endnotes Shapes Undo",
    "Range": "Text WordOpenXML",
    "UndoRecord": "IsRecordingCustomRecord StartCustomRecord EndCustomRecord",
    "Section": "Headers Footers",
    "HeaderFooter": "Exists Range",
    "Comment": "Range",
    "Paragraphs": "Count",
    "Tables": "Count",
    "InlineShapes": "Count",
    "Sections": "Count Item",
    "HeadersFooters": "Item",
    "Comments": "Count Item",
    "Footnotes": "Count",
    "Endnotes": "Count",
    "Shapes": "Count",
}
CALLS = {
    "_Application.Intersect": ("Arg1", "Arg2"),
    "Names.Add": ("Name", "RefersTo", "Visible"),
    "Names.Item": ("Index",),
    "Workbooks.Add": ("Template",),
    "Sheets.Add": ("Before", "After"),
    "_Workbook.Close": ("SaveChanges",),
    "_Worksheet.Copy": ("Before", "After"),
    "_Worksheet.Move": ("Before", "After"),
    "_Worksheet.PivotTables": (),
    "Range.Copy": ("Destination",),
    "Range.PasteSpecial": ("Paste", "Operation", "SkipBlanks", "Transpose"),
    "Range.End": ("Direction",),
    "Range.SpecialCells": ("Type",),
    "_Document.Undo": ("Times",),
    "UndoRecord.StartCustomRecord": ("Name",),
}
SETTERS = {
    "_Application": "DisplayAlerts EnableEvents Calculation CutCopyMode",
    "_Workbook": "Saved",
    "_Worksheet": "Name Visible",
    "Range": "Value Formula Formula2 ColumnWidth RowHeight Hidden NumberFormat HorizontalAlignment VerticalAlignment WrapText ShrinkToFit IndentLevel Orientation Style",
    "Font": "Name Size Bold Italic Underline Color Strikethrough ThemeColor TintAndShade",
    "Interior": "Color ThemeColor TintAndShade Pattern",
    "Border": "Color ThemeColor TintAndShade Weight LineStyle",
    "Window": "Visible SplitRow SplitColumn FreezePanes Zoom DisplayGridlines DisplayHeadings ScrollRow ScrollColumn",
    "Tab": "Color ColorIndex",
}


def verify(guid, major, minor, wanted):
    lib = pythoncom.LoadRegTypeLib(pywintypes.IID(guid), major, minor, 0)
    found = {}
    for index in range(lib.GetTypeInfoCount()):
        ti = lib.GetTypeInfo(index)
        name = ti.GetDocumentation(-1)[0]
        if name not in wanted:
            continue
        members = {}
        for j in range(ti.GetTypeAttr().cFuncs):
            desc = ti.GetFuncDesc(j)
            names = ti.GetNames(desc.memid)
            members.setdefault(names[0], []).append((names[1:], desc.invkind, len(desc.args), desc.cParamsOpt))
        missing = set(wanted[name].split()) - members.keys()
        assert not missing, (name, missing)
        found[name] = members
        print(name + ": " + wanted[name])
    assert wanted.keys() <= found.keys(), wanted.keys() - found.keys()
    if wanted is EXCEL:
        for interface, names in SETTERS.items():
            for member in names.split():
                assert any(entry[1] in (4, 8) for entry in found[interface][member]), (interface, member, "read-only")
        print("Verified property setters: " + "; ".join(k + "." + v.replace(" ", ",") for k, v in SETTERS.items()))
    for key, parameters in CALLS.items():
        interface, member = key.split(".")
        if interface not in wanted or member not in wanted[interface].split():
            continue
        matches = [entry for entry in found[interface][member] if entry[0][:len(parameters)] == parameters]
        assert matches, (key, parameters, found[interface][member])
        assert any(count - optional <= len(parameters) <= count for _, _, count, optional in matches), (key, "argument count")
        print(key, matches)
    return sum(len(names.split()) for names in wanted.values())


EXCEL_LIB = ("{00020813-0000-0000-C000-000000000046}", 1, 9)
WORD_LIB = ("{00020905-0000-0000-C000-000000000046}", 8, 7)


def _check(lib, wanted):
    try:
        pythoncom.LoadRegTypeLib(pywintypes.IID(lib[0]), lib[1], lib[2], 0)
    except pythoncom.com_error:
        pytest.skip("the Office type library is not registered on this machine")
    return verify(*lib, wanted)


def test_excel_members_used_by_history_exist():
    assert _check(EXCEL_LIB, EXCEL) > 0


def test_word_members_used_by_history_exist():
    assert _check(WORD_LIB, WORD) > 0
    assert "Hwnd" not in WORD.get("_Application", "")  # у Word.Application его нет — ключ отмены на него не опирается


if __name__ == "__main__":
    count = verify(*EXCEL_LIB, EXCEL)
    count += verify(*WORD_LIB, WORD)
    print(f"Verified {count} interface/member pairs; no Office activation.")
