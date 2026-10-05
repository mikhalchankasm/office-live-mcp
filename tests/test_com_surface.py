"""COM members used by the journal/undo code exist on the real Excel/Word interfaces, with matching argument order and
writable properties. Reads the registered type libraries only (LoadRegTypeLib never starts Office); skipped where Office
is not installed, e.g. on CI runners."""

import pythoncom
import pytest
import pywintypes


EXCEL = {
    "_Application": "Hwnd ScreenUpdating ActiveWorkbook ActiveSheet Workbooks DisplayAlerts EnableEvents Calculation ActiveWindow CutCopyMode Calculate Selection Intersect International",
    "Workbooks": "Count Item Add",
    "_Workbook": "Name FullName Path Names Worksheets Sheets Windows Saved Application Activate Close SlicerCaches Date1904 Protect Unprotect ProtectStructure ProtectWindows ReadOnly",
    "_Worksheet": "Names Name Parent Rows Columns Cells Range UsedRange Shapes ListObjects PivotTables Comments CommentsThreaded Index Visible Tab Activate Delete Copy Move StandardWidth ProtectContents",
    "Sheets": "Count Item Add",
    "Names": "Count Item Add",
    "Name": "Name RefersTo Visible Delete",
    "Range": "Address Row Column Worksheet Rows Columns Cells Item Count Hidden Value Value2 HasFormula Formula Formula2 Font Interior NumberFormat HorizontalAlignment VerticalAlignment WrapText MergeCells MergeArea PrefixCharacter Areas ColumnWidth RowHeight Copy PasteSpecial UnMerge Insert Delete End FormatConditions Borders ShrinkToFit IndentLevel Orientation ReadingOrder AddIndent Locked FormulaHidden Hyperlinks SpecialCells Comment Validation ClearComments Style Select",
    "Areas": "Count",
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
    "PivotTables": "Count Item",
    "PivotTable": "Name TableRange2",
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
    "Protection": "AllowFormattingCells AllowFormattingColumns AllowFormattingRows AllowInsertingColumns AllowInsertingRows AllowInsertingHyperlinks AllowDeletingColumns AllowDeletingRows AllowSorting AllowFiltering AllowUsingPivotTables",
    "ListColumns": "Item Count",
    "ListColumn": "DataBodyRange",
}
EXCEL["_Worksheet"] += " Protect Unprotect Protection ProtectDrawingObjects ProtectScenarios ProtectionMode EnableSelection"
EXCEL["Range"] += " DirectPrecedents DirectDependents HasSpill"
EXCEL["Areas"] += " Item"
EXCEL["Name"] += " RefersToRange"
EXCEL["ListObject"] += " ListColumns"
EXCEL["ListObjects"] += " Add"
EXCEL["_Worksheet"] += " Hyperlinks"
EXCEL["Hyperlinks"] += " Add"
EXCEL["_Application"] += " Goto"
WORD = {
    "_Application": "UndoRecord Documents ActiveDocument CompareDocuments AutomationSecurity",
    "_Document": "Name FullName Path Content Paragraphs Tables InlineShapes Sections Comments Footnotes Endnotes Shapes Undo Close Revisions Saved TrackRevisions ProtectionType",
    "Documents": "Open Count Item",
    "Range": "Text WordOpenXML Cells LanguageID",
    "Table": "Uniform Tables Rows Columns Range Cell Sort NestingLevel",
    "Rows": "Count",
    "Columns": "Count",
    "Cells": "Count",
    "Cell": "Range",
    "Revisions": "Count Item",
    "Revision": "Type Author Range",
    "UndoRecord": "IsRecordingCustomRecord StartCustomRecord EndCustomRecord",
    "Section": "Headers Footers",
    "HeaderFooter": "Exists Range",
    "Comment": "Range",
    "Paragraphs": "Count",
    "Tables": "Count Item",
    "InlineShapes": "Count",
    "Sections": "Count Item",
    "HeadersFooters": "Item",
    "Comments": "Count Item",
    "Footnotes": "Count",
    "Endnotes": "Count",
    "Shapes": "Count",
}
WORD["_Document"] += " Activate Range Bookmarks"
WORD["_Application"] += " ActiveWindow"
WORD["Range"] += " Select Start End"
WORD["Paragraphs"] += " Item"
WORD["Paragraph"] = "Range"
WORD["Window"] = "Hwnd ScrollIntoView"
WORD["Bookmarks"] = "Exists Item"
WORD["Bookmark"] = "Range"
CALLS = {
    "Hyperlinks.Add": ("Anchor", "Address", "SubAddress", "ScreenTip", "TextToDisplay"),
    "_Application.Goto": ("Reference", "Scroll"),
    "Window.ScrollIntoView": ("obj", "Start"),
    "Bookmarks.Exists": ("Name",), "Bookmarks.Item": ("Index",),
    "_Document.Range": ("Start", "End"), "_Document.Activate": (), "Range.Select": (),
    "Paragraphs.Item": ("Index",),
    "_Worksheet.Protect": ("Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly", "AllowFormattingCells", "AllowFormattingColumns",
                           "AllowFormattingRows", "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks", "AllowDeletingColumns",
                           "AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables"),
    "_Worksheet.Unprotect": ("Password",),
    "_Workbook.Protect": ("Password", "Structure", "Windows"),
    "_Workbook.Unprotect": ("Password",),
    "Range.DirectPrecedents": (), "Range.DirectDependents": (), "Range.HasSpill": (), "Range.Locked": (),
    "Name.RefersToRange": (), "ListColumn.DataBodyRange": (), "ListColumns.Item": ("Index",), "Areas.Item": ("Index",),
    "_Worksheet.ProtectionMode": (), "_Worksheet.EnableSelection": (), "_Workbook.Date1904": (),
    "ListObjects.Add": ("SourceType", "Source", "LinkSource", "XlListObjectHasHeaders"),
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
    "_Document.Close": ("SaveChanges",),
    "Documents.Open": ("FileName", "ConfirmConversions", "ReadOnly", "AddToRecentFiles", "PasswordDocument", "PasswordTemplate", "Revert",
                       "WritePasswordDocument", "WritePasswordTemplate", "Format", "Encoding", "Visible", "OpenAndRepair", "DocumentDirection", "NoEncodingDialog"),
    "_Application.CompareDocuments": ("OriginalDocument", "RevisedDocument", "Destination", "Granularity", "CompareFormatting", "CompareCaseChanges",
                                      "CompareWhitespace", "CompareTables", "CompareHeaders", "CompareFootnotes", "CompareTextboxes", "CompareFields",
                                      "CompareComments", "CompareMoves", "RevisedAuthor", "IgnoreAllComparisonWarnings"),
    "Table.Sort": ("ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", "FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", "SortFieldType3",
                   "SortOrder3", "CaseSensitive", "BidiSort", "IgnoreThe", "IgnoreKashida", "IgnoreDiacritics", "IgnoreHe", "LanguageID"),
    "Table.Cell": ("Row", "Column"),
    "Documents.Item": ("Index",),
    "Revisions.Item": ("Index",),
    "Tables.Item": ("Index",),
    "PivotTables.Item": ("Index",),
    "Range.Item": ("RowIndex", "ColumnIndex"),
    "Range.Value2": (),
    "Range.PrefixCharacter": (),
    "Range.HasFormula": (),
    "Range.MergeArea": (),
    "Table.Uniform": (),
    "Table.NestingLevel": (),
    "UndoRecord.StartCustomRecord": ("Name",),
}
CALLS.update({"Protection." + member: () for member in EXCEL["Protection"].split()})
SETTERS = {
    "_Application": "DisplayAlerts EnableEvents Calculation CutCopyMode",
    "_Workbook": "Saved Date1904",
    "_Worksheet": "Name Visible EnableSelection",
    "Range": "Value Value2 Formula Formula2 ColumnWidth RowHeight Hidden NumberFormat HorizontalAlignment VerticalAlignment WrapText ShrinkToFit IndentLevel Orientation Style Locked",
    "Font": "Name Size Bold Italic Underline Color Strikethrough ThemeColor TintAndShade",
    "Interior": "Color ThemeColor TintAndShade Pattern",
    "Border": "Color ThemeColor TintAndShade Weight LineStyle",
    "Window": "Visible SplitRow SplitColumn FreezePanes Zoom DisplayGridlines DisplayHeadings ScrollRow ScrollColumn",
    "Tab": "Color ColorIndex",
}
WORD_SETTERS = {"_Application": "AutomationSecurity"}
EXCEL_READONLY = {"Range": "DirectPrecedents DirectDependents", "Protection": EXCEL["Protection"],
                  "_Workbook": "ProtectStructure ProtectWindows", "Name": "RefersToRange"}


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
        for interface, names in EXCEL_READONLY.items():
            for member in names.split():
                entries = found[interface][member]
                assert any(entry[1] == 2 for entry in entries), (interface, member, "not a property getter")
                assert not any(entry[1] in (4, 8) for entry in entries), (interface, member, "unexpected setter")
    setters = SETTERS if wanted is EXCEL else WORD_SETTERS
    if setters:
        for interface, names in setters.items():
            for member in names.split():
                assert any(entry[1] in (4, 8) for entry in found[interface][member]), (interface, member, "read-only")
        print("Verified property setters: " + "; ".join(k + "." + v.replace(" ", ",") for k, v in setters.items()))
    for key, parameters in CALLS.items():
        if wanted is WORD and key == "Range.End":
            parameters = ()  # Word property, unlike Excel's Range.End(Direction)
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


def test_word_comparison_and_sort_enum_values():
    try:
        lib = pythoncom.LoadRegTypeLib(*WORD_LIB, 0)
    except pythoncom.com_error:
        pytest.skip("the Office type library is not registered on this machine")
    expected = {"wdCompareDestinationNew": 2, "wdGranularityCharLevel": 0, "wdGranularityWordLevel": 1,
                "wdSortFieldAlphanumeric": 0, "wdSortFieldNumeric": 1, "wdSortFieldDate": 2,
                "wdSortOrderAscending": 0, "wdSortOrderDescending": 1}
    found = {}
    for index in range(lib.GetTypeInfoCount()):
        ti = lib.GetTypeInfo(index)
        if ti.GetDocumentation(-1)[0] not in {"WdCompareDestination", "WdGranularity", "WdSortFieldType", "WdSortOrder"}:
            continue
        for j in range(ti.GetTypeAttr().cVars):
            desc = ti.GetVarDesc(j)
            name = ti.GetNames(desc.memid)[0]
            if name in expected:
                found[name] = desc.value
    assert found == expected


if __name__ == "__main__":
    count = verify(*EXCEL_LIB, EXCEL)
    count += verify(*WORD_LIB, WORD)
    print(f"Verified {count} interface/member pairs; no Office activation.")
