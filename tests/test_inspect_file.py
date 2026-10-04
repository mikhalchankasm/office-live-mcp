"""Разбор xlsx/docx без Office: порядок атрибутов, inlineStr, потоковый разбор, потолок размера частей."""

import zipfile

import pytest

from office_live import inspect_file
from office_live.errors import ToolError

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

WORKBOOK = f"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="{MAIN}" xmlns:r="{REL}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet state="hidden" name="Calc" sheetId="2" r:id="rId2"/></sheets>
  <definedNames><definedName hidden="1" name="_xlnm._FilterDatabase">Data!$A$1:$C$3</definedName><definedName name="Total">Data!$C$4</definedName></definedNames>
</workbook>"""
RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="x"/>
  <Relationship Id="rId2" Target="/xl/worksheets/sheet2.xml" Type="x"/>
</Relationships>"""
STRINGS = f"""<?xml version="1.0"?><sst xmlns="{MAIN}"><si><t>Region</t></si><si><r><t>Rev</t></r><r><t>enue</t></r></si><si><t>A &amp; B</t></si></sst>"""
# атрибуты в «неудобном» порядке, inlineStr, общая строка, формула, слитые ячейки, правило УФ с формулой
SHEET1 = f"""<?xml version="1.0"?>
<worksheet xmlns="{MAIN}">
  <dimension ref="A1:C4"/>
  <sheetViews><sheetView zoomScale="85"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols><col width="12" hidden="1" min="3" max="3" customWidth="1"/><col min="1" max="2" width="20"/></cols>
  <sheetData>
    <row spans="1:3" r="1"><c t="s" s="1" r="A1"><v>0</v></c><c s="1" t="s" r="B1"><v>1</v></c><c r="C1" t="inlineStr"><is><t>Inline</t></is></c></row>
    <row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>10</v></c><c r="C2"><f>B2*2</f><v>20</v></c></row>
    <row r="3"><c r="A3" t="b"><v>1</v></c><c r="B3"><f t="shared" si="0" ref="B3:B4">B2+1</f><v>11</v></c><c r="C3"/></row>
    <row r="4"><c r="B4"><f t="shared" si="0"/><v>12</v></c></row>
  </sheetData>
  <autoFilter ref="A1:C3"/>
  <mergeCells count="1"><mergeCell ref="A4:C4"/></mergeCells>
  <conditionalFormatting sqref="B2:B4"><cfRule priority="1" type="expression" dxfId="0"><formula>$B2&gt;10</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="A2"><formula1>"a,b"</formula1></dataValidation></dataValidations>
  <drawing r:id="rId9" xmlns:r="{REL}"/>
</worksheet>"""
SHEET2 = f"""<?xml version="1.0"?><worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"""


def make_xlsx(path, sheet1=SHEET1):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", RELS)
        z.writestr("xl/sharedStrings.xml", STRINGS)
        z.writestr("xl/worksheets/sheet1.xml", sheet1)
        z.writestr("xl/worksheets/sheet2.xml", SHEET2)
        z.writestr("xl/styles.xml", f'<styleSheet xmlns="{MAIN}"><fonts count="1"><font/></fonts><fills count="2"><fill/><fill/></fills><cellXfs count="2"><xf/><xf/></cellXfs><dxfs count="1"><dxf/></dxfs></styleSheet>')
    return str(path)


def test_xlsx_sheets_names_and_hidden_state(tmp_path):
    info = inspect_file.inspect_xlsx(make_xlsx(tmp_path / "a.xlsx"))
    assert [(s["name"], s["state"]) for s in info["sheets"]] == [("Data", "visible"), ("Calc", "hidden")]
    names = {n["name"]: n for n in info["defined_names"]}
    assert names["_xlnm._FilterDatabase"]["hidden"] is True and names["Total"]["hidden"] is False  # не зависит от порядка атрибутов
    assert info["styles"] == {"cell_formats": 2, "fonts": 1, "fills": 2, "dxfs": 1}


def test_xlsx_sheet_details_independent_of_attribute_order(tmp_path):
    data = inspect_file.inspect_xlsx(make_xlsx(tmp_path / "a.xlsx"))["sheets"][0]
    assert data["dimension"] == "A1:C4" and data["zoom"] == 85
    assert data["freeze"] == {"ySplit": "1", "topLeftCell": "A2", "state": "frozen"}
    assert data["hidden_columns"] == ["3-3"] and data["columns_defined"] == 2
    assert data["merged_ranges"] == {"count": 1, "examples": ["A4:C4"]}
    assert data["conditional_formats"]["rules"] == [{"applies_to": "B2:B4", "type": "expression", "priority": "1", "formula": "$B2>10"}]
    assert data["data_validations"] == 1 and data["has_autofilter"] and data["has_drawing"]
    assert data["rows_with_cells"] == 4 and data["formula_cells"] == 3  # две общие формулы тоже считаются


def test_xlsx_peek_reads_shared_inline_bool_and_formula_cells(tmp_path):
    peek = inspect_file.inspect_xlsx(make_xlsx(tmp_path / "a.xlsx"), peek_rows=4)["sheets"][0]["peek"]
    row1 = {c["cell"]: c["value"] for c in peek[0]["cells"]}
    assert row1 == {"A1": "Region", "B1": "Revenue", "C1": "Inline"}  # rich-text и inlineStr читаются
    row2 = {c["cell"]: c for c in peek[1]["cells"]}
    assert row2["A2"]["value"] == "A & B" and row2["C2"]["formula"] == "B2*2" and row2["C2"]["value"] == "20"
    row3 = {c["cell"]: c for c in peek[2]["cells"]}
    assert row3["A3"]["value"] == "TRUE" and "C3" not in row3  # пустая ячейка без значения пропущена
    assert row3["B3"]["formula"] == "B2+1"


def test_xlsx_peek_can_be_switched_off(tmp_path):
    assert "peek" not in inspect_file.inspect_xlsx(make_xlsx(tmp_path / "a.xlsx"), peek_rows=0)["sheets"][0]


def test_oversized_part_is_refused_not_loaded(tmp_path, monkeypatch):
    path = make_xlsx(tmp_path / "a.xlsx")
    with zipfile.ZipFile(path) as z:
        small = max(z.getinfo(n).file_size for n in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/sharedStrings.xml", "xl/styles.xml"))
        big = z.getinfo("xl/worksheets/sheet1.xml").file_size
    assert big > small
    monkeypatch.setattr(inspect_file, "MAX_PART_BYTES", small)  # служебные части влезают, большой лист — нет
    info = inspect_file.inspect_xlsx(path)
    assert "safety limit" in info["sheets"][0]["skipped"] and "rows_with_cells" in info["sheets"][1]
    monkeypatch.setattr(inspect_file, "MAX_PART_BYTES", 10)
    with pytest.raises(ToolError, match="safety limit"):
        inspect_file.inspect_xlsx(path)  # даже workbook.xml не читаем


def test_docx_paragraphs_tables_and_placeholders(tmp_path):
    body = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Dear {{Name}}</w:t></w:r></w:p>'
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>'
    )
    path = tmp_path / "a.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", body)
    info = inspect_file.inspect_docx(str(path))
    assert info["tables"] == 1 and info["paragraphs"] == 3 and info["placeholders"] == ["Name"]
    assert info["peek"][0] == {"style": "Heading1", "text": "Title"} and info["sections"] == 1
