"""Watermark XML reads, linked-section boundaries, COM failures and native undo on fakes."""

from dataclasses import replace
from types import SimpleNamespace as NS
from xml.sax.saxutils import escape, quoteattr

import pytest

from office_live import config, registry, undo, word_layout as layout
from office_live.errors import ToolError
from tests.history_fakes import Collection, Content, fake_office  # noqa: F401


STORIES = {1: 7, 2: 10, 3: 6}


def anchor_for(section, typ):
    return NS(Text="", StoryType=STORIES[typ], Sections=lambda i: NS(Index=section))


class Shape:
    def __init__(self, collection, name, text, anchor=None):
        self.collection, self.Name, self.text = collection, name, text
        self.Anchor = anchor or collection.anchor
        self.Width = len(text) * 25
        self.Fill = NS(Solid=lambda: None, ForeColor=NS(), Transparency=0)
        self.Line, self.WrapFormat = NS(), NS()

    def Delete(self):
        self.collection.items.remove(self)


class Shapes(Collection):
    fail = False

    def AddTextEffect(self, preset, text, font, size, bold, italic, left, top, anchor, /):
        assert (preset, size, bold, italic, left, top) == (0, 36, 0, 0, 0, 0)
        shape = Shape(self, "WordArt1", text, anchor)
        self.items.append(shape)
        if self.fail:
            raise ValueError("injected after insertion")
        return shape


class Header:
    def __init__(self, section=1, typ=1):
        self.exists, self.touched, self.Shapes = False, 0, Shapes()
        self.Shapes.anchor = anchor_for(section, typ)

    @property
    def Range(self):
        self.exists = True
        self.touched += 1
        return self.Shapes.anchor


class XMLContent(Content):
    @property
    def WordOpenXML(self):  # noqa: N802
        parts, specs, rels = [], [], []
        for i, sec in enumerate(self._doc.Sections, 1):
            refs = []
            for typ, hf in enumerate(sec.Headers, 1):
                if not hf.exists:
                    continue
                key = str(i * 10 + typ)  # Word uses header<number>.xml; undo fingerprints these package parts.
                refs.append(f'<w:headerReference w:type="{["default", "first", "even"][typ - 1]}" r:id="{key}"/>')
                rels.append(f'<Relationship Id="{key}" Target="header{key}.xml"/>')
                shapes = ''.join(f'<v:shape id={quoteattr(s.Name)}><v:textpath string={quoteattr(s.text)}/></v:shape>' for s in hf.Shapes)
                parts.append((f"/word/header{key}.xml", f"<w:hdr>{shapes}</w:hdr>"))
            specs.append("<w:sectPr>" + ''.join(refs) + ("<w:titlePg/>" if sec.first else "") + "</w:sectPr>")
        body = ''.join(f"<w:p><w:pPr>{s}</w:pPr></w:p>" for s in specs[:-1]) + specs[-1]
        parts += [("/word/document.xml", f"<w:document><w:body><w:p><w:r><w:t>{escape(self.Text)}</w:t></w:r></w:p>{body}</w:body></w:document>"),
                  ("/word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(rels) + '</Relationships>'),
                  ("/word/settings.xml", "<w:settings><w:evenAndOddHeaders/></w:settings>")]
        ns = ' '.join(f'xmlns:{k}="{v}"' for k, v in layout._WM_NS.items())
        return f'<pkg:package {ns}>' + ''.join(f'<pkg:part pkg:name="{n}"><pkg:xmlData>{b}</pkg:xmlData></pkg:part>' for n, b in parts) + '</pkg:package>'

    def __deepcopy__(self, memo):
        return XMLContent(self._doc, self.Text)


@pytest.fixture
def o(office):
    office.doc.Sections = Collection(NS(Headers=Collection(Header(i, typ) for typ in (1, 2, 3)), first=True,
                                       PageSetup=NS(PageWidth=600, PageHeight=800)) for i in (1, 2))
    office.doc.Content = XMLContent(office.doc, "original")
    return office


def call(o, action="status", **kw):
    return o.call("word_watermark", document=o.doc.Name, action=action, **kw)


def test_status_readonly_protected_autosave_never_touches_headers(o, monkeypatch):
    o.doc.ProtectionType, o.doc.AutoSaveOn = 3, True
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    assert not call(o)["present"]
    assert not o.word.starts
    assert all(not h.touched and not h.exists for s in o.doc.Sections for h in s.Headers)
    with pytest.raises(ToolError, match="Read-only"):
        call(o, "add", text="Draft")
    assert registry.CATALOG["word_watermark"].read_actions == ("status",)


def test_add_inherited_first_even_status_and_both_undo_directions(o):
    hf = o.doc.Sections(1).Headers(1)
    hf.exists = True
    hf.Shapes.items.append(Shape(hf.Shapes, "Logo", "untouched"))
    before = undo.word_fingerprint(o.doc, styles=False)
    result = call(o, "add", text="ЧЕРНОВИК", color="#123456")
    assert len(result["changed"]) == 3 and result["undo"] == "available"
    status = call(o)
    assert len(status["watermarks"]) == 3
    assert all(m["sections"] == [1, 2] and m["text"] == "ЧЕРНОВИК" for m in status["watermarks"])
    shape = hf.Shapes(2)
    assert (shape.Fill.ForeColor.RGB, shape.Fill.Transparency, shape.Line.Visible) == (0x563412, .5, 0)
    assert (shape.WrapFormat.Type, shape.Left, shape.Top, shape.Rotation) == (5, -999995, -999995, 315)
    added = undo.word_fingerprint(o.doc, styles=False)
    with pytest.raises(ToolError, match="already exists"):
        call(o, "add", text="again")
    assert len(o.word.starts) == 1
    assert len(call(o, "remove")["changed"]) == 3
    assert [s.Name for s in hf.Shapes] == ["Logo"] and not call(o)["present"]
    o.call("office_undo", file=o.doc.Name)
    assert undo.word_fingerprint(o.doc, styles=False) == added
    o.call("office_undo", file=o.doc.Name)
    assert undo.word_fingerprint(o.doc, styles=False) == before


@pytest.mark.parametrize("kw", [{"text": ""}, {"text": "x" * 51}, {"text": "a\nb"}, {"font": ""},
                               {"color": "wrong"}, {"transparency": float("nan")}, {"transparency": 1.1},
                               {"layout": "wrong"}, {"sections": []}, {"sections": [0]}, {"sections": [True]}, {"sections": [3]}])
def test_validation_before_undo_or_header_access(o, kw):
    with pytest.raises(ToolError):
        call(o, "add", **{"text": "Draft", **kw})
    assert not o.word.starts and not any(h.touched for s in o.doc.Sections for h in s.Headers)


@pytest.mark.parametrize("action", ["add", "remove"])
def test_subset_link_refused_before_header_access(o, action):
    if action == "remove":
        call(o, "add", text="Draft")
    before = undo.word_fingerprint(o.doc)
    starts = len(o.word.starts)
    with pytest.raises(ToolError, match="Linked header"):
        call(o, action, text="Draft", sections=[2])
    assert undo.word_fingerprint(o.doc) == before and len(o.word.starts) == starts


def test_independent_subset_and_horizontal(o):
    for hf in o.doc.Sections(2).Headers:
        hf.exists = True
    call(o, "add", text="Confidential", layout="horizontal", sections=[2])
    assert not call(o, sections=[1])["present"]
    assert o.doc.Sections(2).Headers(1).Shapes(1).Rotation == 0
    assert len(call(o, "remove", sections=[2])["changed"]) == 3


@pytest.mark.parametrize("guard", ["AutoSaveOn", "ProtectionType", "ReadOnly"])
def test_picker_guards(o, guard):
    setattr(o.doc, guard, 3 if guard == "ProtectionType" else True)
    with pytest.raises(ToolError):
        call(o, "add", text="Draft")
    assert not o.word.starts


def test_partial_failure_retains_undo_and_reports_uncertain_change(o):
    o.doc.Sections(1).Headers(2).Shapes.fail = True
    before = undo.word_fingerprint(o.doc, styles=False)
    with pytest.raises(ToolError, match="completed=.*possibly changed=.*office_undo"):
        call(o, "add", text="Draft")
    assert len(call(o)["watermarks"]) == 1
    o.call("office_undo", file=o.doc.Name)
    assert undo.word_fingerprint(o.doc, styles=False) == before


def test_long_text_fits_the_smallest_linked_page(o):
    o.doc.Sections(2).PageSetup.PageWidth = 300
    call(o, "add", text="W" * 50)
    shape = o.doc.Sections(1).Headers(1).Shapes(1)
    assert shape.Width < 300 and shape.LockAspectRatio == -1


def test_drawingml_and_vml_fallback_report_once_without_xml_whitespace(o):
    hf = o.doc.Sections(1).Headers(1)
    hf.exists = True
    name = layout.WATERMARK_PREFIX + "42"
    hf.Shapes.items.append(Shape(hf.Shapes, name, "D & D"))
    xml = o.doc.Content.WordOpenXML.replace('</w:hdr>', f'<wp:anchor>\n<wp:docPr name="{name}"/>\n<w:t>D &amp; D</w:t>\n</wp:anchor></w:hdr>')
    o.doc.Content = NS(Text="original", WordOpenXML=xml)
    assert call(o)["watermarks"] == [{"name": name, "text": "D & D", "sections": [1, 2], "header": "default"}]


@pytest.mark.parametrize("guard", ["strict", "folder", "tracking", "undo"])
def test_preflight_guards_preserve_document(o, monkeypatch, guard):
    name = o.doc.Name
    if guard == "strict":
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, strict_target=True))
        name = ""
    elif guard == "folder":
        o.doc.Path, o.doc.FullName = str(o.tmp), str(o.tmp / o.doc.Name)
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, allowed_dirs=(o.tmp / "allowed",)))
    elif guard == "tracking":
        o.doc.ProtectionType = 0
    else:
        o.word.UndoRecord.IsRecordingCustomRecord = True
    before = undo.word_fingerprint(o.doc)
    with pytest.raises(ToolError):
        o.call("word_watermark", document=name, action="add", text="Draft")
    assert undo.word_fingerprint(o.doc) == before and not o.word.starts


def test_remove_ignores_other_sections_shapes_in_the_document_wide_collection(o):
    # Word quirk: HeaderFooter.Shapes lists the shapes of ALL headers/footers. Removing section 1's watermark must not
    # delete section 2's even though Word shows it in section 1's collection too.
    for sec in (1, 2):
        for hf in o.doc.Sections(sec).Headers:
            hf.exists = True
    call(o, "add", text="DRAFT", sections=[1])
    call(o, "add", text="DRAFT", sections=[2])
    own, other = o.doc.Sections(1).Headers(1).Shapes, o.doc.Sections(2).Headers(1).Shapes
    foreign = other(1)
    own.items.append(foreign)  # what Word returns: section 2's shape visible through section 1's header
    changed = call(o, "remove", sections=[1])["changed"]
    assert all(c["sections"] == [1] for c in changed) and len(changed) == 3
    assert foreign in other.items and call(o, sections=[2])["present"]
