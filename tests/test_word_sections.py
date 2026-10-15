"""Section boundaries, XML-only reads and native undo contracts; no Office activation."""

import copy
from dataclasses import replace
from types import SimpleNamespace as NS
from xml.sax.saxutils import escape

import pytest

from office_live import config, undo, word_layout as layout
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401


class Content:
    def __init__(self, doc, text):
        self.doc, self.Text = doc, text

    def __deepcopy__(self, memo):
        return Content(self.doc, self.Text)

    Start = 0

    @property
    def End(self):
        return len(self.Text)

    @property
    def WordOpenXML(self):
        props = []
        for s in self.doc.Sections:
            number = '<w:pgNumType w:start="7"/>' if s.Footers(1).PageNumbers.RestartNumberingAtSection else ''
            ref = '' if s.Footers(1).LinkToPrevious else '<w:footerReference w:type="default"/>'
            props.append(f'<w:sectPr>{ref}{number}<w:test>{escape(repr(vars(s.PageSetup)))}</w:test></w:sectPr>')
        return f'<w:document xmlns:w="{layout._W[1:-1]}"><w:body>' + ''.join(
            f'<w:p><w:pPr>{p}</w:pPr></w:p>' for p in props[:-1]) + props[-1] + '</w:body></w:document>'


class Range:
    def __init__(self, doc, start, end):
        self.doc, self.Start, self.End = doc, start, end

    @property
    def Paragraphs(self):
        return Collection(p for p in self.doc.Paragraphs if p.Range.Start < self.End and p.Range.End > self.Start)

    def InsertBreak(self, kind, /):
        d, pos = self.doc, self.Start
        d.calls.append((pos, kind))
        index = next(i for i, s in enumerate(d.Sections) if s.Range.Start <= pos < s.Range.End)
        source = d.Sections.items[index]
        new = copy.deepcopy(source)
        new.Range = NS(Start=pos + 1, End=source.Range.End + 1)
        for later in d.Sections.items[index + 1:]:
            later.Range.Start += 1
            later.Range.End += 1
        source.Range.End = pos + 1
        d.Sections.items.insert(index + 1, new)
        d.Content.Text = d.Content.Text[:pos] + '\f' + d.Content.Text[pos:]
        refresh(d)
        if d.fail == 'insert':
            raise RuntimeError('RPC_E_DISCONNECTED')


def refresh(d):
    paragraphs, start = [], 0
    for i, char in enumerate(d.Content.Text):
        if char in '\r\f':
            paragraphs.append(NS(Range=NS(Start=start, End=i + 1)))
            start = i + 1
    d.Paragraphs = Collection(paragraphs)


@pytest.fixture
def o(office):
    d = office.doc
    d.Content = Content(d, 'one\rtwo\rthree\rfour\r')
    d.Range = lambda a, b: Range(d, a, b)
    d.Repaginate = lambda: None
    d.calls, d.fail = [], ''
    hf = lambda: NS(LinkToPrevious=False, PageNumbers=NS(RestartNumberingAtSection=True))
    d.Sections = Collection([NS(Range=NS(Start=0, End=d.Content.End),
        PageSetup=NS(Orientation=0, PaperSize=7, PageWidth=600, PageHeight=800, TopMargin=20,
                     BottomMargin=30, LeftMargin=40, RightMargin=50, DifferentFirstPageHeaderFooter=False),
        Headers=Collection(hf() for _ in range(3)), Footers=Collection(hf() for _ in range(3)))])
    refresh(d)
    return office


def call(o, **kw):
    return o.call('word_sections', document=o.doc.Name, **kw)


def test_orientation_preserves_footer_continuity_rotates_and_undo(o):
    before = undo.word_fingerprint(o.doc, styles=False)
    result = call(o, action='orientation', start_paragraph=2, end_paragraph=3)
    assert o.doc.calls == [(14, 2), (4, 2)]
    assert result['range_sections'] == [2]
    assert [s['orientation'] for s in result['sections']] == ['portrait', 'landscape', 'portrait']
    assert [s['restart_numbering_at_section'] for s in result['sections']] == [True, False, False]
    assert all(s['linked_to_previous']['footer_primary'] for s in result['sections'][1:])
    assert [getattr(o.doc.Sections(2).PageSetup, p) for p in layout._PAGE_DIMS] == [800, 600, 40, 50, 20, 30]
    assert len(o.word.starts) == o.word.ends == 1 and result['undo'] == 'available'
    o.call('office_undo', file=o.doc.Name)
    assert undo.word_fingerprint(o.doc, styles=False) == before


@pytest.mark.parametrize('start,end,count', [(1, 4, 0), (1, 2, 1), (3, 4, 1), (2, 3, 2)])
def test_missing_boundaries_only_and_repeat_noop(o, start, end, count):
    result = call(o, action='orientation', start_paragraph=start, end_paragraph=end)
    assert len(result['inserted_breaks']) == count
    sec = next(s for s in result['sections'] if s['orientation'] == 'landscape')
    o.doc.calls.clear()
    result = call(o, action='orientation', start_paragraph=sec['start_paragraph'], end_paragraph=sec['end_paragraph'])
    assert not result['inserted_breaks'] and not o.doc.calls


@pytest.mark.parametrize('kind,code', list(layout.SECTION_BREAKS.items()))
def test_all_break_kinds(o, kind, code):
    call(o, action='insert_break', paragraph=3, kind=kind)
    assert o.doc.calls == [(8, code)] and o.doc.Sections.Count == 2


@pytest.mark.parametrize('cause', ['table_start', 'table_end', 'bad_range', 'missing', 'kind', 'orientation', 'action', 'autosave', 'protected'])
def test_refusals_before_first_mutation(o, cause):
    args = dict(action='orientation', start_paragraph=2, end_paragraph=3)
    if cause.startswith('table'):
        o.doc.Tables = Collection([NS(Range=NS(Start=3 if cause == 'table_start' else 10, End=6 if cause == 'table_start' else 17))])
    elif cause == 'bad_range':
        args['end_paragraph'] = 1
    elif cause == 'missing':
        args['end_paragraph'] = 99
    elif cause == 'autosave':
        o.doc.AutoSaveOn = True
    elif cause == 'protected':
        o.doc.ProtectionType = 3
    else:
        args[cause] = 'invalid'
    with pytest.raises(ToolError):
        call(o, **args)
    assert not o.doc.calls and not o.word.starts


def test_readonly_protected_autosave_and_no_header_access(o, monkeypatch):
    xml = o.doc.Content.WordOpenXML
    monkeypatch.setattr(Content, 'WordOpenXML', property(lambda _: xml))
    for section in o.doc.Sections:
        section.Headers = section.Footers = lambda _: pytest.fail('Reading a header creates it')
    o.doc.ProtectionType, o.doc.AutoSaveOn = 3, True
    monkeypatch.setattr(config, 'SETTINGS', replace(config.SETTINGS, mode='readonly'))
    assert call(o)['sections'][0]['starting_number'] == 7
    with pytest.raises(ToolError, match='Read-only'):
        call(o, action='insert_break', paragraph=2)
    assert not o.word.starts


def test_partial_failure_has_undo_and_shape_warning(o):
    o.doc.Shapes = Collection([NS(Anchor=NS(StoryType=1, Start=6), RelativeHorizontalPosition=1,
                                  RelativeVerticalPosition=0, Name='floating')])
    o.doc.fail = 'insert'
    with pytest.raises(ToolError, match='office_undo'):
        call(o, action='orientation', start_paragraph=2, end_paragraph=3)
    assert o.word.ends == 1 and undo.STACKS
    o.call('office_undo', file=o.doc.Name)
    o.doc.fail = ''
    assert call(o, action='orientation', start_paragraph=2, end_paragraph=3)['warnings']


def test_existing_sections_keep_restart_and_whole_table_allowed(o):
    call(o, action='insert_break', paragraph=3)
    o.doc.Sections(2).Footers(1).LinkToPrevious = False
    o.doc.Sections(2).Footers(1).PageNumbers.RestartNumberingAtSection = True
    o.doc.Tables = Collection([NS(Range=NS(Start=5, End=7))])
    o.doc.calls.clear()
    result = call(o, action='orientation', start_paragraph=1, end_paragraph=o.doc.Paragraphs.Count)
    assert not o.doc.calls and result['range_sections'] == [1, 2]
    assert all(s['restart_numbering_at_section'] for s in result['sections'])
    assert not result['sections'][1]['linked_to_previous']['footer_primary']
    result = call(o, action='orientation', start_paragraph=1, end_paragraph=o.doc.Paragraphs.Count, orientation='portrait')
    assert all(s['orientation'] == 'portrait' for s in result['sections'])
    assert o.doc.Sections(2).PageSetup.LeftMargin == 40


def test_late_page_setup_failure_retains_undo(o):
    class FailingPage(NS):
        def __setattr__(self, name, value):
            if name == 'PageWidth':
                raise RuntimeError('RPC_E_DISCONNECTED')
            super().__setattr__(name, value)
    o.doc.Sections(1).PageSetup = FailingPage(**vars(o.doc.Sections(1).PageSetup))
    with pytest.raises(ToolError, match='setting section 2 attempted.*office_undo'):
        call(o, action='orientation', start_paragraph=2, end_paragraph=3)
    assert o.doc.Sections.Count == 3 and undo.STACKS
    o.call('office_undo', file=o.doc.Name)
    assert o.doc.Sections.Count == 1


def test_xml_section_mismatch_refuses_before_undo(o, monkeypatch):
    monkeypatch.setattr(Content, 'WordOpenXML', property(lambda _: f'<w:body xmlns:w="{layout._W[1:-1]}"/>'))
    with pytest.raises(ToolError, match='counts disagree'):
        call(o, action='orientation', start_paragraph=2, end_paragraph=3)
    assert not o.doc.calls and not o.word.starts


@pytest.mark.parametrize('outside', [False, True])
def test_strict_target_and_allowed_folders_before_changes(o, monkeypatch, outside):
    settings = replace(config.SETTINGS, strict_target=True)
    if outside:
        settings = replace(settings, allowed_dirs=(o.tmp / 'allowed',))
        o.doc.Path, o.doc.FullName = str(o.tmp), str(o.tmp / o.doc.Name)
    monkeypatch.setattr(config, 'SETTINGS', settings)
    with pytest.raises(ToolError):
        o.call('word_sections', document=o.doc.Name if outside else '', action='insert_break', paragraph=2)
    assert not o.doc.calls and not o.word.starts


def test_xml_header_shape_warning_inherits_without_com():
    xml = f'''<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
        xmlns:w="{layout._W[1:-1]}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <pkg:part pkg:name="/word/document.xml"><w:body><w:p><w:pPr><w:sectPr>
        <w:headerReference w:type="first" r:id="h1"/>
      </w:sectPr></w:pPr></w:p><w:sectPr/></w:body></pkg:part>
      <pkg:part pkg:name="/word/_rels/document.xml.rels">
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="h1" Target="header1.xml"/>
        </Relationships></pkg:part>
      <pkg:part pkg:name="/word/header1.xml"><w:positionH relativeFrom="page"/></pkg:part>
    </pkg:package>'''
    meta = layout._section_xml(NS(Content=NS(WordOpenXML=xml), Sections=NS(Count=2)))
    assert all(s['_header_page_shapes'] and s['_existing'] == [('header', 2)] for s in meta)
    assert not meta[0]['linked_to_previous']['header_first_page']
    assert meta[1]['linked_to_previous']['header_first_page']
