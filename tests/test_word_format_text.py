"""Paragraph shading, borders and pagination with native undo; no Office activation."""

from types import SimpleNamespace as NS

import pytest

from office_live import undo, word_format as core
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401


class Borders:
    def __init__(self, edges):
        self.edges = edges

    def __call__(self, index, /):
        return self.edges[index]


@pytest.fixture
def formatted(office, monkeypatch):
    o = office
    paragraphs = []
    for i in range(8):
        edges = {s: NS(LineStyle=i % 2, LineWidth=6, Color=123 + i) for s in (-1, -2, -3, -4)}
        pf = NS(Shading=NS(Texture=10, ForegroundPatternColor=42, BackgroundPatternColor=i),
                Borders=Borders(edges), KeepTogether=False)
        rng = NS(Start=i * 10, End=(i + 1) * 10, StoryType=1, Font=NS(), ParagraphFormat=pf)
        paragraph = NS(Range=rng)
        rng.Paragraphs = Collection([paragraph])
        paragraphs.append(paragraph)
    o.doc.Paragraphs = Collection(paragraphs)

    def target(first=1, last=8):
        return NS(Font=NS(), ParagraphFormat=NS(), Paragraphs=Collection(o.doc.Paragraphs.items[first-1:last]))

    monkeypatch.setattr(core, 'paragraphs_range', lambda doc, first, last: target(first, last))
    monkeypatch.setattr(core, 'find_all', lambda *a, **kw: [target(5, 6), target(6, 7)])
    o.doc.Content.Font, o.doc.Content.ParagraphFormat = NS(), NS()
    o.doc.Content.Paragraphs = o.doc.Paragraphs
    o.doc.ActiveWindow.Selection = NS(Range=target(5, 7))
    # Fingerprinting observes formatting, rather than only text, as real WordOpenXML does.
    monkeypatch.setattr(undo, '_word_body_xml', lambda doc: repr([
        (vars(p.Range.ParagraphFormat.Shading), p.Range.ParagraphFormat.KeepTogether,
         [vars(p.Range.ParagraphFormat.Borders(s)) for s in (-1, -2, -3, -4)]) for p in doc.Paragraphs]))
    return o


def apply(o, **kw):
    return o.call('word_format_text', document=o.doc.Name, start_paragraph=5, end_paragraph=7, **kw)


@pytest.mark.parametrize('scope', ['paragraphs', 'find', 'selection', 'document'])
def test_format_text_paragraph_acceptance_and_undo(formatted, scope):
    o = formatted
    before = undo._word_body_xml(o.doc)
    result = apply(o, scope=scope, find_text='row', paragraph_shading='#F2F2F2',
                   paragraph_border='left', border_color='#123456', border_width_pt=0.9, keep_together=True)
    assert result['undo'] == 'available' and result['border_width_pt'] == 1
    expected = range(8) if scope == 'document' else range(4, 7)
    for i in expected:
        pf = o.doc.Paragraphs(i + 1).Range.ParagraphFormat
        assert pf.Shading.BackgroundPatternColor == 0xF2F2F2 and pf.Shading.Texture == 0
        assert pf.KeepTogether
        assert [pf.Borders(s).LineStyle for s in (-1, -2, -3, -4)] == [0, 1, 0, 0]
        assert pf.Borders(-2).Color == 0x563412 and pf.Borders(-2).LineWidth == 8
    o.call('office_undo', file=o.doc.Name)
    assert undo._word_body_xml(o.doc) == before
    assert o.doc.undo_calls == [1]


def test_format_text_refuses_later_user_edit(formatted):
    o = formatted
    apply(o, paragraph_shading='red')
    o.doc.Paragraphs(5).Range.ParagraphFormat.KeepTogether = True
    with pytest.raises(ToolError, match='changed after the agent'):
        o.call('office_undo', file=o.doc.Name)
    assert not o.doc.undo_calls


def test_format_text_busy_undo_record_refuses_before_changes(formatted):
    o = formatted
    before = undo._word_body_xml(o.doc)
    o.word.UndoRecord.IsRecordingCustomRecord = True
    with pytest.raises(ToolError, match='could not prepare office_undo'):
        apply(o, bold=True, paragraph_border='left')
    assert undo._word_body_xml(o.doc) == before


@pytest.mark.parametrize('mode,sides', [('box', [1, 1, 1, 1]), ('top', [1, 0, 0, 0]),
                                       ('bottom', [0, 0, 1, 0]), ('none', [0, 0, 0, 0])])
def test_format_text_border_modes_and_clear(formatted, mode, sides):
    apply(formatted, paragraph_shading='none', paragraph_border=mode, keep_together=False)
    pf = formatted.doc.Paragraphs(5).Range.ParagraphFormat
    assert pf.Shading.BackgroundPatternColor == -16777216 and pf.Shading.Texture == 0
    assert [pf.Borders(s).LineStyle for s in (-1, -2, -3, -4)] == sides
    assert pf.KeepTogether is False


@pytest.mark.parametrize('width,actual', [(0.01, .25), (.625, .5), (1.8, 1.5), (2.2, 2.25), (99, 6)])
def test_format_text_width_rounding(formatted, width, actual):
    assert apply(formatted, paragraph_border='box', border_width_pt=width)['border_width_pt'] == actual


@pytest.mark.parametrize('kw', [
    {'paragraph_shading': 'bad-color'}, {'paragraph_border': 'right'},
    {'paragraph_border': 'box', 'border_color': 'bad-color'},
    *({'paragraph_border': 'left', 'border_width_pt': v} for v in (0, -1, float('inf'), float('nan'), True)),
    {'border_color': 'red'}, {'border_width_pt': 1},
    {'paragraph_border': 'none', 'border_width_pt': 1}, {'keep_together': 1},
    {'paragraph_shading': '#F2F2F2', 'font_color': 'invalid'},
])
def test_format_text_validation_precedes_changes(formatted, kw):
    before = undo._word_body_xml(formatted.doc)
    with pytest.raises(ToolError):
        apply(formatted, bold=True, **kw)
    assert undo._word_body_xml(formatted.doc) == before
    assert not undo.STACKS
    assert not vars(formatted.doc.Paragraphs(5).Range.Font)


def test_format_text_partial_failure_keeps_undo(formatted):
    o = formatted
    before = undo._word_body_xml(o.doc)

    class Refuse(NS):
        def __setattr__(self, name, value):
            if name == 'BackgroundPatternColor':
                raise RuntimeError('simulated write failure')
            super().__setattr__(name, value)

    shading = o.doc.Paragraphs(6).Range.ParagraphFormat.Shading
    o.doc.Paragraphs(6).Range.ParagraphFormat.Shading = Refuse(**vars(shading))
    with pytest.raises(ToolError, match='formatting interrupted'):
        apply(o, paragraph_shading='red')
    assert o.doc.Paragraphs(5).Range.ParagraphFormat.Shading.BackgroundPatternColor == 255
    o.call('office_undo', file=o.doc.Name)
    assert undo._word_body_xml(o.doc) == before
