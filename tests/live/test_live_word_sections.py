"""Written-only section probes on marked temporary documents; owner approval required to run."""

from pathlib import Path

import pytest

from tests.live.mcpclient import ToolFailed
from tests.live.office_cleanup import DialogWatch
from tests.live.test_live_office_operations import srv as isolated_server
from tests.live.test_live_stage1 import in_word, word_state

srv = isolated_server
pytestmark = pytest.mark.live


@pytest.fixture
def sections_doc(srv, doc, tmp):
    path = str(Path(tmp) / 'sections.docx')
    srv.call('word_save_as', document=doc, path=path)
    def setup(d):
        d.Content.Text = '\r'.join(f'Paragraph {i}' for i in range(1, 41))
        for i in range(5, 41, 4):
            d.Paragraphs(i).Range.ParagraphFormat.PageBreakBefore = True
        d.Sections(1).PageSetup.Orientation = 0  # wdOrientPortrait
        ps = d.Sections(1).PageSetup
        ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin = 20, 30, 40, 50
        d.Repaginate()
        assert int(d.ComputeStatistics(2)) == 10  # wdStatisticPages
    in_word(path, setup)
    srv.call('word_headers_footers', document=path, action='set', text='Section probe', page_numbers='page')
    def numbering(d):
        numbers = d.Sections(1).Footers(1).PageNumbers
        numbers.RestartNumberingAtSection, numbers.StartingNumber = True, 1
    in_word(path, numbering)
    return path


def test_sections_landscape_numbering_footer_and_undo(srv, sections_doc):
    path = sections_doc
    before = word_state(path)[0]
    with DialogWatch() as watch:
        result = srv.call('word_sections', document=path, action='orientation', start_paragraph=20, end_paragraph=25)
        assert result['undo'] == 'available' and len(result['inserted_breaks']) == 2
        assert [s['orientation'] for s in result['sections']] == ['portrait', 'landscape', 'portrait']
        assert result['range_sections'] == [2]
        assert all(s['linked_to_previous']['footer_primary'] and not s['restart_numbering_at_section'] for s in result['sections'][1:])
        def verify(d):
            d.Repaginate()
            texts = []
            for s in d.Sections:
                r = s.Footers(1).Range  # существующий footer, созданный fixture
                texts.append(str(r.Text).replace(str(r.Fields(1).Result.Text), '{PAGE}', 1))
            assert texts[0] == texts[1] == texts[2]
            ps = d.Sections(2).PageSetup
            assert [float(getattr(ps, n)) for n in ('TopMargin', 'BottomMargin', 'LeftMargin', 'RightMargin')] == [40, 50, 20, 30]
            assert float(ps.PageWidth) == float(d.Sections(1).PageSetup.PageHeight)
            count = int(d.ComputeStatistics(2))  # wdStatisticPages
            assert [int(d.GoTo(1, 1, n).Information(1)) for n in range(1, count + 1)] == list(range(1, count + 1))  # wdGoToPage, wdGoToAbsolute, wdActiveEndAdjustedPageNumber
        in_word(path, verify)
        middle = result['sections'][1]
        repeated = srv.call('word_sections', document=path, action='orientation',
                            start_paragraph=middle['start_paragraph'], end_paragraph=middle['end_paragraph'])
        assert not repeated['inserted_breaks']
        restored = srv.call('office_undo', file=path)
        assert 1 <= restored['native_steps'] <= 3  # один основной шаг и до двух дополнительных
        print(f"Section breaks + PageSetup: {restored['native_steps']} native undo steps")
        assert word_state(path)[0] == before  # styles=False; native undo may collect extra steps
    assert not watch.seen


def test_sections_list_no_headers_and_table_refusal(srv, doc, tmp):
    path = str(Path(tmp) / 'sections-refusal.docx')
    srv.call('word_create_table', document=doc, data=[['one'], ['two']])
    srv.call('word_save_as', document=doc, path=path)
    before = word_state(path)
    srv.call('word_sections', document=path)
    assert word_state(path) == before  # no empty headers created, Saved unchanged
    from office_live.wd_common import paragraph_index_at
    start = in_word(path, lambda d: paragraph_index_at(d, int(d.Tables(1).Range.Start)))
    with pytest.raises(ToolFailed, match='table'):
        srv.call('word_sections', document=path, action='orientation', start_paragraph=start, end_paragraph=start)
    assert word_state(path) == before


@pytest.mark.parametrize('kind', ['next_page', 'continuous', 'even_page', 'odd_page'])
def test_section_break_kind_undo(srv, sections_doc, kind):
    path, before = sections_doc, word_state(sections_doc)[0]
    srv.call('word_sections', document=path, action='insert_break', paragraph=3, kind=kind)
    assert len(srv.call('word_sections', document=path)['sections']) == 2
    srv.call('office_undo', file=path)
    assert word_state(path)[0] == before
