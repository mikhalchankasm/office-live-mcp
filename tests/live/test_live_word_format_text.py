"""Issue #60 acceptance on a marked temporary document; run only with owner approval."""

from pathlib import Path

import pytest

from tests.live.office_cleanup import DialogWatch
from tests.live.test_live_office_operations import srv as isolated_server
from tests.live.test_live_stage1 import in_word

srv = isolated_server
pytestmark = pytest.mark.live


def paragraph_state(doc):
    out = []
    for i in range(1, int(doc.Paragraphs.Count) + 1):
        pf = doc.Paragraphs(i).Range.ParagraphFormat
        out.append((int(pf.Shading.BackgroundPatternColor), int(pf.Shading.Texture), bool(pf.KeepTogether),
                    [(int(pf.Borders(s).LineStyle), int(pf.Borders(s).LineWidth), int(pf.Borders(s).Color))
                     for s in (-1, -2, -3, -4)]))  # wdBorderTop/Left/Bottom/Right
    return out


def test_probe_format_text_paragraph_shading_left_bar_and_undo(srv, doc, tmp):
    path = str(Path(tmp) / "paragraph-format.docx")
    srv.call("word_save_as", document=doc, path=path)
    srv.call("word_insert_text", document=path, text="\n".join(f"Paragraph {i}" for i in range(1, 10)))
    srv.call("word_format_text", document=path, scope="document", paragraph_shading="none",
             paragraph_border="none", keep_together=False)
    before = in_word(path, paragraph_state)
    with DialogWatch() as watch:
        result = srv.call("word_format_text", document=path, scope="paragraphs", start_paragraph=5,
                          end_paragraph=7, paragraph_shading="#F2F2F2", paragraph_border="left", keep_together=True)
        assert result["undo"] == "available" and result["border_width_pt"] == 0.5
        after = in_word(path, paragraph_state)
        for i, props in enumerate(after):
            if 4 <= i <= 6:
                assert props[:3] == (0xF2F2F2, 0, True)  # wdTextureNone
                assert [b[0] for b in props[3]] == [0, 1, 0, 0]  # wdLineStyleNone / wdLineStyleSingle
                assert props[3][1][1:] == (4, 0)  # wdLineWidth050pt, black
            else:
                assert props == before[i]
        srv.call("office_undo", file=path)
        assert in_word(path, paragraph_state) == before
    assert not watch.seen
