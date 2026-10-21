"""Caption contracts and refusals on fake Word; no Office activation."""

import bisect
from dataclasses import replace
from types import SimpleNamespace as NS

import pytest

from office_live import config, registry, undo, word_structure as c
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401
from tests.office_operations_fakes import WordRange, refresh_word, word


class CaptionRange(WordRange):
    def __deepcopy__(self, memo):
        return CaptionRange(self.doc, self.Start, self.End)

    @property
    def Paragraphs(self):
        return Collection(p for p in self.doc.Paragraphs if p.Range.Start <= self.End and p.Range.End > self.Start)

    @property
    def InlineShapes(self):
        return self.doc.InlineShapes

    @property
    def Fields(self):
        def update():
            self.doc.calls.append(("update", self.Start, self.End))
            return int(self.doc.fail == "update")
        return NS(Update=update)

    def insert(self, args, start):
        self.doc.calls.append(args)
        old = self.doc.Content.Text
        self.doc.Content.Text = old[:start] + "new\r" + old[start:]
        refresh_word(self.doc)
        if self.doc.fail == "insert":
            raise RuntimeError("RPC_E_DISCONNECTED")

    def InsertCaption(self, label, title, auto, position, exclude, /):
        bounds = self.Paragraphs(1).Range
        self.insert(("caption", label, title, auto, position, exclude), bounds.Start if position == 0 else bounds.End)

    def InsertCrossReference(self, typ, kind, item, hyperlink, include, separate, separator, /):
        self.insert(("reference", typ, kind, item, hyperlink, include, separate, separator), self.Start)


@pytest.fixture
def o(office):
    d = word(office)
    d.Content.Text = "See here\rTable 1 Title\rDATA\rend\r"
    refresh_word(d)
    d.Range = lambda a, b: CaptionRange(d, a, b)
    d.Paragraphs = Collection(NS(Range=d.Range(p.Range.Start, p.Range.End)) for p in d.Paragraphs)
    d.Tables = Collection([NS(Range=d.Range(23, 28), NestingLevel=1)])
    d.Content.Fields = Collection()
    d.Bookmarks = NS(Exists=lambda n: False, Add=lambda n, r: d.calls.append(("bookmark", n, r.Start, r.End)))
    d.GetCrossReferenceItems = lambda typ: ["Table 1 Title"] if isinstance(typ, str) else ["anchor", "Heading two"]
    office.word.CaptionLabels = lambda i: NS(Name={-2: "Table", -1: "Figure", -3: "Equation"}[i])
    return office


def seq(o, label="Table", index=2):
    r = o.doc.Paragraphs(index).Range
    field = NS(Type=12, Code=NS(Text=f' SEQ "{label}" \\* ARABIC ', Start=r.Start + 7),
               Result=NS(Text="1", Start=r.Start + 8, End=r.Start + 9, Paragraphs=Collection([NS(Range=r)])))
    o.doc.Content.Fields.items.append(field)
    return field


def call(o, **kw):
    return o.call("word_captions", document=o.doc.Name, **kw)


def test_list_local_labels_orphans_missing_and_no_undo(o):
    o.word.CaptionLabels = lambda i: NS(Name={-2: "Таблица", -1: "Рисунок", -3: "Формула"}[i])
    seq(o, "Таблица")
    seq(o, "Custom", 1)
    result = call(o)
    cap = result["captions"][0]
    assert (cap["label_id"], cap["label"], cap["number"], cap["caption"]) == (-2, "Таблица", "1", 1)
    assert cap["objects"] == [{"target": "table", "index": 1}] and cap["link"].startswith("officelive://")
    assert len(result["orphans"]) == 1 and result["orphans"][0]["label_id"] is None
    assert result["tables_without_caption"] == [] and not o.word.starts and not o.doc.calls
    o.doc.Content.Fields.items.clear()
    assert call(o)["tables_without_caption"] == [1]


@pytest.mark.parametrize("local", ["Table", "Таблица"])
def test_add_native_arguments_local_update_and_one_undo(o, local):
    o.word.CaptionLabels = lambda i: NS(Name=local if i == -2 else str(i))
    before = o.doc.Content.Text
    result = call(o, action="add", target="table", table_index=1, text=" — Title")
    assert result["undo"] == "available" and result["label_id"] == -2
    assert result["label"] == local
    # TitleAutoText "" not Missing: Missing mid-list nulls Position in Word (review finding)
    assert o.doc.calls[0] == ("caption", -2, " — Title", "", 0, False)
    assert o.doc.calls[1] == ("update", 23, 27) and len(o.word.starts) == o.word.ends == 1
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == before


@pytest.mark.parametrize("target,label,ident,start", [("picture", "", -1, 28), ("paragraph", "equation", -3, 9)])
def test_picture_and_equation_default_below(o, target, label, ident, start):
    o.doc.InlineShapes.items.append(NS(Type=3, Range=o.doc.Range(23, 24)))
    assert call(o)["pictures_without_caption"] == [1]
    call(o, action="add", target=target, picture_index=1, paragraph=1, label=label)
    assert o.doc.calls[0][1] == ident and o.doc.calls[0][4] == 1
    assert o.doc.calls[1] == ("update", start, start + 4)


@pytest.mark.parametrize("cause", ["duplicate", "nested", "floating", "bad_label", "bad_position", "multiline", "protected", "autosave", "missing", "paragraph_label"])
def test_refusals_before_undo_or_mutation(o, cause):
    args = dict(action="add", target="table", table_index=1)
    if cause == "duplicate":
        seq(o)
    elif cause == "nested":
        o.doc.Tables(1).NestingLevel = 2
    elif cause == "floating":
        o.doc.Shapes.items.append(NS())
        args.update(target="picture", picture_index=1)
    elif cause == "protected":
        o.doc.ProtectionType = 3
    elif cause == "autosave":
        o.doc.AutoSaveOn = True
    else:
        args.update({"bad_label": {"label": "Table"}, "bad_position": {"position": "left"}, "multiline": {"text": "a\nb"},
                     "missing": {"table_index": 9}, "paragraph_label": {"target": "paragraph", "paragraph": 1}}[cause])
    before = o.doc.Content.Text
    with pytest.raises(ToolError):
        call(o, **args)
    assert before == o.doc.Content.Text and not o.doc.calls and not o.word.starts


@pytest.mark.parametrize("reference,kind", [("label_and_number", 3), ("text", 2), ("page", 7), ("number", -1)])
def test_reference_matching_and_number_bookmark(o, reference, kind):
    seq(o)
    call(o, action="cross_reference", caption=1, paragraph=4, reference=reference)
    native = next(a for a in o.doc.calls if a[0] == "reference")
    assert native[1:3] == ((2 if reference == "number" else "Table"), kind)
    assert native[4:7] == (True, False, False)
    assert (native[3].startswith("_OfficeLiveCaption") if reference == "number" else native[3] == 1)
    assert len(o.word.starts) == o.word.ends == 1


@pytest.mark.parametrize("target,item", [("heading", 2), ("bookmark", "anchor")])
def test_heading_bookmark_and_literal_occurrence(o, monkeypatch, target, item):
    monkeypatch.setattr(c.wd, "find_all", lambda *a, **kw: [o.doc.Range(0, 1), o.doc.Range(2, 3)])
    result = call(o, action="cross_reference", caption=target, heading=2, bookmark="anchor", at="after_text", find_text="x", occurrence=2, hyperlink=False)
    assert result["label_id"] is None and result["label"] is None
    assert o.doc.calls[0][1:5] == (1 if target == "heading" else 2, -1, item, False)
    assert o.doc.calls[1] == ("update", 3, 7)


@pytest.mark.parametrize("cause", ["ambiguous", "missing", "inside_field", "bad_at", "bad_reference"])
def test_reference_refusals(o, cause):
    seq(o)
    args = dict(action="cross_reference", caption=1, paragraph=4)
    if cause in ("ambiguous", "missing"):
        o.doc.GetCrossReferenceItems = lambda _: ["Table 1 Title"] * (2 if cause == "ambiguous" else 0)
    elif cause == "inside_field":
        o.doc.Content.Fields.items[0].Result.End = 40
    else:
        args.update({"bad_at": {"at": "selection"}, "bad_reference": {"reference": "invalid"}}[cause])
    with pytest.raises(ToolError):
        call(o, **args)
    assert not o.doc.calls and not o.word.starts


@pytest.mark.parametrize("failure", ["insert", "update"])
def test_partial_failure_retains_undo(o, failure):
    o.doc.fail = failure
    with pytest.raises(ToolError, match="office_undo"):
        call(o, action="add", target="table", table_index=1)
    assert o.word.ends == 1 and undo.STACKS


def test_readonly_list_on_protected_autosave_document(o, monkeypatch):
    o.doc.ProtectionType, o.doc.AutoSaveOn = 3, True
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    assert registry.CATALOG["word_captions"].read_actions == ("list",)
    assert call(o)["tables_without_caption"] == [1]
    with pytest.raises(ToolError, match="Read-only"):
        call(o, action="add", target="table", table_index=1)
    assert not o.doc.calls and not o.word.starts


class LongDocument:
    """Paragraphs of 10 characters; Range.Paragraphs counts like Word, which walks every paragraph it counts."""

    def __init__(self, paragraphs):
        self.starts, self.walked = [10 * i for i in range(paragraphs)], 0
        self.Content = NS(End=10 * paragraphs, Fields=Collection())
        self.Tables = self.InlineShapes = Collection()

    def Range(self, start, end):
        lo, hi = bisect.bisect_right(self.starts, start) - 1, bisect.bisect_left(self.starts, end)
        self.walked += hi - lo
        return NS(Start=start, End=end, Paragraphs=NS(Count=hi - lo))


def test_list_walks_the_document_once(monkeypatch):
    doc = LongDocument(5000)
    for n, p in enumerate(range(7, 5000, 16), 1):  # 300+ captions spread over the document, listed in field order
        r = NS(Start=10 * p, End=10 * p + 10, Text=f"Table {n}\r")
        doc.Content.Fields.items.append(NS(Type=12, Code=NS(Text=" SEQ Table "), Result=NS(
            Text=str(n), Paragraphs=Collection([NS(Range=r)]))))
    monkeypatch.setattr(c.wd, "range_link", lambda doc, paragraph, rng: "")
    app = NS(CaptionLabels=lambda i: NS(Name={-2: "Table", -1: "Figure", -3: "Equation"}[i]))
    captions, _ = c._inventory(app, doc)
    assert len(captions) >= 300 and [x["paragraph"] for x in captions] == list(range(8, 5001, 16))
    assert doc.walked <= 5000 + len(captions)  # per-caption counting from the start walked ~750 000 paragraphs


@pytest.mark.parametrize("positions", [[0], [9, 10, 11], [49999, 120, 0, 120, 60000], list(range(0, 50000, 777))])
def test_paragraph_indexes_match_paragraph_index_at(positions):
    doc = LongDocument(5000)
    assert c.wd.paragraph_indexes(doc, positions) == {p: c.wd.paragraph_index_at(doc, p) for p in positions}
