"""Sequential Word replacement batches and refusals without Office (#58)."""

import re

import pytest

from office_live import config, undo, word_core
from office_live.errors import ToolError
from tests.history_fakes import fake_office  # noqa: F401


class Hit:
    def __init__(self, content, start, end):
        self.content, self.start, self.end = content, start, end

    @property
    def Text(self):
        return self.content.Text[self.start:self.end]

    @Text.setter
    def Text(self, value):
        self.content.Text = self.content.Text[:self.start] + value + self.content.Text[self.end:]


@pytest.fixture
def replacement(office, monkeypatch):
    def find(content, text, match_case=False, whole_word=False, wildcards=False, limit=100000):
        pattern = re.escape(text.replace("^^", "^"))
        if whole_word:
            pattern = r"\b" + pattern + r"\b"
        hits = re.finditer(pattern, content.Text, 0 if match_case else re.IGNORECASE)
        return [Hit(content, m.start(), m.end()) for m in hits][:limit]

    monkeypatch.setattr(word_core, "find_all", find)
    return office


def call(o, **kw):
    return o.call("word_replace_text", document=o.doc.Name, **kw)


def test_replace_25_pairs_one_record_journal_and_undo(replacement):
    o = replacement
    original = "\r".join(f"В{i}. Вопрос" for i in range(2, 27))
    o.doc.Content.Text = original
    pairs = [{"find": f"В{i}", "replace": f"В{i-1}"} for i in range(2, 27)]
    result = call(o, pairs=pairs, whole_word=True)
    assert o.doc.Content.Text == "\r".join(f"В{i}. Вопрос" for i in range(1, 26))
    assert result["pairs"] == [{**p, "count": 1} for p in pairs]
    assert result["replacements"] == 25 and result["undo"] == "available"
    assert len(o.word.starts) == o.word.ends == 1
    assert len(undo.STACKS[undo.stack_key("document", o.word, o.doc)]) == 1
    lines = next(config.SETTINGS.journal_dir.glob("*.md")).read_text(encoding="utf-8")
    assert lines.count("· word_replace_text ·") == 1
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == original and o.doc.undo_calls == [1]


def test_replace_sequential_cascade_deletion_and_single_response(replacement):
    o = replacement
    o.doc.Content.Text = "a a"
    result = call(o, pairs=[{"find": "a", "replace": "b"}, {"find": "b", "replace": ""}])
    assert o.doc.Content.Text == " " and result["replacements"] == 4
    result = call(o, find="missing", replace="x")
    assert result["replacements"] == 0 and "pairs" not in result


@pytest.mark.parametrize("kw", [
    {"pairs": []}, {"pairs": [{"find": "a", "replace": "b"}] * 201},
    {"pairs": {}}, {"pairs": [None]}, {"pairs": [{"find": "a"}]},
    {"pairs": [{"find": "", "replace": "b"}]},
    {"pairs": [{"find": "x" * 256, "replace": "b"}]},
    {"pairs": [{"find": "a", "replace": 3}]},
    {"pairs": [{"find": "a", "replace": "b"}], "find": ""},
    {"pairs": [{"find": "a", "replace": "b"}], "replace": ""},
    {"pairs": [{"find": "a", "replace": "b"}, {"find": "", "replace": "x"}]},
    {"pairs": [{"find": "a", "replace": "b"}], "scope": "bad"},
    {"pairs": [{"find": "a", "replace": "x" * 256}], "wildcards": True},
    {"find": "a"}, {"replace": "b"}, {},
])
def test_replace_invalid_arguments_refuse_before_changes(replacement, kw):
    o = replacement
    with pytest.raises(ToolError):
        call(o, **kw)
    assert o.doc.Content.Text == "original" and not o.word.starts and not undo.STACKS


def test_replace_200_pairs_allowed(replacement):
    result = call(replacement, pairs=[{"find": "absent", "replace": ""}] * 200)
    assert len(result["pairs"]) == 200 and result["replacements"] == 0


def test_replace_all_wildcards_preflight_before_undo(replacement, monkeypatch):
    o = replacement
    seen = []
    monkeypatch.setattr(word_core.wd, "list_separator", lambda app: ";")

    def find(story, text, *flags, limit):
        assert not o.word.starts and limit == 1
        seen.append(text)
        if text == "[":
            raise ToolError("Invalid wildcard pattern")
        return []

    monkeypatch.setattr(word_core, "find_all", find)
    with pytest.raises(ToolError, match="Invalid wildcard"):
        call(o, pairs=[{"find": "a{1,2}", "replace": "b"}, {"find": "[", "replace": "x"}], wildcards=True)
    assert seen == ["a{1;2}", "["] and not o.word.starts
    assert o.doc.Content.Text == "original"


@pytest.mark.parametrize("failed_find", ["a", "b"])
def test_replace_partial_failure_keeps_one_undo(replacement, monkeypatch, failed_find):
    o = replacement
    real = word_core.replace_in_doc

    def fail(doc, find, *args):
        count = real(doc, find, *args)
        if find == failed_find:
            raise ValueError("failure after write")
        return count

    monkeypatch.setattr(word_core, "replace_in_doc", fail)
    o.doc.Content.Text = "a"
    with pytest.raises(ToolError, match="interrupted.*office_undo will restore"):
        call(o, pairs=[{"find": "a", "replace": "b"}, {"find": "b", "replace": "c"}])
    assert o.doc.Content.Text == ("b" if failed_find == "a" else "c")
    assert len(o.word.starts) == o.word.ends == 1
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == "a"


def test_replace_undo_unavailable_refuses_before_changes(replacement):
    o = replacement
    o.word.UndoRecord.IsRecordingCustomRecord = True
    with pytest.raises(ToolError, match="could not prepare office_undo"):
        call(o, pairs=[{"find": "original", "replace": "changed"}])
    assert o.doc.Content.Text == "original"


@pytest.mark.parametrize("pattern", [r"\(a\)", "[(]a[)]", "(a)"])
def test_replace_missing_wildcard_group_refuses_before_changes(replacement, pattern):
    o = replacement
    with pytest.raises(ToolError, match="missing wildcard group"):
        call(o, pairs=[{"find": "original", "replace": "b"}, {"find": pattern, "replace": r"\2"}], wildcards=True)
    assert o.doc.Content.Text == "original" and not o.word.starts


@pytest.mark.parametrize("scope", ["main", "all", "selection"])
def test_replace_shared_flags_scopes_and_wildcard_groups(replacement, monkeypatch, scope):
    o = replacement
    calls = []
    monkeypatch.setattr(word_core, "story_ranges", lambda doc, s: [doc.Content])
    monkeypatch.setattr(word_core.wd, "list_separator", lambda app: ";")
    monkeypatch.setattr(word_core, "find_all", lambda *a, **kw: [])

    def apply(doc, find, replace, *flags):
        calls.append((find, replace, flags))
        doc.Content.Text += "!"
        return 2

    monkeypatch.setattr(word_core, "replace_in_doc", apply)
    pairs = [{"find": "(a{1,2})", "replace": r"\1"}, {"find": "[(]", "replace": r"\\1"}]
    result = call(o, pairs=pairs, match_case=True, whole_word=True, wildcards=True, special_codes=True, scope=scope)
    assert calls == [(p["find"], p["replace"], (True, True, True, True, scope)) for p in pairs]
    assert result["replacements"] == 4 and result["scope"] == scope


def test_replace_single_literal_preserves_response(replacement):
    o = replacement
    o.doc.Content.Text = "A a"
    o.doc.TrackRevisions = True
    result = call(o, find="A", replace="^text", match_case=True)
    assert o.doc.Content.Text == "^text a"
    assert result["replacements"] == 1 and result["track_changes_on"] is True
    assert "pairs" not in result and result["scope"] == "main"


def test_replace_document_remains_required(replacement):
    o = replacement
    with pytest.raises(ToolError, match="document.*required"):
        o.call("word_replace_text", pairs=[{"find": "original", "replace": "changed"}])
    assert o.doc.Content.Text == "original" and not o.word.starts
