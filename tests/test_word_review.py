"""Fragment edits and comment threads: preflight, readonly, restoration, journaling and undo on fakes."""

from dataclasses import replace
from types import SimpleNamespace as NS

import pytest

from office_live import config, journal, links, navigation, registry, undo, word_edits, word_layout
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401
from tests.word_review_fakes import Comment, edits, refresh


@pytest.fixture
def o(office):
    edits(office)
    return office


def call(o, name="word_edit_preview", **kw):
    return o.call(name, **{"document": o.doc.Name, "paragraph": 1, "expected_text": "target", "new_text": "replacement", **kw})


def history(o):
    return undo.STACKS.get(undo.stack_key("document", o.word, o.doc), [])


def test_preview_readonly_and_no_side_effects(o, monkeypatch):
    before = o.doc.Content.Text
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, audit_log=o.tmp / "audit.jsonl"))
    monkeypatch.setattr(registry, "_audit_write", lambda *a, **k: pytest.fail("preview audit"))
    monkeypatch.setattr(journal, "append", lambda *a, **k: pytest.fail("preview journal"))
    monkeypatch.setattr(undo, "require_undo", lambda *a, **k: pytest.fail("preview undo"))
    result = call(o)
    assert result["original"] == "target" and result["location"]["start"] == 7
    assert result["context_before"] == "before " and result["context_after"] == " after"
    assert result["formatting"] == "uniform" and not result["blocked"]
    assert not history(o) and not o.word.starts and not list(o.tmp.glob("journal/*"))
    assert o.doc.Content.Text == before and not o.doc.calls
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    registered = {}
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda fn, name, **kw: registered.__setitem__(name, (fn, kw))))
    registry.office_tool("word_core", "read")(word_edits.word_edit_preview)
    registry.office_tool("word_core", "write")(word_edits.word_edit_apply)
    assert "word_edit_apply" not in registered
    assert registered["word_edit_preview"][1]["annotations"].read_only_hint


@pytest.mark.parametrize("shift", ["missing", "paragraph", "occurrence", "fingerprint"])
def test_stale_before_undo_and_write(o, shift):
    result = call(o)
    if shift == "missing":
        o.doc.Content.Text = "before gone after\rnext\r"
    elif shift == "paragraph":
        o.doc.Content.Text = "changed target after\rnext\r"
    elif shift == "occurrence":
        o.doc.Content.Text = "target before target after\rnext\r"
    refresh(o.doc)
    before = o.doc.Content.Text
    with pytest.raises(ToolError, match="изменился после предпросмотра"):
        call(o, "word_edit_apply", fingerprint="bad" if shift == "fingerprint" else result["fingerprint"])
    assert o.doc.Content.Text == before and not history(o) and not o.word.starts


@pytest.mark.parametrize("member", ["Fields", "Footnotes", "Endnotes", "Hyperlinks", "ContentControls", "InlineShapes", "Revisions"])
def test_structural_collections_block(o, member):
    o.doc.range_objects[member] = [NS(Code=NS(Start=8, StoryType=1), Result=NS(End=12))] if member == "Fields" else [NS()]
    refresh(o.doc)
    assert call(o)["blocked"]
    with pytest.raises(ToolError, match=member):
        call(o, "word_edit_apply")
    assert not o.word.starts and not o.doc.calls


@pytest.mark.parametrize("structure", ["field", "bookmark", "bookmark_partial", "hyperlink", "control", "comment_partial", "story", "protected", "autosave"])
def test_partial_structures_block(o, structure):
    if structure == "field":
        field = NS(Code=NS(Start=5, StoryType=1), Result=NS(End=15))
        o.doc.Fields.items.append(field)
    elif structure.startswith("bookmark"):
        o.doc.Bookmarks.items.append(NS(Range=o.doc.Range(8, 10 if structure == "bookmark" else 15)))
    elif structure in {"hyperlink", "control"}:
        getattr(o.doc, "Hyperlinks" if structure == "hyperlink" else "ContentControls").items.append(NS(Range=o.doc.Range(1, 16)))
    elif structure == "comment_partial":
        o.doc.Comments.items.append(Comment(o.doc, "comment", 4, 9))
    elif structure == "story":
        o.doc.story = 2
    elif structure == "protected":
        o.doc.ProtectionType = 3
    else:
        o.doc.AutoSaveOn = True
    refresh(o.doc)
    assert call(o)["blocked"]
    with pytest.raises(ToolError):
        call(o, "word_edit_apply")
    assert not o.word.starts and not o.doc.calls


@pytest.mark.parametrize("text", ["", "\r", "\x07", "\x0c", "\x0e", "\x01"])
def test_fragment_markers_block(o, text):
    o.doc.Content.Text = "before " + text + " after\rnext\r"
    refresh(o.doc)
    # Paragraph markers may split the paragraph; a missing target must also refuse without changes.
    try:
        result = call(o, expected_text=text)
    except ToolError:
        pass
    else:
        assert result["blocked"]
    assert not o.doc.calls


@pytest.mark.parametrize("text", ["bad\r", "bad\n", "\x07", "\x0c", "\x0e", "\x01", "\x00", "\x1f", "a" * 10001])
def test_new_text_limits(o, text):
    assert call(o, new_text=text)["blocked"]
    with pytest.raises(ToolError):
        call(o, "word_edit_apply", new_text=text)
    assert not o.doc.calls


@pytest.mark.parametrize("track", [False, True])
@pytest.mark.parametrize("fail", [False, True])
def test_apply_track_restore_error_undo_and_journal(o, track, fail):
    before = o.doc.Content.Text
    o.doc.TrackRevisions = track
    o.doc.fail = "edit" if fail else ""
    preview = call(o)
    if fail:
        with pytest.raises(ToolError, match="office_undo will restore"):
            call(o, "word_edit_apply", fingerprint=preview["fingerprint"])
    else:
        result = call(o, "word_edit_apply", fingerprint=preview["fingerprint"])
        assert result["undo"] == "available" and result["revisions_added"] == 1
        assert not result["warnings"]
    assert o.doc.TrackRevisions == track
    assert o.doc.calls[-1][-1] is True and len(o.word.starts) == o.word.ends == 1
    log = journal.journal_path(o.doc.Name).read_text(encoding="utf-8")
    assert "word_edit_apply" in log and "target" in log and "replacement" in log
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == before


def test_mixed_plain_table_unicode_and_deletion(o):
    o.doc.Content.Text = "😀 before target target after\rnext\r"
    o.doc.mixed, o.doc.in_table = True, True
    refresh(o.doc)
    result = call(o, occurrence=2)
    assert result["formatting"] == "mixed" and result["warnings"]
    assert result["location"]["start"] == 17 and result["location"]["row"] == 2
    result = call(o, "word_edit_apply", occurrence=2, new_text="")
    assert result["revisions_added"] == 1 and "target  after" in o.doc.Content.Text


def test_undo_conflict_after_user_edit(o):
    call(o, "word_edit_apply")
    o.doc.Content.Text += "user\r"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.doc.Name)


def test_journal_fragment_limit(o):
    _, summary = journal.describe({"expected_text": "a" * 250, "new_text": "b" * 250})
    assert "a" * 200 in summary and "a" * 201 not in summary
    assert "b" * 200 in summary and "b" * 201 not in summary


def test_strict_preview_and_exact_name(o, monkeypatch):
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, strict_target=True))
    with pytest.raises(ToolError, match="required"):
        call(o, document="")
    with pytest.raises(ToolError, match="not found"):
        call(o, document="Draft")


def comments(o, **kw):
    return o.call("word_manage_comments", document=o.doc.Name, **kw)


def test_comment_threads_paging_text_limits_and_stale_key(o):
    root = Comment(o.doc, "long comment", 7, 13)
    o.doc.Comments.items.append(root)
    root.Replies.Add(root.Scope, "reply text")
    o.doc.calls.clear()
    legacy = comments(o)
    assert legacy["comments"][0]["comment"] == "long comment" and legacy["count"] == 2
    page = comments(o, limit=1, max_chars=4, context_chars=3)
    item = page["comments"][0]
    assert page["total"] == 2 and page["returned"] == 1 and page["next_offset"] == 1
    assert item["comment"] == "long" and item["truncated"] and item["replies"][0]["text"] == "repl"
    assert item["date"] == "2026-01-02T03:04:05" and item["initials"] == "RV"
    assert item["context_before"] == "re " and item["context_after"] == " af"
    reply = comments(o, offset=1, limit=1)["comments"][0]
    assert reply["is_reply"] and reply["parent_index"] == 1 and reply["thread_key"] == item["thread_key"]
    root.Done = True
    assert comments(o, only_open=True)["total"] == 1
    o.doc.Comments.items.insert(0, Comment(o.doc, "another", author="other"))
    for action in ("reply", "resolve", "delete"):
        with pytest.raises(ToolError, match="другой ветке"):
            comments(o, action=action, index=1, text="agree", thread_key=item["thread_key"])
    assert not o.doc.calls and not o.word.starts
    comments(o, action="reply", index=2, text="agree", thread_key=item["thread_key"])
    assert root.Replies.Count == 2


def test_unavailable_comment_properties(o):
    root = Comment(o.doc, "comment")
    for key in ("Done", "Replies", "Ancestor", "Date", "Initial"):
        delattr(root, key)
    o.doc.Comments.items.append(root)
    result = comments(o, only_open=True)
    item = result["comments"][0]
    assert item["resolved"] is item["date"] is item["initials"] is item["is_reply"] is None
    assert {"Done", "Ancestor", "Date", "Initial"} <= set(result["unavailable_properties"])


@pytest.mark.parametrize("kw", [{"limit": 301}, {"limit": 0}, {"offset": -1}, {"max_chars": 0}, {"context_chars": -1}])
def test_comment_limits(o, kw):
    with pytest.raises(ToolError):
        comments(o, **kw)
    assert not o.doc.calls


def test_comment_list_no_journal_undo(o):
    o.doc.Comments.items.append(Comment(o.doc, "comment"))
    comments(o)
    assert not history(o) and not list(o.tmp.glob("journal/*")) and not o.word.starts
    assert "untrusted" in word_layout.word_manage_comments.__doc__



def test_comment_default_pagination_opt_in_preserves_legacy(o):
    o.doc.Comments.items[:] = [Comment(o.doc, "x" * 1100, author=f"Reviewer {i}") for i in range(55)]
    legacy = comments(o)
    assert legacy["returned"] == 55 and len(legacy["comments"][0]["comment"]) == 1100
    page = comments(o, offset=0)
    assert page["returned"] == 50 and page["next_offset"] == 50
    assert len(page["comments"][0]["comment"]) == 1000 and page["comments"][0]["truncated"]



def test_fragment_links_resolve_exact_range_without_selection(o):
    result = call(o)
    app, params = links.parse(result["link"])
    assert app == "word" and params["span"] == "7-13"
    _, doc, _, rng = navigation.resolve(app, params)
    assert doc is o.doc and rng.Text == "target" and not o.doc.calls
    item = comments(o)  # empty collection does not mutate anything
    assert item["returned"] == 0
    with pytest.raises(ToolError, match="no longer exists"):
        navigation.resolve("word", {"doc": o.doc.Name, "span": "0-999"})


@pytest.mark.parametrize("span", ["-1-2", "2-1", "01-2", "0-2147483648", "0-2-3", "a-2", "0-١", "0-2&paragraph=1"])
def test_invalid_fragment_link_before_com(span):
    with pytest.raises(ToolError):
        links.parse("officelive://word/open?doc=Draft.docx&span=" + span)



@pytest.mark.parametrize("text", ["x" * 10000, "tab\tline\vnext"])
def test_new_text_boundary_and_allowed_controls(o, text):
    assert not call(o, new_text=text)["blocked"]
