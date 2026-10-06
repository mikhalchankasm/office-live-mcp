"""Office operations Word tools and shared protection guard, using COM fakes only."""

import json
from dataclasses import asdict, replace
from types import SimpleNamespace as NS

import pytest

from office_live import com, config, registry, word_layout as layout, word_core as core, word_tables as tables, undo, wd_common
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401
from tests.office_operations_fakes import WordTable, word


@pytest.fixture
def o(office):
    word(office)
    return office


def restrict(o, **kw):
    return o.call("word_restrict_editing", document=o.doc.Name, **kw)


def insert(o, **kw):
    path = o.tmp / "source.docx"
    path.write_text("source")
    return o.call("word_insert_document", **{"document": o.doc.Name, "file_path": str(path), **kw})


def convert(o, **kw):
    return o.call("word_table_text", **{"document": o.doc.Name, "action": "text_to_table", "start_paragraph": 1, "end_paragraph": 2, **kw})


def history(o):
    return undo.STACKS.get(undo.stack_key("document", o.word, o.doc), [])


@pytest.mark.parametrize("mode,value", [(k, v) for k, v in layout.PROTECTION_MODES.items() if k != "none"])
def test_restrictions_no_native_undo_record_custom_inverse(o, mode, value):
    result = restrict(o, action="protect", mode=mode)
    assert result["mode"] == mode and result["undo"] == "available"
    assert o.doc.calls[0] == ("protect", value, True, "", False, False)
    assert not o.word.starts
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.ProtectionType == -1 and not o.doc.undo_calls and not o.doc.TrackRevisions
    restrict(o, action="protect", mode=mode)
    result = restrict(o, action="unprotect")
    assert result["undo"] == "available" and o.doc.ProtectionType == -1
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.ProtectionType == value


def test_restriction_status_editors(o):
    o.doc.Content.Editors = NS(Count=1, Item=lambda i: NS(ID="id", Name="name", Range=NS(Start=0, End=3)))
    result = restrict(o)
    assert result["mode"] == "none" and result["editors"] == [{"id": "id", "name": "name", "start": 0, "end": 3}]
    assert not history(o) and not o.doc.calls


@pytest.mark.parametrize("action", ["protect", "unprotect"])
def test_password_redacted_everywhere_barrier(o, monkeypatch, action):
    secret = "secret-never-retain"
    records = []
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, audit_log=o.tmp / "audit.jsonl", audit_content=True))
    monkeypatch.setattr(registry, "_audit_write", records.append)
    if action == "unprotect":
        o.doc.ProtectionType, o.doc.secret = 3, secret
    result = restrict(o, action=action, password=secret)
    assert "Password operation" in result["undo"]
    state = json.dumps({"history": [asdict(e) for e in history(o)], "audit": records,
                        "journal": [p.read_text(encoding="utf-8") for p in config.SETTINGS.journal_dir.glob("*.md")]}, default=str)
    assert secret not in state
    assert "<redacted>" in state and records


def test_wrong_password_and_duplicate_protection_no_change(o):
    restrict(o, action="protect", password="private")
    before = len(history(o))
    with pytest.raises(ToolError, match="Неверный пароль") as exc:
        restrict(o, action="unprotect", password="wrong-secret")
    assert "wrong-secret" not in str(exc.value) and o.doc.ProtectionType == 3 and len(history(o)) == before
    with pytest.raises(ToolError, match="уже защищён"):
        restrict(o, action="protect")


@pytest.mark.parametrize("mode", [1, 2, 3])
def test_shared_guard_refuses_all_write_tools_before_selected(o, monkeypatch, mode):
    o.doc.ProtectionType = mode
    picked = []
    monkeypatch.setattr(undo, "selected", lambda *a: picked.append(a))
    # Exercise every current Word writer under its actual name, even those with required COM arguments.
    writers = [info for info in registry.CATALOG.values() if info.group.startswith("word_") and info.kind in wd_common.WRITE_KINDS
               and info.name != "word_restrict_editing" and info.name not in wd_common.PROTECTION_EXEMPT
               and not (mode == 1 and info.name == "word_manage_comments")]
    assert len(writers) >= 20
    for info in writers:
        def operation(document):
            wd_common.pick_document(document)
            pytest.fail("protected write got past document selection")
        operation.__name__ = info.name
        with pytest.raises(ToolError, match="word_restrict_editing action=unprotect"):
            com.run_com(lambda info=info, operation=operation: undo.record_call(operation, (o.doc.Name,), {}, info.name, "write", {"document": o.doc.Name}, {}), kind="write")
    assert not picked and not o.word.starts


def test_comments_exception_and_tracked_mode_warning(o):
    o.doc.ProtectionType = 1
    o.word_write("comment", tool="word_manage_comments")
    o.doc.ProtectionType = 0
    result = o.word_write("tracked text")
    assert "tracked_changes" in result["warnings"][0]


@pytest.mark.parametrize("tool", ["bridge_word_table_to_excel", "bridge_word_text_to_excel", "bridge_excel_to_word_documents"])
def test_protected_bridge_source_can_be_read(o, tool):
    o.doc.ProtectionType = 3
    def source(document):
        wd_common.pick_document(document)
        return {"unchanged": True}
    result = com.run_com(lambda: undo.record_call(source, (o.doc.Name,), {}, tool, "write", {"document": o.doc.Name}, {}), kind="write")
    assert result["unchanged"] and not history(o)
    if tool != "bridge_excel_to_word_documents":
        assert not o.word.starts


def test_restriction_user_edits_block_inverse(o):
    restrict(o, action="protect")
    o.doc.Content.Text += "changed"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.doc.Name)
    assert o.doc.ProtectionType == 3


@pytest.mark.parametrize("position,kw,anchor", [("start", {}, 0), ("end", {}, 7), ("after_paragraph", {"paragraph": 1}, 4),
                                            ("after_bookmark", {"bookmark": "anchor"}, 4)])
def test_insert_positions_positional_contract_and_one_record(o, position, kw, anchor):
    before = o.doc.Content.Text
    result = insert(o, position=position, source_bookmark="source", **kw)
    assert (result["start"], result["end"]) == (anchor, anchor + len(o.doc.inserted_text))
    assert o.doc.calls[-1][2:] == ("source", False, False, False)
    assert o.word.AutomationSecurity == 1
    assert len(o.word.starts) == 1 and o.word.ends == 1
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == before and o.doc.undo_calls == [1]


@pytest.mark.parametrize("problem", ["extension", "missing", "outside", "self", "paragraph", "bookmark", "protected", "inside_table"])
def test_insert_preflight_refusals(o, monkeypatch, problem):
    source = o.tmp / "source.docx"
    source.write_text("source")
    kw = {"file_path": str(source)}
    if problem == "extension":
        kw["file_path"] = str(o.tmp / "macro.docm")
    elif problem == "missing":
        kw["file_path"] = str(o.tmp / "missing.docx")
    elif problem == "outside":
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, allowed_dirs=(o.tmp / "allowed",)))
    elif problem == "self":
        o.doc.Path, o.doc.FullName = str(o.tmp), str(source)
    elif problem == "paragraph":
        kw.update(position="after_paragraph", paragraph=99)
    elif problem == "bookmark":
        kw.update(position="after_bookmark", bookmark="absent")
    elif problem == "protected":
        o.doc.ProtectionType = 3
    else:
        o.doc.in_table = True
    before = o.doc.Content.Text
    with pytest.raises(ToolError):
        o.call("word_insert_document", document=o.doc.Name, **kw)
    assert o.doc.Content.Text == before and not o.doc.calls and not o.word.starts and not history(o)


def test_insert_saved_version_warning_and_partial_undo(o):
    source = NS(Path=str(o.tmp), FullName=str(o.tmp / "source.docx"), Name="source.docx", Saved=False)
    o.word.Documents.items.append(source)
    o.doc.fail = "insert"
    before = o.doc.Content.Text
    with pytest.raises(ToolError, match="applied=.*characters.*office_undo will restore"):
        insert(o)
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == before
    # Content's fake deepcopy only restores text; restore its positions for a second call.
    from tests.office_operations_fakes import refresh_word
    refresh_word(o.doc)
    o.doc.fail = ""
    result = insert(o)
    assert "сохранённая на диске версия" in result["warnings"][0]


@pytest.mark.parametrize("sep,value", [("tab", 1), ("comma", 2), ("semicolon", ";"), ("|", "|")])
def test_text_to_table_separator_header_style_undo(o, sep, value):
    before = o.doc.Content.Text
    result = convert(o, separator=sep, style="Grid")
    assert result["table"] == 1 and result["rows"] == 2
    assert o.doc.calls[-1] == ("to_table", value, result["rows"], result["columns"])  # explicit NumRows (live probe)
    assert o.doc.Tables(1).Rows(1).HeadingFormat
    assert o.doc.Tables(1).Style.Type == 3
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == before and o.doc.Tables.Count == 0


def test_text_to_table_auto_columns_uneven(o):
    o.doc.Content.Text = "a\tb\tc\rd\te\r"
    from tests.office_operations_fakes import refresh_word
    refresh_word(o.doc)
    result = convert(o, header_row=False)
    assert result["columns"] == 3 and result["uneven_rows"] == 1
    assert not o.doc.Tables(1).Rows(1).HeadingFormat


def test_text_to_table_counts_selected_empty_last_paragraph(o):
    o.doc.Content.Text = "a\tb\r\r"
    from tests.office_operations_fakes import refresh_word
    refresh_word(o.doc)
    result = convert(o)
    assert result["columns"] == 2 and result["uneven_rows"] == 1


@pytest.mark.parametrize("separator,value", [("tab", 1), ("comma", 2), ("paragraph", 0), (";", ";")])
def test_table_to_text_returns_paragraphs_and_undo(o, separator, value):
    tbl = WordTable(o.doc, [["a", "b"], ["c", "d"]])
    o.doc.Tables = Collection([tbl])
    result = convert(o, action="table_to_text", separator=separator)
    assert o.doc.calls[-1] == ("to_text", value, False)
    assert result["start_paragraph"] == 1 and result["end_paragraph"] >= 2
    assert o.doc.Tables.Count == 0
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Tables(1).data == [["a", "b"], ["c", "d"]]


@pytest.mark.parametrize("kw", [{"start_paragraph": 0}, {"end_paragraph": 3}, {"start_paragraph": 2, "end_paragraph": 1},
                               {"columns": True}, {"columns": -1}, {"columns": 64}, {"separator": "paragraph"},
                               {"separator": "bad"}, {"style": "Unknown"}, {"action": "bad"}])
def test_conversion_inputs_refuse_before_record(o, kw):
    with pytest.raises(ToolError):
        convert(o, **kw)
    assert not o.doc.calls and not o.word.starts


@pytest.mark.parametrize("problem", ["empty", "protected", "table", "nested", "merged", "nonuniform"])
def test_conversion_state_refusals(o, problem):
    if problem == "empty":
        o.doc.Content.Text = "   \r \r"
        kw = {}
    elif problem == "protected":
        o.doc.ProtectionType = 2
        kw = {}
    elif problem == "table":
        o.doc.in_table = True
        kw = {}
    else:
        tbl = WordTable(o.doc, [["a", "b"], ["c", "d"]])
        o.doc.Tables = Collection([tbl])
        if problem == "nested":
            tbl.Tables = Collection([NS()])
        elif problem == "merged":
            tbl.Range.Cells.items.pop()
        else:
            tbl.Uniform = False
        kw = {"action": "table_to_text"}
    with pytest.raises(ToolError):
        convert(o, **kw)
    assert not o.doc.calls and not o.word.starts


def test_conversion_partial_failure_one_record_undo(o):
    o.doc.fail = "to_table"
    with pytest.raises(ToolError, match="office_undo will restore"):
        convert(o)
    assert len(o.word.starts) == 1 and o.word.ends == 1
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Tables.Count == 0


def test_word_restriction_readonly_and_writer_registration(o, monkeypatch):
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    registered = {}
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda fn, name, **kw: registered.__setitem__(name, fn)))
    registry.office_tool("word_layout", "write", read_actions=("status",))(layout.word_restrict_editing)
    assert registered["word_restrict_editing"](document=o.doc.Name)["mode"] == "none"
    with pytest.raises(ToolError, match="Read-only"):
        registered["word_restrict_editing"](document=o.doc.Name, action="protect")
    for name in ("word_insert_document", "word_table_text"):
        registry.office_tool(registry.CATALOG[name].group, "write")(getattr(core if name == "word_insert_document" else tables, name))
        assert not registry.CATALOG[name].registered


@pytest.mark.parametrize("tool", sorted(wd_common.PROTECTION_EXEMPT))
def test_protection_guard_allows_non_content_tools(o, tool):
    # Live probe: without the exemption a protected document could not even be closed.
    o.doc.ProtectionType = 2

    def operation(document):
        return wd_common.pick_document(document)[1].Name
    operation.__name__ = tool
    assert com.run_com(lambda: undo.record_call(operation, (o.doc.Name,), {}, tool, "write", {"document": o.doc.Name}, {}), kind="write") == o.doc.Name


def test_word_fingerprint_survives_unreadable_track_revisions(o, monkeypatch):
    # Live probe: forms protection makes even reading TrackRevisions raise.
    import pywintypes

    class Locked(type(o.doc)):
        @property
        def TrackRevisions(self):
            raise pywintypes.com_error(-2147352567, "Error", (0, "Microsoft Word", "protected", None, 0, -2146823683), None)
    o.doc.__class__ = Locked
    assert undo.word_fingerprint(o.doc) and undo.word_fingerprint(o.doc, styles=False)


def test_tracked_protection_check_skips_documents_closed_during_the_call():
    # Live regression: word_compare_documents closes its temporary copy; reading ProtectionType afterwards raised
    # RPC_E_DISCONNECTED and failed the whole compare call.
    import pywintypes

    class Closed:
        @property
        def ProtectionType(self):
            raise pywintypes.com_error(-2147417848, "The object invoked has disconnected from its clients.", None, None)

    assert undo._tracked_protection(Closed()) is False
    assert undo._tracked_protection(type("Tracked", (), {"ProtectionType": 0})()) is True
