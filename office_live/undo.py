"""Сеансовая отмена: Word UndoRecord и собственные снимки Excel. В стеке нет COM-прокси."""

import re
import uuid

import pywintypes

from . import com, config, journal, wd_common as wd, xl_common as xl
from .errors import PartialChangeError, ToolError
from .privacy import redact
from .registry import office_tool
from .util import parse_a1

from .undo_excel import (
    MARKER as MARKER,
    SESSION as SESSION,
    _hash as _hash,
    _shown as _shown,
    _quiet as _quiet,
    _internal_events as _internal_events,
    _keep_sheet as _keep_sheet,
    _process_alive as _process_alive,
    _close_orphan_backups as _close_orphan_backups,
    _backup as _backup,
    _area as _area,
    _range as _range,
    _size as _size,
    _dimensions as _dimensions,
    _restore_dimensions as _restore_dimensions,
    _formula_state as _formula_state,
    _format_state as _format_state,
    _optional_state as _optional_state,
    _metadata_state as _metadata_state,
    _rules as _rules,
    _rule_state as _rule_state,
    _inside as _inside,
    _rule_closure as _rule_closure,
    _restore_rules as _restore_rules,
    _property_object as _property_object,
    _property_state as _property_state,
    _restore_properties as _restore_properties,
    _repair_formulas as _repair_formulas,
    Entry as Entry,
    _snapshots as _snapshots,
    _drop as _drop,
    _view as _view,
    _forget_marker as _forget_marker,
    _name_state as _name_state,
    _sheet_content_state as _sheet_content_state,
    _sheet_state as _sheet_state,
    excel_fingerprint as excel_fingerprint,
    _simple_range as _simple_range,
    IMPORT_BOOK_BARRIER as IMPORT_BOOK_BARRIER,
    _union_range as _union_range,
    _import_view as _import_view,
    _import_table as _import_table,
    _write_range as _write_range,
    _clean_text as _clean_text,
    _split_column as _split_column,
    _protection as _protection,
    _what_if as _what_if,
    _goal_seek as _goal_seek,
    _subtotals as _subtotals,
    _sparklines as _sparklines,
    _autofill as _autofill,
    _replace as _replace,
    _dimensions_plan as _dimensions_plan,
    _format_ranges as _format_ranges,
    _hidden as _hidden,
    _rows_columns as _rows_columns,
    _add_sheet as _add_sheet,
    _delete_sheet as _delete_sheet,
    _manage_sheet as _manage_sheet,
    _names as _names,
    _shape as _shape,
    _slicer as _slicer,
    _table as _table,
    _sheet_view as _sheet_view,
    _calculate as _calculate,
    _bridge as _bridge,
    RESOLVERS as RESOLVERS,
    _finish_excel as _finish_excel,
    restore_excel as restore_excel,
)
STACKS: dict[tuple, list] = {}
_active = None  # доступ сериализован com.LOCK; живёт только внутри run_com


def identity(kind, obj):
    return (xl.workbook_path(obj) if kind == "workbook" else wd.document_path(obj)) or obj.Name


def stack_key(kind, app, obj):
    # Несохранённые Book1 в разных процессах Excel не должны делить историю. У Word.Application нет Hwnd
    # (обращение к нему ломало бы каждую правку в Word), а Word работает одним процессом на пользователя.
    return kind, identity(kind, obj).lower(), int(app.Hwnd) if kind == "workbook" else 0




_PART = re.compile(r'<pkg:part pkg:name="(/word/[^"]+)"(.*?)</pkg:part>', re.S)
# Части, которые меняет пользователь: тело, колонтитулы, примечания, сноски, стили, нумерация. settings/theme/шрифты — нет.
_USER_PARTS = re.compile(r"/word/(document|header\d*|footer\d*|comments\w*|footnotes|endnotes|styles|numbering)\.xml")
# Что Word генерирует сам при каждом чтении WordOpenXML (проверено на живом Word): идентификаторы абзацев и текста
# (w14:paraId/textId), сеансов правки (rsid), пометки фоновой проверки орфографии (proofErr).
# wp14:anchorId/editId: Word regenerates drawing ids on undo/redo (live probe 0.5, inline picture after caption undo).
_XML_NOISE = re.compile(r'<w:proofErr\b[^>]*/>|\s(?:w14:paraId|w14:textId|w:rsid\w*|wp14:anchorId|wp14:editId)="[^"]*"')


def _word_body_xml(doc):
    """Пользовательские части документа из WordOpenXML без шума — без единого обращения к колонтитулам через COM:
    даже Headers(i).Exists создаёт в документе пустые колонтитулы (живая проверка), то есть меняет документ пользователя
    и сам отпечаток. WordOpenXML всей области Content уже содержит колонтитулы, примечания и сноски."""
    xml = doc.Content.WordOpenXML
    parts = [(name, _XML_NOISE.sub("", body)) for name, body in _PART.findall(xml) if _USER_PARTS.fullmatch(name)]
    return _without_empty_headers(parts) if parts else _XML_NOISE.sub("", xml)


_HF_NAME = re.compile(r"/word/(header|footer)\d*\.xml")
_HF_REF = re.compile(r'<w:(?:header|footer)Reference\b[^>]*/>')
_NOTES_NAME = re.compile(r"/word/(footnotes|endnotes)\.xml")
_NOTE_SEPARATOR = re.compile(r'<w:(foot|end)note\b[^>]*\bw:type="[^"]*".*?</w:\1note>', re.S)
_HF_OBJECT = re.compile(r'<v:shape[\s>]|<w:drawing\b|<w:fldSimple\b|<w:instrText\b|<w:fldChar\b')


def _without_empty_headers(parts):
    """Empty header/footer parts are invisible but unstable: Word creates them when a header is touched, keeps them
    after undo, and WordOpenXML sometimes omits their references (live probes 0.5). Drop empty ones and all references;
    compare non-empty headers/footers by content (their part numbers may shift)."""
    kept, headers = [], []
    for name, body in parts:
        if _NOTES_NAME.fullmatch(name) and not re.search(r"<w:(?:foot|end)note\b", _NOTE_SEPARATOR.sub("", body)):
            continue  # only Word's separator notes: added with the first header (live probe 0.5), invisible
        if _HF_NAME.fullmatch(name):
            text = re.sub(r"<[^>]+>", "", body.split("<pkg:xmlData>", 1)[-1]).strip()
            if text or _HF_OBJECT.search(body):
                headers.append(body)
        else:
            kept.append((name, _HF_REF.sub("", body)))
    return kept + [("/word/header-footer", body) for body in sorted(headers)]


def watermark_fingerprint(doc):
    """Fingerprint used to verify the custom watermark-add undo (styles excluded, like the native-undo check)."""
    return word_fingerprint(doc, styles=False)


def entry_fingerprint(entry, doc):
    """Отпечаток для проверки «пользователь менял документ после агента». Для защиты — без styles.xml: Word сам дописывает
    туда служебные стили, пока документ защищён «только исправления» (живая проверка), это не правка пользователя."""
    if entry.ops and entry.ops[0].get("op") == "word_watermark_add":
        return watermark_fingerprint(doc)  # same light form as the verification after the custom undo
    return word_fingerprint(doc, styles=entry.tool != "word_restrict_editing")


def _tracked_protection(doc) -> bool:
    """Защищён ли документ «только исправления». Закрытый за время вызова документ (временная копия сравнения) даёт
    RPC_E_DISCONNECTED — живая проверка: такой документ просто пропускаем."""
    try:
        return int(doc.ProtectionType) == 0
    except pywintypes.com_error as exc:
        if com.is_busy(exc):
            raise
        return False


# Word records some agent operations outside the custom UndoRecord as separate native steps: header-story shapes
# (watermarks), section breaks with PageSetup and header/footer links (live probes 0.5 needed more than two). The loop
# stops at the exact state before the agent and redoes everything if it never gets there, so a generous cap is safe.
WORD_EXTRA_STEPS = 30


def word_fingerprint(doc, styles=True):
    """styles=False — для проверки отмены защиты: Word сам дописывает в styles.xml служебные стили (скрытый «Revision»
    при защите «только исправления», стили примечаний при добавлении примечания) — живая проверка."""
    state = [doc.Content.Text, int(doc.Paragraphs.Count), int(doc.Tables.Count), int(doc.InlineShapes.Count)]
    try:
        tracking = bool(doc.TrackRevisions)
    except pywintypes.com_error as exc:
        if com.is_busy(exc) or com.is_dead(exc):
            raise
        tracking = "unavailable"  # защита «только формы»: Word не даёт даже прочитать TrackRevisions (живая проверка)
    state.extend([int(doc.ProtectionType), tracking])
    try:
        body = _word_body_xml(doc)  # замечает и ручное форматирование: без этого Undo(1) отменил бы правку пользователя
        if not styles and isinstance(body, list):
            body = [(name, part) for name, part in body if not name.startswith("/word/styles")]
        state.append(body)
    except (AttributeError, pywintypes.com_error):
        pass
    state.extend([int(doc.Comments.Count), int(doc.Footnotes.Count), int(doc.Endnotes.Count), int(doc.Shapes.Count)])
    return _hash(state)




EXCLUDED = {"office_journal", "office_undo", "excel_new_workbook", "word_new_document", "excel_create_from_template",
            "excel_close_workbook", "word_close_document", "word_reload_document", "word_compare_documents"}

DEFERRED = {"word_replace_text", "excel_what_if", "word_sections", "word_watermark", "excel_import_table", "excel_clean_text", "excel_split_column", "excel_protection", "word_sort_table", "excel_format_ranges",
            "excel_goal_seek", "excel_subtotals", "excel_sparklines", "word_insert_document", "word_table_text", "word_restrict_editing", "word_edit_apply", "word_manage_comments", "word_captions"}


class Recording:
    def __init__(self, tool, kind, arguments, supplied):
        self.tool, self.kind, self.args, self.supplied = tool, kind, redact(arguments), redact(supplied)
        self.args["_has_secret"] = any("password" in k.casefold() and bool(v) for k, v in arguments.items())
        self.targets = {}
        self.pending = {}
        self.records = []
        self.preparing = False
        self.resolved = False
        self.enabled = config.SETTINGS.undo and kind in {"write", "destructive"} and tool not in EXCLUDED

    def target(self, kind, app, obj, ready=False):
        if self.preparing:
            return
        if self.tool.startswith("bridge_"):
            into_excel = self.tool in {"bridge_word_table_to_excel", "bridge_word_text_to_excel"}
            if (kind == "workbook") != into_excel:
                return  # источник моста только читается
        key = stack_key(kind, app, obj)
        self.targets[key] = (app, obj)
        if self.tool in DEFERRED and not ready:
            return
        if not self.enabled or key in self.pending:
            return
        if self.tool == "excel_calculate" and not self.args.get("mode"):
            return
        if self.resolved and self.tool in {"excel_copy_range", "excel_manage_sheet"}:
            return
        entry = Entry(self.tool, journal.describe(self.supplied)[0], kind)
        self.pending[key] = entry
        self.preparing = True
        try:
            if kind == "document":
                if self.tool == "word_restrict_editing":
                    from .excel_format import PASSWORD_BARRIER

                    entry.before_fingerprint = word_fingerprint(obj, styles=False)
                    if self.args.get("_has_secret"):
                        entry.barrier(PASSWORD_BARRIER)
                    else:
                        entry.ops.append({"op": "word_protection", "mode": self.args["_word_protection_before"], "tracking": self.args.get("_word_tracking_before")})
                elif bool(app.UndoRecord.IsRecordingCustomRecord):
                    entry.barrier("Word is already recording another custom undo record.")
                elif self.tool == "word_watermark" and self.args.get("_watermark_add"):
                    # Custom inverse: Word's native undo of header-story shapes is unreliable (live probe 0.5 left a
                    # watermark behind after undo). Undo deletes the header shapes this call created: every shape in
                    # its target headers that did not exist before (also an unnamed one left by an interrupted call).
                    entry.before_fingerprint = watermark_fingerprint(obj)
                    entry.ops.append({"op": "word_watermark_add", "targets": list(self.args["_watermark_add"]),
                                      "before_names": list(self.args["_header_shapes_before"])})
                else:
                    record = app.UndoRecord
                    entry.before_fingerprint = word_fingerprint(obj)
                    entry.before_light = word_fingerprint(obj, styles=False)
                    record.StartCustomRecord(f"Office Live: {self.tool}")
                    self.records.append(record)
            elif self.tool == "excel_copy_range":
                self._copy(app, obj, key, entry)
                self.resolved = True
            elif self.tool == "excel_manage_sheet" and self.args.get("action", "").lower() == "copy" and self.args.get("dest_workbook"):
                dest_app, dest = xl.pick_workbook(self.args["dest_workbook"])
                dest_key = stack_key(kind, dest_app, dest)
                self.targets.pop(key, None)
                self.pending.pop(key, None)
                self.targets[dest_key] = (dest_app, dest)
                self.pending[dest_key] = entry
                _add_sheet(dest_app, dest, entry, self.args)
                self.resolved = True
            else:
                resolver = RESOLVERS.get(self.tool)
                if resolver is None:
                    entry.barrier("This operation has no safe undo strategy.")
                else:
                    resolver(app, obj, entry, self.args)
            for pending_key, pending in self.pending.items():
                if pending.kind == "workbook" and pending.undoable and not pending.before_fingerprint:
                    pa, po = self.targets[pending_key]
                    pending.before_fingerprint = excel_fingerprint(pa, po, pending)
                prior = STACKS.get(pending_key, [])
                if prior and prior[-1].undoable:
                    pa, po = self.targets[pending_key]
                    try:
                        current = entry_fingerprint(prior[-1], po) if pending.kind == "document" else excel_fingerprint(pa, po, prior[-1])
                        pending.refresh_previous = current == prior[-1].fingerprint
                    except Exception:  # noqa: BLE001
                        pending.refresh_previous = False
        except Exception as exc:  # noqa: BLE001 — невозможность снимка создаёт барьер, но не ломает инструмент
            reason = com.com_error_text(exc) if isinstance(exc, pywintypes.com_error) else str(exc)
            if isinstance(exc, AttributeError) and reason.startswith("<unknown>."):
                reason = f"Office did not answer the call {reason[10:]} (busy?)"
            for pending in self.pending.values():
                pending.barrier(f"Could not prepare undo: {reason}")
        finally:
            self.preparing = False

    def _copy(self, app, wb, key, entry):
        args = self.args
        ws, source = xl.get_range(wb, args["sheet"], args["source"], empty_means_used=False)
        da, dw = xl.pick_workbook(args["dest_workbook"]) if args.get("dest_workbook") else (app, wb)
        dest_key = stack_key("workbook", da, dw)
        dest_ws = xl.pick_sheet(dw, args["dest_sheet"]) if args.get("dest_sheet") else (ws if not args.get("dest_workbook") else xl.pick_sheet(dw, ""))
        r1, c1, _, _ = parse_a1(args["dest_cell"])
        rows, cols = int(source.Rows.Count), int(source.Columns.Count)
        if args.get("transpose"):
            rows, cols = cols, rows
        dest = xl.sub_range(dest_ws, r1, c1, r1 + rows - 1, c1 + cols - 1)
        if ":" in args["dest_cell"] and args.get("what", "all").lower() in {"all", "formats"}:
            explicit = dest_ws.Range(args["dest_cell"])
            if _size(explicit) > 1:  # D1:D1 — тоже одиночная ячейка, Copy расширяет её до размера источника
                dest = explicit
        if args.get("move"):
            entry.warning = "Undo restores the source and destination cells, but formulas elsewhere that pointed to the moved cells keep pointing to the destination."
        if key == dest_key:
            mode = "contents" if args.get("what", "all").lower() in {"values", "formulas"} else "formats" if args.get("what", "all").lower() == "formats" else "all"
            _snapshots(app, wb, entry, [dest, source] if args.get("move") else [dest], restore=mode,
                       properties=("NumberFormat",) if mode == "contents" else ())
        else:
            de = Entry(self.tool, f"{dest_ws.Name}!{xl.addr_of(dest)}", "workbook")
            de.warning = entry.warning
            self.targets[dest_key] = da, dw
            self.pending[dest_key] = de
            if args.get("move") and _size(source) + _size(dest) > config.SETTINGS.undo_max_cells:
                for item in (entry, de):
                    item.barrier("Combined source and destination exceeds OFFICE_LIVE_UNDO_MAX_CELLS")
                return
            mode = "contents" if args.get("what", "all").lower() in {"values", "formulas"} else "formats" if args.get("what", "all").lower() == "formats" else "all"
            _snapshots(da, dw, de, [dest], restore=mode, properties=("NumberFormat",) if mode == "contents" else ())
            if args.get("move"):
                _snapshots(app, wb, entry, [source])
            else:
                self.targets.pop(key, None)
                self.pending.pop(key, None)

    def end_records(self):
        while self.records:
            record = self.records.pop()
            try:
                record.EndCustomRecord()
            except Exception as exc:  # noqa: BLE001
                for entry in self.pending.values():
                    if entry.kind == "document":
                        entry.barrier(f"Could not end Word undo record: {exc}")

    def finish(self, result, partial=False):
        self.end_records()
        if self.tool == "word_reload_document":
            if isinstance(result, dict) and partial:
                result.setdefault("undo", "not available: reload closed the old instance and dropped its history")
            return  # reload already dropped history; never inspect the closed COM instance
        if self.tool in {"excel_close_workbook", "word_close_document"}:
            for key, (app, _) in self.targets.items():
                forget_closed_target(key, app)
                if key[0] == "workbook":
                    try:
                        _close_orphan_backups(app)
                    except Exception:  # noqa: BLE001 — уборка не должна портить успешное закрытие книги
                        pass
            return  # закрытый документ уже нельзя читать или дополнять листом «Лог»
        if self.kind == "save":
            for key, (app, obj) in self.targets.items():
                new_key = stack_key(key[0], app, obj)
                if new_key != key:
                    com.note_target(f"{key[0]}:{identity(key[0], obj)}")
                    for entry in STACKS.pop(new_key, []):
                        _drop(app, entry)
                    if key in STACKS:
                        STACKS[new_key] = STACKS.pop(key)
                    for stack in STACKS.values():
                        for entry in stack:
                            entry.peers = [new_key if peer == key else peer for peer in entry.peers]
        if len(self.pending) > 1:
            transaction = uuid.uuid4().hex
            for entry in self.pending.values():
                entry.transaction = transaction
                entry.peers = list(self.pending)
        # Идентификация созданного листа должна предшествовать созданию служебного листа «Лог».
        for key, entry in self.pending.items():
            app, obj = self.targets[key]
            if entry.kind == "workbook" and entry.undoable:
                try:
                    _finish_excel(obj, entry, result if isinstance(result, dict) else {})
                except Exception as exc:  # noqa: BLE001
                    if partial:
                        entry.force_reason = f"Could not read the partial result: {exc}; pass force=true to restore the pre-call snapshot."
                    else:
                        entry.barrier(f"Could not finish undo snapshot: {exc}")
        from .navigation import result_links

        result_links(self, result)
        if self.kind in {"write", "destructive", "save"} and isinstance(result, dict):
            if any(key[0] == "document" and _tracked_protection(obj) for key, (_, obj) in self.targets.items()):
                result.setdefault("warnings", []).append("Документ защищён в режиме tracked_changes: Word записывает правки исправлениями.")
        if self.kind in {"write", "destructive"} and self.tool != "office_undo":
            for key, (app, obj) in self.targets.items():
                if key[0] == "workbook":
                    warning = journal.append_sheet(app, obj, self.tool, self.supplied, links=result.get("links", []) if isinstance(result, dict) else [])
                    if warning and isinstance(result, dict):
                        result["log_sheet_warning"] = warning
        statuses = []
        for key, entry in self.pending.items():
            app, obj = self.targets[key]
            if entry.undoable:
                try:
                    entry.fingerprint = entry_fingerprint(entry, obj) if entry.kind == "document" else excel_fingerprint(app, obj, entry)
                except Exception as exc:  # noqa: BLE001
                    if partial:
                        entry.force_reason = f"Could not fingerprint the partial result: {exc}; pass force=true to restore the pre-call state."
                    else:
                        entry.barrier(f"Could not fingerprint the result: {exc}")
            if not entry.undoable:
                _drop(app, entry)
            if entry.kind == "document" and entry.undoable and entry.fingerprint == entry.before_fingerprint:
                statuses.append("not available: no document change detected")
                continue  # Word не создаёт отдельного undo-действия для пустой правки
            stack = STACKS.setdefault(key, [])
            stack.append(entry)
            while len(stack) > config.SETTINGS.undo_depth:
                _drop(app, stack.pop(0))
            statuses.append("available" if entry.undoable else "not available: " + entry.reason)
            if entry.warning and isinstance(result, dict):
                result["undo_warning"] = entry.warning
            if entry.force_reason and isinstance(result, dict):
                result["undo_warning"] = " ".join(filter(None, (entry.warning, entry.force_reason)))
        if statuses and isinstance(result, dict):
            result["undo"] = next((s for s in statuses if s != "available"), "available")

    def close(self, success):
        self.end_records()
        if not success:
            for key, entry in self.pending.items():
                app, obj = self.targets[key]
                stack = STACKS.get(key, [])
                if stack:
                    try:
                        current = entry_fingerprint(entry, obj) if entry.kind == "document" else excel_fingerprint(app, obj, entry)
                    except Exception:  # noqa: BLE001
                        current = None
                    if not entry.before_fingerprint or current != entry.before_fingerprint:
                        # Нового шага нет. Но перескочить через частично выполненную правку тоже нельзя.
                        stack[-1].barrier(f"A failed {self.tool} may have changed the document or native undo history.")
                _drop(app, entry)
        self.targets.clear()  # все COM-ссылки отпускаем до CoUninitialize
        self.pending.clear()


def forget_closed_target(key, app):
    """Drop stale history immediately after closing, including a reload whose reopen might fail."""
    for entry in STACKS.pop(key, []):
        _drop(app, entry)


def selected(kind, app, obj):
    """Вызывается выборщиками после AutoSave/событий, но до первой правки инструмента."""
    if _active is not None:
        _active.target(kind, app, obj)


def require_undo(kind, app, obj, *, plan=None, barrier_reason=""):
    """Новые операции начинаются только после проверок и успешной подготовки отмены."""
    if _active is None:
        raise ToolError("This operation requires the office_tool undo context.")
    if not config.SETTINGS.undo:
        return  # отмена выключена владельцем (OFFICE_LIVE_UNDO=0): работаем, как остальные инструменты
    if not _active.enabled:
        raise ToolError("This operation requires the office_tool undo context.")
    if plan:
        _active.args.update(plan)
    _active.target(kind, app, obj, ready=True)
    entry = _active.pending.get(stack_key(kind, app, obj))
    if entry is not None and barrier_reason and entry.reason == barrier_reason:
        return  # explicitly chosen password barrier; snapshot failures are still refused
    if entry is None or not entry.undoable:
        if entry is not None:
            _drop(app, entry)
            _active.pending.pop(stack_key(kind, app, obj), None)
        raise ToolError("No changes made: could not prepare office_undo. " + (entry.reason if entry else "No snapshot."))


def cancel_prepared(kind, app, obj):
    """Only for a tool that has verified complete local rollback after a failed mutation."""
    if _active is not None:
        entry = _active.pending.pop(stack_key(kind, app, obj), None)
        if entry is not None:
            _drop(app, entry)


def record_call(fn, args, kwargs, tool, kind, arguments, supplied):
    global _active
    recording = Recording(tool, kind, arguments, supplied)
    previous, _active = _active, recording
    success = False
    try:
        result = fn(*args, **kwargs)
        if not isinstance(result, dict) or result.get("ok", True):
            recording.finish(result)
            if tool == "excel_import_table":
                from .excel_import import _compact_import_result

                _compact_import_result(result)
            success = True
        return result
    except PartialChangeError as exc:
        result = {}
        recording.finish(result, partial=True)
        success = True
        status = result.get("undo", "not available: no change detected")
        hint = "office_undo will restore the state before this call." if status == "available" else f"office_undo: {status}."
        if result.get("undo_warning"):
            hint += " " + result["undo_warning"]
        raise ToolError(f"{exc} {hint}") from None
    finally:
        recording.close(success)
        _active = previous


def _pick_file(file):
    # Ошибка неоднозначности/защиты Excel не должна маскироваться выбором Word.
    try:
        app, obj = xl.pick_workbook(file)
        return "workbook", app, obj
    except ToolError as exc:
        text = str(exc)
        if not any(s in text for s in ("is not running", "has no open workbooks", "not found. Open workbooks")):
            raise
    app, obj = wd.pick_document(file)
    return "document", app, obj


def _available(app, entry):
    if entry.undoable and entry.backup_sheets:
        try:
            backup = _backup(app, entry.backup_token)
            for name in entry.backup_sheets:
                backup.Sheets(name)
        except Exception as exc:  # noqa: BLE001
            entry.barrier(str(exc))
    return entry.undoable


def _participants(key, app, obj, entry):
    """Один move между книгами отменяется целиком; нельзя обходить новые правки второй книги."""
    out = []
    for peer in entry.peers or [key]:
        stack = STACKS.get(peer, [])
        if peer != key:
            if not any(e.transaction == entry.transaction for e in stack):
                raise ToolError(f"{entry.tool}: related undo history for {peer[1]} is no longer available.")
            if stack[-1].transaction != entry.transaction:
                raise ToolError(f"{entry.tool}: undo the newer changes in {peer[1]} first.")
            pa, po = xl.pick_workbook(peer[1]) if peer[0] == "workbook" else wd.pick_document(peer[1])
            if stack_key(peer[0], pa, po) != peer:
                raise ToolError(f"{entry.tool}: the original Office instance for {peer[1]} is no longer available.")
        else:
            pa, po = app, obj
        out.append((peer, pa, po, stack[-1]))
    return out


@office_tool("history", "destructive", title="Undo agent changes", read_actions=("history",))
def office_undo(file: str = "", action: str = "undo", steps: int = 1, force: bool = False) -> dict:
    """Undo recent agent changes to an open workbook/document, or read the session's undo history. Stops at unsupported operations (barriers). Later user edits block undo unless force=true. In Word, force also undoes the user's later edits to reach the state before the agent change, up to 20 native steps per entry; reports native_steps and stops at a barrier if that state cannot be reached. Excel is tried first, then Word.

    Args:
        file: workbook/document name/path (empty means active; exact targets are safer).
        action: 'undo' or 'history' (history is also available in read-only mode).
        steps: number of agent changes to undo, newest first.
        force: accept that later user edits may be overwritten; in Word also undo the user's later edits (up to 20 native steps per entry).
    """
    act = action.lower()
    if act not in {"undo", "history"}:
        raise ToolError("action must be undo or history.")
    if steps < 1:
        raise ToolError("steps must be positive.")
    kind, app, obj = _pick_file(file)
    key = stack_key(kind, app, obj)
    stack = STACKS.get(key, [])
    if act == "history":
        return {"file": identity(kind, obj), "history": [
            {"tool": e.tool, "where": e.where, "time": e.stamp, "undoable": _available(app, e),
             "needs_force": bool(e.force_reason), "reason": e.reason or e.force_reason}
            for e in reversed(stack)
        ]}
    undone = []
    out = {"file": identity(kind, obj), "undone": undone}
    if kind == "document":
        out["native_steps"] = 0
    for _ in range(steps):
        if not stack:
            if not undone:
                raise ToolError("No recorded changes to undo in this server session.")
            break
        entry = stack[-1]
        try:
            participants = _participants(key, app, obj, entry)
            for pk, pa, po, pe in participants:
                if not _available(pa, pe):
                    raise ToolError(f"Undo stopped at {pe.tool}: not available: {pe.reason}")
                if pe.force_reason and not force:
                    raise ToolError(f"{pe.tool}: {pe.force_reason}")
                try:
                    current = entry_fingerprint(pe, po) if pk[0] == "document" else excel_fingerprint(pa, po, pe)
                except Exception as exc:  # noqa: BLE001
                    current = f"unavailable: {exc}"
                if current != pe.fingerprint and not force:
                    raise ToolError(f"{pe.tool}: the document changed after the agent. Undo would also overwrite/undo the user's later edits; pass force=true to proceed.")
        except ToolError as exc:
            if not undone:
                raise
            out["stopped"] = str(exc)
            break
        native_steps = 0
        try:
            for pk, pa, po, pe in participants:
                if pk[0] == "document":
                    if pe.ops and pe.ops[0].get("op") == "word_watermark_add":
                        from .word_structure import _in_header

                        existed = set(pe.ops[0]["before_names"])
                        for owner, typ, _ in pe.ops[0]["targets"]:
                            shapes = po.Sections(owner).Headers(typ).Shapes
                            for i in range(int(shapes.Count), 0, -1):
                                shape = shapes(i)
                                if str(shape.Name) not in existed and _in_header(shape, owner, typ):
                                    shape.Delete()
                        if watermark_fingerprint(po) != pe.before_fingerprint:
                            raise ToolError("Watermark undo could not be verified.")
                        continue
                    if pe.tool == "word_restrict_editing":
                        op = pe.ops[0]
                        if int(po.ProtectionType) != -1:
                            po.Unprotect("")
                        if op["tracking"] is not None:
                            po.TrackRevisions = op["tracking"]  # set while unprotected; protected Word may refuse even the same value
                        if op["mode"] != -1:
                            po.Protect(op["mode"], True, "", False, False)
                        if word_fingerprint(po, styles=False) != pe.before_fingerprint:
                            raise ToolError("Word protection undo could not be verified.")
                        continue
                    for _ in range(20 if force else 1):
                        if force and word_fingerprint(po) == pe.before_fingerprint:
                            break
                        if not po.Undo(1):
                            raise ToolError("Word refused Undo(1); its undo history may have been cleared.")
                        native_steps += 1
                        out["native_steps"] += 1
                    if not force and pe.before_light and word_fingerprint(po, styles=False) != pe.before_light:
                        # Word sometimes splits one agent change into several native steps (live probe: a paragraph style
                        # under tracked-changes protection lands outside the custom record). No user edit followed the agent
                        # (checked above), so up to two extra steps are the agent's own; if that still does not reach the
                        # state before the agent, Redo them back and say so.
                        extra = 0
                        while extra < WORD_EXTRA_STEPS and word_fingerprint(po, styles=False) != pe.before_light and po.Undo(1):
                            extra += 1
                        if word_fingerprint(po, styles=False) == pe.before_light:
                            native_steps += extra
                            out["native_steps"] += extra
                        else:
                            if extra:
                                po.Redo(extra)
                            out.setdefault("warnings", []).append(f"{pe.tool}: Word's single undo step did not fully restore the document; check it or use Ctrl+Z.")
                    if force and word_fingerprint(po) != pe.before_fingerprint:
                        raise ToolError("The limit of 20 native undo steps was reached.")
                else:
                    restore_excel(pa, po, pe)
        except Exception as exc:
            for _, _, _, pe in participants:
                if pe.kind == "document" and force:
                    pe.barrier(f"Could not reach the state before {pe.tool}; {native_steps} native steps were undone, including the user's edits. {exc}")
                else:
                    pe.barrier(f"Undo failed (possibly partially applied): {exc}")
            if not undone:
                raise ToolError(f"{entry.tool}: {entry.reason}") from None
            out["stopped"] = entry.reason
            break
        undone.append({"tool": entry.tool, "where": entry.where, **({"native_steps": native_steps} if kind == "document" else {})})
        for pk, pa, po, pe in participants:
            peer_stack = STACKS[pk]
            peer_stack.pop()
            _drop(pa, pe)
            summary = f"отменено: {pe.tool} ({pe.where})"
            from .navigation import undo_links

            linked = undo_links(pa, po, pe)
            places = linked.get("links", [])
            if places:
                existing = out.get("links", [])
                combined = list({link["uri"]: link for link in [*existing, *places]}.values())
                out.update(links=combined[:20], links_truncated=out.get("links_truncated", False) or linked.get("links_truncated", False) or len(combined) > 20)
            journal.append([f"{pk[0]}:{identity(pk[0], po)}"], "office_undo", {}, where=pe.where, summary=summary, links=places)
            if pk[0] == "workbook":
                warning = journal.append_sheet(pa, po, "office_undo", {}, where=pe.where, summary=summary, links=places)
                if warning:
                    out["log_sheet_warning"] = warning
            if pe.warning:
                out["warning"] = pe.warning
            # Ручные правки МЕЖДУ вызовами не становятся «своими» после отмены нового шага.
            if peer_stack and peer_stack[-1].undoable and pe.refresh_previous:
                try:
                    peer_stack[-1].fingerprint = entry_fingerprint(peer_stack[-1], po) if pk[0] == "document" else excel_fingerprint(pa, po, peer_stack[-1])
                except Exception as exc:  # noqa: BLE001
                    peer_stack[-1].barrier(f"Could not fingerprint after undo: {exc}")
    return out
