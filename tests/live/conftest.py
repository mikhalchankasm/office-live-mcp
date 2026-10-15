"""Общие фикстуры и помощники живых тестов (реальные Excel/Word через настоящий MCP-протокол).

Фикстуры создают СВОИ черновые книги/документы и закрывают только их (discard); приложения не закрываются, кроме Word,
и Excel, если их запустили сами тесты и в них не осталось документов/книг.
"""

import os
import tempfile
import threading
import time
import uuid

import pytest

from tests.live.mcpclient import LIVE_CLIENTS, Stdio, ToolFailed
from tests.live import office_cleanup as cleanup


# Быстрый набор (pytest tests/live -m smoke, ~3 мин): по одному-два ключевых сценария на группу плюс тесты, которые уже
# ловили настоящие ошибки Office. Полный набор (~12 мин) — перед релизом.
SMOKE = {
    "test_live_word_captions.py": ("test_caption_table_reference_selection_and_undo", "test_caption_picture_equation_duplicate_and_number"),
    "test_live_word_watermark.py": ("test_probe_watermark_status_leaves_empty_headers_untouched",
                                    "test_probe_watermark_linked_headers_shapes_and_undo"),
    "test_live_save_old_format.py": ("test_xls_save_with_modern_formula_restores_compatibility",),
    "test_live.py": ("test_write_read_and_text_mode", "test_number_formats_roundtrip_in_any_locale", "test_formulas_fill_down_and_r1c1",
                     "test_sort_filter_find_replace_dedupe", "test_conditional_format_and_validation", "test_chart_table_profile_and_issues",
                     "test_pivot_filters_slicers_and_hiding", "test_render_range_image_is_a_real_png", "test_readonly_server_exposes_no_writing_tools",
                     "test_errors_are_actionable", "test_strict_target_requires_explicit_names", "test_word_insert_structure_find_literal_replace",
                     "test_word_layout_headers_toc_comments_bookmarks_revisions", "test_word_tables_stay_separate_and_multiline_cells",
                     "test_bridge_word_table_to_excel_numbers_and_codes", "test_bridge_excel_to_word_table_picture_and_chart"),
    "test_live_demo_fixes.py": ("test_live_mail_merge_shows_no_windows_and_keeps_the_user_document",),
    "test_live_undo.py": ("test_live_write_range_undo", "test_live_word_insert_text_undo"),
    "test_live_links.py": ("test_link_selects_own_workbook_range",),
    "test_live_stage1.py": ("test_live_clean_text_preview_types_formats_and_undo", "test_live_compare_ranges_key_two_workbooks",
                            "test_live_compare_documents_sources_saved_and_temporary_cleanup"),
    "test_live_stage2.py": ("test_live_import_cp1251_new_book", "test_live_import_occupied_and_undo",
                            "test_live_split_preview_literal_types_full_undo_and_conflict", "test_live_trace_cross_sheet_names_table_and_no_ui",
                            "test_live_protection_no_password_inverse_flags_locked"),
    "test_live_office_operations.py": ("test_probe_goal_seek_bool_no_dialog_manual_settings_and_undo",
                                       "test_probe_sparkline_clear_properties_theme_date_axis_roundtrip",
                                       "test_probe_word_protection_inverse_guard_comments_tracking"),
    "test_live_word_review.py": ("test_probe_tracked_fragment_style_numbering_selection_and_undo",
                                 "test_probe_comment_threads_reply_key_date_and_context"),
    "test_live_window.py": ("test_own_excel_and_word_side_by_side",),
}


def pytest_collection_modifyitems(config, items):
    found = set()
    for item in items:
        if item.originalname in SMOKE.get(item.path.name, ()):
            item.add_marker(pytest.mark.smoke)
            found.add((item.path.name, item.originalname))
    collected = {item.path.name for item in items}
    missing = [f"{f}::{n}" for f, names in SMOKE.items() if f in collected for n in names if (f, n) not in found]
    # Explicit nodeids (especially --resume) intentionally collect only part of a file.
    if missing and "smoke" in (config.option.markexpr or "") and not any('::' in arg for arg in config.args):
        raise pytest.UsageError(f"SMOKE names not found (renamed tests?): {missing}")


SESSION = {"watch": None}
WATCH_EVERY, HUNG_FOR, DIALOG_FOR = 5, 45, 20


def _kill_test_servers():
    for client in list(LIVE_CLIENTS):
        try:
            if client.proc.poll() is None:
                client.proc.kill()
        except OSError:
            pass


def _watch_office(stop, watch):
    while not stop.wait(WATCH_EVERY):
        try:
            watch.tick(time.monotonic(), cleanup.office_processes(),
                       cleanup.hung_office_windows(), cleanup.office_dialogs())
        except Exception as exc:
            # Monitoring failure cannot silently disable ownership protection for the rest of the run.
            watch.fatal = f"Сторож Office не может продолжить безопасно: {exc}"
            cleanup.live_event('watchdog_error', reason=watch.fatal)
            _kill_test_servers()
        if watch.fatal:
            print(f"\n[office-watchdog] {watch.fatal}", flush=True)
            return


@pytest.fixture(scope="session", autouse=True)
def _office_watchdog():
    if os.environ.get('OFFICE_LIVE_LIVE_TESTS') != '1':
        yield
        return
    started = time.time_ns() // 1000
    try:
        before = cleanup.office_processes()
        ownership = cleanup.Ownership(started, before)
        cleanup.OWNERSHIP = ownership
        cleanup.live_event('session_start', started=started, before=[cleanup.asdict(p) for p in before])
        cleanup.clean_leftovers(before, ownership)
    except Exception as exc:
        cleanup.live_event('startup_failed', reason=str(exc))
        cleanup.OWNERSHIP = None
        pytest.exit(f'Остатки Office: {exc}', returncode=3)
    watch = cleanup.OfficeWatch(ownership, _kill_test_servers, HUNG_FOR, DIALOG_FOR)
    SESSION['watch'] = watch
    stop = threading.Event()
    thread = threading.Thread(target=_watch_office, args=(stop, watch), daemon=True, name="office-watchdog")
    thread.start()
    try:
        yield
        # Keep the watchdog running through final COM Quit, too.
        if not watch.fatal:
            cleanup.quit_word_if_idle()
            cleanup.quit_excel_if_idle()
        if watch.fatal:
            pytest.exit(watch.fatal, returncode=3)
    finally:
        stop.set()
        thread.join(timeout=WATCH_EVERY + 1)
        cleanup.OWNERSHIP = None


def pytest_sessionfinish(session, exitstatus):
    # Retain recovery evidence through the last test's teardown report (which follows session fixtures).
    SESSION['watch'] = None


@pytest.fixture(autouse=True)
def _stop_when_office_hung(request, _office_watchdog):
    watch = SESSION['watch']
    fatal, recoveries = watch.checkpoint() if watch else ('', [])
    if fatal:
        pytest.exit(fatal, returncode=3)
    request.node._office_recoveries = len(recoveries)
    yield
    fatal, _ = watch.checkpoint() if watch else ('', [])
    if fatal:
        pytest.exit(fatal, returncode=3)


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    report = outcome.get_result()
    watch = SESSION['watch']
    _, recoveries = watch.checkpoint() if watch else ('', [])
    start = getattr(item, '_office_recoveries', 0)
    if len(recoveries) > start and report.outcome == 'passed':
        report.outcome = 'failed'
        report.longrepr = 'Office восстановлен во время теста: ' + '; '.join(recoveries[start:])


@pytest.fixture(autouse=True)
def _fresh_excel_when_bloated(_stop_when_office_hung):
    yield
    watch = SESSION['watch']
    if watch and not watch.fatal:
        cleanup.quit_excel_if_idle(gdi_over=6000, mem_over_mb=900)


_SERVER = {"client": None}


@pytest.fixture
def srv(_office_watchdog):
    # One MCP server per session (startup costs ~1–2 s × ~150 tests); a fresh one only after the watchdog killed it.
    client = _SERVER["client"]
    if client is None or client.proc.poll() is not None:
        if client is not None:
            client.close()
        client = _SERVER["client"] = Stdio()
    return client


@pytest.fixture(scope="session", autouse=True)
def _close_session_server(_office_watchdog):
    # Depends on the watchdog: torn down first, so the server lets go of Office before the final idle Quit.
    yield
    if _SERVER["client"] is not None:
        _SERVER["client"].close()
        _SERVER["client"] = None


# уникальна для этого запуска: файлы других запусков и пользователя не совпадут. Через окружение — потому что модуль
# грузится дважды (pytest как conftest и тесты как tests.live.conftest), а метка должна быть одна.
RUN_TAG = os.environ.setdefault("OFFICE_LIVE_TEST_RUN_TAG", f"ol_pytest_{uuid.uuid4().hex[:8]}")


def _in_run_dir(path) -> bool:
    """Файл лежит во временной папке ЭТОГО запуска (туда тесты делают Save As): путь уникален, чужим быть не может."""
    return bool(path) and RUN_TAG in path


def _full(item) -> str:
    return item["path"] + "\\" + item["name"] if item["path"] else item["name"]


@pytest.fixture
def wb(srv):
    name = srv.call("excel_new_workbook", sheets=["Data"])["workbook"]
    marker = f"{RUN_TAG}_{uuid.uuid4().hex[:8]}"
    # имя несохранённой книги («Книга3») может совпасть с книгой пользователя в другом экземпляре Excel:
    # свою узнаём по уникальной метке, а не по имени
    srv.call("excel_manage_names", workbook=name, action="add", name=marker, formula="=1")
    yield name
    close_own_workbooks(srv, name, marker)


def close_own_workbooks(srv, name, marker):
    """Закрывает (discard) только книги этого теста: файлы в папке запуска и несохранённую книгу с нашей меткой."""
    for w in srv.call("excel_list_workbooks")["workbooks"]:
        if _in_run_dir(w["path"]):
            srv.call("excel_close_workbook", workbook=_full(w), discard=True)
        elif not w["path"] and w["name"] == name and _excel_marked(srv, name, marker):
            srv.call("excel_close_workbook", workbook=name, discard=True)


def _excel_marked(srv, name, marker) -> bool:
    try:  # имя в нескольких экземплярах -> сервер откажет; тогда не закрываем ничего
        names = srv.call("excel_manage_names", workbook=name, action="list")["names"]
    except ToolFailed:
        return False
    return any(n["name"] == marker for n in names)


@pytest.fixture
def doc(srv):
    name = srv.call("word_new_document")["document"]
    marker = f"{RUN_TAG}_{uuid.uuid4().hex[:8]}"
    srv.call("word_document_properties", document=name, action="set", properties={"Comments": marker})
    yield name
    close_own_documents(srv, name, marker)


def close_own_documents(srv, name, marker):
    for d in srv.call("word_list_documents")["documents"]:
        if _in_run_dir(d["path"]):
            srv.call("word_close_document", document=_full(d), discard=True)
        elif not d["path"] and d["name"] == name and _word_marked(srv, name, marker):
            srv.call("word_close_document", document=name, discard=True)


def _word_marked(srv, name, marker) -> bool:
    try:
        props = srv.call("word_document_properties", document=name)["properties"]
    except ToolFailed:
        return False
    return props.get("Comments") == marker


@pytest.fixture
def tmp():
    d = tempfile.mkdtemp(prefix=RUN_TAG + "_")
    yield d
    import shutil

    shutil.rmtree(d, ignore_errors=True)


def assert_real_picture(png: bytes, min_size=(60, 20)):
    """PNG декодируется, имеет разумный размер и не залит одним цветом (пустой снимок вне видимой области)."""
    import io

    from PIL import Image as PILImage

    assert png[:8] == b"\x89PNG\r\n\x1a\n"
    im = PILImage.open(io.BytesIO(png)).convert("L")
    assert im.width >= min_size[0] and im.height >= min_size[1], im.size
    hist = im.histogram()
    assert max(hist) / float(im.width * im.height) < 0.97, "the picture is a single flat colour"


def read(srv, wb, cells, sheet="Data", **kw):
    return srv.call("excel_read_range", workbook=wb, sheet=sheet, cells=cells, **kw)["values"]


