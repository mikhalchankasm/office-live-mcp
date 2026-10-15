"""Live harness contracts, entirely fake Office/processes (including timeout tree termination)."""

from datetime import datetime, timezone
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import threading

import pytest

from scripts import live_run as runner
from tests.live import office_cleanup as office


@pytest.fixture(autouse=True)
def no_process_actions(monkeypatch):
    def forbidden(*args, **kwargs):
        raise AssertionError('No native process/Office operations in harness tests')

    monkeypatch.setattr(office.win32api, 'TerminateProcess', forbidden)
    monkeypatch.setattr(office.win32api, 'OpenProcess', forbidden)
    monkeypatch.setattr(office.win32com.client, 'GetObject', forbidden)
    monkeypatch.setattr(office, 'office_processes', forbidden)
    monkeypatch.setattr(runner.subprocess, 'Popen', forbidden)
    monkeypatch.setattr(runner.subprocess, 'run', forbidden)
    monkeypatch.delenv('OFFICE_LIVE_LIVE_EVENTS', raising=False)


def process(pid=10, created=101, command='/automation -Embedding', image='EXCEL.EXE'):
    return office.OfficeProcess(pid, created, image, command)


@pytest.mark.parametrize('candidate,before,owned', [
    (process(), [], True),
    (process(command='"C:\\Program Files\\Office\\EXCEL.EXE" /AUTOMATION -EMBEDDING'), [], True),
    (process(), [process()], False),
    (process(created=99), [], False),
    (process(created=100), [], False),
    (process(command='-Embedding'), [], False),
    (process(command='/automation'), [], False),
    (process(command='/automationish -Embedding'), [], False),
    (process(command='/automation -EmbeddingExtra'), [], False),
    (process(command='excel.exe "C:\\Temp\\ /automation -Embedding file.xlsx"'), [], False),
    (process(command=''), [], False),
])
def test_ownership(candidate, before, owned):
    ownership = office.Ownership(100, before)
    ownership.observe([candidate])
    ownership.observe([candidate])
    assert ownership.owns(candidate) is owned


def test_wmi_snapshot_uses_command_line_and_utc(monkeypatch):
    monkeypatch.setattr(office, 'office_processes', REAL_OFFICE_PROCESSES)
    row = SimpleNamespace(ProcessId=22, CreationDate='20261015123000.123456+180', Name='excel.exe',
                          CommandLine='/AUTOMATION -Embedding')
    service = Mock()
    service.ExecQuery.return_value = [row]
    get = Mock(return_value=service)
    monkeypatch.setattr(office.win32com.client, 'GetObject', get)
    result = office.office_processes()
    get.assert_called_once_with('winmgmts:')
    assert 'Win32_Process' in service.ExecQuery.call_args.args[0]
    assert result[0].created == office.timestamp_us(datetime(2026, 10, 15, 9, 30, 0, 123456, tzinfo=timezone.utc))
    assert result[0].automation


REAL_OFFICE_PROCESSES = office.office_processes


@pytest.mark.parametrize('created,owned,expected', [(101, True, True), (102, True, False), (101, False, False)])
def test_termination_uses_same_handle_and_rejects_reused_pid(created, owned, expected):
    p = process()
    ownership = office.Ownership(100, [])
    if owned:
        ownership.observe([p])
    handle = Mock()
    opened = Mock(return_value=handle)
    get_time = Mock(return_value=created)
    terminate = Mock()
    assert office.terminate_owned(p, ownership, open_process=opened, creation_time=get_time, terminate=terminate) is expected
    if owned:
        opened.assert_called_once_with(0x1001, False, p.pid)
        get_time.assert_called_once_with(handle)
        handle.Close.assert_called_once()
    else:
        opened.assert_not_called()
    if expected:
        terminate.assert_called_once_with(handle, 1)
    else:
        terminate.assert_not_called()


def test_native_termination_adapter_gets_time_through_open_handle(monkeypatch):
    dt = datetime(2026, 10, 15, microsecond=416457, tzinfo=timezone.utc)
    p = process(created=office.timestamp_us(dt))
    owner = office.Ownership(100, [])
    owner.observe([p])

    class Handle:
        Close = Mock()

        def __int__(self):
            return 123456789

    handle = Handle()

    def get_times(raw_handle, creation, exit_time, kernel, user):
        assert raw_handle == int(handle)
        # Preserve sub-millisecond precision and truncate only the final 100ns digit, like WMI.
        ticks = p.created * 10 + 116444736000000000 + 7
        creation._obj.dwLowDateTime = ticks & 0xFFFFFFFF
        creation._obj.dwHighDateTime = ticks >> 32
        return True

    opened, times, terminate = Mock(return_value=handle), Mock(side_effect=get_times), Mock()
    monkeypatch.setattr(office.win32api, 'OpenProcess', opened)
    monkeypatch.setattr(office.ctypes.windll.kernel32, 'GetProcessTimes', times)
    monkeypatch.setattr(office.win32api, 'TerminateProcess', terminate)
    assert office.terminate_owned(p, owner)
    assert times.call_count == 1
    terminate.assert_called_once_with(handle, 1)
    handle.Close.assert_called_once()


@pytest.mark.parametrize('path,expected', [
    (r'C:\Temp\ol_pytest_a\book.xlsx', True),
    (r'c:\TEMP\OL_PYTEST_a\nested\book.xlsx', True),
    (r'C:\Temp\ol_pytest_a\..\private.xlsx', False),
    (r'C:\Other\ol_pytest_a\book.xlsx', False),
    (r'C:\Temp\ol_pytest_a.xlsx', False),
    (r'C:\Temp\elsewhere\ol_pytest_a\book.xlsx', False),
    ('Book1', False), ('', False),
])
def test_leftover_paths(path, expected):
    assert office.is_test_document(path, r'C:\Temp') is expected


def fake_app(paths, quit_action=None):
    docs = [SimpleNamespace(Path=str(Path(path).parent) if path else '', FullName=path, Close=Mock()) for path in paths]
    collection = Mock(side_effect=lambda index: docs[index - 1])
    collection.Count = len(docs)
    app = SimpleNamespace(Workbooks=collection, Documents=collection, Quit=Mock(side_effect=quit_action))
    return app, docs


@pytest.mark.parametrize('paths,image', [([], 'EXCEL.EXE'), (['book.xlsx'], 'EXCEL.EXE'), (['doc.docx'], 'WINWORD.EXE')])
def test_leftover_only_test_documents_closed_without_saving(monkeypatch, tmp_path, paths, image):
    monkeypatch.setattr(office.tempfile, 'gettempdir', lambda: str(tmp_path))
    paths = [str(tmp_path / 'ol_pytest_old' / name) for name in paths]
    p = process(created=90, image=image)
    ownership = office.Ownership(100, [p])
    running = [p]
    app, docs = fake_app(paths, lambda *args: running.clear())
    monkeypatch.setattr(office, 'office_processes', lambda: running[:])
    monkeypatch.setattr(office, '_application_for', lambda p: app)
    office.clean_leftovers([p], ownership)
    assert ownership.owns(p)
    for doc in docs:
        doc.Close.assert_called_once_with(False if image == 'EXCEL.EXE' else 0)
    app.Quit.assert_called_once()


def test_leftover_foreign_document_aborts_without_any_close(monkeypatch, tmp_path):
    p = process(created=90)
    ownership = office.Ownership(100, [p])
    monkeypatch.setattr(office.tempfile, 'gettempdir', lambda: str(tmp_path))
    app, docs = fake_app([str(tmp_path / 'ol_pytest_old' / 'ours.xlsx'), str(tmp_path / 'private.xlsx')])
    monkeypatch.setattr(office, 'office_processes', lambda: [p])
    monkeypatch.setattr(office, '_application_for', lambda p: app)
    terminate = Mock()
    with pytest.raises(RuntimeError, match='чужой документ'):
        office.clean_leftovers([p], ownership, terminate=terminate)
    assert not ownership.owns(p)
    app.Quit.assert_not_called()
    terminate.assert_not_called()
    for doc in docs:
        doc.Close.assert_not_called()


def test_leftover_graceful_quit_then_verified_termination(monkeypatch):
    p = process(created=90)
    ownership = office.Ownership(100, [p])
    monkeypatch.setattr(office, 'office_processes', lambda: [p])
    actions = []

    def clean(p, owner):
        office.check_leftover(p, [], owner)
        actions.append('quit')

    def terminate(p, owner):
        assert owner.owns(p)
        actions.append('terminate')
        return True

    office.clean_leftovers([p], ownership, inspect_and_quit=clean, terminate=terminate, wait_seconds=0.01)
    assert actions == ['quit', 'terminate']


def test_uninspectable_leftover_times_out_and_late_worker_cannot_close(monkeypatch, tmp_path):
    p = process(created=90)
    owner = office.Ownership(100, [p])
    release = threading.Event()
    threads = []
    thread_type = threading.Thread
    app, docs = fake_app([str(tmp_path / 'ol_pytest_old' / 'book.xlsx')])
    monkeypatch.setattr(office.tempfile, 'gettempdir', lambda: str(tmp_path))
    monkeypatch.setattr(office, 'office_processes', lambda: [p])

    def get_app(p):
        release.wait(2)
        return app

    def start_thread(**kwargs):
        thread = thread_type(**kwargs)
        threads.append(thread)
        return thread

    monkeypatch.setattr(office, '_application_for', get_app)
    monkeypatch.setattr(office.threading, 'Thread', start_thread)
    try:
        with pytest.raises(RuntimeError, match='процесс не закрыт'):
            office.clean_leftovers([p], owner, wait_seconds=0.01)
    finally:
        release.set()
        for thread in threads:
            thread.join(timeout=2)
            assert not thread.is_alive()
    assert not owner.owns(p)
    app.Quit.assert_not_called()
    docs[0].Close.assert_not_called()


def test_idle_cleanup_never_attaches_to_foreign_office(monkeypatch):
    before = [process(created=90), process(pid=11, created=101)]
    owner = office.Ownership(100, before)
    monkeypatch.setattr(office, 'OWNERSHIP', owner)
    monkeypatch.setattr(office, 'office_processes', lambda: [*before, process(pid=12, command='-Embedding')])
    attach = Mock()
    monkeypatch.setattr(office, '_application_for', attach)
    office.quit_excel_if_idle()
    office.quit_word_if_idle()
    attach.assert_not_called()


def test_com_binding_rejects_reused_pid_before_graceful_quit(monkeypatch):
    from office_live import com

    p = process()
    snapshots = iter([[p], [process(created=999)]])
    app = SimpleNamespace(Hwnd=123, Quit=Mock())
    monkeypatch.setattr(office, 'office_processes', lambda: next(snapshots))
    monkeypatch.setattr(com, 'apps', lambda kind, launch: [app])
    monkeypatch.setattr(office.win32process, 'GetWindowThreadProcessId', lambda hwnd: (1, p.pid))
    with pytest.raises(RuntimeError, match='оставлен без изменений'):
        office._application_for(p)
    app.Quit.assert_not_called()


@pytest.mark.parametrize('reason,owned', [('hung', True), ('dialog', True), ('hung', False), ('dialog', False)])
def test_watchdog_own_and_foreign(reason, owned):
    p = process(created=101 if owned else 90)
    owner = office.Ownership(100, [] if owned else [p])
    actions, events = [], []
    watch = office.OfficeWatch(owner, lambda: actions.append('mcp'), hung_for=10, dialog_for=20,
                              terminate=lambda p, o: actions.append('office') or True,
                              emit=lambda event, **kw: events.append((event, kw)))
    hung = [(p.image, p.pid, 'Not responding')] if reason == 'hung' else []
    dialogs = [(p.pid, '#32770', 'Recalculate? | OK | Cancel')] if reason == 'dialog' else []
    watch.tick(0, [p], hung, dialogs)
    watch.tick(9, [p], hung, dialogs)
    assert not actions
    watch.tick(21, [p], hung, dialogs)
    if owned:
        assert actions == ['office', 'mcp']
        assert len(watch.recoveries) == 1
        assert not watch.fatal
        assert events[-1][0] == 'recovery'
        watch.tick(45, [p], hung, dialogs)
        assert len(watch.recoveries) == 1
    else:
        assert actions == ['mcp']
        assert 'Чужой Office' in watch.fatal
    if reason == 'dialog':
        assert 'Recalculate? | OK | Cancel' in events[-1][1]['text']


def test_watchdog_requires_continuous_problem_and_stops_at_three_recoveries():
    owner = office.Ownership(100, [])
    terminate, kill = Mock(return_value=True), Mock()
    watch = office.OfficeWatch(owner, kill, hung_for=10, terminate=terminate)
    for i in range(3):
        p = process(pid=10 + i)
        hung = [(p.image, p.pid, 'stuck')]
        watch.tick(i * 100, [p], hung, [])
        watch.tick(i * 100 + 5, [p], [], [])
        watch.tick(i * 100 + 9, [p], hung, [])
        watch.tick(i * 100 + 18, [p], hung, [])
        assert len(watch.recoveries) == i
        watch.tick(i * 100 + 20, [p], hung, [])
        assert bool(watch.fatal) is (i == 2)
    assert terminate.call_count == kill.call_count == 3
    assert 'Третье восстановление' in watch.fatal
    assert all(f'PID {pid}' in watch.fatal for pid in (10, 11, 12))


def test_watchdog_identity_change_resets_timer_and_failed_termination_stops():
    watch = office.OfficeWatch(office.Ownership(100, []), Mock(), hung_for=10, terminate=Mock(return_value=False))
    first, reused = process(), process(created=200)
    hung = [('EXCEL.EXE', 10, 'hung')]
    watch.tick(0, [first], hung, [])
    watch.tick(11, [reused], hung, [])
    assert not watch.fatal
    watch.terminate.assert_not_called()
    watch.tick(22, [reused], hung, [])
    assert watch.fatal
    assert not watch.recoveries


def test_events_append_json_lines(monkeypatch, tmp_path):
    path = tmp_path / 'events.jsonl'
    monkeypatch.setenv('OFFICE_LIVE_LIVE_EVENTS', str(path))
    office.live_event('owned', pid=123)
    office.live_event('recovery', reason='Пересчитать формулы?')
    assert [e['event'] for e in runner.read_events(path)] == ['owned', 'recovery']
    with path.open('a', encoding='utf-8') as stream:
        stream.write('{"interrupted":')
    assert len(runner.read_events(path)) == 2


def test_parse_pytest_output_spaced_params_and_teardown_errors():
    output = '''tests/live/test_a.py::test_a[with spaces] PASSED [ 25%]
tests/live/test_a.py::test_b FAILED [ 50%]
tests/live/test_a.py::test_c SKIPPED (reason) [ 75%]
tests/live/test_a.py::test_a[with spaces] ERROR [100%]
ERROR tests/live/test_a.py::test_d - RuntimeError: setup
FAILED tests/live/test_a.py::test_b - assert False
'''
    assert runner.parse_output(output) == {
        'tests/live/test_a.py::test_a[with spaces]': 'errors', 'tests/live/test_a.py::test_b': 'failed',
        'tests/live/test_a.py::test_c': 'skipped', 'tests/live/test_a.py::test_d': 'errors',
    }


def report(nodeid, when, outcome):
    return dict(event='test_report', nodeid=nodeid, when=when, outcome=outcome)


def test_summary_phase_precedence_and_not_run():
    events = [report('a', 'call', 'passed'), report('a', 'teardown', 'failed'), report('b', 'call', 'failed'),
              report('c', 'setup', 'passed'), report('d', 'setup', 'skipped'), report('e', 'call', 'passed'),
              dict(event='recovery', reason='hung'), dict(event='leftover', process={'pid': 10})]
    result = runner.summarize(list('abcdef'), events, exit_code=1, timed_out=True, duration=2.3456)
    assert result['passed'] == ['e']
    assert result['failed'] == ['b']
    assert result['errors'] == ['a']
    assert result['skipped'] == ['d']
    assert result['not_run'] == ['c', 'f']
    assert len(result['recoveries']) == len(result['leftovers']) == 1
    assert result['timed_out'] and result['exit_code'] == 1 and result['duration_seconds'] == 2.346


def test_resume_latest_summary_and_no_successful_nodes(tmp_path):
    for date, nodes in [('20260101-000000', ['old']), ('20260102-000000', ['a', 'b'])]:
        directory = tmp_path / date
        directory.mkdir()
        (directory / 'summary.json').write_text(json.dumps(dict(failed=nodes, errors=['b', 'c'], not_run=['d'], passed=['success'])))
    nodes = runner.resume_nodeids(tmp_path)
    assert nodes == ['a', 'b', 'c', 'd']
    assert runner.build_pytest_args(['-x'], resume=nodes) == ['a', 'b', 'c', 'd', '-x']
    with pytest.raises(ValueError, match='not additional test paths'):
        runner.build_pytest_args(['tests/live'], resume=nodes)
    with pytest.raises(ValueError, match='No previous'):
        runner.resume_nodeids(tmp_path / 'missing')


def test_pytest_arguments_default_smoke_and_explicit_selection():
    assert runner.build_pytest_args(['-x'], smoke=True) == ['tests/live', '-m', 'smoke', '-x']
    assert runner.build_pytest_args(['tests/live/test_a.py::test_a', '-x']) == ['tests/live/test_a.py::test_a', '-x']
    assert runner.build_pytest_args(['--ignore', 'tests/live/test_a.py']) == ['tests/live', '--ignore', 'tests/live/test_a.py']
    assert runner.build_pytest_args(['--ignore', 'tests/live/test_a.py'], resume=['node']) == ['node', '--ignore', 'tests/live/test_a.py']


def test_runner_timeout_kills_only_its_child_and_reports_progress():
    now = [0]
    proc = SimpleNamespace(pid=123, returncode=None, poll=lambda: None)
    popen = Mock(return_value=proc)
    progress, kills = Mock(), []

    def kill(child):
        assert child is proc
        kills.append(child.pid)
        child.returncode = 1

    def sleep(seconds):
        now[0] += seconds

    assert runner.run_child(['fake-pytest'], {}, None, 61, progress, popen=popen, kill_tree=kill,
                            clock=lambda: now[0], sleep=sleep) == (1, True)
    assert kills == [123]
    progress.assert_called_once()


def test_taskkill_command_is_scoped_to_pytest_child(monkeypatch):
    proc = SimpleNamespace(pid=123, poll=lambda: None, wait=Mock())
    run = Mock()
    monkeypatch.setattr(runner.subprocess, 'run', run)
    runner.kill_pytest_tree(proc)
    assert run.call_args.args[0] == ['taskkill', '/T', '/F', '/PID', '123']
    proc.wait.assert_called_once_with(timeout=15)
    run.reset_mock()
    proc.poll = lambda: 0
    runner.kill_pytest_tree(proc)
    run.assert_not_called()


def test_main_collects_same_selection_and_writes_timeout_summary(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(runner, 'ARTIFACTS', tmp_path)
    commands = []
    nodes = ['tests/live/test_a.py::test_a', 'tests/live/test_a.py::test_b']

    def child(command, env, log, deadline, progress):
        commands.append(command)
        assert env['OFFICE_LIVE_LIVE_TESTS'] == '1'
        path = Path(env['OFFICE_LIVE_LIVE_EVENTS'])
        with path.open('a', encoding='utf-8') as stream:
            if '--collect-only' in command:
                stream.write(json.dumps(dict(event='collection', nodeids=nodes)) + '\n')
                return 0, False
            stream.write(json.dumps(report(nodes[0], 'call', 'passed')) + '\n')
        log.write('fake pytest output\n')
        return 1, True

    monkeypatch.setattr(runner, 'run_child', child)
    assert runner.main(['--smoke', '--timeout-min', '2', '-x']) == 124
    assert commands[0][7:-2] == commands[1][7:-2] == ['tests/live', '-m', 'smoke', '-x']
    assert commands[1][-2:] == ['-v', '-rfE']
    summary = json.loads(next(tmp_path.glob('*/summary.json')).read_text())
    assert summary['passed'] == nodes[:1] and summary['not_run'] == nodes[1:]
    assert summary['timed_out'] and summary['exit_code'] == 1
    assert next(tmp_path.glob('*/pytest.log')).read_text() == 'fake pytest output\n'
    assert 'RESULT passed=1 failed=0 errors=0 not_run=1' in capsys.readouterr().out


def test_resume_nothing_does_not_start_pytest(monkeypatch, tmp_path):
    directory = tmp_path / '20260101-000000'
    directory.mkdir()
    (directory / 'summary.json').write_text(json.dumps(dict(failed=[], errors=[], not_run=[])))
    monkeypatch.setattr(runner, 'ARTIFACTS', tmp_path)
    child = Mock()
    monkeypatch.setattr(runner, 'run_child', child)
    assert runner.main(['--resume']) == 0
    child.assert_not_called()


def test_dialog_child_text_and_hidden_hung_windows(monkeypatch):
    p = process()
    monkeypatch.setattr(office, 'office_processes', lambda: [p])
    monkeypatch.setattr(office.win32gui, 'EnumWindows', lambda callback, arg: [callback(h, arg) for h in (1, 2)])
    monkeypatch.setattr(office.win32gui, 'EnumChildWindows', lambda h, callback, arg: callback(3, arg))
    monkeypatch.setattr(office.win32gui, 'GetWindowText', lambda h: {1: 'Question', 2: 'Hidden Excel', 3: 'Recalculate?'}[h])
    monkeypatch.setattr(office.win32gui, 'GetClassName', lambda h: '#32770' if h == 1 else 'XLMAIN')
    monkeypatch.setattr(office.win32gui, 'IsWindowVisible', lambda h: h == 1)
    monkeypatch.setattr(office.win32process, 'GetWindowThreadProcessId', lambda h: (1, p.pid))
    monkeypatch.setattr(office.ctypes.windll.user32, 'IsHungAppWindow', Mock(side_effect=lambda h: h == 2))
    assert office.office_dialogs() == [(p.pid, '#32770', 'Question | Recalculate?')]
    assert office.hung_office_windows() == [('EXCEL.EXE', p.pid, 'Hidden Excel')]


def test_recovery_marks_current_test_failed_and_next_test_starts_fresh(monkeypatch):
    from tests.live import conftest as live

    watch = SimpleNamespace(recoveries=[], fatal='')
    watch.checkpoint = lambda: (watch.fatal, watch.recoveries[:])
    monkeypatch.setitem(live.SESSION, 'watch', watch)
    request = SimpleNamespace(node=SimpleNamespace())
    fixture = live._stop_when_office_hung.__wrapped__(request, None)
    next(fixture)
    watch.recoveries.append('PID 10: dialog')
    result = SimpleNamespace(outcome='passed')
    hook = live.pytest_runtest_makereport(request.node, None)
    next(hook)
    with pytest.raises(StopIteration):
        hook.send(SimpleNamespace(get_result=lambda: result))
    assert result.outcome == 'failed' and 'PID 10: dialog' in result.longrepr
    with pytest.raises(StopIteration):
        next(fixture)
    request2 = SimpleNamespace(node=SimpleNamespace())
    fixture2 = live._stop_when_office_hung.__wrapped__(request2, None)
    next(fixture2)
    assert request2.node._office_recoveries == 1
    with pytest.raises(StopIteration):
        next(fixture2)
    # One server per session: reused while alive, replaced after the watchdog killed it.
    first, second = Mock(), Mock()
    first.proc.poll.return_value = None
    second.proc.poll.return_value = None
    monkeypatch.setitem(live._SERVER, 'client', None)
    monkeypatch.setattr(live, 'Stdio', Mock(side_effect=[first, second]))
    assert live.srv.__wrapped__(None) is first
    assert live.srv.__wrapped__(None) is first
    first.proc.poll.return_value = -9  # killed by the watchdog
    assert live.srv.__wrapped__(None) is second
    first.close.assert_called_once()
    second.close.assert_not_called()


def test_third_recovery_exits_pytest_at_boundary(monkeypatch):
    from tests.live import conftest as live

    watch = SimpleNamespace(recoveries=['one', 'two', 'three'], fatal='Третье восстановление: one; two; three')
    watch.checkpoint = lambda: (watch.fatal, watch.recoveries[:])
    monkeypatch.setitem(live.SESSION, 'watch', watch)
    fixture = live._stop_when_office_hung.__wrapped__(SimpleNamespace(node=SimpleNamespace()), None)
    with pytest.raises(pytest.exit.Exception, match='Третье восстановление'):
        next(fixture)


def test_smoke_resume_allows_partial_file_collection():
    from tests.live import conftest as live

    item = SimpleNamespace(path=Path('test_live.py'), originalname='test_write_read_and_text_mode', add_marker=Mock())
    config = SimpleNamespace(args=['tests/live/test_live.py::test_write_read_and_text_mode'], option=SimpleNamespace(markexpr='smoke'))
    live.pytest_collection_modifyitems(config, [item])
    item.add_marker.assert_called_once()
    config.args = ['tests/live']
    with pytest.raises(pytest.UsageError, match='SMOKE names not found'):
        live.pytest_collection_modifyitems(config, [item])


class _Names:
    def __init__(self, names):
        self.items = [type("N", (), {"Name": n})() for n in names]
        self.Count = len(self.items)

    def __call__(self, index):
        if isinstance(index, str):
            raise office.pythoncom.com_error(-2147352567, "no such name", None, None)
        return self.items[index - 1]


def test_own_unsaved_recognises_only_marked_test_documents():
    book = type("B", (), {"Names": _Names(["Print_Area", "ol_pytest_1a2b_c3d4"])})()
    assert office.own_unsaved(book, "EXCEL.EXE") is True
    foreign = type("B", (), {"Names": _Names(["Budget"])})()
    assert office.own_unsaved(foreign, "EXCEL.EXE") is False

    def props(value):
        return staticmethod(lambda index: type("P", (), {"Value": value})())

    assert office.own_unsaved(type("D", (), {"BuiltInDocumentProperties": props("ol_pytest_9f_1")})(), "WINWORD.EXE") is True
    assert office.own_unsaved(type("D", (), {"BuiltInDocumentProperties": props("draft")})(), "WINWORD.EXE") is False
    broken = type("D", (), {})()  # unreadable document -> fail closed
    assert office.own_unsaved(broken, "EXCEL.EXE") is False and office.own_unsaved(broken, "WINWORD.EXE") is False


def test_runner_module_is_a_valid_pytest_plugin():
    # live_run loads itself with -p scripts.live_run: every pytest_* name must be a real hook, or pluggy aborts the
    # whole run ("unknown hook"), as the first live run of #25 did.
    import _pytest.hookspec as spec

    hooks = {name for name in dir(spec) if name.startswith("pytest_")}
    assert {name for name in dir(runner) if name.startswith("pytest_")} <= hooks
