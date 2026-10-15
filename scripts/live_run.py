"""Local unattended live-test runner; this module never connects to Office.

Also loaded as a pytest plugin in the child to record exact nodeids and phase outcomes.
"""

import argparse
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import time


ROOT = Path(__file__).resolve().parents[1]
ARTIFACTS = ROOT / 'artifacts' / 'live'
RESULTS = ('passed', 'failed', 'errors', 'skipped')


def _event(event, **fields):
    path = os.environ.get('OFFICE_LIVE_LIVE_EVENTS')
    if path:
        with open(path, 'a', encoding='utf-8') as stream:
            stream.write(json.dumps({'event': event, **fields}, ensure_ascii=False) + '\n')


def pytest_collection_finish(session):
    _event('collection', nodeids=[item.nodeid for item in session.items])


def pytest_runtest_logreport(report):
    _event('test_report', nodeid=report.nodeid, when=report.when, outcome=report.outcome)


def read_events(path):
    if not path.exists():
        return []
    events = []
    for line in path.read_text(encoding='utf-8', errors='replace').splitlines():
        try:
            entry = json.loads(line)
            if isinstance(entry, dict):
                events.append(entry)
        except ValueError:
            pass  # concurrent append / interrupted last line
    return events


def parse_output(text):
    """Fallback for a killed/plugin-less pytest: verbose and -rfE summary lines, including spaced params."""
    results = {}
    text = re.sub(r'\x1b\[[0-9;]*m', '', text)
    for line in text.splitlines():
        match = re.match(r'^(\S.*?::.+?)\s+(PASSED|FAILED|ERROR|SKIPPED|XFAIL|XPASS)(?:\s|$)', line)
        if match:
            nodeid, result = match.groups()
        else:
            match = re.match(r'^(FAILED|ERROR)\s+(.+?)(?:\s+-\s+.*)?$', line)
            if not match:
                continue
            result, nodeid = match.groups()
        category = {'PASSED': 'passed', 'FAILED': 'failed', 'ERROR': 'errors', 'SKIPPED': 'skipped',
                    'XFAIL': 'skipped', 'XPASS': 'passed'}[result]
        _merge_result(results, nodeid, category)
    return results


def _merge_result(results, nodeid, category):
    priority = {'skipped': 0, 'passed': 1, 'failed': 2, 'errors': 3}
    if nodeid not in results or priority[category] > priority[results[nodeid]]:
        results[nodeid] = category


def summarize(collected, events, output='', *, exit_code=None, timed_out=False, duration=0):
    results = parse_output(output)
    for event in events:
        if event.get('event') != 'test_report':
            continue
        phase, outcome = event['when'], event['outcome']
        if outcome == 'failed':
            category = 'failed' if phase == 'call' else 'errors'
        elif outcome == 'skipped':
            category = 'skipped'
        elif phase == 'call':
            category = 'passed'
        else:
            continue  # setup/teardown success alone is not a test result
        _merge_result(results, event['nodeid'], category)
    summary = {key: [nodeid for nodeid in collected if results.get(nodeid) == key] for key in RESULTS}
    summary.update(not_run=[nodeid for nodeid in collected if nodeid not in results],
                   recoveries=[e for e in events if e.get('event') == 'recovery'],
                   leftovers=[e for e in events if e.get('event', '').startswith('leftover')],
                   diagnostics=[e for e in events if e.get('event') in
                                {'startup_failed', 'foreign_office', 'watchdog_error', 'recovery_failed'}],
                   exit_code=exit_code, timed_out=timed_out, duration_seconds=round(duration, 3))
    return summary


def resume_nodeids(artifacts):
    summaries = sorted(artifacts.glob('*/summary.json'), reverse=True)
    if not summaries:
        raise ValueError('No previous artifacts/live/*/summary.json to resume')
    data = json.loads(summaries[0].read_text(encoding='utf-8'))
    return list(dict.fromkeys(nodeid for key in ('failed', 'errors', 'not_run') for nodeid in data[key]))


def build_pytest_args(args, smoke=False, resume=None):
    selections = _test_selections(args)
    if resume is not None:
        # Resume nodeids are the exact selection. Extra options (-k/-m etc.) still apply.
        if selections:
            raise ValueError('--resume accepts pytest options, not additional test paths')
        selection = resume
    else:
        selection = [] if selections else ['tests/live']
    return [*selection, *(['-m', 'smoke'] if smoke else []), *args]


def _test_selections(args):
    # These option values may look exactly like test paths; they are not positional selectors.
    value_options = {'--ignore', '--ignore-glob', '--deselect', '--rootdir', '--confcutdir', '--basetemp',
                     '--junitxml', '--junit-xml', '--log-file', '-c', '--config-file', '-k', '-m', '-o', '--override-ini', '-p'}
    skip, selections = False, []
    for arg in args:
        if skip:
            skip = False
            continue
        if arg in value_options:
            skip = True
            continue
        path = arg.split('::', 1)[0].replace('\\', '/')
        if not arg.startswith('-') and (path.endswith('.py') or path.startswith('tests/') or path in ('tests', '.', './tests')):
            selections.append(arg)
    return selections


def kill_pytest_tree(proc):
    """Only the subprocess object created by this runner is accepted, never an Office PID."""
    if proc.poll() is None:
        subprocess.run(['taskkill', '/T', '/F', '/PID', str(proc.pid)], check=True, capture_output=True)
        proc.wait(timeout=15)


def run_child(command, env, log, deadline, progress, *, popen=None, kill_tree=None, clock=None, sleep=None):
    popen = popen or subprocess.Popen
    kill_tree = kill_tree or kill_pytest_tree
    clock, sleep = clock or time.monotonic, sleep or time.sleep
    proc = popen(command, cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT)
    next_progress = clock() + 60
    while proc.poll() is None:
        now = clock()
        if now >= deadline:
            kill_tree(proc)
            return proc.returncode, True
        if now >= next_progress:
            progress()
            next_progress = now + 60
        sleep(min(1, max(0, deadline - now)))
    return proc.returncode, False


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--smoke', action='store_true')
    parser.add_argument('--resume', action='store_true')
    parser.add_argument('--timeout-min', type=float, default=40)
    options, extra = parser.parse_known_args(argv)
    if not math.isfinite(options.timeout_min) or options.timeout_min <= 0:
        parser.error('--timeout-min must be positive')
    try:
        resume = resume_nodeids(ARTIFACTS) if options.resume else None
        args = build_pytest_args(extra, options.smoke, resume)
    except (ValueError, KeyError, OSError) as exc:
        parser.error(str(exc))
    if resume == []:
        print('RESULT passed=0 failed=0 errors=0 not_run=0 (nothing to resume)')
        return 0
    directory = ARTIFACTS / time.strftime('%Y%m%d-%H%M%S')
    directory.mkdir(parents=True, exist_ok=False)
    event_path, log_path = directory / 'events.jsonl', directory / 'pytest.log'
    event_path.touch()
    env = {**os.environ, 'OFFICE_LIVE_LIVE_TESTS': '1', 'OFFICE_LIVE_LIVE_EVENTS': str(event_path),
           'PYTHONUNBUFFERED': '1', 'PYTHONIOENCODING': 'utf-8', 'PYTEST_ADDOPTS': ''}
    base = [sys.executable, '-m', 'pytest', '-p', 'scripts.live_run', '-o', 'testpaths=tests/live']
    started = time.monotonic()
    deadline = started + options.timeout_min * 60
    collected = list(resume or [])

    def snapshot(**kwargs):
        return summarize(collected, read_events(event_path), log_path.read_text(encoding='utf-8', errors='replace'), **kwargs)

    def progress():
        report = snapshot()
        done = sum(len(report[key]) for key in RESULTS)
        print(f"progress: {done}/{len(collected)} done, {len(report['failed']) + len(report['errors'])} failed", flush=True)

    exit_code, timed_out = 3, False
    try:
        with log_path.open('w', encoding='utf-8') as log:
            # Collection uses the exact same selection/pytest options. Fixtures never execute in this phase.
            exit_code, timed_out = run_child([*base, *args, '--collect-only', '-q'], env, log, deadline, progress)
            inventories = [e['nodeids'] for e in read_events(event_path) if e.get('event') == 'collection']
            if inventories:
                collected = list(dict.fromkeys(inventories[-1]))
            if exit_code == 0 and not timed_out:
                exit_code, timed_out = run_child([*base, *args, '-v', '-rfE'], env, log, deadline, progress)
    finally:
        summary = snapshot(exit_code=exit_code, timed_out=timed_out, duration=time.monotonic() - started)
        (directory / 'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        print(f"RESULT passed={len(summary['passed'])} failed={len(summary['failed'])} "
              f"errors={len(summary['errors'])} not_run={len(summary['not_run'])}", flush=True)
        print(f'Artifacts: {directory}', flush=True)
    return 124 if timed_out else exit_code


if __name__ == '__main__':
    raise SystemExit(main())
