"""MCP без Office: изолированный профиль и только initialize/tools/list."""

import os
import subprocess
import sys
import time

import pytest

from office_live.probe import ProtocolError, StdioClient, smoke
from tests.live.mcpclient import Stdio


@pytest.mark.parametrize("mode", ["full", "readonly"])
def test_source_serve_has_only_protocol_and_exits_on_eof(mode):
    assert smoke([sys.executable, "-m", "office_live"], {"OFFICE_LIVE_MODE": mode, "PYTHONIOENCODING": "cp1251"}) > 10


def fake_server(tmp_path, body):
    script = tmp_path / "server.py"
    script.write_text(body, encoding="utf-8")
    return [sys.executable, "-u", str(script)]


def test_stderr_invalid_utf8_is_drained_without_blocking(tmp_path):
    command = fake_server(tmp_path, '''import sys,json
sys.stderr.buffer.write(b'\\xff diagnostic\\n' * 20000)
sys.stderr.flush()
for line in sys.stdin:
    msg=json.loads(line)
    if 'id' in msg:
        print(json.dumps({'jsonrpc':'2.0','id':msg['id'],'result':{'tools':[]}}), flush=True)
''')
    c = Stdio(command=command, timeout=5)
    assert c.list_tools() == []
    c.close()
    assert any("�" in line for line in c.stderr_lines)


@pytest.mark.parametrize("output", ["b'not JSON\\n'", "b'\\xff\\n'", "b'[]\\n'"])
def test_stdout_corruption_is_an_immediate_error(tmp_path, output):
    cmd = fake_server(tmp_path, f"import sys,time\nsys.stdout.buffer.write({output})\nsys.stdout.flush()\ntime.sleep(10)")
    with pytest.raises(ProtocolError, match="stdout"):
        StdioClient(cmd, timeout=2)


def test_timeout_kills_only_probe_child(tmp_path):
    cmd = fake_server(tmp_path, "import time\ntime.sleep(20)")
    start = time.monotonic()
    with pytest.raises(TimeoutError):
        StdioClient(cmd, timeout=0.3)
    assert time.monotonic() - start < 5


def test_notifications_cannot_extend_deadline(tmp_path):
    cmd = fake_server(tmp_path, '''import json,time
while True:
    print(json.dumps({'jsonrpc':'2.0','method':'notification'}),flush=True)
    time.sleep(.005)
''')
    start = time.monotonic()
    with pytest.raises(TimeoutError):
        StdioClient(cmd, timeout=0.3)
    assert time.monotonic() - start < 5


def test_eof_timeout_is_an_error_not_silent_kill(tmp_path):
    cmd = fake_server(tmp_path, '''import sys,json,time
for line in sys.stdin:
    msg=json.loads(line)
    if 'id' in msg:
        print(json.dumps({'jsonrpc':'2.0','id':msg['id'],'result':{}}),flush=True)
time.sleep(20)
''')
    client = StdioClient(cmd, timeout=1)
    with pytest.raises(TimeoutError, match="stdin"):
        client.close()
    assert client.proc.poll() is not None


def test_serve_config_error_is_utf8_on_stderr_and_stdout_empty():
    env = {**os.environ, "PYTHONIOENCODING": "cp1251", "OFFICE_LIVE_TOOLSETS": "bad-toolset"}
    result = subprocess.run([sys.executable, "-m", "office_live", "serve"], env=env, capture_output=True, timeout=20)
    assert result.returncode != 0 and result.stdout == b""
    assert "Ошибка настройки" in result.stderr.decode("utf-8")
