"""Ограниченная по времени проверка stdio MCP, без вызовов инструментов Office."""

import json
import os
import queue
import subprocess
import threading
import time
from collections import deque


class ProtocolError(RuntimeError):
    pass


class StdioClient:
    def __init__(self, command, env=None, cwd=None, timeout=30):
        self.timeout = timeout
        self.proc = subprocess.Popen(command, cwd=cwd, env={**os.environ, **(env or {})},
                                     stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.stderr_lines = deque(maxlen=200)
        self._lines = queue.Queue()
        self._errors = []
        self._threads = [threading.Thread(target=f, daemon=True) for f in (self._drain, self._read_stdout)]
        for thread in self._threads:
            thread.start()
        self._id = 0
        try:
            init = self.rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                                          "clientInfo": {"name": "office-live-check", "version": "1"}})
            self.server_info = init["result"]
            self._send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        except BaseException:
            self.abort()
            raise

    def _drain(self):
        try:
            for line in self.proc.stderr:
                self.stderr_lines.append(line.decode("utf-8", errors="replace").rstrip())
        except OSError as exc:
            self.stderr_lines.append(str(exc))

    def _read_stdout(self):
        try:
            for line in self.proc.stdout:
                try:
                    message = json.loads(line.decode("utf-8", errors="strict"))
                    if not isinstance(message, dict) or message.get("jsonrpc") != "2.0":
                        raise ValueError("ожидался объект JSON-RPC 2.0")
                    if "method" not in message and ("id" not in message or ("result" in message) == ("error" in message)):
                        raise ValueError("ответ должен содержать id и result либо error")
                    self._lines.put(message)
                except (ValueError, UnicodeError) as exc:
                    error = ProtocolError(f"Повреждён stdout MCP (UTF-8 JSON-RPC): {exc}")
                    self._errors.append(error)
                    self._lines.put(error)
        except OSError as exc:
            self._errors.append(exc)
            self._lines.put(exc)
        finally:
            self._lines.put(None)

    def _send(self, msg):
        self.proc.stdin.write((json.dumps(msg) + "\n").encode("utf-8"))
        self.proc.stdin.flush()

    def rpc(self, method, params=None, timeout=None):
        self._id += 1
        rid = self._id
        self._send({"jsonrpc": "2.0", "id": rid, "method": method, **({"params": params} if params is not None else {})})
        deadline = time.monotonic() + (self.timeout if timeout is None else timeout)
        while True:
            try:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise queue.Empty
                item = self._lines.get(timeout=remaining)
            except queue.Empty:
                self.abort()
                raise TimeoutError(f"MCP не ответил на {method} вовремя. stderr: {'; '.join(self.stderr_lines)}") from None
            if isinstance(item, Exception):
                self.abort()
                raise item
            if item is None:
                raise ProtocolError(f"MCP закрыл stdout. stderr: {'; '.join(self.stderr_lines)}")
            if item.get("id") == rid:
                if "error" in item:
                    raise ProtocolError(f"Ошибка MCP {method}: {item['error']}")
                return item

    def list_tools(self):
        result = self.rpc("tools/list", {}).get("result")
        if not isinstance(result, dict) or not isinstance(result.get("tools"), list):
            raise ProtocolError("Ответ tools/list не содержит список tools")
        return result["tools"]

    def abort(self):
        if self.proc.poll() is None:
            self.proc.kill()
        self.proc.wait(timeout=10)
        self._finish()

    def _finish(self):
        for thread in self._threads:
            thread.join(timeout=5)
        for stream in (self.proc.stdin, self.proc.stdout, self.proc.stderr):
            try:
                stream.close()
            except OSError:
                pass

    def close(self):
        try:
            self.proc.stdin.close()
            rc = self.proc.wait(timeout=self.timeout)
        except subprocess.TimeoutExpired:
            self.abort()
            raise TimeoutError("MCP не завершился после закрытия stdin") from None
        self._finish()
        if self._errors:
            raise self._errors[0]
        if rc:
            raise ProtocolError(f"MCP завершился с кодом {rc}. stderr: {'; '.join(self.stderr_lines)}")


def smoke(command: list[str], env=None, on_started=None) -> int:
    """Каталог и протокол должны совпадать в текущем режиме, в том числе readonly."""
    out = subprocess.run([*command, "tools", "--json"], env={**os.environ, **(env or {})}, capture_output=True, timeout=60)
    if out.returncode:
        raise ProtocolError(f"tools --json: код {out.returncode}; stderr: {out.stderr.decode('utf-8', 'replace')[-2000:]}")
    catalog = json.loads(out.stdout.decode("utf-8"))
    expected = {t["name"] for t in catalog if t["registered"]}
    if not expected:
        raise ProtocolError("Каталог инструментов пуст")
    if on_started:
        on_started()
    client = StdioClient([*command, "serve"], env=env)
    try:
        listed = client.list_tools()
        if {t["name"] for t in listed} != expected:
            raise ProtocolError("tools/list не совпадает с каталогом")
        if (env or {}).get("OFFICE_LIVE_MODE", os.environ.get("OFFICE_LIVE_MODE")) == "readonly":
            if any(t["registered"] and t["kind"] not in {"read", "ui", "open"} and not t["read_actions"] for t in catalog):
                raise ProtocolError("В readonly зарегистрированы пишущие инструменты")
    except BaseException:
        client.abort()
        raise
    client.close()
    return len(listed)
