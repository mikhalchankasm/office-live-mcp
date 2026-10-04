"""Минимальный JSON-RPC клиент stdio для живых тестов: сервер запускается в отдельном процессе."""

import json
import os
import queue
import subprocess
import sys
import threading

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class ToolFailed(Exception):
    """Сервер вернул isError=true; в .text — сообщение для агента."""

    def __init__(self, tool, text):
        super().__init__(f"{tool}: {text}")
        self.tool, self.text = tool, text


class Stdio:
    def __init__(self, env=None, python=None):
        e = dict(os.environ)
        e.update(env or {})
        e["PYTHONIOENCODING"] = "utf-8"
        self.proc = subprocess.Popen(
            [python or sys.executable, "-m", "office_live"],
            cwd=ROOT, env=e, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8",
        )
        self.stderr_lines = []
        self._lines: queue.Queue = queue.Queue()
        threading.Thread(target=self._drain, daemon=True).start()
        threading.Thread(target=self._read_stdout, daemon=True).start()
        self._id = 0
        init = self.rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "live-test", "version": "1"}})
        self.server_info = init["result"]
        self._send({"jsonrpc": "2.0", "method": "notifications/initialized"})

    def _drain(self):
        for line in self.proc.stderr:
            self.stderr_lines.append(line.rstrip())

    def _read_stdout(self):
        for line in self.proc.stdout:
            self._lines.put(line)
        self._lines.put(None)  # конец потока: сервер закрыл stdout

    def _send(self, msg):
        self.proc.stdin.write(json.dumps(msg) + "\n")
        self.proc.stdin.flush()

    def rpc(self, method, params=None, timeout=180.0):
        """Отправляет запрос и ждёт ответ не дольше timeout секунд (зависший сервер не должен вешать весь прогон)."""
        self._id += 1
        rid = self._id
        msg = {"jsonrpc": "2.0", "id": rid, "method": method}
        if params is not None:
            msg["params"] = params
        self._send(msg)
        while True:
            try:
                line = self._lines.get(timeout=timeout)
            except queue.Empty:
                self.proc.kill()
                tail = "\n".join(self.stderr_lines[-20:])
                raise TimeoutError(f"no answer to {method} within {timeout:.0f}s; server stderr:\n{tail}") from None
            if line is None:
                tail = "\n".join(self.stderr_lines[-20:])
                raise RuntimeError(f"server closed stdout; stderr:\n{tail}")
            data = json.loads(line)
            if data.get("id") == rid:
                return data

    def list_tools(self):
        return self.rpc("tools/list", {})["result"]["tools"]

    def raw_call(self, name, args, /):
        resp = self.rpc("tools/call", {"name": name, "arguments": args})
        if "error" in resp:
            raise ToolFailed(name, json.dumps(resp["error"], ensure_ascii=False))
        return resp["result"]

    def call(self, name, /, **args):
        """Возвращает dict из JSON-текста ответа; для картинок — ключ '_images'."""
        result = self.raw_call(name, args)
        texts = [c["text"] for c in result.get("content", []) if c.get("type") == "text"]
        images = [c for c in result.get("content", []) if c.get("type") == "image"]
        if result.get("isError"):
            raise ToolFailed(name, "\n".join(texts))
        data = {}
        if texts:
            try:
                data = json.loads(texts[0])
            except json.JSONDecodeError:
                data = {"text": "\n".join(texts)}
        if images:
            data["_images"] = images
        return data

    def close(self):
        try:
            self.proc.stdin.close()
            self.proc.wait(timeout=20)
        except Exception:
            self.proc.kill()
