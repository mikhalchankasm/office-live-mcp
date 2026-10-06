"""Минимальный JSON-RPC клиент stdio для живых тестов: сервер запускается в отдельном процессе."""

import json
import os
import sys
import weakref

from office_live.probe import StdioClient

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


LIVE_CLIENTS: "weakref.WeakSet[Stdio]" = weakref.WeakSet()  # серверы тестов: сторож обрывает их при зависании Office


class ToolFailed(Exception):
    """Сервер вернул isError=true; в .text — сообщение для агента."""

    def __init__(self, tool, text):
        super().__init__(f"{tool}: {text}")
        self.tool, self.text = tool, text


class Stdio(StdioClient):
    def __init__(self, env=None, python=None, command=None, timeout=180):
        executable = python or sys.executable
        command = command or ([executable] if os.path.basename(executable).lower() == "office-live-mcp.exe"
                              else [executable, "-m", "office_live"])
        super().__init__(command, env=env, cwd=ROOT, timeout=timeout)
        LIVE_CLIENTS.add(self)

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
