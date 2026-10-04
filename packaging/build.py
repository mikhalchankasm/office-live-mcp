"""Сборка автономного архива dist/office-live-mcp-<версия>-win64.zip: PyInstaller, проверка exe, упаковка с install.cmd.

Запуск (в окружении с зависимостями проекта и pyinstaller):  python packaging/build.py
"""

import argparse
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from office_live import __version__  # noqa: E402


def smoke(exe: Path) -> int:
    """Собранный exe отдаёт каталог инструментов и отвечает по протоколу MCP. Возвращает число инструментов."""
    out = subprocess.run([str(exe), "tools", "--json"], capture_output=True, timeout=120, check=True)
    catalog = json.loads(out.stdout.decode("utf-8"))
    registered = sum(1 for t in catalog if t["registered"])
    proc = subprocess.Popen([str(exe)], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8")
    try:
        def rpc(rid, method, params=None):
            msg = {"jsonrpc": "2.0", "id": rid, "method": method, **({"params": params} if params is not None else {})}
            proc.stdin.write(json.dumps(msg) + "\n")
            proc.stdin.flush()
            while True:
                line = proc.stdout.readline()
                if not line:
                    raise RuntimeError("the server closed its output: " + proc.stderr.read()[-2000:])
                reply = json.loads(line)
                if reply.get("id") == rid:
                    return reply

        rpc(1, "initialize", {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "build-smoke", "version": "1"}})
        proc.stdin.write(json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}) + "\n")
        proc.stdin.flush()
        listed = rpc(2, "tools/list")["result"]["tools"]
    finally:
        proc.stdin.close()
        proc.wait(30)
    if len(listed) != registered or registered < 50:
        raise RuntimeError(f"the packaged server lists {len(listed)} tools, the catalog says {registered}")
    return len(listed)


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dist", default=str(ROOT / "dist"), help="where to put the archive")
    p.add_argument("--work", default=str(ROOT / "build"), help="PyInstaller work folder")
    ns = p.parse_args()
    dist, work = Path(ns.dist), Path(ns.work)
    pyi_out = dist / "pyinstaller"
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--distpath", str(pyi_out), "--workpath", str(work),
         str(ROOT / "packaging" / "office-live-mcp.spec")],
        check=True,
    )
    app = pyi_out / "office-live-mcp"
    count = smoke(app / "office-live-mcp.exe")
    print(f"smoke test: {count} tools over MCP")

    name = f"office-live-mcp-{__version__}-win64"
    archive = dist / f"{name}.zip"
    archive.unlink(missing_ok=True)
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.write(ROOT / "packaging" / "install.cmd", f"{name}/install.cmd")
        z.write(ROOT / "packaging" / "README.txt", f"{name}/README.txt")
        for f in sorted(app.rglob("*")):
            if f.is_file():
                z.write(f, f"{name}/app/{f.relative_to(app).as_posix()}")
    shutil.rmtree(pyi_out, ignore_errors=True)
    print(f"{archive}  ({archive.stat().st_size / 1_048_576:.1f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
