"""Сборка автономного архива dist/office-live-mcp-<версия>-win64.zip: PyInstaller, проверка exe, упаковка с install.cmd.

Запуск (в окружении с зависимостями проекта и pyinstaller):  python packaging/build.py
"""

import argparse
import shutil
import os
import tempfile
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from office_live import __version__  # noqa: E402


def smoke(exe: Path) -> int:
    from office_live.probe import smoke as check

    with tempfile.TemporaryDirectory(prefix="olp-") as temporary:
        root = Path(temporary)
        env = {key: str(root / part) for key, part in (("USERPROFILE", "home"), ("HOME", "home"),
               ("APPDATA", "roaming"), ("LOCALAPPDATA", "local"), ("CODEX_HOME", "codex"))}
        env.update({key: "" for key in os.environ if key.startswith("OFFICE_LIVE_")})
        env["PYTHONIOENCODING"] = "cp1251"  # frozen обязан сам настроить UTF-8
        count = check([str(exe)], {**env, "OFFICE_LIVE_MODE": "full"})
        readonly = check([str(exe)], {**env, "OFFICE_LIVE_MODE": "readonly"})
        if not 0 < readonly < count:
            raise RuntimeError(f"readonly: {readonly}, full: {count}; ожидалось сокращение каталога")
        return count


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dist", default=str(ROOT / "dist"), help="where to put the archive")
    p.add_argument("--work", default="", help="короткий каталог сборки (по умолчанию временный olb-*)")
    ns = p.parse_args()
    with tempfile.TemporaryDirectory(prefix="olb-") as temporary:
        return build(Path(ns.dist).resolve(), Path(ns.work).resolve() if ns.work else Path(temporary))


def build(dist: Path, work: Path) -> int:
    pyi_out = work / "dist"
    dist.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    for key, part in (("USERPROFILE", "home"), ("HOME", "home"), ("APPDATA", "roaming"),
                      ("LOCALAPPDATA", "local"), ("CODEX_HOME", "codex"), ("PYINSTALLER_CONFIG_DIR", "pyi-cache")):
        env[key] = str(work / "profile" / part)
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--distpath", str(pyi_out), "--workpath", str(work / "work"),
         str(ROOT / "packaging" / "office-live-mcp.spec")],
        check=True, env=env,
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
        z.write(ROOT / "LICENSE", f"{name}/LICENSE.txt")
        for f in sorted(app.rglob("*")):
            if f.is_file():
                z.write(f, f"{name}/app/{f.relative_to(app).as_posix()}")
    shutil.rmtree(pyi_out, ignore_errors=True)
    print(f"{archive}  ({archive.stat().st_size / 1_048_576:.1f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
