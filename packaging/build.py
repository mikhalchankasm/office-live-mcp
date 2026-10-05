"""Сборка setup.exe и запасного win64.zip: PyInstaller, проверка MCP, Inno Setup 6.7+.

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
        result = subprocess.run([str(exe.with_name("office-live-link.exe")), "open-link", "javascript:invalid"],
                                env={**os.environ, **env, "OFFICE_LIVE_LINK_NO_DIALOG": "1"},
                                stdin=subprocess.DEVNULL, capture_output=True, timeout=15)
        if result.returncode == 0:
            raise RuntimeError("GUI link handler accepted an invalid URI")
        return count


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dist", default=str(ROOT / "dist"), help="where to put the archive")
    p.add_argument("--work", default="", help="короткий каталог сборки (по умолчанию временный olb-*)")
    p.add_argument("--no-setup", action="store_true", help="локальная сборка только zip, без Inno Setup")
    ns = p.parse_args()
    with tempfile.TemporaryDirectory(prefix="olb-") as temporary:
        return build(Path(ns.dist).resolve(), Path(ns.work).resolve() if ns.work else Path(temporary), no_setup=ns.no_setup)


def find_iscc() -> Path:
    candidates = [os.environ.get("ISCC", ""),
                  str(Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")) / "Inno Setup 6/ISCC.exe"),
                  str(Path(os.environ.get("LOCALAPPDATA", "")) / "Programs/Inno Setup 6/ISCC.exe"),
                  shutil.which("ISCC.exe") or ""]
    for value in candidates:
        if value and Path(value).is_file():
            return Path(value).resolve()
    raise FileNotFoundError("ISCC.exe не найден. Установите Inno Setup 6.7+ (winget install JRSoftware.InnoSetup), "
                            "задайте ISCC или используйте --no-setup для локальной сборки только zip.")


def build(dist: Path, work: Path, *, no_setup: bool = False) -> int:
    iscc = None if no_setup else find_iscc()  # До подмены LOCALAPPDATA и долгой сборки PyInstaller.
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
    print(f"{archive}  ({archive.stat().st_size / 1_048_576:.1f} MB)")
    if iscc:
        subprocess.run([str(iscc), f"/DAppVersion={__version__}", f"/DSourceDir={app}", f"/O{dist}",
                        str(ROOT / "packaging/office-live-mcp.iss")], check=True)
        setup = dist / f"office-live-mcp-{__version__}-setup.exe"
        print(f"{setup}  ({setup.stat().st_size / 1_048_576:.1f} MB)")
    shutil.rmtree(pyi_out, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
