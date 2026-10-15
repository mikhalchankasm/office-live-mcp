# PyInstaller: автономная сборка office-live-mcp.exe (Python и все зависимости внутри, папка onedir).
# Сборка: python packaging/build.py (вызывает pyinstaller с этим файлом).
# ruff: noqa
from PyInstaller.utils.hooks import collect_data_files, collect_submodules
from PyInstaller.utils.win32.versioninfo import (
    FixedFileInfo, StringFileInfo, StringStruct, StringTable, VarFileInfo, VarStruct, VSVersionInfo,
)
from pathlib import Path
import sys

ROOT = SPECPATH + "\\.."
sys.path.insert(0, SPECPATH)
sys.path.insert(0, ROOT)
from office_live import __version__
from versioninfo import numeric_version

release = numeric_version(__version__)
copyright_text = next(line for line in (Path(ROOT) / "LICENSE").read_text(encoding="utf-8").splitlines()
                      if line.startswith("Copyright"))


def exe_version(name, description):
    return VSVersionInfo(
        ffi=FixedFileInfo(filevers=release, prodvers=release, mask=0x3F, flags=0,
                          OS=0x40004, fileType=1, subtype=0, date=(0, 0)),
        kids=[StringFileInfo([StringTable("040904B0", [StringStruct(key, value) for key, value in {
            "FileVersion": __version__, "ProductVersion": __version__, "ProductName": "Office Live MCP",
            "FileDescription": description, "OriginalFilename": name + ".exe", "InternalName": name,
            "CompanyName": "mikhalchankasm", "LegalCopyright": copyright_text,
        }.items()])]), VarFileInfo([VarStruct("Translation", [0x0409, 1200])])],
    )

# mcp.cli нужен только для `mcp dev` и тянет typer — не берём
hidden = collect_submodules("office_live") + collect_submodules("mcp", filter=lambda n: not n.startswith("mcp.cli")) + collect_submodules("mcp_types")
datas = collect_data_files("mcp") + collect_data_files("mcp_types")

a = Analysis(
    [SPECPATH + "\\entry.py"],
    pathex=[ROOT],
    hiddenimports=hidden + ["win32timezone"],
    datas=datas,
    excludes=["tkinter", "pytest", "ruff", "IPython", "matplotlib", "numpy"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="office-live-mcp",
    console=True,  # stdio-сервер MCP и консольный установщик
    upx=False,  # UPX чаще вызывает ложные срабатывания антивирусов
    version=exe_version("office-live-mcp", "Office Live MCP server"),
)
link_exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="office-live-link", console=False, upx=False,
               version=exe_version("office-live-link", "Office Live MCP link handler"))
coll = COLLECT(exe, link_exe, a.binaries, a.datas, name="office-live-mcp", upx=False)
