# PyInstaller: автономная сборка office-live-mcp.exe (Python и все зависимости внутри, папка onedir).
# Сборка: python packaging/build.py (вызывает pyinstaller с этим файлом).
# ruff: noqa
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

ROOT = SPECPATH + "\\.."

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
)
link_exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="office-live-link", console=False, upx=False)
coll = COLLECT(exe, link_exe, a.binaries, a.datas, name="office-live-mcp", upx=False)
