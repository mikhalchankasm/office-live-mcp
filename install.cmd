@echo off
rem Office Live MCP - one-click installer for Windows. Double-click it or run it from a terminal.
rem Extra arguments go to "python -m office_live setup", e.g.:
rem     install.cmd --yes --clients claude-code,cursor --readonly
rem Set OFFICE_LIVE_NO_PAUSE=1 to skip the final "Press any key" prompt.
setlocal
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
set "RC=0"
echo.
echo === Office Live MCP: installation ===
echo.

rem --- 1. Python 3.10 or newer
set "PY="
for %%C in ("py -3" "python") do (
  if not defined PY (
    %%~C -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1 && set "PY=%%~C"
  )
)
if not defined PY goto :nopython
echo [1/4] Python found: %PY%

rem --- 2. Virtual environment inside this folder (no admin rights needed)
if not exist ".venv\Scripts\python.exe" (
  echo [2/4] Creating virtual environment .venv ...
  %PY% -m venv .venv || goto :fail
) else (
  echo [2/4] Virtual environment .venv already exists
)

rem --- 3. Dependencies
echo [3/4] Installing dependencies - this needs access to PyPI ...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt || goto :fail

rem --- 4. Environment check (informational) and connection to agents (interactive)
echo [4/4] Checking Excel/Word and connecting agents ...
".venv\Scripts\python.exe" -m office_live doctor
".venv\Scripts\python.exe" -m office_live setup %*
if errorlevel 1 goto :fail
goto :end

:nopython
echo.
echo Python 3.10 or newer was not found.
echo Install it from https://www.python.org/downloads/  (tick "Add python.exe to PATH")
echo or run:  winget install Python.Python.3.13
echo then start install.cmd again.
set "RC=1"
goto :end

:fail
echo.
echo Installation did not finish - see the messages above.
set "RC=1"

:end
echo.
if not defined OFFICE_LIVE_NO_PAUSE pause
exit /b %RC%
