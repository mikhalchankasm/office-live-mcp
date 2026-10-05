@echo off
rem Office Live MCP - installer of the packaged build (Python is inside, nothing else is needed).
rem Extract the whole archive, then double-click this file.
rem Extra arguments go to "office-live-mcp.exe install", e.g.:
rem     install.cmd --yes --clients claude-code,cursor --readonly
rem Set OFFICE_LIVE_NO_PAUSE=1 to skip the final "Press any key" prompt.
setlocal DisableDelayedExpansion
cd /d "%~dp0"
set "RC=0"
if not exist "app\office-live-mcp.exe" goto :notextracted
set "OFFICE_LIVE_INTERNAL=%~dp0app\_internal"
if not "%OFFICE_LIVE_INTERNAL:~200,1%"=="" echo WARNING: Extract the archive into a shorter folder, for example Downloads, to avoid Windows path length errors.
"app\office-live-mcp.exe" install %*
set "RC=%errorlevel%"
goto :end

:notextracted
echo.
echo The program files were not found next to install.cmd.
echo Extract the WHOLE archive first (right click - Extract All...), then run install.cmd from the extracted folder.
set "RC=1"

:end
echo.
if not defined OFFICE_LIVE_NO_PAUSE pause
exit /b %RC%
