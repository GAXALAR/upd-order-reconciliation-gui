@echo off
setlocal
pushd "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0powershell\upd_editor.ps1"

set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
    echo Application finished with errors.
)

pause
exit /b %EXIT_CODE%