@echo off
setlocal
pushd "%~dp0"

set "PYTHON_CMD="
where py.exe >nul 2>nul
if not errorlevel 1 (
    py -3 -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
    where python.exe >nul 2>nul
    if not errorlevel 1 (
        python -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>nul
        if not errorlevel 1 set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo Python 3.9 or newer is required to run the primary application.
    echo Install Python and dependencies with: python -m pip install -r python\requirements.txt
    set "EXIT_CODE=9009"
    goto :finish
)

%PYTHON_CMD% "%~dp0python\upd_editor.py"

set "EXIT_CODE=%ERRORLEVEL%"
:finish
popd

if not "%EXIT_CODE%"=="0" (
    echo Application finished with errors.
    echo If a Python package is missing, install dependencies with:
    echo python -m pip install -r python\requirements.txt
)

pause
exit /b %EXIT_CODE%