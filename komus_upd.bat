@echo off
pushd "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "upd_editor.ps1"
pause