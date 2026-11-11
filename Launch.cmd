@echo off
rem Master runner (Windows). Pick Windows or WSL Ubuntu, then a project, then Claude Code opens in it.
rem The root is this folder, wherever it is copied. -NoExit keeps the shell open after Claude exits;
rem -ExecutionPolicy Bypass applies to this process only.
powershell.exe -NoLogo -NoProfile -NoExit -ExecutionPolicy Bypass -File "%~dp0_launcher\Launch.ps1"
