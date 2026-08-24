@echo off
chcp 65001 >nul
title Settlement Automation
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0settlement.ps1" %*
echo.
echo Press any key to close...
pause >nul
