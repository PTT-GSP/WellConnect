@echo off
title Update GSP Well Connect Dashboard
echo ==========================================================
echo           Updating GSP Well Connect Dashboard
echo ==========================================================
echo.
echo Running compilation script...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Compile_Dashboard.ps1"

echo.
echo ==========================================================
echo Done! Please check for any errors above.
echo ==========================================================
pause
