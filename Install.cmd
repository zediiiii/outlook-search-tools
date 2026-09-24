@echo off
rem Double-click to install FADOE for this Windows user (no admin rights needed).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-FADOE.ps1"
echo.
pause
