@echo off
rem Launches the timesheet entry server and opens the browser.
rem Close this window (or press Ctrl+C) to stop it and save the workbook.
title Timesheet entry server
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
if errorlevel 1 (
  echo.
  echo The server stopped with an error. Read the message above.
  pause
)
