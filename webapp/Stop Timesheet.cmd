@echo off
rem Stops the timesheet entry server gracefully: the workbook is saved and the
rem hidden Excel instance is closed properly. Use this if you started the
rem server without a visible console window.
title Stop timesheet server

echo Asking the timesheet server to shut down...
echo.

rem A body is required - HTTP.sys rejects a POST with no Content-Length (411).
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-RestMethod -Uri 'http://localhost:8777/api/shutdown' -Method POST -Body '{}' -ContentType 'application/json' -TimeoutSec 5 | Out-Null; Start-Sleep -Seconds 2; Write-Host '  Server stopped and workbook saved.' -ForegroundColor Green } catch { Write-Host '  No server was responding on port 8777 - it may already be stopped.' -ForegroundColor Yellow }"

echo.
pause
