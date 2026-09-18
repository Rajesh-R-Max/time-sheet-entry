@echo off
rem Stops the timesheet entry server. Asks it to shut down first so the workbook
rem is saved and Excel is released; only kills the process if it will not answer.
title Stop timesheet server

echo Stopping Timesheet Server...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "^
  $stopped = $false; ^
  try { ^
    Invoke-RestMethod -Uri 'http://localhost:8777/api/shutdown' -Method POST -Body '{}' -ContentType 'application/json' -TimeoutSec 5 | Out-Null; ^
    Start-Sleep -Seconds 2; ^
    $stopped = $true; ^
    Write-Host '  Server stopped and workbook saved.' -ForegroundColor Green; ^
  } catch { ^
    Write-Host '  No answer on port 8777.' -ForegroundColor Yellow; ^
  }; ^
  if (-not $stopped) { ^
    $hit = $false; ^
    foreach ($p in @(Get-Process powershell -ErrorAction SilentlyContinue)) { ^
      $cmd = (Get-CimInstance Win32_Process -Filter \"ProcessId = $($p.Id)\" -ErrorAction SilentlyContinue).CommandLine; ^
      if ($cmd -like '*server.ps1*') { ^
        $hit = $true; ^
        Write-Host ('  Force-stopping unresponsive server (PID ' + $p.Id + ').') -ForegroundColor Yellow; ^
        Write-Host '  The workbook was NOT saved by the server - check Excel before closing it.' -ForegroundColor Red; ^
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue; ^
      } ^
    }; ^
    if (-not $hit) { Write-Host '  Nothing to stop - the server is not running.' -ForegroundColor DarkGray } ^
  }^
"
pause
