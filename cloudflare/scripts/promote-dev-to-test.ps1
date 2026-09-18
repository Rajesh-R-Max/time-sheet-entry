[CmdletBinding()]
param(
    [switch]$ReplaceExisting
)

$ErrorActionPreference = 'Stop'
if (-not $ReplaceExisting) {
    throw 'This refresh replaces Test application data. Re-run with -ReplaceExisting after reviewing the impact.'
}

Set-Location (Split-Path -Parent $PSScriptRoot)
$workspace = Join-Path $PSScriptRoot 'promotion-work'
New-Item -ItemType Directory -Path $workspace -Force | Out-Null

$tables = @(
    'users',
    'projects',
    'project_tasks',
    'team_members',
    'team_activity',
    'holidays',
    'entry_options',
    'timesheet_entries',
    'daily_tasks',
    'reminders'
)

try {
    Write-Host 'Exporting Dev application data...' -ForegroundColor Cyan
    foreach ($table in $tables) {
        $output = Join-Path $workspace "$table.sql"
        npx wrangler d1 export timesheet-db-dev --env dev --remote --table $table --no-schema --output $output
        if ($LASTEXITCODE -ne 0) { throw "Dev export failed for table '$table'." }
    }

    $clear = Join-Path $workspace 'test-clear.sql'
    @"
PRAGMA foreign_keys = OFF;
DELETE FROM timesheet_entries;
DELETE FROM team_activity;
DELETE FROM project_tasks;
DELETE FROM projects;
DELETE FROM team_members;
DELETE FROM holidays;
DELETE FROM entry_options;
DELETE FROM daily_tasks;
DELETE FROM reminders;
DELETE FROM sessions;
DELETE FROM password_reset_tokens;
DELETE FROM users;
PRAGMA foreign_keys = ON;
"@ | Set-Content -Path $clear -Encoding UTF8

    Write-Host 'Clearing Test application data...' -ForegroundColor Yellow
    npx wrangler d1 execute timesheet-db-test --env test --remote --yes --file $clear
    if ($LASTEXITCODE -ne 0) { throw 'Test data cleanup failed.' }

    foreach ($table in $tables) {
        $inputFile = Join-Path $workspace "$table.sql"
        if (Test-Path $inputFile) {
            $content = Get-Content -Path $inputFile -Raw
            if ($content -and $content.Trim().Length -gt 0) {
                Write-Host "Importing $table..." -ForegroundColor Cyan
                npx wrangler d1 execute timesheet-db-test --env test --remote --yes --file $inputFile
                if ($LASTEXITCODE -ne 0) { throw "Test import failed for table '$table'." }
            } else {
                Write-Host "Skipping empty table $table..." -ForegroundColor DarkGray
            }
        }
    }

    Write-Host 'Dev-to-Test promotion completed.' -ForegroundColor Green
}
finally {
    Remove-Item -Path $workspace -Recurse -Force -ErrorAction SilentlyContinue
}
