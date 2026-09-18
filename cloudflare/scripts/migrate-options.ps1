[CmdletBinding()]
param(
    [string]$Workbook = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx',
    [ValidateSet('dev', 'test', 'prod')]
    [string]$Environment = 'dev',
    [switch]$ReplaceExisting
)

$ErrorActionPreference = 'Stop'
if (-not $ReplaceExisting) { throw 'Pass -ReplaceExisting to confirm replacing the selected D1 entry options.' }
$export = Join-Path $PSScriptRoot 'options-export.json'
& (Join-Path $PSScriptRoot 'export-options.ps1') -Workbook $Workbook -Output $export
if (-not $? -or -not (Test-Path $export)) { throw 'Options export failed.' }
$env:CF_ENV = $Environment
$env:REPLACE_OPTIONS = '1'
node (Join-Path $PSScriptRoot 'import-options.mjs') $export
$exitCode = $LASTEXITCODE
Remove-Item Env:CF_ENV -ErrorAction SilentlyContinue
Remove-Item Env:REPLACE_OPTIONS -ErrorAction SilentlyContinue
Remove-Item -Path $export -Force -ErrorAction SilentlyContinue
exit $exitCode
