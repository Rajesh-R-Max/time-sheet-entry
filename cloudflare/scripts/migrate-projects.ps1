[CmdletBinding()]
param(
    [string]$Workbook = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx',
    [ValidateSet('dev', 'test', 'prod')]
    [string]$Environment = 'dev',
    [switch]$ReplaceExisting
)

$ErrorActionPreference = 'Stop'
if (-not $ReplaceExisting) { throw 'Pass -ReplaceExisting to confirm replacing existing Cloudflare projects and tasks.' }
$export = Join-Path $PSScriptRoot 'projects-export.json'
& (Join-Path $PSScriptRoot 'export-projects.ps1') -Workbook $Workbook -Output $export
if (-not $? -or -not (Test-Path $export)) { throw 'Project export failed.' }
$env:CF_ENV = $Environment
$env:REPLACE_PROJECTS = '1'
node (Join-Path $PSScriptRoot 'import-projects.mjs') $export
$exitCode = $LASTEXITCODE
Remove-Item Env:CF_ENV -ErrorAction SilentlyContinue
Remove-Item Env:REPLACE_PROJECTS -ErrorAction SilentlyContinue
Remove-Item -Path $export -Force -ErrorAction SilentlyContinue
exit $exitCode
