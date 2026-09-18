[CmdletBinding()]
param(
    [string]$Workbook = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx',
    [ValidateSet('dev', 'test', 'prod')]
    [string]$Environment = 'dev',
    [string]$UserName = 'RAJESH',
    [switch]$ReplaceExisting,
    [switch]$PreserveExistingUser
)

$ErrorActionPreference = 'Stop'
if (-not $ReplaceExisting) { throw 'Pass -ReplaceExisting to confirm replacing existing D1 entries for the migrated user.' }
if (-not $PreserveExistingUser -and -not $env:MIGRATION_USER_PASSWORD) { throw 'Set MIGRATION_USER_PASSWORD or pass -PreserveExistingUser.' }
$export = Join-Path $PSScriptRoot 'excel-export.json'
& (Join-Path $PSScriptRoot 'export-excel.ps1') -Workbook $Workbook -Output $export -UserName $UserName
$exportSucceeded = $?
if (-not $exportSucceeded -or -not (Test-Path $export)) { throw 'Excel export failed.' }
$env:CF_ENV = $Environment
$env:MIGRATION_USER = $UserName
if ($PreserveExistingUser) { $env:PRESERVE_EXISTING_USER = '1' }
node (Join-Path $PSScriptRoot 'import-excel.mjs') $export
$exitCode = $LASTEXITCODE
Remove-Item Env:CF_ENV -ErrorAction SilentlyContinue
Remove-Item Env:MIGRATION_USER -ErrorAction SilentlyContinue
Remove-Item Env:PRESERVE_EXISTING_USER -ErrorAction SilentlyContinue
Remove-Item -Path $export -Force -ErrorAction SilentlyContinue
exit $exitCode
