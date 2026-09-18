[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('dev', 'test', 'prod')]
    [string]$Environment,
    [switch]$ApplySchema
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$config = Get-Content .\wrangler.toml -Raw
$placeholder = switch ($Environment) {
    'dev'  { 'REPLACE_WITH_DEV_D1_DATABASE_ID' }
    'test' { 'REPLACE_WITH_TEST_D1_DATABASE_ID' }
    'prod' { $null }
}
if ($placeholder -and $config.Contains($placeholder)) {
    throw "Configure the $Environment D1 database_id in wrangler.toml before deploying."
}

if ($ApplySchema) {
    $databaseName = if ($Environment -eq 'prod') { 'timesheet-db' } else { "timesheet-db-$Environment" }
    Write-Host "Applying D1 schema to $Environment..." -ForegroundColor Cyan
    npx wrangler d1 execute $databaseName --env $Environment --remote --file .\schema.sql
    if ($LASTEXITCODE -ne 0) { throw 'D1 schema execution failed.' }
}

Write-Host "Deploying Timesheet to Cloudflare environment: $Environment" -ForegroundColor Cyan
npx wrangler deploy --env $Environment
if ($LASTEXITCODE -ne 0) { throw 'Cloudflare deployment failed.' }
