[CmdletBinding()]
param(
    [string]$Workbook = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx',
    [string]$Output = (Join-Path $PSScriptRoot 'excel-export.json'),
    [string]$UserName = 'RAJESH'
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbookObject = $null
try {
    if (-not (Test-Path $Workbook)) { throw "Workbook not found: $Workbook" }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbookObject = $excel.Workbooks.Open((Resolve-Path $Workbook).Path, $false, $true)
    $table = $workbookObject.Worksheets.Item('Timesheet').ListObjects.Item('tblTimesheet')
    $rows = @()
    if ($table.ListRows.Count -gt 0) {
        $grid = $table.DataBodyRange.Value2
        for ($i = 1; $i -le $table.ListRows.Count; $i++) {
            $dateValue = $grid.GetValue($i, 1)
            $hoursValue = $grid.GetValue($i, 6)
            if ($null -eq $dateValue -or $null -eq $hoursValue) { continue }
            $date = if ($dateValue -is [double]) { [datetime]::FromOADate($dateValue).ToString('yyyy-MM-dd') } else { ([datetime]$dateValue).ToString('yyyy-MM-dd') }
            $hours = [double]$hoursValue
            $project = [string]$grid.GetValue($i, 2)
            $budget = [string]$grid.GetValue($i, 10)
            if (-not $project.Trim() -or -not $budget.Trim() -or $hours -le 0 -or $hours -gt 24) { continue }
            $rows += [ordered]@{
                date = $date; project = $project.Trim(); task = [string]$grid.GetValue($i, 3)
                hours = $hours; billable = [string]$grid.GetValue($i, 8); category = [string]$grid.GetValue($i, 9)
                budget = $budget.Trim(); notes = [string]$grid.GetValue($i, 11); ticket = [string]$grid.GetValue($i, 12)
                incidentType = [string]$grid.GetValue($i, 13); nonBillableReason = [string]$grid.GetValue($i, 14)
            }
        }
    }
    [ordered]@{ username = $UserName.ToUpperInvariant(); source = 'Excel tblTimesheet'; exportedAt = [datetime]::UtcNow.ToString('o'); entries = @($rows) } |
        ConvertTo-Json -Depth 5 | Set-Content -Path $Output -Encoding UTF8
    Write-Host "Exported $($rows.Count) valid entries to $Output" -ForegroundColor Green
}
finally {
    if ($workbookObject) { try { $workbookObject.Close($false) } catch { } }
    if ($excel) { try { $excel.Quit() } catch { } }
    foreach ($comObject in @($table, $workbookObject, $excel)) { if ($comObject) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
