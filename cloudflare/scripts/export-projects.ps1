[CmdletBinding()]
param(
    [string]$Workbook = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx',
    [string]$Output = (Join-Path $PSScriptRoot 'projects-export.json')
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbookObject = $null
$table = $null
try {
    if (-not (Test-Path $Workbook)) { throw "Workbook not found: $Workbook" }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbookObject = $excel.Workbooks.Open((Resolve-Path $Workbook).Path, $false, $true)
    $sheet = $workbookObject.Worksheets.Item('Projects')
    $last = $sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
    $projects = [ordered]@{}
    if ($last -ge 2) {
        $grid = $sheet.Range($sheet.Cells.Item(2, 1), $sheet.Cells.Item($last, 3)).Value2
        for ($i = 1; $i -le ($last - 1); $i++) {
            $name = ([string]$grid.GetValue($i, 1)).Trim()
            if (-not $name) { continue }
            $key = $name.ToLowerInvariant()
            if (-not $projects.Contains($key)) {
                $status = ([string]$grid.GetValue($i, 2)).Trim()
                if (@('New', 'In-Progress', 'Hold', 'Complete') -notcontains $status) { $status = 'New' }
                $projects[$key] = [ordered]@{ name = $name; status = $status; tasks = [System.Collections.ArrayList]::new() }
            }
            $task = ([string]$grid.GetValue($i, 3)).Trim()
            if ($task -and -not $projects[$key].tasks.Contains($task)) { [void]$projects[$key].tasks.Add($task) }
        }
    }
    [ordered]@{ source = 'Excel Projects sheet'; exportedAt = [datetime]::UtcNow.ToString('o'); projects = @($projects.Values) } |
        ConvertTo-Json -Depth 5 | Set-Content -Path $Output -Encoding UTF8
    Write-Host "Exported $($projects.Count) projects to $Output" -ForegroundColor Green
}
finally {
    if ($workbookObject) { try { $workbookObject.Close($false) } catch { } }
    if ($excel) { try { $excel.Quit() } catch { } }
    foreach ($comObject in @($table, $workbookObject, $excel)) { if ($comObject) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
