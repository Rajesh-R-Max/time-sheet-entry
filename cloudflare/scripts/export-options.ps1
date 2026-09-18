[CmdletBinding()]
param(
    [string]$Workbook = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx',
    [string]$Output = (Join-Path $PSScriptRoot 'options-export.json')
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
    $sheet = $workbookObject.Worksheets.Item('Lists')
    $columns = @(@{ name = 'category'; col = 3 }, @{ name = 'budget'; col = 4 }, @{ name = 'billable'; col = 5 }, @{ name = 'nonBillableReason'; col = 7 })
    $options = New-Object System.Collections.ArrayList
    foreach ($item in $columns) {
        $last = $sheet.Cells.Item($sheet.Rows.Count, $item.col).End(-4162).Row
        if ($last -lt 2) { continue }
        for ($row = 2; $row -le $last; $row++) {
            $value = ([string]$sheet.Cells.Item($row, $item.col).Value2).Trim()
            if ($value) { [void]$options.Add([ordered]@{ listName = $item.name; value = $value }) }
        }
    }
    $optionArray = @($options | ForEach-Object { [pscustomobject]@{ listName = $_.listName; value = $_.value } } | Sort-Object listName, value -Unique)
    $document = [ordered]@{ source = 'Excel Lists sheet'; exportedAt = [datetime]::UtcNow.ToString('o'); options = $optionArray }
    $json = $document | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($Output, $json, [System.Text.UTF8Encoding]::new($false))
    Write-Host "Exported $($options.Count) option values to $Output" -ForegroundColor Green
}
finally {
    if ($workbookObject) { try { $workbookObject.Close($false) } catch { } }
    if ($excel) { try { $excel.Quit() } catch { } }
    foreach ($comObject in @($workbookObject, $excel)) { if ($comObject) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

