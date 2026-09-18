# Read-back verification of Timesheet.xlsx
$ErrorActionPreference = 'Stop'
$f = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($f)
    Write-Output "Sheets: $(($wb.Worksheets | ForEach-Object { $_.Name }) -join ', ')"

    $ts = $wb.Worksheets.Item('Timesheet')
    $lo = $ts.ListObjects.Item('tblTimesheet')
    Write-Output "Table range: $($lo.Range.Address($false,$false))  rows=$($lo.ListRows.Count)"

    foreach ($r in 2..4) {
        $vals = @()
        foreach ($c in @(1,4,5,6,7,10,15,16,17)) {
            $vals += "$($ts.Cells.Item($r,$c).Text)"
        }
        Write-Output ("row{0}: Date={1} | Start={2} End={3} | Hours={4} Disp={5} | Budget={6} | Day={7} Week={8} Month={9}" -f $r, $vals[0],$vals[1],$vals[2],$vals[3],$vals[4],$vals[5],$vals[6],$vals[7],$vals[8])
    }

    foreach ($c in @('B','C','H','I','J','M','N')) {
        try   { Write-Output "validation $c -> $($ts.Range("$c`5").Validation.Formula1)" }
        catch { Write-Output "validation $c -> MISSING" }
    }

    $pv = $wb.Worksheets.Item('Pivot')
    $pt = $pv.PivotTables('TimesheetPivot')
    $pt.RefreshTable() | Out-Null
    Write-Output "Pivot source: $($pt.SourceData)"
    Write-Output "Pivot range: $($pt.TableRange1.Address($false,$false))"
    $pt.TableRange1.Rows | ForEach-Object { Write-Output ("  | " + (($_.Cells | ForEach-Object { $_.Text }) -join ' | ')) }

    Write-Output "Named ranges: $(($wb.Names | ForEach-Object { $_.Name }) -join ', ')"
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
