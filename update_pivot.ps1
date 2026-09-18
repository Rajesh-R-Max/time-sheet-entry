# Reconfigures the Pivot sheet in the EXISTING Timesheet.xlsx.
# Does not touch entered data. Adds a 'Week Number' calculated column to the
# table if one is not already there, then relays the pivot as:
#   Week Number > Project > Task > Date > Notes > Hours
$ErrorActionPreference = 'Stop'

$f = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx'

$XL_HIDDEN = 0; $XL_ROW = 1; $XL_PAGE = 3; $XL_SUM = -4157
$XL_TABULAR = 1; $XL_REPEAT = 2
$CLR_HDR = 4076605; $CLR_DERIVED = 15921906

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($f)
    $ts = $wb.Worksheets.Item('Timesheet')
    $lo = $ts.ListObjects.Item('tblTimesheet')

    # ---- 1. ensure a Week Number column exists -------------------------------
    $existing = @($lo.ListColumns | ForEach-Object { $_.Name })
    if ($existing -notcontains 'Week Number') {
        $lc = $lo.ListColumns.Add()
        $lc.Name = 'Week Number'
        Write-Output "Added 'Week Number' column."
    } else {
        $lc = $lo.ListColumns.Item('Week Number')
        Write-Output "'Week Number' column already present - refreshing it."
    }
    $idx    = $lc.Index
    $letter = $ts.Cells.Item(1, $idx).Address($false, $false) -replace '\d', ''

    # Must be General BEFORE the formula goes in - a Text-formatted cell stores
    # the formula as a literal string instead of evaluating it.
    $ts.Range("$letter`2:$letter`1000").NumberFormat = 'General'
    $lc.DataBodyRange.Formula = '=IF([@Date]="","",TEXT([@Date],"yyyy")&"-W"&TEXT(ISOWEEKNUM([@Date]),"00"))'

    $hc = $ts.Cells.Item(1, $idx)
    $hc.Font.Bold = $true
    $hc.Interior.Color = $CLR_HDR
    $hc.Font.Color = 16777215
    $hc.HorizontalAlignment = -4108
    $hc.WrapText = $true
    $ts.Range("$letter`2:$letter`1000").Interior.Color = $CLR_DERIVED
    $ts.Columns.Item($letter).ColumnWidth = 12
    Write-Output "Week Number is column $letter; sample value: '$($ts.Cells.Item(2, $idx).Text)'"

    # ---- 2. relay the pivot ---------------------------------------------------
    $pv = $wb.Worksheets.Item('Pivot')
    $pt = $pv.PivotTables('TimesheetPivot')
    $pt.PivotCache().Refresh()

    $pt.ManualUpdate = $true

    $dfs = $pt.DataFields()
    for ($i = $dfs.Count; $i -ge 1; $i--) {
        try { $dfs.Item($i).Orientation = $XL_HIDDEN } catch {}
    }
    $pfs = $pt.PivotFields()
    for ($i = 1; $i -le $pfs.Count; $i++) {
        $p = $pfs.Item($i)
        try { if ($p.Orientation -ne $XL_HIDDEN) { $p.Orientation = $XL_HIDDEN } } catch {}
    }

    $order = @('Week Number','Project','Task','Date','Notes')
    for ($i = 0; $i -lt $order.Count; $i++) {
        $fld = $pt.PivotFields($order[$i])
        $fld.Orientation = $XL_ROW
        $fld.Position = $i + 1
    }

    # Excel auto-groups date row fields into Years/Quarters - undo that so the
    # actual date shows.
    try { $pt.PivotFields('Date').Ungroup() | Out-Null } catch {}
    $rfs = $pt.RowFields()
    for ($i = $rfs.Count; $i -ge 1; $i--) {
        $p = $rfs.Item($i)
        if ($order -notcontains $p.Name) { try { $p.Orientation = $XL_HIDDEN } catch {} }
    }
    try { $pt.PivotFields('Date').NumberFormat = 'dd-mmm-yyyy' } catch {}

    # Month stays as a page filter so you can cut a single month quickly.
    $mf = $pt.PivotFields('Month')
    $mf.Orientation = $XL_PAGE
    $mf.Position = 1

    $df = $pt.AddDataField($pt.PivotFields('Hours'), 'Total Hours', $XL_SUM)
    $df.NumberFormat = '0.00'

    # Flat, report-style layout: one line per entry, labels repeated.
    $pt.RowAxisLayout($XL_TABULAR)
    $pt.RepeatAllLabels($XL_REPEAT)
    $off = @($false) * 12
    foreach ($n in @('Task','Date','Notes')) {
        try { $pt.PivotFields($n).Subtotals = $off } catch {}
    }
    $pt.ColumnGrand = $false

    $pt.ManualUpdate = $false
    $pt.RefreshTable() | Out-Null

    $pv.Range('A1').Value2 = 'Hours by Week / Project / Task / Date'
    $pv.Columns.Item('A').ColumnWidth = 14
    $pv.Columns.Item('B').ColumnWidth = 40
    $pv.Columns.Item('C').ColumnWidth = 22
    $pv.Columns.Item('D').ColumnWidth = 14
    $pv.Columns.Item('E').ColumnWidth = 46
    $pv.Columns.Item('F').ColumnWidth = 10

    $wb.Save()

    $rf = $pt.RowFields(); $pf = $pt.PageFields(); $vf = $pt.DataFields()
    Write-Output "Row fields: $((1..$rf.Count | ForEach-Object { $rf.Item($_).Name }) -join ' > ')"
    Write-Output "Page fields: $((1..$pf.Count | ForEach-Object { $pf.Item($_).Name }) -join ', ')"
    Write-Output "Data fields: $((1..$vf.Count | ForEach-Object { $vf.Item($_).Name }) -join ', ')"
    Write-Output '--- pivot ---'
    $pt.TableRange1.Rows | ForEach-Object {
        Write-Output ('  ' + ((($_.Cells | ForEach-Object { $_.Text }) -join ' | ')))
    }
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
