# Builds Timesheet.xlsx : entry table with dropdowns + live PivotTable.
# Re-runnable: overwrites the output file. Edit the $lists hashtable below to
# change dropdown contents, then re-run.

$ErrorActionPreference = 'Stop'

$outFile = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx'

# ---------------------------------------------------------------- dropdowns
# Placeholder values - replace with the real option lists from the tool.
$lists = [ordered]@{
    # Projects and Tasks are grown by hand - just type new values down these
    # columns and they appear in the Timesheet dropdowns straight away.
    'Project'              = @('300.006 - KEELE Patching and 9.1 Upgrade')
    'Task'                 = @()
    'Category'             = @('Customer Meeting','Individual work','Internal Meeting','On call/Stand by','Other','Travel','Upskilling')
    'Budget'               = @('Non Billable','Fixed Fee')
    'Billable'             = @('Yes','No')
    'Incident Type'        = @()
    'Non-billable Reason'  = @('Naviam issue / Re-work','Product bug','IBM / Vendor Issue','Client-Caused issue','GCS effort on project','Resource','Handover/Training','Sales disconnect','Other')
}
# Range name used for each list column (must be valid Excel names).
$listNames = @('Projects','Tasks','Categories','Budgets','BillableOpts','IncidentTypes','NonBillableReasons')

# ------------------------------------------------------------ entry columns
$headers = @(
    'Date','Project','Task','Start','End','Hours','Time Tracked','Billable',
    'Category','Budget','Notes','Support Ticket ID','Incident Type',
    'Non-billable Reason','Day','Week Starting','Month'
)
$LAST_ROW = 1000   # rows pre-armed with validation

# Derived-column formulas (structured references; Excel fills them down).
$fHours = '=IF(OR([@Start]="",[@End]=""),"",MOD([@End]-[@Start],1)*24)'
$fDisp  = '=IF([@Hours]="","",INT([@Hours])&"h "&TEXT(ROUND(MOD([@Hours],1)*60,0),"00")&"m")'
$fDay   = '=IF([@Date]="","",TEXT([@Date],"ddd"))'
$fWeek  = '=IF([@Date]="","",[@Date]-WEEKDAY([@Date],3))'
$fMonth = '=IF([@Date]="","",TEXT([@Date],"yyyy-mm"))'

# Sample rows: Date, Project, Task, Start, End, Billable, Category, Budget,
#              Notes, Ticket, IncidentType, NonBillableReason
$samples = @(
    @('2026-08-17','300.006 - KEELE Patching and 9.1 Upgrade','UAT Support','09:00','10:30','No','Individual work','Non Billable','KEELE-UAT-ISSUES-Presentation','','','Other'),
    @('2026-08-17','300.006 - KEELE Patching and 9.1 Upgrade','Development','10:30','13:00','Yes','Individual work','Fixed Fee','9.1 upgrade scripting','','',''),
    @('2026-08-18','300.006 - KEELE Patching and 9.1 Upgrade','Meeting','14:00','15:00','No','Customer Meeting','Non Billable','Daily standup','','','GCS effort on project')
)

# ------------------------------------------------------------------- colours
$CLR_HDR     = 4076605    # dark slate  (BGR)
$CLR_DERIVED = 15921906   # light grey  - calculated columns
$CLR_INPUT   = 16777215   # white       - you type here
$CLR_REQ     = 12648447   # pale amber  - required

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    # ================================================================= Lists
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Lists'
    $col = 1
    foreach ($key in $lists.Keys) {
        $ws.Cells.Item(1, $col).Value2 = $key
        $vals = $lists[$key]
        for ($i = 0; $i -lt $vals.Count; $i++) {
            $ws.Cells.Item($i + 2, $col).Value2 = $vals[$i]
        }
        # Dynamic named range so adding an item extends the dropdown automatically.
        $letter = $ws.Cells.Item(1, $col).Address($false, $false) -replace '\d', ''
        $name   = $listNames[$col - 1]
        $refers = "=OFFSET(Lists!`$$letter`$2,0,0,MAX(1,COUNTA(Lists!`$$letter`:`$$letter`)-1),1)"
        $wb.Names.Add($name, $refers) | Out-Null
        $col++
    }
    $hdr = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1, $lists.Count))
    $hdr.Font.Bold = $true
    $hdr.Interior.Color = $CLR_HDR
    $hdr.Font.Color = 16777215
    $ws.Columns.Item("A:G").ColumnWidth = 26
    $ws.Range('I1').Value2 = 'Type new values down any column here and the matching dropdown on the Timesheet sheet picks them up immediately - no rebuild needed. Keep each list contiguous (no blank rows in the middle). Project and Task start out empty/short on purpose: add yours as you go.'
    $ws.Range('I1').Font.Italic = $true

    # ============================================================= Timesheet
    $ts = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws)
    $ts.Name = 'Timesheet'

    for ($i = 0; $i -lt $headers.Count; $i++) {
        $ts.Cells.Item(1, $i + 1).Value2 = $headers[$i]
    }

    # sample data rows
    $r = 2
    foreach ($s in $samples) {
        $ts.Cells.Item($r, 1).Value2  = [double]([datetime]::ParseExact($s[0],'yyyy-MM-dd',$null).ToOADate())
        $ts.Cells.Item($r, 2).Value2  = $s[1]
        $ts.Cells.Item($r, 3).Value2  = $s[2]
        $ts.Cells.Item($r, 4).Value2  = [double]([datetime]::ParseExact($s[3],'HH:mm',$null).TimeOfDay.TotalDays)
        $ts.Cells.Item($r, 5).Value2  = [double]([datetime]::ParseExact($s[4],'HH:mm',$null).TimeOfDay.TotalDays)
        $ts.Cells.Item($r, 8).Value2  = $s[5]
        $ts.Cells.Item($r, 9).Value2  = $s[6]
        $ts.Cells.Item($r,10).Value2  = $s[7]
        $ts.Cells.Item($r,11).Value2  = $s[8]
        $ts.Cells.Item($r,12).Value2  = $s[9]
        $ts.Cells.Item($r,13).Value2  = $s[10]
        $ts.Cells.Item($r,14).Value2  = $s[11]
        $r++
    }
    $lastData = $r - 1

    # table
    $rng = $ts.Range($ts.Cells.Item(1,1), $ts.Cells.Item($lastData, $headers.Count))
    $lo  = $ts.ListObjects.Add(1, $rng, [System.Reflection.Missing]::Value, 1)  # xlSrcRange, xlYes
    $lo.Name = 'tblTimesheet'
    $lo.TableStyle = 'TableStyleLight9'

    # derived formulas (whole column of the table)
    $lo.ListColumns.Item(6).DataBodyRange.Formula  = $fHours
    $lo.ListColumns.Item(7).DataBodyRange.Formula  = $fDisp
    $lo.ListColumns.Item(15).DataBodyRange.Formula = $fDay
    $lo.ListColumns.Item(16).DataBodyRange.Formula = $fWeek
    $lo.ListColumns.Item(17).DataBodyRange.Formula = $fMonth

    # number formats down the sheet so future rows look right too
    $ts.Range("A2:A$LAST_ROW").NumberFormat = 'dd-mmm-yyyy'
    $ts.Range("D2:E$LAST_ROW").NumberFormat = 'hh:mm'
    $ts.Range("F2:F$LAST_ROW").NumberFormat = '0.00'
    $ts.Range("P2:P$LAST_ROW").NumberFormat = 'dd-mmm-yyyy'
    $ts.Range("Q2:Q$LAST_ROW").NumberFormat = '@'

    # header styling
    $h = $ts.Range("A1:Q1")
    $h.Font.Bold = $true
    $h.Interior.Color = $CLR_HDR
    $h.Font.Color = 16777215
    $h.HorizontalAlignment = -4108
    $h.WrapText = $true
    $ts.Rows.Item(1).RowHeight = 32

    # shade derived + required columns
    foreach ($c in @('F','G','O','P','Q')) {
        $ts.Range("$c`2:$c$LAST_ROW").Interior.Color = $CLR_DERIVED
    }
    $ts.Range("J2:J$LAST_ROW").Interior.Color = $CLR_REQ   # Budget (required)

    # column widths
    $widths = @{ 'A'=13; 'B'=38; 'C'=20; 'D'=8; 'E'=8; 'F'=8; 'G'=12; 'H'=9;
                 'I'=18; 'J'=14; 'K'=42; 'L'=17; 'M'=16; 'N'=24; 'O'=7; 'P'=14; 'Q'=10 }
    foreach ($k in $widths.Keys) { $ts.Columns.Item($k).ColumnWidth = $widths[$k] }

    # data validation  (column letter -> named range)
    $validations = @{
        'B' = 'Projects'; 'C' = 'Tasks'; 'H' = 'BillableOpts'; 'I' = 'Categories'
        'J' = 'Budgets';  'M' = 'IncidentTypes'; 'N' = 'NonBillableReasons'
    }
    foreach ($k in $validations.Keys) {
        $vr = $ts.Range("$k`2:$k$LAST_ROW")
        $vr.Validation.Delete()
        $vr.Validation.Add(3, 1, 1, "=$($validations[$k])") | Out-Null   # xlValidateList, Stop, Between
        $vr.Validation.IgnoreBlank = $true
        $vr.Validation.InCellDropdown = $true
        $vr.Validation.ShowError = $false   # allow typing a value not yet in the list
    }

    $ts.Activate()
    $ts.Range('A2').Select()
    $excel.ActiveWindow.FreezePanes = $false
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.SplitColumn = 1
    $excel.ActiveWindow.FreezePanes = $true

    # ================================================================= Pivot
    $pv = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ts)
    $pv.Name = 'Pivot'

    $pc = $wb.PivotCaches().Create(1, 'tblTimesheet')          # xlDatabase
    $pt = $pc.CreatePivotTable($pv.Range('A5'), 'TimesheetPivot')

    $pt.PivotFields('Month').Orientation   = 3; $pt.PivotFields('Month').Position = 1   # page
    $pt.PivotFields('Week Starting').Orientation = 3; $pt.PivotFields('Week Starting').Position = 2
    $pt.PivotFields('Project').Orientation = 1; $pt.PivotFields('Project').Position = 1 # row
    $pt.PivotFields('Task').Orientation    = 1; $pt.PivotFields('Task').Position = 2
    $pt.PivotFields('Budget').Orientation  = 2; $pt.PivotFields('Budget').Position = 1  # column

    $df = $pt.AddDataField($pt.PivotFields('Hours'), 'Total Hours', -4157)              # xlSum
    $df.NumberFormat = '0.00'

    $pt.RowAxisLayout(1)          # xlTabularRow
    $pt.HasAutoFormat = $false
    $pt.TableStyle2 = 'PivotStyleMedium9'
    $pt.RefreshTable() | Out-Null

    $pv.Range('A1').Value2 = 'Hours by Project / Task'
    $pv.Range('A1').Font.Size = 14
    $pv.Range('A1').Font.Bold = $true
    $pv.Range('A2').Value2 = 'Right-click anywhere in the pivot > Refresh after adding rows on the Timesheet sheet. Drag fields in the PivotTable Fields pane to re-cut the report.'
    $pv.Range('A2').Font.Italic = $true
    $pv.Columns.Item('A').ColumnWidth = 40
    $pv.Columns.Item('B').ColumnWidth = 20

    # ================================================================== save
    $ts.Activate()
    if (Test-Path $outFile) { Remove-Item $outFile -Force }
    $wb.SaveAs($outFile, 51)   # xlOpenXMLWorkbook
    $wb.Close($false)
    Write-Output "OK -> $outFile"
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
