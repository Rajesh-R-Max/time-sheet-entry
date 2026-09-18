# =====================================================================
#  Timesheet web entry - local server
#  PowerShell 5.1 + System.Net.HttpListener + Excel COM
#
#  Excel is the database: every write goes through the tblTimesheet
#  ListObject so the PivotTable, formulas and validation all survive.
#
#  Start with:  .\server.ps1        (or double-click "Start Timesheet.cmd")
#  Stop with:   Ctrl+C
# =====================================================================
[CmdletBinding()]
param(
    [int]    $Port     = 8777,
    [string] $Workbook,
    [switch] $NoBrowser
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot is not reliably populated while param defaults are evaluated,
# so resolve everything relative to the script here instead.
$script:ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Workbook) { $Workbook = 'C:\Users\RajeshRamannavar\OneDrive - Naviam\Documents\Office\Project\TimeSheet.xlsx' }

$script:WWW     = Join-Path $script:ROOT 'www'
$script:JOURNAL = Join-Path $script:ROOT 'journal.jsonl'
$script:USERS_FILE = Join-Path $script:ROOT 'users.json'
$script:SHEET   = 'Timesheet'
$script:TABLE   = 'tblTimesheet'
$script:PIVOTWS = 'Pivot'
$script:PIVOT   = 'TimesheetPivot'
$script:RECENT  = 20

# Master project list. One row per project/task pair, Status repeated on each of
# a project's rows so the sheet still reads correctly on its own in Excel.
$script:PROJWS       = 'Projects'
$script:PROJ_STATUS  = @('New', 'In-Progress', 'Hold', 'Complete')
$script:PROJ_ACTIVE  = 'In-Progress'

# Team activity tracking. WFH is booked as whole Mon-Fri weeks, four per person
# per year, so a week is stored by its Monday and counted as one against the
# allowance - no day arithmetic needed anywhere.
$script:TEAMWS       = 'Team'
$script:TEAMACTWS    = 'TeamActivity'
$script:TEAM_STATUS  = @('Active', 'Inactive')
$script:TEAM_SEED    = @('Ramya', 'Bhagya', 'Sunil', 'Pallavi', 'Prajwal')
$script:TEAM_ACTS    = @('WFH')
$script:TEAM_ALLOW   = 4

# Task board. The sheet is the source of truth; the journal only ever seeds it
# the first time, for boards that predate the sheet.
$script:TASKWS      = 'Tasks'
$script:TASK_HEAD   = @('Id', 'Title', 'Status', 'Priority', 'Seq', 'Raised', 'Created', 'Completed')
$script:TASK_COLS   = 8

# Public holidays. The sheet is the source of truth once it exists - this list
# only seeds it the first time, so edits in Excel are picked up on next load.
$script:HOLWS     = 'Holidays'
$script:HOL_SEED  = @(
    @{ date = '2026-01-01'; name = 'New Year' },
    @{ date = '2026-01-26'; name = 'Republic Day' },
    @{ date = '2026-03-19'; name = 'Ugadi' },
    @{ date = '2026-04-03'; name = 'Good Friday' },
    @{ date = '2026-05-01'; name = "Labour's Day" },
    @{ date = '2026-09-14'; name = 'Ganesh Chaturthi' },
    @{ date = '2026-10-02'; name = 'Gandhi Jayanthi' },
    @{ date = '2026-10-20'; name = 'Vijaya Dashami' },
    @{ date = '2026-11-09'; name = 'Deepavali' },
    @{ date = '2026-12-25'; name = 'Christmas' }
)

# Only these may POST. A page on any other site can still reach localhost, so
# without this a tab you happen to have open could write into your timesheet.
$script:ORIGINS = @("http://localhost:$Port", "http://127.0.0.1:$Port")

# The Date column is formatted dd-mmm-yyyy in the workbook; Get-Entries reads
# raw values in bulk, so it has to reproduce that itself.
$script:DATE_FORMAT = 'dd-MMM-yyyy'

# tblTimesheet column indexes. 4/5 (Start/End) and 7,15-18 (derived) are
# deliberately never written - Excel fills the derived ones from formulas.
$script:COL = @{
    Date = 1; Project = 2; Task = 3; Start = 4; End = 5; Hours = 6
    TimeTracked = 7; Billable = 8; Category = 9; Budget = 10; Notes = 11
    Ticket = 12; IncidentType = 13; NonBillableReason = 14
    Day = 15; WeekStarting = 16; Month = 17; WeekNumber = 18; User = 19
}
# Lists sheet: one column per pick list, header in row 1, values from row 2 down.
$script:LIST_COL = @{
    project  = 1; task         = 2; category          = 3; budget = 4
    billable = 5; incidentType = 6; nonBillableReason = 7
}
# Written by the form. Deliberately an array of pairs, not [ordered]@{2='project'}:
# indexing an OrderedDictionary with an int is a POSITIONAL lookup, not a key
# lookup, which silently writes values into the wrong columns.
$script:WRITE_MAP = @(
    @{ Col = 2;  Key = 'project'           }
    @{ Col = 3;  Key = 'task'              }
    @{ Col = 8;  Key = 'billable'          }
    @{ Col = 9;  Key = 'category'          }
    @{ Col = 10; Key = 'budget'            }
    @{ Col = 11; Key = 'notes'             }
    @{ Col = 12; Key = 'ticket'            }
    @{ Col = 13; Key = 'incidentType'      }
    @{ Col = 14; Key = 'nonBillableReason' }
)

# --------------------------------------------------------------- Excel COM
$script:Excel        = $null
$script:Wb           = $null
$script:OwnsExcel    = $false   # we started EXCEL.EXE, so we may quit it
$script:OwnsWorkbook = $false   # we opened the file, so we must close it
$script:ExcelPid     = 0

function Disable-AutoSave {
    # Co-authoring cannot merge COM-driven structural edits (new sheets, formats).
    try {
        if ($script:Wb.AutoSaveOn) {
            $script:Wb.AutoSaveOn = $false
            Write-Host "  Excel      : AutoSave turned off (avoids OneDrive merge conflicts)" -ForegroundColor DarkYellow
        }
    } catch { }
}

function Connect-Excel {
    <#  Reuse a running Excel that already has the workbook open, so the form
        keeps working while the file is on screen. Otherwise open our own
        hidden instance. #>
    $full = (Resolve-Path $Workbook).Path

    try {
        $running = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($w in $running.Workbooks) {
            if ($w.FullName -eq $full) {
                $script:Excel = $running; $script:Wb = $w
                $script:OwnsExcel = $false; $script:OwnsWorkbook = $false
                Disable-AutoSave
                Write-Host "  Excel      : attached to already-open workbook" -ForegroundColor DarkGray
                return
            }
        }
        # Excel is running but does not have our file - open it there.
        $script:Excel = $running
        $script:Wb = $running.Workbooks.Open($full)
        if ($null -eq $script:Wb) { throw "Excel opened the workbook but returned a null object: $full" }
        $script:OwnsExcel = $false; $script:OwnsWorkbook = $true
        Disable-AutoSave
        try { $script:Excel.Visible = $true; $script:Excel.WindowState = -4137; $script:Wb.Activate() } catch { }
        Write-Host "  Excel      : opened workbook in running Excel" -ForegroundColor DarkGray
        return
    } catch { }

    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    $script:Excel = New-Object -ComObject Excel.Application
    $new = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    if ($new.Count) { $script:ExcelPid = $new[0].Id }
    $script:Excel.Visible = $true
    $script:Excel.DisplayAlerts = $false
    $script:Excel.AutomationSecurity = 1   # msoAutomationSecurityLow - bypass Protected View / macro prompts
    $script:Wb = $script:Excel.Workbooks.Open($full)
    if ($null -eq $script:Wb) { throw "Excel opened the workbook but returned a null object: $full" }
    $script:OwnsExcel = $true; $script:OwnsWorkbook = $true
    Disable-AutoSave
    try { $script:Excel.WindowState = -4137; $script:Wb.Activate() } catch { }
    Write-Host "  Excel      : opened workbook" -ForegroundColor DarkGray
}

function Disconnect-Excel {
    try { if ($script:Wb) { $script:Wb.Save() } } catch { }

    # Leave the instance alone if the user has other files open in it.
    $solo = $true
    try { $solo = ($script:Excel.Workbooks.Count -le 1) } catch { }

    if ($script:OwnsWorkbook) { try { $script:Wb.Close($false) } catch { } }
    if ($script:OwnsExcel -and $solo) { try { $script:Excel.Quit() } catch { } }

    # Get-Table and friends mint a fresh wrapper on every request, and each one
    # pins EXCEL.EXE. Quit only takes effect once they have all been collected.
    $script:Wb = $null; $script:Excel = $null
    for ($i = 0; $i -lt 2; $i++) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }

    if ($script:OwnsExcel -and $solo -and $script:ExcelPid) {
        $p = Get-Process -Id $script:ExcelPid -ErrorAction SilentlyContinue
        if ($p) {
            if (-not $p.WaitForExit(5000)) {
                Write-Host '  Excel      : did not exit, closing it' -ForegroundColor DarkYellow
                Stop-Process -Id $script:ExcelPid -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Get-Table { $script:Wb.Worksheets.Item($script:SHEET).ListObjects.Item($script:TABLE) }

function Initialize-UserColumn {
    $lo = Get-Table
    $changed = $false
    if ($lo.ListColumns.Count -lt $script:COL.User) {
        $column = $lo.ListColumns.Add()
        $column.Name = 'User'
        $changed = $true
    } elseif ([string]$lo.ListColumns.Item($script:COL.User).Name -ne 'User') {
        $lo.ListColumns.Item($script:COL.User).Name = 'User'
        $changed = $true
    }

    if ($lo.ListRows.Count -gt 0) {
        $userRange = $lo.DataBodyRange.Columns.Item($script:COL.User)
        for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
            if (-not ([string]$userRange.Cells.Item($i, 1).Value2).Trim()) {
                $userRange.Cells.Item($i, 1).Value2 = 'RAJESH'
                $changed = $true
            }
        }
    }
    if ($changed) { $script:Wb.Save() }
}

function Update-Pivot {
    try { $script:Wb.Worksheets.Item($script:PIVOTWS).PivotTables($script:PIVOT).PivotCache().Refresh() } catch { }
}

# --------------------------------------------------------------- reading
function Get-ListColumnLast {
    <# Last non-empty row in a Lists column. Row 1 is the header, so anything
       below 2 means the column holds no values yet. #>
    param($Ws, [int]$Col)
    $r = $Ws.Cells.Item($Ws.Rows.Count, $Col).End(-4162).Row      # -4162 = xlUp
    if ($r -lt 1) { 1 } else { $r }
}

function Get-ListSheetOptions {
    <# Curated option lists from the Lists sheet, in the order kept there. Read
       to the real end of each column rather than a fixed row, so lists can grow
       past whatever was there when the workbook was built. Pulled as one block
       for the same reason Get-Entries is: this runs on every request, writes
       included. #>
    $ws   = $script:Wb.Worksheets.Item('Lists')
    $wide = ($script:LIST_COL.Values | Measure-Object -Maximum).Maximum

    $out = @{}
    foreach ($key in @($script:LIST_COL.Keys)) { $out[$key] = New-Object System.Collections.ArrayList }

    $last = 1
    foreach ($c in $script:LIST_COL.Values) {
        $r = Get-ListColumnLast $ws $c
        if ($r -gt $last) { $last = $r }
    }
    if ($last -lt 2) { return $out }

    $grid = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, $wide)).Value2
    foreach ($key in @($script:LIST_COL.Keys)) {
        $c = $script:LIST_COL[$key]
        for ($i = 1; $i -le $last - 1; $i++) {
            $v = $grid.GetValue($i, $c)
            if ($null -eq $v) { continue }
            $t = ([string]$v).Trim()
            if ($t) { [void]$out[$key].Add($t) }
        }
    }
    $out
}

function Add-Option {
    <# Append values to the Lists sheet so they appear in the pickers straight
       away, instead of only after time has been booked against them. Matching
       is case-insensitive, so a value already there is reported back as skipped
       rather than added a second time under different casing. #>
    param($Items)
    if (-not $Items) { throw 'Nothing to add.' }

    Write-Journal -Payload $Items -Action 'option'

    $ws      = $script:Wb.Worksheets.Item('Lists')
    $added   = New-Object System.Collections.ArrayList
    $skipped = New-Object System.Collections.ArrayList

    foreach ($it in $Items) {
        $list = [string]$it.list
        $val  = ([string]$it.value).Trim()
        if (-not $val) { continue }
        if (-not $script:LIST_COL.ContainsKey($list)) { throw "Unknown list '$list'." }
        if ($val.Length -gt 120) { throw "'$val' is too long for a list entry." }

        $c      = $script:LIST_COL[$list]
        $last   = Get-ListColumnLast $ws $c
        $exists = $false
        if ($last -ge 2) {
            foreach ($r in 2..$last) {
                if (([string]$ws.Cells.Item($r, $c).Text).Trim() -eq $val) { $exists = $true; break }
            }
        }
        if ($exists) { [void]$skipped.Add(@{ list = $list; value = $val }); continue }

        $ws.Cells.Item(([math]::Max($last, 1) + 1), $c).Value2 = $val
        [void]$added.Add(@{ list = $list; value = $val })
    }

    if ($added.Count) { $script:Wb.Save() }
    @{ added = @($added); skipped = @($skipped) }
}

# --------------------------------------------------------------- projects
function Get-ProjectSheet {
    foreach ($w in $script:Wb.Worksheets) { if ($w.Name -eq $script:PROJWS) { return $w } }
    $null
}

function Initialize-ProjectSheet {
    <# Seeded once from whatever is already in the workbook, so an existing
       timesheet does not have to be retyped into the new sheet. #>
    if (Get-ProjectSheet) { return }

    $ws = $script:Wb.Worksheets.Add([System.Reflection.Missing]::Value,
                                    $script:Wb.Worksheets.Item($script:Wb.Worksheets.Count))
    $ws.Name = $script:PROJWS
    $ws.Cells.Item(1, 1).Value2 = 'Project'
    $ws.Cells.Item(1, 2).Value2 = 'Status'
    $ws.Cells.Item(1, 3).Value2 = 'Task'
    $ws.Rows.Item(1).Font.Bold = $true

    $seed = [ordered]@{}
    foreach ($v in (Get-ListSheetOptions).project) {
        if ($v -and -not $seed.Contains($v)) { $seed[$v] = New-Object System.Collections.ArrayList }
    }
    foreach ($e in (Get-Entries)) {
        if (-not $e.project) { continue }
        if (-not $seed.Contains($e.project)) { $seed[$e.project] = New-Object System.Collections.ArrayList }
        if ($e.task -and -not $seed[$e.project].Contains($e.task)) { [void]$seed[$e.project].Add($e.task) }
    }

    $r = 2
    foreach ($name in $seed.Keys) {
        $tasks = @($seed[$name])
        if (-not $tasks.Count) { $tasks = @('') }
        foreach ($t in $tasks) {
            $ws.Cells.Item($r, 1).Value2 = $name
            $ws.Cells.Item($r, 2).Value2 = $script:PROJ_ACTIVE
            if ($t) { $ws.Cells.Item($r, 3).Value2 = $t }
            $r++
        }
    }
    $ws.Columns.Item(1).ColumnWidth = 46
    $ws.Columns.Item(2).ColumnWidth = 14
    $ws.Columns.Item(3).ColumnWidth = 46
    $script:Wb.Save()
    Write-Host "  Projects   : sheet created, seeded with $($seed.Count) project(s)" -ForegroundColor DarkGray
}

function Get-Projects {
    <# Sheet order is preserved. A project's status is whatever its first row
       says, so a half-edited column cannot split one project into two. #>
    $out = New-Object System.Collections.ArrayList
    $ws  = Get-ProjectSheet
    if (-not $ws) { return $out }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -lt 2) { return $out }

    $grid = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 3)).Value2
    $idx  = @{}
    for ($i = 1; $i -le $last - 1; $i++) {
        $name = ([string]$grid.GetValue($i, 1)).Trim()
        if (-not $name) { continue }
        $key = $name.ToLower()
        if (-not $idx.ContainsKey($key)) {
            $status = ([string]$grid.GetValue($i, 2)).Trim()
            if ($script:PROJ_STATUS -notcontains $status) { $status = 'New' }
            $idx[$key] = [ordered]@{ name = $name; status = $status; tasks = New-Object System.Collections.ArrayList }
            [void]$out.Add($idx[$key])
        }
        $task = ([string]$grid.GetValue($i, 3)).Trim()
        if ($task -and -not $idx[$key].tasks.Contains($task)) { [void]$idx[$key].tasks.Add($task) }
    }
    $out
}

function Find-Project {
    param($Projects, [string]$Name)
    foreach ($p in $Projects) { if ($p.name -eq $Name) { return $p } }
    $null
}

function Save-Projects {
    <# Rewrites the whole block. A project's rows move as tasks are added and
       removed, so patching individual cells would be more fragile than this. #>
    param($Projects)
    $ws = Get-ProjectSheet
    if (-not $ws) { throw 'The Projects sheet is missing.' }

    $rows = New-Object System.Collections.ArrayList
    foreach ($p in $Projects) {
        $tasks = @($p.tasks)
        if (-not $tasks.Count) { $tasks = @('') }
        foreach ($t in $tasks) { [void]$rows.Add(@($p.name, $p.status, $t)) }
    }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -ge 2) { [void]$ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 3)).ClearContents() }

    # Written a cell at a time on purpose: assigning an object[,] to Range.Value2
    # fails through PowerShell's COM binder, and the list is only ever tens of rows.
    $r = 2
    foreach ($row in $rows) {
        $ws.Cells.Item($r, 1).Value2 = $row[0]
        $ws.Cells.Item($r, 2).Value2 = $row[1]
        if ($row[2]) { $ws.Cells.Item($r, 3).Value2 = $row[2] }
        $r++
    }
    $script:Wb.Save()
}

function Update-Projects {
    <# The local is $proj, never $p: PowerShell variable names are case
       insensitive, so $p would quietly overwrite the $P payload. #>
    param($P)
    $action = [string]$P.action
    $name   = ([string]$P.name).Trim()
    $list   = @(Get-Projects)

    Write-Journal -Payload $P -Action 'project'

    switch ($action) {
        'addProject' {
            if (-not $name) { throw 'A project needs a name.' }
            if ($name.Length -gt 120) { throw 'That project name is too long.' }
            if (Find-Project $list $name) { throw "'$name' is already on the list." }
            $status = [string]$P.status
            if ($script:PROJ_STATUS -notcontains $status) { $status = 'New' }
            $list += , ([ordered]@{ name = $name; status = $status; tasks = New-Object System.Collections.ArrayList })
        }
        'setStatus' {
            $status = [string]$P.status
            $proj = Find-Project $list $name
            if (-not $proj) { throw "No project called '$name'." }
            if ($script:PROJ_STATUS -notcontains $status) { throw "Unknown status '$status'." }
            $proj.status = $status
        }
        'removeProject' {
            if (-not (Find-Project $list $name)) { throw "No project called '$name'." }
            $list = @($list | Where-Object { $_.name -ne $name })
        }
        'addTask' {
            $task = ([string]$P.task).Trim()
            $proj = Find-Project $list $name
            if (-not $proj) { throw "No project called '$name'." }
            if (-not $task) { throw 'A task needs a name.' }
            if ($task.Length -gt 120) { throw 'That task name is too long.' }
            if ($proj.tasks -contains $task) { throw "'$task' is already on $name." }
            [void]$proj.tasks.Add($task)
        }
        'removeTask' {
            $task = ([string]$P.task).Trim()
            $proj = Find-Project $list $name
            if (-not $proj) { throw "No project called '$name'." }
            $keep = New-Object System.Collections.ArrayList
            foreach ($t in $proj.tasks) { if ($t -ne $task) { [void]$keep.Add($t) } }
            $proj.tasks = $keep
        }
        default { throw "Unknown action '$action'." }
    }

    Save-Projects $list
    @{ ok = $true; action = $action }
}

# --------------------------------------------------------------- team
function Get-Monday {
    param([datetime]$D)
    $dow = [int]$D.DayOfWeek
    if ($dow -eq 0) { $dow = 7 }        # PowerShell counts Sunday as 0
    $D.Date.AddDays(1 - $dow)
}

function Get-WeekLabel {
    <# ISO-8601 week of the Monday's own week, matching the Week Number column
       on the timesheet and the label the browser shows. #>
    param([datetime]$Monday)
    $thu = $Monday.AddDays(3)
    $jan1 = New-Object datetime $thu.Year, 1, 1
    '{0}-W{1:00}' -f $thu.Year, ([math]::Floor(($thu - $jan1).Days / 7) + 1)
}

function Get-Sheet {
    param([string]$Name)
    foreach ($w in $script:Wb.Worksheets) { if ($w.Name -eq $Name) { return $w } }
    $null
}

function New-Sheet {
    param([string]$Name, [string[]]$Headers, [int[]]$Widths)
    $ws = $script:Wb.Worksheets.Add([System.Reflection.Missing]::Value,
                                    $script:Wb.Worksheets.Item($script:Wb.Worksheets.Count))
    $ws.Name = $Name
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $ws.Cells.Item(1, $i + 1).Value2 = $Headers[$i]
        $ws.Columns.Item($i + 1).ColumnWidth = $Widths[$i]
    }
    $ws.Rows.Item(1).Font.Bold = $true
    $ws
}

function Initialize-TeamSheets {
    $made = $false

    if (-not (Get-Sheet $script:TEAMWS)) {
        $ws = New-Sheet $script:TEAMWS @('Person', 'Status') @(28, 14)
        $r = 2
        foreach ($p in $script:TEAM_SEED) {
            $ws.Cells.Item($r, 1).Value2 = $p
            $ws.Cells.Item($r, 2).Value2 = 'Active'
            $r++
        }
        $made = $true
        Write-Host "  Team       : sheet created with $($script:TEAM_SEED.Count) people" -ForegroundColor DarkGray
    }

    if (-not (Get-Sheet $script:TEAMACTWS)) {
        [void](New-Sheet $script:TEAMACTWS @('Week Starting', 'Person', 'Activity', 'Week', 'Notes') @(15, 20, 12, 12, 50))
        $made = $true
        Write-Host "  Team       : activity sheet created" -ForegroundColor DarkGray
    }

    if ($made) { $script:Wb.Save() }
}

function Get-TeamPeople {
    $out = New-Object System.Collections.ArrayList
    $ws  = Get-Sheet $script:TEAMWS
    if (-not $ws) { return $out }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -lt 2) { return $out }

    $grid = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 2)).Value2
    for ($i = 1; $i -le $last - 1; $i++) {
        $name = ([string]$grid.GetValue($i, 1)).Trim()
        if (-not $name) { continue }
        $status = ([string]$grid.GetValue($i, 2)).Trim()
        if ($script:TEAM_STATUS -notcontains $status) { $status = 'Active' }
        [void]$out.Add([ordered]@{ name = $name; status = $status })
    }
    $out
}

function Save-TeamPeople {
    param($People)
    $ws = Get-Sheet $script:TEAMWS
    if (-not $ws) { throw 'The Team sheet is missing.' }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -ge 2) { [void]$ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 2)).ClearContents() }

    $r = 2
    foreach ($p in $People) {
        $ws.Cells.Item($r, 1).Value2 = $p.name
        $ws.Cells.Item($r, 2).Value2 = $p.status
        $r++
    }
    $script:Wb.Save()
}

function Get-TeamWeeks {
    $out = New-Object System.Collections.ArrayList
    $ws  = Get-Sheet $script:TEAMACTWS
    if (-not $ws) { return $out }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -lt 2) { return $out }

    $grid = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 5)).Value2
    for ($i = 1; $i -le $last - 1; $i++) {
        $dv = $grid.GetValue($i, 1)
        if ($dv -isnot [double]) { continue }
        $monday = Get-Monday ([datetime]::FromOADate($dv))
        $person = ([string]$grid.GetValue($i, 2)).Trim()
        if (-not $person) { continue }
        $act = ([string]$grid.GetValue($i, 3)).Trim()
        if (-not $act) { $act = 'WFH' }

        [void]$out.Add([ordered]@{
            row       = $i + 1
            weekStart = $monday.ToString('yyyy-MM-dd')
            weekEnd   = $monday.AddDays(4).ToString('yyyy-MM-dd')
            week      = Get-WeekLabel $monday
            person    = $person
            activity  = $act
            notes     = ([string]$grid.GetValue($i, 5)).Trim()
        })
    }
    $out
}

function Save-TeamWeeks {
    param($Weeks)
    $ws = Get-Sheet $script:TEAMACTWS
    if (-not $ws) { throw 'The TeamActivity sheet is missing.' }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -ge 2) { [void]$ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 5)).ClearContents() }

    $sorted = @($Weeks | Sort-Object weekStart, person)
    $r = 2
    foreach ($w in $sorted) {
        $d = [datetime]::ParseExact($w.weekStart, 'yyyy-MM-dd', $null)
        $ws.Cells.Item($r, 1).Value2  = [double]$d.ToOADate()
        $ws.Cells.Item($r, 1).NumberFormat = 'dd-mmm-yyyy'
        $ws.Cells.Item($r, 2).Value2  = $w.person
        $ws.Cells.Item($r, 3).Value2  = $w.activity
        # Left as a formula so the label follows the date if it is edited in Excel.
        $ws.Cells.Item($r, 4).Formula = "=IF(A$r=`"`",`"`",TEXT(A$r,`"yyyy`")&`"-W`"&TEXT(ISOWEEKNUM(A$r),`"00`"))"
        if ($w.notes) { $ws.Cells.Item($r, 5).Value2 = $w.notes }
        $r++
    }
    $script:Wb.Save()
}

function Get-Team {
    [ordered]@{
        allowance  = $script:TEAM_ALLOW
        activities = @($script:TEAM_ACTS)
        people     = @(Get-TeamPeople)
        entries    = @(Get-TeamWeeks)
    }
}

function Update-Team {
    <# $member / $entry rather than $p: PowerShell variable names are case
       insensitive, so $p would quietly overwrite the $P payload. #>
    param($P)
    $action = [string]$P.action
    $name   = ([string]$P.person).Trim()
    $result = @{ ok = $true; action = $action }

    Write-Journal -Payload $P -Action 'team'

    switch ($action) {
        'addPerson' {
            if (-not $name) { throw 'A person needs a name.' }
            if ($name.Length -gt 80) { throw 'That name is too long.' }
            $people = @(Get-TeamPeople)
            foreach ($m in $people) { if ($m.name -eq $name) { throw "$name is already on the team." } }
            $people += , ([ordered]@{ name = $name; status = 'Active' })
            Save-TeamPeople $people
        }
        'setPersonStatus' {
            $status = [string]$P.status
            if ($script:TEAM_STATUS -notcontains $status) { throw "Unknown status '$status'." }
            $people = @(Get-TeamPeople)
            $member = $null
            foreach ($m in $people) { if ($m.name -eq $name) { $member = $m } }
            if (-not $member) { throw "No one called '$name' on the team." }
            $member.status = $status
            Save-TeamPeople $people
        }
        'removePerson' {
            $people = @(Get-TeamPeople)
            $keep = @($people | Where-Object { $_.name -ne $name })
            if ($keep.Count -eq $people.Count) { throw "No one called '$name' on the team." }
            if (@(Get-TeamWeeks | Where-Object { $_.person -eq $name }).Count) {
                throw "$name has weeks logged. Set them to Inactive instead so the history is kept."
            }
            Save-TeamPeople $keep
        }
        'addWeek' {
            $people = @(Get-TeamPeople)
            $member = $null
            foreach ($m in $people) { if ($m.name -eq $name) { $member = $m } }
            if (-not $member) { throw "No one called '$name' on the team." }
            if ($member.status -ne 'Active') { throw "$name is Inactive." }

            $act = ([string]$P.activity).Trim()
            if ($script:TEAM_ACTS -notcontains $act) { throw "Unknown activity '$act'." }
            if (-not $P.weekStart) { throw 'Pick a week.' }

            $monday = Get-Monday ([datetime]::ParseExact([string]$P.weekStart, 'yyyy-MM-dd', $null))
            $label  = Get-WeekLabel $monday
            $iso    = $monday.ToString('yyyy-MM-dd')

            $weeks = @(Get-TeamWeeks)
            foreach ($w in $weeks) {
                if ($w.person -eq $name -and $w.weekStart -eq $iso -and $w.activity -eq $act) {
                    throw "$name already has $act booked for $label."
                }
            }

            $weeks += , ([ordered]@{
                weekStart = $iso; person = $name; activity = $act
                notes = ([string]$P.notes).Trim()
            })
            Save-TeamWeeks $weeks

            # Over-allowance is flagged, not blocked - exceptions get granted.
            $year = $label.Substring(0, 4)
            $used = @($weeks | Where-Object {
                $_.person -eq $name -and $_.activity -eq $act -and
                (Get-WeekLabel ([datetime]::ParseExact($_.weekStart, 'yyyy-MM-dd', $null))).StartsWith($year)
            }).Count
            if ($used -gt $script:TEAM_ALLOW) {
                $result.warning = "$name is now on $used $act weeks in $year - over the allowance of $($script:TEAM_ALLOW)."
            }
            $result.week = $label
        }
        'removeWeek' {
            $row = [int]$P.row
            $weeks = @(Get-TeamWeeks)
            $keep = @($weeks | Where-Object { $_.row -ne $row })
            if ($keep.Count -eq $weeks.Count) { throw 'That entry is no longer there.' }
            Save-TeamWeeks $keep
        }
        default { throw "Unknown action '$action'." }
    }

    $result
}

# --------------------------------------------------------------- holidays
function Initialize-HolidaySheet {
    if (Get-Sheet $script:HOLWS) { return }

    $ws = New-Sheet $script:HOLWS @('Date', 'Day', 'Holiday') @(14, 14, 30)
    $r = 2
    foreach ($h in $script:HOL_SEED) {
        $d = [datetime]::ParseExact($h.date, 'yyyy-MM-dd', $null)
        $ws.Cells.Item($r, 1).Value2 = [double]$d.ToOADate()
        $ws.Cells.Item($r, 1).NumberFormat = 'dd-mmm-yyyy'
        # Left as a formula so the day name follows the date if it is edited in Excel.
        $ws.Cells.Item($r, 2).Formula = "=IF(A$r=`"`",`"`",TEXT(A$r,`"dddd`"))"
        $ws.Cells.Item($r, 3).Value2 = $h.name
        $r++
    }
    $script:Wb.Save()
    Write-Host "  Holidays   : sheet created with $($script:HOL_SEED.Count) dates" -ForegroundColor DarkGray
}

function Get-Holidays {
    <# Whatever is on the Holidays sheet, oldest first. The Day column is
       recomputed from the date so a stale hand-typed day name cannot lie. #>
    $out = New-Object System.Collections.ArrayList
    $ws  = Get-Sheet $script:HOLWS
    if (-not $ws) { return $out }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -lt 2) { return $out }

    $grid = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 3)).Value2
    for ($i = 1; $i -le $last - 1; $i++) {
        $dv = $grid.GetValue($i, 1)
        if ($dv -isnot [double]) { continue }
        $name = ([string]$grid.GetValue($i, 3)).Trim()
        if (-not $name) { continue }
        $d = [datetime]::FromOADate($dv)
        [void]$out.Add([ordered]@{
            date = $d.ToString('yyyy-MM-dd')
            day  = $d.DayOfWeek.ToString()
            name = $name
        })
    }
    @($out | Sort-Object { $_.date })
}

function Save-Holidays {
    param($Holidays)
    $ws = Get-Sheet $script:HOLWS
    if (-not $ws) { throw 'The Holidays sheet is missing.' }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -ge 2) { [void]$ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, 3)).ClearContents() }

    $r = 2
    foreach ($h in @($Holidays | Sort-Object { $_.date })) {
        $d = [datetime]::ParseExact($h.date, 'yyyy-MM-dd', $null)
        $ws.Cells.Item($r, 1).Value2 = [double]$d.ToOADate()
        $ws.Cells.Item($r, 1).NumberFormat = 'dd-mmm-yyyy'
        $ws.Cells.Item($r, 2).Formula = "=IF(A$r=`"`",`"`",TEXT(A$r,`"dddd`"))"
        $ws.Cells.Item($r, 3).Value2 = $h.name
        $r++
    }
    $script:Wb.Save()
}

function Update-Holidays {
    param($P)
    $action = [string]$P.action
    $date   = ([string]$P.date).Trim()

    Write-Journal -Payload $P -Action 'holiday'

    $d = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($date, 'yyyy-MM-dd',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        throw 'Pick a date.'
    }
    $list = @(Get-Holidays)

    switch ($action) {
        'addHoliday' {
            $name = ([string]$P.name).Trim()
            if (-not $name) { throw 'A holiday needs a name.' }
            if ($name.Length -gt 60) { throw 'That name is too long.' }
            foreach ($h in $list) {
                if ($h.date -eq $date) { throw "$($h.name) is already on $($d.ToString($script:DATE_FORMAT))." }
            }
            $list += , ([ordered]@{ date = $date; name = $name })
            Save-Holidays $list
        }
        'removeHoliday' {
            $keep = @($list | Where-Object { $_.date -ne $date })
            if ($keep.Count -eq $list.Count) { throw 'That holiday is no longer there.' }
            Save-Holidays $keep
        }
        default { throw "Unknown action '$action'." }
    }

    @{ ok = $true; action = $action }
}

function Get-GridText {
    param($Grid, [int]$R, [int]$C)
    $v = $Grid.GetValue($R, $C)
    if ($null -eq $v) { '' } else { [string]$v }
}

function Get-Entries {
    <# Every data row, newest last. Row numbers are worksheet rows so delete
       can target them directly.

       The whole table comes back in one Value2 call. Reading cell by cell is a
       COM round trip each - roughly fifteen per row - and this runs on every
       request including every write, so it was the thing that made the form get
       slower as the timesheet grew. Every column the form uses is either text
       or a formula that evaluates to text, so Value2 matches what .Text gave;
       only Date needs formatting back. #>
    $lo = Get-Table
    $entries = New-Object System.Collections.ArrayList
    if ($lo.ListRows.Count -eq 0) { return $entries }

    $first = $lo.DataBodyRange.Row
    $grid  = $lo.DataBodyRange.Value2      # object[,], 1-based [row, column]

    for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
        $iso = ''; $dateText = ''
        $dv = $grid.GetValue($i, $script:COL.Date)
        if ($dv -is [double]) {
            $d = [datetime]::FromOADate($dv)
            $iso      = $d.ToString('yyyy-MM-dd')
            $dateText = $d.ToString($script:DATE_FORMAT)
        }
        $hours = 0.0
        $hv = $grid.GetValue($i, $script:COL.Hours)
        if ($null -ne $hv -and $hv -ne '') { $hours = [double]$hv }

        [void]$entries.Add([ordered]@{
            row               = $first + $i - 1
            date              = $iso
            dateText          = $dateText
            day               = Get-GridText $grid $i $script:COL.Day
            weekNumber        = Get-GridText $grid $i $script:COL.WeekNumber
            project           = Get-GridText $grid $i $script:COL.Project
            task              = Get-GridText $grid $i $script:COL.Task
            hours             = $hours
            timeTracked       = Get-GridText $grid $i $script:COL.TimeTracked
            billable          = Get-GridText $grid $i $script:COL.Billable
            category          = Get-GridText $grid $i $script:COL.Category
            budget            = Get-GridText $grid $i $script:COL.Budget
            notes             = Get-GridText $grid $i $script:COL.Notes
            ticket            = Get-GridText $grid $i $script:COL.Ticket
            incidentType      = Get-GridText $grid $i $script:COL.IncidentType
            nonBillableReason = Get-GridText $grid $i $script:COL.NonBillableReason
            user             = Get-GridText $grid $i $script:COL.User
        })
    }
    $entries
}

function Merge-Options {
    <# Lists-sheet values first (preserves the curated order), then any value
       actually in use that is not already there. Means a project typed into
       the form shows up in the dropdown next load with no write-back. #>
    param($Curated, $InUse)    $seen = @{}
    $out  = New-Object System.Collections.ArrayList
    foreach ($v in $Curated) {
        if ($v -and -not $seen.ContainsKey($v.ToLower())) { $seen[$v.ToLower()] = $true; [void]$out.Add($v) }
    }
    foreach ($v in ($InUse | Sort-Object)) {
        if ($v -and -not $seen.ContainsKey($v.ToLower())) { $seen[$v.ToLower()] = $true; [void]$out.Add($v) }
    }
    ,$out
}

function Get-Bootstrap {
    $curated = Get-ListSheetOptions
    $entries = Get-Entries

    # Project and Task are driven entirely by the Projects sheet now, so only
    # projects that are actually being worked on can be booked against.
    $projects = @(Get-Projects)
    $active   = @($projects | Where-Object { $_.status -eq $script:PROJ_ACTIVE })

    $projectTasks = [ordered]@{}
    foreach ($p in $active) { $projectTasks[$p.name] = @($p.tasks) }

    $col = { param($k) ,@($entries | ForEach-Object { $_.$k } | Where-Object { $_ } | Select-Object -Unique) }

    $today    = (Get-Date).Date
    $monday   = $today.AddDays(-(([int]$today.DayOfWeek + 6) % 7))
    $todayIso = $today.ToString('yyyy-MM-dd')

    $todayTotal = 0.0; $weekTotal = 0.0
    foreach ($e in $entries) {
        if (-not $e.date) { continue }
        $d = [datetime]::ParseExact($e.date, 'yyyy-MM-dd', $null)
        if ($e.date -eq $todayIso) { $todayTotal += $e.hours }
        if ($d -ge $monday -and $d -lt $monday.AddDays(7)) { $weekTotal += $e.hours }
    }

    $recent = @($entries)
    if ($recent.Count -gt $script:RECENT) { $recent = @($recent[($recent.Count - $script:RECENT)..($recent.Count - 1)]) }

    [ordered]@{
        options = [ordered]@{
            project           = @($active | ForEach-Object { $_.name })
            task              = @($active | ForEach-Object { $_.tasks } | Select-Object -Unique)
            category          = Merge-Options $curated.category          (& $col 'category')
            budget            = Merge-Options $curated.budget            (& $col 'budget')
            billable          = Merge-Options $curated.billable          (& $col 'billable')
            incidentType      = Merge-Options $curated.incidentType      (& $col 'incidentType')
            nonBillableReason = Merge-Options $curated.nonBillableReason (& $col 'nonBillableReason')
        }
        projects     = @($projects | ForEach-Object { [ordered]@{ name = $_.name; status = $_.status; tasks = @($_.tasks) } })
        projectTasks = $projectTasks
        entries      = @($recent)
        totals       = [ordered]@{ today = [math]::Round($todayTotal,2); week = [math]::Round($weekTotal,2); count = $entries.Count }
        weekStart    = $monday.ToString('yyyy-MM-dd')
        today        = $todayIso
    }
}

function Get-WeekReport {
    <# Every entry inside the Mon-Sun week that $Start falls in. The browser
       does the grouping; the server just scopes the rows. #>
    param([string]$Start)
    if (-not $Start) {
        $today = (Get-Date).Date
        $monday = $today.AddDays(-(([int]$today.DayOfWeek + 6) % 7))
    } else {
        $d = [datetime]::ParseExact($Start, 'yyyy-MM-dd', $null).Date
        $monday = $d.AddDays(-(([int]$d.DayOfWeek + 6) % 7))
    }
    $end  = $monday.AddDays(7)
    $rows = New-Object System.Collections.ArrayList
    foreach ($e in (Get-Entries)) {
        if (-not $e.date) { continue }
        $dd = [datetime]::ParseExact($e.date, 'yyyy-MM-dd', $null)
        if ($dd -ge $monday -and $dd -lt $end) { [void]$rows.Add($e) }
    }
    [ordered]@{
        weekStart = $monday.ToString('yyyy-MM-dd')
        weekEnd   = $monday.AddDays(6).ToString('yyyy-MM-dd')
        entries   = @($rows)
    }
}

function Get-AllEntries {
    <# Every row, unscoped. The browser filters by date range and groups it. #>
    $rows  = @(Get-Entries)
    $dates = @($rows | ForEach-Object { $_.date } | Where-Object { $_ } | Sort-Object)
    $first = ''
    $last  = ''
    if ($dates.Count) { $first = $dates[0]; $last = $dates[$dates.Count - 1] }
    [ordered]@{
        entries = $rows
        first   = $first
        last    = $last
    }
}

# --------------------------------------------------------------- writing
function ConvertTo-Hours {
    <# Accepts 1.5 | .5 | 1:30 | 1h 30m | 1h | 90m -> decimal hours. #>
    param([string]$Raw)
    $s = ($Raw -replace '\s', '').ToLower()
    if (-not $s) { throw 'Hours is required.' }

    if ($s -match '^(\d+):([0-5]?\d)$')            { return [math]::Round([double]$Matches[1] + [double]$Matches[2]/60, 4) }
    if ($s -match '^(\d+(?:\.\d+)?)h(\d+)m?$')     { return [math]::Round([double]$Matches[1] + [double]$Matches[2]/60, 4) }
    if ($s -match '^(\d+(?:\.\d+)?)h$')            { return [math]::Round([double]$Matches[1], 4) }
    if ($s -match '^(\d+(?:\.\d+)?)m$')            { return [math]::Round([double]$Matches[1]/60, 4) }
    if ($s -match '^\.?\d+(\.\d+)?$' -or $s -match '^\d*\.\d+$') { return [math]::Round([double]$s, 4) }

    throw "Could not read '$Raw' as a duration. Try 1.5, 1:30, 1h 30m or 90m."
}

function Write-Journal {
    param($Payload, [string]$Action)
    $line = [ordered]@{
        at      = (Get-Date).ToString('s')
        action  = $Action
        payload = $Payload
    } | ConvertTo-Json -Depth 6 -Compress
    Add-Content -Path $script:JOURNAL -Value $line -Encoding UTF8
}

$script:Tasks = $null      # board state, cached from the Tasks sheet on first use

function Update-TaskState {
    <# Fold one task event into the board. Named $Evt, not $Event, because
       $Event is a PowerShell automatic variable. #>
    param($Tasks, $Evt)
    $taskId = [string]$Evt.taskId
    if (-not $taskId) { return }

    switch ([string]$Evt.type) {
        'task:created' {
            $prio = 3
            if ($Evt.priority) { $prio = [int]$Evt.priority }
            $seq = 0
            if ($Evt.seq) { $seq = [int]$Evt.seq }
            $Tasks[$taskId] = @{
                id          = $taskId
                title       = $Evt.title
                status      = $Evt.status
                priority    = $prio
                seq         = $seq
                date        = $Evt.date
                createdAt   = $Evt.timestamp
                completedAt = $null
            }
        }
        'task:status_changed'   {
            if ($Tasks.ContainsKey($taskId)) {
                $Tasks[$taskId].status = $Evt.newStatus
                # Completion time is the stamp on the event that closed the task,
                # so replaying an old journal backfills it.
                $Tasks[$taskId].completedAt =
                    if ([string]$Evt.newStatus -eq 'Complete') { $Evt.timestamp } else { $null }
            }
        }
        'task:priority_changed' { if ($Tasks.ContainsKey($taskId)) { $Tasks[$taskId].priority = [int]$Evt.newPriority } }
        'task:seq_changed'      { if ($Tasks.ContainsKey($taskId)) { $Tasks[$taskId].seq      = [int]$Evt.newSeq      } }
        'task:title_changed'    { if ($Tasks.ContainsKey($taskId)) { $Tasks[$taskId].title    = $Evt.newTitle        } }
        'task:deleted'          { $Tasks.Remove($taskId) }
    }
}

function Get-Tasks {
    <# The Tasks sheet is the store; it is read once and the cache is kept in
       step as events arrive, the same way the board used to be cached. #>
    if ($null -eq $script:Tasks) { $script:Tasks = Read-TaskSheet }
    @{ tasks = $script:Tasks; today = (Get-Date).Date.ToString('yyyy-MM-dd') }
}

function Read-TaskJournal {
    <# One-off migration path: the board used to live only in journal.jsonl. #>
    $tasks = @{}
    if (-not (Test-Path $script:JOURNAL)) { return $tasks }
    foreach ($line in @(Get-Content -Path $script:JOURNAL -Encoding UTF8 -ErrorAction SilentlyContinue)) {
        if (-not $line) { continue }
        $entry = $null
        try { $entry = $line | ConvertFrom-Json } catch { continue }   # skip malformed lines
        if ($entry.action -eq 'task_event') { Update-TaskState $tasks $entry.payload }
    }
    $tasks
}

function Initialize-TaskSheet {
    if (Get-Sheet $script:TASKWS) { return }

    [void](New-Sheet $script:TASKWS $script:TASK_HEAD @(22, 52, 14, 10, 8, 14, 20, 20))
    $seed = Read-TaskJournal
    if ($seed.Count) {
        Save-Tasks $seed
        Write-Host "  Tasks      : sheet created, $($seed.Count) tasks moved in from the journal" -ForegroundColor DarkGray
    } else {
        $script:Wb.Save()
        Write-Host '  Tasks      : sheet created' -ForegroundColor DarkGray
    }
}

function ConvertTo-LocalDate {
    <# Anything the board carries - a real DateTime, an Excel OADate, an ISO
       stamp with a Z, or a plain yyyy-MM-dd - as a local DateTime, or $null.
       Parsed invariantly on purpose: under a dd-MM culture "09/04/2026" reads
       as 9 April, and "08/25/2026" does not parse at all. #>
    param($Value)
    if ($null -eq $Value)    { return $null }
    if ($Value -is [double]) { return [datetime]::FromOADate($Value) }

    if ($Value -is [datetime]) {
        $d = $Value
    } else {
        $t = ([string]$Value).Trim()
        if (-not $t) { return $null }
        $parsed = [datetime]::MinValue
        if (-not [datetime]::TryParse($t, [System.Globalization.CultureInfo]::InvariantCulture,
                                      [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $null }
        $d = $parsed
    }
    # The browser stamps events in UTC; the sheet is read by a human, so it holds
    # wall-clock time.
    if ($d.Kind -eq [System.DateTimeKind]::Utc) { return $d.ToLocalTime() }
    $d
}

function ConvertTo-Stamp {
    <# Cell value to a local date/time string, or $null. Excel hands dates back
       as OADate doubles, but a hand-typed cell comes through as text. #>
    param($V, [string]$Format)
    $d = ConvertTo-LocalDate $V
    if ($null -eq $d) { return $null }
    $d.ToString($Format, [System.Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-Count {
    param($V, [int]$Default)
    $n = 0.0
    if ($null -ne $V -and [double]::TryParse(([string]$V).Trim(), [ref]$n)) { return [int]$n }
    $Default
}

function Read-TaskSheet {
    $out = @{}
    $ws  = Get-Sheet $script:TASKWS
    if (-not $ws) { return $out }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -lt 2) { return $out }

    $grid = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, $script:TASK_COLS)).Value2
    for ($i = 1; $i -le $last - 1; $i++) {
        $id = ([string]$grid.GetValue($i, 1)).Trim()
        if (-not $id) { continue }
        $status = ([string]$grid.GetValue($i, 3)).Trim()
        if (-not $status) { $status = 'New' }
        $out[$id] = @{
            id          = $id
            title       = ([string]$grid.GetValue($i, 2)).Trim()
            status      = $status
            priority    = ConvertTo-Count $grid.GetValue($i, 4) 3
            seq         = ConvertTo-Count $grid.GetValue($i, 5) 0
            date        = ConvertTo-Stamp $grid.GetValue($i, 6) 'yyyy-MM-dd'
            createdAt   = ConvertTo-Stamp $grid.GetValue($i, 7) 's'
            completedAt = ConvertTo-Stamp $grid.GetValue($i, 8) 's'
        }
    }
    $out
}

function Set-DateCell {
    <# Real Excel dates, not text, so the sheet sorts and filters on its own. #>
    param($Ws, [int]$R, [int]$C, $Value, [string]$Format)
    $d = ConvertTo-LocalDate $Value
    if ($null -eq $d) { return }
    $cell = $Ws.Cells.Item($R, $C)
    $cell.Value2 = [double]$d.ToOADate()
    $cell.NumberFormat = $Format
}

function Save-Tasks {
    <# Rewrites the whole block. The board is small and rows move as tasks are
       reordered, so tracking row numbers would buy nothing. #>
    param($Tasks)
    $ws = Get-Sheet $script:TASKWS
    if (-not $ws) { throw 'The Tasks sheet is missing.' }

    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -ge 2) {
        [void]$ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($last, $script:TASK_COLS)).ClearContents()
    }

    $ordered = @($Tasks.Values | Sort-Object `
        @{ Expression = { [int]$_.priority } },
        @{ Expression = { if ($_.seq) { [int]$_.seq } else { 999 } } },
        @{ Expression = { [string]$_.createdAt } })

    $r = 2
    foreach ($t in $ordered) {
        $ws.Cells.Item($r, 1).Value2 = [string]$t.id
        $ws.Cells.Item($r, 2).Value2 = [string]$t.title
        $ws.Cells.Item($r, 3).Value2 = [string]$t.status
        $ws.Cells.Item($r, 4).Value2 = [double]$t.priority
        $ws.Cells.Item($r, 5).Value2 = [double]$t.seq
        Set-DateCell $ws $r 6 $t.date        'dd-mmm-yyyy'
        Set-DateCell $ws $r 7 $t.createdAt   'dd-mmm-yyyy hh:mm'
        Set-DateCell $ws $r 8 $t.completedAt 'dd-mmm-yyyy hh:mm'
        $r++
    }
    $script:Wb.Save()
}

function Assert-Payload {
    <# Shared validation for add and update. Returns the parsed date + hours. #>
    param($P)
    if (-not $P.date)    { throw 'Date is required.' }
    if (-not $P.project) { throw 'Project is required.' }
    if (-not $P.budget)  { throw 'Budget is required.' }

    # The form only offers In-Progress projects, but the field is free text, so
    # a stale or mistyped name would otherwise slip straight into the timesheet.
    $project = ([string]$P.project).Trim()
    $active  = @(Get-Projects | Where-Object { $_.status -eq $script:PROJ_ACTIVE } | ForEach-Object { $_.name })
    if ($active -notcontains $project) {
        throw "'$project' is not an In-Progress project. Set it to In-Progress on the Projects tab to book time against it."
    }

    $hours = ConvertTo-Hours ([string]$P.hours)
    if ($hours -le 0)  { throw 'Hours must be greater than zero.' }
    if ($hours -gt 24) { throw 'Hours must be 24 or less.' }
    @{
        Hours = $hours
        Date  = [datetime]::ParseExact([string]$P.date, 'yyyy-MM-dd', $null)
    }
}

function Add-Entry {
    param($P, [string]$Username)
    $v = Assert-Payload $P

    # Journal before touching Excel so nothing is lost if the write fails.
    Write-Journal -Payload $P -Action 'add'

    $lo = Get-Table
    $lr = $lo.ListRows.Add()          # copies the derived-column formulas down
    $r  = $lr.Range.Row
    $ws = $script:Wb.Worksheets.Item($script:SHEET)

    $ws.Cells.Item($r, $script:COL.Date).Value2  = [double]$v.Date.ToOADate()
    $ws.Cells.Item($r, $script:COL.Hours).Value2 = [double]$v.Hours
    $ws.Cells.Item($r, $script:COL.User).Value2 = $Username.ToUpperInvariant()
    foreach ($m in $script:WRITE_MAP) {
        $val = [string]$P.($m.Key)
        if ($val) { $ws.Cells.Item($r, $m.Col).Value2 = $val }
    }

    Update-Pivot
    $script:Wb.Save()
    @{ row = $r; hours = $v.Hours }
}

function Update-Entry {
    <# Edit in place. Unlike Add-Entry this clears emptied fields, otherwise an
       edit could never remove a value once it had been set. Columns 7 and
       15-18 are formulas and are never touched. #>
    param($P, [string]$Username)
    $row = [int]$P.row
    $v   = Assert-Payload $P

    $lo    = Get-Table
    $first = $lo.DataBodyRange.Row
    $last  = $first + $lo.ListRows.Count - 1
    if ($row -lt $first -or $row -gt $last) { throw "Row $row is not part of the timesheet table." }

    Write-Journal -Payload $P -Action 'update'

    $ws = $script:Wb.Worksheets.Item($script:SHEET)
    $ws.Cells.Item($row, $script:COL.Date).Value2  = [double]$v.Date.ToOADate()
    $ws.Cells.Item($row, $script:COL.Hours).Value2 = [double]$v.Hours
    $ws.Cells.Item($row, $script:COL.User).Value2 = $Username.ToUpperInvariant()
    foreach ($m in $script:WRITE_MAP) {
        $val = [string]$P.($m.Key)
        if ($val) { $ws.Cells.Item($row, $m.Col).Value2 = $val }
        else      { [void]$ws.Cells.Item($row, $m.Col).ClearContents() }
    }

    Update-Pivot
    $script:Wb.Save()
    @{ row = $row; hours = $v.Hours }
}

function Remove-Entry {
    param([int]$Row)
    $lo    = Get-Table
    $first = $lo.DataBodyRange.Row
    $last  = $first + $lo.ListRows.Count - 1
    if ($Row -lt $first -or $Row -gt $last) { throw "Row $Row is not part of the timesheet table." }

    Write-Journal -Payload @{ row = $Row } -Action 'delete'
    $lo.ListRows.Item($Row - $first + 1).Delete()
    Update-Pivot
    $script:Wb.Save()
    @{ deleted = $Row }
}

# --------------------------------------------------------------- http
$MIME = @{ '.html' = 'text/html; charset=utf-8'; '.js' = 'application/javascript; charset=utf-8'
           '.css'  = 'text/css; charset=utf-8';  '.ico' = 'image/x-icon'; '.jpg' = 'image/jpeg' }

# Bing's "picture of the day" changes once every 24h, so caching one file per
# day means only the first request each day ever hits the network - every
# other page load (and every browser re-fetch) is served from disk.
$script:WALLDIR = Join-Path $script:ROOT 'wallpaper-cache'

function Get-Wallpaper {
    $today = Get-Date -Format 'yyyy-MM-dd'
    $file  = Join-Path $script:WALLDIR "$today.jpg"
    if (Test-Path $file) { return [IO.File]::ReadAllBytes($file) }

    $meta = Invoke-RestMethod 'https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US'
    $url  = 'https://www.bing.com' + $meta.images[0].url
    if (-not (Test-Path $script:WALLDIR)) { New-Item $script:WALLDIR -ItemType Directory | Out-Null }
    Invoke-WebRequest $url -OutFile $file

    # Yesterday's file is no longer reachable by date, so it would sit there forever otherwise.
    Get-ChildItem $script:WALLDIR -Filter '*.jpg' | Where-Object { $_.Name -ne "$today.jpg" } | Remove-Item -Force
    [IO.File]::ReadAllBytes($file)
}

function Send-Response {
    param($Context, [int]$Status, [string]$ContentType, [byte[]]$Bytes)
    $res = $Context.Response
    $res.StatusCode  = $Status
    $res.ContentType = $ContentType
    $res.ContentLength64 = $Bytes.Length
    $res.Headers.Add('Cache-Control', 'no-store')
    try { $res.OutputStream.Write($Bytes, 0, $Bytes.Length) } catch { }
    $res.OutputStream.Close()
}

function Send-Json {
    param($Context, $Obj, [int]$Status = 200)
    $json = $Obj | ConvertTo-Json -Depth 8
    Send-Response $Context $Status 'application/json; charset=utf-8' ([Text.Encoding]::UTF8.GetBytes($json))
}

function Send-File {
    param($Context, [string]$Rel)
    $path = Join-Path $script:WWW $Rel
    $full = [IO.Path]::GetFullPath($path)
    if (-not $full.StartsWith([IO.Path]::GetFullPath($script:WWW))) {   # path traversal guard
        Send-Json $Context @{ error = 'Forbidden' } 403; return
    }
    if (-not (Test-Path $full -PathType Leaf)) {
        Send-Response $Context 404 'text/plain; charset=utf-8' ([Text.Encoding]::UTF8.GetBytes('Not found')); return
    }
    $ext = [IO.Path]::GetExtension($full).ToLower()
    $ct  = if ($MIME.ContainsKey($ext)) { $MIME[$ext] } else { 'application/octet-stream' }
    Send-Response $Context 200 $ct ([IO.File]::ReadAllBytes($full))
}

function Read-Body {
    param($Context)
    $reader = New-Object IO.StreamReader($Context.Request.InputStream, [Text.Encoding]::UTF8)
    $raw = $reader.ReadToEnd(); $reader.Close()
    if (-not $raw) { return @{} }
    $raw | ConvertFrom-Json
}

$script:AuthUsers = @{}
$script:AuthSessions = @{}

function Save-AuthUsers {
    $users = @($script:AuthUsers.Values | ForEach-Object {
        [ordered]@{
            username = $_.username
            passwordHash = $_.passwordHash
            passwordSalt = $_.passwordSalt
            passwordIterations = [int]$_.passwordIterations
            name = $_.name
            mustChangePassword = [bool]$_.mustChangePassword
        }
    })
    $users | ConvertTo-Json -Depth 3 | Set-Content -Path $script:USERS_FILE -Encoding UTF8
}

function New-PasswordRecord {
    param([string]$Password, [bool]$MustChangePassword)
    $salt = New-Object byte[] 16
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($salt)
    $iterations = 120000
    $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes(
        ([string]$Password), $salt, $iterations)
    try { $hash = $derive.GetBytes(32) } finally { $derive.Dispose() }
    @{
        passwordHash = [Convert]::ToBase64String($hash)
        passwordSalt = [Convert]::ToBase64String($salt)
        passwordIterations = $iterations
        mustChangePassword = $MustChangePassword
    }
}

function Test-Password {
    param($User, [string]$Password)
    if (-not $User.passwordHash -or -not $User.passwordSalt) { return $false }
    try {
        $salt = [Convert]::FromBase64String($User.passwordSalt)
        $expected = [Convert]::FromBase64String($User.passwordHash)
        $iterations = [int]$User.passwordIterations
        if ($iterations -lt 100000) { return $false }
        $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes(
            ([string]$Password), $salt, $iterations)
        try { $actual = $derive.GetBytes($expected.Length) } finally { $derive.Dispose() }
        if ($actual.Length -ne $expected.Length) { return $false }
        $different = 0
        for ($i = 0; $i -lt $expected.Length; $i++) { $different = $different -bor ($actual[$i] -bxor $expected[$i]) }
        return $different -eq 0
    } catch { return $false }
}

function Initialize-AuthUsers {
    if (Test-Path $script:USERS_FILE) {
        $stored = @(Get-Content -Path $script:USERS_FILE -Raw | ConvertFrom-Json)
        foreach ($user in $stored) {
            $username = ([string]$user.username).Trim()
            if ($username -match '^[A-Za-z0-9._-]{2,40}$' -and -not $script:AuthUsers.ContainsKey($username)) {
                $mustChange = $false
                if ($null -ne $user.mustChangePassword) { $mustChange = [bool]$user.mustChangePassword }
                if ($user.passwordHash -and $user.passwordSalt) {
                    $script:AuthUsers[$username] = @{
                        username = $username; passwordHash = [string]$user.passwordHash
                        passwordSalt = [string]$user.passwordSalt; passwordIterations = [int]$user.passwordIterations
                        name = [string]$user.name; mustChangePassword = $mustChange
                    }
                } elseif ($user.password) {
                    $record = New-PasswordRecord ([string]$user.password) $mustChange
                    $script:AuthUsers[$username] = @{ username = $username; name = [string]$user.name } + $record
                }
            }
        }
    }

    if (-not $script:AuthUsers.ContainsKey('admin')) {
        $script:AuthUsers['admin'] = @{ username = 'admin'; name = 'Administrator' } + (New-PasswordRecord 'admin123' $false)
    }
    if (-not $script:AuthUsers.ContainsKey('RAJESH')) {
        $script:AuthUsers['RAJESH'] = @{ username = 'RAJESH'; name = 'Rajesh' } + (New-PasswordRecord 'rajesh123' $false)
    }
    Save-AuthUsers
}

function Get-PublicAuthUsers {
    @($script:AuthUsers.Values | Sort-Object username | ForEach-Object {
        [ordered]@{ username = $_.username; name = $_.name }
    })
}

function Set-AuthPassword {
    param([string]$Username, [string]$Password, [bool]$MustChangePassword)
    $user = $script:AuthUsers[$Username]
    $record = New-PasswordRecord $Password $MustChangePassword
    $user.passwordHash = $record.passwordHash
    $user.passwordSalt = $record.passwordSalt
    $user.passwordIterations = $record.passwordIterations
    $user.mustChangePassword = $record.mustChangePassword
    Save-AuthUsers
}

function Assert-NewPassword {
    param([string]$Password)
    if ([string]::IsNullOrWhiteSpace($Password) -or $Password.Length -lt 8) {
        throw 'Password must be at least 8 characters.'
    }
}

function Ensure-Admin {
    param($Context)
    $session = Ensure-Authenticated $Context
    if ($session.username -ine 'admin') { throw 'Administrator access required.' }
    $session
}

function Get-SessionCookie {
    param($Request)
    $cookieHeader = $Request.Headers['Cookie']
    if (-not $cookieHeader) { return $null }
    $cookies = @{}
    foreach ($part in ($cookieHeader -split ';')) {
        $nameValue = $part.Trim()
        if (-not $nameValue) { continue }
        $eq = $nameValue.IndexOf('=')
        if ($eq -lt 0) { continue }
        $cookies[[string]$nameValue.Substring(0, $eq)] = [string]$nameValue.Substring($eq + 1)
    }
    if ($cookies.ContainsKey('timesheet_session')) { return $cookies['timesheet_session'] }
    return $null
}

function Ensure-Authenticated {
    param($Context)
    $sessionId = Get-SessionCookie $Context.Request
    if (-not $sessionId) { throw 'Authentication required.' }
    if (-not $script:AuthSessions.ContainsKey($sessionId)) { throw 'Session expired or invalid.' }
    $script:AuthSessions[$sessionId]
}

function Set-SessionCookie {
    param($Context, [string]$SessionId)
    $cookie = [System.Net.Cookie]::new('timesheet_session', $SessionId)
    $cookie.HttpOnly = $true
    $cookie.Path = '/'
    $Context.Response.AppendCookie($cookie)
}

function Clear-SessionCookie {
    param($Context)
    $cookie = [System.Net.Cookie]::new('timesheet_session', '')
    $cookie.Expires = [datetime]::UtcNow.AddDays(-1)
    $cookie.Path = '/'
    $Context.Response.AppendCookie($cookie)
}

function Test-SameOrigin {
    <# The server is only reachable from this machine, but "this machine"
       includes every page open in the browser. Anything cross-site that a page
       can trigger - fetch, XHR, a plain form post - carries an Origin header,
       so refusing the ones that are not ours keeps another tab from adding or
       deleting rows. No Origin at all means it did not come from a page (the
       Stop script, curl), which is fine. #>
    param($Request)
    $origin = $Request.Headers['Origin']
    if (-not $origin) { return $true }
    $script:ORIGINS -contains $origin
}

# --------------------------------------------------------------- main
if (-not (Test-Path $Workbook)) { throw "Workbook not found: $Workbook" }

Write-Host ''
Write-Host '  Timesheet entry' -ForegroundColor Cyan
Write-Host '  ---------------' -ForegroundColor DarkGray
Write-Host "  Workbook   : $Workbook" -ForegroundColor DarkGray

# Bind the port BEFORE touching Excel, so launching a second copy fails
# cleanly instead of opening the workbook and then backing out.
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
Initialize-AuthUsers
try { $listener.Start() }
catch { throw "Could not bind port $Port. Is the server already running? ($($_.Exception.Message))" }

try {
    Connect-Excel
    Initialize-ProjectSheet
    Initialize-TeamSheets
    Initialize-HolidaySheet
    Initialize-TaskSheet
    Initialize-UserColumn
}
catch {
    # 0x800A01A8 "Object required" - the workbook reference died mid-setup,
    # which on a OneDrive file means co-authoring detached it.
    $ex = $_.Exception
    while ($ex -and $ex.HResult -ne -2146827864) { $ex = $ex.InnerException }
    if (-not $ex) { throw }
    throw @"
Excel dropped the workbook while the server was setting it up (COM 0x800A01A8).

The workbook is on OneDrive, and creating the Projects/Team sheets is a
structural change that co-authoring cannot merge, so Excel detached the file.

To fix:
  1. Close every Excel window, including the workbook in Excel Online.
  2. Wait for OneDrive to finish syncing.
  3. Start again - AutoSave is now turned off on open.
Or keep the workbook outside OneDrive and pass -Workbook <local path>.
"@
}

$url = "http://localhost:$Port/"
Write-Host "  Listening  : $url" -ForegroundColor Green
Write-Host '  Ctrl+C to stop.' -ForegroundColor DarkGray
Write-Host ''
if (-not $NoBrowser) { Start-Process $url | Out-Null }

$script:StopRequested = $false

try {
    while ($listener.IsListening -and -not $script:StopRequested) {
        $ctx  = $listener.GetContext()
        $path = $ctx.Request.Url.AbsolutePath
        $verb = $ctx.Request.HttpMethod

        if ($verb -eq 'POST' -and -not (Test-SameOrigin $ctx.Request)) {
            Write-Host ("  {0,-6} {1,-22} REFUSED cross-origin post from {2}" -f $verb, $path, $ctx.Request.Headers['Origin']) -ForegroundColor Red
            Send-Json $ctx @{ error = 'Cross-origin request refused.' } 403
            continue
        }

        try {
            # Each branch ends in `break`: PowerShell's switch -Regex runs EVERY
            # matching branch otherwise, so /api/bootstrap would also fall into
            # the '^GET /api/.*$' and '^GET /.*$' catch-alls below.
            switch -Regex ("$verb $path") {
                '^GET /$'               { Send-File $ctx 'index.html'; break }
                '^POST /api/login$' {
                    $p = Read-Body $ctx
                    $u = [string]($p.username)
                    $pw = [string]($p.password)
                    if (-not $u -or -not $pw) { Send-Json $ctx @{ error = 'Username and password are required.' } 401; break }
                    if (-not $script:AuthUsers.ContainsKey($u) -or -not (Test-Password $script:AuthUsers[$u] $pw)) {
                        Send-Json $ctx @{ error = 'Invalid username or password.' } 401
                        break
                    }
                    $sid = [System.Guid]::NewGuid().ToString('N')
                    $script:AuthSessions[$sid] = @{ username = $u; name = $script:AuthUsers[$u].name; mustChangePassword = [bool]$script:AuthUsers[$u].mustChangePassword }
                    Set-SessionCookie $ctx $sid
                    Send-Json $ctx @{ ok = $true; user = @{ username = $u; name = $script:AuthUsers[$u].name; mustChangePassword = [bool]$script:AuthUsers[$u].mustChangePassword } }
                    break
                }
                '^POST /api/logout$' {
                    $sid = Get-SessionCookie $ctx.Request
                    if ($sid -and $script:AuthSessions.ContainsKey($sid)) { $script:AuthSessions.Remove($sid) }
                    Clear-SessionCookie $ctx
                    Send-Json $ctx @{ ok = $true }
                    break
                }
                '^GET /api/session$' {
                    try {
                        $session = Ensure-Authenticated $ctx
                        Send-Json $ctx @{ ok = $true; user = @{ username = $session.username; name = $session.name; mustChangePassword = [bool]$session.mustChangePassword } }
                    }
                    catch {
                        Send-Json $ctx @{ error = $_.Exception.Message } 401
                    }
                    break
                }
                '^POST /api/password/change$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    if (-not $session.mustChangePassword -and -not (Test-Password $script:AuthUsers[$session.username] ([string]$p.currentPassword))) { throw 'Current password is incorrect.' }
                    Assert-NewPassword ([string]$p.newPassword)
                    Set-AuthPassword $session.username ([string]$p.newPassword) $false
                    $session.mustChangePassword = $false
                    Send-Json $ctx @{ ok = $true }
                    break
                }
                '^GET /api/users$' {
                    $session = Ensure-Admin $ctx
                    Send-Json $ctx @{ users = @(Get-PublicAuthUsers) }
                    break
                }
                '^POST /api/users$' {
                    $session = Ensure-Admin $ctx
                    $p = Read-Body $ctx
                    $username = ([string]$p.username).Trim().ToUpperInvariant()
                    $password = [string]$p.password
                    $name = ([string]$p.name).Trim()
                    if ($username -notmatch '^[A-Z0-9._-]{2,40}$') { throw 'Username must be 2-40 letters, numbers, dots, hyphens, or underscores.' }
                    Assert-NewPassword $password
                    if (-not $name) { $name = $username }
                    if ($name.Length -gt 80) { throw 'Display name is too long.' }
                    foreach ($existing in $script:AuthUsers.Keys) {
                        if ($existing -ieq $username) { throw "The username '$username' already exists." }
                    }
                    $script:AuthUsers[$username] = @{ username = $username; name = $name } + (New-PasswordRecord $password $true)
                    Save-AuthUsers
                    Send-Json $ctx @{ ok = $true; users = @(Get-PublicAuthUsers) }
                    break
                }
                '^POST /api/users/reset-password$' {
                    $session = Ensure-Admin $ctx
                    $p = Read-Body $ctx
                    $username = ([string]$p.username).Trim()
                    $password = [string]$p.password
                    if (-not $script:AuthUsers.ContainsKey($username)) { throw "No user called '$username'." }
                    Assert-NewPassword $password
                    Set-AuthPassword $username $password $true
                    Send-Json $ctx @{ ok = $true; username = $username }
                    break
                }
                '^GET /api/bootstrap$'  {
                    $session = Ensure-Authenticated $ctx
                    Send-Json $ctx (Get-Bootstrap)
                    break
                }
                '^GET /api/week$'       {
                    $session = Ensure-Authenticated $ctx
                    Send-Json $ctx (Get-WeekReport $ctx.Request.QueryString['start'])
                    break
                }
                '^GET /api/all$'        {
                    $session = Ensure-Authenticated $ctx
                    Send-Json $ctx (Get-AllEntries)
                    break
                }

                '^POST /api/update$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    $r = Update-Entry $p $session.username
                    $out = Get-Bootstrap
                    $out.result = $r
                    Send-Json $ctx $out
                    break
                }

                '^POST /api/entry$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    $r = Add-Entry $p $session.username
                    $out = Get-Bootstrap
                    $out.result = $r
                    Send-Json $ctx $out
                    break
                }
                '^POST /api/option$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    $r = Add-Option $p.items
                    $out = Get-Bootstrap
                    $out.result = $r
                    Send-Json $ctx $out
                    break
                }
                '^POST /api/project$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    $r = Update-Projects $p
                    $out = Get-Bootstrap
                    $out.result = $r
                    Send-Json $ctx $out
                    break
                }
                '^POST /api/delete$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    $r = Remove-Entry ([int]$p.row)
                    $out = Get-Bootstrap
                    $out.result = $r
                    Send-Json $ctx $out
                    break
                }
                '^POST /api/shutdown$' {
                    $session = Ensure-Authenticated $ctx
                    # Lets "Stop Timesheet.cmd" end the server the tidy way:
                    # the finally block still saves the workbook and quits Excel.
                    Send-Json $ctx @{ ok = $true; message = 'Server shutting down.' }
                    $script:StopRequested = $true
                    break
                }
                '^GET /api/tasks$' {
                    $session = Ensure-Authenticated $ctx
                    $out = Get-Tasks
                    Send-Json $ctx $out
                    break
                }
                '^GET /api/team$' {
                    $session = Ensure-Authenticated $ctx
                    Send-Json $ctx (Get-Team)
                    break
                }
                '^GET /api/holidays$' {
                    $session = Ensure-Authenticated $ctx
                    Send-Json $ctx @{ holidays = @(Get-Holidays) }
                    break
                }
                '^GET /api/wallpaper$' {
                    try { Send-Response $ctx 200 'image/jpeg' (Get-Wallpaper) }
                    catch { Send-Json $ctx @{ error = 'Wallpaper unavailable' } 502 }
                    break
                }
                '^POST /api/holiday$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    $r = Update-Holidays $p
                    Send-Json $ctx @{ ok = $true; holidays = @(Get-Holidays); result = $r }
                    break
                }
                '^POST /api/team$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    $r = Update-Team $p
                    $out = Get-Team
                    $out.result = $r
                    Send-Json $ctx $out
                    break
                }
                '^POST /api/events$' {
                    $session = Ensure-Authenticated $ctx
                    $p = Read-Body $ctx
                    Write-Journal -Payload $p -Action 'task_event'
                    $board = (Get-Tasks).tasks
                    Update-TaskState $board $p
                    Save-Tasks $board
                    Send-Json $ctx @{ ok = $true; event = $p.type }
                    break
                }
                '^GET /api/.*$'  { Send-Json $ctx @{ error = 'Unknown endpoint' } 404; break }
                '^GET /.*$'      { Send-File $ctx ($path.TrimStart('/')); break }
                default          { Send-Json $ctx @{ error = 'Method not allowed' } 405 }
            }
            Write-Host ("  {0,-6} {1,-22} ok" -f $verb, $path) -ForegroundColor DarkGray
        }
        catch {
            $msg  = $_.Exception.Message
            $line = $_.InvocationInfo.ScriptLineNumber
            $at   = ($_.InvocationInfo.Line).Trim()
            Write-Host ("  {0,-6} {1,-22} ERROR {2}" -f $verb, $path, $msg) -ForegroundColor Red
            Write-Host ("         at line {0}: {1}" -f $line, $at) -ForegroundColor DarkRed
            try { Send-Json $ctx @{ error = $msg; line = $line; at = $at } 400 } catch { }
        }
    }
}
finally {
    Write-Host ''
    Write-Host '  Shutting down, saving workbook...' -ForegroundColor Yellow
    try { $listener.Stop(); $listener.Close() } catch { }
    Disconnect-Excel
    Write-Host '  Done.' -ForegroundColor Green
}
