# Exporter.ps1
# Exports productivity entries to CSV, Markdown, or a color-coded Excel log

function global:Get-ExportPivotRows {
    <#
        Shared pivot used by Export-EntriesToCsv and Export-EntriesToXlsx.

        Builds one row object per entry (sorted by date ascending) with columns
        Date, Days, one "<h> Hour" column per distinct goal-hour value anywhere
        in the dataset, then Note. Goal cells hold the display mark (check for
        done, cross otherwise) or "" when the entry has no goal for that hour.

        Also returns the hour columns, the date-sorted entries, and a per-entry
        Hours -> status map so the Excel path can color the exact cells without
        scanning $entry.Goals a second time.
    #>
    param([Parameter(Mandatory)][array]$Entries)

    $checkMark = [string][char]0x2713   # v
    $crossMark = [string][char]0x2717   # x

    # Every distinct goal-hour value that appears anywhere becomes its own column (e.g. "6 Hour", "9 Hour", "4 Hour")
    $allHours = @($Entries | ForEach-Object { $_.Goals } | ForEach-Object { $_.Hours } | Sort-Object -Unique)
    $sortedEntries = @($Entries | Sort-Object { [datetime]$_.Date })

    $rows = [System.Collections.Generic.List[object]]::new()
    $statusByHour = [System.Collections.Generic.List[object]]::new()

    foreach ($entry in $sortedEntries) {
        # Build a Hours -> Goal map once per entry and reuse it for every hour
        # column - eliminates the repeated Where-Object scan over $entry.Goals.
        # First match wins, mirroring the old `Where-Object { $_.Hours -eq $h }`
        # semantics (a later goal with the same hour is ignored).
        $goalByHour = @{}
        foreach ($g in @($entry.Goals)) {
            if ($null -ne $g -and $null -ne $g.Hours -and -not $goalByHour.ContainsKey($g.Hours)) {
                $goalByHour[$g.Hours] = $g
            }
        }

        $row = [ordered]@{
            Date = $entry.Date
            Days = "Day $($entry.DayNumber)"
        }
        $statuses = @{}
        foreach ($hours in $allHours) {
            $goal = $goalByHour[$hours]
            if ($goal) {
                # Same value rule as the old per-cell logic: "done" renders the
                # check mark, anything else renders the cross mark.
                $statuses[$hours] = $goal.Status
                $row["$hours Hour"] = if ($goal.Status -eq "done") { $checkMark } else { $crossMark }
            } else {
                $row["$hours Hour"] = ""
            }
        }
        $row["Note"] = $entry.Note
        $rows.Add([PSCustomObject]$row)
        $statusByHour.Add($statuses)
    }

    return @{
        Rows          = $rows
        AllHours      = $allHours
        SortedEntries = $sortedEntries
        StatusByHour  = $statusByHour
    }
}

function global:Get-ExcelColumnLetter {
    <#
        Converts a 1-based column index to its Excel column letters, e.g.
        1 -> A, 26 -> Z, 27 -> AA. The old [char](64 + $i) silently produced
        invalid references past column Z.
    #>
    param([int]$Index)
    $letters = ""
    $n = $Index
    while ($n -gt 0) {
        $n--
        $rem = $n % 26
        $letters = [char](65 + $rem) + $letters
        # Truncating integer division. Plain `[int](...)` would ROUND (not
        # truncate), and a bare [math]::Floor() would leak a [double] into the
        # next % (which then yields a [decimal] that [char] can't cast) - so
        # both a Floor and an immediate int cast are needed.
        $n = [int][math]::Floor($n / 26)
    }
    return $letters
}

function global:Export-EntriesToCsv {
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Path
    )

    $pivot = Get-ExportPivotRows -Entries $Entries
    $pivot.Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function global:Export-EntriesToMarkdown {
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Path
    )

    $lines = @("# Productivity log", "")
    foreach ($entry in ($Entries | Sort-Object { [datetime]$_.Date })) {
        $lines += "## Day $($entry.DayNumber) - $($entry.Date)"
        foreach ($goal in $entry.Goals) {
            $mark = switch ($goal.Status) {
                "done"   { "[x]" }
                "missed" { "[!]" }
                default  { "[ ]" }
            }
            $lines += "- $mark $($goal.Hours) hour goal"
        }
        if ($entry.Note) {
            $lines += ""
            $lines += "> $($entry.Note)"
        }
        $lines += ""
    }

    ($lines -join "`r`n") | Set-Content -Path $Path -Encoding UTF8
}

function global:Export-EntriesToXlsx {
    <#
        Exports the same pivoted layout as the CSV, but as a real .xlsx with
        goal cells directly colored green (done) / red (missed) - no conditional
        formatting rules involved, so there's nothing that can fail to match.
        Requires the ImportExcel module: Install-Module ImportExcel -Scope CurrentUser
    #>
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "ImportExcel module not found. Install it once with: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $greenColor = [System.Drawing.ColorTranslator]::FromHtml("#C6EFCE")
    $redColor   = [System.Drawing.ColorTranslator]::FromHtml("#FFC7CE")
    $blackColor = [System.Drawing.Color]::Black

    $pivot = Get-ExportPivotRows -Entries $Entries

    if (Test-Path $Path) { Remove-Item $Path -Force }

    $excel = $pivot.Rows | Export-Excel -Path $Path -WorksheetName "Log" -AutoSize -FreezeTopRow -BoldTopRow -PassThru
    $ws = $excel.Workbook.Worksheets["Log"]

    # Column layout: A=Date, B=Days, one column per goal hour starting at C, then Note.
    # Color each goal column's cells in contiguous same-status runs - one
    # Set-ExcelRange call per run instead of one per cell (the old per-cell loop
    # was the dominant cost for large entry counts). Cells whose entry has no
    # goal for that hour stay untouched, exactly as before; a goal that is present
    # is always colored (green when "done", red otherwise - mirroring the old
    # `else -> red` branch, so a goal with an unexpected status still gets red).
    $firstGoalCol = 3
    $statusByHour = $pivot.StatusByHour
    $n = $pivot.SortedEntries.Count
    for ($c = 0; $c -lt $pivot.AllHours.Count; $c++) {
        $colLetter = Get-ExcelColumnLetter -Index ($firstGoalCol + $c)
        $hours = $pivot.AllHours[$c]
        $i = 0
        while ($i -lt $n) {
            $stMap = $statusByHour[$i]
            if (-not $stMap.ContainsKey($hours)) { $i++; continue }
            $status = $stMap[$hours]
            $bg = if ($status -eq "done") { $greenColor } else { $redColor }
            # Extend the run while the next cell in this column has a goal with
            # the same status (a gap with no goal splits the run).
            $runEnd = $i
            while ($runEnd + 1 -lt $n -and
                   $statusByHour[$runEnd + 1].ContainsKey($hours) -and
                   $statusByHour[$runEnd + 1][$hours] -eq $status) { $runEnd++ }
            $cellRef = "$colLetter$($i + 2):$colLetter$($runEnd + 2)"
            Set-ExcelRange -Worksheet $ws -Range $cellRef -BackgroundColor $bg -FontColor $blackColor -HorizontalAlignment Center
            $i = $runEnd + 1
        }
    }

    Close-ExcelPackage $excel
}
