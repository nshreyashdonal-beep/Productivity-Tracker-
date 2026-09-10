# Main.ps1
# Entry point for the Productivity Tracker (WPF + PowerShell)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Windows 11 DWM API: enables rounded corners on borderless windows
# without AllowsTransparency=True.
Add-Type -TypeDefinition '
using System;
using System.Runtime.InteropServices;
public class DwmApi {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(
        IntPtr hwnd, int attr, IntPtr attrValue, int attrSize);
}
' -ErrorAction SilentlyContinue

# App root directory. $PSScriptRoot is empty when running as a PS2EXE-compiled
# exe, so fall back to the exe's own folder (where the XAML and loose scripts
# live). Raw-script and compiled-exe launches then find the same files.
$root = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}
. (Join-Path $root "DataStore.ps1")
. (Join-Path $root "Exporter.ps1")

# ---------- Load window ----------
[xml]$xaml = Get-Content -Path (Join-Path $root "MainWindow.xaml") -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# App icon (title bar / taskbar). A relative Icon= URI can't resolve in loose
# XAML (WPF treats it as an application resource), so set it from an absolute
# path here. Prefer Assets\app.ico, else the newest .ico in Assets. A missing
# or corrupt icon must not crash the app - fall back to the default.
$iconPath = Join-Path $root "Assets\app.ico"
if (-not (Test-Path $iconPath)) {
    $candidate = Get-ChildItem (Join-Path $root "Assets") -Filter "*.ico" -File |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($candidate) { $iconPath = $candidate.FullName }
}
if (Test-Path $iconPath) {
    try {
        $window.Icon = [System.Windows.Media.Imaging.BitmapImage]::new([System.Uri]$iconPath)
    } catch {
        Write-Host "Warning: could not load app icon from $iconPath - $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ---------- Header ----------
$JourneyNameText    = $window.FindName("JourneyNameText")
$CreateJourneyBtn   = $window.FindName("CreateJourneyBtn")
$EditTemplateBtn    = $window.FindName("EditTemplateBtn")
$global:AddDayBtn   = $window.FindName("AddDayBtn")
$global:SyncBtn     = $window.FindName("SyncBtn")
$SyncStatusText     = $window.FindName("SyncStatusText")

# ---------- Sidebar ----------
$SidebarColumn        = $window.FindName("SidebarColumn")
$Sidebar              = $window.FindName("Sidebar")
$CollapseSidebarBtn   = $window.FindName("CollapseSidebarBtn")
$CollapseLabel        = $window.FindName("CollapseLabel")
# Maximize/Focus layout elements (see Toggle-FocusMaximize)
$HeaderStack          = $window.FindName("HeaderStack")
$HeaderDivider        = $window.FindName("HeaderDivider")
$SidebarSplitterCol   = $window.FindName("SidebarSplitterCol")
$SidebarSplitter      = $window.FindName("SidebarSplitter")

$DashboardHeaderBtn   = $window.FindName("DashboardHeaderBtn")
$DashboardIcon        = $window.FindName("DashboardIcon")
$DashboardLabel       = $window.FindName("DashboardLabel")
$DashboardSubs        = $window.FindName("DashboardSubs")
$FilterAllBtn         = $window.FindName("FilterAllBtn")
$FilterCompletedBtn   = $window.FindName("FilterCompletedBtn")
$FilterMissedBtn      = $window.FindName("FilterMissedBtn")
$ReportBtn            = $window.FindName("ReportBtn")

$ExportHeaderBtn      = $window.FindName("ExportHeaderBtn")
$ExportIcon           = $window.FindName("ExportIcon")
$ExportLabel          = $window.FindName("ExportLabel")
$ExportSubs           = $window.FindName("ExportSubs")
$ExportBtn            = $window.FindName("ExportBtn")

$SwitchJourneyBtn     = $window.FindName("SwitchJourneyBtn")   # sidebar section header (Border)
$SwitchIcon           = $window.FindName("SwitchIcon")
$SwitchLabel          = $window.FindName("SwitchLabel")
$SwitchJourneySubs    = $window.FindName("SwitchJourneySubs")
$SwitchJourneyListPanel = $window.FindName("SwitchJourneyListPanel")

$BackupHeaderBtn      = $window.FindName("BackupHeaderBtn")   # sidebar section header (Border)
$BackupIcon           = $window.FindName("BackupIcon")
$BackupLabel          = $window.FindName("BackupLabel")
$BackupSubs           = $window.FindName("BackupSubs")
$SetBackupFolderBtn   = $window.FindName("SetBackupFolderBtn")
$ImportBtn            = $window.FindName("ImportBtn")

$LoadArchiveBtn       = $window.FindName("LoadArchiveBtn")     # sidebar leaf (Border)
$ArchiveIcon          = $window.FindName("ArchiveIcon")
$ArchiveLabel         = $window.FindName("ArchiveLabel")

$FocusModeBtn         = $window.FindName("FocusModeBtn")       # sidebar leaf (Border)
$FocusIcon            = $window.FindName("FocusIcon")
# NB: named FocusSidebarLabel (not $FocusLabel) - under `powershell -File` the script
# scope is the global scope, so an unprefixed $FocusLabel would collide with
# $global:FocusLabel (the session-title string) and get overwritten to a string,
# breaking Set-SidebarNarrow.
$FocusSidebarLabel    = $window.FindName("FocusLabel")

$CollapseIcon         = $window.FindName("CollapseIcon")

# ---------- Content area ----------
$ReadOnlyBanner     = $window.FindName("ReadOnlyBanner")
$ReadOnlyBannerText = $window.FindName("ReadOnlyBannerText")
$BackToCurrentBtn   = $window.FindName("BackToCurrentBtn")

$StartJourneyPanel  = $window.FindName("StartJourneyPanel")
$JourneyPanelTitle  = $window.FindName("JourneyPanelTitle")
$JourneyNameBox     = $window.FindName("JourneyNameBox")
$JourneyDateRow     = $window.FindName("JourneyDateRow")
$JourneyDatePicker  = $window.FindName("JourneyDatePicker")
$GoalEditPanel      = $window.FindName("GoalEditPanel")
$NewGoalHoursBox    = $window.FindName("NewGoalHoursBox")
$AddGoalBtn         = $window.FindName("AddGoalBtn")
$SetTemplateBtn     = $window.FindName("SetTemplateBtn")
$CancelJourneyBtn   = $window.FindName("CancelJourneyBtn")

$DaysScrollViewer   = $window.FindName("DaysScrollViewer")
$DaysList           = $window.FindName("DaysList")
$EmptyStateText     = $window.FindName("EmptyStateText")

$BulkDeleteToggleBtn   = $window.FindName("BulkDeleteToggleBtn")
$BulkDeleteBanner      = $window.FindName("BulkDeleteBanner")
$BulkDeleteCountText   = $window.FindName("BulkDeleteCountText")
$BulkDeleteConfirmBtn  = $window.FindName("BulkDeleteConfirmBtn")
$BulkDeleteCancelBtn   = $window.FindName("BulkDeleteCancelBtn")

$InfoAreaScrollViewer = $window.FindName("InfoAreaScrollViewer")
$InfoAreaPanel        = $window.FindName("InfoAreaPanel")

$FocusModePanel       = $window.FindName("FocusModePanel")
$FocusPomoBtn         = $window.FindName("FocusPomoBtn")
$FocusStopBtn         = $window.FindName("FocusStopBtn")
$FocusLabelBox        = $window.FindName("FocusLabelBox")
$FocusLabelHint       = $window.FindName("FocusLabelHint")
$FocusTimerCircle     = $window.FindName("FocusTimerCircle")
$FocusTimerText       = $window.FindName("FocusTimerText")
$FocusMaximizeBtn     = $window.FindName("FocusMaximizeBtn")
$FocusButtonPanel     = $window.FindName("FocusButtonPanel")
$FocusDurationEditPanel = $window.FindName("FocusDurationEditPanel")
$FocusDurationBox     = $window.FindName("FocusDurationBox")
$FocusDurationOkBtn   = $window.FindName("FocusDurationOkBtn")
$FocusDurationCancelBtn = $window.FindName("FocusDurationCancelBtn")
$FocusLogTargetText   = $window.FindName("FocusLogTargetText")

# ---------- State ----------
$global:Entries          = Import-Entries
$global:Settings         = Import-Settings

# Fill in empty rows for any day skipped between the last time the app was
# opened and today (e.g. last opened Aug 20, opens again Aug 23 - Aug 21 and
# Aug 22 get created here as ordinary missed-goal days). Today itself is
# never created by this - only ADD DAY or starting a Focus session does that.
$global:Entries = Backfill-SkippedDays -Entries $global:Entries -Settings $global:Settings
Save-Entries -Entries $global:Entries

# ---------- Backup folder ----------
# Defaults to the app's own Backup\ folder, but the user can point it anywhere
# via SET BACKUP FOLDER; that choice is persisted in settings.json so a reinstall
# can re-point to the same location and restore everything with IMPORT.
$global:BackupDir = if ($global:Settings.BackupFolder) { $global:Settings.BackupFolder } else { Join-Path $root "Backup" }

$global:Filter           = "All"
$global:ExpandedIds      = @{}
# Rebuilt on every Render-Days; lets a single goal toggle update just the
# affected chips + fraction instead of rebuilding the whole list.
$global:ChipRegistry     = @{}   # "EntryId|Hours" -> @(chip Borders)
$global:CountRegistry    = @{}   # EntryId -> fraction TextBlock
$global:ExpandRegistry   = @{}   # EntryId -> expanded-section container (toggleable visibility)
$global:StackRegistry    = @{}   # EntryId -> row root StackPanel (so lazy expand can append)
$global:RowRegistry      = @{}   # EntryId -> row root Border (so filters/add/delete can show/hide/remove rows without rebuilding)
$global:DayLabelRegistry = @{}   # EntryId -> "DAY N" TextBlock (so delete can renumber in place)

# Virtualized day list: the ItemsControl's ItemsSource. Only the rows near the
# scroll position are ever realized (see Populate-RealizedRows), so the registries
# above only hold realized rows. Kept observable so add/delete/filter update WPF.
$global:DayItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$global:DaysVSP  = $null          # cached VirtualizingStackPanel under $DaysList (found lazily)
$global:RealizedIds = @{}         # EntryIds currently realized; used to prune registries for recycled rows
$DaysList.ItemsSource = $global:DayItems

$global:PendingGoalHours = @()
$global:JourneyPanelMode = "New"   # "New" (setting up a journey) or "Edit" (editing the goal template)
$global:ReadOnlyMode     = $false  # $true while viewing an archived journey via LOAD ARCHIVE
$global:LiveEntries      = $global:Entries   # the real, editable Data/entries.json contents
$global:LiveSettings     = $global:Settings  # the real, editable Data/settings.json contents
$global:SidebarNarrow    = $false
$global:ActiveSidebarSection = "dash"        # "dash" | "export" | "switch" | "archive"
$global:SyncPending = $false                 # $true when changes exist since the last manual SYNC
# Bulk-select delete mode. Selection is keyed by Entry.Id so it stays
# correct through container recycling in the virtualized list.
$global:SelectModeActive = $false
$global:SelectedEntryIds = @{}

# Focus Mode session state. A "started" session (running OR paused, i.e.
# $FocusElapsedSec > 0) locks the mode toggle and the focus-target label.
$global:FocusMode      = "Pomo"     # "Pomo" | "Stopwatch"
$global:FocusRunning   = $false
$global:FocusElapsedSec = 0
$global:FocusLabel     = ""
$global:FocusLabelWasEmpty = $true  # last known empty-state of the label box (drives watermark toggle)
$global:FocusDurationMin = 25       # working Pomodoro length (minutes); editable via the timer circle
# Session persistence: an in-progress session is written to Data/focus-session.json
# on every start/pause/resume, on window close, and every 5s while running (so a
# task-kill loses at most ~5s of elapsed time - a hard kill can't be handled
# gracefully, but periodic writes make the loss negligible). LastActive lets a
# resume credit time the app was closed toward a clock that was still "running".
$global:FocusStateFile = Join-Path $root "Data/focus-session.json"
$global:FocusPersistCounter = 0
# Timestamp-based session model: the persisted file is the source of truth (never
# the tick counter - see Get-FocusElapsedSec). startTime is the last resume moment
# (null while paused); elapsedBankedSec holds the completed run segments.
$global:FocusStartTime = $null
$global:FocusBankedSec = 0
$global:FocusStatus    = "RUNNING"   # "RUNNING" | "PAUSED" | "COMPLETED"
# Completion notification (on-screen toast popup): a small always-on-top WPF window
# appears in the top-right of the screen even while the main window is minimized.
# It stays up until the user clicks OK (or the app closes).
$global:FocusToastWindow = $null
# Maximized Focus view: hides the header + sidebar so the Focus panel fills the
# whole window. Saved sidebar geometry lets Restore-FocusLayout put it back exactly
# (including the user's narrow/wide choice).
$global:FocusMaximized = $false
$global:SavedSidebarWidth  = "190"
$global:SavedSidebarMin    = "52"
$global:SavedSidebarMax    = "260"
# Live day-row registries for focus sessions. Rebuilt on every Render-Days.
$global:TotalRegistry   = @{}   # EntryId -> collapsed "Total Hour : Xh Ym" TextBlock
$global:FocusRegistry   = @{}   # EntryId -> @{ Timeline; Total } for the expanded section
# Excel copy is deferred until ~2s after the last edit so disk I/O is batched.
$global:AutoSyncDebounceTimer = New-Object System.Windows.Threading.DispatcherTimer
$global:AutoSyncDebounceTimer.Interval = [TimeSpan]::FromSeconds(2)
$global:AutoSyncDebounceTimer.Add_Tick({
    # Fired ~2s after the last edit: run the deferred Excel backup.
    Write-LiveExcel
}.GetNewClosure())

# Debounces note saving: typing only touches the TextBox (no per-keystroke writes).
# When the note box loses focus we update the in-memory note immediately and defer
# the disk writes ~1s so blur never blocks the UI thread - same idea as the Excel
# debounce above. $global:NoteDirty tracks whether a save is actually pending.
$global:NoteDirty = $false
$global:NoteDebounceTimer = New-Object System.Windows.Threading.DispatcherTimer
$global:NoteDebounceTimer.Interval = [TimeSpan]::FromSeconds(1)
$global:NoteDebounceTimer.Add_Tick({
    $global:NoteDebounceTimer.Stop()
    if ($global:NoteDirty) {
        $global:NoteDirty = $false
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync
    }
}.GetNewClosure())

# Debounces goal-chip saves: toggling a chip updates the in-memory state
# immediately (UI is instant) but defers the disk writes ~1s so rapid
# clicks batch into one write instead of one per click.
$global:GoalDirty = $false
$global:GoalDebounceTimer = New-Object System.Windows.Threading.DispatcherTimer
$global:GoalDebounceTimer.Interval = [TimeSpan]::FromSeconds(1)
$global:GoalDebounceTimer.Add_Tick({
    $global:GoalDebounceTimer.Stop()
    if ($global:GoalDirty) {
        $global:GoalDirty = $false
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync
    }
}.GetNewClosure())

# Focus Mode clock: ticks once per second while a session is running. Pomodoro
# counts down from 25:00 and auto-saves when it reaches 0; Stopwatch counts up.
# The tick handler runs on the UI thread, so the timer can keep running (and
# keep accruing elapsed time) even while the Focus panel is hidden - navigating
# away from the Focus view no longer stops it (see Show-ContentArea).
$global:FocusTimer = New-Object System.Windows.Threading.DispatcherTimer
$global:FocusTimer.Interval = [TimeSpan]::FromSeconds(1)
$global:FocusTimer.Add_Tick({
    if (-not $global:FocusRunning) { return }
    # Elapsed is recomputed from the persisted startTime + banked, never from the
    # tick count - a blocked UI thread, sleep, or clock change can't make the
    # display drift; the tick is only the clock's heartbeat.
    $global:FocusElapsedSec = Get-FocusElapsedSec
    if ($global:FocusMode -eq "Pomo") {
        $secsLeft = [Math]::Max(0, ($global:FocusDurationMin * 60) - $global:FocusElapsedSec)
        $FocusTimerText.Text = Format-FocusClock -TotalSec $secsLeft -Mode $global:FocusMode
        if ($secsLeft -le 0) {
            # Pomodoro finished naturally -> stop the session, save it, and raise
            # the background completion notification. Read the label straight from
            # the box: Save-FocusSession resets $global:FocusLabel, so capturing
            # the global here would toast the stale (usually empty) label.
            $doneLabel = $FocusLabelBox.Text
            $doneMin   = $global:FocusDurationMin
            # Do NOT flip $global:FocusRunning here. Save-FocusSession recomputes
            # elapsed from Get-FocusElapsedSec, which only credits the current run
            # segment while FocusRunning is true - pre-stopping it made elapsed
            # compute to 0 (banked-only) and the completed session got DISCARDED
            # instead of recorded. Save-FocusSession stops the timer and clears
            # running itself after computing elapsed.
            $global:FocusTimer.Stop()
            Save-FocusSession
            Show-FocusCompletionNotification -Label $doneLabel -DurationMin $doneMin
            return
        }
    } else {
        $FocusTimerText.Text = Format-FocusClock -TotalSec $global:FocusElapsedSec -Mode $global:FocusMode
    }
    # Belt-and-suspenders periodic write; startTime/elapsedBankedSec are already
    # saved at every state change, so a task-kill (which bypasses Add_Closing) can
    # only ever lose a few seconds of a *running* segment at worst.
    if (++$global:FocusPersistCounter -ge 5) {
        $global:FocusPersistCounter = 0
        Save-FocusSessionState
    }
}.GetNewClosure())

$window.Add_Closing({
    param($s, $e)
    # Flush any pending debounced Excel write so closing the app never loses the
    # latest live.xlsx, even if the 2s window hadn't elapsed yet.
    Flush-AutoSync
    # Also flush any pending debounced note save so a note typed just before close
    # is still persisted.
    Flush-NoteSave
    # Flush any pending debounced goal-chip save so toggling a chip just before
    # close is still persisted.
    Flush-GoalSave
    # Persist an in-progress focus session so it survives the restart. A RUNNING
    # session's startTime was written at Start/Resume, so the elapsed gap is
    # measured from THAT moment (the clock kept running while the app was closed).
    Save-FocusSessionState
    # Drop any lingering completion-notification tray icon (non-blocking).
    Cleanup-FocusNotification
    # Stop the Focus clock so it can't keep firing during shutdown.
    $global:FocusTimer.Stop()
}.GetNewClosure())

# ---------- Colors (logbook theme) ----------
$bc = [System.Windows.Media.BrushConverter]::new()
$InkBrush      = $bc.ConvertFromString("#2B3A55")
$InkColor      = [System.Windows.Media.Color]::FromRgb(43, 58, 85)
$InkLightBrush = $bc.ConvertFromString("#5A6B8C")
$PaperBrush    = $bc.ConvertFromString("#F6F1E4")
$PlusBlueBrush = $bc.ConvertFromString("#2B3A55")   # navy "+" vector icon (matches InkBrush)
$RowBrush      = $bc.ConvertFromString("#FFFDF7")
$HoverNavyBrush  = $bc.ConvertFromString("#38507A")
$HoverCreamBrush = $bc.ConvertFromString("#FBF3DE")
$IconHoverBrush  = $bc.ConvertFromString("#F0E8DC")   # index2.html .icon-danger-btn hover
$ErrorBrush      = $bc.ConvertFromString("#993C1D")   # index2.html modalError
$DividerDotBrush = $bc.ConvertFromString("#C9BFA5")
$GreenBrush    = $bc.ConvertFromString("#2F6E4F")
$GreenBgBrush  = $bc.ConvertFromString("#E3EEE4")
$RedBrush      = $bc.ConvertFromString("#A13B2E")
$AmberBrush    = $bc.ConvertFromString("#B8860B")
$RedBgBrush    = $bc.ConvertFromString("#F4E3DE")
$DividerBrush  = $bc.ConvertFromString("#D9D2BC")
$HighlightBrush = $bc.ConvertFromString("#FBF3DE")
$TransparentBrush = $bc.ConvertFromString("#00000000")

# Cached chip brushes/glyphs - avoids re-parsing a color string (and allocating a
# new SolidColorBrush) for every goal chip on every render.
$WhiteBrush     = $bc.ConvertFromString("#FFFFFF")
$ChipDoneBg     = $bc.ConvertFromString("#EDF7E8")
$ChipDoneBd     = $bc.ConvertFromString("#6B9955")
$ChipDoneFg     = $bc.ConvertFromString("#6B9955")
$ChipMissBg     = $bc.ConvertFromString("#FFFFFF")
$ChipMissBd     = $bc.ConvertFromString("#C55A47")
$ChipMissFg     = $bc.ConvertFromString("#C55A47")
$CheckGlyph     = [string][char]0x2713
$CrossGlyph     = [string][char]0x2717
$DeleteGlyph    = [string][char]0x2715
$ClockGlyph     = [string][char]0x23F1   # stopwatch (BMP, renders in Consolas)

function global:Match-Filter {
    # Pure predicate: does $Entry match the current $global:Filter?
    # "Completed" = has goals and all are done; "Missed" = at least one missed.
    # Tight foreach (not Where-Object pipelines) - this runs per entry on every
    # filter switch, so its cost matters.
    param($Entry)
    switch ($global:Filter) {
        "Completed" {
            $total = 0; $done = 0
            foreach ($g in $Entry.Goals) { $total++; if ($g.Status -eq "done") { $done++ } }
            return ($total -gt 0 -and $done -eq $total)
        }
        "Missed" {
            foreach ($g in $Entry.Goals) { if ($g.Status -eq "missed") { return $true } }
            return $false
        }
        default { return $true }
    }
}

function global:Get-FilteredSortedEntries {
    # The virtualized list's backing order: entries newest-first, restricted to
    # the current filter. Only realized rows get controls; everything else is just
    # data waiting to be scrolled into view.
    $sorted = @($global:Entries | Sort-Object -Property @{ Expression = { [datetime]$_.Date } } -Descending)
    if ($global:Filter -eq "All") { return $sorted }
    $out = @()
    foreach ($e in $sorted) {
        if (Match-Filter -Entry $e) { $out += $e }
    }
    return $out
}

function global:Update-EmptyState {
    # Centralized empty-state text/visibility for the virtualized day list. Cheap:
    # DayItems holds exactly the currently-filtered set, so its Count is the truth.
    if ($global:DayItems.Count -eq 0) {
        $EmptyStateText.Visibility = "Visible"
        if (@($global:Entries).Count -eq 0) {
            if ($global:Settings.StartDate) {
                $EmptyStateText.Text = "Journey set. Click + ADD DAY to log your first day."
            } else {
                $EmptyStateText.Text = "Click CREATE NEW JOURNEY to set your start date and goals."
            }
        } else {
            $EmptyStateText.Text = "Nothing matches this filter."
        }
    } else {
        $EmptyStateText.Visibility = "Collapsed"
    }
}

function global:Update-FilterView {
    # Reconciles the virtualized ItemsSource (DayItems) with the filtered,
    # newest-first entry set, using in-place add/remove so the scroll position is
    # preserved where possible. Realized rows are untouched (their content is keyed
    # to the item, not the index); unrealized rows are just data. This replaces the
    # old visibility-toggle fast path, which can't work once only a window of rows
    # is materialized.

    # Keep the journey name header in sync — this function runs on every filter
    # change and after journey switches, where Render-Days is not always reached.
    if ($global:Settings.JourneyName) {
        $JourneyNameText.Text = "Journey: $($global:Settings.JourneyName)"
        $JourneyNameText.Visibility = "Visible"
    } else {
        $JourneyNameText.Visibility = "Collapsed"
    }

    $desired = Get-FilteredSortedEntries
    $desiredSet = @{}
    foreach ($e in $desired) { $desiredSet[$e.Id] = $true }

    # Drop items that no longer match the filter (also forget their expand state,
    # matching the old behavior of collapsing hidden rows).
    for ($i = $global:DayItems.Count - 1; $i -ge 0; $i--) {
        if (-not $desiredSet.ContainsKey($global:DayItems[$i].Id)) {
            $global:ExpandedIds.Remove($global:DayItems[$i].Id)
            $global:DayItems.RemoveAt($i)
        }
    }
    # Ensure the survivors are present, in the right order (newest first).
    $idx = 0
    foreach ($e in $desired) {
        $pos = -1
        for ($j = 0; $j -lt $global:DayItems.Count; $j++) {
            if ($global:DayItems[$j].Id -eq $e.Id) { $pos = $j; break }
        }
        if ($pos -lt 0) {
            $global:DayItems.Insert($idx, $e)
        } elseif ($pos -ne $idx) {
            $global:DayItems.RemoveAt($pos)
            $global:DayItems.Insert($idx, $e)
        }
        $idx++
    }
    Update-EmptyState
}

function global:New-GoalChip {
    param($Goal, [string]$EntryId)

    $border = [System.Windows.Controls.Border]::new()
    $border.CornerRadius   = 16
    $border.Padding        = "12,4,12,4"
    $border.Margin         = "0,0,6,0"
    $border.BorderThickness = 1.5
    $border.Tag            = "markfirst"   # collapsed chips read "✓ 4h" (mark first)

    $label = "$($Goal.Hours)h"
    switch ($Goal.Status) {
        "done"   {
            $border.Background = $ChipDoneBg
            $border.BorderBrush = $ChipDoneBd
            $textForeground = $ChipDoneFg
            $label = "$CheckGlyph $($Goal.Hours)h"
        }
        "missed" {
            $border.Background = $ChipMissBg
            $border.BorderBrush = $ChipMissBd
            $textForeground = $ChipMissFg
            $label = "$CrossGlyph $($Goal.Hours)h"
        }
        default  {
            $border.Background = $WhiteBrush
            $border.BorderBrush = $InkBrush
            $textForeground = $InkBrush
        }
    }

    $text = [System.Windows.Controls.TextBlock]::new()
    $text.Text       = $label
    $text.FontFamily = "Consolas"
    $text.FontSize   = 12
    $text.Foreground = $textForeground
    $border.Child    = $text

    # Register so a goal toggle can restyle this chip in place.
    Register-Chip -EntryId $EntryId -Hours $Goal.Hours -Border $border

    return $border
}

function global:Register-Chip {
    # Tracks every chip (header + body) for a given goal so a status toggle can
    # restyle it in place. Key is "EntryId|Hours".
    param([string]$EntryId, [int]$Hours, $Border)
    $key = "$EntryId|$Hours"
    if (-not $global:ChipRegistry.ContainsKey($key)) {
        $global:ChipRegistry[$key] = @()
    }
    $global:ChipRegistry[$key] = @($global:ChipRegistry[$key]) + $Border
}

function global:Register-Count {
    # Tracks the "done/total" fraction TextBlock for a row so it can be refreshed
    # in place after a goal toggle.
    param([string]$EntryId, $TextBlock)
    $global:CountRegistry[$EntryId] = $TextBlock
}

function global:Set-GoalChipVisual {
    # Re-applies just the status-dependent visuals of an existing chip border
    # (background / border / label / color) without rebuilding it. Keeps the
    # done/missed colors identical to the ones used at construction time.
    # Chip label format follows construction: collapsed chips are "✓ 4h"
    # (Tag "markfirst"), big chips are "4h ✓" (Tag "hoursfirst").
    param($Border, [string]$Status, [int]$Hours)
    $markFirst = ($Border.Tag -eq "markfirst")
    switch ($Status) {
        "done"   {
            $Border.Background = $ChipDoneBg; $Border.BorderBrush = $ChipDoneBd; $fg = $ChipDoneFg
            $label = if ($markFirst) { "$CheckGlyph ${Hours}h" } else { "${Hours}h $CheckGlyph" }
        }
        "missed" {
            $Border.Background = $ChipMissBg; $Border.BorderBrush = $ChipMissBd; $fg = $ChipMissFg
            $label = if ($markFirst) { "$CrossGlyph ${Hours}h" } else { "${Hours}h $CrossGlyph" }
        }
        default  { $Border.Background = $WhiteBrush; $Border.BorderBrush = $InkBrush; $fg = $InkBrush; $label = "${Hours}h" }
    }
    $tb = $Border.Child
    if ($tb -is [System.Windows.Controls.TextBlock]) {
        $tb.Text = $label
        $tb.Foreground = $fg
    }
}

function global:Update-GoalChipsInPlace {
    # Post-toggle fast path (filter "All" only): restyles the affected chips and
    # refreshes the done/total fraction for one row. Falls back to a full
    # Render-Days if the entry can no longer be found.
    param([string]$EntryId, [int]$Hours)
    $entry = $global:Entries | Where-Object { $_.Id -eq $EntryId }
    if (-not $entry) { Render-Days; return }

    $goal = $entry.Goals | Where-Object { $_.Hours -eq $Hours }
    $status = if ($goal) { $goal.Status } else { "missed" }

    $key = "$EntryId|$Hours"
    if ($global:ChipRegistry.ContainsKey($key)) {
        foreach ($border in $global:ChipRegistry[$key]) {
            Set-GoalChipVisual -Border $border -Status $status -Hours $Hours
        }
    }

    if ($global:CountRegistry.ContainsKey($EntryId)) {
        $done = 0; $total = 0
        foreach ($g in $entry.Goals) { $total++; if ($g.Status -eq "done") { $done++ } }
        $global:CountRegistry[$EntryId].Text = "$done/$total"
    }
}

function global:New-ExpandedSection {
    <#
        Builds (once, lazily) the expandable body of a day row: a divider + the
        big interactive goal chips + the note box. Returns a container whose
        Visibility is toggled on expand/collapse so the app never has to rebuild
        the whole list. Registers the container in ExpandRegistry.
    #>
    param($Entry)

    $region = [System.Windows.Controls.StackPanel]::new()

    $divider = [System.Windows.Controls.Border]::new()
    $divider.BorderBrush = $InkBrush
    $divider.BorderThickness = "0,1,0,0"
    $divider.Margin = "0,0,0,0"
    $divider.SnapsToDevicePixels = $true
    [void]$region.Children.Add($divider)

    $bodyStack = [System.Windows.Controls.StackPanel]::new()
    $bodyStack.Margin = "14,4,14,14"

    $tipLabel = [System.Windows.Controls.TextBlock]::new()
    $tipLabel.Text = "Tap a goal to mark done/missed:"
    $tipLabel.FontFamily = "Consolas"
    $tipLabel.FontSize = 11
    $tipLabel.Foreground = $InkLightBrush
    $tipLabel.Margin = "0,0,0,6"
    [void]$bodyStack.Children.Add($tipLabel)

    $bigChipsPanel = [System.Windows.Controls.WrapPanel]::new()
    foreach ($goal in $Entry.Goals) {
        $bigChip = [System.Windows.Controls.Border]::new()
        $bigChip.CornerRadius = 20
        $bigChip.Padding = "16,6,16,6"
        $bigChip.Margin = "0,10,8,6"
        $bigChip.BorderThickness = 1.5
        $bigChip.Cursor = "Hand"
        $bigChip.Tag = "hoursfirst"   # big chips read "4h ✓" (hours first)

        $label = "$($goal.Hours)h"
        switch ($goal.Status) {
            "done"   {
                $bigChip.Background   = $ChipDoneBg
                $bigChip.BorderBrush  = $ChipDoneBd
                $textForeground       = $ChipDoneFg
                $label = "$($goal.Hours)h $CheckGlyph"
            }
            "missed" {
                $bigChip.Background   = $ChipMissBg
                $bigChip.BorderBrush  = $ChipMissBd
                $textForeground       = $ChipMissFg
                $label = "$($goal.Hours)h $CrossGlyph"
            }
            default  {
                $bigChip.Background   = $WhiteBrush
                $bigChip.BorderBrush  = $ChipMissBd
                $textForeground       = $ChipMissFg
            }
        }

        $text = [System.Windows.Controls.TextBlock]::new()
        $text.Text = $label
        $text.FontFamily = "Consolas"
        $text.FontSize = 13
        $text.Foreground = $textForeground
        $bigChip.Child = $text

        # Register so a goal toggle can restyle this chip in place.
        Register-Chip -EntryId $Entry.Id -Hours $goal.Hours -Border $bigChip

        $bigChip.Add_MouseLeftButtonUp({
            param($s, $e)
            $e.Handled = $true
            if ($global:ReadOnlyMode) { return }
            $global:Entries = Set-GoalStatus -Entries $global:Entries -Id $Entry.Id -Hours $goal.Hours
            $global:LiveEntries = $global:Entries
            # Defer disk writes ~1s so rapid toggles batch into one save.
            $global:GoalDirty = $true
            $global:GoalDebounceTimer.Stop()
            $global:GoalDebounceTimer.Start()
            # Always update the chips in place; with a filter active the row may
            # enter/leave the set, so adjust its visibility too (no full rebuild).
            Update-GoalChipsInPlace -EntryId $Entry.Id -Hours $goal.Hours
            if ($global:Filter -ne "All") {
                Update-FilterView
            }
        }.GetNewClosure())

        [void]$bigChipsPanel.Children.Add($bigChip)
    }
    [void]$bodyStack.Children.Add($bigChipsPanel)

    # FOCUS RECORD - one row per logged session (built once; refreshed in place
    # by Update-FocusViewsInPlace when a new session is saved). A small "+" on the
    # header opens the manual-log dialog (retroactive entry, no timer runs).
    $focusHeaderPanel = [System.Windows.Controls.StackPanel]::new()
    $focusHeaderPanel.Orientation = "Horizontal"
    # Top margin separates this section from the goal chips above; the bottom
    # margin gives room before the dashed divider that sits under the header.
    $focusHeaderPanel.Margin = "0,14,0,14"
    $focusHeaderText = [System.Windows.Controls.TextBlock]::new()
    $focusHeaderText.Text = "Focus Record"
    $focusHeaderText.FontFamily = "Consolas"
    $focusHeaderText.FontSize = 20
    $focusHeaderText.FontWeight = "Bold"
    $focusHeaderText.Foreground = $InkBrush
    [void]$focusHeaderPanel.Children.Add($focusHeaderText)

    # Manual add only makes sense for a day that has already happened, so the "+"
    # is restricted to today and earlier (a future blank day has no session to log
    # retroactively) and is hidden in read-only archive view.
    $focusRowDate = [datetime]::ParseExact($Entry.Date, "yyyy-MM-dd", $null)
    if ($focusRowDate -le (Get-Date).Date -and -not $global:ReadOnlyMode) {
        $addFocusBtn = [System.Windows.Controls.Border]::new()
        $addFocusBtn.Width = 26
        $addFocusBtn.Height = 26
        $addFocusBtn.CornerRadius = 13
        $addFocusBtn.Background = $InkBrush
        $addFocusBtn.Cursor = "Hand"
        $addFocusBtn.ToolTip = "Add Focus Record"
        $addFocusBtn.Margin = "12,0,0,0"
        $addFocusBtn.VerticalAlignment = "Center"
        # SVG vector plus: two crossed lines inside a Viewbox, crisp at any size.
        $plusViewbox = [System.Windows.Controls.Viewbox]::new()
        $plusViewbox.Width = 18; $plusViewbox.Height = 18
        $plusViewbox.Stretch = "Uniform"
        $plusCanvas = [System.Windows.Controls.Canvas]::new()
        $plusCanvas.Width = 24; $plusCanvas.Height = 24
        $plusH = [System.Windows.Shapes.Path]::new()
        $plusH.Stroke = $PaperBrush; $plusH.StrokeThickness = 2.5
        $plusH.StrokeStartLineCap = "Round"; $plusH.StrokeEndLineCap = "Round"
        $plusH.Data = [System.Windows.Media.StreamGeometry]::Parse("M6,12 L18,12")
        [void]$plusCanvas.Children.Add($plusH)
        $plusV = [System.Windows.Shapes.Path]::new()
        $plusV.Stroke = $PaperBrush; $plusV.StrokeThickness = 2.5
        $plusV.StrokeStartLineCap = "Round"; $plusV.StrokeEndLineCap = "Round"
        $plusV.Data = [System.Windows.Media.StreamGeometry]::Parse("M12,6 L12,18")
        [void]$plusCanvas.Children.Add($plusV)
        $plusViewbox.Child = $plusCanvas
        $addFocusBtn.Child = $plusViewbox
        $addFocusBtn.Add_MouseLeftButtonUp({
            param($s, $e)
            $e.Handled = $true
            Show-FocusRecordDialog -Entry $Entry
        }.GetNewClosure())
        [void]$focusHeaderPanel.Children.Add($addFocusBtn)
    }
    [void]$bodyStack.Children.Add($focusHeaderPanel)

    # 1px dashed divider directly under the header row (index2.html), with ~16px
    # of air before the first record row.
    $focusDivider = New-SectionDivider
    $focusDivider.Margin = "0,6,0,16"
    [void]$bodyStack.Children.Add($focusDivider)

    $timelinePanel = [System.Windows.Controls.StackPanel]::new()
    $timelinePanel.Margin = "0,0,0,16"
    Add-FocusTimelineRows -Timeline $timelinePanel -Entry $Entry
    [void]$bodyStack.Children.Add($timelinePanel)

    # TOTAL FOCUS - heading matches Focus Record header style (20px bold)
    [void]$bodyStack.Children.Add((New-SectionDivider))
    $totalFocusHeader = [System.Windows.Controls.TextBlock]::new()
    $totalFocusHeader.Text = "Total Focus"
    $totalFocusHeader.FontFamily = "Consolas"
    $totalFocusHeader.FontSize = 20
    $totalFocusHeader.FontWeight = "Bold"
    $totalFocusHeader.Foreground = $InkBrush
    $totalFocusHeader.Margin = "0,14,0,0"
    [void]$bodyStack.Children.Add($totalFocusHeader)
    $totalFocusText = [System.Windows.Controls.TextBlock]::new()
    $totalFocusText.Text = "$ClockGlyph  $(Format-FocusTotal (Get-EntryFocusTotal $Entry))"
    $totalFocusText.FontFamily = "Consolas"
    $totalFocusText.FontSize = 14
    $totalFocusText.FontWeight = "Bold"
    $totalFocusText.Foreground = $InkBrush
    $totalFocusText.Margin = "0,4,0,0"
    [void]$bodyStack.Children.Add($totalFocusText)

    $global:FocusRegistry[$Entry.Id] = @{ Timeline = $timelinePanel; Total = $totalFocusText }

    $noteBox = [System.Windows.Controls.TextBox]::new()
    $noteBox.Text = $Entry.Note
    $noteBox.FontFamily = "Consolas"
    $noteBox.FontSize = 12
    $noteBox.AcceptsReturn = $true
    $noteBox.TextWrapping = "Wrap"
    # Auto-grow: the box expands in line-units as text is added (min 3 lines, up
    # to 40), so the panel grows with the note instead of the note scrolling.
    $noteBox.Height = [Double]::NaN
    $noteBox.MinLines = 3
    $noteBox.MaxLines = 40
    $noteBox.VerticalScrollBarVisibility = "Auto"
    $noteBox.Padding = "6,6,6,6"
    $noteBox.Margin = "0,8,0,0"
    $noteBox.BorderBrush = $InkBrush
    $noteBox.BorderThickness = 1
    $noteBox.Background = $WhiteBrush
    $noteBox.IsReadOnly = $global:ReadOnlyMode
    $noteBox.Add_LostFocus({
        if ($global:ReadOnlyMode) { return }
        # Update the in-memory note immediately (cheap); defer the disk writes ~1s
        # so losing focus never blocks the UI thread.
        $global:Entries = Set-EntryNote -Entries $global:Entries -Id $Entry.Id -Note $noteBox.Text
        $global:LiveEntries = $global:Entries
        $global:NoteDirty = $true
        $global:NoteDebounceTimer.Stop()
        $global:NoteDebounceTimer.Start()
    }.GetNewClosure())
    [void]$bodyStack.Children.Add($noteBox)

    [void]$region.Children.Add($bodyStack)
    $global:ExpandRegistry[$Entry.Id] = $region
    return $region
}

function global:New-DayRow {
    param($Entry)

    $outer = [System.Windows.Controls.Border]::new()
    $outer.BorderBrush     = $InkBrush
    $outer.BorderThickness = 1.5
    $outer.CornerRadius    = 6
    $outer.Background      = $RowBrush
    $outer.Margin          = "0,0,20,10"
    $outer.Cursor          = "Hand"
    # Lets Populate-RealizedRows identify (and de-register) this row when its
    # container is recycled to a different entry.
    $outer.Tag = $Entry.Id

    $stack = [System.Windows.Controls.StackPanel]::new()

    # ---- Header row (click anywhere to expand/collapse) ----
    # Two rows: row 0 is the compact summary line (date, done-count, total,
    # delete/select) right-anchored via a flexible spacer column; row 1 is the
    # goal chips, given the FULL row width on their own line. Previously chips
    # shared row 0 inside a "*" column squeezed between fixed columns, which at
    # normal window widths wrapped each chip to its own line and left a large
    # dead gap before the count/total text (the "no proper gaps" look).
    $headerGrid = [System.Windows.Controls.Grid]::new()
    $headerGrid.Margin = "14,10,14,10"
    # Widths set directly (faster than New-Object -Property in a hot per-row loop)
    $cdDate = [System.Windows.Controls.ColumnDefinition]::new(); $cdDate.Width = "100"
    $cdSpacer = [System.Windows.Controls.ColumnDefinition]::new(); $cdSpacer.Width = "*"
    $cdCount = [System.Windows.Controls.ColumnDefinition]::new(); $cdCount.Width = "Auto"
    $cdTotal = [System.Windows.Controls.ColumnDefinition]::new(); $cdTotal.Width = "Auto"
    $cdDel = [System.Windows.Controls.ColumnDefinition]::new(); $cdDel.Width = "Auto"
    [void]$headerGrid.ColumnDefinitions.Add($cdDate)
    [void]$headerGrid.ColumnDefinitions.Add($cdSpacer)
    [void]$headerGrid.ColumnDefinitions.Add($cdCount)
    [void]$headerGrid.ColumnDefinitions.Add($cdTotal)
    [void]$headerGrid.ColumnDefinitions.Add($cdDel)

    $rdSummary = [System.Windows.Controls.RowDefinition]::new(); $rdSummary.Height = "Auto"
    $rdChips = [System.Windows.Controls.RowDefinition]::new(); $rdChips.Height = "Auto"
    [void]$headerGrid.RowDefinitions.Add($rdSummary)
    [void]$headerGrid.RowDefinitions.Add($rdChips)

    # Date column
    $dateStack = [System.Windows.Controls.StackPanel]::new()
    $dateText = [System.Windows.Controls.TextBlock]::new()
    $dateText.Text = $Entry.Date
    $dateText.FontFamily = "Consolas"
    $dateText.FontWeight = "Bold"
    $dateText.FontSize = 13
    $dateText.Foreground = $InkBrush
    $dayText = [System.Windows.Controls.TextBlock]::new()
    $dayText.Text = "DAY $($Entry.DayNumber)"
    $dayText.FontFamily = "Consolas"
    $dayText.FontSize = 10
    $dayText.Foreground = $InkLightBrush
    $dayText.Margin = "0,2,0,0"
    $global:DayLabelRegistry[$Entry.Id] = $dayText   # for in-place renumber after delete
    [void]$dateStack.Children.Add($dateText)
    [void]$dateStack.Children.Add($dayText)
    [System.Windows.Controls.Grid]::SetColumn($dateStack, 0)
    [System.Windows.Controls.Grid]::SetRow($dateStack, 0)
    [void]$headerGrid.Children.Add($dateStack)

    # Chips panel - own full-width row beneath the summary line, so it always
    # has room to lay chips out horizontally instead of wrapping one-per-line.
    $chipsPanel = [System.Windows.Controls.WrapPanel]::new()
    $chipsPanel.Margin = "0,10,0,0"
    foreach ($goal in $Entry.Goals) {
        [void]$chipsPanel.Children.Add((New-GoalChip -Goal $goal -EntryId $Entry.Id))
    }
    [System.Windows.Controls.Grid]::SetRow($chipsPanel, 1)
    [System.Windows.Controls.Grid]::SetColumn($chipsPanel, 0)
    [System.Windows.Controls.Grid]::SetColumnSpan($chipsPanel, 5)
    [void]$headerGrid.Children.Add($chipsPanel)

    # Goal fraction (col 2)
    $doneCount = 0; $totalCount = 0
    foreach ($g in $Entry.Goals) { $totalCount++; if ($g.Status -eq "done") { $doneCount++ } }
    $countText = [System.Windows.Controls.TextBlock]::new()
    $countText.Text = "$doneCount/$totalCount"
    $countText.FontFamily = "Consolas"
    $countText.FontSize = 12
    $countText.Foreground = $InkLightBrush
    $countText.HorizontalAlignment = "Right"
    $countText.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($countText, 2)
    [System.Windows.Controls.Grid]::SetRow($countText, 0)
    [void]$headerGrid.Children.Add($countText)
    Register-Count -EntryId $Entry.Id -TextBlock $countText

    # Focus Total Hour (col 3), computed live from the day's logged sessions.
    $sumMin = 0
    foreach ($s in @($Entry.Sessions)) { if ($s) { $sumMin += [int]$s.duration } }
    $h = [math]::Floor($sumMin / 60); $m = $sumMin % 60
    $fmtTotal = if ($h -eq 0) { "$m min" } elseif ($m -eq 0) { "$h h" } else { "$h h $m min" }
    $totalText = [System.Windows.Controls.TextBlock]::new()
    $totalText.Text = "Total Hour : $fmtTotal"
    $totalText.FontFamily = "Consolas"
    $totalText.FontSize = 12
    $totalText.FontWeight = "Bold"
    $totalText.Foreground = $InkBrush
    $totalText.Margin = "16,0,12,0"
    $totalText.VerticalAlignment = "Center"
    $global:TotalRegistry[$Entry.Id] = $totalText
    [System.Windows.Controls.Grid]::SetColumn($totalText, 3)
    [System.Windows.Controls.Grid]::SetRow($totalText, 0)
    [void]$headerGrid.Children.Add($totalText)

    # Delete button / Select-mode checkbox — only for today-and-future dates.
    # In select mode the checkbox replaces the delete button (same column).
    $entryDate = [datetime]::ParseExact($Entry.Date, "yyyy-MM-dd", $null)
    $isEligible = ($entryDate -ge (Get-Date).Date)

    # ---- Select-mode checkbox (col 4) ----
    $selCheckBox = [System.Windows.Controls.CheckBox]::new()
    $selCheckBox.Margin = "4,0,4,0"
    $selCheckBox.Cursor = "Hand"
    $selCheckBox.ToolTip = "Select for bulk delete"
    $selCheckBox.IsChecked = $global:SelectedEntryIds.ContainsKey($Entry.Id)
    if ($isEligible -and $global:SelectModeActive) {
        $selCheckBox.Visibility = "Visible"
    } else {
        $selCheckBox.Visibility = "Collapsed"
    }
    $selCheckBox.Add_Checked({
        param($s, $e)
        $e.Handled = $true
        $global:SelectedEntryIds[$Entry.Id] = $true
        Update-BulkDeleteBanner
    }.GetNewClosure())
    $selCheckBox.Add_Unchecked({
        param($s, $e)
        $e.Handled = $true
        $global:SelectedEntryIds.Remove($Entry.Id)
        Update-BulkDeleteBanner
    }.GetNewClosure())
    [System.Windows.Controls.Grid]::SetColumn($selCheckBox, 4)
    [System.Windows.Controls.Grid]::SetRow($selCheckBox, 0)
    [void]$headerGrid.Children.Add($selCheckBox)

    # ---- Single-row delete button (col 4) ----
    $deleteBtn = [System.Windows.Controls.Button]::new()
    $deleteBtn.Content = $DeleteGlyph   # small x
    $deleteBtn.FontSize = 11
    $deleteBtn.Padding = "6,2,6,2"
    $deleteBtn.Margin = "0,0,4,0"
    $deleteBtn.Background = $TransparentBrush
    $deleteBtn.BorderBrush = $TransparentBrush
    $deleteBtn.Foreground = $ChipMissBd   # cached brush (avoid per-row creation)
    $deleteBtn.Cursor = "Hand"
    $deleteBtn.ToolTip = "Delete this day"
    if (-not $isEligible -or $global:SelectModeActive) {
        $deleteBtn.Visibility = "Collapsed"
    }
    $deleteBtn.Add_Click({
        param($s, $e)
        $e.Handled = $true   # stop expand/collapse from firing
        if ($global:ReadOnlyMode) { return }
        $confirm = [System.Windows.MessageBox]::Show(
            "Delete the entry for $($Entry.Date)? This cannot be undone.",
            "Delete Day",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        $global:Entries = Remove-Entry -Entries $global:Entries -Id $Entry.Id
        $global:LiveEntries = $global:Entries
        $global:Entries = Renumber-Entries -Entries $global:Entries
        $global:LiveEntries = $global:Entries
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync

        # Incremental: drop this row's registry entries and renumber the surviving
        # realized rows' "DAY N" labels in place. The virtualized ItemsControl
        # removes the container itself once the item leaves DayItems (Update-FilterView
        # below reconciles that) - no full list rebuild.
        $global:RowRegistry.Remove($Entry.Id)
        $global:StackRegistry.Remove($Entry.Id)
        $global:CountRegistry.Remove($Entry.Id)
        $global:DayLabelRegistry.Remove($Entry.Id)
        $global:ExpandRegistry.Remove($Entry.Id)
        $global:ExpandedIds.Remove($Entry.Id)
        $global:TotalRegistry.Remove($Entry.Id)
        $global:FocusRegistry.Remove($Entry.Id)
        @($global:ChipRegistry.Keys) | Where-Object { $_ -like "$($Entry.Id)|*" } | ForEach-Object {
            $global:ChipRegistry.Remove($_)
        }

        foreach ($e in $global:Entries) {
            if ($global:DayLabelRegistry.ContainsKey($e.Id)) {
                $global:DayLabelRegistry[$e.Id].Text = "DAY $($e.DayNumber)"
            }
        }

        Update-FilterView
    }.GetNewClosure())
    [System.Windows.Controls.Grid]::SetColumn($deleteBtn, 4)
    [System.Windows.Controls.Grid]::SetRow($deleteBtn, 0)
    [void]$headerGrid.Children.Add($deleteBtn)

    $headerGrid.Add_MouseLeftButtonUp({
        $wasExpanded = $global:ExpandedIds.ContainsKey($Entry.Id)

        # Only one row is expanded at a time: collapse whichever other row is shown.
        if (-not $wasExpanded) {
            $prev = @($global:ExpandedIds.Keys)[0]
            if ($prev -and $global:ExpandRegistry.ContainsKey($prev)) {
                $global:ExpandRegistry[$prev].Visibility = "Collapsed"
            }
        }

        $global:ExpandedIds.Clear()
        if ($wasExpanded) {
            if ($global:ExpandRegistry.ContainsKey($Entry.Id)) {
                $global:ExpandRegistry[$Entry.Id].Visibility = "Collapsed"
            }
        } else {
            $global:ExpandedIds[$Entry.Id] = $true
            if ($global:ExpandRegistry.ContainsKey($Entry.Id)) {
                $global:ExpandRegistry[$Entry.Id].Visibility = "Visible"
            } else {
                # First expand for this row: build its section lazily, then show it.
                $region = New-ExpandedSection -Entry $Entry
                [void]$global:StackRegistry[$Entry.Id].Children.Add($region)
                $region.Visibility = "Visible"
                # Safety net: ensure the Focus Record timeline is populated.
                # New-ExpandedSection calls Add-FocusTimelineRows, but in some
                # edge cases the children don't render (race with
                # Populate-RealizedRows or WPF layout). Force-refresh via the
                # same in-place update path used when a session completes.
                Update-FocusViewsInPlace -EntryId $Entry.Id
            }
        }
    }.GetNewClosure())

    [void]$stack.Children.Add($headerGrid)
    $global:StackRegistry[$Entry.Id] = $stack
    $global:RowRegistry[$Entry.Id] = $outer   # for show/hide/remove without rebuilding

    # ---- Expanded section (eagerly built only for the one row already expanded) ----
    if ($global:ExpandedIds.ContainsKey($Entry.Id)) {
        [void]$stack.Children.Add((New-ExpandedSection -Entry $Entry))
        # Same safety net as the lazy-build path above.
        Update-FocusViewsInPlace -EntryId $Entry.Id
    }

    $outer.Child = $stack
    return $outer
}

function global:Get-DaysPanel {
    # The VirtualizingStackPanel under the ItemsControl. Found lazily (the ItemsControl's
    # template is only instantiated once the window lays out) and cached.
    if ($global:DaysVSP) { return $global:DaysVSP }
    $stack = New-Object System.Collections.Stack
    $stack.Push($DaysList)
    while ($stack.Count -gt 0) {
        $d = $stack.Pop()
        if ($d -is [System.Windows.Controls.VirtualizingStackPanel]) {
            $global:DaysVSP = $d
            return $d
        }
        $n = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($d)
        for ($i = 0; $i -lt $n; $i++) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($d, $i)) }
    }
    return $null
}

function global:Remove-RowFromRegistries {
    # Drops every registry entry for a day row that is leaving the realized window
    # (its container was recycled to a different entry), so stale control references
    # don't accumulate or get double-registered.
    param([string]$EntryId)
    $global:RowRegistry.Remove($EntryId)
    $global:StackRegistry.Remove($EntryId)
    $global:CountRegistry.Remove($EntryId)
    $global:DayLabelRegistry.Remove($EntryId)
    $global:ExpandRegistry.Remove($EntryId)
    $global:TotalRegistry.Remove($EntryId)
    $global:FocusRegistry.Remove($EntryId)
    # ExpandedIds is data (not a control reference) - kept so a row that was expanded
    # re-expands when scrolled back into view.
    @($global:ChipRegistry.Keys) | Where-Object { $_ -like "$EntryId|*" } | ForEach-Object {
        $global:ChipRegistry.Remove($_)
    }
}

function global:Populate-RealizedRows {
    # Rebuilds the content of each realized day-row container to match its item.
    # Called whenever the item generator realizes/recycles containers (initial
    # layout, scrolling, add/delete, filter changes). Idempotent by default: a
    # container whose content already matches its item is left untouched, so
    # goal toggles / note typing never rebuild a row in place.
    #
    # -Force bypasses that idempotency check and rebuilds every realized row
    # regardless of whether its Tag already matches. This exists for the
    # startup safety net: the very first (premature) populate pass can run
    # while DaysList still has ~0 actual width (window not shown/sized yet),
    # which collapses the header Grid's Star column and the chip WrapPanel to
    # near-zero size - rows get built, just built wrong. Because that first
    # pass DID set curId = item.Id, the normal idempotent check then treats
    # those rows as "already correct" forever after, so a later plain
    # Populate-RealizedRows call (even with an UpdateLayout()) can't repair
    # them - only -Force can, by unconditionally re-running New-DayRow once
    # the panel has its real width.
    param([switch]$Force)
    $panel = Get-DaysPanel
    if (-not $panel) { return }
    # Only materialize items that are actually in the current ItemsSource - a
    # recycled container can briefly hold an item that a filter/delete just removed.
    $live = @{}
    foreach ($e in $global:DayItems) { $live[$e.Id] = $true }
    $now = @{}
    foreach ($cont in @($panel.Children)) {
        $item = $cont.DataContext
        if ($null -eq $item -or $item -isnot [pscustomobject] -or -not $item.PSObject.Properties['Id']) { continue }
        $now[$item.Id] = $true
        if (-not $live.ContainsKey($item.Id)) { continue }
        $cur = $cont.Content
        $curId = if ($null -ne $cur -and $null -ne $cur.Tag) { [string]$cur.Tag } else { $null }
        if ($Force -or $curId -ne $item.Id) {
            if ($curId -and $global:RowRegistry.ContainsKey($curId)) {
                Remove-RowFromRegistries -EntryId $curId
            }
            $cont.Content = New-DayRow -Entry $item
        }
    }
    # Prune registries for any row that was realized last pass but is no longer -
    # covers recycled containers whose content was reset by the ListBox before we
    # saw it, so stale control references never accumulate.
    if ($null -ne $global:RealizedIds) {
        foreach ($oldId in @($global:RealizedIds.Keys)) {
            if (-not $now.ContainsKey($oldId) -and $global:RowRegistry.ContainsKey($oldId)) {
                Remove-RowFromRegistries -EntryId $oldId
            }
        }
    }
    $global:RealizedIds = $now
}

function global:Render-Days {
    if ($global:Settings.JourneyName) {
        $JourneyNameText.Text = "Journey: $($global:Settings.JourneyName)"
        $JourneyNameText.Visibility = "Visible"
    } else {
        $JourneyNameText.Visibility = "Collapsed"
    }

    # Rebuild the virtualized ItemsSource (DayItems) from the filtered, newest-first
    # entry set. Only the rows near the scroll position are materialized afterwards
    # (Populate-RealizedRows); everything else is cheap data until scrolled into view.
    $global:ChipRegistry   = @{}
    $global:CountRegistry  = @{}
    $global:ExpandRegistry = @{}
    $global:StackRegistry  = @{}
    $global:RowRegistry    = @{}
    $global:DayLabelRegistry = @{}
    $global:TotalRegistry  = @{}
    $global:FocusRegistry  = @{}

    $desired = Get-FilteredSortedEntries
    $desiredSet = @{}
    foreach ($e in $desired) { $desiredSet[$e.Id] = $true }
    foreach ($k in @($global:ExpandedIds.Keys)) {
        if (-not $desiredSet.ContainsKey($k)) { $global:ExpandedIds.Remove($k) }
    }

    $global:DayItems.Clear()
    foreach ($entry in $desired) { $global:DayItems.Add($entry) }

    Populate-RealizedRows
    Update-EmptyState
}

# ---------- Bulk-select delete mode ----------

function global:Update-BulkDeleteBanner {
    # Single source of truth for the banner visibility and delete button state.
    $count = $global:SelectedEntryIds.Count
    if ($global:SelectModeActive -and $count -gt 0) {
        $BulkDeleteBanner.Visibility = "Visible"
        $BulkDeleteCountText.Text = "$count day$(if ($count -ne 1) { 's' }) selected"
        $BulkDeleteConfirmBtn.Content = "DELETE SELECTED ($count)"
    } else {
        $BulkDeleteBanner.Visibility = "Collapsed"
    }
}

function global:Enter-SelectMode {
    if ($global:ReadOnlyMode) { return }
    $global:SelectModeActive = $true
    $global:SelectedEntryIds = @{}
    $BulkDeleteToggleBtn.Content = "EXIT SELECT"
    $BulkDeleteToggleBtn.ToolTip = "Exit select mode"
    $BulkDeleteBanner.Visibility = "Collapsed"
    Render-Days   # rebuilds rows so checkboxes appear instead of delete buttons
}

function global:Exit-SelectMode {
    $global:SelectModeActive = $false
    $global:SelectedEntryIds = @{}
    $BulkDeleteToggleBtn.Content = "SELECT"
    $BulkDeleteToggleBtn.ToolTip = "Select days for bulk delete"
    $BulkDeleteBanner.Visibility = "Collapsed"
    Render-Days   # rebuilds rows so delete buttons appear instead of checkboxes
}

function global:Reset-SelectMode {
    # Lightweight reset without triggering a full rebuild. Used by journey/archive
    # transitions that rebuild the view themselves (so we don't double-rebuild).
    $global:SelectModeActive = $false
    $global:SelectedEntryIds = @{}
    $BulkDeleteToggleBtn.Content = "SELECT"
    $BulkDeleteToggleBtn.ToolTip = "Select days for bulk delete"
    $BulkDeleteBanner.Visibility = "Collapsed"
}

function global:Invoke-BulkDelete {
    if ($global:SelectedEntryIds.Count -eq 0) { return }
    $count = $global:SelectedEntryIds.Count
    $confirm = [System.Windows.MessageBox]::Show(
        "Delete $count selected day$(if ($count -ne 1) { 's' })? This cannot be undone.",
        "Bulk Delete",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    # Single pass: remove all selected entries from $global:Entries.
    # One renumber, one save, one sync — not N of each.
    $global:Entries = @($global:Entries | Where-Object { -not $global:SelectedEntryIds.ContainsKey($_.Id) })
    $global:LiveEntries = $global:Entries
    $global:Entries = Renumber-Entries -Entries $global:Entries
    $global:LiveEntries = $global:Entries
    Save-Entries -Entries $global:Entries
    Invoke-AutoSync

    Exit-SelectMode   # clears selection, hides banner, rebuilds view
}

# ---------- Content area routing ----------
function global:Show-ContentArea {
    param([ValidateSet("Days","Info","Focus")][string]$Mode)

    # Leaving the Focus view hides the on-screen clock but does NOT pause the
    # session. The DispatcherTimer keeps ticking while the user browses Days/
    # Report/etc., so a running session keeps accruing elapsed time (and keeps
    # its periodic focus-session.json write) - only an explicit user action
    # (PAUSE / STOP / DISCARD) or a journey/archive transition that swaps the
    # dataset force-stops it (see Stop-FocusSessionAtTransition).
    if ($Mode -ne "Focus") {
        # A maximized Focus view always restores the header/sidebar when the user
        # navigates away, so the sidebar isn't left hidden elsewhere in the app.
        if ($global:FocusMaximized) { Restore-FocusLayout }
        $FocusModePanel.Visibility = "Collapsed"
    } else {
        $DaysScrollViewer.Visibility = "Collapsed"
        $InfoAreaScrollViewer.Visibility = "Collapsed"
        $FocusModePanel.Visibility = "Visible"
        return
    }

    if ($Mode -eq "Days") {
        # Navigating back to the Days view dismisses any switch-confirm, so make
        # sure ADD DAY / SYNC are never left disabled. But keep them disabled
        # while the Create/Edit journey panel is open (StartJourneyPanel visible).
        if ($StartJourneyPanel.Visibility -ne "Visible") {
            Set-JourneyEditButtons -Disabled $false
        }
        # Only show DaysScrollViewer if the journey panel is not visible
        if ($StartJourneyPanel.Visibility -eq "Visible") {
            $DaysScrollViewer.Visibility = "Collapsed"
        } else {
            $DaysScrollViewer.Visibility = "Visible"
        }
        $InfoAreaScrollViewer.Visibility = "Collapsed"
    } else {
        $DaysScrollViewer.Visibility = "Collapsed"
        $InfoAreaScrollViewer.Visibility = "Visible"
    }
}

# ---------- Sidebar: collapse / expand to icon rail ----------
function global:Set-SidebarNarrow {
    param([bool]$Narrow)
    $global:SidebarNarrow = $Narrow
    foreach ($lbl in @($CollapseLabel, $DashboardLabel, $ExportLabel, $SwitchLabel, $BackupLabel, $ArchiveLabel, $FocusSidebarLabel)) {
        $lbl.Visibility = if ($Narrow) { "Collapsed" } else { "Visible" }
    }
    # Icons should always stay visible
    foreach ($icon in @($CollapseIcon, $DashboardIcon, $ExportIcon, $SwitchIcon, $BackupIcon, $ArchiveIcon, $FocusIcon)) {
        $icon.Visibility = "Visible"
    }
    if ($Narrow) {
        $DashboardSubs.Visibility = "Collapsed"
        $ExportSubs.Visibility = "Collapsed"
        $SwitchJourneySubs.Visibility = "Collapsed"
        $BackupSubs.Visibility = "Collapsed"
    } else {
        # Always show Dashboard sub-items when expanded
        $DashboardSubs.Visibility = "Visible"
        Set-SidebarActive $global:ActiveSidebarSection
    }
}

$Sidebar.Add_SizeChanged({
    param($s, $e)
    $isNarrow = $Sidebar.ActualWidth -le 80
    if ($isNarrow -ne $global:SidebarNarrow) { Set-SidebarNarrow $isNarrow }
}.GetNewClosure())

$CollapseSidebarBtn.Add_MouseLeftButtonUp({
    if ($global:SidebarNarrow) {
        $SidebarColumn.Width = [System.Windows.GridLength]::new(190)
    } else {
        $SidebarColumn.Width = [System.Windows.GridLength]::new(52)
    }
})

# ---------- Sidebar: accordion ----------
function global:Set-SidebarActive {
    param([string]$Section)
    $global:ActiveSidebarSection = $Section

    foreach ($hdr in @($DashboardHeaderBtn, $ExportHeaderBtn, $SwitchJourneyBtn, $BackupHeaderBtn, $LoadArchiveBtn, $FocusModeBtn)) {
        $hdr.Background = $TransparentBrush
    }

    # Only hide Export, Switch and Backup subs - Dashboard is always visible
    $ExportSubs.Visibility        = "Collapsed"
    $SwitchJourneySubs.Visibility = "Collapsed"
    $BackupSubs.Visibility        = "Collapsed"

    if ($global:SidebarNarrow) {
        $DashboardSubs.Visibility = "Collapsed"
        return
    }

    # Dashboard sub-items are always visible when sidebar is not narrow
    $DashboardSubs.Visibility = "Visible"

    switch ($Section) {
        "dash"   { $DashboardHeaderBtn.Background = $HighlightBrush }
        "export" { $ExportHeaderBtn.Background     = $HighlightBrush; $ExportSubs.Visibility = "Visible" }
        "switch" { $SwitchJourneyBtn.Background    = $HighlightBrush; $SwitchJourneySubs.Visibility = "Visible" }
        "backup" { $BackupHeaderBtn.Background     = $HighlightBrush; $BackupSubs.Visibility = "Visible" }
        "archive"{ $LoadArchiveBtn.Background       = $HighlightBrush }
        "focus"  { $FocusModeBtn.Background         = $HighlightBrush }
    }
}

function global:Toggle-SidebarSection {
    param([string]$Section)
    if ($global:ActiveSidebarSection -eq $Section) {
        # collapse it back (accordion: clicking the open section again closes it)
        Set-SidebarActive ""
    } else {
        Set-SidebarActive $Section
    }
}

function global:Set-DashboardActiveButton {
    # Highlights one Dashboard sub-button (a filter or Report) in the active "blue"
    # style, and resets the others - so Report highlights exactly like the filters.
    param([string]$Name)   # "All" | "Completed" | "Missed" | "Report"
    foreach ($btn in @($FilterAllBtn, $FilterCompletedBtn, $FilterMissedBtn, $ReportBtn)) {
        $btn.Background = $TransparentBrush
        $btn.Foreground = $InkLightBrush
    }
    $active = switch ($Name) {
        "All"       { $FilterAllBtn }
        "Completed" { $FilterCompletedBtn }
        "Missed"    { $FilterMissedBtn }
        "Report"    { $ReportBtn }
    }
    $active.Background = $InkBrush
    $active.Foreground = $PaperBrush
}

function global:Set-Filter {
    param([string]$Filter)
    $global:Filter = $Filter
    Set-DashboardActiveButton $Filter
    Update-ButtonVisibility
    Show-ContentArea "Days"
    Set-SidebarActive "dash"
    # A full build is only needed when the virtualized ItemsSource is empty but
    # entries exist; otherwise reconcile DayItems in place (preserves scroll).
    if ($global:DayItems.Count -eq 0 -and @($global:Entries).Count -gt 0) {
        Render-Days
    } else {
        Update-FilterView
    }
}

# ---------- Goal template editor (used only during journey setup) ----------
function global:Render-GoalEditChips {
    $GoalEditPanel.Children.Clear()
    foreach ($hours in ($global:PendingGoalHours | Sort-Object)) {
        $chip = [System.Windows.Controls.Border]::new()
        $chip.CornerRadius = 12
        $chip.Padding = "8,3,6,3"
        $chip.Margin = "0,0,6,6"
        $chip.BorderThickness = 1
        $chip.BorderBrush = $InkBrush
        $chip.Background = $RowBrush

        $inner = [System.Windows.Controls.StackPanel]::new()
        $inner.Orientation = "Horizontal"

        $label = [System.Windows.Controls.TextBlock]::new()
        $label.Text = "$hours h"
        $label.FontFamily = "Consolas"
        $label.FontSize = 12
        $label.Foreground = $InkBrush
        $label.VerticalAlignment = "Center"
        [void]$inner.Children.Add($label)

        $removeText = [System.Windows.Controls.TextBlock]::new()
        $removeText.Text = "  x"
        $removeText.FontFamily = "Consolas"
        $removeText.FontSize = 12
        $removeText.Foreground = $RedBrush
        $removeText.Cursor = "Hand"
        $removeText.Margin = "4,0,0,0"
        $removeText.Add_MouseLeftButtonUp({
            $global:PendingGoalHours = @($global:PendingGoalHours | Where-Object { $_ -ne $hours })
            Render-GoalEditChips
        }.GetNewClosure())
        [void]$inner.Children.Add($removeText)

        $chip.Child = $inner
        [void]$GoalEditPanel.Children.Add($chip)
    }
}

function global:Add-PendingGoalFromInput {
    $val = 0
    if ([int]::TryParse($NewGoalHoursBox.Text, [ref]$val) -and $val -gt 0) {
        if ($global:PendingGoalHours -notcontains $val) {
            $global:PendingGoalHours = @($global:PendingGoalHours) + $val
            Render-GoalEditChips
        }
        $NewGoalHoursBox.Text = ""
        $NewGoalHoursBox.Focus() | Out-Null
    } else {
        [System.Windows.MessageBox]::Show("Enter a whole number of hours greater than 0.", "Productivity Tracker") | Out-Null
    }
}

function global:Get-ComputedDate {
    # Was: StartDate + EntryCount days. That assumed one entry gets added
    # per real calendar day with zero gaps, so any skipped day (app not
    # opened) silently shifted every day added afterward earlier than its
    # real date, forever. Real "today" everywhere - matches Focus Mode,
    # Show-Report, and Backfill-SkippedDays, which all already used the
    # real clock. Kept as a thin wrapper (no remaining internal callers)
    # in case anything external still expects this name.
    return (Get-Date).Date
}

function global:Update-SyncStatus {
    param([datetime]$Timestamp)
    $daysAgo = ((Get-Date).Date - $Timestamp.Date).Days
    $timeStr = $Timestamp.ToString('h:mm tt')
    if ($daysAgo -eq 0) {
        $SyncStatusText.Foreground = $InkLightBrush
        $SyncStatusText.Text = "Last synced: $timeStr"
    } else {
        $dateStr = $Timestamp.ToString('MMM d, yyyy')
        $unit = if ($daysAgo -eq 1) { "day" } else { "days" }
        $SyncStatusText.Foreground = $AmberBrush
        $SyncStatusText.Text = "Last synced: $dateStr, $timeStr ($daysAgo $unit ago)"
    }
}

function global:Get-SafeFileLabel {
    param([string]$Label, [string]$Fallback = "untitled")
    $safe = ($Label -replace '[^a-zA-Z0-9\-]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($safe)) { return $Fallback }
    return $safe
}

function global:Get-LiveJourneyDir {
    $liveLabel = if ($global:LiveSettings.JourneyName) { $global:LiveSettings.JourneyName } else { "current" }
    $liveSafe  = Get-SafeFileLabel -Label $liveLabel -Fallback "current"
    return Join-Path (Join-Path $global:BackupDir "Journeys") $liveSafe
}

function global:Set-SyncPending {
    param([bool]$Pending)
    $global:SyncPending = $Pending
    if ($Pending) {
        $SyncBtn.Background = $AmberBrush
        $SyncBtn.Foreground = $PaperBrush
        $SyncBtn.FontWeight = "Bold"
        $SyncBtn.ToolTip = "Unsynced changes - click to back up"
    } else {
        $SyncBtn.Background = $TransparentBrush
        $SyncBtn.Foreground = $InkBrush
        $SyncBtn.FontWeight = "Normal"
        $SyncBtn.ToolTip = $null
    }
}

function global:Write-LiveJson {
    # Durability half of the safety net: writes the live JSON immediately on
    # every edit. Cheap, so it always runs - the Excel copy is what gets debounced.
    try {
        $liveDir = Get-LiveJourneyDir
        if (-not (Test-Path $liveDir)) {
            New-Item -ItemType Directory -Path $liveDir -Force | Out-Null
        }
        $global:LiveEntries | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $liveDir "live.json") -Encoding UTF8
        $global:LiveSettings | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $liveDir "live-settings.json") -Encoding UTF8
        Set-SyncPending $true
    } catch {
        $SyncStatusText.Foreground = $RedBrush
        $SyncStatusText.Text = "Auto-sync failed - click SYNC to retry"
    }
}

function global:Write-LiveExcel {
    # Expensive half of the safety net: the live.xlsx copy + sync timestamp.
    # Only runs after the debounce window elapses (or on an explicit flush), so
    # rapid edits collapse into a single write instead of one per click.
    try {
        $liveDir = Get-LiveJourneyDir
        Export-EntriesToXlsx -Entries $global:LiveEntries -Path (Join-Path $liveDir "live.xlsx")
        Update-SyncStatus -Timestamp (Get-Date)
    } catch {
        $SyncStatusText.Foreground = $RedBrush
        $SyncStatusText.Text = "Auto-sync failed - click SYNC to retry"
    }
}

function global:Flush-AutoSync {
    # Forces the debounced Excel write to run now. Used by manual SYNC so its
    # status check and failure detection stay synchronous.
    $global:AutoSyncDebounceTimer.Stop()
    Write-LiveExcel
}

function global:Flush-NoteSave {
    # Forces any pending debounced note save to run now (on window close) so a
    # note typed just before closing is still persisted.
    $global:NoteDebounceTimer.Stop()
    if ($global:NoteDirty) {
        $global:NoteDirty = $false
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync
    }
}

function global:Flush-GoalSave {
    # Forces any pending debounced goal-chip save to run now (on window close)
    # so toggling a chip just before closing is still persisted.
    $global:GoalDebounceTimer.Stop()
    if ($global:GoalDirty) {
        $global:GoalDirty = $false
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync
    }
}

function global:Invoke-AutoSync {
    <#
        Instant safety net - fires on every edit. Overwrites a single "live"
        slot inside the live journey's own Backup\Journeys\<name>\ folder.
        This is NOT part of the ranked version history - only Sync-AllJourneysToBackup
        (the SYNC button) adds a new rotated checkpoint.

        JSON is written immediately for durability; the live.xlsx export is
        debounced ~2s so bursts of edits batch into one write.
    #>
    Write-LiveJson
    $global:AutoSyncDebounceTimer.Stop()
    $global:AutoSyncDebounceTimer.Start()
}

function global:Write-JourneySnapshots {
    <#
        Writes exactly the given snapshots into $Dir, ranked and labeled
        newest-first (e.g. "01 - newest (2026-07-31 11-54 PM).json"). Wipes
        whatever files are currently in $Dir first - callers must have
        already read back anything from $Dir they want to keep, since this
        is the only place that decides final file names/ranks.
    #>
    param([string]$Dir, [array]$Snapshots)
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
    Get-ChildItem -Path $Dir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match '^\d{2} - ' } |
        Remove-Item -Force

    $ordered = @($Snapshots | Sort-Object Time -Descending)
    for ($i = 0; $i -lt $ordered.Count; $i++) {
        $rank = $i + 1
        $tag = if ($ordered.Count -eq 1) { "latest" }
               elseif ($rank -eq 1) { "newest" }
               elseif ($rank -eq $ordered.Count) { "oldest" }
               else { "" }
        $readableDate = $ordered[$i].Time.ToString("yyyy-MM-dd hh-mm tt")
        $label = if ($tag) { "{0:00} - {1} ({2})" -f $rank, $tag, $readableDate } else { "{0:00} - ({1})" -f $rank, $readableDate }

        $ordered[$i].Entries | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $Dir "$label.json") -Encoding UTF8
        if (@($ordered[$i].Entries).Count -gt 0) {
            Export-EntriesToXlsx -Entries $ordered[$i].Entries -Path (Join-Path $Dir "$label.xlsx")
        }
    }
}

function global:Sync-AllJourneysToBackup {
    <#
        Backup\Journeys\<journey name>\ - one folder per journey, holding up
        to $VersionsPerJourney ranked snapshots (newest/oldest labeled).

        The LIVE journey's folder is a rolling version history across syncs:
        each call reads back whatever snapshots are already sitting in its
        folder from previous syncs, adds this sync's current state, and keeps
        only the newest N - so the history survives from one sync to the
        next instead of being wiped and rebuilt from nothing every time.

        Archived journeys are regenerated from Data\Archive each call, since
        that folder is itself the durable source of truth for those.

        If a past archive shares the live journey's name, both are merged
        into one combined, re-ranked history in the same folder.
    #>
    param([int]$VersionsPerJourney = 3)

    $journeysDir = Join-Path $global:BackupDir "Journeys"
    if (-not (Test-Path $journeysDir)) {
        New-Item -ItemType Directory -Path $journeysDir | Out-Null
    }

    $liveDir = Get-LiveJourneyDir

    # Read back whatever ranked snapshots the live journey's folder already
    # holds from past syncs - never the "live.*" safety-net files, those
    # aren't part of the rotation.
    $liveSnapshots = @()
    if (Test-Path $liveDir) {
        Get-ChildItem -Path $liveDir -Filter "*.json" -ErrorAction SilentlyContinue |
            Where-Object { $_.BaseName -match '^\d{2} - ' } |
            ForEach-Object {
                try {
                    $raw = Get-Content -Path $_.FullName -Raw
                    if (-not [string]::IsNullOrWhiteSpace($raw)) {
                        $data = $raw | ConvertFrom-Json
                        $entries = if ($null -eq $data) { @() } elseif ($data -isnot [System.Array]) { @($data) } else { $data }
                        $liveSnapshots += [PSCustomObject]@{ Entries = $entries; Time = $_.LastWriteTime }
                    }
                } catch { }
            }
    }
    # This sync's snapshot of the live journey
    $liveSnapshots += [PSCustomObject]@{ Entries = $global:LiveEntries; Time = (Get-Date) }

    $groups = (Get-ArchivedJourneys) | Group-Object { if ($_.Name -eq "(untitled journey)") { $_.FolderName } else { $_.Name } }
    $handledDirs = @{}

    foreach ($group in $groups) {
        $toKeep = @($group.Group | Select-Object -First $VersionsPerJourney)
        $displayName = $toKeep[0].Name
        $safeName = if ($displayName -eq "(untitled journey)") {
            Get-SafeFileLabel -Label $toKeep[0].FolderName -Fallback "untitled"
        } else {
            Get-SafeFileLabel -Label $displayName
        }
        $journeyDir = Join-Path $journeysDir $safeName
        $handledDirs[$journeyDir] = $true

        $archivedSnaps = $toKeep | ForEach-Object {
            $restored = Restore-JourneyFromArchive -ArchivePath $_.Path
            [PSCustomObject]@{ Entries = $restored.Entries; Time = $_.ArchivedAt }
        }

        if ($journeyDir -eq $liveDir) {
            # Shares a name with the live journey - merge sync history + archived history together
            $combined = @($liveSnapshots) + @($archivedSnaps)
            $top = @($combined | Sort-Object Time -Descending | Select-Object -First $VersionsPerJourney)
            Write-JourneySnapshots -Dir $journeyDir -Snapshots $top
        } else {
            $top = @($archivedSnaps | Sort-Object Time -Descending | Select-Object -First $VersionsPerJourney)
            Write-JourneySnapshots -Dir $journeyDir -Snapshots $top
        }
    }

    if (-not $handledDirs.ContainsKey($liveDir)) {
        $top = @($liveSnapshots | Sort-Object Time -Descending | Select-Object -First $VersionsPerJourney)
        Write-JourneySnapshots -Dir $liveDir -Snapshots $top
    }

    return $journeysDir
}

# ---------- Inline info-area builders (replace the old modal pickers) ----------
function global:New-InfoButton {
    param([string]$Text, [string]$Tag)
    $b = [System.Windows.Controls.Button]::new()
    $b.Content = $Text
    $b.HorizontalAlignment = "Left"
    $b.Padding = "10,6,10,6"
    $b.Margin = "0,0,0,8"
    $b.FontFamily = "Consolas"
    return $b
}

function global:New-InfoHeading {
    param([string]$Text)
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $Text
    $t.FontFamily = "Consolas"
    $t.FontSize = 14
    $t.FontWeight = "Bold"
    $t.Foreground = $InkBrush
    $t.Margin = "0,0,0,12"
    return $t
}

function global:New-InfoNote {
    param([string]$Text)
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $Text
    $t.FontFamily = "Consolas"
    $t.FontSize = 11
    $t.Foreground = $InkLightBrush
    $t.TextWrapping = "Wrap"
    $t.Margin = "0,0,0,12"
    return $t
}

function global:New-RuleLine {
    # A thin dotted divider used to separate sections in the Info area.
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = "------------------------------------------------"
    $t.FontFamily = "Consolas"
    $t.FontSize = 11
    $t.Foreground = $InkLightBrush
    $t.Margin = "0,0,0,8"
    return $t
}

function global:Refresh-SwitchJourneyList {
    $SwitchJourneyListPanel.Children.Clear()
    $journeys = Get-LatestArchivedJourneys
    if ($journeys.Count -eq 0) {
        $none = [System.Windows.Controls.TextBlock]::new()
        $none.Text = "No other journeys yet"
        $none.FontFamily = "Consolas"
        $none.FontSize = 11
        $none.Foreground = $InkLightBrush
        $none.Margin = "6,4,0,0"
        [void]$SwitchJourneyListPanel.Children.Add($none)
        return
    }
    foreach ($j in $journeys) {
        $btn = [System.Windows.Controls.Button]::new()
        $btn.Content = $j.Name
        $btn.HorizontalContentAlignment = "Left"
        $btn.Background = $TransparentBrush
        $btn.BorderThickness = 0
        $btn.Padding = "6,4,6,4"
        $btn.FontSize = 12
        $btn.FontFamily = "Consolas"
        $btn.Add_Click({ Show-SwitchConfirm -Journey $j }.GetNewClosure())
        [void]$SwitchJourneyListPanel.Children.Add($btn)
    }
}

function global:Update-ButtonVisibility {
    <#
        Refreshes AddDayBtn / EditTemplateBtn / BulkDeleteToggleBtn visibility
        to match the current journey state. Call after any switch or journey
        setup so buttons stay in sync with what data actually exists.
    #>
    if ($global:Settings.StartDate -and -not $global:ReadOnlyMode) {
        $AddDayBtn.Visibility          = "Visible"
        $EditTemplateBtn.Visibility    = "Visible"
        $BulkDeleteToggleBtn.Visibility = "Visible"
    } else {
        $AddDayBtn.Visibility          = "Collapsed"
        $EditTemplateBtn.Visibility    = "Collapsed"
        $BulkDeleteToggleBtn.Visibility = "Collapsed"
    }
    # Journey/archive transitions always reset select mode — selection is
    # meaningless across a different dataset.
    if (-not $global:SelectModeActive) { return }
    Reset-SelectMode
}

function global:Set-JourneyEditButtons {
    param([bool]$Disabled)
    # Disable/Enable ADD DAY, SYNC, and SELECT while a journey operation is in
    # progress (Create/Edit journey panel or a Switch confirm). Buttons are
    # referenced via $global: so this works from inside GetNewClosure handlers too.
    $global:SyncBtn.IsEnabled   = -not $Disabled
    $global:SyncBtn.Opacity     = if ($Disabled) { 0.4 } else { 1.0 }
    $global:AddDayBtn.IsEnabled = -not $Disabled
    $global:AddDayBtn.Opacity   = if ($Disabled) { 0.4 } else { 1.0 }
    $BulkDeleteToggleBtn.IsEnabled = -not $Disabled
    $BulkDeleteToggleBtn.Opacity   = if ($Disabled) { 0.4 } else { 1.0 }
}

function global:Show-SwitchConfirm {
    param($Journey)

    # If the selected journey is the one already active, do nothing: don't
    # archive, restore, refresh, or change views - just close the dialog.
    if ($Journey.Name -eq $global:Settings.JourneyName) {
        Show-ContentArea "Days"
        Set-JourneyEditButtons -Disabled $false
        return
    }

    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Switch to `"$($Journey.Name)`"?"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "This archives your current journey. Nothing is deleted."))

    # Disable ADD DAY and SYNC while the confirm is showing so the user can't
    # trigger either mid-switch
    Set-JourneyEditButtons -Disabled $true

    $confirmBtn = [System.Windows.Controls.Button]::new()
    $confirmBtn.Content = "CONFIRM SWITCH"
    $confirmBtn.Width = 160
    $confirmBtn.HorizontalAlignment = "Left"
    $confirmBtn.Margin = "0,0,0,8"
    $confirmBtn.Add_Click({
        # Swapping to another journey replaces $global:Entries - a running focus
        # session would keep ticking against the old dataset. Force-stop it first;
        # on cancel, back out of the switch (buttons were disabled above).
        if (-not (Stop-FocusSessionAtTransition)) {
            Set-JourneyEditButtons -Disabled $false
            return
        }
        try {
            # Only archive the old journey if one actually exists (has a start date).
            # Archiving a blank journey would create a junk "(untitled journey)" entry
            # that never gets deduped and clutters the Load Archive / Switch lists.
            if ($global:Settings.StartDate) {
                Save-JourneyArchive -Entries $global:Entries -Settings $global:Settings | Out-Null
            }
            $restored = Restore-JourneyFromArchive -ArchivePath $Journey.Path
            $global:Entries  = $restored.Entries
            $global:Settings = $restored.Settings
            $global:LiveEntries  = $global:Entries
            $global:LiveSettings = $global:Settings
            $global:ExpandedIds = @{}

            Save-Entries -Entries $global:Entries
            Save-Settings -Settings $global:Settings
            Invoke-AutoSync
            Update-ButtonVisibility
            Set-Filter "All"
        } catch {
            [System.Windows.MessageBox]::Show("Switch failed: $($_.Exception.Message)", "Productivity Tracker") | Out-Null
        } finally {
            # Re-enable ADD DAY and SYNC regardless of success or failure
            Set-JourneyEditButtons -Disabled $false
        }
    }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($confirmBtn)

    Show-ContentArea "Info"
}

function global:Show-ExportResult {
    param([string]$Path)
    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Exported"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote $Path))
    Show-ContentArea "Info"
    Set-SidebarActive "export"
}

# ---------- Backup & Import explanation dialogs ----------

function global:Set-BackupFolderFlow {
    # The actual "pick a folder + persist + sync" work. Shown only after the user
    # confirms the explanation dialog (Feature 4).
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Choose the folder where backups will be stored"
    if (Test-Path $global:BackupDir) { $dlg.SelectedPath = $global:BackupDir }
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Show-ContentArea "Days"
        return
    }
    $path = $dlg.SelectedPath

    $global:Settings.BackupFolder    = $path
    $global:LiveSettings.BackupFolder = $path
    Save-Settings -Settings $global:Settings
    $global:BackupDir = $path

    try {
        Sync-AllJourneysToBackup | Out-Null
        [System.Windows.MessageBox]::Show(
            "Backup folder set to:`n$path`n`nCurrent journeys synced there.",
            "Productivity Tracker"
        ) | Out-Null
    } catch {
        [System.Windows.MessageBox]::Show(
            "Backup folder set to:`n$path`n`nBut syncing failed: $($_.Exception.Message)",
            "Productivity Tracker"
        ) | Out-Null
    }
    Show-ContentArea "Days"
}

function global:Show-BackupExplanationDialog {
    # Explains what backups are before asking where to put them (Feature 4).
    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Backup & Restore"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "This chooses where a safe copy of all your journeys is stored."))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "• Live Data — your current journey's days + settings, stored inside the app folder."))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "• Backup Folder — a copy of every journey, stored wherever you pick (OneDrive, D:\, a USB drive)."))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "• SYNC — copies every journey into the backup folder."))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "• IMPORT — restores journeys from a backup folder back into the app."))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "• Deleting the application folder does NOT delete your backups."))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "• Backups can be re-imported any time."))

    $contBtn = [System.Windows.Controls.Button]::new()
    $contBtn.Content = "CONTINUE"
    $contBtn.Width = 120
    $contBtn.HorizontalAlignment = "Left"
    $contBtn.Margin = "0,4,0,8"
    $contBtn.Add_Click({ Set-BackupFolderFlow }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($contBtn)

    $cancelBtn = [System.Windows.Controls.Button]::new()
    $cancelBtn.Content = "CANCEL"
    $cancelBtn.Width = 90
    $cancelBtn.HorizontalAlignment = "Left"
    $cancelBtn.Margin = "0,0,0,8"
    $cancelBtn.Add_Click({ Show-ContentArea "Days" }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($cancelBtn)

    Show-ContentArea "Info"
    Set-SidebarActive "backup"
}

function global:Start-ImportFlow {
    # Actual "pick folder + import" work, invoked after the info dialog (Feature 3).
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Choose the backup folder to restore from (the one containing a Journeys subfolder)"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Show-ContentArea "Days"
        return
    }
    $srcRoot = $dlg.SelectedPath

    try {
        $result = Import-JourneysFromFolder -SourceRoot $srcRoot
        $restored = @($result.Restored)

        if ($restored.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No journeys were found in:`n$srcRoot",
                "Import",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
            return
        }

        # If the current app has no journey yet, activate the backup's live journey
        # so the user sees restored data immediately (never overwrites existing data).
        if ($result.Live -and -not $global:Settings.StartDate) {
            $global:Entries      = $result.Live.Entries
            $global:Settings     = $result.Live.Settings
            $global:LiveEntries  = $global:Entries
            $global:LiveSettings = $global:Settings
            $global:ExpandedIds  = @{}
            Save-Entries -Entries $global:Entries
            Save-Settings -Settings $global:Settings
            Invoke-AutoSync
            Update-ButtonVisibility
        }

        # Feature 3.1: prompt user to rename each imported journey before saving.
        $global:ImportRenameQueue  = @($restored)
        $global:ImportRenameIndex  = 0
        $global:ImportRenameSummary = @()
        $sourceWord = if ($result.Source -eq "export") { "journey export" } else { "backup" }
        Show-ImportRenamePrompt -SourceWord $sourceWord
    } catch {
        [System.Windows.MessageBox]::Show(
            "Import failed: $($_.Exception.Message)",
            "Import",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
}

function global:Show-ImportInfoDialog {
    # Explains the two import modes before the user picks a folder (Feature 3).
    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Import Journeys"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "Two import formats are supported. Choose the folder that matches:"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "Option 1 — Backup Import`nThe folder should contain a `Journeys\` subfolder, with each journey as a sub-folder inside it (Journey A, Journey B, etc.)."))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "Option 2 — Single Journey Import`nThe folder must directly contain:"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "    entries.json`n    settings.json`n`nThe file names MUST be exactly entries.json and settings.json. The folder name may be anything.`nThese files are created by Export."))

    $nextBtn = [System.Windows.Controls.Button]::new()
    $nextBtn.Content = "NEXT"
    $nextBtn.Width = 120
    $nextBtn.HorizontalAlignment = "Left"
    $nextBtn.Margin = "0,4,0,8"
    $nextBtn.Add_Click({ Start-ImportFlow }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($nextBtn)

    $cancelBtn = [System.Windows.Controls.Button]::new()
    $cancelBtn.Content = "CANCEL"
    $cancelBtn.Width = 90
    $cancelBtn.HorizontalAlignment = "Left"
    $cancelBtn.Margin = "0,0,0,8"
    $cancelBtn.Add_Click({ Show-ContentArea "Days" }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($cancelBtn)

    Show-ContentArea "Info"
    Set-SidebarActive "backup"
}

function global:Show-ImportRenamePrompt {
    # Shows a rename prompt for each imported journey in sequence (Feature 3.1).
    # Uses $global:ImportRenameQueue, $global:ImportRenameIndex, $global:ImportRenameSummary.
    param([string]$SourceWord)

    if ($global:ImportRenameIndex -ge $global:ImportRenameQueue.Count) {
        # All journeys processed — show final summary.
        Refresh-SwitchJourneyList
        Show-ContentArea "Days"
        Set-SidebarActive "archive"
        Set-Filter "All"
        $summary = ($global:ImportRenameSummary -join "`n")
        [System.Windows.MessageBox]::Show(
            "Restored $($global:ImportRenameSummary.Count) journey(s) from $SourceWord.`n`n$summary",
            "Import"
        ) | Out-Null
        return
    }

    $j = $global:ImportRenameQueue[$global:ImportRenameIndex]
    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Imported: `"$($j.Name)`""))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "Keep the original name or rename before saving. Duplicates are not allowed."))

    $label = [System.Windows.Controls.TextBlock]::new()
    $label.Text = "JOURNEY NAME"
    $label.FontFamily = "Consolas"
    $label.FontSize = 11
    $label.Foreground = $bc.ConvertFromString("#5A6B8C")
    $label.Margin = "0,0,0,6"
    [void]$InfoAreaPanel.Children.Add($label)

    $nameBox = [System.Windows.Controls.TextBox]::new()
    $nameBox.Text = $j.Name
    $nameBox.FontFamily = "Consolas"
    $nameBox.Width = 300
    $nameBox.Padding = "8,6"
    $nameBox.Margin = "0,0,0,10"
    [void]$InfoAreaPanel.Children.Add($nameBox)

    $keepBtn = [System.Windows.Controls.Button]::new()
    $keepBtn.Content = "KEEP NAME"
    $keepBtn.Width = 120
    $keepBtn.HorizontalAlignment = "Left"
    $keepBtn.Margin = "0,0,0,8"
    $keepBtn.Add_Click({
        # Accept the current name unchanged.
        $global:ImportRenameSummary += "$($j.Name) ($($j.EntryCount) day(s))"
        $global:ImportRenameIndex++
        Show-ImportRenamePrompt -SourceWord $SourceWord
    }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($keepBtn)

    $renameBtn = [System.Windows.Controls.Button]::new()
    $renameBtn.Content = "RENAME"
    $renameBtn.Width = 120
    $renameBtn.HorizontalAlignment = "Left"
    $renameBtn.Margin = "10,0,0,8"
    $renameBtn.Add_Click({
        $newName = $nameBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($newName)) {
            [System.Windows.MessageBox]::Show("Enter a name.", "Rename") | Out-Null
            return
        }
        # Reject duplicate names.
        $existing = Get-LatestArchivedJourneys
        if ($existing | Where-Object { $_.Name -eq $newName -and $_.Path -ne $j.Path }) {
            [System.Windows.MessageBox]::Show("A journey with this name already exists.", "Rename") | Out-Null
            return
        }
        $j.Path = Rename-ArchivedJourney -ArchivePath $j.Path -NewName $newName
        $j.Name = $newName
        $global:ImportRenameSummary += "$newName ($($j.EntryCount) day(s))"
        $global:ImportRenameIndex++
        Show-ImportRenamePrompt -SourceWord $SourceWord
    }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($renameBtn)

    Show-ContentArea "Info"
    Set-SidebarActive "backup"
}

function global:Show-ArchiveBrowser {
    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Archived journeys"))

    $journeys = Get-LatestArchivedJourneys
    if ($journeys.Count -eq 0) {
        [void]$InfoAreaPanel.Children.Add((New-InfoNote "No archived journeys yet."))
    } else {
        foreach ($j in $journeys) {
            $startStr = if ($j.StartDate) { $j.StartDate } else { "no start date" }

            # Row container: Load button + Rename + Delete
            $row = [System.Windows.Controls.StackPanel]::new()
            $row.Orientation = "Horizontal"
            $row.Margin = "0,0,0,8"

            $btn = [System.Windows.Controls.Button]::new()
            $btn.Content = "LOAD  $($j.Name)  -  $startStr  -  $($j.EntryCount) day(s)"
            $btn.HorizontalContentAlignment = "Left"
            $btn.Padding = "10,6,10,6"
            $btn.FontFamily = "Consolas"
            $btn.Add_Click({
                try { Enter-ReadOnlyArchiveView -Archive $j }
                catch { [System.Windows.MessageBox]::Show("Load failed: $($_.Exception.Message)", "Productivity Tracker") | Out-Null }
            }.GetNewClosure())
            [void]$row.Children.Add($btn)

            $renameBtn = [System.Windows.Controls.Button]::new()
            $renameBtn.Content = "RENAME"
            $renameBtn.Padding = "10,6,10,6"
            $renameBtn.Margin = "6,0,0,0"
            $renameBtn.FontFamily = "Consolas"
            $renameBtn.FontSize = 11
            $renameBtn.Background = $TransparentBrush
            $renameBtn.BorderBrush = $InkBrush
            $renameBtn.BorderThickness = 1
            $renameBtn.Foreground = $InkBrush
            $renameBtn.Add_Click({
                Show-RenameDialog -Archive $j
            }.GetNewClosure())
            [void]$row.Children.Add($renameBtn)

            $deleteBtn = [System.Windows.Controls.Button]::new()
            $deleteBtn.Content = "DELETE"
            $deleteBtn.Padding = "10,6,10,6"
            $deleteBtn.Margin = "6,0,0,0"
            $deleteBtn.FontFamily = "Consolas"
            $deleteBtn.FontSize = 11
            $deleteBtn.Background = $TransparentBrush
            $deleteBtn.BorderBrush = $bc.ConvertFromString("#C55A47")
            $deleteBtn.BorderThickness = 1
            $deleteBtn.Foreground = $bc.ConvertFromString("#C55A47")
            $deleteBtn.Add_Click({
                # Guard: never delete the active journey.
                if ($j.Name -eq $global:Settings.JourneyName) {
                    [System.Windows.MessageBox]::Show(
                        "Cannot delete the active journey.",
                        "Delete",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    ) | Out-Null
                    return
                }
                $confirm = [System.Windows.MessageBox]::Show(
                    "Delete archived journey `"$($j.Name)`"? This cannot be undone.",
                    "Delete",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning
                )
                if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
                try {
                    Delete-ArchivedJourney -ArchivePath $j.Path | Out-Null
                    Refresh-SwitchJourneyList
                    Show-ArchiveBrowser
                } catch {
                    [System.Windows.MessageBox]::Show("Delete failed: $($_.Exception.Message)", "Delete") | Out-Null
                }
            }.GetNewClosure())
            [void]$row.Children.Add($deleteBtn)

            [void]$InfoAreaPanel.Children.Add($row)
        }
    }

    Show-ContentArea "Info"
    Set-SidebarActive "archive"
}

function global:Show-RenameDialog {
    # Allows renaming an archived journey (Feature 1).
    param([Parameter(Mandatory)]$Archive)

    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Rename `"$($Archive.Name)`""))

    $label = [System.Windows.Controls.TextBlock]::new()
    $label.Text = "JOURNEY NAME"
    $label.FontFamily = "Consolas"
    $label.FontSize = 11
    $label.Foreground = $bc.ConvertFromString("#5A6B8C")
    $label.Margin = "0,0,0,6"
    [void]$InfoAreaPanel.Children.Add($label)

    $nameBox = [System.Windows.Controls.TextBox]::new()
    $nameBox.Text = $Archive.Name
    $nameBox.FontFamily = "Consolas"
    $nameBox.Width = 300
    $nameBox.Padding = "8,6"
    $nameBox.Margin = "0,0,0,10"
    [void]$InfoAreaPanel.Children.Add($nameBox)

    $renameBtn = [System.Windows.Controls.Button]::new()
    $renameBtn.Content = "RENAME"
    $renameBtn.Width = 120
    $renameBtn.HorizontalAlignment = "Left"
    $renameBtn.Margin = "0,0,0,8"
    $renameBtn.Add_Click({
        $newName = $nameBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($newName)) {
            [System.Windows.MessageBox]::Show("Enter a name.", "Rename") | Out-Null
            return
        }
        try {
            Rename-ArchivedJourney -ArchivePath $Archive.Path -NewName $newName | Out-Null
            Refresh-SwitchJourneyList
            Show-ArchiveBrowser
        } catch {
            [System.Windows.MessageBox]::Show("Rename failed: $($_.Exception.Message)", "Rename") | Out-Null
        }
    }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($renameBtn)

    $cancelBtn = [System.Windows.Controls.Button]::new()
    $cancelBtn.Content = "CANCEL"
    $cancelBtn.Width = 90
    $cancelBtn.HorizontalAlignment = "Left"
    $cancelBtn.Margin = "10,0,0,8"
    $cancelBtn.Add_Click({ Show-ArchiveBrowser }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($cancelBtn)

    Show-ContentArea "Info"
    Set-SidebarActive "archive"
}

function global:Show-ExportDialog {
    # Single export dialog with checkboxes for each format (Feature 7).
    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Export Journey"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "Journey: $($global:Settings.JourneyName)"))
    [void]$InfoAreaPanel.Children.Add((New-InfoNote "Choose formats and a destination folder:"))

    # Checkboxes
    $cbXlsx = New-Object System.Windows.Controls.CheckBox
    $cbXlsx.Content = "Excel File (.xlsx)"
    $cbXlsx.FontFamily = "Consolas"
    $cbXlsx.FontSize = 12
    $cbXlsx.IsChecked = $true
    $cbXlsx.Margin = "0,0,0,6"
    [void]$InfoAreaPanel.Children.Add($cbXlsx)

    $cbCsv = New-Object System.Windows.Controls.CheckBox
    $cbCsv.Content = "CSV File (.csv)"
    $cbCsv.FontFamily = "Consolas"
    $cbCsv.FontSize = 12
    $cbCsv.IsChecked = $true
    $cbCsv.Margin = "0,0,0,6"
    [void]$InfoAreaPanel.Children.Add($cbCsv)

    $cbMd = New-Object System.Windows.Controls.CheckBox
    $cbMd.Content = "Markdown File (.md)"
    $cbMd.FontFamily = "Consolas"
    $cbMd.FontSize = 12
    $cbMd.IsChecked = $true
    $cbMd.Margin = "0,0,0,6"
    [void]$InfoAreaPanel.Children.Add($cbMd)

    $cbEntries = New-Object System.Windows.Controls.CheckBox
    $cbEntries.Content = "entries.json"
    $cbEntries.FontFamily = "Consolas"
    $cbEntries.FontSize = 12
    $cbEntries.IsChecked = $true
    $cbEntries.Margin = "0,0,0,6"
    [void]$InfoAreaPanel.Children.Add($cbEntries)

    $cbSettings = New-Object System.Windows.Controls.CheckBox
    $cbSettings.Content = "settings.json"
    $cbSettings.FontFamily = "Consolas"
    $cbSettings.FontSize = 12
    $cbSettings.IsChecked = $true
    $cbSettings.Margin = "0,0,0,12"
    [void]$InfoAreaPanel.Children.Add($cbSettings)

    # Destination
    $destRow = [System.Windows.Controls.StackPanel]::new()
    $destRow.Orientation = "Horizontal"
    $destRow.Margin = "0,0,0,10"

    $destLabel = [System.Windows.Controls.TextBlock]::new()
    $destLabel.Text = "DESTINATION FOLDER"
    $destLabel.FontFamily = "Consolas"
    $destLabel.FontSize = 11
    $destLabel.Foreground = $bc.ConvertFromString("#5A6B8C")
    $destLabel.VerticalAlignment = "Center"
    $destLabel.Margin = "0,0,10,0"
    [void]$destRow.Children.Add($destLabel)

    $destBox = [System.Windows.Controls.TextBox]::new()
    $destBox.FontFamily = "Consolas"
    $destBox.Width = 280
    $destBox.Padding = "8,6"
    $destBox.IsEnabled = $false
    [void]$destRow.Children.Add($destBox)

    $browseBtn = [System.Windows.Controls.Button]::new()
    $browseBtn.Content = "Browse..."
    $browseBtn.Padding = "8,4,8,4"
    $browseBtn.Margin = "8,0,0,0"
    $browseBtn.FontFamily = "Consolas"
    $browseBtn.FontSize = 11
    $browseBtn.Background = $TransparentBrush
    $browseBtn.BorderBrush = $InkBrush
    $browseBtn.BorderThickness = 1
    $browseBtn.Foreground = $InkBrush
    $browseBtn.Cursor = "Hand"
    $browseBtn.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Choose destination folder"
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $destBox.Text = $fbd.SelectedPath
        }
    }.GetNewClosure())
    [void]$destRow.Children.Add($browseBtn)
    [void]$InfoAreaPanel.Children.Add($destRow)

    # Action buttons
    $actionRow = [System.Windows.Controls.StackPanel]::new()
    $actionRow.Orientation = "Horizontal"
    $actionRow.Margin = "0,0,0,0"

    $exportBtn = [System.Windows.Controls.Button]::new()
    $exportBtn.Content = "EXPORT"
    $exportBtn.Width = 120
    $exportBtn.HorizontalAlignment = "Left"
    $exportBtn.Margin = "0,0,0,8"
    $exportBtn.Add_Click({
        $dest = $destBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($dest) -or -not (Test-Path $dest)) {
            [System.Windows.MessageBox]::Show("Choose a valid destination folder.", "Export") | Out-Null
            return
        }
        $exported = @()
        $errors   = @()

        if ($cbXlsx.IsChecked) {
            try {
                $path = Join-Path $dest "Journey Report.xlsx"
                Export-EntriesToXlsx -Entries $global:Entries -Path $path
                $exported += "Journey Report.xlsx"
            } catch { $errors += "Excel: $($_.Exception.Message)" }
        }
        if ($cbCsv.IsChecked) {
            try {
                $path = Join-Path $dest "Journey Report.csv"
                Export-EntriesToCsv -Entries $global:Entries -Path $path
                $exported += "Journey Report.csv"
            } catch { $errors += "CSV: $($_.Exception.Message)" }
        }
        if ($cbMd.IsChecked) {
            try {
                $path = Join-Path $dest "Journey Report.md"
                Export-EntriesToMarkdown -Entries $global:Entries -Path $path
                $exported += "Journey Report.md"
            } catch { $errors += "Markdown: $($_.Exception.Message)" }
        }
        if ($cbEntries.IsChecked) {
            try {
                $path = Join-Path $dest "entries.json"
                $global:Entries | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding UTF8
                $exported += "entries.json"
            } catch { $errors += "entries.json: $($_.Exception.Message)" }
        }
        if ($cbSettings.IsChecked) {
            try {
                $path = Join-Path $dest "settings.json"
                $global:Settings | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
                $exported += "settings.json"
            } catch { $errors += "settings.json: $($_.Exception.Message)" }
        }

        $summary = if ($exported.Count -gt 0) { "Exported:`n" + ($exported -join "`n") } else { "No files exported." }
        if ($errors.Count -gt 0) {
            $summary += "`n`nErrors:`n" + ($errors -join "`n")
        }
        [System.Windows.MessageBox]::Show($summary, "Export") | Out-Null
    }.GetNewClosure())
    [void]$actionRow.Children.Add($exportBtn)

    $cancelBtn = [System.Windows.Controls.Button]::new()
    $cancelBtn.Content = "CANCEL"
    $cancelBtn.Width = 90
    $cancelBtn.HorizontalAlignment = "Left"
    $cancelBtn.Margin = "10,0,0,8"
    $cancelBtn.Add_Click({ Show-ContentArea "Days" }.GetNewClosure())
    [void]$actionRow.Children.Add($cancelBtn)

    [void]$InfoAreaPanel.Children.Add($actionRow)

    Show-ContentArea "Info"
    Set-SidebarActive "export"
}

function global:Enter-ReadOnlyArchiveView {
    param([Parameter(Mandatory)]$Archive)

    # Entering an archive swaps $global:Entries to a different dataset - a focus
    # session left running would target the wrong day. Force-stop it first (this
    # still targets the live dataset, since the swap happens below).
    if (-not (Stop-FocusSessionAtTransition)) { return }

    $restored = Restore-JourneyFromArchive -ArchivePath $Archive.Path
    $global:ReadOnlyMode = $true
    $global:Entries = $restored.Entries
    $global:Settings = $restored.Settings
    $global:ExpandedIds = @{}

    $ReadOnlyBannerText.Text = "Viewing archived journey (read-only): $($Archive.Name)"
    $ReadOnlyBanner.Visibility = "Visible"
    $CreateJourneyBtn.Visibility = "Collapsed"
    $AddDayBtn.Visibility = "Collapsed"
    $EditTemplateBtn.Visibility = "Collapsed"
    $LoadArchiveBtn.Visibility = "Collapsed"

    $global:Filter = "All"
    Show-ContentArea "Days"
    Set-SidebarActive "dash"
    Render-Days
}

function global:Exit-ReadOnlyArchiveView {
    # Leaving an archive swaps $global:Entries back to the live dataset. A session
    # still running (it kept ticking while the archive was on screen) must not be
    # allowed to save onto the archived data - force-stop it first. Read-only mode
    # is still $true here, so the helper only offers Discard/Cancel.
    if (-not (Stop-FocusSessionAtTransition)) { return }

    $global:ReadOnlyMode = $false
    $global:Entries = $global:LiveEntries
    $global:Settings = $global:LiveSettings
    $global:ExpandedIds = @{}

    $ReadOnlyBanner.Visibility = "Collapsed"
    $LoadArchiveBtn.Visibility = "Visible"
    $CreateJourneyBtn.Visibility = "Visible"
    if ($global:Settings.StartDate) {
        $AddDayBtn.Visibility = "Visible"
        $EditTemplateBtn.Visibility = "Visible"
    } else {
        $AddDayBtn.Visibility = "Collapsed"
        $EditTemplateBtn.Visibility = "Collapsed"
    }

    $global:Filter = "All"
    Show-ContentArea "Days"
    Set-SidebarActive "dash"
    Render-Days
}

# ---------- Focus Mode ----------

function global:Get-EntrySessions {
    # Defensive: returns an entry's focus sessions, or @() when the property is
    # missing/null (e.g. an old archived journey created before Focus Mode).
    # IMPORTANT: do NOT add a leading unary comma here. PowerShell unwraps a
    # single-element array when it's written to the output pipeline via
    # "return @(...)", so a caller doing "$sessions = Get-EntrySessions $Entry"
    # would get a bare scalar (not a 1-element array) when there's exactly one
    # session - that was the original bug (the first focus record's card
    # failing to render). The fix belongs at the CALL SITE instead: every
    # caller must wrap the call in @(...), e.g. "@(Get-EntrySessions $Entry)".
    # That forces a real array regardless of what this function returns.
    # (Wrapping here too, with a comma, would double-wrap multi-session
    # results into a 1-element array containing the whole array - do not do
    # both.)
    param($Entry)
    if (-not $Entry.PSObject.Properties['Sessions'] -or $null -eq $Entry.Sessions) { return @() }
    return @($Entry.Sessions)
}

function global:Get-EntryFocusTotal {
    # Total focus minutes for an entry = sum of its session durations.
    param($Entry)
    $sum = 0
    foreach ($s in (Get-EntrySessions $Entry)) { $sum += [int]$s.duration }
    return $sum
}

function global:Format-FocusClock {
    # Timer face text. Pomodoro is always MM:SS (counting down); Stopwatch is
    # MM:SS and switches to HH:MM:SS once it passes an hour.
    param([int]$TotalSec, [string]$Mode)
    if ($Mode -eq "Pomo") {
        return ("{0:00}:{1:00}" -f [math]::Floor($TotalSec / 60), ($TotalSec % 60))
    }
    $h = [math]::Floor($TotalSec / 3600)
    $m = [math]::Floor(($TotalSec % 3600) / 60)
    $s = $TotalSec % 60
    if ($h -gt 0) { return ("{0:00}:{1:00}:{2:00}" -f $h, $m, $s) }
    return ("{0:00}:{1:00}" -f $m, $s)
}

function global:Format-FocusTotal {
    # "50 min" / "1 h 30 min" from a total in minutes (matches the mockup).
    param([int]$Minutes)
    $h = [math]::Floor($Minutes / 60)
    $m = $Minutes % 60
    if ($h -eq 0) { return "$m min" }
    if ($m -eq 0) { return "$h h" }
    return "$h h $m min"
}

function global:Format-FocusHM {
    param([datetime]$Time)
    return $Time.ToString("HH:mm")
}

function global:Format-FocusTime12 {
    # "HH:mm" stored session time → "h:mm tt" display (e.g. "9:08 AM").
    param([string]$Time24)
    try {
        $dt = [datetime]::ParseExact($Time24, "HH:mm", $null)
        return $dt.ToString("h:mm tt")
    } catch {
        return $Time24   # degrade gracefully
    }
}

function global:Format-FocusDurCompact {
    # Minutes → compact display: "1h", "50m", or "1h30m".
    param([int]$Minutes)
    $h = [math]::Floor($Minutes / 60)
    $m = $Minutes % 60
    if ($h -gt 0 -and $m -gt 0) { return "${h}h${m}m" }
    if ($h -gt 0)               { return "${h}h" }
    return "${m}m"
}

function global:Get-FocusTargetEntry {
    # The day a session logs to: the entry dated TODAY. Pre-built future days
    # (added ahead of time so goal days exist) are intentionally NOT targeted -
    # "newest by date" would log a session started today into a future day.
    # $null when there is no entry for today's calendar date.
    $todayStr = (Get-Date).ToString("yyyy-MM-dd")
    foreach ($e in $global:Entries) {
        if ($e.Date -eq $todayStr) { return $e }
    }
    return $null
}

function global:Get-OrCreateEntryForDate {
    <#
        Returns the entry for the given yyyy-MM-dd date, creating it only if
        necessary - and always through the app's existing Auto Day Creation
        path, never a second/incompatible mechanism. Used when a manually
        logged overnight focus session spills into the following calendar
        day and that day's row doesn't exist yet.

        Two-step lookup, matching the pattern Start-FocusSession and
        Write-FocusSessionRecord already use for "today":
          1. Backfill-SkippedDays - fills any backward gap between the last
             known day and today (a no-op if there's no gap, which is the
             common case since the app keeps this current on every launch).
          2. If the date is still missing (it's today or in the future
             relative to "today" - e.g. logging an overnight session on the
             day it's still in progress - Backfill never creates those), the
             day is created directly with New-Entry, the same shape every
             other day in the app has.
        Never creates a duplicate: the existing entry is returned as-is
        whenever one is already found, at either step.
    #>
    param([Parameter(Mandatory)][string]$DateStr)

    $existing = $global:Entries | Where-Object { $_.Date -eq $DateStr }
    if ($existing) { return $existing }

    $global:Entries = Backfill-SkippedDays -Entries $global:Entries -Settings $global:Settings
    $existing = $global:Entries | Where-Object { $_.Date -eq $DateStr }

    if (-not $existing) {
        $goals = @($global:Settings.GoalHours | Where-Object { $null -ne $_ })
        $newEntry = New-Entry -Date $DateStr -GoalHours $goals -ExistingEntries $global:Entries
        $global:Entries = Renumber-Entries -Entries (@($global:Entries) + $newEntry)
        $existing = $newEntry
    }

    $global:LiveEntries = $global:Entries
    Save-Entries -Entries $global:Entries
    Invoke-AutoSync
    # A brand new (or freshly backfilled) day needs to appear in the Days
    # list right away, same as Start-FocusSession does when it auto-creates
    # today's entry.
    Render-Days
    return $existing
}

function global:Show-FocusMode {
    # Sidebar leaf -> swap the content area to the Focus Mode view. The clock is
    # NOT restarted here: leaving the Focus view pauses the session, so the user
    # explicitly resumes it.
    Set-SidebarActive "focus"
    Show-ContentArea "Focus"
    Update-FocusView
}

function global:Maximize-FocusLayout {
    # Fullscreen-ish Focus view: hide the header + sidebar + splitter so the Focus
    # panel takes the whole window. The running timer is untouched (this only
    # changes visibility, never state). Sidebar geometry is remembered so
    # Restore-FocusLayout can put it back exactly, including the narrow/wide choice.
    # MinWidth/MaxWidth are Double (not GridLength), so no .Value - capture the
    # raw values. Width is a GridLength whose ToString() round-trips ("190", "1*",
    # "Auto"). If Min/Max were captured via .Value they'd come out null->"" and
    # coerce to 0 on restore, and WPF clamps the column to max-width 0 => the
    # sidebar stays invisible even though Visibility says Visible.
    $global:SavedSidebarWidth = "$($SidebarColumn.Width)"
    $global:SavedSidebarMin   = "$($SidebarColumn.MinWidth)"
    $global:SavedSidebarMax   = "$($SidebarColumn.MaxWidth)"
    $HeaderStack.Visibility       = "Collapsed"
    $HeaderDivider.Visibility     = "Collapsed"
    $Sidebar.Visibility           = "Collapsed"
    $SidebarSplitter.Visibility   = "Collapsed"
    $SidebarColumn.Width  = "0"; $SidebarColumn.MinWidth  = "0"; $SidebarColumn.MaxWidth  = "0"
    $SidebarSplitterCol.Width     = "0"
    $global:FocusMaximized = $true
    $FocusMaximizeBtn.Content = "RESTORE"
    $FocusMaximizeBtn.ToolTip = "Restore the header and sidebar"
}

function global:Restore-FocusLayout {
    # Puts the header + sidebar back exactly as they were before maximize.
    $HeaderStack.Visibility       = "Visible"
    $HeaderDivider.Visibility     = "Visible"
    $Sidebar.Visibility           = "Visible"
    $SidebarSplitter.Visibility   = "Visible"
    # Defensive defaults: fall back to the XAML geometry if a save is empty or
    # zero (a zero MaxWidth clamps the column to 0px, which makes the sidebar
    # invisible even though Visibility says Visible).
    $w = $global:SavedSidebarWidth; if ([string]::IsNullOrWhiteSpace($w) -or [double]$w -le 0) { $w = "190" }
    $m = $global:SavedSidebarMin;   if ([string]::IsNullOrWhiteSpace($m) -or [double]$m -le 0) { $m = "52" }
    $x = $global:SavedSidebarMax;   if ([string]::IsNullOrWhiteSpace($x) -or [double]$x -le 0) { $x = "260" }
    $SidebarColumn.Width  = $w
    $SidebarColumn.MinWidth = $m
    $SidebarColumn.MaxWidth = $x
    $SidebarSplitterCol.Width     = "5"
    $global:FocusMaximized = $false
    $FocusMaximizeBtn.Content = "MAXIMIZE"
    $FocusMaximizeBtn.ToolTip = "Hide sidebar and header for a distraction-free focus view"
    # The user's narrow/wide choice is re-applied to the label/icon layout.
    Set-SidebarNarrow $global:SidebarNarrow
}

function global:Toggle-FocusMaximize {
    if ($global:FocusMaximized) { Restore-FocusLayout } else { Maximize-FocusLayout }
}

function global:Update-FocusView {
    # Reflects the current focus session state into the Focus panel.
    $sessionStarted = ($global:FocusRunning -or $global:FocusElapsedSec -gt 0)

    # Mode toggle: active one is navy-filled; locked once a session has started.
    Set-FocusModeButtonVisual -Button $FocusPomoBtn -Active ($global:FocusMode -eq "Pomo") -Enabled (-not $sessionStarted)
    Set-FocusModeButtonVisual -Button $FocusStopBtn -Active ($global:FocusMode -eq "Stopwatch") -Enabled (-not $sessionStarted)

    # Focus-target label: stays editable for the entire session (running or
    # paused). Whatever is typed at STOP & SAVE time is what gets saved —
    # Save-FocusSession reads it straight from the box.
    $FocusLabelBox.IsEnabled = $true
    Update-FocusLabelWatermark

    # Timer face (elapsed recomputed from timestamps, never trusted from memory)
    $elapsedNow = Get-FocusElapsedSec
    $global:FocusElapsedSec = $elapsedNow   # cache for callers that read the global
    if ($global:FocusMode -eq "Pomo") {
        $secs = [Math]::Max(0, ($global:FocusDurationMin * 60) - $elapsedNow)
        $FocusTimerText.Text = Format-FocusClock -TotalSec $secs -Mode $global:FocusMode
    } else {
        $FocusTimerText.Text = Format-FocusClock -TotalSec $elapsedNow -Mode $global:FocusMode
    }

    # Timer circle: clickable only while idle in Pomodoro mode (opens the inline
    # duration editor). Locked once a session has started, and Stopwatch mode has
    # no target duration to edit - same gating as the mode toggle + label above.
    $FocusTimerCircle.Cursor = if (-not $sessionStarted -and $global:FocusMode -eq "Pomo") { "Hand" } else { "Arrow" }
    if ($FocusDurationEditPanel.Visibility -eq "Visible" -and $sessionStarted) {
        Close-FocusDurationEditor
    }

    # Log target line (names the day the session will log to). In read-only
    # archive view there is no live day to write to, so sessions are blocked.
    # In normal mode START auto-creates today's day, so it is always actionable.
    $target = Get-FocusTargetEntry
    if ($global:ReadOnlyMode) {
        $FocusLogTargetText.Text = "Read-only mode - focus sessions cannot be saved."
    } elseif ($target) {
        $FocusLogTargetText.Text = "Sessions are logged to Day $($target.DayNumber)  -  $($target.Date)"
    } elseif (@($global:Entries).Count -eq 0) {
        $FocusLogTargetText.Text = "No days yet - START will create today's day."
    } else {
        $FocusLogTargetText.Text = "No day for today yet - START will create it."
    }

    Render-FocusButtons
}

function global:Open-FocusDurationEditor {
    # Swaps the focus buttons for the inline duration editor, pre-filled with the
    # current Pomodoro length in minutes.
    $FocusDurationBox.Text = "$global:FocusDurationMin"
    $FocusButtonPanel.Visibility = "Collapsed"
    $FocusDurationEditPanel.Visibility = "Visible"
    [void]$FocusDurationBox.Focus()
    $FocusDurationBox.SelectAll()
}

function global:Close-FocusDurationEditor {
    # Discards any in-progress edit and restores the normal focus buttons.
    $FocusDurationEditPanel.Visibility = "Collapsed"
    $FocusButtonPanel.Visibility = "Visible"
}

function global:Confirm-FocusDuration {
    # Validates the editor input, clamps it to 2-180 minutes, and applies it as
    # the working Pomodoro length. Non-numeric input is rejected (editor stays
    # open with an explanation); out-of-range input is clamped with feedback.
    $val = 0
    if (-not [int]::TryParse($FocusDurationBox.Text.Trim(), [ref]$val)) {
        [System.Windows.MessageBox]::Show("Enter a whole number of minutes between 2 and 180.", "Focus Duration") | Out-Null
        return
    }
    if ($val -lt 2 -or $val -gt 180) {
        $val = [Math]::Max(2, [Math]::Min(180, $val))
        [System.Windows.MessageBox]::Show("Focus duration must be between 2 and 180 minutes. Set to $val min.", "Focus Duration") | Out-Null
    }
    $global:FocusDurationMin = $val
    Update-FocusView          # refresh the countdown face with the new length
    Close-FocusDurationEditor
}

function global:Set-FocusModeButtonVisual {
    param($Button, [bool]$Active, [bool]$Enabled)
    $Button.IsEnabled = $Enabled
    $Button.Opacity   = if ($Enabled) { 1 } else { 0.4 }
    if ($Active) {
        $Button.Background = $InkBrush
        $Button.Foreground = $PaperBrush
    } else {
        $Button.Background = $WhiteBrush
        $Button.Foreground = $InkBrush
    }
}

function global:Update-FocusLabelWatermark {
    # Label is now always editable — the hint shows when the box is empty and
    # hidden whenever the user has typed something.
    $FocusLabelHint.Visibility = if ([string]::IsNullOrEmpty($FocusLabelBox.Text)) { "Visible" } else { "Collapsed" }
}

function global:Render-FocusButtons {
    # Buttons change with state: idle shows a single START; a started (running or
    # paused) session shows PAUSE/RESUME + STOP & SAVE + DISCARD. In read-only
    # archive view every action is rendered disabled so a save can never write
    # the archived journey over live data.
    $FocusButtonPanel.Children.Clear()
    $canAct = -not $global:ReadOnlyMode

    if ($global:FocusElapsedSec -eq 0 -and -not $global:FocusRunning) {
        $startBtn = [System.Windows.Controls.Button]::new()
        $startBtn.Content = "START"
        $startBtn.MinWidth = 120
        $startBtn.Margin = "0,0,0,0"
        # No START when read-only. In normal mode the day to log to is created
        # on the fly inside Start-FocusSession, so the button is always active.
        $startBtn.IsEnabled = $canAct
        $startBtn.Add_Click({ Start-FocusSession }.GetNewClosure())
        [void]$FocusButtonPanel.Children.Add($startBtn)
        Apply-FocusButtonStyle $startBtn
        return
    }

    $pauseBtn = [System.Windows.Controls.Button]::new()
    $pauseBtn.Content = if ($global:FocusRunning) { "PAUSE" } else { "RESUME" }
    $pauseBtn.MinWidth = 100
    $pauseBtn.Margin = "0,0,10,0"
    $pauseBtn.IsEnabled = $canAct
    $pauseBtn.Add_Click({ Toggle-PauseFocusSession }.GetNewClosure())
    [void]$FocusButtonPanel.Children.Add($pauseBtn)
    Apply-FocusButtonStyle $pauseBtn

    $saveBtn = [System.Windows.Controls.Button]::new()
    $saveBtn.Content = "STOP & SAVE"
    $saveBtn.MinWidth = 120
    $saveBtn.Margin = "0,0,10,0"
    $saveBtn.IsEnabled = $canAct
    $saveBtn.Add_Click({ Save-FocusSession }.GetNewClosure())
    [void]$FocusButtonPanel.Children.Add($saveBtn)
    Apply-FocusButtonStyle $saveBtn

    $discardBtn = [System.Windows.Controls.Button]::new()
    $discardBtn.Content = "DISCARD"
    $discardBtn.MinWidth = 90
    $discardBtn.IsEnabled = $canAct
    $discardBtn.Add_Click({ Discard-FocusSession }.GetNewClosure())
    [void]$FocusButtonPanel.Children.Add($discardBtn)
    Apply-FocusButtonStyle $discardBtn
}

function global:Apply-FocusButtonStyle {
    # Same look as the app's top buttons (white fill, navy border) - no separate
    # danger color for Discard, per the design system.
    param($Button)
    $Button.FontFamily = "Consolas"
    $Button.FontSize = 11
    $Button.Padding = "10,6,10,6"
    $Button.Background = $WhiteBrush
    $Button.BorderBrush = $InkBrush
    $Button.BorderThickness = 1
    $Button.Foreground = $InkBrush
    $Button.Cursor = "Hand"
}

function global:Start-FocusSession {
    # A fresh session: banked time resets, a new run segment starts at now, and the
    # state is persisted immediately (a crash right after Start still recovers).
    # If today's day entry doesn't exist yet, auto-create it (with the journey's
    # goal template) so a session always has a day to log to - no + ADD DAY needed.
    if ($null -eq (Get-FocusTargetEntry) -and @($global:Settings.GoalHours).Count -gt 0) {
        $global:Entries = Backfill-SkippedDays -Entries $global:Entries -Settings $global:Settings
        $todayStr = (Get-Date).ToString("yyyy-MM-dd")
        $newEntry = New-Entry -Date $todayStr -GoalHours $global:Settings.GoalHours -ExistingEntries $global:Entries
        $global:Entries = @($global:Entries) + $newEntry
        $global:LiveEntries = $global:Entries
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync
        # Rebuild DayItems from the live set so the new day shows in the Days
        # view later. Render-Days (not Set-Filter "All") because the latter would
        # switch the content area away from the Focus panel.
        Render-Days
    }
    $global:FocusRunning = $true
    $global:FocusStatus = "RUNNING"
    $global:FocusBankedSec = 0
    $global:FocusStartTime = Get-Date
    $global:FocusElapsedSec = 0
    $global:FocusTimer.Start()
    Save-FocusSessionState
    Update-FocusView
}

function global:Toggle-PauseFocusSession {
    if ($global:FocusRunning) {
        # Pause: bank the just-completed run segment and drop the startTime, then
        # persist immediately - a crash right after Pause must still recover the
        # correct elapsed time.
        $global:FocusBankedSec = Get-FocusElapsedSec
        $global:FocusStartTime = $null
        $global:FocusRunning = $false
        $global:FocusStatus = "PAUSED"
        $global:FocusTimer.Stop()
    } else {
        # Resume: keep the banked time and start a new segment at now (Scenario 5).
        $global:FocusStartTime = Get-Date
        $global:FocusRunning = $true
        $global:FocusStatus = "RUNNING"
        $global:FocusTimer.Start()
    }
    $global:FocusElapsedSec = Get-FocusElapsedSec
    Save-FocusSessionState
    Update-FocusView
}

function global:Discard-FocusSession {
    # Throws the in-progress timer away with no session recorded.
    $global:FocusTimer.Stop()
    $global:FocusRunning = $false
    $global:FocusStatus = "PAUSED"
    $global:FocusStartTime = $null
    $global:FocusBankedSec = 0
    $global:FocusElapsedSec = 0
    $global:FocusLabel = ""
    $global:FocusMode = "Pomo"
    Clear-FocusSessionState    # discarded -> no unfinished session to persist
    Update-FocusView
}

function global:Stop-FocusSessionAtTransition {
    <#
        THE EXCEPTION to "navigation never pauses the session": journey/archive
        transitions (Enter/Exit-ReadOnlyArchiveView, create-journey, switch-journey)
        swap $global:Entries wholesale, and Get-FocusTargetEntry resolves against
        it. A session left running across one of those would keep ticking against a
        day that no longer belongs to the active dataset, and Save-FocusSession
        could log a session onto the wrong journey (or, in read-only archive view,
        write the archived dataset back over the live journey).

        So these transitions force-stop the session BEFORE the swap, letting the
        user choose what happens to it:
          Yes    -> save it to the active journey's today entry, then proceed.
          No     -> discard it, then proceed.
          Cancel -> abort the transition.
        In read-only archive view there is no live day to save to, so only
        Discard/Cancel are offered.

        Returns $true to proceed with the transition, $false to abort.
    #>
    if (-not (Get-FocusSessionInProgress)) { return $true }

    $elapsedStr = Format-FocusClock -TotalSec (Get-FocusElapsedSec) -Mode $global:FocusMode
    $saveable = -not $global:ReadOnlyMode
    $prompt  = "An unfinished focus session ($elapsedStr) is in progress.`n`n"
    $prompt += "Switching journeys or loading an archive ends it. "
    if ($saveable) {
        # YesNoCancel: Yes = save & leave, No = discard & leave, Cancel = abort.
        $prompt += "Save it to today's entry before leaving?"
        $resp = [System.Windows.MessageBox]::Show($prompt, "Focus Session In Progress",
            [System.Windows.MessageBoxButton]::YesNoCancel, [System.Windows.MessageBoxImage]::Question)
        if ($resp -eq [System.Windows.MessageBoxResult]::Yes)   { Save-FocusSession }   # logs onto the still-active (pre-swap) dataset
        elseif ($resp -eq [System.Windows.MessageBoxResult]::No) { Discard-FocusSession }
        else { return $false }                                   # Cancel -> abort the transition
    } else {
        # Read-only archive view cannot write a session anywhere - only two ways out:
        # Yes = discard & leave, No = stay in the archive (abort the transition).
        $prompt += "It cannot be saved in read-only mode, so it will be discarded."
        $resp = [System.Windows.MessageBox]::Show($prompt, "Focus Session In Progress",
            [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($resp -eq [System.Windows.MessageBoxResult]::Yes) { Discard-FocusSession }
        else { return $false }
    }
    return $true
}

function global:Get-FocusSessionInProgress {
    # A session is "in progress" if it's running or has any elapsed time.
    return ($global:FocusRunning -or $global:FocusElapsedSec -gt 0)
}

function global:Get-FocusElapsedSec {
    <#
        THE canonical elapsed computation - the in-memory tick counter is never
        trusted as the source of truth.
          RUNNING: (now - startTime) + elapsedBankedSec
          PAUSED : elapsedBankedSec
        A negative gap (backwards clock change) clamps to the banked value.
    #>
    $sec = [int]$global:FocusBankedSec
    if ($global:FocusRunning -and $global:FocusStartTime) {
        $gap = ((Get-Date) - $global:FocusStartTime).TotalSeconds
        if ($gap -gt 0) { $sec += [int]$gap }
    }
    return $sec
}

function global:Get-FocusElapsedSecFromState {
    <#
        Same canonical computation against a persisted/normalized state object,
        used by the recovery dialog before the session is restored into globals.
    #>
    param($State)
    $sec = [int]$State.elapsedBankedSec
    if ($State.status -eq "RUNNING" -and $State.startTime) {
        $gap = ((Get-Date) - $State.startTime).TotalSeconds
        if ($gap -gt 0) { $sec += [int]$gap }
    }
    return $sec
}

function global:Save-FocusSessionState {
    <#
        Persists the in-progress session to Data/focus-session.json under the
        timestamp-based model (mode/status/label/startTime/elapsedBankedSec/
        plannedDurationSec), so elapsed can always be recomputed from the file.
        Written on EVERY state change (start, pause, resume, completion, manual
        end, discard), on window close, and every 5s while running. Idle (no
        session) deletes the file. A crash immediately after Pause therefore still
        recovers the correct elapsed time.
    #>
    if (-not (Get-FocusSessionInProgress) -and $global:FocusStatus -ne "COMPLETED") {
        Clear-FocusSessionState
        return
    }
    $state = [PSCustomObject]@{
        mode               = if ($global:FocusMode -eq "Pomo") { "Pomodoro" } else { "Stopwatch" }
        status             = $global:FocusStatus
        label              = [string]$global:FocusLabel
        startTime          = if ($global:FocusStartTime) { $global:FocusStartTime.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
        elapsedBankedSec   = [int]$global:FocusBankedSec
        plannedDurationSec = if ($global:FocusMode -eq "Pomo") { [int]($global:FocusDurationMin * 60) } else { 0 }
    }
    $state | ConvertTo-Json -Depth 4 | Set-Content -Path $global:FocusStateFile -Encoding UTF8
}

function global:Clear-FocusSessionState {
    # Removes the persisted session snapshot (session completed, discarded, or
    # fully restored back into memory).
    if (Test-Path $global:FocusStateFile) {
        Remove-Item $global:FocusStateFile -Force -ErrorAction SilentlyContinue
    }
}

function global:Read-FocusSessionState {
    <#
        Loads Data/focus-session.json and normalizes it into the canonical model:
        @{ mode; status; label; startTime (datetime|null); elapsedBankedSec;
           plannedDurationSec }. Also migrates the legacy schema
        (Mode/Running/ElapsedSec/LastActive): a legacy RUNNING session maps to
        startTime=LastActive + elapsedBankedSec=ElapsedSec, so recomputing from
        timestamps reproduces exactly the old gap-credited elapsed. Returns $null
        when the file is absent or unreadable.
    #>
    if (-not (Test-Path $global:FocusStateFile)) { return $null }
    try { $data = Get-Content $global:FocusStateFile -Raw | ConvertFrom-Json } catch { $data = $null }
    if (-not $data) { return $null }

    if ($data.PSObject.Properties['status'] -and $data.PSObject.Properties['mode']) {
        $startTime = $null
        if ($data.PSObject.Properties['startTime'] -and $data.startTime) {
            try { $startTime = [datetime]::ParseExact([string]$data.startTime, "yyyy-MM-ddTHH:mm:ss", $null) } catch { $startTime = $null }
        }
        $mode   = if ([string]$data.mode -eq "Stopwatch") { "Stopwatch" } else { "Pomodoro" }
        $status = if ([string]$data.status -eq "PAUSED") { "PAUSED" }
                  elseif ([string]$data.status -eq "COMPLETED") { "COMPLETED" }
                  else { "RUNNING" }
        return [PSCustomObject]@{
            mode               = $mode
            status             = $status
            label              = [string]$data.label
            startTime          = $startTime
            elapsedBankedSec   = [Math]::Max(0, [int]$data.elapsedBankedSec)
            plannedDurationSec = [Math]::Max(0, [int]$data.plannedDurationSec)
        }
    }

    if ($data.PSObject.Properties['Mode']) {
        $wasRunning = [bool]$data.Running
        $start = $null
        if ($wasRunning -and $data.PSObject.Properties['LastActive'] -and $data.LastActive) {
            try { $start = [datetime]::ParseExact([string]$data.LastActive, "yyyy-MM-ddTHH:mm:ss", $null) } catch { $start = $null }
        }
        $mode = if ([string]$data.Mode -eq "Stopwatch") { "Stopwatch" } else { "Pomodoro" }
        return [PSCustomObject]@{
            mode               = $mode
            status             = if ($wasRunning) { "RUNNING" } else { "PAUSED" }
            label              = [string]$data.Label
            startTime          = $start
            elapsedBankedSec   = [Math]::Max(0, [int]$data.ElapsedSec)
            plannedDurationSec = if ($mode -eq "Pomodoro") { [Math]::Max(0, [int]$data.DurationMin) * 60 } else { 0 }
        }
    }
    return $null
}

function global:Write-FocusSessionRecord {
    <#
        Logs one completed focus session onto the current/today day entry and
        refreshes just that row in place (no full list rebuild). Shared by STOP &
        SAVE, recovery End Session, and recovery Complete so the recorded shape
        stays identical. Duration is rounded the same half-up way Save-FocusSession
        always did; a sub-minute session records as 1 minute.
    #>
    param([string]$Type, [string]$Label, [datetime]$Start, [datetime]$End, [int]$ElapsedSec)
    $session = [PSCustomObject]@{
        type     = $Type
        label    = $Label.Trim()
        start    = Format-FocusHM $Start
        end      = Format-FocusHM $End
        duration = [Math]::Max(1, [math]::Floor(($ElapsedSec / 60) + 0.5))
    }
    # Ensure today's day entry exists so the record always has a day to land on.
    # Sessions started live auto-create it in Start-FocusSession; sessions resolved
    # at startup (recovery Complete / End Session) reach this directly, and without
    # this guard a completed session is silently dropped when today isn't in the
    # dataset yet (e.g. the journey hasn't been advanced to today).
    $target = Get-FocusTargetEntry
    if (-not $target) {
        $global:Entries = Backfill-SkippedDays -Entries $global:Entries -Settings $global:Settings
        $todayStr = (Get-Date).ToString("yyyy-MM-dd")
        $goals = @($global:Settings.GoalHours | Where-Object { $null -ne $_ })
        $target = New-Entry -Date $todayStr -GoalHours $goals -ExistingEntries $global:Entries
        $global:Entries = @($global:Entries) + $target
        $global:LiveEntries = $global:Entries
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync
    }
    $global:Entries = Add-EntrySession -Entries $global:Entries -Id $target.Id -Session $session
    $global:LiveEntries = $global:Entries
    Save-Entries -Entries $global:Entries
    Invoke-AutoSync
    Update-FocusViewsInPlace -EntryId $target.Id
    # Ensure the entry is visible in the Days view. For a brand-new entry
    # not yet in DayItems, reconcile via Update-FilterView (in-place diff,
    # honors the current filter) rather than Render-Days — Render-Days clears
    # ExpandRegistry/FocusRegistry, destroying the expanded section the user
    # may be looking at. Update-FilterView inserts the new row into the
    # ObservableCollection; the ItemsControl's StatusChanged handler
    # materializes it, and New-DayRow reads Entry.Sessions live, so the row
    # (collapsed Total Hour and the Focus Record section) is correct without
    # any extra patching.
    Update-FilterView
}

function global:New-FieldRow {
    # A single .fieldrow from index2.html: bold 12px label on the left (46px),
    # the control on the right (flexes). Used by the Focus Record dialog.
    param([string]$Label, $Control, [int]$BottomMargin = 10)
    $grid = [System.Windows.Controls.Grid]::new()
    $grid.Margin = "0,0,0,$BottomMargin"
    $col0 = [System.Windows.Controls.ColumnDefinition]::new(); $col0.Width = [System.Windows.GridLength]::new(46)
    $col1 = [System.Windows.Controls.ColumnDefinition]::new(); $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    [void]$grid.ColumnDefinitions.Add($col0); [void]$grid.ColumnDefinitions.Add($col1)
    $lbl = [System.Windows.Controls.TextBlock]::new()
    $lbl.Text = $Label; $lbl.FontFamily = "Consolas"; $lbl.FontSize = 12
    $lbl.FontWeight = "Bold"; $lbl.Foreground = $InkBrush; $lbl.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($lbl, 0); [void]$grid.Children.Add($lbl)
    $control.VerticalAlignment = "Center"; $control.HorizontalAlignment = "Stretch"
    [System.Windows.Controls.Grid]::SetColumn($control, 1); [void]$grid.Children.Add($control)
    return $grid
}

function global:New-TimeFieldRow {
    # Start/End row: label, a time textbox (flexes), and an AM/PM combo on the right.
    param([string]$Label, $Box, $Ampm)
    $grid = [System.Windows.Controls.Grid]::new()
    $grid.Margin = "0,0,0,10"
    $col0 = [System.Windows.Controls.ColumnDefinition]::new(); $col0.Width = [System.Windows.GridLength]::new(46)
    $col1 = [System.Windows.Controls.ColumnDefinition]::new(); $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $col2 = [System.Windows.Controls.ColumnDefinition]::new(); $col2.Width = [System.Windows.GridLength]::new(58)
    [void]$grid.ColumnDefinitions.Add($col0); [void]$grid.ColumnDefinitions.Add($col1); [void]$grid.ColumnDefinitions.Add($col2)
    $lbl = [System.Windows.Controls.TextBlock]::new()
    $lbl.Text = $Label; $lbl.FontFamily = "Consolas"; $lbl.FontSize = 12
    $lbl.FontWeight = "Bold"; $lbl.Foreground = $InkBrush; $lbl.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($lbl, 0); [void]$grid.Children.Add($lbl)
    $Box.Margin = "0,0,8,0"; $Box.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($Box, 1); [void]$grid.Children.Add($Box)
    $Ampm.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($Ampm, 2); [void]$grid.Children.Add($Ampm)
    return $grid
}

function global:New-FlatButton {
    # Matches index2.html's .topbtn / .topbtn.primary: 11px Consolas, white (or
    # navy) fill, 1px #2B3A55 border, 4px corner radius. Uses a custom template
    # so the default WPF button chrome (hover overlay, focus ring) doesn't fight
    # the app look; hover swaps to #FBF3DE (secondary) / #38507A (primary).
    param([string]$Text, [switch]$Primary)
    $btn = [System.Windows.Controls.Button]::new()
    $btn.Content = $Text
    $btn.FontFamily = "Consolas"
    $btn.FontSize = 11
    $btn.Padding = "12,6,12,6"
    $btn.Cursor = "Hand"
    $btn.BorderThickness = "1"
    $btn.FocusVisualStyle = $null
    $btn.Background = if ($Primary) { $InkBrush } else { $WhiteBrush }
    $btn.BorderBrush = $InkBrush
    $btn.Foreground = if ($Primary) { $PaperBrush } else { $InkBrush }
    $tplXaml = '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Button"><Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate>'
    $btn.Template = [System.Windows.Markup.XamlReader]::Parse($tplXaml)
    $baseBg = $btn.Background
    $hoverBg = if ($Primary) { $HoverNavyBrush } else { $HoverCreamBrush }
    $btn.Add_MouseEnter({ $btn.Background = $hoverBg; if ($Primary) { $btn.BorderBrush = $hoverBg } }.GetNewClosure())
    $btn.Add_MouseLeave({ $btn.Background = $baseBg; if ($Primary) { $btn.BorderBrush = $InkBrush } }.GetNewClosure())
    return $btn
}

function global:Show-FocusRecordDialog {
    <#
        Unified modal dialog for adding AND editing focus records (index2.html).
        - Add mode ($SessionIndex not provided): shows Start/End time pickers,
          Type, Title. Save appends a new session.
        - Edit mode ($SessionIndex provided): shows read-only date range and
          duration, Type and Title editable, Delete button visible.
        Button row layout: [Delete (edit only)] --- [Cancel] [Save]
    #>
    param($Entry, [int]$SessionIndex = -1)
    $isEdit = ($SessionIndex -ge 0)
    # @() at the call site is a belt-and-suspenders guard: even if
    # Get-EntrySessions (or a future edit to it) ever returns a bare scalar
    # for a single-item result, wrapping it here forces a real array so
    # .Count and indexing behave correctly regardless.
    $sessions = @(Get-EntrySessions $Entry)
    $session = if ($isEdit -and $SessionIndex -lt $sessions.Count) { $sessions[$SessionIndex] } else { $null }

    # --- Dialog window ---
    # AllowsTransparency=True forces WPF to render via an off-screen composition
    # bitmap, which disables ClearType sub-pixel text rendering — the root cause
    # of the blurry text. Instead, use a normal opaque borderless window:
    #   WindowStyle=None  → no title bar
    #   AllowsTransparency=False → normal DWM rendering path (crisp text)
    #   DWM rounded corners (Win11) → keeps the card visual shape
    $dlg = [System.Windows.Window]::new()
    $dlg.Title = if ($isEdit) { "Edit Focus Record" } else { "Add Focus Record" }
    $dlg.Width = 352
    $dlg.SizeToContent = "Height"
    $dlg.ResizeMode = "NoResize"
    $dlg.WindowStartupLocation = "CenterOwner"
    $dlg.Owner = $window
    $dlg.FontFamily = "Consolas"
    $dlg.ShowInTaskbar = $false
    $dlg.WindowStyle = [System.Windows.WindowStyle]::None
    $dlg.AllowsTransparency = $false
    # Opaque background matching the card color — the window IS the card.
    $dlg.Background = $RowBrush

    # Enable Windows 11 DWM rounded corners (DWMWA_WINDOW_CORNER_PREFERENCE = 33,
    # DWMWCP_ROUND = 2). Fires after the HWND is created but before first paint.
    $dlg.Add_SourceInitialized({
        try {
            $h = [System.Windows.Interop.WindowInteropHelper]::new($dlg).Handle
            $pAttr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal(4)
            [System.Runtime.InteropServices.Marshal]::WriteInt32($pAttr, 2)
            [DwmApi]::DwmSetWindowAttribute([IntPtr]$h, 33, $pAttr, 4) | Out-Null
            [System.Runtime.InteropServices.Marshal]::FreeHGlobal($pAttr)
        } catch { }
    }.GetNewClosure())

    # Card fills the window exactly (no margin). The 1.5px navy border and 8px
    # corner radius match the original design; DWM rounds the window corners.
    $card = [System.Windows.Controls.Border]::new()
    $card.Background = $RowBrush
    $card.BorderBrush = $InkBrush
    $card.BorderThickness = "1.5"
    $card.CornerRadius = 8
    $card.Padding = "20,18,20,18"
    $card.Margin = "0"
    # TextFormattingMode=Display uses pixel-snapped metrics (crisper at small sizes).
    [System.Windows.Media.TextOptions]::SetTextFormattingMode($card, [System.Windows.Media.TextFormattingMode]::Display)

    $root = [System.Windows.Controls.StackPanel]::new()

    # --- Title ---
    $title = [System.Windows.Controls.TextBlock]::new()
    $title.Text = if ($isEdit) { "Edit Focus Record" } else { "Add Focus Record" }
    $title.FontSize = 14
    $title.FontWeight = "Bold"
    $title.Foreground = $InkBrush
    $title.Margin = "0,0,0,16"
    $title.Cursor = "Hand"
    # No title bar on the borderless window, so dragging the title text moves it.
    $title.Add_MouseLeftButtonDown({
        param($s, $e)
        try { $dlg.DragMove() } catch {}
    }.GetNewClosure())
    [void]$root.Children.Add($title)

    # --- Read-only time block (edit mode only) ---
    $roTimeBlock = [System.Windows.Controls.StackPanel]::new()
    $roTimeBlock.Margin = "0,0,0,14"
    if ($isEdit) {
        $roTimeBlock.Visibility = "Visible"
        $entryDate = [datetime]::ParseExact($Entry.Date, "yyyy-MM-dd", $null)
        $sp = [datetime]::MinValue; $ep = [datetime]::MinValue
        [void][datetime]::TryParseExact($session.start, "HH:mm", $null, [System.Globalization.DateTimeStyles]::None, [ref]$sp)
        [void][datetime]::TryParseExact($session.end, "HH:mm", $null, [System.Globalization.DateTimeStyles]::None, [ref]$ep)
        $sDt = $entryDate.AddHours($sp.Hour).AddMinutes($sp.Minute)
        $eDt = $entryDate.AddHours($ep.Hour).AddMinutes($ep.Minute)
        if ($eDt -le $sDt) { $sDt = $sDt.AddDays(-1) }
        # Range line with SVG clock
        $roRange = [System.Windows.Controls.StackPanel]::new()
        $roRange.Orientation = "Horizontal"
        $roRange.Margin = "0,0,0,6"
        $roClk1 = [System.Windows.Controls.Viewbox]::new()
        $roClk1.Width = 14; $roClk1.Height = 14; $roClk1.Stretch = "Uniform"; $roClk1.Margin = "0,0,8,0"
        $roClk1.VerticalAlignment = "Center"
        $roCvs1 = [System.Windows.Controls.Canvas]::new(); $roCvs1.Width = 24; $roCvs1.Height = 24
        $roE1 = [System.Windows.Shapes.Ellipse]::new(); $roE1.Width = 18; $roE1.Height = 18
        $roE1.Stroke = $InkLightBrush; $roE1.StrokeThickness = 2; $roE1.Fill = [System.Windows.Media.Brushes]::Transparent
        [System.Windows.Controls.Canvas]::SetLeft($roE1, 3); [System.Windows.Controls.Canvas]::SetTop($roE1, 3)
        [void]$roCvs1.Children.Add($roE1)
        $roP1 = [System.Windows.Shapes.Path]::new(); $roP1.Stroke = $InkLightBrush; $roP1.StrokeThickness = 2
        $roP1.StrokeLineJoin = "Round"; $roP1.StrokeStartLineCap = "Round"; $roP1.StrokeEndLineCap = "Round"
        $roP1.Data = [System.Windows.Media.StreamGeometry]::Parse("M12,7 L12,12 L15.5,13.5")
        [void]$roCvs1.Children.Add($roP1); $roClk1.Child = $roCvs1
        [void]$roRange.Children.Add($roClk1)
        $roRangeTxt = [System.Windows.Controls.TextBlock]::new()
        # Date printed once, like index2.html ("Aug 7 5:21 PM - 5:46 PM").
        $roRangeTxt.Text = "$($sDt.ToString("MMM d h:mm tt")) - $($eDt.ToString("h:mm tt"))"
        $roRangeTxt.FontFamily = "Consolas"; $roRangeTxt.FontSize = 12; $roRangeTxt.Foreground = $InkBrush
        [void]$roRange.Children.Add($roRangeTxt)
        [void]$roTimeBlock.Children.Add($roRange)
        # Duration line with SVG clock
        $roDur = [System.Windows.Controls.StackPanel]::new()
        $roDur.Orientation = "Horizontal"
        $roClk2 = [System.Windows.Controls.Viewbox]::new()
        $roClk2.Width = 14; $roClk2.Height = 14; $roClk2.Stretch = "Uniform"; $roClk2.Margin = "0,0,8,0"
        $roClk2.VerticalAlignment = "Center"
        $roCvs2 = [System.Windows.Controls.Canvas]::new(); $roCvs2.Width = 24; $roCvs2.Height = 24
        $roE2 = [System.Windows.Shapes.Ellipse]::new(); $roE2.Width = 18; $roE2.Height = 18
        $roE2.Stroke = $InkLightBrush; $roE2.StrokeThickness = 2; $roE2.Fill = [System.Windows.Media.Brushes]::Transparent
        [System.Windows.Controls.Canvas]::SetLeft($roE2, 3); [System.Windows.Controls.Canvas]::SetTop($roE2, 3)
        [void]$roCvs2.Children.Add($roE2)
        $roP2 = [System.Windows.Shapes.Path]::new(); $roP2.Stroke = $InkLightBrush; $roP2.StrokeThickness = 2
        $roP2.StrokeLineJoin = "Round"; $roP2.StrokeStartLineCap = "Round"; $roP2.StrokeEndLineCap = "Round"
        $roP2.Data = [System.Windows.Media.StreamGeometry]::Parse("M12,7 L12,12 L15.5,13.5")
        [void]$roCvs2.Children.Add($roP2); $roClk2.Child = $roCvs2
        [void]$roDur.Children.Add($roClk2)
        $roDurTxt = [System.Windows.Controls.TextBlock]::new()
        # Plain minutes ("25m"), matching the row's duration display (fmtDur).
        $roDurTxt.Text = "$([int]$session.duration)m"
        $roDurTxt.FontFamily = "Consolas"; $roDurTxt.FontSize = 12; $roDurTxt.Foreground = $InkBrush
        [void]$roDur.Children.Add($roDurTxt)
        [void]$roTimeBlock.Children.Add($roDur)
    } else {
        $roTimeBlock.Visibility = "Collapsed"
    }
    [void]$root.Children.Add($roTimeBlock)

    # --- Time fields (add mode only) - label-left rows like index2.html .fieldrow ---
    $timeFieldsBlock = [System.Windows.Controls.StackPanel]::new()
    if ($isEdit) {
        $timeFieldsBlock.Visibility = "Collapsed"
    } else {
        $timeFieldsBlock.Visibility = "Visible"
        $defaultStart = (Get-Date).AddMinutes(-25)
        $startBox = [System.Windows.Controls.TextBox]::new()
        $startBox.Text = $defaultStart.ToString("h:mm"); $startBox.FontSize = 12
        $startBox.Padding = "6,4,6,4"; $startBox.BorderBrush = $InkBrush; $startBox.BorderThickness = 1
        $startAmpm = [System.Windows.Controls.ComboBox]::new()
        $startAmpm.FontSize = 12; $startAmpm.Width = 58
        [void]$startAmpm.Items.Add("AM"); [void]$startAmpm.Items.Add("PM")
        $startAmpm.SelectedIndex = if ($defaultStart.Hour -ge 12) { 1 } else { 0 }
        [void]$timeFieldsBlock.Children.Add((New-TimeFieldRow -Label "Start" -Box $startBox -Ampm $startAmpm))

        $defaultEnd = Get-Date
        $endBox = [System.Windows.Controls.TextBox]::new()
        $endBox.Text = $defaultEnd.ToString("h:mm"); $endBox.FontSize = 12
        $endBox.Padding = "6,4,6,4"; $endBox.BorderBrush = $InkBrush; $endBox.BorderThickness = 1
        $endAmpm = [System.Windows.Controls.ComboBox]::new()
        $endAmpm.FontSize = 12; $endAmpm.Width = 58
        [void]$endAmpm.Items.Add("AM"); [void]$endAmpm.Items.Add("PM")
        $endAmpm.SelectedIndex = if ($defaultEnd.Hour -ge 12) { 1 } else { 0 }
        [void]$timeFieldsBlock.Children.Add((New-TimeFieldRow -Label "End" -Box $endBox -Ampm $endAmpm))
    }
    [void]$root.Children.Add($timeFieldsBlock)

    # --- Type ---
    $typeBox = [System.Windows.Controls.ComboBox]::new()
    $typeBox.FontSize = 12
    [void]$typeBox.Items.Add("Pomo"); [void]$typeBox.Items.Add("Stopwatch")
    if ($isEdit -and $session) {
        $typeBox.SelectedIndex = if ($session.type -eq "Pomodoro") { 0 } else { 1 }
    } else {
        $typeBox.SelectedIndex = 0
    }
    [void]$root.Children.Add((New-FieldRow -Label "Type" -Control $typeBox))

    # --- Title (with a placeholder hint overlay; WPF TextBox has no native
    # placeholder, so mirror the Focus label hint approach from MainWindow.xaml) ---
    $titleWrap = [System.Windows.Controls.Grid]::new()
    $titleBox = [System.Windows.Controls.TextBox]::new()
    $titleBox.FontSize = 12; $titleBox.Padding = "6,4,6,4"
    $titleBox.BorderBrush = $InkBrush; $titleBox.BorderThickness = 1
    $titleBox.VerticalContentAlignment = "Center"
    if ($isEdit -and $session) { $titleBox.Text = if ($session.label) { $session.label } else { "" } }
    [void]$titleWrap.Children.Add($titleBox)
    $titleHint = [System.Windows.Controls.TextBlock]::new()
    $titleHint.Text = "e.g. React part 5"
    $titleHint.FontFamily = "Consolas"; $titleHint.FontSize = 12; $titleHint.Foreground = $InkLightBrush
    $titleHint.Margin = "7,0,0,0"; $titleHint.VerticalAlignment = "Center"; $titleHint.HorizontalAlignment = "Left"
    $titleHint.IsHitTestVisible = $false
    [void]$titleWrap.Children.Add($titleHint)
    $titleBox.Add_TextChanged({
        $titleHint.Visibility = if ($titleBox.Text.Length -eq 0 -and -not $titleBox.IsKeyboardFocused) { "Visible" } else { "Collapsed" }
    }.GetNewClosure())
    $titleBox.Add_GotKeyboardFocus({ $titleHint.Visibility = "Collapsed" }.GetNewClosure())
    $titleBox.Add_LostKeyboardFocus({
        $titleHint.Visibility = if ($titleBox.Text.Length -eq 0) { "Visible" } else { "Collapsed" }
    }.GetNewClosure())
    $titleHint.Visibility = if ($titleBox.Text.Length -eq 0) { "Visible" } else { "Collapsed" }
    [void]$root.Children.Add((New-FieldRow -Label "Title" -Control $titleWrap -BottomMargin 0))

    # --- Error text (index2.html shows it between the fields and the buttons) ---
    $errorText = [System.Windows.Controls.TextBlock]::new()
    $errorText.Text = ""
    $errorText.FontSize = 11
    $errorText.Foreground = $ErrorBrush
    $errorText.TextWrapping = "Wrap"
    $errorText.Margin = "6,0,0,0"
    $errorText.Visibility = "Collapsed"
    [void]$root.Children.Add($errorText)

    # --- Button row: [Delete (edit)] --- [CANCEL] [SAVE] ---
    $btnRow = [System.Windows.Controls.Grid]::new()
    $btnRow.Margin = "0,18,0,0"
    $btnRow.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
    $btnRow.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
    $btnRow.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $btnRow.ColumnDefinitions[1].Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)

    # Delete button (edit mode only, left side) - borderless icon, hover tint
    $deleteBtn = [System.Windows.Controls.Border]::new()
    $deleteBtn.Cursor = "Hand"; $deleteBtn.Padding = "4"; $deleteBtn.ToolTip = "Delete"
    $deleteBtn.VerticalAlignment = "Center"
    $deleteBtn.CornerRadius = 4
    $deleteBtn.Visibility = if ($isEdit) { "Visible" } else { "Collapsed" }
    $deleteSvg = [System.Windows.Controls.Canvas]::new()
    $deleteSvg.Width = 17; $deleteSvg.Height = 17
    # Trash can body
    $trashPath = [System.Windows.Shapes.Path]::new()
    $trashPath.Stroke = $InkBrush; $trashPath.StrokeThickness = 1.6
    $trashPath.Fill = [System.Windows.Media.Brushes]::Transparent
    $trashPath.StrokeLineJoin = "Round"; $trashPath.StrokeStartLineCap = "Round"; $trashPath.StrokeEndLineCap = "Round"
    $trashPath.Data = [System.Windows.Media.StreamGeometry]::Parse("M6,7 L18,7 M9,7 L9,5 A1,1,0,0,1,11,4 L13,4 A1,1,0,0,1,15,5 L15,7 M8,7 L9,20 A1,1,0,0,0,10,21 L14,21 A1,1,0,0,0,15,20 L16,7")
    [void]$deleteSvg.Children.Add($trashPath)
    # Vertical lines inside
    $trashLine1 = [System.Windows.Shapes.Path]::new()
    $trashLine1.Stroke = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#5A9BD5"))
    $trashLine1.StrokeThickness = 1.6; $trashLine1.StrokeStartLineCap = "Round"; $trashLine1.StrokeEndLineCap = "Round"
    $trashLine1.Data = [System.Windows.Media.StreamGeometry]::Parse("M10,10 L10,17")
    [void]$deleteSvg.Children.Add($trashLine1)
    $trashLine2 = [System.Windows.Shapes.Path]::new()
    $trashLine2.Stroke = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#5A9BD5"))
    $trashLine2.StrokeThickness = 1.6; $trashLine2.StrokeStartLineCap = "Round"; $trashLine2.StrokeEndLineCap = "Round"
    $trashLine2.Data = [System.Windows.Media.StreamGeometry]::Parse("M14,10 L14,17")
    [void]$deleteSvg.Children.Add($trashLine2)
    $deleteBtn.Child = $deleteSvg
    # index2.html .icon-danger-btn hover tints the icon area #F0E8DC
    $deleteBtn.Add_MouseEnter({ $deleteBtn.Background = $IconHoverBrush }.GetNewClosure())
    $deleteBtn.Add_MouseLeave({ $deleteBtn.Background = $TransparentBrush }.GetNewClosure())
    $deleteBtn.Add_MouseLeftButtonUp({
        param($sender, $evtArgs)
        $evtArgs.Handled = $true
        $confirm = [System.Windows.MessageBox]::Show(
            "Delete this focus session?", "Confirm Delete",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
            $global:Entries = Remove-EntrySession -Entries $global:Entries -Id $Entry.Id -SessionIndex $SessionIndex
            $global:LiveEntries = $global:Entries
            Save-Entries -Entries $global:Entries
            Invoke-AutoSync
            Update-FocusViewsInPlace -EntryId $Entry.Id
            $dlg.Close()
        }
    }.GetNewClosure())
    [System.Windows.Controls.Grid]::SetColumn($deleteBtn, 0)
    [void]$btnRow.Children.Add($deleteBtn)

    # Cancel + Save (right side) - uppercase labels like index2.html's .topbtn
    $rightBtns = [System.Windows.Controls.StackPanel]::new()
    $rightBtns.Orientation = "Horizontal"
    $cancelBtn = New-FlatButton -Text "CANCEL"
    $cancelBtn.Margin = "0,0,8,0"
    $cancelBtn.Add_Click({ $dlg.Close() }.GetNewClosure())
    [void]$rightBtns.Children.Add($cancelBtn)
    $saveBtn = New-FlatButton -Text "SAVE" -Primary
    [void]$rightBtns.Children.Add($saveBtn)
    [System.Windows.Controls.Grid]::SetColumn($rightBtns, 1)
    [void]$btnRow.Children.Add($rightBtns)
    [void]$root.Children.Add($btnRow)

    # --- Save handler (shared) ---
    $saveBtn.Add_Click({
        if ($isEdit) {
            # Edit: update type + label only
            $typeStr = if ($typeBox.SelectedItem -eq "Pomo") { "Pomodoro" } else { "Stopwatch" }
            $newLabel = $titleBox.Text.Trim()
            $global:Entries = Set-EntrySessionType -Entries $global:Entries -Id $Entry.Id -SessionIndex $SessionIndex -NewType $typeStr
            $global:Entries = Set-EntrySessionLabel -Entries $global:Entries -Id $Entry.Id -SessionIndex $SessionIndex -NewLabel $newLabel
            $global:LiveEntries = $global:Entries
            Save-Entries -Entries $global:Entries
            Invoke-AutoSync
            Update-FocusViewsInPlace -EntryId $Entry.Id
            $dlg.Close()
        } else {
            # Add: parse the typed times. "h:mm tt" has no date component, so
            # TryParseExact would otherwise silently default the date to
            # today - always anchor to the row's own date instead, the same
            # way the read-only edit-mode block above (~line 3232) already
            # reconstructs session.start/session.end against $Entry.Date.
            [datetime]$startT = [datetime]::MinValue
            [datetime]$endT   = [datetime]::MinValue
            $startCombined = "$($startBox.Text.Trim()) $($startAmpm.SelectedItem)"
            $endCombined   = "$($endBox.Text.Trim()) $($endAmpm.SelectedItem)"
            $startOk = [datetime]::TryParseExact($startCombined, "h:mm tt", $null, [System.Globalization.DateTimeStyles]::None, [ref]$startT)
            $endOk   = [datetime]::TryParseExact($endCombined,   "h:mm tt", $null, [System.Globalization.DateTimeStyles]::None, [ref]$endT)
            if (-not $startOk -or -not $endOk) {
                $errorText.Text = "Enter start and end as h:mm AM/PM (e.g. 9:30 AM)."
                $errorText.Visibility = "Visible"; return
            }

            $rowDate = [datetime]::ParseExact($Entry.Date, "yyyy-MM-dd", $null).Date
            $startDT = $rowDate.AddHours($startT.Hour).AddMinutes($startT.Minute)
            $endDT   = $rowDate.AddHours($endT.Hour).AddMinutes($endT.Minute)

            # Exact equality is almost certainly a typo, not a 24-hour
            # session - reject it outright rather than letting it fall
            # through to the overnight rule below and get capped to 180.
            if ($endDT -eq $startDT) {
                $errorText.Text = "End time must be different from Start time."
                $errorText.Visibility = "Visible"; return
            }

            # Overnight / cross-midnight detection: compare the full
            # reconstructed instants, never just .Hour (which ignores
            # minutes and can misread a same-day typo like 2:15 PM -> 1:45 PM
            # as a session spanning almost a full day). Only roll to the next
            # calendar day when End is genuinely earlier than Start.
            if ($endDT -lt $startDT) {
                $endDT = $endDT.AddDays(1)
            }

            # Hard cap: no single focus session logs more than 180 minutes.
            # Longer spans are truncated to exactly 180 minutes from Start,
            # never rejected - this is also what keeps an ambiguous same-day
            # typo (above) from ever being stored as a near-24-hour session.
            $actualMinutes = ($endDT - $startDT).TotalMinutes
            $cappedMinutes = [Math]::Min([math]::Round($actualMinutes), 180)
            $cappedEndDT   = $startDT.AddMinutes($cappedMinutes)

            # Future-time validation runs against the actual reconstructed
            # (and capped) dates, so a historical row's late-evening time is
            # never confused with today, and an overnight end that lands on
            # today is checked against the real current moment rather than a
            # value silently shifted onto today's date.
            $now = Get-Date
            if ($startDT -gt $now) {
                $errorText.Text = "Start time cannot be in the future."
                $errorText.Visibility = "Visible"; return
            }
            if ($cappedEndDT -gt $now) {
                $errorText.Text = "End time cannot be in the future."
                $errorText.Visibility = "Visible"; return
            }

            $typeStr  = if ($typeBox.SelectedItem -eq "Pomo") { "Pomodoro" } else { "Stopwatch" }
            $labelStr = $titleBox.Text.Trim()
            $midnight = $rowDate.AddDays(1)
            $isOvernight = $cappedEndDT -gt $midnight
            $nextEntry = $null

            if (-not $isOvernight) {
                # Ordinary same-day session (the common case, including a
                # session capped down to same-day per the note above).
                $newSession = [PSCustomObject]@{
                    type = $typeStr; label = $labelStr
                    start = Format-FocusHM $startDT; end = Format-FocusHM $cappedEndDT
                    duration = [int][Math]::Max(1, $cappedMinutes)
                }
                $global:Entries = Add-EntrySession -Entries $global:Entries -Id $Entry.Id -Session $newSession
            } else {
                # Genuine overnight: split the (already-capped) duration at
                # midnight and log each portion on its own calendar day,
                # matching the app's one-row-per-day data model. The next
                # day's entry is found or created through the app's existing
                # Auto Day Creation path (Get-OrCreateEntryForDate) - it is
                # updated in place if it already exists, never duplicated.
                $day1Minutes = [int][Math]::Max(1, [math]::Round(($midnight - $startDT).TotalMinutes))
                $day2Minutes = [int][Math]::Max(1, $cappedMinutes - $day1Minutes)

                $session1 = [PSCustomObject]@{
                    type = $typeStr; label = $labelStr
                    start = Format-FocusHM $startDT; end = "00:00"
                    duration = $day1Minutes
                }
                $session2 = [PSCustomObject]@{
                    type = $typeStr; label = $labelStr
                    start = "00:00"; end = Format-FocusHM ($midnight.AddMinutes($day2Minutes))
                    duration = $day2Minutes
                }

                $nextEntry = Get-OrCreateEntryForDate -DateStr ($midnight.ToString("yyyy-MM-dd"))
                $global:Entries = Add-EntrySession -Entries $global:Entries -Id $Entry.Id -Session $session1
                $global:Entries = Add-EntrySession -Entries $global:Entries -Id $nextEntry.Id -Session $session2
            }

            $global:LiveEntries = $global:Entries
            Save-Entries -Entries $global:Entries
            Invoke-AutoSync
            Update-FocusViewsInPlace -EntryId $Entry.Id
            if ($isOvernight) {
                Update-FocusViewsInPlace -EntryId $nextEntry.Id
            }
            $dlg.Close()
        }
    }.GetNewClosure())

    $card.Child = $root
    $dlg.Content = $card
    $dlg.ShowDialog() | Out-Null
}

function global:Show-FocusRecoveryDialog {
    <#
        Dedicated modal WPF window (never MessageBox) that resolves a RUNNING
        session found at startup. Two modes:
          "Resume"    -> "Focus session detected" with live Elapsed/Remaining text
                         ticking once per second (own DispatcherTimer); Resume /
                         End Session buttons.
          "Completed" -> "Welcome back!" completion prompt; Complete / Discard.
        A Resume dialog whose elapsed crosses the planned duration is swapped to
        the Completed mode on the spot (threshold transition), so Resume is never
        offered after expiration. Every button recomputes elapsed from the
        persisted timestamps at click time - never from values captured when the
        dialog opened.
        Returns @{ Action = "Resume"|"EndSession"|"Complete"|"Discard";
                   ElapsedSec = <seconds at click time> }.
    #>
    param($State, [ValidateSet("Resume","Completed")][string]$Mode)
    try {
        [xml]$xaml = Get-Content -Path (Join-Path $root "RecoveryWindow.xaml") -Raw
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $dlg = [System.Windows.Markup.XamlReader]::Load($reader)

        $ResumePanel     = $dlg.FindName("ResumePanel")
        $CompletionPanel = $dlg.FindName("CompletionPanel")
        $LabelText       = $dlg.FindName("LabelText")
        $ElapsedText     = $dlg.FindName("ElapsedText")
        $RemainingText   = $dlg.FindName("RemainingText")
        $ResumeBtn       = $dlg.FindName("ResumeBtn")
        $EndSessionBtn   = $dlg.FindName("EndSessionBtn")
        $CompleteBtn     = $dlg.FindName("CompleteBtn")
        $DiscardBtn      = $dlg.FindName("DiscardBtn")
        $CompletionLabelText   = $dlg.FindName("CompletionLabelText")
        $CompletionDurationText = $dlg.FindName("CompletionDurationText")

        $clockMode = if ($State.mode -eq "Pomodoro") { "Pomo" } else { "Stopwatch" }
        $labelStr  = if ([string]::IsNullOrWhiteSpace([string]$State.label)) { "(untitled)" } else { [string]$State.label }
        $isPomoExpiring = ($State.mode -eq "Pomodoro" -and $State.plannedDurationSec -gt 0)

        # --- 1 Hz refresh: recompute elapsed from the persisted timestamps and
        # drive the threshold transition (Resume -> Completed when expired). ---
        $global:FocusRecoveryResult = $null
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(1)
        $timer.Add_Tick({
            $elapsed  = Get-FocusElapsedSecFromState -State $State
            $remaining = if ($isPomoExpiring) { [Math]::Max(0, $State.plannedDurationSec - $elapsed) } else { 0 }
            $ElapsedText.Text   = "Elapsed:   " + (Format-FocusClock -TotalSec $elapsed  -Mode $clockMode)
            $RemainingText.Text = "Remaining: " + (Format-FocusClock -TotalSec $remaining -Mode $clockMode)
            if ($isPomoExpiring -and $elapsed -ge $State.plannedDurationSec) {
                # Threshold transition: never leave Resume available after expiry.
                $ResumePanel.Visibility = "Collapsed"
                $CompletionPanel.Visibility = "Visible"
                $CompletionLabelText.Text = $labelStr
                $CompletionDurationText.Text = "$([math]::Round($State.plannedDurationSec / 60)) minutes"
                $timer.Stop()
            }
        }.GetNewClosure())

        # --- Buttons: every value is computed at click time. ---
        $ResumeBtn.Add_Click({
            # Bank the elapsed-at-click and start a fresh run segment at now
            # (remaining = plannedDuration - elapsed at click).
            $State.elapsedBankedSec = [int](Get-FocusElapsedSecFromState -State $State)
            $State.startTime = Get-Date
            $State.status = "RUNNING"
            $global:FocusRecoveryResult = @{ Action = "Resume"; ElapsedSec = [int]$State.elapsedBankedSec }
            $timer.Stop()
            $dlg.Close()
        }.GetNewClosure())
        $EndSessionBtn.Add_Click({
            $global:FocusRecoveryResult = @{ Action = "EndSession"; ElapsedSec = [int](Get-FocusElapsedSecFromState -State $State) }
            $timer.Stop()
            $dlg.Close()
        }.GetNewClosure())
        $CompleteBtn.Add_Click({
            $global:FocusRecoveryResult = @{ Action = "Complete"; ElapsedSec = 0 }
            $dlg.Close()
        }.GetNewClosure())
        $DiscardBtn.Add_Click({
            $global:FocusRecoveryResult = @{ Action = "Discard"; ElapsedSec = 0 }
            $dlg.Close()
        }.GetNewClosure())

        # --- Initial paint ---
        $LabelText.Text = $labelStr
        if ($Mode -eq "Completed") {
            $ResumePanel.Visibility = "Collapsed"
            $CompletionPanel.Visibility = "Visible"
            $CompletionLabelText.Text = $labelStr
            $CompletionDurationText.Text = "$([math]::Round($State.plannedDurationSec / 60)) minutes"
        } else {
            $elapsed  = Get-FocusElapsedSecFromState -State $State
            $remaining = if ($isPomoExpiring) { [Math]::Max(0, $State.plannedDurationSec - $elapsed) } else { 0 }
            $ElapsedText.Text   = "Elapsed:   " + (Format-FocusClock -TotalSec $elapsed  -Mode $clockMode)
            $RemainingText.Text = "Remaining: " + (Format-FocusClock -TotalSec $remaining -Mode $clockMode)
            $timer.Start()
        }

        $dlg.ShowDialog() | Out-Null
        $timer.Stop()
        if (-not $global:FocusRecoveryResult) {
            # Closed without a button (e.g. Alt+F4) - safe default: discard.
            $global:FocusRecoveryResult = @{ Action = "Discard"; ElapsedSec = 0 }
        }
        $result = $global:FocusRecoveryResult
        $global:FocusRecoveryResult = $null
        return $result
    } catch {
        # Never let a dialog failure block the whole app: discard the session and
        # continue with an idle launch.
        $global:FocusRecoveryResult = $null
        return @{ Action = "Discard"; ElapsedSec = 0 }
    }
}

function global:Resolve-FocusRecovery {
    <#
        Startup gate, run BEFORE the window opens. Reads the persisted session and
        resolves it so the user can never begin a second session on top of an
        unrecovered one:
          COMPLETED -> already finalized; clear and continue idle.
          PAUSED    -> restore the paused state silently (spec).
          RUNNING   -> elapsed < target  -> live Resume dialog (Scenario 1/2);
                       elapsed >= target -> completion dialog that logs exactly
                       plannedDuration on Complete (Scenario 3).
        Returns the state to restore into the app, or $null to continue idle.
    #>
    $state = Read-FocusSessionState
    if (-not $state) { return $null }

    if ($state.status -eq "COMPLETED") {
        Clear-FocusSessionState
        return $null
    }
    if ($state.status -eq "PAUSED") {
        return $state
    }

    # RUNNING: decide Scenario 1 vs Scenario 3 from the persisted timestamps.
    $elapsedNow = Get-FocusElapsedSecFromState -State $state
    $expired = ($state.mode -eq "Pomodoro" -and $state.plannedDurationSec -gt 0 -and $elapsedNow -ge $state.plannedDurationSec)

    if ($expired) {
        $choice = Show-FocusRecoveryDialog -State $state -Mode "Completed"
        if ($choice.Action -eq "Complete") {
            # Log exactly plannedDuration - never the real-world elapsed (a 25-min
            # Pomodoro with the laptop off 90 min records 25 minutes).
            $endNow = Get-Date
            Write-FocusSessionRecord -Type "Pomodoro" -Label $state.label -Start $endNow.AddSeconds(-$state.plannedDurationSec) -End $endNow -ElapsedSec $state.plannedDurationSec
        }
        # The session genuinely finished while the app was closed: notify in both
        # paths (Complete and Discard).
        Show-FocusCompletionNotification -Label $state.label -DurationMin ([int][math]::Round($state.plannedDurationSec / 60))
        Clear-FocusSessionState
        return $null
    }

    $choice = Show-FocusRecoveryDialog -State $state -Mode "Resume"
    if ($choice.Action -eq "EndSession") {
        $elapsed = [int]$choice.ElapsedSec
        $endNow  = Get-Date
        $type    = if ($state.mode -eq "Pomodoro") { "Pomodoro" } else { "Stopwatch" }
        Write-FocusSessionRecord -Type $type -Label $state.label -Start $endNow.AddSeconds(-$elapsed) -End $endNow -ElapsedSec $elapsed
        Clear-FocusSessionState
        return $null
    }
    if ($choice.Action -eq "Discard") {
        Clear-FocusSessionState
        return $null
    }
    # Resume: the dialog banked the elapsed-at-click and restarted the run segment.
    return $state
}

function global:Restore-FocusSessionIntoApp {
    <#
        Applies a resolved/recovered session state into the running app's globals
        and jumps to the Focus view. A RUNNING session keeps running (its startTime
        is already in the past and the clock is credited); a PAUSED one stays
        paused until the user hits RESUME. Persists immediately so a crash right
        after restore still recovers the correct elapsed.
    #>
    param($State)
    $global:FocusMode = if ($State.mode -eq "Stopwatch") { "Stopwatch" } else { "Pomo" }
    if ($State.mode -eq "Pomodoro" -and $State.plannedDurationSec -gt 0) {
        $global:FocusDurationMin = [Math]::Max(2, [Math]::Min(180, [math]::Round($State.plannedDurationSec / 60)))
    }
    $global:FocusBankedSec   = [int]$State.elapsedBankedSec
    $global:FocusStartTime   = $State.startTime
    $global:FocusRunning     = ($State.status -eq "RUNNING")
    $global:FocusStatus      = $State.status
    $global:FocusElapsedSec  = Get-FocusElapsedSec
    $global:FocusLabel       = [string]$State.label
    $FocusLabelBox.Text      = [string]$State.label
    $global:FocusLabelWasEmpty = [string]::IsNullOrEmpty($FocusLabelBox.Text)
    Save-FocusSessionState
    if ($global:FocusRunning) { $global:FocusTimer.Start() }
    Show-FocusMode
}

function global:Cleanup-FocusNotification {
    # Closes the on-screen completion toast. Idempotent - also called before showing
    # a new one so a rapid second completion can't stack toasts.
    if ($global:FocusToastWindow) {
        $global:FocusToastWindow.Close()
        $global:FocusToastWindow = $null
    }
}

function global:Show-FocusCompletionNotification {
    <#
        On-screen completion toast for a session that finished naturally (a Pomodoro
        hitting 0:00 while the app is open, or one that completed while the app was
        closed). A small always-on-top WPF window in the top-right of the screen
        shows even while the main window is minimized or another view is active.
        It stays up until the user clicks OK (or the app closes) - never auto-dismisses.
        Best-effort: a failure here never breaks the app.
    #>
    param([string]$Label, [int]$DurationMin)
    try {
        Cleanup-FocusNotification
        $body = if ([string]::IsNullOrWhiteSpace($Label)) { "Duration: $DurationMin min" } else { "`"$Label`" - $DurationMin min" }
        $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="300" SizeToContent="Height" WindowStyle="None" ResizeMode="NoResize"
        AllowsTransparency="True" Background="Transparent" Topmost="True"
        ShowActivated="False" ShowInTaskbar="False">
    <Border CornerRadius="10" Background="#F4F0E6" BorderBrush="#2B3A55" BorderThickness="1.5"
            Padding="14,12">
        <Border.Effect>
            <DropShadowEffect BlurRadius="14" ShadowDepth="2" Opacity="0.30" Color="#2B3A55"/>
        </Border.Effect>
        <StackPanel>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Ellipse Grid.Column="0" Width="10" Height="10" Fill="#2F6E4F" VerticalAlignment="Top" Margin="2,5,10,0"/>
                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="FOCUS SESSION COMPLETE" FontFamily="Consolas" FontWeight="Bold" FontSize="12" Foreground="#2B3A55"/>
                    <TextBlock x:Name="BodyText" FontFamily="Consolas" FontSize="11" Foreground="#5A6B8C" TextWrapping="Wrap" Margin="0,4,0,0"/>
                </StackPanel>
            </Grid>
            <Border Height="1" Background="#DCD3BE" Margin="0,12,0,0"/>
            <Button x:Name="OkBtn" Content="OK" MinWidth="72" HorizontalAlignment="Center" Margin="0,12,0,0"
                    Padding="18,5" Background="#FFFFFF" BorderBrush="#2B3A55" BorderThickness="1"
                    Foreground="#2B3A55" FontFamily="Consolas" FontSize="11" Cursor="Hand"/>
        </StackPanel>
    </Border>
</Window>
'@
        $toast = [System.Windows.Markup.XamlReader]::Parse($xaml)
        $toast.FindName("BodyText").Text = $body
        # Render off-screen first, then place in the top-right corner so the final
        # position never flickers in.
        $toast.Left = -10000; $toast.Top = -10000
        $toast.Show()
        $toast.UpdateLayout()
        $work = [System.Windows.SystemParameters]::WorkArea
        $toast.Left = $work.Right - $toast.ActualWidth - 16
        $toast.Top  = $work.Top + 16
        # Only the OK button dismisses the toast - it stays up until acknowledged
        # (or the app closes, which also calls Cleanup-FocusNotification).
        $toast.FindName("OkBtn").Add_Click({ Cleanup-FocusNotification }.GetNewClosure())
        $global:FocusToastWindow = $toast
    } catch {
        Cleanup-FocusNotification
    }
}

function global:Save-FocusSession {
    <#
        STOP & SAVE (or a Pomodoro that reached 0:00): records a session on the
        current/today day and resets the view to idle. A session shorter than one
        second is discarded instead. Status COMPLETED is persisted immediately
        (Scenario 4), so a crash right after can't re-offer this session.
    #>
    $elapsed = Get-FocusElapsedSec
    if ($elapsed -lt 1) { Discard-FocusSession; return }
    $global:FocusTimer.Stop()
    $global:FocusRunning = $false
    $global:FocusStatus  = "COMPLETED"
    Save-FocusSessionState    # persist completion immediately

    $endTime   = Get-Date
    $startTime = $endTime.AddSeconds(-$elapsed)
    # Label is read straight from the box now (it's editable for the whole
    # session, and we stopped writing the global on every keystroke).
    $global:FocusLabel = $FocusLabelBox.Text
    $sessionType = if ($global:FocusMode -eq "Pomo") { "Pomodoro" } else { "Stopwatch" }
    Write-FocusSessionRecord -Type $sessionType -Label $global:FocusLabel -Start $startTime -End $endTime -ElapsedSec $elapsed

    $global:FocusElapsedSec  = 0
    $global:FocusBankedSec   = 0
    $global:FocusStartTime   = $null
    $global:FocusStatus      = "PAUSED"
    $global:FocusLabel       = ""
    $global:FocusMode        = "Pomo"
    Clear-FocusSessionState    # recorded -> no unfinished session to persist
    Update-FocusView
}

function global:Test-ControlIsLive {
    # True only if $Control is still attached to a rendered PresentationSource
    # (i.e. actually part of the on-screen visual tree right now). A control
    # that was built earlier but whose container was since recycled/replaced by
    # Populate-RealizedRows (virtualized ItemsControl) fails this check even
    # though the registry still holds a reference to it - patching such a
    # detached control succeeds silently (no exception) but is invisible to
    # the user, which is exactly the "totals update, card doesn't appear" bug.
    param($Control)
    if (-not $Control) { return $false }
    try {
        return $null -ne [System.Windows.PresentationSource]::FromVisual($Control)
    } catch {
        return $false
    }
}

function global:Update-FocusViewsInPlace {
    # After a session is saved, patch the cached day-row visuals for that entry
    # only: the collapsed "Total Hour" text and, if the row is expanded, the
    # FOCUS RECORD + TOTAL FOCUS sections. No full re-render.
    param([string]$EntryId)
    $entry = $global:Entries | Where-Object { $_.Id -eq $EntryId }
    if (-not $entry) { return }

    if ($global:TotalRegistry.ContainsKey($EntryId)) {
        if (Test-ControlIsLive $global:TotalRegistry[$EntryId]) {
            $global:TotalRegistry[$EntryId].Text = "Total Hour : $(Format-FocusTotal (Get-EntryFocusTotal $entry))"
        } else {
            # Stale reference: the collapsed row's control was recycled to a
            # different container. Drop it so a future Populate-RealizedRows /
            # Render-Days repopulates it; nothing else to patch here.
            $global:TotalRegistry.Remove($EntryId)
        }
    }

    if ($global:FocusRegistry.ContainsKey($EntryId)) {
        $ref = $global:FocusRegistry[$EntryId]
        if (Test-ControlIsLive $ref.Timeline) {
            $ref.Timeline.Children.Clear()
            Add-FocusTimelineRows -Timeline $ref.Timeline -Entry $entry
            $ref.Total.Text = "$ClockGlyph  $(Format-FocusTotal (Get-EntryFocusTotal $entry))"
            # Defensive: force WPF to re-measure/re-arrange the panel and its
            # ancestor row right away instead of waiting for the next layout
            # pass, closing the "race with WPF layout" window called out below.
            $ref.Timeline.InvalidateMeasure()
            $ref.Timeline.InvalidateArrange()
            if ($global:RowRegistry.ContainsKey($EntryId)) {
                $global:RowRegistry[$EntryId].InvalidateVisual()
                $global:RowRegistry[$EntryId].UpdateLayout()
            }
        } else {
            # Registry pointed at a Timeline panel that is no longer part of the
            # live visual tree (its row container was recycled/replaced since
            # the row was expanded). Patching it is a silent no-op from the
            # user's perspective: totals update elsewhere, the card doesn't.
            # Rebuild the expanded section fresh, from the live entry, and
            # re-register it so the fix isn't just cosmetic-once.
            $global:FocusRegistry.Remove($EntryId)
            $global:TotalRegistry.Remove($EntryId)
            $global:ExpandRegistry.Remove($EntryId)
            if ($global:StackRegistry.ContainsKey($EntryId) -and (Test-ControlIsLive $global:StackRegistry[$EntryId])) {
                $stack = $global:StackRegistry[$EntryId]
                # Remove any old (now-orphaned) expanded-section child before adding the new one.
                for ($i = $stack.Children.Count - 1; $i -ge 1; $i--) { $stack.Children.RemoveAt($i) }
                $region = New-ExpandedSection -Entry $entry
                [void]$stack.Children.Add($region)
                $region.Visibility = "Visible"
                # New-ExpandedSection already calls Add-FocusTimelineRows with the
                # current (live) entry, so the freshly built card is correct as-is.
            }
        }
    }
}

function global:New-SectionHeading {
    # 11px bold uppercase section header used inside the expanded day row.
    param([string]$Text)
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $Text
    $t.FontFamily = "Consolas"
    $t.FontSize = 11
    $t.FontWeight = "Bold"
    $t.Foreground = $InkBrush
    $t.Margin = "0,6,0,6"
    return $t
}

function global:New-SectionDivider {
    # Thin dashed horizontal rule separating the expanded row's sections.
    # ("2,2" is coerced to a DoubleCollection by WPF's type converter.)
    $r = [System.Windows.Shapes.Rectangle]::new()
    $r.Height = 1
    $r.Margin = "0,6,0,2"
    $r.Stroke = $InkBrush
    $r.StrokeDashArray = "2,2"
    $r.StrokeThickness = 1
    $r.HorizontalAlignment = "Stretch"
    return $r
}

function global:Add-FocusTimelineRows {
    # Populates the FOCUS RECORD panel: each session is a card-style row
    # (cream background, rounded, shadow on hover) with a 34px navy clock
    # badge, time range + label, and compact duration. Clickable → opens the
    # unified edit dialog. Matches index2.html exactly.
    param($Timeline, $Entry)
    # @() call-site guard - see comment on the same pattern in
    # Show-FocusRecordDialog. This is the exact spot that produced the
    # "first record saves but the card never renders" bug: with exactly one
    # session, $sessions.Count used to be $null, so the render loop below
    # never ran even though the session data was there all along.
    $sessions = @(Get-EntrySessions $Entry)
    if ($sessions.Count -eq 0) {
        $empty = [System.Windows.Controls.TextBlock]::new()
        $empty.Text = "No focus records yet."
        $empty.FontFamily = "Consolas"
        $empty.FontSize = 12
        $empty.Foreground = $InkLightBrush
        [void]$Timeline.Children.Add($empty)
        return
    }

    for ($i = 0; $i -lt $sessions.Count; $i++) {
        $s = $sessions[$i]
        $si = $i
        $shortType = if ($s.type -eq "Pomodoro") { "Pomo" } else { "Stopwatch" }
        $labelPart = if ($s.label) { "$shortType : $($s.label)" } else { "$shortType : (no title)" }
        $durText = "$([int]$s.duration)m"

        # --- Card-style row: [34px clock badge] [time + label ... duration] ---
        $contentRow = [System.Windows.Controls.Grid]::new()
        $contentRow.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        $contentRow.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        $contentRow.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        $contentRow.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(34)
        $contentRow.ColumnDefinitions[1].Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $contentRow.ColumnDefinitions[2].Width = [System.Windows.GridLength]::new(40)

        # 34px navy circle badge with SVG clock icon
        $circle = [System.Windows.Controls.Border]::new()
        $circle.Width = 34
        $circle.Height = 34
        $circle.CornerRadius = 17
        $circle.Background = $InkBrush
        $circle.HorizontalAlignment = "Center"
        $circle.VerticalAlignment = "Center"
        $circle.ClipToBounds = $true

        # SVG clock: Viewbox(24x24) → Canvas with Ellipse + Path
        $clockViewbox = [System.Windows.Controls.Viewbox]::new()
        $clockViewbox.Width = 16
        $clockViewbox.Height = 16
        $clockViewbox.Stretch = "Uniform"
        $clockCanvas = [System.Windows.Controls.Canvas]::new()
        $clockCanvas.Width = 24
        $clockCanvas.Height = 24
        $clockFace = [System.Windows.Shapes.Ellipse]::new()
        $clockFace.Width = 18; $clockFace.Height = 18
        $clockFace.Stroke = $PaperBrush; $clockFace.StrokeThickness = 2
        $clockFace.Fill = [System.Windows.Media.Brushes]::Transparent
        [System.Windows.Controls.Canvas]::SetLeft($clockFace, 3)
        [System.Windows.Controls.Canvas]::SetTop($clockFace, 3)
        [void]$clockCanvas.Children.Add($clockFace)
        $clockHand = [System.Windows.Shapes.Path]::new()
        $clockHand.Stroke = $PaperBrush; $clockHand.StrokeThickness = 2
        $clockHand.StrokeLineJoin = "Round"
        $clockHand.StrokeStartLineCap = "Round"; $clockHand.StrokeEndLineCap = "Round"
        $clockHand.Data = [System.Windows.Media.StreamGeometry]::Parse("M12,7 L12,12 L15.5,13.5")
        [void]$clockCanvas.Children.Add($clockHand)
        $clockViewbox.Child = $clockCanvas
        $circle.Child = $clockViewbox
        [System.Windows.Controls.Grid]::SetColumn($circle, 0)
        [void]$contentRow.Children.Add($circle)

        # Time range + label text block
        $infoPanel = [System.Windows.Controls.StackPanel]::new()
        $infoPanel.Orientation = "Vertical"
        $infoPanel.Margin = "14,0,0,0"
        $infoPanel.VerticalAlignment = "Center"
        $timeText = [System.Windows.Controls.TextBlock]::new()
        $timeText.Text = "$(Format-FocusTime12 $s.start) - $(Format-FocusTime12 $s.end)"
        $timeText.FontFamily = "Consolas"; $timeText.FontSize = 13
        $timeText.Foreground = $InkBrush
        [void]$infoPanel.Children.Add($timeText)
        $labelText = [System.Windows.Controls.TextBlock]::new()
        $labelText.Text = $labelPart
        $labelText.FontFamily = "Consolas"; $labelText.FontSize = 12
        $labelText.Foreground = $InkLightBrush; $labelText.Margin = "0,3,0,0"
        # index2.html renders the type : title line flat - no shadow.
        [void]$infoPanel.Children.Add($labelText)
        [System.Windows.Controls.Grid]::SetColumn($infoPanel, 1)
        [void]$contentRow.Children.Add($infoPanel)

        # Duration on the far right (bold)
        $durLabel = [System.Windows.Controls.TextBlock]::new()
        $durLabel.Text = $durText
        $durLabel.FontFamily = "Consolas"; $durLabel.FontSize = 12
        $durLabel.FontWeight = "Bold"; $durLabel.Foreground = $InkLightBrush
        $durLabel.HorizontalAlignment = "Right"; $durLabel.VerticalAlignment = "Center"
        $durLabel.Margin = "0,0,8,0"
        [System.Windows.Controls.Grid]::SetColumn($durLabel, 2)
        [void]$contentRow.Children.Add($durLabel)

        # Card-style row wrapper: white bg (contrasts against the cream day-row
        # container so the cards stand out), rounded, hover highlight.
        $rowBorder = [System.Windows.Controls.Border]::new()
        $rowBorder.Cursor = "Hand"
        $rowBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(251, 243, 222))
        $rowBorder.BorderBrush = $DividerDotBrush
        $rowBorder.BorderThickness = 1
        $rowBorder.CornerRadius = 8
        $rowBorder.Padding = "14,12,14,12"   # 12px vertical / 14px horizontal (index2.html .record-row)
        $rowBorder.Margin = "0,0,0,10"
        $rowBorder.Child = $contentRow
        # Pixel-snapped text metrics for crisper rendering at small sizes.
        [System.Windows.Media.TextOptions]::SetTextFormattingMode($rowBorder, [System.Windows.Media.TextFormattingMode]::Display)
        # Hover: tint the background instead of using DropShadowEffect, which
        # forces WPF to an off-screen bitmap and degrades ClearType text rendering.
        $rowBorder.Add_MouseEnter({
            $rowBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255, 255, 255))
        }.GetNewClosure())
        $rowBorder.Add_MouseLeave({
            $rowBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(251, 243, 222))
        }.GetNewClosure())
        $rowBorder.Add_MouseLeftButtonUp({
            param($sender, $evtArgs)
            $evtArgs.Handled = $true
            Show-FocusRecordDialog -Entry $Entry -SessionIndex $si
        }.GetNewClosure())
        [void]$Timeline.Children.Add($rowBorder)
    }
}

# ---------- Report / Analytics ----------

function global:Get-ProductiveHoursForDay {
    # Goals are thresholds: productive hours for a day = the HIGHEST goal hour that
    # was completed (4+6 done -> 6, not 10; 4+6+9 done -> 9; none done -> 0).
    # Tight foreach: tracks the max completed threshold instead of pipelines.
    param($Entry)
    $max = 0
    foreach ($g in $Entry.Goals) {
        if ($g.Status -eq "done" -and $g.Hours -gt $max) { $max = $g.Hours }
    }
    return $max
}

function global:Get-StreakInfo {
    <#
        Given a per-entry "completed" list and its matching Dates (both in
        date order), returns the longest run of consecutive true values and
        the indices of that run. Dates is optional for backward
        compatibility, but when supplied a run also breaks whenever two
        consecutive completed entries aren't exactly one calendar day apart -
        so a real gap (e.g. a day entry that got deleted) can't silently
        stitch two streaks together by array position alone. Backfill-
        SkippedDays normally prevents gaps from existing at all; this is a
        second, independent safety net.
    #>
    param(
        [bool[]]$Completed,
        [string[]]$Dates = $null
    )
    $longest = 0
    $longestStart = -1
    $longestEnd = -1
    $runStart = -1
    for ($i = 0; $i -lt $Completed.Count; $i++) {
        if ($Completed[$i]) {
            $brokeByGap = $false
            if ($runStart -ge 0 -and $Dates -and $Dates.Count -eq $Completed.Count) {
                $prevDate = [datetime]::ParseExact($Dates[$i - 1], "yyyy-MM-dd", $null)
                $curDate  = [datetime]::ParseExact($Dates[$i],     "yyyy-MM-dd", $null)
                if (($curDate - $prevDate).Days -ne 1) { $brokeByGap = $true }
            }
            if ($runStart -lt 0 -or $brokeByGap) { $runStart = $i }
            $len = $i - $runStart + 1
            if ($len -gt $longest) {
                $longest = $len
                $longestStart = $runStart
                $longestEnd = $i
            }
        } else {
            $runStart = -1
        }
    }
    return [PSCustomObject]@{ Longest = $longest; Start = $longestStart; End = $longestEnd }
}

function global:New-ReportCard {
    # A bordered card in the day-row visual language (RowBrush fill, InkBrush
    # border, 6px radius) so the Report panel reads like the rest of the app
    # instead of a wall of plain text.
    param($Child, $Background = $RowBrush, $BorderBrush = $InkBrush, [int]$BottomMargin = 12)
    $border = [System.Windows.Controls.Border]::new()
    $border.Background      = $Background
    $border.BorderBrush     = $BorderBrush
    $border.BorderThickness = 1.5
    $border.CornerRadius    = 6
    $border.Padding         = "14,12,14,12"
    $border.Margin          = "0,0,20,$BottomMargin"
    $border.HorizontalAlignment = "Stretch"
    $border.Child = $Child
    return $border
}

function global:New-ReportLabel {
    # Muted monospace label line, used inside report cards.
    param([string]$Text, $Color = $InkLightBrush, [int]$FontSize = 11, [int]$BottomMargin = 2)
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $Text
    $t.FontFamily = "Consolas"
    $t.FontSize = $FontSize
    $t.Foreground = $Color
    $t.Margin = "0,0,0,$BottomMargin"
    $t.TextWrapping = "Wrap"
    return $t
}

function global:New-CompletionBar {
    # Thin filled-proportion bar (existing brushes) that communicates completion %
    # at a glance, using the same green/amber/red semantics as the goal chips.
    param([int]$Pct, [int]$Width = 220)
    $outer = [System.Windows.Controls.Border]::new()
    $outer.Width = $Width
    $outer.Height = 12
    $outer.CornerRadius = 6
    $outer.Background = $DividerBrush
    $outer.BorderBrush = $InkBrush
    $outer.BorderThickness = 1
    $outer.Padding = "1,1,1,1"
    $outer.HorizontalAlignment = "Left"
    $outer.Margin = "0,2,0,8"
    $fill = [System.Windows.Controls.Border]::new()
    $fill.HorizontalAlignment = "Left"
    $fill.CornerRadius = 4
    $fill.Background = if ($Pct -ge 60) { $ChipDoneBd } elseif ($Pct -ge 35) { $AmberBrush } else { $ChipMissBd }
    $fill.Width = [math]::Round([Math]::Min(100, [Math]::Max(0, $Pct)) / 100 * ($Width - 2))
    $outer.Child = $fill
    return $outer
}

function global:Show-Report {
    $InfoAreaPanel.Children.Clear()
    [void]$InfoAreaPanel.Children.Add((New-InfoHeading "Report / Analytics"))

    # Only count days up to and including today. Future blank days (added ahead
    # of time so goal days exist) don't belong in the report: they'd inflate
    # Total Days, drag down completion %, and break streaks with days that
    # haven't happened yet. Current Day is computed separately from
    # Settings.StartDate and is unaffected.
    $todayStr = (Get-Date).ToString("yyyy-MM-dd")
    $sorted = @($global:Entries | Where-Object { $_.Date -le $todayStr } |
                 Sort-Object { [datetime]$_.Date })
    if ($sorted.Count -eq 0) {
        [void]$InfoAreaPanel.Children.Add((New-InfoNote "No data yet."))
        Show-ContentArea "Info"
        Set-SidebarActive "dash"
        Set-DashboardActiveButton "Report"
        return
    }

    # ---- Journey summary card ----
    $journeyName = if ($global:Settings.JourneyName) { $global:Settings.JourneyName } else { "(untitled)" }
    $startDate = if ($global:Settings.StartDate) { $global:Settings.StartDate } else { "-" }
    $totalDays = $sorted.Count
    $currentDay = try {
        if ($global:Settings.StartDate) {
            [math]::Max(1, ((Get-Date).Date - [datetime]::ParseExact($global:Settings.StartDate, "yyyy-MM-dd", $null)).Days + 1)
        } else { $totalDays }
    } catch { $totalDays }

    $sumStack = [System.Windows.Controls.StackPanel]::new()
    [void]$sumStack.Children.Add((New-ReportLabel -Text "JOURNEY" -Color $InkLightBrush -FontSize 10 -BottomMargin 4))
    [void]$sumStack.Children.Add((New-ReportLabel -Text $journeyName -Color $InkBrush -FontSize 14 -BottomMargin 8))
    [void]$sumStack.Children.Add((New-ReportLabel -Text "Start Date : $startDate" -Color $InkBrush))
    [void]$sumStack.Children.Add((New-ReportLabel -Text "Current Day : $currentDay" -Color $InkBrush))
    [void]$sumStack.Children.Add((New-ReportLabel -Text "Total Days : $totalDays" -Color $InkBrush -BottomMargin 0))
    [void]$InfoAreaPanel.Children.Add((New-ReportCard -Child $sumStack))

    # ---- Per-goal stats ----
    # Goal list = the union of every distinct Hours value that actually appears
    # across entries' Goals arrays (sorted ascending). The current template
    # ($Settings.GoalHours) only affects NEW days - a goal dropped from the
    # template mid-journey would otherwise vanish from the report even though
    # earlier entries still track it.
    $goalSet = @{}
    foreach ($en in $sorted) {
        foreach ($g in $en.Goals) { $goalSet[[int]$g.Hours] = $true }
    }
    $goalList = @($goalSet.Keys | Sort-Object)

    if ($goalList.Count -gt 0) {
        $goalsHeading = [System.Windows.Controls.TextBlock]::new()
        $goalsHeading.Text = "GOAL PERFORMANCE"
        $goalsHeading.FontFamily = "Consolas"
        $goalsHeading.FontSize = 12
        $goalsHeading.FontWeight = "Bold"
        $goalsHeading.Foreground = $InkBrush
        $goalsHeading.Margin = "0,4,0,8"
        [void]$InfoAreaPanel.Children.Add($goalsHeading)

        foreach ($hours in $goalList) {
            # Count a threshold over ONLY the days that actually tracked it (a goal
            # dropped from the template stops existing partway, so later days neither
            # help nor hurt its historical numbers or streak).
            $completedList = [System.Collections.Generic.List[bool]]::new()
            $dateList = [System.Collections.Generic.List[string]]::new()
            $doneCount = 0
            $trackedDays = 0
            foreach ($en in $sorted) {
                $present = $false; $hit = $false
                foreach ($g in $en.Goals) {
                    if ($g.Hours -eq $hours) {
                        $present = $true
                        if ($g.Status -eq "done") { $hit = $true; break }
                    }
                }
                if ($present) {
                    $trackedDays++
                    $completedList.Add($hit)
                    $dateList.Add($en.Date)
                    if ($hit) { $doneCount++ }
                }
            }
            $completed = $completedList.ToArray()
            $dates = $dateList.ToArray()
            $streak = Get-StreakInfo -Completed $completed -Dates $dates

            $startDateStr = "-"; $endDateStr = "-"
            if ($streak.Longest -gt 0 -and $streak.Start -ge 0) {
                $startDateStr = $dateList[$streak.Start]
                $endDateStr   = $dateList[$streak.End]
            }

            $pct = if ($trackedDays -gt 0) { [math]::Round(($doneCount / $trackedDays) * 100) } else { 0 }
            # Green-leaning for a strong completion rate, orange/red-leaning for a
            # weak one - the same chip palette used on the day rows.
            if ($pct -ge 75)     { $bg = $ChipDoneBg; $bd = $ChipDoneBd; $fg = $ChipDoneFg }
            elseif ($pct -lt 50) { $bg = $ChipMissBg; $bd = $ChipMissBd; $fg = $ChipMissFg }
            else                 { $bg = $RowBrush;   $bd = $InkBrush;   $fg = $InkBrush }

            $cardStack = [System.Windows.Controls.StackPanel]::new()

            $hdr = [System.Windows.Controls.Grid]::new()
            $cdA = [System.Windows.Controls.ColumnDefinition]::new(); $cdA.Width = "*"
            $cdB = [System.Windows.Controls.ColumnDefinition]::new(); $cdB.Width = "Auto"
            [void]$hdr.ColumnDefinitions.Add($cdA)
            [void]$hdr.ColumnDefinitions.Add($cdB)

            $title = [System.Windows.Controls.TextBlock]::new()
            $title.Text = "  {0}h  GOAL" -f $hours
            $title.FontFamily = "Consolas"
            $title.FontSize = 14
            $title.FontWeight = "Bold"
            $title.Foreground = $fg
            $title.VerticalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($title, 0)
            [void]$hdr.Children.Add($title)

            $badge = [System.Windows.Controls.Border]::new()
            $badge.Background = $bd
            $badge.CornerRadius = 10
            $badge.Padding = "10,3,10,3"
            $badge.VerticalAlignment = "Center"
            $badgeText = [System.Windows.Controls.TextBlock]::new()
            $badgeText.Text = "$pct%"
            $badgeText.FontFamily = "Consolas"
            $badgeText.FontSize = 12
            $badgeText.FontWeight = "Bold"
            $badgeText.Foreground = $WhiteBrush
            $badge.Child = $badgeText
            [System.Windows.Controls.Grid]::SetColumn($badge, 1)
            [void]$hdr.Children.Add($badge)

            [void]$cardStack.Children.Add($hdr)
            [void]$cardStack.Children.Add((New-ReportLabel -Text "Highest Streak : $($streak.Longest) days  ($startDateStr -> $endDateStr)" -Color $fg -BottomMargin 2))
            [void]$cardStack.Children.Add((New-ReportLabel -Text "Total ${hours}h Completed : $doneCount of $trackedDays tracked day(s)" -Color $fg -BottomMargin 4))
            [void]$cardStack.Children.Add((New-CompletionBar -Pct $pct))

            [void]$InfoAreaPanel.Children.Add((New-ReportCard -Child $cardStack -Background $bg -BorderBrush $bd))
        }
    }

    # ---- Totals card ----
    # Get-ProductiveHoursForDay iterates each entry's own Goals array (highest done
    # threshold), so dropped-template goals still count toward the hours they were
    # actually completed for.
    $phSum = 0
    foreach ($e in $sorted) {
        $phSum += (Get-ProductiveHoursForDay -Entry $e)
    }
    $avgPH = if ($totalDays -gt 0) { [math]::Round($phSum / $totalDays, 1) } else { 0 }

    $totStack = [System.Windows.Controls.StackPanel]::new()
    [void]$totStack.Children.Add((New-ReportLabel -Text "PRODUCTIVE HOURS" -Color $InkLightBrush -FontSize 10 -BottomMargin 4))
    [void]$totStack.Children.Add((New-ReportLabel -Text "$phSum hours" -Color $InkBrush -FontSize 16 -BottomMargin 2))
    [void]$totStack.Children.Add((New-ReportLabel -Text "Average : $avgPH h/day" -Color $InkBrush -BottomMargin 0))
    [void]$InfoAreaPanel.Children.Add((New-ReportCard -Child $totStack))

    # ---- Overall card ----
    # Completion % iterates each entry's Goals directly, so it also counts goals no
    # longer in the current template.
    $doneGoals = 0; $totalGoals = 0; $missedDays = 0
    foreach ($e in $sorted) {
        $d = 0; $t = 0
        foreach ($g in $e.Goals) { $t++; if ($g.Status -eq "done") { $d++ } }
        $doneGoals += $d; $totalGoals += $t
        if ($t -gt 0 -and $d -lt $t) { $missedDays++ }
    }
    $overallPct = if ($totalGoals -gt 0) { [math]::Round(($doneGoals / $totalGoals) * 100) } else { 0 }

    $ovStack = [System.Windows.Controls.StackPanel]::new()
    [void]$ovStack.Children.Add((New-ReportLabel -Text "OVERALL COMPLETION" -Color $InkLightBrush -FontSize 10 -BottomMargin 4))
    [void]$ovStack.Children.Add((New-ReportLabel -Text "$overallPct%" -Color $InkBrush -FontSize 16 -BottomMargin 2))
    [void]$ovStack.Children.Add((New-CompletionBar -Pct $overallPct))
    [void]$ovStack.Children.Add((New-ReportLabel -Text "Missed Days : $missedDays" -Color $InkBrush -BottomMargin 0))
    [void]$InfoAreaPanel.Children.Add((New-ReportCard -Child $ovStack))

    # Back button
    $backBtn = [System.Windows.Controls.Button]::new()
    $backBtn.Content = "BACK TO DAYS"
    $backBtn.Width = 130
    $backBtn.HorizontalAlignment = "Left"
    $backBtn.Margin = "0,4,0,8"
    $backBtn.Add_Click({
        Set-Filter "All"
    }.GetNewClosure())
    [void]$InfoAreaPanel.Children.Add($backBtn)

    Show-ContentArea "Info"
    Set-SidebarActive "dash"
    Set-DashboardActiveButton "Report"
}

# ---------- Event wiring ----------
$FilterAllBtn.Add_Click({ Set-Filter "All" })
$FilterCompletedBtn.Add_Click({ Set-Filter "Completed" })
$FilterMissedBtn.Add_Click({ Set-Filter "Missed" })
$ReportBtn.Add_Click({ Show-Report })

$DashboardHeaderBtn.Add_MouseLeftButtonUp({
    if ($global:SidebarNarrow) { return }
    Toggle-SidebarSection "dash"
    if ($global:ActiveSidebarSection -eq "dash") { Set-Filter "All" } else { Show-ContentArea "Days" }
})
$ExportHeaderBtn.Add_MouseLeftButtonUp({
    if ($global:SidebarNarrow) { return }
    Toggle-SidebarSection "export"
})
$SwitchJourneyBtn.Add_MouseLeftButtonUp({
    if ($global:SidebarNarrow) { return }
    Refresh-SwitchJourneyList
    Toggle-SidebarSection "switch"
})
$LoadArchiveBtn.Add_MouseLeftButtonUp({ Show-ArchiveBrowser })
$FocusModeBtn.Add_MouseLeftButtonUp({ Show-FocusMode })

$FocusPomoBtn.Add_Click({
    if ($global:FocusRunning -or $global:FocusElapsedSec -gt 0) { return }
    $global:FocusMode = "Pomo"
    Update-FocusView
})
$FocusStopBtn.Add_Click({
    if ($global:FocusRunning -or $global:FocusElapsedSec -gt 0) { return }
    $global:FocusMode = "Stopwatch"
    Update-FocusView
})
# Maximize toggle (hides header + sidebar so the Focus view fills the window).
# Toggling never touches the timer state - purely a visibility change.
$FocusMaximizeBtn.Add_Click({
    Toggle-FocusMaximize
})
$FocusLabelBox.Add_TextChanged({
    # Keep the watermark in sync only when the empty/non-empty boundary is
    # crossed (first/last character) - the Visibility toggle that fires on every
    # keystroke re-measures the Focus panel, and after the first character the
    # hint is already hidden. $global:FocusLabel itself is read from the box's
    # text at STOP & SAVE time, so it isn't written on every keystroke.
    $nowEmpty = [string]::IsNullOrEmpty($FocusLabelBox.Text)
    if ($nowEmpty -ne $global:FocusLabelWasEmpty) {
        $global:FocusLabelWasEmpty = $nowEmpty
        Update-FocusLabelWatermark
    }
})
$FocusLabelBox.Add_GotFocus({
    $FocusLabelHint.Visibility = "Collapsed"
})
# When focus leaves the label box, re-evaluate the watermark so it reappears
# if the box is empty — Update-FocusLabelWatermark checks text content.
$FocusLabelBox.Add_LostFocus({
    Update-FocusLabelWatermark
})

# Clicking the timer circle opens the inline duration editor - only while idle in
# Pomodoro mode (Update-FocusView also sets the cursor to match). Started sessions
# and Stopwatch mode are locked, matching the mode-toggle + label gating.
$FocusTimerCircle.Add_MouseLeftButtonUp({
    if ($global:FocusRunning -or $global:FocusElapsedSec -gt 0) { return }
    if ($global:FocusMode -ne "Pomo") { return }
    Open-FocusDurationEditor
})
$FocusDurationOkBtn.Add_Click({
    Confirm-FocusDuration
})
$FocusDurationCancelBtn.Add_Click({
    Close-FocusDurationEditor
})
$FocusDurationBox.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
        Confirm-FocusDuration
    } elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) {
        Close-FocusDurationEditor
    }
})

$BackupHeaderBtn.Add_MouseLeftButtonUp({
    if ($global:SidebarNarrow) { return }
    Toggle-SidebarSection "backup"
})

$SetBackupFolderBtn.Add_Click({
    Show-BackupExplanationDialog
})

$ImportBtn.Add_Click({
    Show-ImportInfoDialog
})

$CreateJourneyBtn.Add_Click({
    $global:JourneyPanelMode = "New"
    $JourneyPanelTitle.Text = "SET UP YOUR JOURNEY"
    $JourneyDateRow.Visibility = "Visible"
    $SetTemplateBtn.Content = "SET TEMPLATE FOR JOURNEY"
    $global:PendingGoalHours = @()
    $JourneyNameBox.Text = ""
    $JourneyDatePicker.SelectedDate = Get-Date
    Render-GoalEditChips
    $NewGoalHoursBox.Text = ""

    # Hide CREATE NEW JOURNEY button and empty state while panel is open
    $CreateJourneyBtn.Visibility = "Collapsed"
    $EmptyStateText.Visibility = "Collapsed"
    $StartJourneyPanel.Visibility = "Visible"
    Set-JourneyEditButtons -Disabled $true
    Show-ContentArea "Days"
})

$EditTemplateBtn.Add_Click({
    $global:JourneyPanelMode = "Edit"
    $JourneyPanelTitle.Text = "EDIT GOAL TEMPLATE"
    $JourneyDateRow.Visibility = "Collapsed"
    $SetTemplateBtn.Content = "SAVE TEMPLATE CHANGES"
    $global:PendingGoalHours = @($global:Settings.GoalHours)
    $JourneyNameBox.Text = $global:Settings.JourneyName
    Render-GoalEditChips
    $NewGoalHoursBox.Text = ""

    # Hide CREATE NEW JOURNEY button while editing
    $CreateJourneyBtn.Visibility = "Collapsed"
    $StartJourneyPanel.Visibility = "Visible"
    Set-JourneyEditButtons -Disabled $true
    Show-ContentArea "Days"
})

$AddGoalBtn.Add_Click({ Add-PendingGoalFromInput })

$CancelJourneyBtn.Add_Click({
    # Close the journey panel without saving. The day rows are unchanged, so
    # reconcile visibility instead of rebuilding the whole list.
    $StartJourneyPanel.Visibility = "Collapsed"
    $CreateJourneyBtn.Visibility = "Visible"
    $DaysScrollViewer.Visibility = "Visible"
    Set-JourneyEditButtons -Disabled $false
    Update-FilterView
})

$NewGoalHoursBox.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
        Add-PendingGoalFromInput
        $e.Handled = $true
    }
})

$SetTemplateBtn.Add_Click({
    $journeyName = $JourneyNameBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($journeyName)) {
        [System.Windows.MessageBox]::Show("Enter a name for this journey.", "Productivity Tracker") | Out-Null
        return
    }

    if ($global:PendingGoalHours.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Add at least one goal.", "Productivity Tracker") | Out-Null
        return
    }

    if ($global:JourneyPanelMode -eq "Edit") {
        # Capture the live backup folder BEFORE the name changes, so a rename can be
        # reconciled below - without this, Get-LiveJourneyDir (called from
        # Invoke-AutoSync) derives its path from the NEW name and just starts a fresh
        # folder there, leaving the old-named folder (and its live.json/live.xlsx and
        # any rotated snapshots) behind untouched. That looked like renaming had
        # "copied" the journey instead of renaming it.
        $oldJourneyDir = Get-LiveJourneyDir
        $oldJourneyName = $global:Settings.JourneyName

        $global:Settings.GoalHours = $global:PendingGoalHours
        $global:Settings.JourneyName = $journeyName
        $global:LiveSettings = $global:Settings
        Save-Settings -Settings $global:Settings

        # Also fold in any Data\Archive folder(s) still sitting under the OLD
        # name (leftover from a previous Switch Journey / Create New Journey
        # that archived this same journey before it was renamed) - otherwise
        # they linger forever as a separate "ghost" entry in SWITCH JOURNEY /
        # LOAD ARCHIVE alongside the new name.
        Reconcile-RenamedJourneyArchives -OldName $oldJourneyName -NewName $journeyName

        $newJourneyDir = Get-LiveJourneyDir
        if ($oldJourneyDir -ne $newJourneyDir -and (Test-Path $oldJourneyDir)) {
            if (-not (Test-Path $newJourneyDir)) {
                # No folder exists under the new name yet - just rename it in place.
                $parentDir = Split-Path $newJourneyDir -Parent
                if (-not (Test-Path $parentDir)) { New-Item -ItemType Directory -Path $parentDir -Force | Out-Null }
                Rename-Item -Path $oldJourneyDir -NewName (Split-Path $newJourneyDir -Leaf)
            } else {
                # A folder for the new name already exists (e.g. renaming back to a
                # name used before) - fold the old folder's files into it, then remove
                # the old folder, so nothing is silently duplicated.
                Get-ChildItem -Path $oldJourneyDir -File -ErrorAction SilentlyContinue | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $newJourneyDir -Force
                }
                Remove-Item -Path $oldJourneyDir -Recurse -Force
            }
        }

        $StartJourneyPanel.Visibility = "Collapsed"
        $CreateJourneyBtn.Visibility = "Visible"
        Set-JourneyEditButtons -Disabled $false
        Show-ContentArea "Days"
        Invoke-AutoSync
        # Editing the template only affects FUTURE days; existing rows are unchanged,
        # so reconcile visibility instead of rebuilding the whole list.
        Update-FilterView
        [System.Windows.MessageBox]::Show(
            "Goal template updated. This applies to days you add from now on - days already logged keep their original goals.",
            "Productivity Tracker"
        ) | Out-Null
        return
    }

    if (-not $JourneyDatePicker.SelectedDate) {
        [System.Windows.MessageBox]::Show("Pick a start date.", "Productivity Tracker") | Out-Null
        return
    }

    # Starting a fresh journey clears $global:Entries - a running focus session
    # would otherwise keep ticking against a dataset that's about to be emptied.
    # Force-stop it before the old journey is archived, so a "save" still lands
    # on the current journey's today entry.
    if (-not (Stop-FocusSessionAtTransition)) { return }

    # Archive the old journey if one exists
    if ($global:Settings.StartDate) {
        Save-JourneyArchive -Entries $global:Entries -Settings $global:Settings | Out-Null
    }

    $global:Settings.StartDate = $JourneyDatePicker.SelectedDate.ToString("yyyy-MM-dd")
    $global:Settings.GoalHours = $global:PendingGoalHours
    $global:Settings.JourneyName = $journeyName
    $global:LiveSettings = $global:Settings
    Save-Settings -Settings $global:Settings

    $global:Entries  = @()
    $global:LiveEntries = $global:Entries
    $global:ExpandedIds = @{}
    Save-Entries -Entries $global:Entries

    $StartJourneyPanel.Visibility = "Collapsed"
    $CreateJourneyBtn.Visibility = "Visible"
    $AddDayBtn.Visibility = "Visible"
    $EditTemplateBtn.Visibility = "Visible"
    $BulkDeleteToggleBtn.Visibility = "Visible"
    Reset-SelectMode
    Set-JourneyEditButtons -Disabled $false
    Invoke-AutoSync
    Render-Days
})

$AddDayBtn.Add_Click({
    if ($global:ReadOnlyMode) { return }
    $global:Entries = Backfill-SkippedDays -Entries $global:Entries -Settings $global:Settings
    $dateStr = (Get-Date).ToString("yyyy-MM-dd")
    if (@($global:Entries | Where-Object { $_.Date -eq $dateStr })) {
        # Backfill (or an earlier click) already created today - don't duplicate it.
        $global:LiveEntries = $global:Entries
        Save-Entries -Entries $global:Entries
        Invoke-AutoSync
        Set-Filter "All"
        [System.Windows.MessageBox]::Show("A day for today already exists. Use the Focus mode to track time for today.", "Add Day", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }
    $newEntry = New-Entry -Date $dateStr -GoalHours $global:Settings.GoalHours -ExistingEntries $global:Entries
    $global:Entries = @($global:Entries) + $newEntry
    $global:LiveEntries = $global:Entries
    Save-Entries -Entries $global:Entries
    Invoke-AutoSync

    # Set-Filter "All" reconciles DayItems, inserting the new row at the top of
    # the virtualized list; the ItemsControl materializes it immediately since
    # scroll is at the top after adding.
    Set-Filter "All"
})

$ExportBtn.Add_Click({
    Show-ExportDialog
})

$BackToCurrentBtn.Add_Click({
    Exit-ReadOnlyArchiveView
})

$SyncBtn.Add_Click({
    Invoke-AutoSync
    Flush-AutoSync   # write live.xlsx + status now so the failure check below is synchronous
    if ($SyncStatusText.Text -notlike "Last synced:*") {
        [System.Windows.MessageBox]::Show(
            "Sync failed. If a live.xlsx or checkpoint .xlsx file is open in Excel, close it and try again.",
            "Sync failed",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        return
    }

    try {
        $journeysDir = Sync-AllJourneysToBackup
        [System.Windows.MessageBox]::Show(
            "Synced.`n`nEverything lives under: $journeysDir`n(live safety-net copy + up to 3 checkpoint versions per journey)",
            "Productivity Tracker"
        ) | Out-Null
    } catch {
        [System.Windows.MessageBox]::Show(
            "Current journey synced, but backing up every journey failed: $($_.Exception.Message)",
            "Sync failed"
        ) | Out-Null
    } finally {
        Set-SyncPending $false
    }
})

# ---------- Bulk-select delete buttons ----------
$BulkDeleteToggleBtn.Add_Click({
    if ($global:SelectModeActive) { Exit-SelectMode } else { Enter-SelectMode }
})

$BulkDeleteConfirmBtn.Add_Click({ Invoke-BulkDelete })

$BulkDeleteCancelBtn.Add_Click({ Exit-SelectMode })

# ---------- Virtualized day-list plumbing ----------
# Whenever the item generator realizes containers (initial layout, scrolling,
# add/delete, filter changes), rebuild the realized rows' content so each
# container shows its item. Standard mode keeps containers alive (no recycling
# jumps) so scrolling is smooth.
$global:DayListGenerator = $DaysList.ItemContainerGenerator
$global:DayListGenerator.add_StatusChanged({
    $st = $global:DayListGenerator.Status
    if ($st -eq [System.Windows.Controls.Primitives.GeneratorStatus]::ContainersGenerated -or
        $st -eq [System.Windows.Controls.Primitives.GeneratorStatus]::Ready) {
        Populate-RealizedRows
    }
}.GetNewClosure())

# ---------- Keyboard scrolling ----------
# Arrow Up/Down scrolls the day list smoothly. Clicking anywhere on the list
# gives the ListBox focus so arrow keys work immediately.
$DaysList.Add_PreviewMouseLeftButtonDown({
    param($s, $e)
    $DaysList.Focus()
}.GetNewClosure())
$DaysList.Add_PreviewKeyDown({
    param($s, $e)
    # Walk the visual tree to find the ListBox's internal ScrollViewer.
    $sv = $null
    $stack = New-Object System.Collections.Stack
    $stack.Push($DaysList)
    while ($stack.Count -gt 0) {
        $d = $stack.Pop()
        if ($d -is [System.Windows.Controls.ScrollViewer]) { $sv = $d; break }
        $n = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($d)
        for ($i = 0; $i -lt $n; $i++) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($d, $i)) }
    }
    if ($sv) {
        switch ($e.Key) {
            "Up"   { $sv.LineUp();   $e.Handled = $true }
            "Down" { $sv.LineDown(); $e.Handled = $true }
        }
    }
}.GetNewClosure())

# ---------- Launch ----------
$CreateJourneyBtn.Visibility = "Visible"
Update-ButtonVisibility
$liveJsonPath = Join-Path (Get-LiveJourneyDir) "live.json"
if (Test-Path $liveJsonPath) {
    Update-SyncStatus -Timestamp (Get-Item $liveJsonPath).LastWriteTime
}
Set-SyncPending $false
# If a focus session was left in progress last run, resolve it through the modal
# recovery dialog BEFORE the window opens. The user can't begin another session
# until the previous one is resolved (Resume / End Session / Complete / Discard).
$recovered = Resolve-FocusRecovery
# Root-cause fix (previous approach built the Days list here, before
# ShowDialog(), then tried to patch around the resulting bad first
# measurement with -Force rebuilds, InvalidateMeasure, and a resize nudge -
# all of which were repairs, not prevention). The actual bug: building rows
# before the window has ever been shown means DaysList has ~0 actual width,
# so the header Grid's Star column and the chip WrapPanel collapse when
# rows are first measured - producing the squashed/overlapping first paint.
# Building the list AFTER the window has completed its first real layout
# pass (ContentRendered) means rows are correctly sized the very first time
# they're measured, so nothing needs correcting afterwards.
$window.Add_ContentRendered({
    # Build the initial Days list after recovery resolves: completing/ending
    # a session can auto-create today's entry (see Write-FocusSessionRecord),
    # and the list must include it. On Resume, Restore-FocusSessionIntoApp
    # below overrides the view to Focus, so this doesn't fight the jump to
    # the Focus panel.
    Set-Filter "All"
    if ($recovered) {
        Restore-FocusSessionIntoApp -State $recovered
    }
    # Safety net for a separate, narrower WPF race: ItemContainerGenerator's
    # StatusChanged can fire before VirtualizingStackPanel has actually
    # realized child containers into its Children collection, so
    # Populate-RealizedRows can find nothing on that particular pass. This
    # doesn't need -Force - Set-Filter above only ever runs after the window
    # is genuinely laid out, so any row it builds is already correct; this
    # is purely a "catch anything not yet realized" pass, not a repair pass.
    $DaysList.Dispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Loaded,
        [action]{ Populate-RealizedRows }
    ) | Out-Null
}.GetNewClosure())
$window.ShowDialog() | Out-Null
