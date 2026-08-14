# DataStore.ps1
# Handles loading, saving, and mutating productivity entries stored in Data/entries.json

$script:DataDir  = Join-Path $PSScriptRoot "Data"
$script:DataFile = Join-Path $script:DataDir "entries.json"

function global:Initialize-DataStore {
    if (-not (Test-Path $script:DataDir)) {
        New-Item -ItemType Directory -Path $script:DataDir | Out-Null
    }
    if (-not (Test-Path $script:DataFile)) {
        "[]" | Set-Content -Path $script:DataFile -Encoding UTF8
    }
}

function global:Import-Entries {
    Initialize-DataStore
    $raw = Get-Content -Path $script:DataFile -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($raw)) { return , @() }

    $data = $raw | ConvertFrom-Json
    if ($null -eq $data) { return , @() }

    # ConvertFrom-Json collapses a single-element array to a scalar object - normalize back to an array
    if ($data -isnot [System.Array]) { $data = @($data) }

    # Backfill Sessions for legacy entries created before Focus Mode existed.
    foreach ($entry in $data) {
        if (-not $entry.PSObject.Properties['Sessions'] -or $null -eq $entry.Sessions) {
            $entry | Add-Member -NotePropertyName Sessions -NotePropertyValue @() -Force
        }
    }
    return , $data
}

function global:Save-Entries {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Entries
    )
    Initialize-DataStore
    # Wrap in @() to prevent PowerShell pipeline enumeration from unwrapping a
    # single-element array into a bare object — without this, ConvertTo-Json
    # produces { ... } (object) instead of [ ... ] (array) when there is exactly
    # one entry, which can confuse downstream readers and the virtualized list.
    @($Entries) | ConvertTo-Json -Depth 4 | Set-Content -Path $script:DataFile -Encoding UTF8
}

function global:New-Entry {
    param(
        [Parameter(Mandatory)][string]$Date,
        [Parameter(Mandatory)][AllowEmptyCollection()][int[]]$GoalHours,
        [array]$ExistingEntries = (Import-Entries)
    )

    $dayNumber = ($ExistingEntries | Measure-Object).Count + 1
    # Wrapped in @(...) so an empty goal template yields an empty array, never a
    # single $null element (@($null) would inject a phantom goal into the entry).
    $goals = @(foreach ($h in ($GoalHours | Sort-Object)) {
        [PSCustomObject]@{ Hours = $h; Status = "missed" }
    })

    [PSCustomObject]@{
        Id        = [guid]::NewGuid().ToString()
        Date      = $Date
        DayNumber = $dayNumber
        Goals     = @($goals)
        Note      = ""
        Sessions  = @()
    }
}

function global:Set-GoalStatus {
    <#
        Toggles a goal's status between missed (red) and done (green).
        New goals start as "missed" by default - user marks green when completed.
    #>
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][int]$Hours
    )

    foreach ($entry in $Entries) {
        if ($entry.Id -eq $Id) {
            foreach ($goal in $entry.Goals) {
                if ($goal.Hours -eq $Hours) {
                    $goal.Status = if ($goal.Status -eq "done") { "missed" } else { "done" }
                }
            }
        }
    }
    return , $Entries
}

function global:Set-EntryNote {
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Note
    )
    foreach ($entry in $Entries) {
        if ($entry.Id -eq $Id) { $entry.Note = $Note }
    }
    return , $Entries
}

function global:Add-EntrySession {
    <#
        Appends a focus session to the entry with the given Id and returns the
        entries. Session shape: @{ type; label; start; end; duration } where
        start/end are wall-clock HH:MM strings and duration is in minutes.
    #>
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)]$Session
    )
    foreach ($entry in $Entries) {
        if ($entry.Id -eq $Id) {
            if ($null -eq $entry.Sessions -or -not $entry.PSObject.Properties['Sessions']) {
                $entry | Add-Member -NotePropertyName Sessions -NotePropertyValue @() -Force
            }
            $entry.Sessions = @($entry.Sessions) + $Session
        }
    }
    return , $Entries
}

function global:Remove-Entry {
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Id
    )
    return , @($Entries | Where-Object { $_.Id -ne $Id })
}

function global:Remove-EntrySession {
    <#
        Removes the session at the given zero-based index from the entry with
        the given Id and returns the entries.
    #>
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][int]$SessionIndex
    )
    foreach ($entry in $Entries) {
        if ($entry.Id -eq $Id) {
            if (-not $entry.PSObject.Properties['Sessions'] -or $null -eq $entry.Sessions) { break }
            $sessions = @($entry.Sessions)
            if ($SessionIndex -ge 0 -and $SessionIndex -lt $sessions.Count) {
                $newSessions = @()
                for ($i = 0; $i -lt $sessions.Count; $i++) {
                    if ($i -ne $SessionIndex) { $newSessions += $sessions[$i] }
                }
                $entry.Sessions = $newSessions
            }
        }
    }
    return , $Entries
}

function global:Set-EntrySessionLabel {
    <#
        Updates the label of the session at the given zero-based index on the
        entry with the given Id and returns the entries.
    #>
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][int]$SessionIndex,
        [Parameter(Mandatory)][string]$NewLabel
    )
    foreach ($entry in $Entries) {
        if ($entry.Id -eq $Id) {
            if (-not $entry.PSObject.Properties['Sessions'] -or $null -eq $entry.Sessions) { break }
            $sessions = @($entry.Sessions)
            if ($SessionIndex -ge 0 -and $SessionIndex -lt $sessions.Count) {
                $sessions[$SessionIndex].label = $NewLabel
                $entry.Sessions = $sessions
            }
        }
    }
    return , $Entries
}

function global:Set-EntrySessionType {
    <#
        Updates the type of the session at the given zero-based index on the
        entry with the given Id and returns the entries.
    #>
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][int]$SessionIndex,
        [Parameter(Mandatory)][string]$NewType
    )
    foreach ($entry in $Entries) {
        if ($entry.Id -eq $Id) {
            if (-not $entry.PSObject.Properties['Sessions'] -or $null -eq $entry.Sessions) { break }
            $sessions = @($entry.Sessions)
            if ($SessionIndex -ge 0 -and $SessionIndex -lt $sessions.Count) {
                $sessions[$SessionIndex].type = $NewType
                $entry.Sessions = $sessions
            }
        }
    }
    return , $Entries
}

# ---------- Settings (journey start date) ----------
$script:SettingsFile = Join-Path $script:DataDir "settings.json"

function global:Initialize-Settings {
    if (-not (Test-Path $script:DataDir)) {
        New-Item -ItemType Directory -Path $script:DataDir | Out-Null
    }
    if (-not (Test-Path $script:SettingsFile)) {
        '{"StartDate": null, "GoalHours": [], "JourneyName": null, "BackupFolder": null}' | Set-Content -Path $script:SettingsFile -Encoding UTF8
    }
}

function global:Import-Settings {
    Initialize-Settings
    $raw = Get-Content -Path $script:SettingsFile -Raw -ErrorAction SilentlyContinue
    $default = [PSCustomObject]@{ StartDate = $null; GoalHours = @(); JourneyName = $null; BackupFolder = $null }
    if ([string]::IsNullOrWhiteSpace($raw)) { return $default }
    $data = $raw | ConvertFrom-Json
    if ($null -eq $data) { return $default }
    if (-not (Get-Member -InputObject $data -Name "GoalHours" -ErrorAction SilentlyContinue)) {
        $data | Add-Member -MemberType NoteProperty -Name "GoalHours" -Value @()
    }
    if ($null -eq $data.GoalHours) { $data.GoalHours = @() }
    if (-not (Get-Member -InputObject $data -Name "JourneyName" -ErrorAction SilentlyContinue)) {
        $data | Add-Member -MemberType NoteProperty -Name "JourneyName" -Value $null
    }
    if (-not (Get-Member -InputObject $data -Name "BackupFolder" -ErrorAction SilentlyContinue)) {
        $data | Add-Member -MemberType NoteProperty -Name "BackupFolder" -Value $null
    }
    return $data
}

function global:Save-Settings {
    param([Parameter(Mandatory)]$Settings)
    Initialize-Settings
    $Settings | ConvertTo-Json -Depth 5 | Set-Content -Path $script:SettingsFile -Encoding UTF8
}

# ---------- Journey archiving (groundwork for future multi-journey support) ----------
$script:ArchiveDir = Join-Path $script:DataDir "Archive"

function global:Save-JourneyArchive {
    <#
        Copies the given entries + settings into a timestamped folder under Data/Archive/
        so nothing is lost when a journey is reset. Returns the archive folder path.
    #>
    param(
        # Not Mandatory so a journey with zero entries can still be archived
        # (switching away from a fresh journey passes an empty collection, and a
        # Mandatory array rejects that at bind time).
        [Parameter(Mandatory = $false)][array]$Entries = @(),
        [Parameter(Mandatory)]$Settings
    )
    if (-not (Test-Path $script:ArchiveDir)) {
        New-Item -ItemType Directory -Path $script:ArchiveDir | Out-Null
    }

    # Never archive an uninitialized journey — it produces a junk "untitled" entry
    # that can never be deduped (no JourneyName to match on) and clutters the list.
    if (-not $Settings.StartDate -and -not $Settings.JourneyName) {
        return $null
    }

    $rawLabel = if ($Settings.JourneyName) { $Settings.JourneyName } elseif ($Settings.StartDate) { $Settings.StartDate } else { "untitled" }
    $safeLabel = ($rawLabel -replace '[^a-zA-Z0-9\-]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($safeLabel)) { $safeLabel = "untitled" }

    # Data\Archive keeps only the latest snapshot per named journey - remove any
    # earlier archive folder(s) for this same journey before writing the new one.
    # (Journeys with no name can't be matched reliably, so those are left alone
    # and will keep accumulating - same as before.)
    if ($Settings.JourneyName) {
        Get-ChildItem -Path $script:ArchiveDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "journey-$safeLabel-*" } |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $folderPath = Join-Path $script:ArchiveDir "journey-$safeLabel-$stamp"
    New-Item -ItemType Directory -Path $folderPath | Out-Null

    $Entries  | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $folderPath "entries.json")  -Encoding UTF8
    $Settings | ConvertTo-Json -Depth 5  | Set-Content -Path (Join-Path $folderPath "settings.json") -Encoding UTF8

    return $folderPath
}

function global:Get-ArchivedJourneys {
    <#
        Scans Data/Archive/ and returns one summary object per archived journey,
        newest first: Path, FolderName, Name, StartDate, EntryCount.
    #>
    if (-not (Test-Path $script:ArchiveDir)) { return @() }

    $folders = Get-ChildItem -Path $script:ArchiveDir -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending

    $results = foreach ($folder in $folders) {
        $settingsPath = Join-Path $folder.FullName "settings.json"
        $entriesPath  = Join-Path $folder.FullName "entries.json"
        if (-not (Test-Path $settingsPath)) { continue }

        $settingsRaw = Get-Content -Path $settingsPath -Raw -ErrorAction SilentlyContinue
        $settingsObj = if ($settingsRaw) { $settingsRaw | ConvertFrom-Json } else { $null }

        $entryCount = 0
        if (Test-Path $entriesPath) {
            $entriesRaw = Get-Content -Path $entriesPath -Raw -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrWhiteSpace($entriesRaw)) {
                $entriesObj = $entriesRaw | ConvertFrom-Json
                $entryCount = if ($entriesObj -is [System.Array]) { $entriesObj.Count } elseif ($null -ne $entriesObj) { 1 } else { 0 }
            }
        }

        [PSCustomObject]@{
            Path       = $folder.FullName
            FolderName = $folder.Name
            Name       = if ($settingsObj -and $settingsObj.JourneyName) { $settingsObj.JourneyName } else { "(untitled journey)" }
            StartDate  = if ($settingsObj) { $settingsObj.StartDate } else { $null }
            EntryCount = $entryCount
            ArchivedAt = $folder.LastWriteTime
        }
    }

    return @($results)
}

function global:Get-LatestArchivedJourneys {
    <#
        Same as Get-ArchivedJourneys, but collapsed to one entry per named
        journey - the newest snapshot only. Switching back and forth between
        journeys archives a fresh snapshot every time you leave one, so the
        raw list can contain several snapshots with the same name; pickers
        should only offer the latest one. Journeys still using the default
        "(untitled journey)" label (pre-dating the naming feature) are kept
        distinct per folder, since collapsing those would hide unrelated
        archives that just happen to share the placeholder name.
    #>
    $all = Get-ArchivedJourneys
    $groups = $all | Group-Object { if ($_.Name -eq "(untitled journey)") { $_.FolderName } else { $_.Name } }
    return @($groups | ForEach-Object { $_.Group | Select-Object -First 1 })
}

function global:Restore-JourneyFromArchive {
    <#
        Reads entries.json + settings.json out of an archive folder and returns
        them as @{ Entries = ...; Settings = ... }. Does not touch the live Data
        files or the archive itself - caller decides what to do with the result.
    #>
    param([Parameter(Mandatory)][string]$ArchivePath)

    $entriesPath  = Join-Path $ArchivePath "entries.json"
    $settingsPath = Join-Path $ArchivePath "settings.json"

    $entries = @()
    if (Test-Path $entriesPath) {
        $raw = Get-Content -Path $entriesPath -Raw -ErrorAction SilentlyContinue
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $data = $raw | ConvertFrom-Json
            if ($null -ne $data) { $entries = if ($data -isnot [System.Array]) { @($data) } else { $data } }
        }
    }

    $settings = [PSCustomObject]@{ StartDate = $null; GoalHours = @(); JourneyName = $null; BackupFolder = $null }
    if (Test-Path $settingsPath) {
        $raw = Get-Content -Path $settingsPath -Raw -ErrorAction SilentlyContinue
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $data = $raw | ConvertFrom-Json
            if ($null -ne $data) {
                if (-not (Get-Member -InputObject $data -Name "JourneyName" -ErrorAction SilentlyContinue)) {
                    $data | Add-Member -MemberType NoteProperty -Name "JourneyName" -Value $null
                }
                $settings = $data
            }
        }
    }

    return @{ Entries = @($entries); Settings = $settings }
}

function global:Restore-JourneysFromBackup {
    <#
        Restores every journey found under a backup folder's Journeys\ subfolder
        into Data\Archive so they show up in LOAD ARCHIVE / SWITCH JOURNEY.

        Each journey folder's newest checkpoint snapshot is used for its entries;
        live-settings.json (or the folder name) provides its settings. Never
        overwrites an existing archive - every restore gets a fresh timestamped
        folder. Returns @{ Restored = <journey summaries>; Live = <latest live
        journey or $null> }.
    #>
    param([Parameter(Mandatory)][string]$BackupRoot)

    $journeysDir = Join-Path $BackupRoot "Journeys"
    if (-not (Test-Path $journeysDir)) {
        throw "No Journeys folder found in: $BackupRoot"
    }

    if (-not (Test-Path $script:ArchiveDir)) {
        New-Item -ItemType Directory -Path $script:ArchiveDir -Force | Out-Null
    }

    $restored = @()
    $live     = $null

    foreach ($folder in (Get-ChildItem -Path $journeysDir -Directory -ErrorAction SilentlyContinue)) {
        $settingsPath = Join-Path $folder.FullName "live-settings.json"
        $hasLiveSettings = Test-Path $settingsPath

        # --- Settings: live-settings.json when present, else defaults from the folder name ---
        $settings = [PSCustomObject]@{ StartDate = $null; GoalHours = @(); JourneyName = $null; BackupFolder = $null }
        if ($hasLiveSettings) {
            $raw = Get-Content -Path $settingsPath -Raw -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $data = $raw | ConvertFrom-Json
                if ($null -ne $data) {
                    foreach ($p in @("GoalHours","JourneyName","BackupFolder","StartDate")) {
                        if (-not (Get-Member -InputObject $data -Name $p -ErrorAction SilentlyContinue)) {
                            $data | Add-Member -MemberType NoteProperty -Name $p -Value $null
                        }
                    }
                    if ($null -eq $data.GoalHours) { $data.GoalHours = @() }
                    $settings = $data
                }
            }
        }
        if (-not $settings.JourneyName) { $settings.JourneyName = $folder.Name }

        # --- Entries: newest ranked checkpoint, else live.json ---
        $entries = @()
        $snap = Get-ChildItem -Path $folder.FullName -Filter "*.json" -ErrorAction SilentlyContinue |
                Where-Object { $_.BaseName -match '^\d{2} - ' } |
                Sort-Object Name -Descending | Select-Object -First 1
        $entriesSrc = if ($snap) { $snap.FullName } else { Join-Path $folder.FullName "live.json" }
        if (Test-Path $entriesSrc) {
            $raw = Get-Content -Path $entriesSrc -Raw -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $data = $raw | ConvertFrom-Json
                if ($null -ne $data) {
                    $entries = if ($data -isnot [System.Array]) { @($data) } else { $data }
                }
            }
        }

        # --- Write a fresh archive folder (never wipes other archives) ---
        $safeName = ($settings.JourneyName -replace '[^a-zA-Z0-9\-]+', '-').Trim('-')
        if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "untitled" }
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $archivePath = Join-Path $script:ArchiveDir "journey-$safeName-$stamp"
        New-Item -ItemType Directory -Path $archivePath -Force | Out-Null
        $entries  | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $archivePath "entries.json")  -Encoding UTF8
        $settings | ConvertTo-Json -Depth 5  | Set-Content -Path (Join-Path $archivePath "settings.json") -Encoding UTF8

        $restored += [PSCustomObject]@{ Name = $settings.JourneyName; Path = $archivePath; EntryCount = $entries.Count }

        # The live journey (the one with live-settings.json) becomes the candidate
        # for restoring the app's current state.
        if ($hasLiveSettings) { $live = @{ Entries = $entries; Settings = $settings } }
    }

    return @{ Restored = @($restored); Live = $live }
}

function global:Get-UniqueArchivePath {
    <#
        Resolves a folder path under Data\Archive for the given safe journey name
        that does not already exist. Appends the current timestamp, then a numeric
        suffix if two imports land in the same second, so an import can never
        overwrite an existing archive.
    #>
    param([string]$SafeName)
    if (-not (Test-Path $script:ArchiveDir)) {
        New-Item -ItemType Directory -Path $script:ArchiveDir -Force | Out-Null
    }
    $base = Join-Path $script:ArchiveDir "journey-$SafeName-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    $candidate = $base
    $i = 1
    while (Test-Path $candidate) {
        $candidate = "$base-$i"
        $i++
    }
    return $candidate
}

function global:Restore-JourneyFromExport {
    <#
        Mode 2 import: reads a single journey from a folder that directly contains
        entries.json + settings.json (a standalone journey export, not a full
        backup tree). Writes it as a fresh archive under Data\Archive so it shows
        up in LOAD ARCHIVE / SWITCH JOURNEY. Never overwrites existing archives.

        Returns the same @{ Restored; Live } shape as Restore-JourneysFromBackup,
        with Live set to the imported journey so a fresh app can activate it.
    #>
    param([Parameter(Mandatory)][string]$ExportRoot)

    $entriesPath  = Join-Path $ExportRoot "entries.json"
    $settingsPath = Join-Path $ExportRoot "settings.json"

    if (-not (Test-Path $entriesPath) -or -not (Test-Path $settingsPath)) {
        throw "The selected folder is not a valid Productivity Tracker backup or journey export."
    }

    # --- Entries ---
    $entries = @()
    try {
        $raw = Get-Content -Path $entriesPath -Raw -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $data = $raw | ConvertFrom-Json
            if ($null -ne $data) {
                $entries = if ($data -isnot [System.Array]) { @($data) } else { $data }
            }
        }
    } catch {
        throw "Could not read entries.json in the selected folder (invalid or unreadable JSON)."
    }

    # --- Settings ---
    $settings = [PSCustomObject]@{ StartDate = $null; GoalHours = @(); JourneyName = $null; BackupFolder = $null }
    try {
        $raw = Get-Content -Path $settingsPath -Raw -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $data = $raw | ConvertFrom-Json
            if ($null -ne $data) {
                foreach ($p in @("StartDate","GoalHours","JourneyName","BackupFolder")) {
                    if (-not (Get-Member -InputObject $data -Name $p -ErrorAction SilentlyContinue)) {
                        $data | Add-Member -MemberType NoteProperty -Name $p -Value $null
                    }
                }
                if ($null -eq $data.GoalHours) { $data.GoalHours = @() }
                $settings = $data
            }
        }
    } catch {
        throw "Could not read settings.json in the selected folder (invalid or unreadable JSON)."
    }
    if (-not $settings.JourneyName) { $settings.JourneyName = "untitled" }

    # --- Write a fresh, unique archive folder ---
    $safeName = ($settings.JourneyName -replace '[^a-zA-Z0-9\-]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "untitled" }
    $archivePath = Get-UniqueArchivePath -SafeName $safeName
    New-Item -ItemType Directory -Path $archivePath -Force | Out-Null
    $entries  | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $archivePath "entries.json")  -Encoding UTF8
    $settings | ConvertTo-Json -Depth 5  | Set-Content -Path (Join-Path $archivePath "settings.json") -Encoding UTF8

    $summary = [PSCustomObject]@{ Name = $settings.JourneyName; Path = $archivePath; EntryCount = $entries.Count }
    return @{
        Restored = @($summary)
        Live     = @{ Entries = $entries; Settings = $settings }
    }
}

function global:Import-JourneysFromFolder {
    <#
        Dispatch import by inspecting the selected folder and choosing the mode:
          - Contains a Journeys\ subfolder            -> full backup tree (Mode 1)
          - Contains entries.json + settings.json     -> single journey export (Mode 2)
          - Otherwise                                 -> friendly error.
        Returns @{ Restored; Live; Source } where Source is "backup" | "export".
    #>
    param([Parameter(Mandatory)][string]$SourceRoot)

    if (Test-Path (Join-Path $SourceRoot "Journeys")) {
        $r = Restore-JourneysFromBackup -BackupRoot $SourceRoot
        $r.Source = "backup"
        return $r
    }

    $entriesPath  = Join-Path $SourceRoot "entries.json"
    $settingsPath = Join-Path $SourceRoot "settings.json"
    if ((Test-Path $entriesPath) -and (Test-Path $settingsPath)) {
        $r = Restore-JourneyFromExport -ExportRoot $SourceRoot
        $r.Source = "export"
        return $r
    }

    throw "The selected folder is not a valid Productivity Tracker backup or journey export."
}

function global:Rename-ArchivedJourney {
    <#
        Renames an archived journey: updates its settings.json and renames the
        folder to match the new safe label. Returns the new folder path.
    #>
    param(
        [Parameter(Mandatory)][string]$ArchivePath,
        [Parameter(Mandatory)][string]$NewName
    )

    $newName = $NewName.Trim()
    if ([string]::IsNullOrWhiteSpace($newName)) {
        throw "Journey name cannot be empty."
    }

    # Reject duplicate names (compare against all known archived journeys).
    $existing = Get-LatestArchivedJourneys
    foreach ($e in $existing) {
        if ($e.Name -eq $newName -and $e.Path -ne $ArchivePath) {
            throw "A journey with the name `"$newName`" already exists."
        }
    }

    # Update settings.json inside the archive folder.
    $settingsPath = Join-Path $ArchivePath "settings.json"
    if (-not (Test-Path $settingsPath)) {
        throw "Archive folder does not contain settings.json: $ArchivePath"
    }
    $raw = Get-Content -Path $settingsPath -Raw -ErrorAction SilentlyContinue
    $data = $raw | ConvertFrom-Json
    $data.JourneyName = $newName
    $data | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsPath -Encoding UTF8

    # Rename the folder to match the new safe label, preserving the timestamp.
    $folderName = Split-Path $ArchivePath -Leaf
    $newSafeLabel = ($newName -replace '[^a-zA-Z0-9\-]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($newSafeLabel)) { $newSafeLabel = "untitled" }
    # Extract the existing timestamp from the folder name (journey-<old>-<stamp>).
    $match = [regex]::Match($folderName, 'journey-[a-zA-Z0-9\-]+-(\d{8}-\d{6})')
    $stamp = if ($match.Success) { $match.Groups[1].Value } else { Get-Date -Format "yyyyMMdd-HHmmss" }
    $newFolderName = "journey-$newSafeLabel-$stamp"
    $parentDir = Split-Path $ArchivePath -Parent
    $newPath = Join-Path $parentDir $newFolderName

    if ($ArchivePath -ne $newPath) {
        Rename-Item -Path $ArchivePath -NewName $newFolderName -ErrorAction Stop
    }

    return $newPath
}

function global:Delete-ArchivedJourney {
    <#
        Completely removes an archived journey folder and all its contents.
    #>
    param([Parameter(Mandatory)][string]$ArchivePath)

    # Safety: only allow deleting folders under the archive directory.
    $archiveDir = Join-Path (Join-Path $PSScriptRoot "Data") "Archive"
    $resolved = (Resolve-Path $ArchivePath -ErrorAction SilentlyContinue).Path
    if (-not $resolved -or -not $resolved.StartsWith($archiveDir, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Invalid archive path: $ArchivePath"
    }

    Remove-Item -Path $resolved -Recurse -Force
    return $true
}

function global:Renumber-Entries {
    <#
        Reassigns DayNumber to every entry sequentially (1, 2, 3…)
        based on ascending date order. Returns the renumbered array.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Entries)

    $sorted = @($Entries | Sort-Object { [datetime]$_.Date })
    $dayNumber = 1
    foreach ($entry in $sorted) {
        $entry.DayNumber = $dayNumber
        $dayNumber++
    }
    return , $sorted
}
