![alt text](Shadow.png)# Productivity Tracker

A PowerShell + WPF desktop app for tracking daily goal-hours across one or more "journeys."

## Running it

```powershell
.\Main.ps1
```

Requires the `ImportExcel` module for XLSX export/backup:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

## Files

- `Main.ps1` — entry point, UI wiring
- `DataStore.ps1` — load/save entries & settings, journey archiving, restore
- `Exporter.ps1` — CSV / Markdown / XLSX export
- `MainWindow.xaml` — UI layout
- `Build.ps1` — syntax-checks the scripts and packages a distributable `ProductivityTracker.zip`
- `Data/` — `entries.json` and `settings.json` for the current ("live") journey
- `Backup/` — auto-sync backups; `SYNC` also writes a json+xlsx pair per journey (live + archived) into `Backup/Journeys/`
- `Data/Archive/` — created automatically the first time you archive a journey (via Create New Journey or Switch Journey)

## Concepts

- A **journey** has a name, a start date, and a goal-hour template (e.g. 4h / 6h / 9h). Each day added gets its own copy of the current template's goals, which you mark done/missed.
- **Create New Journey** archives whatever journey is currently active (if any) and opens setup for a new one.
- **Switch Journey** archives the current journey and loads a previously archived one back as the active, editable journey.
- **Load from Archive** opens an old journey read-only, without touching the live one. **Back to Current** returns to the live journey.
- **Edit Template** changes the goal template going forward (and lets you rename the journey) without altering days already logged.
