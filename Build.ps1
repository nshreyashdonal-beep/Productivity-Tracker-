# Build.ps1
# Syntax-checks the app's PowerShell scripts, compiles Main.ps1 into
# ProductivityTracker.exe via the PS2EXE module, stages everything the exe
# reads from disk at runtime (loose scripts, XAML, Assets), and packages it
# all into a distributable ProductivityTracker.zip.
#
# Usage:
#   .\Build.ps1
#   .\Build.ps1 -OutputDir .\dist
#   .\Build.ps1 -SkipZip

param(
    [string]$OutputDir = (Join-Path $PSScriptRoot "dist"),
    [switch]$SkipZip
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# A fresh machine's Restricted/AllSigned execution policy would block this
# build script itself, and Install-Module/Import-Module for PS2EXE below.
# Bypass is scoped to THIS PROCESS ONLY - it never touches the machine-wide
# or CurrentUser policy, and reverts automatically when the process exits.
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Write-Host "Execution policy for this build (process scope): Bypass" -ForegroundColor DarkGray

$root     = $PSScriptRoot
$exeName  = "ProductivityTracker.exe"
$exePath  = Join-Path $root $exeName
$iconPath = Join-Path $root "Assets\app.ico"

# Files/folders the compiled exe reads from disk at runtime. PS2EXE compiles
# only the entry-point script; the dot-sourced scripts and loose XAML that
# Main.ps1 loads via its $root fallback (see the comment at the top of
# Main.ps1) are NOT embedded and must ship alongside the exe.
$runtimeFiles = @("DataStore.ps1", "Exporter.ps1", "MainWindow.xaml", "RecoveryWindow.xaml")
$runtimeDirs  = @("Assets")

# ---------------------------------------------------------------------------
# 1. Syntax check every script that ships with the app. Parses each file
#    with the PowerShell language parser (no execution) and fails the build
#    on the first error found.
Write-Host "`n== Syntax check ==" -ForegroundColor Cyan
$scriptsToCheck = @("Main.ps1", "DataStore.ps1", "Exporter.ps1") | ForEach-Object { Join-Path $root $_ }
$hadErrors = $false
foreach ($file in $scriptsToCheck) {
    if (-not (Test-Path $file)) {
        Write-Host "  SKIP (not found): $file" -ForegroundColor Yellow
        continue
    }
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        $hadErrors = $true
        Write-Host "  FAIL: $(Split-Path $file -Leaf)" -ForegroundColor Red
        foreach ($e in $parseErrors) {
            Write-Host "    Line $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "  OK:   $(Split-Path $file -Leaf)" -ForegroundColor Green
    }
}
if ($hadErrors) {
    throw "Syntax check failed - fix the errors above before building."
}

# ---------------------------------------------------------------------------
# 2. Make sure PS2EXE is available. Installs to CurrentUser scope if missing;
#    the process-scoped Bypass above ensures the install/import isn't
#    blocked by the machine's policy either.
Write-Host "`n== PS2EXE module ==" -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "  Installing ps2exe (CurrentUser scope)..." -ForegroundColor DarkGray
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe -Force

# ---------------------------------------------------------------------------
# 3. Compile Main.ps1 -> ProductivityTracker.exe
Write-Host "`n== Compiling $exeName ==" -ForegroundColor Cyan
$ps2exeParams = @{
    inputFile   = Join-Path $root "Main.ps1"
    outputFile  = $exePath
    noConsole   = $true
    title       = "Productivity Tracker"
    description = "Daily goal-hours & focus session tracker"
    product     = "Productivity Tracker"
    version     = "1.0.0.0"
}
if (Test-Path $iconPath) { $ps2exeParams["iconFile"] = $iconPath }
Invoke-ps2exe @ps2exeParams

if (-not (Test-Path $exePath)) {
    throw "PS2EXE did not produce $exeName - check the output above."
}
Write-Host "  Built: $exePath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. Stage the exe plus everything it reads from disk at runtime into
#    $OutputDir. A fresh, empty Data\ folder is included so the packaged app
#    starts blank - the developer's own journey data in .\Data is never
#    shipped.
Write-Host "`n== Staging distributable ==" -ForegroundColor Cyan
if (Test-Path $OutputDir) { Remove-Item $OutputDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutputDir | Out-Null

Copy-Item $exePath -Destination $OutputDir
foreach ($f in $runtimeFiles) {
    $src = Join-Path $root $f
    if (Test-Path $src) {
        Copy-Item $src -Destination $OutputDir
    } else {
        Write-Host "  Warning: expected runtime file missing: $f" -ForegroundColor Yellow
    }
}
foreach ($d in $runtimeDirs) {
    $src = Join-Path $root $d
    if (Test-Path $src) { Copy-Item $src -Destination (Join-Path $OutputDir $d) -Recurse }
}
New-Item -ItemType Directory -Path (Join-Path $OutputDir "Data") -Force | Out-Null
"[]" | Set-Content -Path (Join-Path $OutputDir "Data\entries.json") -Encoding UTF8

Write-Host "  Staged in: $OutputDir" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 5. Zip it.
if (-not $SkipZip) {
    $zipPath = Join-Path $root "ProductivityTracker.zip"
    Write-Host "`n== Packaging $(Split-Path $zipPath -Leaf) ==" -ForegroundColor Cyan
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path (Join-Path $OutputDir "*") -DestinationPath $zipPath
    Write-Host "  Packaged: $zipPath" -ForegroundColor Green
}

Write-Host "`nBuild complete." -ForegroundColor Cyan
