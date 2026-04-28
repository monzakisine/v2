<#
.SYNOPSIS
    One-time setup: creates folder structure, generates company
    subfolders from config, and adds the AMC root to the user PATH
    so the 'amc' command is callable from any directory.

.DESCRIPTION
    Run this ONCE on the company Windows PC, after copying the
    AMC-Automation folder to C:\AMC-Automation (or wherever you
    set RootDir in config.psd1).

.EXAMPLE
    cd C:\AMC-Automation\Scripts
    powershell -ExecutionPolicy Bypass -File .\Install.ps1
#>

[CmdletBinding()]
param(
    [string] $ConfigPath
)

$ErrorActionPreference = 'Stop'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir 'config.psd1' }

if (-not (Test-Path $ConfigPath)) { throw "Config not found: $ConfigPath" }
$Cfg = Import-PowerShellDataFile $ConfigPath

$RootDir      = $Cfg.RootDir
$TrackerDir   = Join-Path $RootDir (Split-Path $Cfg.TrackerRelPath -Parent)
$CompaniesDir = Join-Path $RootDir $Cfg.CompaniesDir
$ArchiveDir   = Join-Path $RootDir $Cfg.ArchiveDir
$LogsDir      = Join-Path $RootDir $Cfg.LogsDir

Write-Host '=== AMC Automation: Install ===' -ForegroundColor Cyan
Write-Host "Root: $RootDir"
Write-Host ''

# 1. Create main folders
Write-Host 'Creating folder structure...' -ForegroundColor Yellow
foreach ($d in @($RootDir, $TrackerDir, $CompaniesDir, $ArchiveDir, $LogsDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
        Write-Host "  + $d" -ForegroundColor Green
    } else {
        Write-Host "  = $d (exists)" -ForegroundColor DarkGray
    }
}

# 2. Create one folder per configured company
Write-Host ''
Write-Host 'Creating company subfolders...' -ForegroundColor Yellow
foreach ($key in ($Cfg.Companies.Keys | Sort-Object)) {
    $folder = Join-Path $CompaniesDir $Cfg.Companies[$key].Folder
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Write-Host "  + $folder" -ForegroundColor Green
    } else {
        Write-Host "  = $folder (exists)" -ForegroundColor DarkGray
    }
}

# 3. Add RootDir to USER PATH (so amc.bat works from anywhere)
Write-Host ''
Write-Host 'Updating user PATH...' -ForegroundColor Yellow
$currentPath = [Environment]::GetEnvironmentVariable('Path', 'User')
if (-not $currentPath) { $currentPath = '' }
$pathParts = $currentPath -split ';' | Where-Object { $_ -ne '' }

if ($pathParts -contains $RootDir) {
    Write-Host "  = '$RootDir' already in user PATH" -ForegroundColor DarkGray
} else {
    $newPath = if ($pathParts.Count -eq 0) { $RootDir } else { ($pathParts -join ';') + ';' + $RootDir }
    [Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
    Write-Host "  + Added '$RootDir' to user PATH" -ForegroundColor Green
    Write-Host '  ! Open a NEW terminal window to use the amc command.' -ForegroundColor Yellow
}

# 4. Sanity-check tracker presence
$trackerPath = Join-Path $RootDir $Cfg.TrackerRelPath
Write-Host ''
if (Test-Path $trackerPath) {
    Write-Host "Tracker found: $trackerPath" -ForegroundColor Green
} else {
    Write-Host "Tracker NOT found at: $trackerPath" -ForegroundColor Red
    Write-Host '   Place Contractors_AMC_Tracker_2026.xlsm in the Tracker folder before running.' -ForegroundColor Yellow
}

# 5. Sanity-check amc.bat
$batPath = Join-Path $RootDir 'amc.bat'
if (Test-Path $batPath) {
    Write-Host "Launcher found: $batPath" -ForegroundColor Green
} else {
    Write-Host "Launcher NOT found at: $batPath" -ForegroundColor Red
    Write-Host '   Make sure amc.bat sits in the root AMC-Automation folder.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host '=== Install complete ===' -ForegroundColor Cyan
Write-Host ''
Write-Host 'Try it out (in a NEW terminal):'  -ForegroundColor White
Write-Host '   amc all -DryRun'                -ForegroundColor White
Write-Host ''
