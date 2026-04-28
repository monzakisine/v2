<#
.SYNOPSIS
    Friendly menu UI launched when amc.bat is double-clicked.
    Shows the list of companies, lets the user pick one (or "all"),
    runs Update-Tracker.ps1, then waits for a keypress before closing.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'


$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Network-share fixes: unblock all .ps1 files + force Bypass for this process
# so sub-scripts (Update-Tracker.ps1) don't trigger the "Run only scripts that
# you trust" prompt every time.
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue } catch { }
try {
    Get-ChildItem -Path $ScriptDir -Filter '*.ps1' -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
} catch { }


# Window appearance
try {
    $Host.UI.RawUI.WindowTitle = 'AMC Tracker Automation'
    # Try to make the window a bit larger if possible
    $size = $Host.UI.RawUI.WindowSize
    if ($size.Width -lt 100) {
        try {
            $bs = $Host.UI.RawUI.BufferSize
            $bs.Width = 110
            $Host.UI.RawUI.BufferSize = $bs
            $size.Width = 100
            $size.Height = 35
            $Host.UI.RawUI.WindowSize = $size
        } catch { }
    }
} catch { }

function Show-Banner {
    Clear-Host
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host '                  A M C   A U T O M A T I O N                ' -ForegroundColor Cyan
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host ''
}

function Show-Menu {
    param([hashtable] $Companies)

    Write-Host '  Choose what to process:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '     0. ALL companies' -ForegroundColor Green
    Write-Host ''

    $sortedKeys = $Companies.Keys | Sort-Object { $Companies[$_].Sheet }
    $i = 1
    foreach ($key in $sortedKeys) {
        $sheet = $Companies[$key].Sheet
        $folder = Join-Path (Join-Path $RootDir $Cfg.CompaniesDir) $Companies[$key].Folder
        $count = 0
        if (Test-Path $folder) {
            $count = @(Get-ChildItem -Path $folder -Filter '*.xlsm' -File -ErrorAction SilentlyContinue).Count
        }
        $line = '    {0,2}. {1,-22}' -f $i, $sheet
        if ($count -gt 0) {
            Write-Host -NoNewline $line
            Write-Host (' ({0} new patient file{1})' -f $count, ($(if($count -eq 1){''}else{'s'}))) -ForegroundColor Green
        } else {
            Write-Host $line -ForegroundColor DarkGray
        }
        $i++
    }

    Write-Host ''
    Write-Host '     P. Preview ALL (dry run, no changes written)' -ForegroundColor DarkCyan
    Write-Host '     Q. Quit' -ForegroundColor DarkGray
    Write-Host ''

    return $sortedKeys
}

# Load config to populate menu
$ConfigPath = Join-Path $ScriptDir 'config.psd1'
if (-not (Test-Path $ConfigPath)) {
    Show-Banner
    Write-Host "  ERROR: config.psd1 not found at $ConfigPath" -ForegroundColor Red
    Write-Host ''
    $null = Read-Host '  Press Enter to close'
    exit 1
}
$Cfg = Import-PowerShellDataFile $ConfigPath
$RootDir = $Cfg.RootDir

Show-Banner
$sortedKeys = Show-Menu -Companies $Cfg.Companies

$choice = Read-Host '  Your choice'
$choice = $choice.Trim()

$engineArgs = @()


switch -Regex ($choice) {
    '^[Qq]$' {
        Write-Host '' ; Write-Host '  Goodbye!' -ForegroundColor DarkGray ; Start-Sleep -Seconds 1
        exit 0
    }
    '^[Pp]$' {
        $engineArgs = @('all', '-DryRun')
        break
    }
    '^0$' {
        $engineArgs = @('all')
        break
    }
    '^[1-9]\d*$' {
        $idx = [int]$choice - 1
        if ($idx -ge $sortedKeys.Count) {
            Write-Host '' ; Write-Host "  Invalid number. Quitting." -ForegroundColor Red
            $null = Read-Host '  Press Enter to close'
            exit 1
        }
        $engineArgs = @($sortedKeys[$idx])
        break
    }
    default {
        Write-Host '' ; Write-Host "  Invalid choice. Quitting." -ForegroundColor Red
        $null = Read-Host '  Press Enter to close'
        exit 1
    }
}


# Confirmation banner before launching the engine
Show-Banner
$action = if ($engineArgs -contains '-DryRun') {
    "PREVIEW (dry-run) of ALL companies"
} elseif ($engineArgs[0] -eq 'all') {
    "REAL run of ALL companies"
} else {
    "REAL run of: $($Cfg.Companies[$engineArgs[0]].Sheet)"
}
Write-Host "  -> $action" -ForegroundColor Yellow
Write-Host ''
Start-Sleep -Milliseconds 600

# Run the engine
$enginePath = Join-Path $ScriptDir 'Update-Tracker.ps1'
$exitCode = 0
try {
    & $enginePath @engineArgs
    $exitCode = $LASTEXITCODE
} catch {
    Write-Host ''
    Write-Host "  FATAL: $_" -ForegroundColor Red
    $exitCode = 99
}

Write-Host ''
Write-Host '  ============================================================' -ForegroundColor Cyan
if ($exitCode -eq 0) {
    Write-Host '                       FINISHED SUCCESSFULLY                  ' -ForegroundColor Green
} else {
    Write-Host '                  FINISHED WITH ERRORS / WARNINGS             ' -ForegroundColor Yellow
}
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host '  Review the messages above. Log files are in:' -ForegroundColor DarkGray
Write-Host ('  {0}' -f (Join-Path $RootDir $Cfg.LogsDir)) -ForegroundColor DarkGray
Write-Host ''

$null = Read-Host '  Press Enter to close this window'
exit $exitCode
