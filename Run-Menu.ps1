<#
.SYNOPSIS
    AMC Automation menu launcher. Calls amc_engine.py via Python.
    Double-click amc.bat to launch this.
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue } catch { }
try {
    Get-ChildItem -Path $ScriptDir -Filter '*.ps1' -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
} catch { }

$ConfigPath = Join-Path $ScriptDir 'config.psd1'
$Cfg        = Import-PowerShellDataFile $ConfigPath
$RootDir    = $Cfg.RootDir
$EnginePath = Join-Path $ScriptDir 'amc_engine.py'

# ── Window setup ──────────────────────────────────────────────────────────────
try {
    $Host.UI.RawUI.WindowTitle = 'AMC Tracker Automation'
    try {
        $bs = $Host.UI.RawUI.BufferSize; $bs.Width = 110; $Host.UI.RawUI.BufferSize = $bs
        $ws = $Host.UI.RawUI.WindowSize; $ws.Width = 100; $ws.Height = 45; $Host.UI.RawUI.WindowSize = $ws
    } catch { }
} catch { }

# ── Find Python ───────────────────────────────────────────────────────────────
function Find-Python {
    foreach ($cmd in @('python', 'python3', 'py')) {
        try {
            $ver = & $cmd --version 2>&1
            if ($ver -match 'Python 3') { return $cmd }
        } catch { }
    }
    return $null
}

function Show-Banner {
    Clear-Host
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host '                  A M C   A U T O M A T I O N                ' -ForegroundColor Cyan
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host ''
}

function Show-Menu {
    $sortedKeys = $Cfg.Companies.Keys | Sort-Object { $Cfg.Companies[$_].Sheet }
    Write-Host '  Choose what to process:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '     0. ALL companies' -ForegroundColor Green
    Write-Host ''
    $i = 1
    foreach ($key in $sortedKeys) {
        $sheet  = $Cfg.Companies[$key].Sheet
        $folder = Join-Path (Join-Path $RootDir $Cfg.CompaniesDir) $Cfg.Companies[$key].Folder
        $count  = 0
        if (Test-Path $folder) {
            $count = @(Get-ChildItem -Path $folder -Filter '*.xlsm' -File -ErrorAction SilentlyContinue).Count
        }
        $line = '    {0,2}. {1,-22}' -f $i, $sheet
        if ($count -gt 0) {
            Write-Host -NoNewline $line
            Write-Host (' ({0} new file{1})' -f $count, $(if($count -eq 1){''} else {'s'})) -ForegroundColor Green
        } else {
            Write-Host $line -ForegroundColor DarkGray
        }
        $i++
    }
    Write-Host ''
    Write-Host '     P. Preview ALL (dry run - no changes written)' -ForegroundColor DarkCyan
    Write-Host '     Q. Quit' -ForegroundColor DarkGray
    Write-Host ''
    return $sortedKeys
}

Show-Banner
$pythonCmd = Find-Python
if (-not $pythonCmd) {
    Write-Host '  ERROR: Python 3 is not installed or not in PATH.' -ForegroundColor Red
    Write-Host ''
    Write-Host '  Please install Python from: https://www.python.org/downloads/' -ForegroundColor Yellow
    Write-Host '  Tick "Add Python to PATH" during installation.' -ForegroundColor Yellow
    Write-Host ''
    $null = Read-Host '  Press Enter to close'
    exit 1
}
Write-Host "  Python: $pythonCmd" -ForegroundColor DarkGray
Write-Host ''

# Ensure openpyxl is installed silently
$hasOpenpyxl = & $pythonCmd -c "import openpyxl; print('ok')" 2>&1
if ($hasOpenpyxl -ne 'ok') {
    Write-Host '  Installing openpyxl (one-time setup)...' -ForegroundColor Yellow
    & $pythonCmd -m pip install openpyxl --quiet
    Write-Host '  Done.' -ForegroundColor Green
    Write-Host ''
}

$sortedKeys = Show-Menu
$choice = (Read-Host '  Your choice').Trim()

$engineKey = $null
$dryRun    = $false

switch -Regex ($choice) {
    '^[Qq]$' {
        Write-Host '  Goodbye!' -ForegroundColor DarkGray; Start-Sleep -Seconds 1; exit 0
    }
    '^[Pp]$' {
        $engineKey = 'all'; $dryRun = $true; break
    }
    '^0$' {
        $engineKey = 'all'; break
    }
    '^[1-9]\d*$' {
        $idx = [int]$choice - 1
        if ($idx -ge $sortedKeys.Count) {
            Write-Host '  Invalid number.' -ForegroundColor Red
            $null = Read-Host '  Press Enter to close'; exit 1
        }
        $engineKey = $sortedKeys[$idx]; break
    }
    default {
        Write-Host '  Invalid choice.' -ForegroundColor Red
        $null = Read-Host '  Press Enter to close'; exit 1
    }
}

Show-Banner
$action = if ($dryRun) { "PREVIEW (dry-run): $engineKey" } else { "LIVE run: $engineKey" }
Write-Host "  -> $action" -ForegroundColor Yellow
Write-Host ''
Start-Sleep -Milliseconds 400

$pyArgs = @($EnginePath, $engineKey, '--root', $RootDir)
if ($dryRun) { $pyArgs += '--dry-run' }

$exitCode = 0
try {
    & $pythonCmd @pyArgs
    $exitCode = $LASTEXITCODE
} catch {
    Write-Host "  FATAL: $_" -ForegroundColor Red
    $exitCode = 99
}

Write-Host ''
Write-Host '  ============================================================' -ForegroundColor Cyan
if ($exitCode -eq 0) {
    Write-Host '                       FINISHED SUCCESSFULLY                  ' -ForegroundColor Green
} else {
    Write-Host '               FINISHED WITH ERRORS / WARNINGS                ' -ForegroundColor Yellow
}
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''
$null = Read-Host '  Press Enter to close'
exit $exitCode
