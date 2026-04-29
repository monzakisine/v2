<#
.SYNOPSIS
    Reads the actual column headers from the tracker and compares
    them to what config.psd1 says. Prints the correct mapping.
#>

$ErrorActionPreference = 'Continue'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$ConfigPath = Join-Path $ScriptDir 'config.psd1'
$Cfg = Import-PowerShellDataFile $ConfigPath
$TrackerPath = Join-Path $Cfg.RootDir $Cfg.TrackerRelPath

function ColLetter {
    param([int]$n)
    $r = ''
    while ($n -gt 0) {
        $n--
        $r = [char]([byte][char]'A' + ($n % 26)) + $r
        $n = [Math]::Floor($n / 26)
    }
    return $r
}

Clear-Host
Write-Host ''
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host '            AMC TRACKER  -  COLUMN MAPPING DIAGNOSTIC          ' -ForegroundColor Cyan
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false; $Excel.DisplayAlerts = $false
$Excel.AskToUpdateLinks = $false
try { $Excel.AutomationSecurity = 3 } catch { }

$wb = $null
try {
    $wb = $Excel.Workbooks.Open($TrackerPath, $false, $true)

    # Test every company sheet
    foreach ($key in ($Cfg.Companies.Keys | Sort-Object)) {
        $sheetName = $Cfg.Companies[$key].Sheet
        $sh = $null
        try { $sh = $wb.Sheets.Item($sheetName) } catch { continue }

        Write-Host ("  === Sheet: {0} ===" -f $sheetName) -ForegroundColor Yellow

        # Read header row (row 1) - go up to column 60
        $headers = @{}
        for ($col = 1; $col -le 60; $col++) {
            $raw = $null
            try { $raw = $sh.Cells.Item(1, $col).Value2 } catch { break }
            if ($null -ne $raw -and [string]$raw -ne '') {
                $letter = ColLetter $col
                $headers[$letter] = [string]$raw
                Write-Host ("    {0,-4} {1}" -f $letter, $raw) -ForegroundColor Gray
            }
        }

        # Check the first data row to understand data
        Write-Host ''
        Write-Host '  First data row (row 2):' -ForegroundColor DarkCyan
        for ($col = 1; $col -le 60; $col++) {
            $raw = $null
            try { $raw = $sh.Cells.Item(2, $col).Value2 } catch { break }
            if ($null -ne $raw -and [string]$raw -ne '') {
                $letter = ColLetter $col
                $hdr = if ($headers.ContainsKey($letter)) { $headers[$letter] } else { '?' }
                Write-Host ("    {0,-4} [{1,-22}] = {2}" -f $letter, $hdr, ([string]$raw).Substring(0, [Math]::Min(30, ([string]$raw).Length))) -ForegroundColor Gray
            }
        }
        Write-Host ''

        # Only show first sheet fully, rest just show header count
        if ($key -ne ($Cfg.Companies.Keys | Sort-Object | Select-Object -First 1)) {
            break
        }
    }

    # Show SCMS and AL MUTAIRI specifically
    foreach ($sheetName in @('AL MUTAIRI', 'SCMS')) {
        $sh = $null
        try { $sh = $wb.Sheets.Item($sheetName) } catch { continue }

        Write-Host ("  === Full header map: {0} ===" -f $sheetName) -ForegroundColor Yellow
        for ($col = 1; $col -le 60; $col++) {
            $raw = $null
            try { $raw = $sh.Cells.Item(1, $col).Value2 } catch { break }
            if ($null -ne $raw -and [string]$raw -ne '') {
                $letter = ColLetter $col
                Write-Host ("    {0,-4} {1}" -f $letter, $raw) -ForegroundColor White
            }
        }
        Write-Host ''

        # Find next empty row and show surrounding rows
        $nextRow = 2
        while ($true) {
            $v = $null
            try { $v = $sh.Cells.Item($nextRow, 1).Value2 } catch { break }
            if ($null -eq $v -or [string]$v -eq '') { break }
            $nextRow++
        }
        Write-Host ("  Next empty row in {0}: {1}" -f $sheetName, $nextRow) -ForegroundColor Green

        Write-Host ("  Row {0} (last data row) col A-F:" -f ($nextRow - 1)) -ForegroundColor DarkCyan
        for ($col = 1; $col -le 6; $col++) {
            $v = $null
            try { $v = $sh.Cells.Item($nextRow - 1, $col).Value2 } catch { }
            $letter = ColLetter $col
            Write-Host ("    {0,-4} = {1}" -f $letter, $v) -ForegroundColor Gray
        }
        Write-Host ''
    }

} catch {
    Write-Host "FATAL: $_" -ForegroundColor Red
} finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($Excel) { try { $Excel.Quit() } catch { } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

Write-Host '  Screenshot this entire window and send it.' -ForegroundColor Cyan
Write-Host ''
Read-Host '  Press Enter to close'
