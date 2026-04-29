<#
.SYNOPSIS
    AMC Tracker write diagnostic.
    Opens the tracker, finds the next empty row in a chosen sheet,
    tries to write a single test value, and reports EXACTLY what
    COM does or doesn't allow. Tells you in plain English what to fix.

.USAGE
    Right-click -> Run with PowerShell
    OR
    powershell -ExecutionPolicy Bypass -File .\Test-OneWrite.ps1
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Load the same config the real script uses
$ConfigPath = Join-Path $ScriptDir 'config.psd1'
if (-not (Test-Path $ConfigPath)) {
    Write-Host "ERROR: config.psd1 not found at $ConfigPath" -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}
$Cfg = Import-PowerShellDataFile $ConfigPath
$TrackerPath = Join-Path $Cfg.RootDir $Cfg.TrackerRelPath

# Pick which sheet to test. Default to AL MUTAIRI since that's what was failing.
$TestSheet = 'AL MUTAIRI'
$TestColumn = 'E'   # the Name column - safe, simple, plain text

Clear-Host
Write-Host ''
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host '            AMC TRACKER  -  WRITE DIAGNOSTIC TOOL              ' -ForegroundColor Cyan
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Tracker:  $TrackerPath" -ForegroundColor DarkGray
Write-Host "  Test sheet:    $TestSheet" -ForegroundColor DarkGray
Write-Host "  Test column:   $TestColumn (Name)" -ForegroundColor DarkGray
Write-Host ''

if (-not (Test-Path $TrackerPath)) {
    Write-Host "  ERROR: Tracker not found." -ForegroundColor Red
    Read-Host '  Press Enter to close'
    exit 1
}

# -------------------- 1. File-system probe --------------------
Write-Host '[1/6] File-system probe...' -ForegroundColor Yellow
$item = Get-Item $TrackerPath
Write-Host ("       Size:           {0:N0} bytes" -f $item.Length) -ForegroundColor Gray
Write-Host ("       Last modified:  {0}" -f $item.LastWriteTime) -ForegroundColor Gray
Write-Host ("       IsReadOnly:     {0}" -f $item.IsReadOnly) -ForegroundColor $(if ($item.IsReadOnly) {'Red'} else {'Green'})

# Try to open exclusively -> tells us if anything else is holding it
$exclusive = $false
try {
    $fs = [System.IO.File]::Open($TrackerPath, 'Open', 'ReadWrite', 'None')
    $fs.Close(); $fs.Dispose()
    $exclusive = $true
} catch {
    Write-Host "       LOCKED BY:      $($_.Exception.Message)" -ForegroundColor Red
}
if ($exclusive) {
    Write-Host '       Lock test:      OK (no other process holds it)' -ForegroundColor Green
}
Write-Host ''

# -------------------- 2. Launch Excel --------------------
Write-Host '[2/6] Launching Excel...' -ForegroundColor Yellow
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible          = $false
$Excel.DisplayAlerts    = $false
$Excel.AskToUpdateLinks = $false
try { $Excel.AutomationSecurity = 3 } catch { }
Write-Host "       Excel version:  $($Excel.Version)" -ForegroundColor Gray
Write-Host ''

$wb = $null
$exitCode = 0
try {
    # -------------------- 3. Open tracker --------------------
    Write-Host '[3/6] Opening tracker...' -ForegroundColor Yellow
    $wb = $Excel.Workbooks.Open(
        $TrackerPath,
        $false,         # UpdateLinks
        $false,         # ReadOnly
        [Type]::Missing,
        [Type]::Missing,
        [Type]::Missing,
        $true,          # IgnoreReadOnlyRecommended
        [Type]::Missing,
        [Type]::Missing,
        $false          # Editable
    )
    Write-Host ("       Opened:         {0}" -f $wb.FullName) -ForegroundColor Gray
    Write-Host ("       ReadOnly:       {0}" -f $wb.ReadOnly) -ForegroundColor $(if ($wb.ReadOnly) {'Red'} else {'Green'})
    try {
        $autosave = $wb.AutoSaveOn
        Write-Host ("       AutoSaveOn:     {0}" -f $autosave) -ForegroundColor $(if ($autosave) {'Yellow'} else {'Green'})
        if ($autosave) {
            try { $wb.AutoSaveOn = $false; Write-Host '       AutoSave -> OFF (forced)' -ForegroundColor Green } catch { }
        }
    } catch { Write-Host '       AutoSaveOn:     n/a' -ForegroundColor DarkGray }
    Write-Host ''

    # -------------------- 4. Inspect target sheet --------------------
    Write-Host "[4/6] Inspecting sheet '$TestSheet'..." -ForegroundColor Yellow
    $sh = $null
    try { $sh = $wb.Sheets.Item($TestSheet) } catch {
        Write-Host "       SHEET NOT FOUND: '$TestSheet'" -ForegroundColor Red
        Write-Host '       Available sheets:' -ForegroundColor Yellow
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            Write-Host ("         - {0}" -f $wb.Sheets.Item($i).Name) -ForegroundColor Gray
        }
        $exitCode = 1
        return
    }

    Write-Host ("       Sheet name:     {0}" -f $sh.Name) -ForegroundColor Gray
    try {
        $protected = $sh.ProtectContents
        Write-Host ("       Protected:      {0}" -f $protected) -ForegroundColor $(if ($protected) {'Red'} else {'Green'})
    } catch { Write-Host '       Protected:      ?' -ForegroundColor DarkGray }

    # ListObjects (Tables) on the sheet
    try {
        $loCount = $sh.ListObjects.Count
        Write-Host ("       ListObjects:    {0}" -f $loCount) -ForegroundColor $(if ($loCount -gt 0) {'Yellow'} else {'Green'})
        if ($loCount -gt 0) {
            for ($i = 1; $i -le $loCount; $i++) {
                $lo = $sh.ListObjects.Item($i)
                Write-Host ("         Table: {0}  Range: {1}" -f $lo.Name, $lo.Range.Address) -ForegroundColor Yellow
            }
        }
    } catch { }

    # Find next empty row
    $row = 2
    while ($null -ne $sh.Cells.Item($row, 1).Value2 -and `
           [string]$sh.Cells.Item($row, 1).Value2 -ne '') { $row++ }
    Write-Host ("       Next empty row: {0}" -f $row) -ForegroundColor Gray

    # Check the target cell
    $targetCol = 0
    foreach ($c in $TestColumn.ToCharArray()) {
        $targetCol = $targetCol * 26 + ([byte][char]$c.ToString().ToUpper() - [byte][char]'A' + 1)
    }
    $cell = $sh.Cells.Item($row, $targetCol)
    Write-Host ("       Target cell:    {0}{1}" -f $TestColumn, $row) -ForegroundColor Gray
    try { Write-Host ("       Cell.Locked:    {0}" -f $cell.Locked) -ForegroundColor Gray } catch { }
    try { Write-Host ("       Cell.HasFormula: {0}" -f $cell.HasFormula) -ForegroundColor Gray } catch { }
    try { Write-Host ("       MergeCells:     {0}" -f $cell.MergeCells) -ForegroundColor $(if ($cell.MergeCells) {'Red'} else {'Green'}) } catch { }
    try {
        $vCount = $cell.Validation.Type
        Write-Host ("       Has validation: {0}" -f ($vCount -ge 0)) -ForegroundColor Gray
    } catch { Write-Host '       Has validation: no' -ForegroundColor Gray }
    Write-Host ''

    # -------------------- 5. Try writes - 4 different methods --------------------
    Write-Host '[5/6] TESTING WRITES (4 different methods)...' -ForegroundColor Yellow
    Write-Host ''

    $testValue = "DIAGNOSTIC_TEST_$(Get-Date -Format HHmmss)"
    $methods = @(
        @{ Name = 'Method A: Cells.Item().Value2 = ...'; Action = { $sh.Cells.Item($row, $targetCol).Value2 = $testValue } }
        @{ Name = 'Method B: Cells.Item().Value  = ...'; Action = { $sh.Cells.Item($row, $targetCol).Value  = $testValue } }
        @{ Name = "Method C: Range('${TestColumn}${row}').Value2 = ..."; Action = { $sh.Range("${TestColumn}${row}").Value2 = $testValue } }
        @{ Name = 'Method D: Cells.Item().Formula = "=..."'; Action = { $sh.Cells.Item($row, $targetCol).Formula = "=`"$testValue`"" } }
    )

    $successMethods = @()
    foreach ($m in $methods) {
        Write-Host "       $($m.Name)" -ForegroundColor White
        # Clear the cell first so each method starts fresh
        try { $sh.Cells.Item($row, $targetCol).ClearContents() | Out-Null } catch { }
        try {
            & $m.Action
            # Read back
            $readback = $sh.Cells.Item($row, $targetCol).Value2
            if ($readback -like "*$testValue*" -or $readback -like "*DIAGNOSTIC_TEST*") {
                Write-Host "         -> SUCCESS  read back: $readback" -ForegroundColor Green
                $successMethods += $m.Name
            } else {
                Write-Host "         -> WROTE but readback is: '$readback' (mismatch)" -ForegroundColor Yellow
            }
        } catch {
            $ex = $_.Exception
            Write-Host "         -> FAILED" -ForegroundColor Red
            Write-Host "            Message:    $($ex.Message)" -ForegroundColor Red
            Write-Host "            Type:       $($ex.GetType().FullName)" -ForegroundColor DarkRed
            if ($ex.InnerException) {
                Write-Host "            InnerEx:    $($ex.InnerException.Message)" -ForegroundColor DarkRed
            }
            try {
                $hr = $ex.HResult
                Write-Host ("            HResult:    0x{0:X8}" -f $hr) -ForegroundColor DarkRed
            } catch { }
        }
    }

    # Clear the test cell
    try { $sh.Cells.Item($row, $targetCol).ClearContents() | Out-Null } catch { }
    Write-Host ''

    # -------------------- 6. Save test --------------------
    Write-Host '[6/6] Save test (no actual changes - just trying Save)...' -ForegroundColor Yellow
    $beforeTime = (Get-Item $TrackerPath).LastWriteTime
    try {
        $wb.Save()
        Start-Sleep -Seconds 1
        $afterTime = (Get-Item $TrackerPath).LastWriteTime
        if ($afterTime -gt $beforeTime) {
            Write-Host '       Save:           SUCCESS (file timestamp updated)' -ForegroundColor Green
        } else {
            Write-Host '       Save:           Returned OK but timestamp DID NOT change' -ForegroundColor Yellow
        }
    } catch {
        Write-Host "       Save:           FAILED - $($_.Exception.Message)" -ForegroundColor Red
    }

    # ----------------- DIAGNOSIS -----------------
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host '                      D I A G N O S I S                        ' -ForegroundColor Cyan
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host ''
    if ($successMethods.Count -eq 0) {
        Write-Host '  NO WRITE METHOD WORKED.' -ForegroundColor Red
        Write-Host '  Look at the FAILED messages in step 5 above. Most common:' -ForegroundColor Yellow
        Write-Host '    "0x800A03EC" -> sheet is protected, cell is locked, or workbook restricted' -ForegroundColor Gray
        Write-Host '    "Old format" -> Excel opened tracker in Compatibility / Protected View' -ForegroundColor Gray
        Write-Host '    "Sensitivity" -> Microsoft Information Protection label blocking writes' -ForegroundColor Gray
    } elseif ($successMethods.Count -eq 4) {
        Write-Host '  ALL 4 METHODS WORKED FROM THIS DIAGNOSTIC.' -ForegroundColor Green
        Write-Host '  This means COM writing to plain cells works fine here.' -ForegroundColor Green
        Write-Host '  The real script must be hitting a different cell or condition.' -ForegroundColor Yellow
        Write-Host '  Send me this output and I will dig deeper.' -ForegroundColor Yellow
    } else {
        Write-Host "  $($successMethods.Count) of 4 methods worked:" -ForegroundColor Green
        foreach ($m in $successMethods) { Write-Host "    - $m" -ForegroundColor Green }
        Write-Host ''
        Write-Host '  We will switch the real script to use the working method.' -ForegroundColor Yellow
    }
    Write-Host ''

} catch {
    Write-Host ''
    Write-Host "  FATAL EXCEPTION: $_" -ForegroundColor Red
    Write-Host "  Type: $($_.Exception.GetType().FullName)" -ForegroundColor DarkRed
    $exitCode = 1
} finally {
    if ($wb) {
        try { $wb.Close($false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
    }
    if ($Excel) {
        try { $Excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

Write-Host ''
Write-Host '  Send me a screenshot of this whole window and we will fix the real script.' -ForegroundColor Cyan
Write-Host ''
Read-Host '  Press Enter to close'
exit $exitCode
