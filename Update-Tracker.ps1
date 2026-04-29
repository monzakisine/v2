<#
.SYNOPSIS
    Reads filled AMCFORMULA patient files and appends rows to the tracker.
.PARAMETER Company
    Company key from config.psd1, or 'all'.
.PARAMETER DryRun
    Preview only - nothing written, nothing moved.
.PARAMETER NoArchive
    Skip moving files to Archive after writing.
.PARAMETER ConfigPath
    Optional override for the config file path.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)] [string] $Company = 'all',
    [switch] $DryRun,
    [switch] $NoArchive,
    [string] $ConfigPath
)

$ErrorActionPreference = 'Stop'

# ============================================================
# RETRY WRAPPER
# ============================================================
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock] $Action,
        [string] $Label = 'operation',
        [int] $MaxAttempts = 6,
        [int] $InitialDelayMs = 250
    )
    $attempt = 0; $delay = $InitialDelayMs; $lastErr = $null
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try { return & $Action }
        catch {
            $lastErr = $_
            if ($attempt -ge $MaxAttempts) { break }
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, 4000)
        }
    }
    throw ("Failed after {0} attempts ({1}): {2}" -f $MaxAttempts, $Label, $lastErr.Exception.Message)
}

# ============================================================
# COLUMN LETTER -> INDEX
# ============================================================
function ConvertTo-ColIndex {
    param([string] $Letter)
    $idx = 0
    foreach ($c in $Letter.ToUpper().ToCharArray()) {
        $idx = $idx * 26 + ([byte][char]$c - [byte][char]'A' + 1)
    }
    return $idx
}

# ============================================================
# SAFE VALUE FROM COM
# Routes all numerics through [double] to avoid Int32 overflow
# on 10-digit Iqama numbers.
# ============================================================
function Get-SafeValue {
    param($Raw)
    if ($null -eq $Raw) { return $null }
    if ($Raw -is [string]) { return $Raw.Trim() }
    if ($Raw -is [bool])   { return $Raw }
    try {
        $d = [double]$Raw
        if ($d -lt -999999) { return $null }   # Excel error sentinel
        return $d
    } catch {
        try { return [string]$Raw } catch { return $null }
    }
}

# ============================================================
# DETECT ABNORMAL (red fill on cell)
# ============================================================
function Test-CellIsHighlighted {
    param($Cell)
    try {
        $idx = [double]$Cell.Interior.ColorIndex
        if ($idx -eq -4142 -or $idx -eq 0) { return $false }
        $color = [double]$Cell.Interior.Color
        if ($color -eq 16777215) { return $false }   # plain white
        return $true
    } catch { return $false }
}

# ============================================================
# RELEASE EXCEL WORKBOOK + GC
# ============================================================
function Release-Workbook {
    param($Workbook)
    if ($null -ne $Workbook) {
        try { $Workbook.Close($false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    Start-Sleep -Milliseconds 200
}

# ============================================================
# PRE-CACHE: force SharePoint to download file before Excel
# ============================================================
function Invoke-PreCache {
    param([string] $Path)
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
        $buf = New-Object byte[] 512
        [void]$fs.Read($buf, 0, $buf.Length)
        $fs.Close(); $fs.Dispose()
        Start-Sleep -Milliseconds 100
    } catch { }
}

# ============================================================
# READ ONE PATIENT AMCFORMULA FILE
# ============================================================
function Read-PatientFile {
    param(
        [Parameter(Mandatory)] $ExcelApp,
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [hashtable] $Cfg
    )

    Invoke-PreCache -Path $Path

    $wb = Invoke-WithRetry -Label "open '$([IO.Path]::GetFileName($Path))'" -Action {
        $ExcelApp.Workbooks.Open($Path, $false, $true)
    }

    try {
        $ws = $wb.Sheets.Item($Cfg.SourceSheet)

        $data = [ordered]@{
            Name = $null; Company = $null; Iqama = $null
            Age  = $null; DateAMC = $null; DateReview = $null
            BloodPress = $null; Height = $null; Weight = $null
            Status = $null; Comment = $null; Tests = @{}
        }

        # Identity cells
        foreach ($k in $Cfg.PatientCells.Keys) {
            $raw = $null
            try { $raw = $ws.Range($Cfg.PatientCells[$k]).Value2 } catch { }
            $data[$k] = Get-SafeValue $raw
        }

        # Iqama must be a clean string of digits (no decimal point)
        if ($null -ne $data.Iqama) {
            $iqDouble = $null
            try { $iqDouble = [double]$data.Iqama } catch { }
            if ($null -ne $iqDouble) {
                $data.Iqama = ([long]$iqDouble).ToString()
            } else {
                $data.Iqama = ([string]$data.Iqama).Trim()
            }
        }

        # Status: checkmark (ChrW 10003) in I4:I7
        foreach ($s in $Cfg.StatusCandidates) {
            $raw = $null
            try { $raw = $ws.Range($s.CheckCell).Value2 } catch { }
            if ($null -ne $raw -and [string]$raw -ne '') {
                $data.Status = $s.Label; break
            }
        }

        # Test results
        foreach ($t in $Cfg.TestRowMap) {
            $row = [int]$t.FormulaRow

            if ($t.ContainsKey('MinAge') -and $null -ne $data.Age) {
                $ageVal = 0L
                try { $ageVal = [long][double]$data.Age } catch { }
                if ($ageVal -lt [long]$t.MinAge) {
                    $data.Tests[$t.TrackerCol] = $null; continue
                }
            }

            $gCell = $null
            try { $gCell = $ws.Cells.Item($row, 7) } catch { }
            $data.Tests[$t.TrackerCol] = if ($null -ne $gCell -and (Test-CellIsHighlighted $gCell)) {
                'ABNORMAL'
            } else { 'NORMAL' }
        }

        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch { }
        return $data
    } finally {
        Release-Workbook $wb
    }
}

# ============================================================
# WRITE ONE ROW TO TRACKER
# Direct individual cell writes — no script block scope tricks.
# ============================================================
function Write-TrackerRow {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] [int] $RowIndex,
        [Parameter(Mandatory)] [int] $SerialNumber,
        [Parameter(Mandatory)] [hashtable] $Patient,
        [Parameter(Mandatory)] [hashtable] $Cfg
    )
    $fc = $Cfg.FixedColumns

    # Helper used inline so $Sheet and $RowIndex are always in scope
    function Set-Cell {
        param([string]$ColLetter, $Value, [bool]$AsText = $false)
        if ($null -eq $Value) { return }
        try {
            $c = $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $ColLetter))
            if ($AsText) { $c.NumberFormat = '@' }
            $c.Value2 = $Value
        } catch {
            Write-Host ("    cell-write ERR  col={0} val={1}: {2}" -f $ColLetter, $Value, $_.Exception.Message) -ForegroundColor Magenta
        }
    }

    Set-Cell $fc.SerialNumber $SerialNumber
    Set-Cell $fc.DateAMC      $Patient.DateAMC
    Set-Cell $fc.DateReview   $Patient.DateReview
    Set-Cell $fc.Iqama        $Patient.Iqama   $true    # text format
    Set-Cell $fc.Name         $Patient.Name
    Set-Cell $fc.Company      $Patient.Company
    Set-Cell $fc.Height       $Patient.Height
    Set-Cell $fc.Weight       $Patient.Weight
    Set-Cell $fc.Age          $Patient.Age
    Set-Cell $fc.BloodPress   $Patient.BloodPress
    Set-Cell $fc.Status       $Patient.Status
    Set-Cell $fc.Comment      $Patient.Comment

    # BMI formula
    try {
        $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.BMIFormula)).Formula = `
            ('=J{0}/(I{0}/100)^2' -f $RowIndex)
    } catch { }

    # Test results
    foreach ($col in @($Patient.Tests.Keys)) {
        $val = $Patient.Tests[$col]
        if ($null -eq $val) { continue }
        try {
            $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $col)).Value2 = $val
        } catch {
            Write-Host ("    cell-write ERR  col={0} val={1}: {2}" -f $col, $val, $_.Exception.Message) -ForegroundColor Magenta
        }
    }
}

# ============================================================
# FIND NEXT EMPTY ROW
# A row is considered empty if BOTH column A and column E are
# blank. This skips ghost rows that have only a serial number
# from a previous partial run.
# ============================================================
function Get-NextEmptyRow {
    param($Sheet)
    $row = 2
    while ($true) {
        $snVal   = $null
        $nameVal = $null
        try { $snVal   = $Sheet.Cells.Item($row, 1).Value2 } catch { break }
        try { $nameVal = $Sheet.Cells.Item($row, 5).Value2 } catch { }

        $snEmpty   = ($null -eq $snVal   -or [string]$snVal   -eq '')
        $nameEmpty = ($null -eq $nameVal -or [string]$nameVal -eq '')

        # Truly empty row = nothing in SN AND nothing in Name
        if ($snEmpty -and $nameEmpty) { break }
        $row++
    }
    return $row
}

# ============================================================
# GET SERIAL NUMBER FOR NEXT ROW
# Scans up from nextRow to find last real SN value.
# ============================================================
function Get-NextSerialNumber {
    param($Sheet, [int]$NextEmptyRow)
    if ($NextEmptyRow -le 2) { return 1 }
    # Walk backwards to find a row with both SN and Name filled
    for ($r = $NextEmptyRow - 1; $r -ge 2; $r--) {
        $snVal   = $null
        $nameVal = $null
        try { $snVal   = $Sheet.Cells.Item($r, 1).Value2 } catch { }
        try { $nameVal = $Sheet.Cells.Item($r, 5).Value2 } catch { }
        if ($null -ne $snVal -and [string]$snVal -ne '' -and
            $null -ne $nameVal -and [string]$nameVal -ne '') {
            $d = 0.0
            if ([double]::TryParse([string]$snVal, [ref]$d)) {
                return [int]$d + 1
            }
        }
    }
    return 1
}

# ============================================================
# CHECK IQAMA EXISTS IN COLUMN D
# ============================================================
function Test-IqamaExists {
    param($Sheet, [string]$Iqama)
    if ([string]::IsNullOrWhiteSpace($Iqama)) { return $false }
    $row = 2
    while ($true) {
        $sn = $null
        try { $sn = $Sheet.Cells.Item($row, 1).Value2 } catch { break }
        if ($null -eq $sn -or [string]$sn -eq '') { break }
        $iq = $null
        try { $iq = $Sheet.Cells.Item($row, 4).Value2 } catch { }
        if ($null -ne $iq) {
            $iqStr = ''
            try { $iqStr = ([long][double]$iq).ToString() } catch { $iqStr = [string]$iq }
            if ($iqStr.Trim() -eq $Iqama) { return $true }
        }
        $row++
    }
    return $false
}

# ============================================================
# ARCHIVE PATIENT FILE (copy+delete, safer than Move-Item)
# ============================================================
function Move-PatientToArchive {
    param(
        [Parameter(Mandatory)] [string] $XlsmPath,
        [Parameter(Mandatory)] [string] $ArchiveCompanyDir
    )
    if (-not (Test-Path $ArchiveCompanyDir)) {
        New-Item -ItemType Directory -Path $ArchiveCompanyDir -Force | Out-Null
    }
    $base = [IO.Path]::GetFileNameWithoutExtension($XlsmPath)
    $ext  = [IO.Path]::GetExtension($XlsmPath)
    $dest = Join-Path $ArchiveCompanyDir ([IO.Path]::GetFileName($XlsmPath))
    if (Test-Path $dest) {
        $dest = Join-Path $ArchiveCompanyDir ('{0}_{1:yyyyMMdd-HHmmss}{2}' -f $base, (Get-Date), $ext)
    }

    Invoke-WithRetry -Label "copy $([IO.Path]::GetFileName($XlsmPath))" -MaxAttempts 8 -InitialDelayMs 400 -Action {
        Copy-Item -Path $XlsmPath -Destination $dest -Force -ErrorAction Stop
    }

    $srcSize = (Get-Item $XlsmPath).Length
    $dstSize = (Get-Item $dest).Length
    if ($dstSize -eq $srcSize -and $dstSize -gt 0) {
        try {
            Invoke-WithRetry -Label "delete original" -MaxAttempts 5 -InitialDelayMs 500 -Action {
                Remove-Item -Path $XlsmPath -Force -ErrorAction Stop
            }
            Write-Host ("    Archived -> $dest") -ForegroundColor DarkGray
        } catch {
            Write-Host ("    Copied (original stays, skipped next run) -> $dest") -ForegroundColor DarkGray
        }
    } else {
        try { Remove-Item $dest -Force -ErrorAction SilentlyContinue } catch { }
        throw "Copy size mismatch. Original left in place."
    }

    # PDF
    $pdfSrc = Join-Path (Split-Path -Parent $XlsmPath) "$base.pdf"
    if (Test-Path $pdfSrc) {
        $pdfDest = Join-Path $ArchiveCompanyDir "$base.pdf"
        if (Test-Path $pdfDest) {
            $pdfDest = Join-Path $ArchiveCompanyDir ('{0}_{1:yyyyMMdd-HHmmss}.pdf' -f $base, (Get-Date))
        }
        try {
            Invoke-WithRetry -Label "copy pdf" -MaxAttempts 6 -Action {
                Copy-Item -Path $pdfSrc -Destination $pdfDest -Force -ErrorAction Stop
            }
            try { Remove-Item -Path $pdfSrc -Force -ErrorAction SilentlyContinue } catch { }
        } catch {
            Write-Host "    (pdf archive skipped: $_)" -ForegroundColor DarkGray
        }
    }
}

# ============================================================
#                          M A I N
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir 'config.psd1' }
if (-not (Test-Path $ConfigPath)) { throw "Config not found: $ConfigPath" }
$Cfg = Import-PowerShellDataFile $ConfigPath

$RootDir      = $Cfg.RootDir
$TrackerPath  = Join-Path $RootDir $Cfg.TrackerRelPath
$CompaniesDir = Join-Path $RootDir $Cfg.CompaniesDir
$ArchiveDir   = Join-Path $RootDir $Cfg.ArchiveDir

foreach ($d in @($CompaniesDir, $ArchiveDir)) {
    if (-not (Test-Path $d)) {
        try { New-Item -ItemType Directory -Path $d -Force | Out-Null } catch { }
    }
}

Write-Host ''
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host "  AMC Automation  |  Company='$Company'  |  DryRun=$DryRun" -ForegroundColor Cyan
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''

if (-not (Test-Path $TrackerPath)) {
    Write-Host "  ERROR: Tracker not found: $TrackerPath" -ForegroundColor Red
    exit 1
}

# Check tracker is not locked
try {
    $fs = [System.IO.File]::Open($TrackerPath, 'Open', 'ReadWrite', 'None')
    $fs.Close(); $fs.Dispose()
} catch {
    Write-Host '  ERROR: Tracker is open in Excel or locked by another process.' -ForegroundColor Red
    Write-Host '  Close the tracker in Excel and try again.' -ForegroundColor Yellow
    exit 1
}

$AllKeys = $Cfg.Companies.Keys | Sort-Object
$keys = if ($Company.ToLower() -eq 'all') {
    $AllKeys
} else {
    $hit = $AllKeys | Where-Object { $_ -ieq $Company }
    if (-not $hit) {
        Write-Host "  ERROR: Unknown company '$Company'. Valid: $($AllKeys -join ', ')" -ForegroundColor Red
        exit 1
    }
    @($hit)
}

# Backup tracker
if ($Cfg.BackupTrackerBeforeRun -and -not $DryRun) {
    $bkpDir = Join-Path $RootDir $Cfg.LogsDir
    if (-not (Test-Path $bkpDir)) {
        try { New-Item -ItemType Directory -Path $bkpDir -Force | Out-Null } catch { }
    }
    $bkp = Join-Path $bkpDir ('tracker-backup-{0:yyyy-MM-dd_HHmmss}.xlsm' -f (Get-Date))
    try {
        Invoke-WithRetry -Label 'backup tracker' -Action {
            Copy-Item -Path $TrackerPath -Destination $bkp -Force -ErrorAction Stop
        }
        Write-Host "  Tracker backed up -> $bkp" -ForegroundColor DarkGray
    } catch {
        Write-Host "  WARNING: Backup failed ($_). Continuing." -ForegroundColor Yellow
    }
}

Write-Host '  Launching Excel...' -ForegroundColor DarkGray
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false; $Excel.DisplayAlerts = $false; $Excel.AskToUpdateLinks = $false
try { $Excel.AutomationSecurity = 3 } catch { }

$Tracker         = $null
$totalProcessed  = 0
$totalSkipped    = 0
$totalErrors     = 0
$startTime       = Get-Date
$perCompanyStats = [ordered]@{}

try {
    $Tracker = Invoke-WithRetry -Label 'open tracker' -Action {
        $Excel.Workbooks.Open($TrackerPath, $false, $false)
    }
    try { $Tracker.AutoSaveOn = $false } catch { }

    if ($Tracker.ReadOnly) {
        throw "Tracker opened READ-ONLY. Close it in Excel everywhere and retry."
    }
    Write-Host ("  Tracker: {0}  ReadOnly={1}" -f $Tracker.FullName, $Tracker.ReadOnly) -ForegroundColor DarkGray
    Write-Host ''

    $companyIdx   = 0
    $companyTotal = $keys.Count

    foreach ($key in $keys) {
        $companyIdx++
        $info       = $Cfg.Companies[$key]
        $sheetName  = $info.Sheet
        $folderName = $info.Folder
        $folderPath = Join-Path $CompaniesDir $folderName

        $coStats = [ordered]@{ Sheet = $sheetName; Written = 0; Skipped = 0; Errors = 0 }
        $perCompanyStats[$sheetName] = $coStats

        Write-Progress -Id 0 -Activity 'Processing companies' `
            -Status ("$sheetName ($companyIdx of $companyTotal)") `
            -PercentComplete ([int](($companyIdx - 1) / [Math]::Max(1,$companyTotal) * 100))

        Write-Host ("  [{0}/{1}] {2}" -f $companyIdx, $companyTotal, $sheetName) -ForegroundColor White

        if (-not (Test-Path $folderPath)) {
            try { New-Item -ItemType Directory -Path $folderPath -Force | Out-Null } catch { }
            Write-Host "         Folder created." -ForegroundColor DarkGray
            continue
        }

        $files = @(Get-ChildItem -Path $folderPath -Filter '*.xlsm' -File -ErrorAction SilentlyContinue)
        if (-not $files -or $files.Count -eq 0) {
            Write-Host "         No files." -ForegroundColor DarkGray
            continue
        }

        $trackerSheet = $null
        try { $trackerSheet = $Tracker.Sheets.Item($sheetName) }
        catch {
            Write-Host "         ERROR: Sheet '$sheetName' not found in tracker." -ForegroundColor Red
            $totalErrors += $files.Count; $coStats.Errors += $files.Count
            continue
        }

        $nextRow = Get-NextEmptyRow $trackerSheet
        $nextSN  = Get-NextSerialNumber $trackerSheet $nextRow
        Write-Host ("         Next row: {0}  Next SN: {1}  Files: {2}" -f $nextRow, $nextSN, $files.Count) -ForegroundColor DarkGray

        $fileIdx = 0
        foreach ($file in $files) {
            $fileIdx++
            Write-Progress -Id 1 -ParentId 0 -Activity "Files in $sheetName" `
                -Status ("$($file.Name) ($fileIdx of $($files.Count))") `
                -PercentComplete ([int](($fileIdx - 1) / [Math]::Max(1,$files.Count) * 100))

            try {
                $patient = Read-PatientFile -ExcelApp $Excel -Path $file.FullName -Cfg $Cfg

                if ([string]::IsNullOrWhiteSpace($patient.Iqama)) {
                    Write-Host "         SKIP $($file.Name) - Iqama empty" -ForegroundColor Yellow
                    $totalSkipped++; $coStats.Skipped++; continue
                }
                if (-not $patient.Status) {
                    Write-Host "         WARN $($file.Name) - no status checkmark" -ForegroundColor Yellow
                }

                $exists = Test-IqamaExists -Sheet $trackerSheet -Iqama $patient.Iqama
                if ($exists -and $Cfg.OnDuplicateIqama -eq 'skip') {
                    Write-Host "         SKIP $($file.Name) - Iqama $($patient.Iqama) already exists" -ForegroundColor Yellow
                    $totalSkipped++; $coStats.Skipped++; continue
                }
                if ($exists) {
                    Write-Host "         WARN Iqama $($patient.Iqama) already exists - adding anyway" -ForegroundColor Yellow
                }

                if ($DryRun) {
                    $abn = ($patient.Tests.GetEnumerator() | Where-Object { $_.Value -eq 'ABNORMAL' } |
                            ForEach-Object { $_.Key }) -join ','
                    Write-Host ("         DRY  row={0} {1} | {2} | status={3} | abn={4}" -f `
                        $nextRow, $patient.Name, $patient.Iqama, $patient.Status,
                        $(if ($abn) {$abn} else {'none'})) -ForegroundColor Cyan
                    $totalProcessed++; $coStats.Written++
                } else {
                    Write-TrackerRow -Sheet $trackerSheet -RowIndex $nextRow `
                        -SerialNumber $nextSN -Patient $patient -Cfg $Cfg

                    Write-Host ("         OK   row={0} SN={1}  {2}  ({3})  status={4}" -f `
                        $nextRow, $nextSN, $patient.Name, $patient.Iqama, $patient.Status) -ForegroundColor Green
                    $totalProcessed++; $coStats.Written++

                    if (-not $NoArchive) {
                        try {
                            Move-PatientToArchive -XlsmPath $file.FullName `
                                -ArchiveCompanyDir (Join-Path $ArchiveDir $folderName)
                        } catch {
                            Write-Host "         WARN archive failed: $_" -ForegroundColor Yellow
                        }
                    }
                    $nextRow++; $nextSN++
                }
            } catch {
                Write-Host ("         ERROR $($file.Name): $($_.Exception.Message)") -ForegroundColor Red
                $totalErrors++; $coStats.Errors++
            }
        }
        Write-Progress -Id 1 -Activity "Files in $sheetName" -Completed
    }

    Write-Progress -Id 0 -Activity 'Processing companies' -Completed

    if (-not $DryRun) {
        Write-Host ''
        Write-Host '  Saving tracker...' -ForegroundColor DarkGray
        $beforeSize = (Get-Item $TrackerPath).Length
        $beforeTime = (Get-Item $TrackerPath).LastWriteTime
        try {
            Invoke-WithRetry -Label 'save tracker' -MaxAttempts 8 -InitialDelayMs 500 -Action {
                $Tracker.Save()
            }
            Start-Sleep -Seconds 1
            $afterInfo = Get-Item $TrackerPath
            if ($afterInfo.LastWriteTime -gt $beforeTime -or $afterInfo.Length -ne $beforeSize) {
                Write-Host ("  Tracker saved.  ({0:N0} -> {1:N0} bytes)" -f $beforeSize, $afterInfo.Length) -ForegroundColor Green
            } else {
                # Fallback: SaveAs with macro-enabled format
                Write-Host '  Save() did not update file — trying SaveAs...' -ForegroundColor Yellow
                $Tracker.SaveAs($TrackerPath, 52)
                Start-Sleep -Seconds 1
                $afterInfo2 = Get-Item $TrackerPath
                if ($afterInfo2.LastWriteTime -gt $beforeTime) {
                    Write-Host ("  SaveAs succeeded.  ({0:N0} bytes)" -f $afterInfo2.Length) -ForegroundColor Green
                } else {
                    Write-Host '  ERROR: File not updated even after SaveAs!' -ForegroundColor Red
                    $totalErrors++
                }
            }
        } catch {
            Write-Host "  ERROR: Save failed: $_" -ForegroundColor Red
            $totalErrors++
        }
    } else {
        Write-Host '  DRY-RUN: tracker not modified.' -ForegroundColor Cyan
    }

} catch {
    Write-Host "  FATAL: $_" -ForegroundColor Red
    $totalErrors++
} finally {
    if ($Tracker) {
        try { $Tracker.Close($false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Tracker) } catch { }
    }
    if ($Excel) {
        try { $Excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

$elapsed = '{0:hh\:mm\:ss}' -f ((Get-Date) - $startTime)
Write-Host ''
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host '                          S U M M A R Y                       ' -ForegroundColor Cyan
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''
if ($DryRun) {
    Write-Host '   Mode:                  ' -NoNewline; Write-Host 'DRY-RUN' -ForegroundColor Yellow
} else {
    Write-Host '   Mode:                  ' -NoNewline; Write-Host 'LIVE (tracker updated)' -ForegroundColor Green
}
Write-Host ('   Companies scanned:     {0}' -f $keys.Count)
Write-Host ('   Patient files written: {0}' -f $totalProcessed) -ForegroundColor $(if($totalProcessed -gt 0){'Green'}else{'Gray'})
Write-Host ('   Files skipped:         {0}' -f $totalSkipped)   -ForegroundColor $(if($totalSkipped  -gt 0){'Yellow'}else{'Gray'})
Write-Host ('   Errors:                {0}' -f $totalErrors)    -ForegroundColor $(if($totalErrors   -gt 0){'Red'}else{'Gray'})
Write-Host ('   Time elapsed:          {0}' -f $elapsed)
Write-Host ''

$active = @($perCompanyStats.GetEnumerator() | Where-Object {
    $_.Value.Written -gt 0 -or $_.Value.Skipped -gt 0 -or $_.Value.Errors -gt 0
})
if ($active.Count -gt 0) {
    Write-Host '   Per-company breakdown:' -ForegroundColor White
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f 'Company','Written','Skipped','Errors') -ForegroundColor DarkGray
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f ('-'*22),('-'*8),('-'*8),('-'*8)) -ForegroundColor DarkGray
    foreach ($e in $active) {
        $s = $e.Value
        $c = if ($s.Errors -gt 0){'Red'} elseif($s.Written -gt 0){'Green'} else {'Yellow'}
        Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f $s.Sheet,$s.Written,$s.Skipped,$s.Errors) -ForegroundColor $c
    }
    Write-Host ''
}

Write-Host '   Tracker: ' -NoNewline -ForegroundColor DarkGray
Write-Host $TrackerPath -ForegroundColor Gray
Write-Host '  ============================================================' -ForegroundColor Cyan
if ($totalErrors -gt 0) {
    Write-Host '               FINISHED WITH ERRORS / WARNINGS               ' -ForegroundColor Yellow
} else {
    Write-Host '                       FINISHED SUCCESSFULLY                  ' -ForegroundColor Green
}
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''

if ($totalErrors -gt 0) { exit 1 } else { exit 0 }