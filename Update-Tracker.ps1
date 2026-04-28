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
# PRE-CACHE: Forces SharePoint/OneDrive to download the file
# locally before Excel tries to open it. Without this, the
# Windows "Downloading..." blue bar appears mid-run and can
# cause a brief lock race on the first open attempt.
# ============================================================
function Invoke-PreCache {
    param([string] $Path)
    try {
        # Reading a few bytes forces the sync client to fully
        # download the file to the local cache before we proceed.
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open,
              [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $buf = New-Object byte[] 512
        [void]$fs.Read($buf, 0, $buf.Length)
        $fs.Close()
        $fs.Dispose()
        Start-Sleep -Milliseconds 100
    } catch {
        # If pre-cache fails, just continue - Excel will trigger
        # the download itself (slower but still works)
    }
}
# ============================================================
# SAFE VALUE READER
# Excel COM returns cell values as untyped VARIANTs.
# Large integers (Iqama numbers > Int32.MaxValue) cause
# "Specified cast is not valid" if PowerShell tries to box
# them as Int32. We route everything through Double first,
# then decide what to return.
# ============================================================
function Get-SafeValue {
    param($RawValue)
    if ($null -eq $RawValue) { return $null }
    try {
        # String cells - return as-is
        if ($RawValue -is [string]) { return $RawValue.Trim() }

        # Boolean cells
        if ($RawValue -is [bool]) { return $RawValue }

        # Numeric (including dates stored as doubles, and large iqamas)
        # Force through [double] to avoid Int32 overflow cast errors
        $d = [double]$RawValue
        # Excel error codes are negative numbers like -2146826281
        if ($d -lt -999999) { return $null }
        # If it looks like a whole number, return as string (preserves Iqama precision)
        if ($d -eq [Math]::Floor($d)) {
            # Use [long] not [int] - Iqama numbers exceed Int32
            return ([long]$d).ToString()
        }
        return $d
    } catch {
        # Absolute fallback - just stringify whatever came back
        try { return [string]$RawValue } catch { return $null }
    }
}

# ============================================================
# SAFE INTERIOR COLOR CHECK
# Returns $true if the cell has a non-default fill colour.
# All COM property access is wrapped individually so one bad
# cell doesn't crash the whole file.
# ============================================================
function Test-CellIsHighlighted {
    param($Cell)
    try {
        # Get ColorIndex - use [double] cast to avoid Int32 overflow
        $idx = $null
        try { $idx = [double]$Cell.Interior.ColorIndex } catch { return $false }
        if ($null -eq $idx)    { return $false }
        if ($idx -eq -4142.0)  { return $false }   # xlColorIndexNone
        if ($idx -eq 0.0)      { return $false }   # xlColorIndexAutomatic

        # Check if it's plain white
        $color = $null
        try { $color = [double]$Cell.Interior.Color } catch { return $true }
        if ($null -ne $color -and $color -eq 16777215.0) { return $false }

        return $true
    } catch {
        return $false
    }
}

# ============================================================
# RETRY WRAPPER
# Defends against transient file locks on network shares.
# ============================================================
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock] $Action,
        [string] $Label = 'operation',
        [int] $MaxAttempts = 6,
        [int] $InitialDelayMs = 250
    )
    $attempt = 0
    $delay   = $InitialDelayMs
    $lastErr = $null
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
# Convert column letter to index (e.g. 'AH' -> 34)
# ============================================================
function ConvertTo-ColIndex {
    param([string] $Letter)
    $Letter = $Letter.ToUpper()
    $idx = 0
    foreach ($c in $Letter.ToCharArray()) {
        $idx = $idx * 26 + ([byte][char]$c - [byte][char]'A' + 1)
    }
    return $idx
}

# ============================================================
# Release Excel workbook COM object and run GC.
# Must be done before Move-Item on the same file.
# ============================================================
function Release-Workbook {
    param($Workbook)
    if ($null -ne $Workbook) {
        try { $Workbook.Close($false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    Start-Sleep -Milliseconds 200
}

# ============================================================
# Read one filled patient file
# ============================================================
function Read-PatientFile {
    param(
        [Parameter(Mandatory)] $ExcelApp,
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [hashtable] $Cfg
    )

    # Force local cache before Excel opens (eliminates the SharePoint download bar)
    Invoke-PreCache -Path $Path

    $wb = Invoke-WithRetry -Label "open '$([IO.Path]::GetFileName($Path))'" -Action {
        $ExcelApp.Workbooks.Open($Path, $false, $true)   # UpdateLinks=false, ReadOnly=true
    }

    try {
        $ws = $wb.Sheets.Item($Cfg.SourceSheet)

        $data = [ordered]@{
            SourceFile = $Path
            Name       = $null; Company   = $null; Iqama     = $null
            Age        = $null; DateAMC   = $null; DateReview= $null
            BloodPress = $null; Height    = $null; Weight    = $null
            Status     = $null; Comment   = $null; Tests     = @{}
        }

        # Read patient identity cells
        foreach ($k in $Cfg.PatientCells.Keys) {
            $addr = $Cfg.PatientCells[$k]
            $raw  = $null
            try { $raw = $ws.Range($addr).Value2 } catch { }
            $data[$k] = Get-SafeValue $raw
        }

        # Status: VBA writes ChrW(10003) into exactly one of I4:I7
        $detectedStatus = $null
        foreach ($s in $Cfg.StatusCandidates) {
            $raw = $null
            try { $raw = $ws.Range($s.CheckCell).Value2 } catch { }
            $val = Get-SafeValue $raw
            if ($null -ne $val -and [string]$val -ne '') {
                $detectedStatus = $s.Label
                break
            }
        }
        $data.Status = $detectedStatus

        # Test results: red fill on column G = ABNORMAL
        foreach ($t in $Cfg.TestRowMap) {
            $row = [int] $t.FormulaRow

            # PSA only for patients >= MinAge
            if ($t.ContainsKey('MinAge') -and $null -ne $data.Age) {
                $ageVal = 0L
                try { $ageVal = [long]$data.Age } catch { }
                if ($ageVal -lt [long]$t.MinAge) {
                    $data.Tests[$t.TrackerCol] = $null
                    continue
                }
            }

            $gCell = $null
            try { $gCell = $ws.Cells.Item($row, 7) } catch { }
            $data.Tests[$t.TrackerCol] = if ($null -ne $gCell -and (Test-CellIsHighlighted $gCell)) {
                'ABNORMAL'
            } else {
                'NORMAL'
            }
        }

        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch { }
        return $data

    } finally {
        Release-Workbook $wb
    }
}

# ============================================================
# Write one patient row into the tracker sheet
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

    # Helper: write to a cell safely (never throws on a single cell)
    $w = {
        param($col, $val)
        try { $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $col)).Value2 = $val } catch { }
    }

    & $w $fc.SerialNumber $SerialNumber
    & $w $fc.DateAMC      $Patient.DateAMC
    & $w $fc.DateReview   $Patient.DateReview
    & $w $fc.Name         $Patient.Name
    & $w $fc.Company      $Patient.Company
    & $w $fc.Height       $Patient.Height
    & $w $fc.Weight       $Patient.Weight
    & $w $fc.Age          $Patient.Age
    & $w $fc.BloodPress   $Patient.BloodPress
    & $w $fc.Status       $Patient.Status
    & $w $fc.Comment      $Patient.Comment

    # Iqama must be stored as TEXT to keep all 10 digits
    try {
        $iqCell = $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Iqama))
        $iqCell.NumberFormat = '@'
        $iqCell.Value2 = $Patient.Iqama
    } catch { }

    # BMI formula
    try {
        $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.BMIFormula)).Formula = `
            ('=J{0}/(I{0}/100)^2' -f $RowIndex)
    } catch { }

    # Test result columns
    foreach ($col in @($Patient.Tests.Keys)) {
        $val = $Patient.Tests[$col]
        if ($null -ne $val) {
            try { $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $col)).Value2 = $val } catch { }
        }
    }
}

# ============================================================
# Find next empty row (column A = serial number, starts row 2)
# ============================================================
function Get-NextEmptyRow {
    param($Sheet)
    $row = 2
    while ($true) {
        $raw = $null
        try { $raw = $Sheet.Cells.Item($row, 1).Value2 } catch { break }
        if ($null -eq $raw) { break }
        $str = ''
        try { $str = [string]$raw } catch { }
        if ($str.Trim() -eq '') { break }
        $row++
    }
    return $row
}

# ============================================================
# Check if Iqama already exists in column D
# ============================================================
function Test-IqamaExists {
    param($Sheet, [string] $Iqama)
    if ([string]::IsNullOrWhiteSpace($Iqama)) { return $false }
    $row = 2
    while ($true) {
        $snRaw = $null
        try { $snRaw = $Sheet.Cells.Item($row, 1).Value2 } catch { break }
        if ($null -eq $snRaw) { break }
        $snStr = ''
        try { $snStr = [string]$snRaw } catch { }
        if ($snStr.Trim() -eq '') { break }

        $iqRaw = $null
        try { $iqRaw = $Sheet.Cells.Item($row, 4).Value2 } catch { }
        $existing = ''
        try { $existing = [string](Get-SafeValue $iqRaw) } catch { }
        if ($existing -eq $Iqama) { return $true }
        $row++
    }
    return $false
}

# ============================================================
# Move xlsm (+ matching pdf) to Archive with retry
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

    # COPY first (doesn't need exclusive lock, works even if SharePoint sync
    # client is holding the file), then delete the original.
    # If the delete fails we leave the original in place - the tracker row is
    # already written, and the Iqama check will skip it on the next run.
    Invoke-WithRetry -Label "copy to archive $([IO.Path]::GetFileName($XlsmPath))" -MaxAttempts 8 -InitialDelayMs 400 -Action {
        Copy-Item -Path $XlsmPath -Destination $dest -Force -ErrorAction Stop
    }

    # Verify the copy is complete before deleting the original
    $srcSize  = (Get-Item $XlsmPath).Length
    $destSize = (Get-Item $dest).Length
    if ($destSize -eq $srcSize -and $destSize -gt 0) {
        try {
            Invoke-WithRetry -Label "delete original $([IO.Path]::GetFileName($XlsmPath))" -MaxAttempts 5 -InitialDelayMs 500 -Action {
                Remove-Item -Path $XlsmPath -Force -ErrorAction Stop
            }
            Write-Host ("    Archived -> $dest") -ForegroundColor DarkGray
        } catch {
            # Copy succeeded, delete failed - file stays in company folder.
            # Next run will detect the Iqama as duplicate and skip it.
            Write-Host ("    Copied to archive (original stays - will be skipped next run): $dest") -ForegroundColor DarkGray
        }
    } else {
        # Copy looks incomplete - remove the bad copy, leave original
        try { Remove-Item $dest -Force -ErrorAction SilentlyContinue } catch { }
        throw "Archive copy size mismatch (src=$srcSize dest=$destSize). Original left in place."
    }

    # Handle matching PDF the same way
    $pdfSrc = Join-Path (Split-Path -Parent $XlsmPath) "$base.pdf"
    if (Test-Path $pdfSrc) {
        $pdfDest = Join-Path $ArchiveCompanyDir "$base.pdf"
        if (Test-Path $pdfDest) {
            $pdfDest = Join-Path $ArchiveCompanyDir ('{0}_{1:yyyyMMdd-HHmmss}.pdf' -f $base, (Get-Date))
        }
        try {
            Invoke-WithRetry -Label "copy pdf $base.pdf" -MaxAttempts 6 -Action {
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
# Pre-flight: is the tracker already open in another Excel instance?
# If yes, our Save() will silently go to a conflict copy.
$lockProbe = $TrackerPath + '.amclock'
$isLocked = $false
try {
    [System.IO.File]::Move($TrackerPath, $lockProbe)
    [System.IO.File]::Move($lockProbe, $TrackerPath)
} catch {
    $isLocked = $true
}
if ($isLocked) {
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor Red
    Write-Host '  ERROR: The tracker file is currently LOCKED.' -ForegroundColor Red
    Write-Host '  ============================================================' -ForegroundColor Red
    Write-Host '  Someone (probably you, or another nurse) has the tracker' -ForegroundColor Red
    Write-Host '  open in Excel right now. Any changes the script makes will' -ForegroundColor Red
    Write-Host '  be lost or saved to a conflict copy.' -ForegroundColor Red
    Write-Host ''
    Write-Host '  -> Close the tracker in Excel everywhere, then try again.' -ForegroundColor Yellow
    Write-Host ''
    exit 1
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
$Excel.Visible          = $false
$Excel.DisplayAlerts    = $false
$Excel.AskToUpdateLinks = $false
try { $Excel.AutomationSecurity = 3 } catch { }   # msoAutomationSecurityForceDisable

$Tracker         = $null
$totalProcessed  = 0
$totalSkipped    = 0
$totalErrors     = 0
$startTime       = Get-Date
$perCompanyStats = [ordered]@{}

try {
    # Notify=$false stops Excel from popping a "file in use" dialog and
    # silently dropping us to read-only mode.
    $Tracker = Invoke-WithRetry -Label 'open tracker' -Action {
        $Excel.Workbooks.Open(
            $TrackerPath,   # Filename
            $false,         # UpdateLinks
            $false,         # ReadOnly
            [Type]::Missing,# Format
            [Type]::Missing,# Password
            [Type]::Missing,# WriteResPassword
            $true,          # IgnoreReadOnlyRecommended
            [Type]::Missing,# Origin
            [Type]::Missing,# Delimiter
            $false          # Editable
        )
    }

    # CRITICAL: confirm we got a WRITABLE workbook, not a read-only copy
    if ($Tracker.ReadOnly) {
        throw "Tracker opened in READ-ONLY mode (another Excel instance has it). Close it everywhere and retry."
    }

    # CRITICAL: disable SharePoint AutoSave for our session.
    # AutoSave intercepts Save() and can route it to a sync queue
    # instead of immediately writing the file.
    try { $Tracker.AutoSaveOn = $false } catch { }

    # Diagnostic: show exactly which file Excel actually opened.
    Write-Host ("  Opened: {0}" -f $Tracker.FullName) -ForegroundColor DarkGray
    Write-Host ("  Read-only: {0}   AutoSave: {1}" -f $Tracker.ReadOnly, `
        $(try { $Tracker.AutoSaveOn } catch { 'n/a' })) -ForegroundColor DarkGray

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
            -Status ("$sheetName ({0} of {1})" -f $companyIdx, $companyTotal) `
            -PercentComplete ([int](($companyIdx - 1) / [Math]::Max(1, $companyTotal) * 100))

        Write-Host ("  [{0}/{1}] {2}" -f $companyIdx, $companyTotal, $sheetName) -ForegroundColor White

        if (-not (Test-Path $folderPath)) {
            try { New-Item -ItemType Directory -Path $folderPath -Force | Out-Null } catch { }
            Write-Host "         Folder created (was missing)." -ForegroundColor DarkGray
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
            $totalErrors += $files.Count
            $coStats.Errors += $files.Count
            continue
        }

        $nextRow = Get-NextEmptyRow $trackerSheet
        $prevSN  = $null
        try { $prevSN = $trackerSheet.Cells.Item($nextRow - 1, 1).Value2 } catch { }
        $nextSN = if ($nextRow -eq 2) {
            1
        } elseif ($null -ne $prevSN -and ($prevSN -as [double]) -ne $null) {
            [int][double]$prevSN + 1
        } else {
            $nextRow - 1
        }

        $fileIdx   = 0
        $fileTotal = $files.Count

        foreach ($file in $files) {
            $fileIdx++
            Write-Progress -Id 1 -ParentId 0 `
                -Activity "Files in $sheetName" `
                -Status ("$($file.Name)  ($fileIdx of $fileTotal)") `
                -PercentComplete ([int](($fileIdx - 1) / [Math]::Max(1, $fileTotal) * 100))

            try {
                $patient = Read-PatientFile -ExcelApp $Excel -Path $file.FullName -Cfg $Cfg

                if ([string]::IsNullOrWhiteSpace($patient.Iqama)) {
                    Write-Host "         SKIP $($file.Name) - Iqama empty" -ForegroundColor Yellow
                    $totalSkipped++; $coStats.Skipped++
                    continue
                }
                if (-not $patient.Status) {
                    Write-Host "         WARN $($file.Name) - no status checkmark" -ForegroundColor Yellow
                }

                $exists = Test-IqamaExists -Sheet $trackerSheet -Iqama $patient.Iqama
                if ($exists -and $Cfg.OnDuplicateIqama -eq 'skip') {
                    Write-Host "         SKIP $($file.Name) - Iqama $($patient.Iqama) already exists" -ForegroundColor Yellow
                    $totalSkipped++; $coStats.Skipped++
                    continue
                }
                if ($exists) {
                    Write-Host "         WARN Iqama $($patient.Iqama) already exists - adding anyway" -ForegroundColor Yellow
                }

                if ($DryRun) {
                    $abn = ($patient.Tests.GetEnumerator() |
                            Where-Object { $_.Value -eq 'ABNORMAL' } |
                            ForEach-Object { $_.Key }) -join ','
                    Write-Host ("         DRY  row={0}  {1}  |  {2}  |  status={3}  |  abnormal={4}" -f `
                        $nextRow, $patient.Name, $patient.Iqama, $patient.Status,
                        $(if ($abn) { $abn } else { 'none' })) -ForegroundColor Cyan
                    $totalProcessed++; $coStats.Written++
                } else {
                    Write-TrackerRow -Sheet $trackerSheet -RowIndex $nextRow `
                        -SerialNumber $nextSN -Patient $patient -Cfg $Cfg
                    Write-Host ("         OK   row={0}  {1}  ({2})  status={3}" -f `
                        $nextRow, $patient.Name, $patient.Iqama, $patient.Status) -ForegroundColor Green
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

        # Capture file size + timestamp BEFORE save
        $beforeInfo = Get-Item $TrackerPath
        $beforeTime = $beforeInfo.LastWriteTime
        $beforeSize = $beforeInfo.Length

        try {
            # Force xlOpenXMLWorkbookMacroEnabled (52) so macro-enabled .xlsm
            # is preserved and we don't accidentally save to a converted format.
            Invoke-WithRetry -Label 'save tracker' -MaxAttempts 6 -InitialDelayMs 500 -Action {
                $Tracker.Save()
            }

            # Wait briefly for SharePoint sync to commit
            Start-Sleep -Seconds 1

            # VERIFY the file was actually changed on disk
            $afterInfo = Get-Item $TrackerPath
            if ($afterInfo.LastWriteTime -gt $beforeTime -or $afterInfo.Length -ne $beforeSize) {
                Write-Host ("  Tracker saved.  ({0:N0} bytes -> {1:N0} bytes, modified {2:HH:mm:ss})" -f `
                    $beforeSize, $afterInfo.Length, $afterInfo.LastWriteTime) -ForegroundColor Green
            } else {
                Write-Host '' -ForegroundColor Yellow
                Write-Host '  WARNING: Save() returned but the file timestamp did NOT change!' -ForegroundColor Yellow
                Write-Host '  This usually means SharePoint AutoSave intercepted the write.' -ForegroundColor Yellow
                Write-Host '  Trying SaveAs as a fallback...' -ForegroundColor Yellow

                # Fallback: explicit SaveAs with the macro-enabled format (52)
                try {
                    $Tracker.SaveAs($TrackerPath, 52)
                    Start-Sleep -Seconds 1
                    $afterInfo2 = Get-Item $TrackerPath
                    if ($afterInfo2.LastWriteTime -gt $beforeTime) {
                        Write-Host ("  SaveAs succeeded.  ({0:N0} bytes, modified {1:HH:mm:ss})" -f `
                            $afterInfo2.Length, $afterInfo2.LastWriteTime) -ForegroundColor Green
                    } else {
                        Write-Host '  ERROR: Even SaveAs did not update the file!' -ForegroundColor Red
                        Write-Host '  Check that you have write permission to the tracker folder.' -ForegroundColor Red
                        $totalErrors++
                    }
                } catch {
                    Write-Host "  ERROR: SaveAs also failed: $_" -ForegroundColor Red
                    $totalErrors++
                }
            }
        } catch {
            Write-Host "  ERROR: Tracker save failed: $_" -ForegroundColor Red
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
    Write-Host '   Mode:                  ' -NoNewline
    Write-Host 'DRY-RUN (no changes written)' -ForegroundColor Yellow
} else {
    Write-Host '   Mode:                  ' -NoNewline
    Write-Host 'LIVE (tracker updated)' -ForegroundColor Green
}
Write-Host ('   Companies scanned:     {0}' -f $keys.Count)
Write-Host ('   Patient files written: {0}' -f $totalProcessed) `
    -ForegroundColor $(if ($totalProcessed -gt 0) { 'Green' } else { 'Gray' })
Write-Host ('   Files skipped:         {0}' -f $totalSkipped) `
    -ForegroundColor $(if ($totalSkipped -gt 0) { 'Yellow' } else { 'Gray' })
Write-Host ('   Errors:                {0}' -f $totalErrors) `
    -ForegroundColor $(if ($totalErrors -gt 0) { 'Red' } else { 'Gray' })
Write-Host ('   Time elapsed:          {0}' -f $elapsed)
Write-Host ''

$active = @($perCompanyStats.GetEnumerator() | Where-Object {
    $_.Value.Written -gt 0 -or $_.Value.Skipped -gt 0 -or $_.Value.Errors -gt 0
})
if ($active.Count -gt 0) {
    Write-Host '   Per-company breakdown:' -ForegroundColor White
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f 'Company','Written','Skipped','Errors') `
        -ForegroundColor DarkGray
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f ('-'*22),('-'*8),('-'*8),('-'*8)) `
        -ForegroundColor DarkGray
    foreach ($e in $active) {
        $s = $e.Value
        $c = if ($s.Errors -gt 0) { 'Red' } elseif ($s.Written -gt 0) { 'Green' } else { 'Yellow' }
        Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f $s.Sheet,$s.Written,$s.Skipped,$s.Errors) `
            -ForegroundColor $c
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