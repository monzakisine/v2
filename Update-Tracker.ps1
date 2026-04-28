<#
.SYNOPSIS
    Reads filled AMCFORMULA patient files from company folders and
    appends the data as new rows in the master tracker workbook.
.PARAMETER Company
    Company key as defined in config.psd1, or 'all'.
.PARAMETER DryRun
    Preview only - no changes written, no files moved.
.PARAMETER NoArchive
    Skip moving processed files to Archive.
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
# Retry wrapper - defends against transient file locks on
# network shares (AV scans, sync clients, Excel handle lag).
# ============================================================
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock] $Action,
        [string] $Label = 'operation',
        [int] $MaxAttempts = 6,
        [int] $InitialDelayMs = 250
    )
    $attempt = 0
    $delay = $InitialDelayMs
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
# Convert column letter -> index  (e.g. 'AH' -> 34)
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
# Returns $true if the cell has a non-default fill colour
# (i.e. doctor double-clicked to mark Abnormal = red fill).
# ============================================================
function Test-CellIsHighlighted {
    param($Cell)
    try {
        $idx = $Cell.Interior.ColorIndex
        if ($idx -eq -4142) { return $false }   # xlColorIndexNone
        if ($idx -eq 0) { return $false }   # xlColorIndexAutomatic
        if ($Cell.Interior.Color -eq 16777215) { return $false }   # plain white
        return $true
    }
    catch { return $false }
}

# ============================================================
# Forcibly release an Excel workbook COM object and run GC.
# Must be called before Move-Item on the same file, otherwise
# Excel's lingering handle causes "in use" errors.
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
# Read one filled patient AMCFORMULA file.
# ============================================================
function Read-PatientFile {
    param(
        [Parameter(Mandatory)] $ExcelApp,
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [hashtable] $Cfg
    )

    $wb = Invoke-WithRetry -Label "open '$([IO.Path]::GetFileName($Path))'" -Action {
        $ExcelApp.Workbooks.Open($Path, $false, $true)   # ReadOnly
    }

    try {
        $sheet = $wb.Sheets.Item($Cfg.SourceSheet)

        $data = [ordered]@{
            SourceFile = $Path; Name = $null; Company = $null
            Iqama = $null; Age = $null; DateAMC = $null; DateReview = $null
            BloodPress = $null; Height = $null; Weight = $null
            Status = $null; Comment = $null; Tests = @{}
        }

        foreach ($k in $Cfg.PatientCells.Keys) {
            $data[$k] = $sheet.Range($Cfg.PatientCells[$k]).Value2
        }
        if ($null -ne $data.Iqama) { $data.Iqama = ([string]$data.Iqama).Trim() }
        if ($null -ne $data.Name) { $data.Name = ([string]$data.Name).Trim() }

        # ----------------------------------------------------------
        # Status detection:
        # VBA stores ChrW(10003) = ✓ (U+2713) in exactly one of
        # I4, I5, I6, I7. We read each CheckCell and take the first
        # non-empty one. Any non-empty value = the checkmark is there.
        # ----------------------------------------------------------
        $detectedStatus = $null
        foreach ($s in $Cfg.StatusCandidates) {
            $cv = $sheet.Range($s.CheckCell).Value2
            if ($null -ne $cv -and [string]$cv -ne '') {
                $detectedStatus = $s.Label
                break
            }
        }
        $data.Status = $detectedStatus

        # ----------------------------------------------------------
        # Test results: red fill on column G of the test row = ABNORMAL
        # ----------------------------------------------------------
        foreach ($t in $Cfg.TestRowMap) {
            $row = [int] $t.FormulaRow

            if ($t.ContainsKey('MinAge') -and $null -ne $data.Age) {
                $ageInt = 0
                if ([int]::TryParse([string]$data.Age, [ref]$ageInt)) {
                    if ($ageInt -lt [int]$t.MinAge) {
                        $data.Tests[$t.TrackerCol] = $null
                        continue
                    }
                }
            }

            $gCell = $sheet.Cells.Item($row, 7)
            $data.Tests[$t.TrackerCol] = if (Test-CellIsHighlighted $gCell) { 'ABNORMAL' } else { 'NORMAL' }
        }

        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) } catch { }
        return $data
    }
    finally {
        Release-Workbook $wb
    }
}

# ============================================================
# Write one patient row into the tracker sheet.
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
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.SerialNumber)).Value2 = $SerialNumber
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.DateAMC)).Value2 = $Patient.DateAMC
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.DateReview)).Value2 = $Patient.DateReview
    $iqCell = $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Iqama))
    $iqCell.NumberFormat = '@'
    $iqCell.Value2 = $Patient.Iqama
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Name)).Value2 = $Patient.Name
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Company)).Value2 = $Patient.Company
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Height)).Value2 = $Patient.Height
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Weight)).Value2 = $Patient.Weight
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.BMIFormula)).Formula = '=J{0}/(I{0}/100)^2' -f $RowIndex
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Age)).Value2 = $Patient.Age
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.BloodPress)).Value2 = $Patient.BloodPress
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Status)).Value2 = $Patient.Status
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Comment)).Value2 = $Patient.Comment

    foreach ($col in @($Patient.Tests.Keys)) {
        $val = $Patient.Tests[$col]
        if ($null -ne $val) {
            $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $col)).Value2 = $val
        }
    }
}

# ============================================================
# Find the next empty row (column A = serial number, starts row 2)
# ============================================================
function Get-NextEmptyRow {
    param($Sheet)
    $row = 2
    while ($null -ne $Sheet.Cells.Item($row, 1).Value2 -and
        [string]$Sheet.Cells.Item($row, 1).Value2 -ne '') { $row++ }
    return $row
}

# ============================================================
# Check if an Iqama already exists in column D
# ============================================================
function Test-IqamaExists {
    param($Sheet, [string] $Iqama)
    if ([string]::IsNullOrWhiteSpace($Iqama)) { return $false }
    $row = 2
    while ($null -ne $Sheet.Cells.Item($row, 1).Value2 -and
        [string]$Sheet.Cells.Item($row, 1).Value2 -ne '') {
        if ([string]$Sheet.Cells.Item($row, 4).Value2 -eq $Iqama) { return $true }
        $row++
    }
    return $false
}

# ============================================================
# Move xlsm (and matching pdf) to Archive, with retry.
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
    $ext = [IO.Path]::GetExtension($XlsmPath)
    $dest = Join-Path $ArchiveCompanyDir ([IO.Path]::GetFileName($XlsmPath))
    if (Test-Path $dest) {
        $dest = Join-Path $ArchiveCompanyDir ('{0}_{1:yyyyMMdd-HHmmss}{2}' -f $base, (Get-Date), $ext)
    }
    Invoke-WithRetry -Label "archive $([IO.Path]::GetFileName($XlsmPath))" -MaxAttempts 8 -InitialDelayMs 300 -Action {
        Move-Item -Path $XlsmPath -Destination $dest -Force -ErrorAction Stop
    }
    Write-Host ("  Archived -> $dest") -ForegroundColor DarkGray

    $pdfSrc = Join-Path (Split-Path -Parent $XlsmPath) "$base.pdf"
    if (Test-Path $pdfSrc) {
        $pdfDest = Join-Path $ArchiveCompanyDir "$base.pdf"
        if (Test-Path $pdfDest) {
            $pdfDest = Join-Path $ArchiveCompanyDir ('{0}_{1:yyyyMMdd-HHmmss}.pdf' -f $base, (Get-Date))
        }
        try {
            Invoke-WithRetry -Label "archive $base.pdf" -MaxAttempts 6 -Action {
                Move-Item -Path $pdfSrc -Destination $pdfDest -Force -ErrorAction Stop
            }
        }
        catch {
            Write-Host "  (pdf archive skipped: $_)" -ForegroundColor DarkGray
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

$RootDir = $Cfg.RootDir
$TrackerPath = Join-Path $RootDir $Cfg.TrackerRelPath
$CompaniesDir = Join-Path $RootDir $Cfg.CompaniesDir
$ArchiveDir = Join-Path $RootDir $Cfg.ArchiveDir

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
}
else {
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
    }
    catch {
        Write-Host "  WARNING: Backup failed ($_). Continuing." -ForegroundColor Yellow
    }
}

Write-Host '  Launching Excel...' -ForegroundColor DarkGray
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
$Excel.AskToUpdateLinks = $false
try { $Excel.AutomationSecurity = 3 } catch { }

$Tracker = $null
$totalProcessed = 0
$totalSkipped = 0
$totalErrors = 0
$startTime = Get-Date
$perCompanyStats = [ordered]@{}

try {
    $Tracker = Invoke-WithRetry -Label 'open tracker' -Action {
        $Excel.Workbooks.Open($TrackerPath, $false, $false)
    }

    $companyIdx = 0
    $companyTotal = $keys.Count

    foreach ($key in $keys) {
        $companyIdx++
        $info = $Cfg.Companies[$key]
        $sheetName = $info.Sheet
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

        $sheet = $null
        try { $sheet = $Tracker.Sheets.Item($sheetName) }
        catch {
            Write-Host "         ERROR: Sheet '$sheetName' not found in tracker." -ForegroundColor Red
            $totalErrors += $files.Count
            $coStats.Errors += $files.Count
            continue
        }

        $nextRow = Get-NextEmptyRow $sheet
        $nextSN = if ($nextRow -eq 2) { 1 } else {
            $prev = $sheet.Cells.Item($nextRow - 1, 1).Value2
            if ($prev -as [int]) { [int]$prev + 1 } else { $nextRow - 1 }
        }

        $fileIdx = 0
        $fileTotal = $files.Count
        foreach ($file in $files) {
            $fileIdx++
            Write-Progress -Id 1 -ParentId 0 `
                -Activity "Files in $sheetName" `
                -Status ("$($file.Name)  ({0} of {1})" -f $fileIdx, $fileTotal) `
                -PercentComplete ([int](($fileIdx - 1) / [Math]::Max(1, $fileTotal) * 100))

            try {
                $patient = Read-PatientFile -ExcelApp $Excel -Path $file.FullName -Cfg $Cfg

                if ([string]::IsNullOrWhiteSpace($patient.Iqama)) {
                    Write-Host "         SKIP $($file.Name) - Iqama empty" -ForegroundColor Yellow
                    $totalSkipped++; $coStats.Skipped++
                    continue
                }
                if (-not $patient.Status) {
                    Write-Host "         WARN $($file.Name) - no status checkmark detected" -ForegroundColor Yellow
                }

                $exists = Test-IqamaExists -Sheet $sheet -Iqama $patient.Iqama
                if ($exists -and $Cfg.OnDuplicateIqama -eq 'skip') {
                    Write-Host "         SKIP $($file.Name) - Iqama $($patient.Iqama) already exists" -ForegroundColor Yellow
                    $totalSkipped++; $coStats.Skipped++
                    continue
                }
                if ($exists) {
                    Write-Host "         WARN Iqama $($patient.Iqama) already exists - adding anyway" -ForegroundColor Yellow
                }

                if ($DryRun) {
                    $abn = ($patient.Tests.GetEnumerator() | Where-Object { $_.Value -eq 'ABNORMAL' } |
                        ForEach-Object { $_.Key }) -join ','
                    Write-Host ("         DRY  row={0} {1} | {2} | status={3} | abnormal={4}" -f `
                            $nextRow, $patient.Name, $patient.Iqama, $patient.Status, $(if ($abn) { $abn }else { 'none' })) -ForegroundColor Cyan
                    $totalProcessed++; $coStats.Written++
                }
                else {
                    Write-TrackerRow -Sheet $sheet -RowIndex $nextRow -SerialNumber $nextSN `
                        -Patient $patient -Cfg $Cfg
                    Write-Host ("         OK   row={0} {1} ({2}) status={3}" -f `
                            $nextRow, $patient.Name, $patient.Iqama, $patient.Status) -ForegroundColor Green
                    $totalProcessed++; $coStats.Written++

                    if (-not $NoArchive) {
                        try {
                            Move-PatientToArchive -XlsmPath $file.FullName `
                                -ArchiveCompanyDir (Join-Path $ArchiveDir $folderName)
                        }
                        catch {
                            Write-Host "         WARN archive failed for $($file.Name): $_" -ForegroundColor Yellow
                        }
                    }
                    $nextRow++; $nextSN++
                }
            }
            catch {
                Write-Host "         ERROR $($file.Name): $_" -ForegroundColor Red
                $totalErrors++; $coStats.Errors++
            }
        }
        Write-Progress -Id 1 -Activity "Files in $sheetName" -Completed
    }

    Write-Progress -Id 0 -Activity 'Processing companies' -Completed

    if (-not $DryRun) {
        Write-Host ''
        Write-Host '  Saving tracker...' -ForegroundColor DarkGray
        try {
            Invoke-WithRetry -Label 'save tracker' -MaxAttempts 6 -InitialDelayMs 500 -Action {
                $Tracker.Save()
            }
            Write-Host '  Tracker saved.' -ForegroundColor Green
        }
        catch {
            Write-Host "  ERROR: Tracker save failed: $_" -ForegroundColor Red
            $totalErrors++
        }
    }
    else {
        Write-Host '  DRY-RUN: tracker not modified.' -ForegroundColor Cyan
    }
}
catch {
    Write-Host "  FATAL: $_" -ForegroundColor Red
    $totalErrors++
}
finally {
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
    Write-Host '   Mode:                  ' -NoNewline; Write-Host 'DRY-RUN (no changes written)' -ForegroundColor Yellow
}
else {
    Write-Host '   Mode:                  ' -NoNewline; Write-Host 'LIVE (tracker updated)' -ForegroundColor Green
}
Write-Host ('   Companies scanned:     {0}' -f $keys.Count)
Write-Host ('   Patient files written: {0}' -f $totalProcessed) -ForegroundColor $(if ($totalProcessed -gt 0) { 'Green' }else { 'Gray' })
Write-Host ('   Files skipped:         {0}' -f $totalSkipped)   -ForegroundColor $(if ($totalSkipped -gt 0) { 'Yellow' }else { 'Gray' })
Write-Host ('   Errors:                {0}' -f $totalErrors)    -ForegroundColor $(if ($totalErrors -gt 0) { 'Red' }else { 'Gray' })
Write-Host ('   Time elapsed:          {0}' -f $elapsed)
Write-Host ''

$active = @($perCompanyStats.GetEnumerator() | Where-Object {
        $_.Value.Written -gt 0 -or $_.Value.Skipped -gt 0 -or $_.Value.Errors -gt 0
    })
if ($active.Count -gt 0) {
    Write-Host '   Per-company breakdown:' -ForegroundColor White
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f 'Company', 'Written', 'Skipped', 'Errors') -ForegroundColor DarkGray
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f ('-' * 22), ('-' * 8), ('-' * 8), ('-' * 8)) -ForegroundColor DarkGray
    foreach ($e in $active) {
        $s = $e.Value
        $c = if ($s.Errors -gt 0) { 'Red' } elseif ($s.Written -gt 0) { 'Green' } else { 'Yellow' }
        Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f $s.Sheet, $s.Written, $s.Skipped, $s.Errors) -ForegroundColor $c
    }
    Write-Host ''
}

Write-Host '   Tracker: ' -NoNewline -ForegroundColor DarkGray
Write-Host $TrackerPath -ForegroundColor Gray
Write-Host '  ============================================================' -ForegroundColor Cyan
if ($totalErrors -gt 0) {
    Write-Host '               FINISHED WITH ERRORS / WARNINGS               ' -ForegroundColor Yellow
}
else {
    Write-Host '                       FINISHED SUCCESSFULLY                  ' -ForegroundColor Green
}
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''

if ($totalErrors -gt 0) { exit 1 } else { exit 0 }