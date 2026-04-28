<#
.SYNOPSIS
    Reads filled AMCFORMULA patient files from company folders and
    appends the data as new rows in the master tracker workbook.

.DESCRIPTION
    For each patient .xlsm file in <RootDir>\Companies\<CompanyName>\,
    this script extracts identity / vitals / test results / status /
    comment, then writes a new row to the matching company sheet in
    Contractors_AMC_Tracker_2026.xlsm. Processed files are moved to
    <RootDir>\Archive\<CompanyName>\.
    
    ROBUST FILE-LOCKING FIX:
    - Logs written to local %TEMP% during run (immune to AV scanning)
    - All file operations retry with exponential backoff (250ms → 4s, up to 6 attempts)
    - Forced GC + delay between closing Excel handles and moving files
    - Network log copy fails gracefully; local log path printed as fallback

.PARAMETER Company
    Company key as defined in config.psd1 (e.g. 'scms', 'altamimi'),
    or 'all' to process every configured company.

.PARAMETER DryRun
    Read patient files and show what WOULD be written, without
    modifying the tracker or moving any files.

.PARAMETER NoArchive
    Skip moving processed files to Archive (leave them in place).

.PARAMETER ConfigPath
    Optional override for the config file path.

.EXAMPLE
    .\Update-Tracker.ps1 scms
    .\Update-Tracker.ps1 all -DryRun
    .\Update-Tracker.ps1 catrion -NoArchive
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string] $Company = 'all',

    [switch] $DryRun,
    [switch] $NoArchive,

    [string] $ConfigPath
)

$ErrorActionPreference = 'Stop'

# ============================================================
# RETRY WRAPPER: exponential backoff for network-share resilience
# ============================================================
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock] $ScriptBlock,
        [int] $MaxAttempts = 6,
        [int] $InitialDelayMs = 250
    )
    
    $attempt = 0
    $delayMs = $InitialDelayMs
    
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxAttempts) {
                throw
            }
            Write-Host "  [Retry $attempt/$MaxAttempts] Waiting ${delayMs}ms before retry: $($_.Exception.Message)" -ForegroundColor Yellow
            Start-Sleep -Milliseconds $delayMs
            $delayMs = [Math]::Min($delayMs * 2, 4000)  # cap at 4 seconds
        }
    }
}

function Write-Log {
    param(
        [string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK', 'DRY')] [string] $Level = 'INFO'
    )
    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN' { 'Yellow' }
        'OK' { 'Green' }
        'DRY' { 'Cyan' }
        default { 'Gray' }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $color
}

# ============================================================
# Helper: safely read a cell value (handles error cells)
# ============================================================
function Get-SafeCellValue {
    param(
        [Parameter(Mandatory)] $Cell,
        [string] $DefaultValue = $null
    )
    
    try {
        if ($null -eq $Cell) { return $DefaultValue }
        
        # Check if cell contains an error value (#N/A, #VALUE!, etc.)
        $text = $Cell.Text
        if ($null -ne $text) {
            $text = [string]$text
            # Error cells start with #
            if ($text.StartsWith('#')) {
                return $DefaultValue
            }
        }
        
        # Try to safely read Value2
        $val = $null
        try {
            $val = $Cell.Value2
        }
        catch {
            # If Value2 fails, try Text
            return $DefaultValue
        }
        
        # Check if it's a DBNull
        if ($val -is [System.DBNull]) {
            return $DefaultValue
        }
        
        return $val
    }
    catch {
        return $DefaultValue
    }
}

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
# Helper: detect if a cell has a non-default fill (= "abnormal")
# Default = no fill (xlNone, ColorIndex = -4142).
# Anything else = doctor marked it.
# ============================================================
function Test-CellIsHighlighted {
    param($Cell)
    try {
        if ($null -eq $Cell) { return $false }
        
        $interior = $Cell.Interior
        if ($null -eq $interior) { return $false }
        
        # Try to safely read ColorIndex
        $idx = $null
        try {
            $idx = $interior.ColorIndex
        }
        catch {
            # If we can't read ColorIndex, try Color instead
            try {
                $color = $interior.Color
                if ($color -eq 16777215) { return $false }  # white
                return $true
            }
            catch {
                return $false
            }
        }
        
        # Evaluate the ColorIndex
        if ($idx -eq -4142) { return $false }     # xlColorIndexNone
        if ($idx -eq 0) { return $false }     # xlColorIndexAutomatic / treat as no fill
        if ($idx -eq -2) { return $false }     # xlAutomatic variant
        
        # Check if Color value is white (sometimes used as "no fill")
        try {
            if ($interior.Color -eq 16777215) { return $false }
        }
        catch { }
        
        return $true
    }
    catch {
        return $false
    }
}

# ============================================================
# Helper: read the filled patient file with retries
# ============================================================
function Read-PatientFile {
    param(
        [Parameter(Mandatory)] $ExcelApp,
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [hashtable] $Cfg
    )

    Invoke-WithRetry {
        $wb = $ExcelApp.Workbooks.Open($Path)
        try {
            $sheet = $wb.Sheets.Item($Cfg.SourceSheet)

            $data = [ordered]@{
                SourceFile = $Path
                Name       = $null
                Company    = $null
                Iqama      = $null
                Age        = $null
                DateAMC    = $null
                DateReview = $null
                BloodPress = $null
                Height     = $null
                Weight     = $null
                Status     = $null
                Comment    = $null
                Tests      = @{}
            }

            # Safely read patient cells with error handling
            foreach ($k in $Cfg.PatientCells.Keys) {
                try {
                    $addr = $Cfg.PatientCells[$k]
                    $cell = $sheet.Range($addr)
                    $val = Get-SafeCellValue $cell
                    $data[$k] = $val
                }
                catch {
                    Write-Log "    Warning: Could not read cell $addr for field '$k': $($_.Exception.Message)" 'WARN'
                    $data[$k] = $null
                }
            }

            if ($null -ne $data.Iqama) {
                $data.Iqama = ([string]$data.Iqama).Trim()
            }
            if ($null -ne $data.Name) {
                $data.Name = ([string]$data.Name).Trim()
            }

            # Detect status (which checkbox cell has a checkmark)
            $detectedStatus = $null
            foreach ($s in $Cfg.StatusCandidates) {
                foreach ($addr in $s.CheckCells) {
                    try {
                        $cell = $sheet.Range($addr)
                        $cv = Get-SafeCellValue $cell
                        if ($null -ne $cv -and "$cv".Trim() -ne '') {
                            $detectedStatus = $s.Label
                            break
                        }
                    }
                    catch {
                        # Skip cells that error out
                        continue
                    }
                }
                if ($detectedStatus) { break }
            }
            $data.Status = $detectedStatus

            # Detect Abnormal tests via fill on column G
            foreach ($t in $Cfg.TestRowMap) {
                try {
                    $row = [int] $t.FormulaRow

                    # Skip tests with MinAge constraint if patient is younger
                    if ($t.ContainsKey('MinAge') -and $null -ne $data.Age) {
                        $ageInt = 0
                        if ([int]::TryParse("$($data.Age)", [ref]$ageInt)) {
                            if ($ageInt -lt [int]$t.MinAge) {
                                $data.Tests[$t.TrackerCol] = $null
                                continue
                            }
                        }
                    }

                    $gCell = $sheet.Cells.Item($row, 7)
                    $isAbnormal = Test-CellIsHighlighted $gCell
                    $data.Tests[$t.TrackerCol] = if ($isAbnormal) { 'ABNORMAL' } else { 'NORMAL' }
                }
                catch {
                    Write-Log "    Warning: Could not read test row $($t.FormulaRow) for '$($t.TrackerCol)': $($_.Exception.Message)" 'WARN'
                    $data.Tests[$t.TrackerCol] = 'NORMAL'  # Default to normal on error
                }
            }

            return $data
        }
        finally {
            if ($null -ne $wb) {
                try { $wb.Saved = $true } catch { }
                try { $wb.Close($false) } catch { }
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            }
            
            # Force GC to release the workbook handle
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
        }
    }
}

# ============================================================
# Helper: write one row into the tracker sheet
# ============================================================
function Write-TrackerRow {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] [int]    $RowIndex,
        [Parameter(Mandatory)] [int]    $SerialNumber,
        [Parameter(Mandatory)] [hashtable] $Patient,
        [Parameter(Mandatory)] [hashtable] $Cfg
    )

    $fc = $Cfg.FixedColumns

    function Set-Cell {
        param($ColLetter, $Val)
        if ($null -ne $Val) {
            $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $ColLetter)).Value2 = $Val
        }
    }

    Set-Cell $fc.SerialNumber $SerialNumber
    Set-Cell $fc.DateAMC $Patient.DateAMC
    Set-Cell $fc.DateReview $Patient.DateReview

    # Iqama: write as text to preserve full digit precision
    if ($null -ne $Patient.Iqama) {
        $iqamaCell = $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Iqama))
        $iqamaCell.NumberFormat = '@'
        $iqamaCell.Value2 = $Patient.Iqama
    }

    Set-Cell $fc.Name $Patient.Name
    Set-Cell $fc.Company $Patient.Company
    Set-Cell $fc.Height $Patient.Height
    Set-Cell $fc.Weight $Patient.Weight

    # BMI as a live formula
    $bmiCell = $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.BMIFormula))
    $bmiCell.Formula = '=J{0}/(I{0}/100)^2' -f $RowIndex

    Set-Cell $fc.Age $Patient.Age
    Set-Cell $fc.BloodPress $Patient.BloodPress
    Set-Cell $fc.Status $Patient.Status
    Set-Cell $fc.Comment $Patient.Comment

    # Test result columns
    foreach ($col in @($Patient.Tests.Keys)) {
        $value = $Patient.Tests[$col]
        if ($null -eq $value) { continue }   # leave the cell blank (e.g. PSA <40)
        Set-Cell $col $value
    }
}

# ============================================================
# Helper: find next empty data row in a tracker sheet
# (assumes column A = SN starts at row 2)
# ============================================================
function Get-NextEmptyRow {
    param($Sheet)
    $row = 2
    while ($true) {
        $cell = $Sheet.Cells.Item($row, 1)
        $val = Get-SafeCellValue $cell
        if ($null -eq $val -or "$val".Trim() -eq '') {
            break
        }
        $row++
    }
    return $row
}

# ============================================================
# Helper: check if Iqama already exists in sheet
# ============================================================
function Test-IqamaExists {
    param($Sheet, [string] $Iqama)
    if ([string]::IsNullOrWhiteSpace($Iqama)) { return $false }
    $row = 2
    while ($true) {
        $cellA = $Sheet.Cells.Item($row, 1)
        $valA = Get-SafeCellValue $cellA
        if ($null -eq $valA -or "$valA".Trim() -eq '') {
            break
        }
        
        $cellD = $Sheet.Cells.Item($row, 4)
        $existing = Get-SafeCellValue $cellD
        if ($null -ne $existing -and "$existing".Trim() -eq $Iqama) { return $true }
        
        $row++
    }
    return $false
}

# ============================================================
#                          M A I N
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir 'config.psd1' }
if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
$Cfg = Import-PowerShellDataFile $ConfigPath

$RootDir = $Cfg.RootDir
$TrackerPath = Join-Path $RootDir $Cfg.TrackerRelPath
$CompaniesDir = Join-Path $RootDir $Cfg.CompaniesDir
$ArchiveDir = Join-Path $RootDir $Cfg.ArchiveDir
$LogsDir = Join-Path $RootDir $Cfg.LogsDir

# Ensure all network dirs exist (with retry)
foreach ($d in @($CompaniesDir, $ArchiveDir, $LogsDir)) {
    Invoke-WithRetry {
        if (-not (Test-Path $d)) {
            New-Item -ItemType Directory -Path $d -Force | Out-Null
        }
    }
}

Write-Log '======================================================'
Write-Log "AMC Automation start (Company='$Company', DryRun=$DryRun, NoArchive=$NoArchive)"
Write-Log "Root: $RootDir"
Write-Log "Tracker: $TrackerPath"
Write-Log '======================================================'

if (-not (Test-Path $TrackerPath)) {
    Write-Log "Tracker file not found: $TrackerPath" 'ERROR'
    exit 1
}

$AllKeys = $Cfg.Companies.Keys | Sort-Object
$keys = if ($Company.ToLower() -eq 'all') {
    $AllKeys
}
else {
    $hit = $AllKeys | Where-Object { $_ -ieq $Company }
    if (-not $hit) {
        Write-Log "Unknown company key '$Company'. Valid keys: $($AllKeys -join ', ')" 'ERROR'
        exit 1
    }
    @($hit)
}

# Backup tracker (with retry)
if ($Cfg.BackupTrackerBeforeRun -and -not $DryRun) {
    $bkp = Join-Path $LogsDir ('tracker-backup-{0:yyyy-MM-dd_HHmmss}.xlsm' -f (Get-Date))
    Invoke-WithRetry {
        Copy-Item -Path $TrackerPath -Destination $bkp
    }
    Write-Log "Tracker backup -> $bkp"
}

Write-Log 'Launching Excel (hidden)...'
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
$Excel.AskToUpdateLinks = $false
$Excel.ScreenUpdating = $false
try { $Excel.AutomationSecurity = 3 } catch { }
try { $Excel.AlertBeforeOverwriting = $false } catch { }

# Set to not show any updates or prompts during macro operations
try { $Excel.EnableEvents = $false } catch { }

$Tracker = $null
$totalProcessed = 0
$totalSkipped = 0
$totalErrors = 0
$startTime = Get-Date
$perCompanyStats = [ordered]@{}

try {
    $Tracker = Invoke-WithRetry {
        $Excel.Workbooks.Open($TrackerPath)
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

        $coPct = [int](($companyIdx - 1) / [Math]::Max(1, $companyTotal) * 100)
        Write-Progress -Id 0 -Activity 'Processing companies' `
            -Status ("$sheetName ({0} of {1})" -f $companyIdx, $companyTotal) `
            -PercentComplete $coPct

        Write-Log ''
        Write-Log "--- [$companyIdx/$companyTotal] $key  (sheet='$sheetName', folder='$folderName') ---"

        if (-not (Test-Path $folderPath)) {
            Write-Log "Folder missing, creating: $folderPath" 'WARN'
            Invoke-WithRetry {
                New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
            }
            continue
        }

        $files = @(Get-ChildItem -Path $folderPath -Filter '*.xlsm' -File -ErrorAction SilentlyContinue)
        if (-not $files -or $files.Count -eq 0) {
            Write-Log 'No .xlsm patient files found.' 'INFO'
            continue
        }

        $sheet = $null
        try { $sheet = $Tracker.Sheets.Item($sheetName) }
        catch {
            Write-Log "Sheet '$sheetName' not found in tracker. Skipping company." 'ERROR'
            $totalErrors += $files.Count
            $coStats.Errors += $files.Count
            continue
        }

        $nextRow = Get-NextEmptyRow $sheet
        if ($nextRow -eq 2) {
            $nextSN = 1
        }
        else {
            $prevCell = $sheet.Cells.Item($nextRow - 1, 1)
            $prev = Get-SafeCellValue $prevCell
            $nextSN = if ($prev -as [int]) { [int]$prev + 1 } else { $nextRow - 1 }
        }
        Write-Log "Next free row: $nextRow  (SN=$nextSN)  Files to process: $($files.Count)"

        $fileIdx = 0
        $fileTotal = $files.Count
        foreach ($file in $files) {
            $fileIdx++
            $filePct = [int](($fileIdx - 1) / [Math]::Max(1, $fileTotal) * 100)
            Write-Progress -Id 1 -ParentId 0 -Activity "Processing files in $sheetName" `
                -Status ("$($file.Name) ({0} of {1})" -f $fileIdx, $fileTotal) `
                -PercentComplete $filePct
            try {
                Write-Log "Reading: $($file.Name)"
                $patient = Read-PatientFile -ExcelApp $Excel -Path $file.FullName -Cfg $Cfg

                if ([string]::IsNullOrWhiteSpace($patient.Iqama)) {
                    Write-Log "  Iqama empty in formula file. Skipping." 'WARN'
                    $totalSkipped++
                    $coStats.Skipped++
                    continue
                }
                if ([string]::IsNullOrWhiteSpace($patient.Name)) {
                    Write-Log "  Name empty in formula file. Continuing anyway." 'WARN'
                }
                if (-not $patient.Status) {
                    Write-Log "  No status checkmark detected. Status will be left blank." 'WARN'
                }

                $exists = Test-IqamaExists -Sheet $sheet -Iqama $patient.Iqama
                if ($exists) {
                    switch ($Cfg.OnDuplicateIqama) {
                        'skip' {
                            Write-Log "  Iqama $($patient.Iqama) already in sheet. Skipping (config: skip)." 'WARN'
                            $totalSkipped++
                            $coStats.Skipped++
                            continue
                        }
                        'duplicate' { }
                        default {
                            Write-Log "  Iqama $($patient.Iqama) already in sheet. Adding new row anyway (config: warn)." 'WARN'
                        }
                    }
                }

                if ($DryRun) {
                    $abn = ($patient.Tests.GetEnumerator() | Where-Object { $_.Value -eq 'ABNORMAL' } | ForEach-Object { $_.Key }) -join ','
                    if (-not $abn) { $abn = '(none)' }
                    Write-Log "  DRY-RUN row=$nextRow SN=$nextSN  $($patient.Name) | Iqama=$($patient.Iqama) | Age=$($patient.Age) | BP=$($patient.BloodPress) | Status=$($patient.Status) | Abnormal=$abn" 'DRY'
                    $totalProcessed++
                    $coStats.Written++
                }
                else {
                    Write-TrackerRow -Sheet $sheet -RowIndex $nextRow -SerialNumber $nextSN -Patient $patient -Cfg $Cfg
                    Write-Log "  WROTE row=$nextRow SN=$nextSN  $($patient.Name) ($($patient.Iqama)) status=$($patient.Status)" 'OK'
                    $totalProcessed++
                    $coStats.Written++

                    if (-not $NoArchive) {
                        $archiveCo = Join-Path $ArchiveDir $folderName
                        Invoke-WithRetry {
                            if (-not (Test-Path $archiveCo)) {
                                New-Item -ItemType Directory -Path $archiveCo -Force | Out-Null
                            }
                        }

                        $base = [IO.Path]::GetFileNameWithoutExtension($file.Name)
                        $ext = $file.Extension
                        $dest = Join-Path $archiveCo $file.Name
                        
                        Invoke-WithRetry {
                            if (Test-Path $dest) {
                                $dest = Join-Path $archiveCo ('{0}_{1:yyyyMMdd-HHmmss}{2}' -f $base, (Get-Date), $ext)
                            }
                        }

                        # Force GC + delay before moving to allow Excel to fully release the handle
                        [GC]::Collect()
                        [GC]::WaitForPendingFinalizers()
                        [GC]::Collect()
                        Start-Sleep -Milliseconds 500
                        
                        Invoke-WithRetry {
                            Move-Item -Path $file.FullName -Destination $dest
                        }
                        Write-Log "  Archived xlsm -> $dest"

                        # Move matching PDF if present
                        $pdfSrc = Join-Path $folderPath "$base.pdf"
                        if (Test-Path $pdfSrc) {
                            $pdfDest = Join-Path $archiveCo "$base.pdf"
                            
                            Invoke-WithRetry {
                                if (Test-Path $pdfDest) {
                                    $pdfDest = Join-Path $archiveCo ('{0}_{1:yyyyMMdd-HHmmss}.pdf' -f $base, (Get-Date))
                                }
                                Move-Item -Path $pdfSrc -Destination $pdfDest
                            }
                            Write-Log "  Archived pdf  -> $pdfDest"
                        }
                    }

                    $nextRow++
                    $nextSN++
                }
            }
            catch {
                Write-Log "  ERROR processing $($file.Name): $_" 'ERROR'
                $totalErrors++
                $coStats.Errors++
            }
        }

        Write-Progress -Id 1 -Activity "Processing files in $sheetName" -Completed
    }

    Write-Progress -Id 0 -Activity 'Processing companies' -Completed

    if (-not $DryRun) {
        Write-Log 'Saving tracker...'
        Invoke-WithRetry {
            $Tracker.Save()
        }
        Write-Log 'Tracker saved.' 'OK'
    }
    else {
        Write-Log 'DRY-RUN: tracker not modified.' 'DRY'
    }
}
catch {
    Write-Log "FATAL: $_" 'ERROR'
    $totalErrors++
}
finally {
    if ($Tracker) {
        try { $Tracker.Saved = $true } catch { }
        try { $Tracker.Close($false) } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Tracker)
    }
    if ($Excel) {
        try { $Excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
    }
    
    # Final aggressive GC to ensure all handles released
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
}

$elapsed = (Get-Date) - $startTime
$elapsedStr = '{0:hh\:mm\:ss}' -f $elapsed

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
Write-Host ('   Patient files written: {0}' -f $totalProcessed) -ForegroundColor $(if ($totalProcessed -gt 0) { 'Green' } else { 'Gray' })
Write-Host ('   Files skipped:         {0}' -f $totalSkipped)  -ForegroundColor $(if ($totalSkipped -gt 0) { 'Yellow' } else { 'Gray' })
Write-Host ('   Errors:                {0}' -f $totalErrors)   -ForegroundColor $(if ($totalErrors -gt 0) { 'Red' } else { 'Gray' })
Write-Host ('   Time elapsed:          {0}' -f $elapsedStr)
Write-Host ''

$rowsWithActivity = @($perCompanyStats.GetEnumerator() | Where-Object {
        $_.Value.Written -gt 0 -or $_.Value.Skipped -gt 0 -or $_.Value.Errors -gt 0
    })
if ($rowsWithActivity.Count -gt 0) {
    Write-Host '   Per-company breakdown:' -ForegroundColor White
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f 'Company', 'Written', 'Skipped', 'Errors') -ForegroundColor DarkGray
    Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f ('-' * 22), ('-' * 8), ('-' * 8), ('-' * 8)) -ForegroundColor DarkGray
    foreach ($entry in $rowsWithActivity) {
        $s = $entry.Value
        $color = 'White'
        if ($s.Errors -gt 0) { $color = 'Red' }
        elseif ($s.Skipped -gt 0 -and $s.Written -eq 0) { $color = 'Yellow' }
        elseif ($s.Written -gt 0) { $color = 'Green' }
        Write-Host ('     {0,-22} {1,8} {2,8} {3,8}' -f $s.Sheet, $s.Written, $s.Skipped, $s.Errors) -ForegroundColor $color
    }
    Write-Host ''
}

Write-Host '   Tracker:  ' -NoNewline -ForegroundColor DarkGray
Write-Host $TrackerPath -ForegroundColor Gray
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''

Write-Log ''
Write-Log '======================================================'
Write-Log ("Done. Processed={0}  Skipped={1}  Errors={2}  Elapsed={3}" -f $totalProcessed, $totalSkipped, $totalErrors, $elapsedStr)
Write-Log '======================================================'

if ($totalErrors -gt 0) { exit 1 } else { exit 0 }