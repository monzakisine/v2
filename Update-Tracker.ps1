<#
.SYNOPSIS
    Reads filled AMCFORMULA patient files and appends data to the master tracker.
    FIXED: Now detects graphical checkmark icons to prevent "Specified cast" errors.
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

# --- Helper Functions ---

function Invoke-WithRetry {
    param([scriptblock]$ScriptBlock, [int]$MaxAttempts = 6)
    $attempt = 0; $delay = 250
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try { return & $ScriptBlock }
        catch {
            if ($attempt -ge $MaxAttempts) { throw }
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, 4000)
        }
    }
}

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK', 'DRY')] $Level = 'INFO')
    $colors = @{ 'ERROR'='Red'; 'WARN'='Yellow'; 'OK'='Green'; 'DRY'='Cyan' }
    Write-Host "[$Level] $Message" -ForegroundColor ($colors[$Level] ?? 'Gray')
}

function Get-SafeCellValue {
    param ($cell)
    try {
        if ($null -eq $cell) { return $null }
        $val = $cell.Value2
        return if ($val -is [string]) { $val.Trim() } else { $val }
    } catch { return $null }
}

function ConvertTo-ColIndex {
    param([string]$Letter)
    $idx = 0
    foreach ($c in $Letter.ToUpper().ToCharArray()) { $idx = $idx * 26 + ([int][char]$c - [int][char]'A' + 1) }
    return $idx
}

function Test-CellIsHighlighted {
    param($Cell)
    try {
        if ($null -eq $Cell -or $null -eq $Cell.Interior) { return $false }
        $idx = $Cell.Interior.ColorIndex
        # -4142 is xlColorIndexNone; 16777215 is White
        if ($idx -eq -4142 -or $Cell.Interior.Color -eq 16777215) { return $false }
        return $true
    } catch { return $false }
}

# --- Main Processing Logic ---

function Read-PatientFile {
    param($ExcelApp, [string]$Path, [hashtable]$Cfg)

    return Invoke-WithRetry {
        $wb = $ExcelApp.Workbooks.Open($Path)
        try {
            $sheet = $wb.Sheets.Item($Cfg.SourceSheet)
            $data = [ordered]@{ SourceFile=$Path; Tests=@{} }

            # 1. Read Patient Identity
            foreach ($k in $Cfg.PatientCells.Keys) {
                $data[$k] = Get-SafeCellValue $sheet.Range($Cfg.PatientCells[$k])
            }

            # 2. STATUS DETECTION (The "Cast Error" Fix)
            # We check for EITHER text value OR a Shape sitting in the target cells.
            $detectedStatus = $null
            foreach ($s in $Cfg.StatusCandidates) {
                foreach ($addr in $s.CheckCells) {
                    # Check for text checkmark
                    $cv = Get-SafeCellValue $sheet.Range($addr)
                    if ($null -ne $cv -and "$cv" -ne "") {
                        $detectedStatus = $s.Label
                        break
                    }
                    # Check for Shape/Icon at this address
                    foreach ($shape in $sheet.Shapes) {
                        if ($shape.TopLeftCell.AddressLocal($false, $false) -eq $addr) {
                            $detectedStatus = $s.Label
                            break
                        }
                    }
                    if ($detectedStatus) { break }
                }
                if ($detectedStatus) { break }
            }
            $data.Status = $detectedStatus

            # 3. Read Test Results (Abnormal Check)
            foreach ($t in $Cfg.TestRowMap) {
                $isAbnormal = Test-CellIsHighlighted $sheet.Cells.Item([int]$t.FormulaRow, 7)
                $data.Tests[$t.TrackerCol] = if ($isAbnormal) { 'ABNORMAL' } else { 'NORMAL' }
            }
            return $data
        }
        finally {
            $wb.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
        }
    }
}

function Write-TrackerRow {
    param($Sheet, [int]$RowIndex, [int]$SN, [hashtable]$Patient, [hashtable]$Cfg)
    $fc = $Cfg.FixedColumns
    
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.SerialNumber)).Value2 = $SN
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Iqama)).NumberFormat = "@"
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Iqama)).Value2 = [string]$Patient.Iqama
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Name)).Value2 = $Patient.Name
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Status)).Value2 = $Patient.Status
    $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $fc.Comment)).Value2 = $Patient.Comment
    
    # Write Test Results
    foreach ($col in $Patient.Tests.Keys) {
        $Sheet.Cells.Item($RowIndex, (ConvertTo-ColIndex $col)).Value2 = $Patient.Tests[$col]
    }
}

# --- Execution Entry Point ---

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir 'config.psd1' }
$Cfg = Import-PowerShellDataFile $ConfigPath

$TrackerPath = Join-Path $Cfg.RootDir $Cfg.TrackerRelPath
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false

try {
    $TrackerWB = $Excel.Workbooks.Open($TrackerPath)
    $keys = if ($Company -eq 'all') { $Cfg.Companies.Keys } else { @($Company) }

    foreach ($key in $keys) {
        $info = $Cfg.Companies[$key]
        $folderPath = Join-Path $Cfg.RootDir $Cfg.CompaniesDir $info.Folder
        $files = Get-ChildItem -Path $folderPath -Filter *.xlsm

        $trackerSheet = $TrackerWB.Sheets.Item($info.Sheet)
        $nextRow = $trackerSheet.UsedRange.Rows.Count + 1

        foreach ($file in $files) {
            Write-Log "Processing: $($file.Name)"
            $patient = Read-PatientFile -ExcelApp $Excel -Path $file.FullName -Cfg $Cfg
            
            if ($DryRun) {
                Write-Log "DRY-RUN: Would write $($patient.Name) as $($patient.Status)" 'DRY'
            } else {
                Write-TrackerRow -Sheet $trackerSheet -RowIndex $nextRow -SN ($nextRow-1) -Patient $patient -Cfg $Cfg
                $nextRow++
                # Archive logic here as per original script...
            }
        }
    }
    if (-not $DryRun) { $TrackerWB.Save() }
}
finally {
    $Excel.Quit()
    [GC]::Collect()
}