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
---------------------------



  ============================================================
            AMC TRACKER  -  COLUMN MAPPING DIAGNOSTIC
  ============================================================

  === Sheet: AL AMSHAWI ===
    A    SN
    B    Date AMC
    C    Date reviewed
    D    Iqama
    E    Name
    F    Company
    G    IM CONSLTATION
    H    OPTHA. CONSULTATION
    I    HEIGHT
    J    WEIGHT
    K    UPDATED BMI
    L    CBC
    M    FBS
    N    SERUM BILIRUBIN
    O    SGPT/ALT
    P    SGOT/AST
    Q    GGT
    R    SERUM ALKALINE PHOSPHATASE
    S    SERUM CHOLE
    T    SERUM TRIGLYCERIDE
    U    HDL
    V    LDL
    W    SERUM CREATININE
    X    URIC ACID
    Y    TSH
    Z    UA
    AA   STOOL
    AB   X-RAY
    AC   AUDIO
    AD   SPIRO
    AE   ECG
    AF   PSA Above 40
    AG   Abdominal Ultrasound
    AH   AGE
    AI   BP
    AJ   status
    AK   Comments
    AL   Nurse 1

  First data row (row 2):
    A    [SN                    ] = 1
    B    [Date AMC              ] = 46098
    C    [Date reviewed         ] = 46112
    D    [Iqama                 ] = 2568412825
    E    [Name                  ] = MUHAMMADMUUAMMAD KHAN
    F    [Company               ] = AL AMSHAWI
    G    [IM CONSLTATION        ] = ABNORMAL
    H    [OPTHA. CONSULTATION   ] = NORMAL
    I    [HEIGHT                ] = 190
    J    [WEIGHT                ] = 97
    K    [UPDATED BMI           ] = 26.9
    L    [CBC                   ] = NORMAL
    M    [FBS                   ] = NORMAL
    N    [SERUM BILIRUBIN       ] = NORMAL
    O    [SGPT/ALT              ] = NORMAL
    P    [SGOT/AST              ] = NORMAL
    Q    [GGT                   ] = NORMAL
    R    [SERUM ALKALINE PHOSPHATASE] = NORMAL
    S    [SERUM CHOLE           ] = NORMAL
    T    [SERUM TRIGLYCERIDE    ] = NORMAL
    U    [HDL                   ] = NORMAL
    V    [LDL                   ] = NORMAL
    W    [SERUM CREATININE      ] = NORMAL
    X    [URIC ACID             ] = NORMAL
    Y    [TSH                   ] = NORMAL
    Z    [UA                    ] = NORMAL
    AA   [STOOL                 ] = NORMAL
    AB   [X-RAY                 ] = NORMAL
    AC   [AUDIO                 ] = NORMAL
    AD   [SPIRO                 ] = NORMAL
    AE   [ECG                   ] = NORMAL
    AF   [PSA Above 40          ] = NORMAL
    AG   [Abdominal Ultrasound  ] = NORMAL
    AH   [AGE                   ] = 23
    AI   [BP                    ] = 120/80
    AJ   [status                ] = FIT
    AL   [Nurse 1               ] = Hakim

  === Sheet: AL JARI ===
    A    SN
    B    Date AMC
    C    Date reviewed
    D    Iqama
    E    Name
    F    Company
    G    IM CONSLTATION
    H    OPTHA. CONSULTATION
    I    HEIGHT
    J    WEIGHT
    K    UPDATED BMI
    L    CBC
    M    FBS
    N    SERUM BILIRUBIN
    O    SGPT/ALT
    P    SGOT/AST
    Q    GGT
    R    SERUM ALKALINE PHOSPHATASE
    S    SERUM CHOLE
    T    SERUM TRIGLYCERIDE
    U    HDL
    V    LDL
    W    SERUM CREATININE
    X    URIC ACID
    Y    TSH
    Z    UA
    AA   STOOL
    AB   X-RAY
    AC   AUDIO
    AD   SPIRO
    AE   ECG
    AF   PSA Above 40
    AG   Abdominal U/S
    AH   AGE
    AI   BP
    AJ   status
    AK   Comments
    AL   Nurse 1

  First data row (row 2):
    A    [SN                    ] = 1
    B    [Date AMC              ] = 46043
    C    [Date reviewed         ] = 46047
    D    [Iqama                 ] = 1112658420
    E    [Name                  ] = Abdul Malik Fraih Alharbi6
    F    [Company               ] = Al JARI
    G    [IM CONSLTATION        ] = NORMAL
    H    [OPTHA. CONSULTATION   ] = NORMAL
    I    [HEIGHT                ] = 174
    J    [WEIGHT                ] = 88
    K    [UPDATED BMI           ] = 29.1
    L    [CBC                   ] = NORMAL
    M    [FBS                   ] = NORMAL
    N    [SERUM BILIRUBIN       ] = NORMAL
    O    [SGPT/ALT              ] = NORMAL
    P    [SGOT/AST              ] = NORMAL
    Q    [GGT                   ] = NORMAL
    R    [SERUM ALKALINE PHOSPHATASE] = NORMAL
    S    [SERUM CHOLE           ] = NORMAL
    T    [SERUM TRIGLYCERIDE    ] = NORMAL
    U    [HDL                   ] = NORMAL
    V    [LDL                   ] = NORMAL
    W    [SERUM CREATININE      ] = NORMAL
    X    [URIC ACID             ] = NORMAL
    Y    [TSH                   ] = NORMAL
    Z    [UA                    ] = NORMAL
    AA   [STOOL                 ] = NORMAL
    AB   [X-RAY                 ] = NORMAL
    AC   [AUDIO                 ] = NORMAL
    AD   [SPIRO                 ] = NORMAL
    AE   [ECG                   ] = NORMAL
    AG   [Abdominal U/S         ] = NORMAL
    AH   [AGE                   ] = 25
    AI   [BP                    ] = 120/80
    AJ   [status                ] = FIT
    AL   [Nurse 1               ] = mohamedin

  === Full header map: AL MUTAIRI ===
    A    SN
    B    Date AMC
    C    Date reviewed
    D    Iqama
    E    Name
    F    Company
    G    IM CONSLTATION
    H    OPTHA. CONSULTATION
    I    HEIGHT
    J    WEIGHT
    K    UPDATED BMI
    L    CBC
    M    FBS
    N    SERUM BILIRUBIN
    O    SGPT/ALT
    P    SGOT/AST
    Q    GGT
    R    SERUM ALKALINE PHOSPHATASE
    S    SERUM CHOLE
    T    SERUM TRIGLYCERIDE
    U    HDL
    V    LDL
    W    SERUM CREATININE
    X    URIC ACID
    Y    TSH
    Z    UA
    AA   STOOL
    AB   X-RAY
    AC   AUDIO
    AD   SPIRO
    AE   ECG
    AF   PSA Above 40
    AG   Abdominal Ultrasound
    AH   AGE
    AI   BP
    AJ   status
    AK   Comments
    AL   Nurse 1

  Next empty row in AL MUTAIRI: 48
  Row 47 (last data row) col A-F:
    A    = 46
    B    =
    C    =
    D    =
    E    =
    F    =

  === Full header map: SCMS ===
    A    SN
    B    Date AMC
    C    Date reviewed
    D    Iqama
    E    Name
    F    Company
    G    IM CONSLTATION
    H    OPTHA. CONSULTATION
    I    HEIGHT
    J    WEIGHT
    K    UPDATED BMI
    L    CBC
    M    FBS
    N    SERUM BILIRUBIN
    O    SGPT/ALT
    P    SGOT/AST
    Q    GGT
    R    SERUM ALKALINE PHOSPHATASE
    S    SERUM CHOLE
    T    SERUM TRIGLYCERIDE
    U    HDL
    V    LDL
    W    SERUM CREATININE
    X    URIC ACID
    Y    TSH
    Z    UA
    AA   STOOL
    AB   X-RAY
    AC   AUDIO
    AD   SPIRO
    AE   ECG
    AF   PSA Above 40
    AG   Abdominal Ultrasound
    AH   AGE
    AI   BP
    AJ   status
    AK   Comments
    AL   Nurse 1

  Next empty row in SCMS: 40
  Row 39 (last data row) col A-F:
    A    = 38
    B    = 46070
    C    = 46140
    D    = 2551598382
    E    = KHALID AININA
    F    = SCMS

  Screenshot this entire window and send it.

  Press Enter to close:
