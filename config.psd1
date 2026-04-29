@{
    # ============================================================
    # AMC AUTOMATION CONFIG
    # Column letters verified directly from tracker headers.
    # ============================================================

    RootDir            = 'S:\AlBaithaMineClinic\AMC-automation'
    TrackerRelPath     = 'Tracker\Contractors_AMC_Tracker_2026.xlsm'
    CompaniesDir       = 'Companies'
    ArchiveDir         = 'Archive'
    LogsDir            = 'Logs'
    BackupTrackerBeforeRun = $true

    # 'warn'  -> add new row and log a warning
    # 'skip'  -> skip duplicate Iqama entirely
    OnDuplicateIqama = 'warn'

    # ============================================================
    # COMPANY SHEET MAP
    # ============================================================
    Companies = @{
        'altamimi'    = @{ Sheet = 'Al Tamimi';     Folder = 'Al Tamimi' }
        'mixed'       = @{ Sheet = 'Mixed';         Folder = 'Mixed' }
        'alsalem'     = @{ Sheet = 'AL SALEM';      Folder = 'AL SALEM' }
        'sar'         = @{ Sheet = 'SAR';           Folder = 'SAR' }
        'jal'         = @{ Sheet = 'JAL';           Folder = 'JAL' }
        'alamshawi'   = @{ Sheet = 'AL AMSHAWI';    Folder = 'AL AMSHAWI' }
        'catrion'     = @{ Sheet = 'CATRION';       Folder = 'CATRION' }
        'scms'        = @{ Sheet = 'SCMS';          Folder = 'SCMS' }
        'alsuwaidi'   = @{ Sheet = 'AL SUWAIDI';    Folder = 'AL SUWAIDI' }
        'aljari'      = @{ Sheet = 'AL JARI';       Folder = 'AL JARI' }
        'alseif'      = @{ Sheet = 'AL SEIF';       Folder = 'AL SEIF' }
        'malakalreem' = @{ Sheet = 'MALAK AL REEM'; Folder = 'MALAK AL REEM' }
        'almutairi'   = @{ Sheet = 'AL MUTAIRI';    Folder = 'AL MUTAIRI' }
        'nal'         = @{ Sheet = 'NAL';           Folder = 'NAL' }
        'reda'        = @{ Sheet = 'REDA';          Folder = 'REDA' }
    }

    # ============================================================
    # SOURCE SHEET NAME inside each patient AMCFORMULA file
    # ============================================================
    SourceSheet = 'Field'

    # ============================================================
    # PATIENT IDENTITY CELLS  (from AMCFORMULA "Field" sheet)
    # ============================================================
    PatientCells = @{
        Name       = 'C4'
        Company    = 'C5'
        Iqama      = 'C6'
        Age        = 'C7'
        DateAMC    = 'C8'
        DateReview = 'C9'
        BloodPress = 'C11'
        Height     = 'E11'
        Weight     = 'G11'
        Comment    = 'B48'
    }

    # ============================================================
    # STATUS DETECTION
    # VBA writes ChrW(10003) = checkmark into column I, rows 4-7
    # ============================================================
    StatusCandidates = @(
        @{ Label = 'FIT';                    CheckCell = 'I4' }
        @{ Label = 'UNFIT';                  CheckCell = 'I5' }
        @{ Label = 'CLINIC VISIT';           CheckCell = 'I6' }
        @{ Label = 'FOR FURTHER EVALUATION'; CheckCell = 'I7' }
    )

    # ============================================================
    # FIXED TRACKER COLUMNS  (verified from actual header row)
    #
    # A  = SN
    # B  = Date AMC
    # C  = Date reviewed
    # D  = Iqama
    # E  = Name
    # F  = Company
    # G  = IM CONSLTATION
    # H  = OPTHA. CONSULTATION
    # I  = HEIGHT
    # J  = WEIGHT
    # K  = UPDATED BMI
    # L  = CBC
    # M  = FBS
    # N  = SERUM BILIRUBIN       <- tracker-only, left blank
    # O  = SGPT/ALT
    # P  = SGOT/AST
    # Q  = GGT
    # R  = SERUM ALKALINE PHOSPHATASE
    # S  = SERUM CHOLE
    # T  = SERUM TRIGLYCERIDE
    # U  = HDL                   <- tracker-only, left blank
    # V  = LDL
    # W  = SERUM CREATININE
    # X  = URIC ACID             <- tracker-only, left blank
    # Y  = TSH
    # Z  = UA
    # AA = STOOL
    # AB = X-RAY
    # AC = AUDIO
    # AD = SPIRO
    # AE = ECG
    # AF = PSA Above 40
    # AG = Abdominal Ultrasound
    # AH = AGE
    # AI = BP
    # AJ = status
    # AK = Comments
    # AL = Nurse 1               <- filled manually by nurse
    # ============================================================
    FixedColumns = @{
        SerialNumber = 'A'
        DateAMC      = 'B'
        DateReview   = 'C'
        Iqama        = 'D'
        Name         = 'E'
        Company      = 'F'
        Height       = 'I'
        Weight       = 'J'
        BMIFormula   = 'K'
        Age          = 'AH'
        BloodPress   = 'AI'
        Status       = 'AJ'
        Comment      = 'AK'
    }

    # ============================================================
    # TEST ROW MAP
    # FormulaRow = row number in AMCFORMULA "Field" sheet
    # TrackerCol = exact column letter in the tracker
    # Columns N (SERUM BILIRUBIN), U (HDL), X (URIC ACID) are
    # NOT in the formula — left blank for manual entry.
    # ============================================================
    TestRowMap = @(
        @{ TrackerCol = 'G';  FormulaRow = 15 }   # IM CONSLTATION
        @{ TrackerCol = 'H';  FormulaRow = 16 }   # OPTHA. CONSULTATION
        @{ TrackerCol = 'L';  FormulaRow = 20 }   # CBC
        @{ TrackerCol = 'M';  FormulaRow = 21 }   # FBS
        @{ TrackerCol = 'O';  FormulaRow = 24 }   # SGPT/ALT
        @{ TrackerCol = 'P';  FormulaRow = 25 }   # SGOT/AST
        @{ TrackerCol = 'Q';  FormulaRow = 26 }   # GGT
        @{ TrackerCol = 'R';  FormulaRow = 27 }   # SERUM ALKALINE PHOSPHATASE
        @{ TrackerCol = 'S';  FormulaRow = 28 }   # SERUM CHOLE
        @{ TrackerCol = 'T';  FormulaRow = 29 }   # SERUM TRIGLYCERIDE
        @{ TrackerCol = 'V';  FormulaRow = 30 }   # LDL
        @{ TrackerCol = 'W';  FormulaRow = 32 }   # SERUM CREATININE
        @{ TrackerCol = 'Y';  FormulaRow = 33 }   # TSH
        @{ TrackerCol = 'Z';  FormulaRow = 35 }   # UA
        @{ TrackerCol = 'AA'; FormulaRow = 34 }   # STOOL
        @{ TrackerCol = 'AB'; FormulaRow = 40 }   # X-RAY
        @{ TrackerCol = 'AC'; FormulaRow = 37 }   # AUDIO
        @{ TrackerCol = 'AD'; FormulaRow = 38 }   # SPIRO
        @{ TrackerCol = 'AE'; FormulaRow = 43 }   # ECG
        @{ TrackerCol = 'AF'; FormulaRow = 45; MinAge = 40 }   # PSA Above 40
        @{ TrackerCol = 'AG'; FormulaRow = 42 }   # Abdominal Ultrasound
    )
}
