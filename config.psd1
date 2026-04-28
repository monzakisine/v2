@{
    # ============================================================
    # AMC AUTOMATION CONFIG
    # Edit this file when paths change. The PowerShell script
    # reads this on every run.
    # ============================================================

    # Root folder of the whole automation system on the Windows PC.
    # Change this when you deploy to the company computer.
    RootDir = 'C:\AMC-Automation'

    # Path of the master tracker file, RELATIVE to RootDir.
    TrackerRelPath = 'Tracker\Contractors_AMC_Tracker_2026.xlsm'

    # Subfolder names under RootDir.
    CompaniesDir = 'Companies'
    ArchiveDir   = 'Archive'
    LogsDir      = 'Logs'

    # Backup the tracker before each run? (Highly recommended.)
    BackupTrackerBeforeRun = $true

    # When the same Iqama already exists in the sheet:
    #   'warn'      -> log a warning, add the new row anyway (good for re-check-ups)
    #   'skip'      -> skip the patient, do not write
    #   'duplicate' -> add silently
    OnDuplicateIqama = 'warn'

    # ============================================================
    # COMPANY SHEET MAP
    # Maps the command-line key you type (e.g. 'scms')
    # to (a) the sheet name in the tracker
    #    (b) the folder name under Companies\
    # Add or remove entries as your company list changes.
    # Keys are case-insensitive when typed at the CLI.
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
    # AMC FORMULA -> TRACKER CELL MAPPING
    # Source rows refer to the "Field" sheet inside the patient's
    # filled AMCFORMULA file.
    # Target columns refer to the company sheet in the tracker.
    # If you change the formula or tracker structure, update here.
    # ============================================================
    SourceSheet = 'Field'

    # Header / patient identity cells in the formula
    PatientCells = @{
        Name        = 'C4'
        Company     = 'C5'
        Iqama       = 'C6'
        Age         = 'C7'
        DateAMC     = 'C8'
        DateReview  = 'C9'
        BloodPress  = 'C11'
        Height      = 'E11'
        Weight      = 'G11'
        Comment     = 'B48'
    }

    # Status (Fit / Unfit / Clinic Visit / Further Evaluation)
    # The script will check each candidate cell beside the label.
    # Whichever cell has a non-empty value (the checkmark) wins.
    StatusCandidates = @(
        @{ Label = 'FIT';                    CheckCells = @('H4','I4','F4') }
        @{ Label = 'UNFIT';                  CheckCells = @('H5','I5','F5') }
        @{ Label = 'CLINIC VISIT';           CheckCells = @('H6','I6','F6') }
        @{ Label = 'FOR FURTHER EVALUATION'; CheckCells = @('I7','H7','F7') }
    )

    # Test result cells in the formula (column where doctor adds
    # red fill on G to mark Abnormal). Default = NORMAL.
    # Each entry maps a tracker column letter to the formula row.
    # Tracker columns NOT listed here stay blank (e.g. Bilirubin,
    # HDL, Uric Acid - tests not present in the formula).
    TestRowMap = @(
        @{ TrackerCol = 'G';  FormulaRow = 15 }   # IM Consultation
        @{ TrackerCol = 'H';  FormulaRow = 16 }   # Ophtha Consultation
        @{ TrackerCol = 'L';  FormulaRow = 20 }   # CBC
        @{ TrackerCol = 'M';  FormulaRow = 21 }   # FBS
        @{ TrackerCol = 'O';  FormulaRow = 24 }   # SGPT/ALT
        @{ TrackerCol = 'P';  FormulaRow = 25 }   # SGOT/AST
        @{ TrackerCol = 'Q';  FormulaRow = 26 }   # GGT
        @{ TrackerCol = 'R';  FormulaRow = 27 }   # Alk. Phosphatase
        @{ TrackerCol = 'S';  FormulaRow = 28 }   # Cholesterol
        @{ TrackerCol = 'T';  FormulaRow = 29 }   # Triglycerides
        @{ TrackerCol = 'V';  FormulaRow = 30 }   # LDL
        @{ TrackerCol = 'W';  FormulaRow = 32 }   # Creatinine
        @{ TrackerCol = 'Y';  FormulaRow = 33 }   # TSH
        @{ TrackerCol = 'Z';  FormulaRow = 35 }   # Urinalysis (UA)
        @{ TrackerCol = 'AA'; FormulaRow = 34 }   # Stool
        @{ TrackerCol = 'AB'; FormulaRow = 40 }   # Chest X-ray
        @{ TrackerCol = 'AC'; FormulaRow = 37 }   # Audiometry
        @{ TrackerCol = 'AD'; FormulaRow = 38 }   # Spirometry
        @{ TrackerCol = 'AE'; FormulaRow = 43 }   # ECG
        @{ TrackerCol = 'AF'; FormulaRow = 45; MinAge = 40 }   # PSA (>=40 yr only)
        @{ TrackerCol = 'AG'; FormulaRow = 42 }   # Abdominal U/S
    )

    # Fixed-position columns in the tracker
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
}
