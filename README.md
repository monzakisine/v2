# AMC Automation 🩺

Automates the transfer of patient data from filled `AMCFORMULA.xlsm`
files into the master `Contractors_AMC_Tracker_2026.xlsm`.

## What it does

1. You drop reviewed patient files (`<iqama>.xlsm` plus optional
   `<iqama>.pdf`) into the matching company folder under `Companies\`.
2. The nurse double-clicks `amc.bat` on the desktop (or anywhere).
3. A friendly menu appears showing every company and **how many new
   patient files** are waiting in each folder. The nurse types the
   number of the company (or `0` for all) and presses Enter.
4. A PowerShell window shows live progress bars while the script
   reads each patient's identity, vitals, status, and Normal/Abnormal
   results, then appends one row per patient to the matching company
   sheet in the tracker.
5. A clean summary appears at the end (companies scanned, files
   written, files skipped, errors, time elapsed, per-company
   breakdown). The window stays open until the nurse presses Enter.
6. Processed files are moved to `Archive\<company>\` so they aren't
   imported twice.

## Folder layout

```
C:\AMC-Automation\
├── amc.bat                         <- THE ICON THE NURSE DOUBLE-CLICKS
├── Tracker\
│   └── Contractors_AMC_Tracker_2026.xlsm
├── Companies\
│   ├── SCMS\           <- drop reviewed files here
│   │   ├── 2627896331.xlsm
│   │   └── 2627896331.pdf
│   ├── Al Tamimi\
│   ├── CATRION\
│   └── ... (one per company)
├── Archive\            <- processed files end up here
├── Logs\               <- one log file per run + tracker backups
└── Scripts\
    ├── Run-Menu.ps1        <- friendly menu + final summary
    ├── Update-Tracker.ps1  <- the engine that does the work
    ├── Install.ps1
    └── config.psd1         <- edit this when paths change
```

## How the nurse uses it (daily)

1. Doctor drops reviewed files into the matching company folder.
2. Nurse double-clicks `amc.bat`.
3. A window like this appears:
   ```
   ============================================================
                A M C   A U T O M A T I O N
   ============================================================

     Choose what to process:

        0. ALL companies

        1. AL AMSHAWI
        2. AL JARI
        3. AL SALEM           (4 new patient files)
        4. AL SEIF
        5. AL SUWAIDI         (2 new patient files)
        ...
       12. SCMS               (8 new patient files)
       ...
       15. REDA

        P. Preview ALL (dry run, no changes written)
        Q. Quit

     Your choice:
   ```
4. Nurse types `0` (Enter) for all, or a single number for one company.
5. Progress bars run live while the engine works.
6. A summary screen shows totals + per-company breakdown.
7. Nurse presses Enter to close. Done.

## First-time setup on the Windows PC

1. Copy the entire `AMC-Automation` folder to `C:\AMC-Automation`.
   (If you put it elsewhere, edit `RootDir` in `Scripts\config.psd1`
   AND the `SCRIPTDIR` line in `amc.bat`.)
2. Place `Contractors_AMC_Tracker_2026.xlsm` inside `Tracker\`.
3. (Optional) Run the install script to create all subfolders and
   add a system PATH entry so `amc` works from a terminal too:
   ```powershell
   cd C:\AMC-Automation\Scripts
   powershell -ExecutionPolicy Bypass -File .\Install.ps1
   ```
4. (Recommended) Right-click `amc.bat` → Send to → Desktop (create
   shortcut). Now the nurse has a one-click icon on the desktop.
5. Double-click `amc.bat`, choose `P` (Preview ALL) to do a dry-run
   that doesn't change anything. If the summary looks right, you're
   ready for real runs.

## Power-user CLI mode (for Shams in a terminal)

If you call `amc` from PowerShell or cmd with arguments, it skips
the menu and runs the engine directly:

```
amc scms              <- process the SCMS folder
amc altamimi          <- process Al Tamimi folder
amc all               <- process every company in one go
amc scms -DryRun      <- preview only, no changes written
amc all -NoArchive    <- run normally but keep files in place
```

The valid company keys are listed in `Scripts\config.psd1` under the
`Companies` table. Keys are case-insensitive on the command line.

## How the script reads the formula file

| Field           | Source on `Field` sheet                                   |
| --------------- | --------------------------------------------------------- |
| Name            | C4                                                        |
| Company         | C5                                                        |
| Iqama           | C6                                                        |
| Age             | C7                                                        |
| Date AMC        | C8                                                        |
| Date Reviewed   | C9                                                        |
| Blood Pressure  | C11                                                       |
| Height / Weight | E11 / G11                                                 |
| BMI             | written as a live formula in the tracker                  |
| Status          | the cell next to FIT / UNFIT / etc. that has a checkmark  |
| Each test       | column G of the test row: red fill = ABNORMAL, else NORMAL|
| Comment         | B48 (merged area below "COMMENT")                         |

Tracker columns the formula doesn't have (Serum Bilirubin, HDL, Uric
Acid) are left blank for the nurse to fill manually if those tests
were ordered separately.

## Safety features

- **Backup before write** — a copy of the tracker is saved to `Logs\`
  before every non-dry-run.
- **Dry run** — `-DryRun` reads everything, prints what would happen,
  but doesn't touch the tracker or move any files.
- **Logging** — every run writes a timestamped log to `Logs\`.
- **Duplicate detection** — warns if an Iqama already exists in the
  sheet (configurable via `OnDuplicateIqama` in `config.psd1`).
- **Macros disabled on read** — patient files are opened with macros
  disabled, so the doctor's double-click-to-mark VBA never runs from
  the script.
- **Iqama as text** — written as text (`@` format) so leading zeros
  and 10-digit precision are preserved.

## When something changes

| What changed                          | Where to fix              |
| ------------------------------------- | ------------------------- |
| You moved the project to a new path   | `config.psd1` + `amc.bat` |
| New company added                     | `Companies` table         |
| Tracker got a new column              | `FixedColumns` or `TestRowMap` |
| Formula form got a new test row       | `TestRowMap`              |
| Different cell holds the comment      | `PatientCells.Comment`    |
| Status checkbox layout changed        | `StatusCandidates`        |

No code edits needed for any of those — just edit `config.psd1`.

## Troubleshooting

**"Cannot find Update-Tracker.ps1"** — the path in `amc.bat` doesn't
match the actual install location. Edit `amc.bat` and fix `PSPATH`.

**"amc is not recognized as a command"** — open a NEW terminal window
(PATH changes don't apply to terminals already open), or re-run
`Install.ps1`.

**"Sheet '<name>' not found in tracker"** — the `Sheet` name in
`config.psd1` doesn't exactly match the tab name in the tracker.
Fix the casing/spelling.

**"No status checkmark detected"** — the script looked at H4/I4/F4
etc. and found no value. Either the doctor didn't mark a status,
or the macro put the checkmark in a different cell. Open a filled
file, find which cell has the checkmark, and add it to
`StatusCandidates` in `config.psd1`.

**Excel process stays open after a crash** — open Task Manager and
end any orphan `EXCEL.EXE` processes. The script tries to clean up,
but a hard crash can leave one behind.
