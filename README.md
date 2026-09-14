# AKS Migration App

**Convert legacy AKS lists to the AKS V2 format.**

A Windows desktop application for migrating old AKS `.xlsb` equipment lists into
the SAP-era `AKS_V2_06.xlsb` template. It is a deterministic engineering tool:
every value is migrated according to `MigrationsMapping_00.xlsx`, and anything
the mapping does not cover is **flagged for human review, never guessed**.

---

## For colleagues using the tool

### Starting it

Double-click **`Start AKS Migration.bat`**.

That is the whole procedure. You do not need Python, a terminal, `pip`, or a
virtual environment.

- The **first** start takes a few minutes: the application installs everything it
  needs into your own Windows user profile, under
  `%LOCALAPPDATA%\AKS Migration App`. If this PC has no Python at all, a private
  copy of Python is downloaded there too — nothing is installed system-wide, no
  administrator rights are required, and your PC's own Python and PATH are not
  touched.
- Every start after that opens the application straight away in your browser.
- Leave the black window open while you work. Closing it stops the application.

You do need **Microsoft Excel for Windows** (the desktop version) on the PC. The
application starts its own hidden Excel in the background to write the `.xlsb`
file. The sidebar shows a system check that confirms this before you start.

### Using it

**Step 1 — Select the old AKS file.** Choose the legacy `.xlsb` list. The app
confirms the worksheet it found and how many equipment rows it detected. If the
file is not a legacy AKS list, it says so and stops.

**Step 2 — Migration configuration.** Nothing to do normally. The shipped
mapping and AKS V2 template are used and are named on screen. *Advanced
settings* lets an authorised user point at a different mapping or template.

**Step 3 — Analyse Migration.** A dry run. Nothing is written. You get:

| Status | Meaning |
|---|---|
| **OK** | migrated automatically |
| **REVIEW** | no migration mapping defined — the engineering team decides |
| **WARNING** | migrated, but worth checking (for example an approximate column match) |
| **ERROR** | the migration cannot run |

Review items are grouped by parameter so you can see the themes, and listed
individually with the old value, the source cell, and the reason.

**Step 4 — Run Migration.** Creates the migrated workbook. Your original file and
the master template are never modified.

**Results.** Three downloads, also saved under `outputs\`:

| File | Contents |
|---|---|
| `<project>_AKS_V2_Migrated.xlsb` | the migrated workbook |
| `<project>_Migration_Audit.xlsx` | every value, before and after, with status and reason |
| `<project>_Migration_Report.xlsx` | timestamp, versions, totals, unresolved mappings |

Every run writes to its own timestamped folder, so an earlier result is never
overwritten. A `.csv` copy of the audit is written alongside the `.xlsx`.

### If something goes wrong

The application shows a plain-language message and writes the technical detail to
`logs\`. Close every Excel window and try again; if it keeps happening, send the
log file from `logs\` to the tool maintainer.

---

## What the migration does

- Reads the legacy `Übersicht` worksheet with `pyxlsb` — Excel is not involved in
  the analysis, so Step 3 works on any Windows PC.
- Reads the active rules from `MigrationsMapping_00.xlsx` at runtime, so mapping
  changes do not require a code change.
- Matches each legacy column by its heading. If a heading cannot be matched with
  confidence, **the run stops** rather than reading the wrong column. Seven
  hand-audited exceptions cover columns that were renamed in the legacy layout.
- Converts values through the mapping's replacement tables, derives the numeric
  status code, and normalises semicolons and commas to periods where the rule
  says so.
- Flags for review: values absent from the template's `MIG-Werte Merkmale`
  allowed-value list, unresolvable status codes, and cases where two legacy
  columns feed one SAP field with conflicting values (the target is left blank).
- Writes with Microsoft Excel via COM so formulas, dropdowns, hidden support
  sheets and the VBA project survive. Calculation is set to manual during the
  write and recalculated once at the end; the operator's Excel setting is
  restored afterwards.
- Colours the migrated cells: green copied, yellow converted, red needs review.
- Adds a `Migration_Report` worksheet to the generated workbook.
- Reopens the finished workbook and verifies the expected rows are present.

### Deliberate behaviours worth knowing

- **The master template is opened read-only.** The result is written with
  `SaveAs`, so the template cannot be modified even in principle.
- **Excel writes to a local folder first.** Inside a OneDrive-synced folder Excel
  silently redirects the path to SharePoint, opens the workbook read-only and
  discards the save without reporting an error. The finished file is moved into
  `outputs\` afterwards.
- **The seed row supplies formulas, not formatting.** The supplied template's
  seed row has columns formatted as Text and lacks the numeric validation that
  the rows below it carry. Copying it wholesale would store migrated numbers as
  text and drop the template's own dropdowns, so rows the template has already
  prepared receive formulas only.
- **Rows below the migrated data are cleared**, so the workbook ends at the last
  migrated row instead of trailing `#N/A` formulas.
- **"Read-only recommended" is cleared on the output.** The template carries that
  flag; the migrated workbook has to be editable for the review step.

---

## For maintainers

### Layout

```
AKS Migration App/
├── Start AKS Migration.bat     double-click launcher
├── install_and_run.ps1         bootstrap: environment, Python, start
├── app.py                      Streamlit user interface (presentation only)
├── migration/                  the engine - no user-interface code
│   ├── api.py                  analyse_migration() / run_migration()
│   ├── engine.py               build_plan: legacy cell -> SAP cell
│   ├── mapping.py              rules, column resolution, value conversion
│   ├── model.py                constants, dataclasses, value helpers
│   ├── excel_reader.py         read .xlsb via pyxlsb (no Excel needed)
│   ├── excel_writer.py         write .xlsb via Excel COM
│   ├── validation.py           environment, input and output checks
│   ├── reporting.py            statuses, audit and report files
│   ├── runlog.py               one log file per run
│   └── paths.py                where files live
├── aks_migrate.py              command-line interface (unchanged V1 commands)
├── config/MigrationsMapping_00.xlsx
├── templates/AKS_V2_06.xlsb
├── pilot/                      pilot project source and reference output
├── outputs/                    one folder per run
└── logs/                       one log per run, plus launcher logs
```

### Files that are not in this repository

Three things are deliberately absent, because they are STADLER engineering
property rather than code:

| Missing | Put it here before first use |
|---|---|
| `MigrationsMapping_00.xlsx` | `config\` |
| `AKS_V2_06.xlsb` | `templates\` |
| Legacy pilot workbooks | `pilot\` (only needed to re-run the regression check) |

The application will not start a migration without the first two. It reports
exactly which one is missing in Step 2.

### The engine is independent of the user interface

`app.py` never reads a workbook, never talks to Excel, and never decides what a
value becomes. It calls two functions:

```python
from migration import analyse_migration, run_migration

analysis = analyse_migration("<project>_AKS_04.xlsb")     # dry run, writes nothing
if analysis.ok:
    print(analysis.rows_processed, analysis.review_count)
    for row in analysis.review_rows(10):
        print(row["Parameter"], row["Old Value"], row["Reason"])
    result = run_migration(analysis)                    # writes the workbook
    print(result.workbook_path, result.audit_path, result.report_path)
```

Neither function raises for an expected migration problem: inspect `.ok`,
`.errors` and `.warnings`. Supporting calls: `inspect_source_file`,
`check_environment`, `describe_configuration`, `stage_source`.

### Command line

The V1 pilot commands still work, and `--template` / `--mapping` now default to
the shipped files:

The interpreter the launcher built lives in
`%LOCALAPPDATA%\AKS Migration App\env\Scripts\python.exe`:

```powershell
$py = "$env:LOCALAPPDATA\AKS Migration App\env\Scripts\python.exe"

& $py aks_migrate.py --source ".\pilot\<project>_AKS_04.xlsb" `
  --dry-run --report "dry_run.csv"

& $py aks_migrate.py --source ".\pilot\<project>_AKS_04.xlsb" `
  --output ".\outputs\manual.xlsb" --replace-template-data --highlight
```

### Regression baseline

The refactor is verified against the V1 pilot:

- `aks_migrate.py --dry-run` produces a CSV **byte-identical** to
  `pilot/<project>_AKS_04_dry_run.csv` (SHA-256 `4C78CBF6...`);
- 208 rows, 50 rules, 10,634 audit entries, 9,237 applied writes;
- the generated workbook matches `pilot/<project>_AKS_04_migrated_verified.xlsb`
  on **10,619 of 10,619 compared cells**, with one intentional difference:
  cell `A46` is left blank and flagged for review instead of holding `#N/A`.

Re-run these checks after any change to `migration/`.

### Launcher options

```powershell
.\install_and_run.ps1 -Reinstall     # rebuild the environment
.\install_and_run.ps1 -Port 8600     # use another port
.\install_and_run.ps1 -NoBrowser     # do not open the browser
.\install_and_run.ps1 -DebugErrors   # show technical errors in the app
```

The environment is considered ready when a stamp file next to the interpreter
matches the SHA-256 of `requirements.txt`, so editing `requirements.txt`
triggers a reinstall on the next start and nothing else does. The app is bound
to `localhost` and is not reachable from the network.

The environment is built in `%LOCALAPPDATA%\AKS Migration App` rather than in the
application folder, which is normally on OneDrive: a ~400 MB Python environment
there would be uploaded to SharePoint and is prone to sync locking. Copying the
application folder to another PC is therefore safe — the next start rebuilds the
environment locally. An environment created by an earlier version in a `.venv`
subfolder is still recognised and reused.
