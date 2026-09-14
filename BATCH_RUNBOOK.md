# AKS migration operator guide

This guide covers the pilot project and later legacy AKS workbooks. The tool
always creates a new output workbook. It never edits the legacy source or the
SAP template.

Most colleagues should use the application (`Start AKS Migration.bat`) described
in `README.md`. This runbook is for the controlled release process and for batch
migrations, which still run from the command line.

---

## Before a controlled migration release

Keep one working folder containing the approved versions of:

- `config\MigrationsMapping_00.xlsx`
- `templates\AKS_V2_06.xlsb`

Record which versions were used. The application shows both file names, sizes
and modification dates on screen, and writes them into every migration report.

---

## Migrating one workbook (the application)

1. Double-click `Start AKS Migration.bat`.
2. Check the **system check** in the sidebar: Microsoft Excel must be green.
3. **Step 1** — upload the legacy `.xlsb` file. Confirm the detected row count
   matches what you expect from the legacy `Übersicht` sheet.
4. **Step 2** — upload the approved migration mapping. Confirm the rule count
   matches the release you intend to use.
5. **Step 3** — upload the approved AKS V2 template. Confirm the SAP field count.
6. **Step 4 — Analyse Migration.** This is mandatory. Read the summary and work
   through the review table before writing anything.
7. **Step 5 — Run Migration.** Download all three result files, or open the run
   folder under `outputs\`.

Uploading the mapping and template every time is deliberate: it puts the
version used on screen, and the migration report records it.

Expected totals for the pilot project (`<project>_AKS_04.xlsb`):

| | |
|---|---|
| Legacy equipment rows | 208 |
| Pre-filled rows excluded | 767 |
| Active mapping rules | 50 |
| Mapped values examined | 10,634 |
| Values written | 9,237 |
| OK | 8,778 |
| Items requiring review | 489 |
| Warnings | 168 |
| Empty legacy cells skipped | 1,199 |

If a new legacy file produces very different numbers, stop and check that the
right source file and the right mapping version were used.

---

## Migrating one workbook (command line)

Always start with a dry run. `$py` is the interpreter the launcher built:

```powershell
$py = "$env:LOCALAPPDATA\AKS Migration App\env\Scripts\python.exe"

& $py aks_migrate.py `
  --source "LEGACY_AKS_FILE.xlsb" `
  --dry-run `
  --report "LEGACY_AKS_FILE_dry_run.csv"
```

Only after reviewing the dry run:

```powershell
& $py aks_migrate.py `
  --source "LEGACY_AKS_FILE.xlsb" `
  --output ".\outputs\LEGACY_AKS_FILE_migrated.xlsb" `
  --replace-template-data `
  --highlight
```

`--template` and `--mapping` default to the shipped files. Pass them explicitly
when validating a new mapping release.

Do not reuse an existing output name unless replacement is intentional. The
optional `--overwrite` switch permits replacement and should be used only after
the previous output has been archived. The application does not need this: it
writes every run into its own timestamped folder.

---

## Batch migration

After the pilot result is approved, place legacy workbooks in a folder named
`legacy_aks`, create an empty folder named `migrated_aks`, and run:

```powershell
& $py aks_migrate.py `
  --input-dir ".\legacy_aks" `
  --output-dir ".\migrated_aks" `
  --replace-template-data `
  --highlight
```

Run a dry run on each new source layout before including it in a batch. A header
mismatch or a missing SAP target key stops the run so the mapping can be reviewed
rather than guessed.

---

## Approval checklist

For each generated workbook:

1. Confirm the source-row count matches the legacy `Übersicht` sheet.
2. Compare project number and process position for every row.
3. Review every red cell and every REVIEW line in the migration audit.
4. Confirm all converted status and equipment codes with the project team.
5. Check formulas, dropdowns, hidden support sheets, external links, and VBA.
6. Confirm the workbook ends at the last migrated row and that the rows below it
   are empty.
7. Save the reviewed result under a controlled revision name.
8. Record who reviewed the exceptions and which mapping version was used.

Do not approve batch migration until the pilot specification questions in
`PILOT_FINDINGS.md` have documented answers and the generated workbook has passed
this checklist.

---

## Colour code in the generated workbook

| Colour | Meaning |
|---|---|
| Green | copied unchanged |
| Yellow | converted by a mapping rule |
| Red | written but outside the template's known SAP values — needs review |

A blank red or REVIEW target may be intentional: when two legacy columns feed one
SAP field with different values, the tool leaves the cell blank rather than
choosing. The engineering team decides.

---

## Common stops and what they mean

| Message | Meaning and action |
|---|---|
| *This file cannot be used as a legacy AKS list* | The workbook has no `Übersicht` sheet or no recognisable headings. Check that the right file was chosen. |
| *The migration mapping could not be applied to this legacy file* | A legacy column heading or SAP target key could not be resolved safely. The tool stops instead of guessing. Review the listed rules and add an audited layout exception before continuing. |
| *Template contains data through row N* | Confirm the rows are demonstration data, then leave the "Replace the demonstration rows" option ticked (command line: `--replace-template-data`). |
| *Output already exists* | Command line only. Choose a new name or archive the old output before using `--overwrite`. |
| *Microsoft Excel refused the request* / *did not open the template* | Close every Excel window, dismiss any Document Recovery message, end any leftover `EXCEL.EXE` in Task Manager, and run again. |
| *Microsoft Excel did not write the migrated workbook* | Excel declined the save. Check that the output folder is writable and that no copy of the file is open. |
| *Output check: the last populated row is N, but row M was expected* | The workbook was created but does not contain what was planned. Do not use it; send the log from `logs\` to the maintainer. |

The application never shows a Python traceback. The full technical detail,
including the traceback, is in the matching file under `logs\`.
