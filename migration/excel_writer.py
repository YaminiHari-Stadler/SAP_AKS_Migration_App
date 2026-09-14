"""Write the migrated workbook with Microsoft Excel via COM automation.

This is the only module that needs Excel. It always works on a copy of the V2
template; the legacy source workbook and the master template are opened
read-only or not at all.

Carried over unchanged from the V1 pilot script, including the manual
calculation strategy: recalculating the template's large lookup formulas after
every single written cell is far too slow, so calculation is switched to manual,
all values are written, one full calculation runs, and the operator's original
Excel calculation setting is restored.
"""

from __future__ import annotations

import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, TypeVar

from .paths import local_staging_dir
from .model import (
    REPORT_COLUMNS,
    TARGET_DATA_START_ROW,
    TARGET_KEY_ROW,
    TARGET_SHEET,
    MigrationEntry,
    MigrationError,
    display,
)


T = TypeVar("T")

# Excel constants
XL_PASTE_ALL = -4104
XL_PASTE_FORMULAS = -4123
XL_CALCULATION_MANUAL = -4135
XL_CELL_TYPE_CONSTANTS = 2
XL_TO_LEFT = -4159
XL_UP = -4162

# Excel is single-threaded. While it is busy — saving a 4 MB .xlsb takes a
# while — it rejects incoming automation calls with one of these HRESULTs
# instead of queueing them. The correct response is to wait and ask again.
COM_BUSY_HRESULTS = frozenset(
    {
        -2147418111,  # RPC_E_CALL_REJECTED   "Call was rejected by callee."
        -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
        -2147417851,  # RPC_E_SERVERFAULT / server busy
    }
)


def com_retry(
    action: Callable[[], T],
    attempts: int = 90,
    delay: float = 1.0,
    retry_attribute_errors: bool = False,
) -> T:
    """Call Excel, waiting politely when it is busy rather than failing the run.

    `retry_attribute_errors` is for the first property read after opening a
    workbook. pywin32 resolves names late: while Excel is still loading a large
    file the name lookup fails and pywin32 reports a plain `AttributeError`
    rather than a COM "server busy" error, which would otherwise look like a
    missing property.
    """
    last_error: BaseException | None = None
    for attempt in range(attempts):
        try:
            return action()
        except AttributeError as exc:
            if not retry_attribute_errors:
                raise
            last_error = exc
            time.sleep(delay if attempt < 30 else delay * 2)
        except Exception as exc:  # pywintypes.com_error and friends
            hresult = None
            args = getattr(exc, "args", ())
            if args and isinstance(args[0], int):
                hresult = args[0]
            if hresult not in COM_BUSY_HRESULTS:
                raise
            last_error = exc
            time.sleep(delay if attempt < 30 else delay * 2)
    raise MigrationError(
        "Microsoft Excel stayed busy and did not accept the request. "
        "Close all Excel windows and run the migration again."
    ) from last_error


def rgb(red: int, green: int, blue: int) -> int:
    return red + green * 256 + blue * 65536


FILL_COLORS = {
    "COPIED": rgb(198, 239, 206),
    "CONVERTED": rgb(255, 235, 156),
    "MANUAL_REVIEW": rgb(255, 199, 206),
}


def populate_report_sheet(
    workbook: Any, entries: list[MigrationEntry], summary: dict[str, Any]
) -> None:
    excel = workbook.Application
    for sheet in workbook.Worksheets:
        if sheet.Name == "Migration_Report":
            sheet.Delete()
            break
    report = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    report.Name = "Migration_Report"
    report.Cells(1, 1).Value = "AKS migration report"
    report.Cells(2, 1).Value = "Source"
    report.Cells(2, 2).Value = summary["source"]
    report.Cells(3, 1).Value = "Created"
    report.Cells(3, 2).Value = datetime.now().isoformat(timespec="seconds")
    report.Cells(4, 1).Value = "Source rows"
    report.Cells(4, 2).Value = summary["source_rows"]

    header_row = 6
    for column, name in enumerate(REPORT_COLUMNS, start=1):
        report.Cells(header_row, column).Value = name.replace("_", " ").title()
    data = [tuple(display(getattr(entry, column)) for column in REPORT_COLUMNS) for entry in entries]
    if data:
        report.Range(
            report.Cells(header_row + 1, 1),
            report.Cells(header_row + len(data), len(REPORT_COLUMNS)),
        ).Value = tuple(data)
    header = report.Range(report.Cells(header_row, 1), report.Cells(header_row, len(REPORT_COLUMNS)))
    header.Font.Bold = True
    header.Font.Color = rgb(255, 255, 255)
    header.Interior.Color = rgb(31, 78, 121)
    header.AutoFilter()
    report.Columns("A:M").EntireColumn.AutoFit()
    report.Columns("C:C").ColumnWidth = 34
    report.Columns("E:E").ColumnWidth = 24
    report.Columns("I:I").ColumnWidth = 24
    report.Columns("K:L").ColumnWidth = 48
    report.Columns("C:M").WrapText = True
    # Freezing the header is cosmetic. On a hidden Excel instance ActiveWindow
    # can be Nothing, and this must not throw away a migration whose values are
    # already written but not yet saved.
    try:
        report.Activate()
        excel.ActiveWindow.SplitRow = header_row
        excel.ActiveWindow.FreezePanes = True
    except Exception:
        pass


def migrate_with_excel(
    template_path: Path,
    output_path: Path,
    entries: list[MigrationEntry],
    summary: dict[str, Any],
    replace_template_data: bool,
    highlight: bool,
    overwrite: bool,
    progress: Callable[[str, float], None] | None = None,
) -> None:
    # Excel resolves relative paths against its own default folder, which can
    # differ from the terminal's working directory. Always hand COM absolute
    # paths so the same command works from PowerShell, Task Scheduler, or an IDE.
    template_path = Path(template_path).resolve()
    output_path = Path(output_path).resolve()

    def report(message: str, fraction: float) -> None:
        if progress is not None:
            progress(message, fraction)

    if output_path.exists() and not overwrite:
        raise MigrationError(f"Output already exists: {output_path}. Use --overwrite to replace it.")
    if output_path.resolve() in {template_path.resolve(), Path(summary["source"]).resolve()}:
        raise MigrationError("Output must be different from both source and template")
    if summary["template_last_data_row"] >= TARGET_DATA_START_ROW and not replace_template_data:
        raise MigrationError(
            f"Template contains data through row {summary['template_last_data_row']}. "
            "Re-run with --replace-template-data after verifying it is demonstration data."
        )
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise MigrationError("pywin32 is required for .xlsb writing: pip install pywin32") from exc

    # Excel is given a plain local folder to write into. In a OneDrive-synced
    # folder Excel silently rewrites the path to the SharePoint URL behind it,
    # opens the workbook read-only, and then discards the save without raising.
    staging_dir = local_staging_dir(output_path.stem)
    staging_path = staging_dir / output_path.name
    if staging_path.exists():
        staging_path.unlink()

    report("Opening the AKS V2 template", 0.05)
    previous_calculation = None
    previous_calculate_before_save = None
    control_workbook = None
    excel = None
    workbook = None
    saved = False
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        # Policy-locked or add-in-heavy Excel installations reject some of these.
        # None of them is essential to the migration, so a refusal must not leave
        # an orphaned EXCEL.EXE behind - which is why the Dispatch and this setup
        # now sit inside the try whose finally quits Excel.
        for attribute, value in (
            ("Visible", False),
            ("ScreenUpdating", False),
            ("DisplayAlerts", False),
            ("EnableEvents", False),
            ("AskToUpdateLinks", False),
            ("AutomationSecurity", 3),
        ):
            try:
                setattr(excel, attribute, value)
            except Exception:
                pass
        try:
            previous_calculation = excel.Calculation
            previous_calculate_before_save = excel.CalculateBeforeSave
            excel.Calculation = XL_CALCULATION_MANUAL
            excel.CalculateBeforeSave = False
        except Exception:
            # Excel may reject Calculation changes when no workbook exists. A
            # temporary blank workbook establishes the application calculation
            # context before the formula-heavy AKS file is opened.
            control_workbook = excel.Workbooks.Add()
            previous_calculation = excel.Calculation
            previous_calculate_before_save = excel.CalculateBeforeSave
            excel.Calculation = XL_CALCULATION_MANUAL
            excel.CalculateBeforeSave = False

        report("Reading the AKS V2 template", 0.10)
        # The template is opened read-only on purpose: Excel then cannot write
        # to it under any circumstance. All edits happen in memory and are
        # written out with SaveAs below, to a new file.
        # Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
        #      IgnoreReadOnlyRecommended, ...) — positional, because pywin32's
        # late binding does not reliably forward named arguments to Excel.
        workbook = com_retry(
            lambda: excel.Workbooks.Open(
                str(template_path), 0, True, None, None, None, True,
            )
        )
        # The empty scratch workbook stays open until the very end. Closing it
        # here, while Excel is still loading the large template, makes Excel
        # stop responding to automation.

        # Excel returns Nothing rather than raising when it declines to open a
        # file, for example while a Document Recovery pane is pending.
        if workbook is None:
            raise MigrationError(
                "Microsoft Excel did not open the AKS V2 template. Open Excel once by hand, "
                "dismiss any 'Document Recovery' or update message, close Excel, and run the "
                "migration again."
            )

        # Wait for Excel to finish loading the workbook before touching it.
        # An AttributeError here means "still busy"; a COM error from the call
        # itself means the worksheet really is not in this template.
        try:
            worksheets = com_retry(
                lambda: workbook.Worksheets, attempts=120, retry_attribute_errors=True
            )
        except MigrationError:
            raise MigrationError(
                "Microsoft Excel opened the AKS V2 template but stopped responding. Close every "
                "Excel window, dismiss any 'Document Recovery' message, and run the migration again."
            ) from None
        try:
            sheet = worksheets(TARGET_SHEET)
        except Exception as exc:
            raise MigrationError(
                f"The AKS V2 template has no '{TARGET_SHEET}' worksheet, so it cannot be used "
                "as a migration target."
            ) from exc
        last_key_column = sheet.Cells(TARGET_KEY_ROW, sheet.Columns.Count).End(-4159).Column  # xlToLeft
        existing_last_row = max(
            sheet.Cells(sheet.Rows.Count, 3).End(-4162).Row,  # xlUp
            sheet.Cells(sheet.Rows.Count, 4).End(-4162).Row,
        )
        target_last_row = summary["last_target_row"]
        if existing_last_row >= TARGET_DATA_START_ROW and not replace_template_data:
            raise MigrationError(
                f"Template contains data through row {existing_last_row}. "
                "Re-run with --replace-template-data after verifying it is demonstration data."
            )

        report("Preparing the target rows", 0.15)
        clear_last_row = max(existing_last_row, target_last_row)
        if clear_last_row > TARGET_DATA_START_ROW:
            seed = sheet.Range(
                sheet.Cells(TARGET_DATA_START_ROW, 1),
                sheet.Cells(TARGET_DATA_START_ROW, last_key_column),
            )
            # The seed row supplies the per-row formulas. Its *formatting* must
            # not be spread across the sheet: in the supplied template the seed
            # row has columns formatted as Text and without the numeric data
            # validation that the rows below it carry, so copying everything
            # would store migrated numbers as text and drop the template's own
            # dropdowns. Rows the template has already prepared therefore
            # receive formulas only and keep their own formats and validation.
            # Which rows has the template already prepared with number formats
            # and data validation? `template_last_data_row` comes from the
            # pyxlsb read, which drops a formula cell whose cached value is ""
            # - Excel's own End(xlUp) counts that cell as occupied. Taking the
            # larger of the two keeps such rows in the formulas-only branch
            # below, instead of letting the seed row's Text format reach them.
            prepared_last_row = max(
                existing_last_row,
                int(summary.get("template_last_data_row", 0)),
                TARGET_DATA_START_ROW,
            )
            inner_last = min(clear_last_row, prepared_last_row)
            if inner_last > TARGET_DATA_START_ROW:
                inner = sheet.Range(
                    sheet.Cells(TARGET_DATA_START_ROW + 1, 1),
                    sheet.Cells(inner_last, last_key_column),
                )
                com_retry(lambda: seed.Copy())
                com_retry(lambda: inner.PasteSpecial(Paste=XL_PASTE_FORMULAS))
            if clear_last_row > inner_last:
                # Rows beyond anything the template prepared have no formatting
                # of their own, so these do get the full seed row.
                outer = sheet.Range(
                    sheet.Cells(max(inner_last + 1, TARGET_DATA_START_ROW + 1), 1),
                    sheet.Cells(clear_last_row, last_key_column),
                )
                com_retry(lambda: seed.Copy())
                com_retry(lambda: outer.PasteSpecial(Paste=XL_PASTE_ALL))
            try:
                excel.CutCopyMode = False
            except Exception:
                pass

        data_range = sheet.Range(
            sheet.Cells(TARGET_DATA_START_ROW, 1),
            sheet.Cells(clear_last_row, last_key_column),
        )
        try:
            data_range.SpecialCells(2).ClearContents()  # xlCellTypeConstants
        except Exception:
            pass

        writable = [
            entry for entry in entries
            if entry.apply and entry.status != "SKIPPED" and entry.target_cell
        ]

        total = len(writable) or 1
        for index, entry in enumerate(writable):
            cell = sheet.Range(entry.target_cell)
            cell.Value = entry.new_value
            if highlight and entry.status in FILL_COLORS:
                cell.Interior.Color = FILL_COLORS[entry.status]
            if index % 250 == 0:
                report(f"Writing migrated values ({index:,} of {len(writable):,})",
                       0.15 + 0.60 * index / total)

        # Rows below the migrated data still carry the formulas that were
        # replicated from the seed row, which display as #N/A once their inputs
        # are gone. Clear those rows completely - values and formulas - so the
        # migrated workbook ends at the last migrated row. Formatting, data
        # validation and conditional formatting are left in place.
        if clear_last_row > target_last_row:
            report("Removing the unused template rows", 0.78)
            com_retry(
                lambda: sheet.Range(
                    sheet.Cells(target_last_row + 1, 1),
                    sheet.Cells(clear_last_row, last_key_column),
                ).ClearContents()
            )

        report("Adding the Migration_Report sheet", 0.80)
        populate_report_sheet(workbook, entries, summary)
        # Recalculate once after all writes. This avoids recalculating the
        # template's large lookup formulas after every migrated cell.
        report("Recalculating the workbook", 0.88)
        com_retry(lambda: excel.CalculateFull())
        # Persist the user's original calculation behavior in the output. The
        # full calculation above has already refreshed all workbook results.
        if previous_calculation is not None:
            excel.Calculation = previous_calculation
            excel.CalculateBeforeSave = previous_calculate_before_save
            previous_calculation = None
            previous_calculate_before_save = None
        report("Saving the migrated workbook", 0.94)
        # SaveAs writes the edited copy to a new file and leaves the template
        # on disk exactly as it was. Saving a formula-heavy .xlsb keeps Excel
        # busy for a while; it rejects further calls until it has finished.
        # SaveAs(Filename, FileFormat, Password, WriteResPassword,
        #        ReadOnlyRecommended, CreateBackup, AccessMode,
        #        ConflictResolution, AddToMru, ...)
        # FileFormat 50 = xlExcel12 (.xlsb, keeps macros and structure).
        # ReadOnlyRecommended False: the template carries that flag, but the
        # migrated workbook has to be editable for the review step.
        com_retry(
            lambda: workbook.SaveAs(
                str(staging_path), 50, None, None, False, False, 1, 2, False, None, None, True,
            )
        )
        saved = True
    except Exception:
        if workbook is not None:
            try:
                com_retry(lambda: workbook.Close(SaveChanges=False), attempts=20)
            except Exception:
                pass
            workbook = None
        raise
    finally:
        if workbook is not None:
            # SaveAs has already written everything to disk, so never ask Excel
            # to save again while closing. A failure to close is not a failure
            # to migrate: the file is on disk and is verified afterwards.
            try:
                com_retry(lambda: workbook.Close(SaveChanges=False))
            except Exception:
                pass
        if control_workbook is not None:
            try:
                com_retry(lambda: control_workbook.Close(SaveChanges=False), attempts=20)
            except Exception:
                pass
        if excel is not None and previous_calculation is not None:
            try:
                excel.Calculation = previous_calculation
                excel.CalculateBeforeSave = previous_calculate_before_save
            except Exception:
                pass
        if excel is not None:
            try:
                com_retry(lambda: excel.Quit(), attempts=20)
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not saved or not staging_path.exists() or staging_path.stat().st_size == 0:
        raise MigrationError(
            "Microsoft Excel did not write the migrated workbook. Close every Excel window "
            "and run the migration again."
        )

    report("Moving the migrated workbook to the results folder", 0.98)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()
    shutil.move(str(staging_path), str(output_path))
    try:
        staging_dir.rmdir()
    except OSError:
        pass

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise MigrationError("Excel did not create a valid output file")
