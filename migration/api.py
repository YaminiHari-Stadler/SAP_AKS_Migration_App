"""The only entry points a user interface needs.

    inspect_source_file(path)      -> is this a legacy AKS workbook?
    check_environment()            -> can this machine run a migration?
    analyse_migration(source)      -> dry run: what would happen?
    run_migration(analysis)        -> write the migrated workbook and artefacts

The migration engine is deliberately independent of Streamlit: these functions
take and return plain paths, dictionaries, and dataclasses, so the same calls
work from the command line, a test, or any future front end.
"""

from __future__ import annotations

import shutil
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Sequence

from . import paths as app_paths
from . import reporting
from .engine import MappingResolutionError, build_plan
from .excel_writer import migrate_with_excel
from .model import MigrationEntry, MigrationError
from .runlog import RunLog, friendly_error, new_run_id
from .validation import (
    EnvironmentReport,
    OutputVerification,
    SourceInspection,
    check_environment,
    inspect_source,
    validate_mapping,
    validate_template,
)

ProgressCallback = Callable[[str, float], None]

MIGRATED_SUFFIX = "_AKS_V2_Migrated"


# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------


@dataclass
class AnalysisResult:
    """What a dry run found. Pass this straight into `run_migration`."""

    ok: bool
    run_id: str
    source_path: Path
    source_name: str
    mapping_path: Path
    template_path: Path
    log_path: Path
    duration_seconds: float = 0.0
    summary: dict[str, Any] = field(default_factory=dict)
    counts: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    unresolved_mappings: list[dict[str, Any]] = field(default_factory=list)
    entries: list[MigrationEntry] = field(default_factory=list)
    log_text: str = ""
    _identities: dict[int, dict[str, str]] = field(default_factory=dict, repr=False)
    _log: RunLog | None = field(default=None, repr=False)

    # -- convenience views used by the interface ---------------------------

    @property
    def rows_processed(self) -> int:
        return int(self.summary.get("source_rows", 0))

    @property
    def values_to_migrate(self) -> int:
        return int(self.summary.get("applied_writes", 0))

    @property
    def review_count(self) -> int:
        return self.counts.get(reporting.STATUS_REVIEW, 0)

    @property
    def warning_count(self) -> int:
        return self.counts.get(reporting.STATUS_WARNING, 0)

    @property
    def error_count(self) -> int:
        return self.counts.get(reporting.STATUS_ERROR, 0) + len(self.errors)

    def table_rows(
        self, statuses: Sequence[str] | None = None, limit: int | None = None
    ) -> list[dict[str, Any]]:
        rows = reporting.to_table_rows(self.entries, self._identities, statuses)
        return rows[:limit] if limit else rows

    def review_rows(self, limit: int | None = None) -> list[dict[str, Any]]:
        return self.table_rows(reporting.REVIEW_STATUSES, limit)

    def review_breakdown(self) -> list[dict[str, Any]]:
        return reporting.review_breakdown(self.entries)


@dataclass
class MigrationResult:
    """What a real migration produced."""

    ok: bool
    run_id: str
    source_name: str
    output_dir: Path | None = None
    workbook_path: Path | None = None
    audit_path: Path | None = None
    audit_csv_path: Path | None = None
    report_path: Path | None = None
    log_path: Path | None = None
    duration_seconds: float = 0.0
    summary: dict[str, Any] = field(default_factory=dict)
    counts: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    verification: OutputVerification | None = None
    log_text: str = ""

    @property
    def rows_processed(self) -> int:
        return int(self.summary.get("source_rows", 0))

    @property
    def values_migrated(self) -> int:
        return int(self.summary.get("applied_writes", 0))

    @property
    def review_count(self) -> int:
        return self.counts.get(reporting.STATUS_REVIEW, 0)

    @property
    def warning_count(self) -> int:
        return self.counts.get(reporting.STATUS_WARNING, 0)

    @property
    def error_count(self) -> int:
        return self.counts.get(reporting.STATUS_ERROR, 0) + len(self.errors)


# ---------------------------------------------------------------------------
# Step 1 — is the machine and the file usable?
# ---------------------------------------------------------------------------


def inspect_source_file(path: Path | str) -> SourceInspection:
    """Check that a file really is a legacy AKS workbook and count its rows."""
    return inspect_source(Path(path))


CONFIG_STAGE_ID = "_config"


def stage_source(data: bytes, filename: str, run_id: str | None = None) -> Path:
    """Save an uploaded workbook to disk so Excel and pyxlsb can read it.

    The user's own copy is never touched: the tool only ever reads this staged
    copy inside the application folder. Re-staging identical bytes is a no-op,
    so a front end may call this on every redraw.
    """
    app_paths.ensure_directories()
    run_id = run_id or new_run_id(filename)
    target_dir = app_paths.WORK_DIR / run_id
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / Path(filename).name
    if target.exists() and target.stat().st_size == len(data) and target.read_bytes() == data:
        return target
    target.write_bytes(data)
    return target


def stage_configuration_file(data: bytes, filename: str) -> Path:
    """Save an uploaded alternative mapping or AKS V2 template."""
    return stage_source(data, filename, run_id=CONFIG_STAGE_ID)


def describe_configuration(
    mapping_path: Path | None = None, template_path: Path | None = None
) -> dict[str, Any]:
    """Which mapping and template are in use, and are they valid?"""
    mapping = Path(mapping_path) if mapping_path else app_paths.default_mapping()
    template = Path(template_path) if template_path else app_paths.default_template()
    return {
        "mapping_path": mapping,
        "template_path": template,
        "mapping_label": app_paths.version_label(mapping),
        "template_label": app_paths.version_label(template),
        "mapping_problems": validate_mapping(mapping),
        "template_problems": validate_template(template),
    }


# ---------------------------------------------------------------------------
# Step 3 — analyse (dry run)
# ---------------------------------------------------------------------------


def analyse_migration(
    source: Path | str,
    mapping_path: Path | str | None = None,
    template_path: Path | str | None = None,
    run_id: str | None = None,
    progress: ProgressCallback | None = None,
) -> AnalysisResult:
    """Dry run. Reads every file, writes none.

    Never raises for an expected migration problem: inspect `ok` and `errors`.
    """
    app_paths.ensure_directories()
    started = time.perf_counter()
    source = Path(source)
    mapping = Path(mapping_path) if mapping_path else app_paths.default_mapping()
    template = Path(template_path) if template_path else app_paths.default_template()
    run_id = run_id or new_run_id(source.name)
    log = RunLog(run_id, app_paths.LOGS_DIR)

    def report(message: str, fraction: float) -> None:
        if progress is not None:
            progress(message, fraction)

    result = AnalysisResult(
        ok=False,
        run_id=run_id,
        source_path=source,
        source_name=source.name,
        mapping_path=mapping,
        template_path=template,
        log_path=log.path,
        _log=log,
    )

    log.section("Analyse migration (dry run)")
    log.info(f"Run ID       : {run_id}")
    log.info(f"Started      : {datetime.now().isoformat(timespec='seconds')}")
    log.info(f"Source file  : {source}")
    log.info(f"Mapping file : {mapping}")
    log.info(f"V2 template  : {template}")

    try:
        report("Checking the input files", 0.05)
        problems: list[str] = []
        problems.extend(validate_mapping(mapping))
        problems.extend(validate_template(template))

        inspection = inspect_source(source)
        if not inspection.ok:
            problems.extend(inspection.problems)
        if problems:
            for message in problems:
                log.error(message)
            result.errors = problems
            result.log_text = log.text
            result.duration_seconds = time.perf_counter() - started
            log.close()
            return result

        log.info(
            f"Source sheet '{inspection.sheet}': {inspection.data_rows} equipment rows "
            f"({inspection.first_row}-{inspection.last_row}), "
            f"{inspection.placeholder_rows} pre-filled rows ignored"
        )

        report("Reading the migration mapping and the AKS V2 template", 0.25)
        entries, summary = build_plan(source, template, mapping)
        report("Comparing every legacy value against the mapping", 0.80)

        result.entries = entries
        result.summary = summary
        result._identities = reporting.row_identities(entries)
        result.counts = reporting.count_statuses(entries)
        result.warnings = list(summary.get("warnings", []))
        result.ok = True

        log.info(f"Active mapping rules   : {summary['active_rules']}")
        log.info(f"Legacy equipment rows  : {summary['source_rows']}")
        log.info(f"Mapped values examined : {summary['entries']}")
        log.info(f"Values to be written   : {summary['applied_writes']}")
        log.info(f"Engine status counts   : {summary['status_counts']}")
        log.info(f"Review view counts     : {result.counts}")
        for message in result.warnings:
            log.warning(message)
        report("Analysis complete", 1.0)

    except MappingResolutionError as exc:
        log.exception("Mapping resolution failed", exc)
        result.unresolved_mappings = exc.problems
        result.errors = [
            "The migration mapping could not be applied to this legacy file. "
            "The tool stopped instead of guessing which legacy column to read."
        ]
        result.errors.extend(
            f"{problem['target_key']} ({problem['target_label'] or 'no label'}): "
            f"legacy column {problem['legacy_label']!r} — {problem['reason']}"
            for problem in exc.problems
        )
    except MigrationError as exc:
        log.exception("Migration stopped", exc)
        result.errors = [str(exc)]
    except Exception as exc:  # noqa: BLE001 - the log keeps the traceback
        log.exception(f"Unexpected {type(exc).__name__} during analysis", exc)
        result.errors = [friendly_error(exc)]

    result.duration_seconds = time.perf_counter() - started
    log.info(f"Analysis finished in {result.duration_seconds:.1f} s (ok={result.ok})")
    result.log_text = log.text
    # Release the file handle. In the long-lived Streamlit process these would
    # otherwise accumulate and keep every run's log locked on Windows.
    # `run_migration` re-opens the same file to continue this run's log.
    log.close()
    return result


# ---------------------------------------------------------------------------
# Step 4 — run the migration
# ---------------------------------------------------------------------------


def planned_output_name(source_name: str) -> str:
    return f"{Path(source_name).stem}{MIGRATED_SUFFIX}.xlsb"


def previous_outputs(source_name: str) -> list[Path]:
    """Earlier migrations of the same legacy file, so the user can be warned."""
    wanted = planned_output_name(source_name).casefold()
    if not app_paths.OUTPUTS_DIR.exists():
        return []
    return sorted(
        (path for path in app_paths.OUTPUTS_DIR.glob("*/*.xlsb") if path.name.casefold() == wanted),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def run_migration(
    analysis: AnalysisResult | Path | str,
    mapping_path: Path | str | None = None,
    template_path: Path | str | None = None,
    highlight: bool = True,
    replace_template_data: bool = True,
    overwrite: bool = False,
    output_dir: Path | str | None = None,
    progress: ProgressCallback | None = None,
) -> MigrationResult:
    """Create the migrated workbook plus its audit and report.

    Accepts the `AnalysisResult` from `analyse_migration` (preferred, so the
    reviewed plan is exactly what gets written) or a source path, in which case
    the analysis is repeated internally.

    Never raises for an expected migration problem: inspect `ok` and `errors`.
    """
    app_paths.ensure_directories()
    started = time.perf_counter()

    def report(message: str, fraction: float) -> None:
        if progress is not None:
            progress(message, fraction)

    if not isinstance(analysis, AnalysisResult):
        report("Analysing the legacy workbook", 0.02)
        analysis = analyse_migration(
            analysis, mapping_path=mapping_path, template_path=template_path
        )

    result = MigrationResult(
        ok=False,
        run_id=analysis.run_id,
        source_name=analysis.source_name,
        summary=analysis.summary,
        counts=analysis.counts,
        warnings=list(analysis.warnings),
        log_path=analysis.log_path,
    )

    if not analysis.ok:
        result.errors = list(analysis.errors) or ["The analysis did not complete, so nothing was migrated."]
        result.log_text = analysis.log_text
        return result

    log = analysis._log or RunLog(analysis.run_id, app_paths.LOGS_DIR)
    # The analysis closed its handler, and the same analysis may be migrated
    # more than once. Re-attach so this run reaches the log file too.
    log.open()
    log.section("Run migration")

    # Each run gets its own output folder, so a migrated workbook is never
    # silently replaced by a later run - including a second run started from the
    # same analysis, which would otherwise land in the folder it already filled.
    target_dir = Path(output_dir) if output_dir else _unique_run_dir(analysis.run_id)
    target_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(analysis.source_name).stem
    workbook_path = target_dir / planned_output_name(analysis.source_name)
    audit_path = target_dir / f"{stem}_Migration_Audit.xlsx"
    audit_csv_path = target_dir / f"{stem}_Migration_Audit.csv"
    report_path = target_dir / f"{stem}_Migration_Report.xlsx"
    result.output_dir = target_dir

    log.info(f"Output folder  : {target_dir}")
    log.info(f"Output workbook: {workbook_path.name}")
    log.info(f"Highlighting   : {'on' if highlight else 'off'}")
    log.info(f"Replace template demonstration rows: {replace_template_data}")

    # Phase 1 - the migrated workbook. This is the only phase whose failure may
    # delete the output: a half-written workbook is worse than no workbook.
    try:
        report("Starting Microsoft Excel", 0.03)
        migrate_with_excel(
            analysis.template_path,
            workbook_path,
            analysis.entries,
            analysis.summary,
            replace_template_data=replace_template_data,
            highlight=highlight,
            overwrite=overwrite,
            progress=lambda message, fraction: report(message, 0.03 + 0.72 * fraction),
        )
        log.info(f"Excel wrote {analysis.summary.get('applied_writes', 0)} values")
        result.workbook_path = workbook_path
    except MigrationError as exc:
        log.exception("Migration stopped", exc)
        result.errors = [str(exc)]
        _cleanup_failed_output(workbook_path)
    except Exception as exc:  # noqa: BLE001 - the log keeps the traceback
        log.exception(f"Unexpected {type(exc).__name__} during migration", exc)
        result.errors = [friendly_error(exc)]
        _cleanup_failed_output(workbook_path)

    # Phase 2 - the checks and the side artefacts. The workbook now exists and
    # has cost several minutes of Excel time, so nothing below may delete it.
    # A failure here is a warning against a migration that did succeed, never
    # "migration failed, nothing was changed".
    if result.workbook_path is not None:
        verification: OutputVerification | None = None
        try:
            report("Verifying the migrated workbook", 0.80)
            verification = verify_migrated_workbook(workbook_path, analysis.summary)
            result.verification = verification
            if verification.ok:
                log.info(
                    f"Verified: rows {verification.expected_first_row}-{verification.actual_last_row} "
                    f"populated ({verification.populated_rows} rows), Migration_Report sheet present"
                )
            else:
                for problem in verification.problems:
                    log.warning(f"Verification: {problem}")
                result.warnings.extend(
                    f"Output check: {problem}" for problem in verification.problems
                )
        except Exception as exc:  # noqa: BLE001 - the log keeps the traceback
            log.exception("The migrated workbook could not be checked", exc)
            result.warnings.append(
                "The migrated workbook was created but could not be checked automatically. "
                "Open it and confirm the migrated rows before using it."
            )

        try:
            report("Writing the migration audit", 0.88)
            reporting.write_audit_xlsx(audit_path, analysis.entries, analysis._identities)
            result.audit_path = audit_path
            reporting.write_csv_report(audit_csv_path, analysis.entries)
            result.audit_csv_path = audit_csv_path
            log.info(f"Audit written: {audit_path.name} and {audit_csv_path.name}")
        except Exception as exc:  # noqa: BLE001 - the log keeps the traceback
            log.exception("The migration audit could not be written", exc)
            result.warnings.append(
                "The migrated workbook was created, but the migration audit could not be "
                "written. The same detail is on the Migration_Report sheet inside the workbook."
            )

        verified = verification is not None and verification.ok
        try:
            report("Writing the migration report", 0.94)
            context = {
                "timestamp": datetime.now().strftime("%d %B %Y %H:%M:%S"),
                "run_id": analysis.run_id,
                "result": "Completed" if verified else "Completed with output warnings",
                "source_name": analysis.source_name,
                "source_sheet": "Übersicht",
                "mapping_name": app_paths.version_label(analysis.mapping_path),
                "template_name": app_paths.version_label(analysis.template_path),
                "output_name": workbook_path.name,
                "log_name": Path(analysis.log_path).name,
                "verification": "yes — reopened and checked" if verified else "see warnings below",
                "summary": analysis.summary,
                "counts": analysis.counts,
                "warnings": result.warnings,
                "errors": result.errors,
            }
            reporting.write_report_xlsx(report_path, context, analysis.entries)
            result.report_path = report_path
            log.info(f"Report written: {report_path.name}")
        except Exception as exc:  # noqa: BLE001 - the log keeps the traceback
            log.exception("The migration report could not be written", exc)
            result.warnings.append(
                "The migrated workbook was created, but the migration report could not be written."
            )

        result.ok = True
        report("Migration complete", 1.0)

    result.duration_seconds = time.perf_counter() - started
    log.info(f"Migration finished in {result.duration_seconds:.1f} s (ok={result.ok})")
    result.log_text = log.text
    log.close()
    return result


def _unique_run_dir(run_id: str) -> Path:
    """An empty results folder for this run, never one that already holds output."""
    candidate = app_paths.OUTPUTS_DIR / run_id
    attempt = 2
    while any(candidate.glob("*.xlsb")):
        candidate = app_paths.OUTPUTS_DIR / f"{run_id}_{attempt}"
        attempt += 1
    return candidate


def verify_migrated_workbook(path: Path, summary: dict[str, Any]) -> OutputVerification:
    """Reopen a generated workbook and confirm the migrated rows are present."""
    from .validation import verify_output

    return verify_output(Path(path), summary)


def _cleanup_failed_output(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


def cleanup_staged_uploads(keep_run_id: str | None = None) -> None:
    """Remove staged copies of uploaded legacy workbooks from earlier runs.

    Uploaded mapping and template files (staged under `_config`) are kept: they
    belong to the current advanced-settings selection, not to one source file.
    """
    if not app_paths.WORK_DIR.exists():
        return
    protected = {keep_run_id, CONFIG_STAGE_ID}
    for directory in app_paths.WORK_DIR.iterdir():
        if directory.is_dir() and directory.name not in protected:
            shutil.rmtree(directory, ignore_errors=True)


__all__ = [
    "AnalysisResult",
    "MigrationResult",
    "EnvironmentReport",
    "SourceInspection",
    "OutputVerification",
    "analyse_migration",
    "run_migration",
    "inspect_source_file",
    "check_environment",
    "describe_configuration",
    "stage_source",
    "stage_configuration_file",
    "planned_output_name",
    "previous_outputs",
    "cleanup_staged_uploads",
]
