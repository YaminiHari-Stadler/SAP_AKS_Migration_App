"""Command-line interface for the AKS migration engine.

The engine itself now lives in the `migration` package and is shared with the
Streamlit application (`app.py`). This script keeps the V1 pilot command line
working unchanged, so the runbook commands and the recorded pilot numbers stay
reproducible:

    python aks_migrate.py --source "legacy.xlsb" --dry-run --report "audit.csv"

`--template` and `--mapping` are now optional and default to the files shipped
with the application in `templates\\` and `config\\`.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

from migration.engine import build_plan
from migration.excel_writer import migrate_with_excel
from migration.model import MigrationError
from migration.paths import default_mapping, default_template
from migration.reporting import write_csv_report


def sources_from_args(args: argparse.Namespace) -> list[Path]:
    if args.source:
        return [args.source]
    sources = sorted(args.input_dir.glob("*.xlsb"))
    template = args.template.resolve()
    return [path for path in sources if path.resolve() != template]


def output_for_source(args: argparse.Namespace, source: Path) -> Path:
    if args.output:
        return args.output
    return args.output_dir / f"{source.stem}_migrated.xlsb"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Migrate legacy AKS .xlsb workbooks")
    source_group = parser.add_mutually_exclusive_group(required=True)
    source_group.add_argument("--source", type=Path, help="one legacy AKS .xlsb file")
    source_group.add_argument("--input-dir", type=Path, help="directory of legacy AKS .xlsb files")
    parser.add_argument(
        "--template", type=Path, default=None,
        help="AKS V2 template (default: templates/AKS_V2_06.xlsb)",
    )
    parser.add_argument(
        "--mapping", type=Path, default=None,
        help="migration mapping (default: config/MigrationsMapping_00.xlsx)",
    )
    output_group = parser.add_mutually_exclusive_group()
    output_group.add_argument("--output", type=Path, help="output path for a single source")
    output_group.add_argument("--output-dir", type=Path, help="output directory, required for batch mode")
    parser.add_argument("--dry-run", action="store_true", help="validate and report without creating a workbook")
    parser.add_argument("--report", type=Path, help="CSV report path for a single-source dry run")
    parser.add_argument("--replace-template-data", action="store_true", help="replace demonstration rows in the template")
    parser.add_argument("--highlight", action="store_true", help="color migrated target cells by status")
    parser.add_argument("--overwrite", action="store_true", help="replace an existing output file")
    args = parser.parse_args(argv)

    if args.template is None:
        args.template = default_template()
    if args.mapping is None:
        args.mapping = default_mapping()
    if args.input_dir and not args.output_dir and not args.dry_run:
        parser.error("--output-dir is required with --input-dir")
    if args.source and not args.output and not args.dry_run:
        parser.error("--output is required with --source")
    if args.report and not args.source:
        parser.error("--report is supported only with --source")
    return args


def validate_inputs(paths: Iterable[Path]) -> None:
    for path in paths:
        if not path.exists() or not path.is_file():
            raise MigrationError(f"Input file not found: {path}")


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        validate_inputs([args.template, args.mapping])
        sources = sources_from_args(args)
        if not sources:
            raise MigrationError("No legacy .xlsb source files found")
        validate_inputs(sources)

        for source in sources:
            entries, summary = build_plan(source, args.template, args.mapping)
            if args.dry_run:
                report_path = args.report
                if report_path is None and args.output_dir:
                    report_path = args.output_dir / f"{source.stem}_migration_report.csv"
                if report_path:
                    write_csv_report(report_path, entries)
                print(f"DRY RUN OK: {source.name}")
            else:
                output = output_for_source(args, source)
                migrate_with_excel(
                    args.template, output, entries, summary,
                    replace_template_data=args.replace_template_data,
                    highlight=args.highlight,
                    overwrite=args.overwrite,
                )
                print(f"CREATED: {output}")
            print(
                f"  rows={summary['source_rows']} rules={summary['active_rules']} "
                f"entries={summary['entries']} statuses={summary['status_counts']}"
            )
        return 0
    except MigrationError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"UNEXPECTED ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
