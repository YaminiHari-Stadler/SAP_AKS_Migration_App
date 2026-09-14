"""Build the migration plan: which legacy cell becomes which SAP cell.

`build_plan` is the heart of the tool and is a verbatim carry-over of the V1
pilot logic, with two additive changes that do not affect any written value:

* unresolved mappings raise `MappingResolutionError` (a `MigrationError`) that
  also carries the problems as structured data for the user interface;
* each entry gets presentation-only metadata (`target_label`,
  `has_mapping_table`) so the review table can explain itself.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from .excel_reader import load_sheet
from .mapping import (
    build_header_candidates,
    build_target_columns,
    load_allowed_values,
    load_rules,
    load_status_codes,
    resolve_source_column,
    resolve_status_code,
    transform_value,
)
from .model import (
    PLACEHOLDERS,
    SOURCE_DATA_START_ROW,
    SOURCE_SHEET,
    TARGET_DATA_START_ROW,
    TARGET_SHEET,
    MigrationEntry,
    MigrationError,
    display,
    excel_col,
    is_blank,
    normalize,
)


class MappingResolutionError(MigrationError):
    """A legacy header or SAP target key could not be resolved safely."""

    def __init__(self, message: str, problems: list[dict[str, Any]]):
        super().__init__(message)
        self.problems = problems


def source_data_rows(source_rows: dict[int, dict[int, Any]]) -> tuple[list[int], int]:
    """Return the real equipment rows and the number of placeholder rows skipped."""
    result: list[int] = []
    skipped = 0
    for row_number in sorted(source_rows):
        if row_number < SOURCE_DATA_START_ROW:
            continue
        values = source_rows[row_number]
        # Legacy AKS workbooks pre-fill hundreds of unused rows with only the
        # project number and numeric defaults.  Those are not equipment rows.
        # Require at least one business-identifying field: status, designation,
        # equipment description, or equipment code.
        if any(
            not is_blank(values.get(column))
            and normalize(values.get(column)) not in PLACEHOLDERS
            for column in (1, 4, 6, 7)
        ):
            result.append(row_number)
        else:
            skipped += 1
    if not result:
        raise MigrationError(
            "No legacy AKS equipment rows were found on the 'Übersicht' sheet "
            f"from row {SOURCE_DATA_START_ROW} downwards."
        )
    return result, skipped


def build_plan(
    source_path: Path,
    template_path: Path,
    mapping_path: Path,
) -> tuple[list[MigrationEntry], dict[str, Any]]:
    rules = load_rules(mapping_path)
    source_rows = load_sheet(source_path, SOURCE_SHEET)
    target_rows = load_sheet(template_path, TARGET_SHEET)
    candidates = build_header_candidates(source_rows)
    target_columns = build_target_columns(target_rows)
    allowed_values = load_allowed_values(template_path)
    status_codes = load_status_codes(target_rows)

    resolved: dict[int, tuple[int, str]] = {}
    problems: list[dict[str, Any]] = []
    warnings: list[str] = []
    for rule in rules:
        source_column, reason = resolve_source_column(rule, candidates)
        if source_column is None:
            problems.append(
                {
                    "mapping_row": rule.mapping_row,
                    "target_key": rule.target_key,
                    "target_label": rule.target_label,
                    "legacy_label": rule.legacy_label,
                    "legacy_row_hint": rule.legacy_row_hint,
                    "reason": reason,
                    "message": (
                        f"Mapping row {rule.mapping_row}: source for {rule.target_key} "
                        f"({rule.legacy_label!r}, {rule.legacy_row_hint!r}) was not resolved"
                    ),
                }
            )
        elif normalize(rule.target_key) not in target_columns:
            problems.append(
                {
                    "mapping_row": rule.mapping_row,
                    "target_key": rule.target_key,
                    "target_label": rule.target_label,
                    "legacy_label": rule.legacy_label,
                    "legacy_row_hint": rule.legacy_row_hint,
                    "reason": "SAP target key is not present in the V2 template header row",
                    "message": (
                        f"Mapping row {rule.mapping_row}: target key {rule.target_key!r} was not found"
                    ),
                }
            )
        else:
            resolved[rule.mapping_row] = (source_column, reason)
            if reason.startswith("close header match"):
                warnings.append(
                    f"{rule.target_key}: legacy column matched approximately "
                    f"({reason.split(':', 1)[1].strip()}) instead of exactly."
                )
    if problems:
        raise MappingResolutionError(
            "Unsafe mapping resolution:\n- " + "\n- ".join(p["message"] for p in problems),
            problems,
        )

    entries: list[MigrationEntry] = []
    legacy_rows, placeholder_rows = source_data_rows(source_rows)
    for offset, source_row in enumerate(legacy_rows):
        target_row = TARGET_DATA_START_ROW + offset
        row_entries: list[MigrationEntry] = []
        for rule in rules:
            source_column, resolution_note = resolved[rule.mapping_row]
            old_value = source_rows[source_row].get(source_column)
            new_value, status, note = transform_value(rule, old_value)
            target_column = target_columns[normalize(rule.target_key)]

            if status != "SKIPPED" and normalize(rule.target_parameter) in allowed_values:
                allowed = allowed_values[normalize(rule.target_parameter)]
                if normalize(new_value) not in allowed and normalize(new_value) not in PLACEHOLDERS:
                    status = "MANUAL_REVIEW"
                    note = "value is not listed in the template's allowed SAP values"

            row_entries.append(
                MigrationEntry(
                    source_row=source_row,
                    target_row=target_row,
                    source_label=rule.legacy_label,
                    source_cell=f"{excel_col(source_column)}{source_row}",
                    old_value=old_value,
                    target_parameter=rule.target_parameter,
                    target_key=rule.target_key,
                    target_cell=f"{excel_col(target_column)}{target_row}",
                    new_value=new_value,
                    status=status,
                    rule=rule.todo,
                    note=f"{note}; {resolution_note}",
                    target_label=rule.target_label,
                    has_mapping_table=bool(rule.replacements),
                )
            )

            if rule.target_key == "STATUS" and status != "SKIPPED":
                status_nr_column = target_columns.get(normalize("STATUS_NR"))
                code, code_note = resolve_status_code(new_value, status_codes)
                code_status = status if code is not None else "MANUAL_REVIEW"
                row_entries.append(
                    MigrationEntry(
                        source_row=source_row,
                        target_row=target_row,
                        source_label=rule.legacy_label or "Status",
                        source_cell=f"{excel_col(source_column)}{source_row}",
                        old_value=old_value,
                        target_parameter="Z_VC_0201",
                        target_key="STATUS_NR",
                        target_cell=(
                            f"{excel_col(status_nr_column)}{target_row}" if status_nr_column else ""
                        ),
                        new_value=code,
                        status=code_status,
                        rule="derive numeric status code",
                        note=code_note,
                        target_label="Status code",
                        has_mapping_table=True,
                    )
                )

        by_target: dict[str, list[MigrationEntry]] = defaultdict(list)
        for entry in row_entries:
            by_target[entry.target_cell].append(entry)
        for target_cell, same_target in by_target.items():
            if len(same_target) == 1:
                entries.extend(same_target)
                continue
            active = [entry for entry in same_target if entry.status != "SKIPPED"]
            if len(active) <= 1:
                for entry in same_target:
                    entry.apply = entry in active
                entries.extend(same_target)
                continue
            distinct = {normalize(entry.new_value) for entry in active}
            if len(distinct) == 1:
                for index, entry in enumerate(active):
                    entry.apply = index == 0
                    if index:
                        entry.note += "; duplicate target value not written twice"
                entries.extend(same_target)
                continue
            for entry in same_target:
                entry.apply = False
                if entry in active:
                    entry.status = "MANUAL_REVIEW"
                    entry.note += "; conflicting source fields map to the same SAP target"
            conflict = MigrationEntry(
                source_row=source_row,
                target_row=target_row,
                source_label=" + ".join(entry.source_label for entry in active),
                source_cell=" + ".join(entry.source_cell for entry in active),
                old_value=" | ".join(display(entry.old_value) for entry in active),
                target_parameter=active[0].target_parameter,
                target_key=active[0].target_key,
                target_cell=target_cell,
                new_value=None,
                status="MANUAL_REVIEW",
                rule="resolve conflicting legacy fields",
                note="target left blank because multiple non-equivalent values were found",
                apply=True,
                target_label=active[0].target_label,
                has_mapping_table=active[0].has_mapping_table,
            )
            entries.extend(same_target)
            entries.append(conflict)

    template_last_data_row = max(
        (
            row_number
            for row_number, values in target_rows.items()
            if row_number >= TARGET_DATA_START_ROW
            and any(not is_blank(values.get(column)) for column in (3, 4, 12, 13))
        ),
        default=TARGET_DATA_START_ROW - 1,
    )
    if template_last_data_row >= TARGET_DATA_START_ROW:
        warnings.append(
            f"The V2 template contains demonstration rows through row {template_last_data_row}. "
            "They are removed from the generated copy; the template file itself is not changed."
        )
    if placeholder_rows:
        warnings.append(
            f"{placeholder_rows} pre-filled legacy rows contained no equipment data and were "
            "excluded from the migration."
        )

    summary = {
        "source": str(source_path),
        "template": str(template_path),
        "mapping": str(mapping_path),
        "active_rules": len(rules),
        "source_rows": len(legacy_rows),
        "placeholder_rows": placeholder_rows,
        "first_source_row": legacy_rows[0],
        "last_source_row": legacy_rows[-1],
        "first_target_row": TARGET_DATA_START_ROW,
        "last_target_row": TARGET_DATA_START_ROW + len(legacy_rows) - 1,
        "template_last_data_row": template_last_data_row,
        "entries": len(entries),
        "applied_writes": sum(
            1 for entry in entries if entry.apply and entry.status != "SKIPPED" and entry.target_cell
        ),
        "status_counts": dict(Counter(entry.status for entry in entries)),
        "warnings": warnings,
    }
    return entries, summary
