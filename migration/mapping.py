"""Mapping rules, column resolution, and value conversion.

This is the migration specification in code form. Every function here is a
verbatim carry-over from the V1 pilot script. Nothing in this module may guess:
when a legacy header or a SAP target key cannot be resolved with confidence the
caller is expected to stop.
"""

from __future__ import annotations

import re
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import openpyxl

from .excel_reader import load_sheet
from .model import (
    MAPPING_SHEET,
    SOURCE_COLUMN_OVERRIDES,
    SOURCE_HEADER_ROWS,
    TARGET_DATA_START_ROW,
    TARGET_KEY_ROW,
    VALUES_SHEET,
    MigrationError,
    Rule,
    display,
    is_blank,
    normalize,
)


def load_rules(path: Path) -> list[Rule]:
    """Read the active migration rules from MigrationsMapping_00.xlsx."""
    workbook = openpyxl.load_workbook(path, data_only=False, read_only=False)
    try:
        if MAPPING_SHEET not in workbook.sheetnames:
            raise MigrationError(f"Sheet {MAPPING_SHEET!r} not found in {Path(path).name}")
        sheet = workbook[MAPPING_SHEET]
        rules: list[Rule] = []
        inherited_target: tuple[str, str, str] | None = None
        for row in range(3, min(sheet.max_row, 1000) + 1):
            todo = sheet.cell(row, 6).value
            target_key = sheet.cell(row, 3).value
            if target_key:
                inherited_target = (
                    display(sheet.cell(row, 1).value).strip(),
                    display(sheet.cell(row, 2).value).strip(),
                    display(target_key).strip(),
                )
            if not todo:
                continue
            if target_key:
                target_parameter, target_label, resolved_target_key = inherited_target
            elif inherited_target:
                target_parameter, target_label, resolved_target_key = inherited_target
            else:
                raise MigrationError(f"Mapping row {row} has a ToDo but no target field")
            replacements: list[tuple[Any, Any]] = []
            if "replace" in str(todo).casefold() or "line color" in str(todo).casefold():
                for column in range(7, 25):
                    source_value = sheet.cell(row, column).value
                    target_value = sheet.cell(row + 1, column).value
                    if source_value not in (None, "") and target_value not in (None, ""):
                        replacements.append((source_value, target_value))
            rules.append(
                Rule(
                    mapping_row=row,
                    target_parameter=target_parameter,
                    target_label=target_label,
                    target_key=resolved_target_key,
                    legacy_row_hint=display(sheet.cell(row, 4).value).strip(),
                    legacy_label=display(sheet.cell(row, 5).value).strip(),
                    todo=display(todo).strip(),
                    replacements=tuple(replacements),
                )
            )
    finally:
        workbook.close()
    if not rules:
        raise MigrationError(f"No active migration rules found in {Path(path).name}")
    return rules


def build_target_columns(target_rows: dict[int, dict[int, Any]]) -> dict[str, int]:
    result: dict[str, int] = {}
    for column, value in target_rows.get(TARGET_KEY_ROW, {}).items():
        key = display(value).strip()
        if key:
            result[normalize(key)] = column
    return result


def build_header_candidates(
    source_rows: dict[int, dict[int, Any]]
) -> dict[str, list[tuple[int, int, str]]]:
    candidates: dict[str, list[tuple[int, int, str]]] = defaultdict(list)
    for row_number in SOURCE_HEADER_ROWS:
        for column, value in source_rows.get(row_number, {}).items():
            key = normalize(value)
            if key:
                candidates[key].append((row_number, column, display(value)))
    return candidates


def hinted_rows(hint: str) -> tuple[int, ...]:
    numbers = tuple(int(n) for n in re.findall(r"\d+", hint))
    return numbers or SOURCE_HEADER_ROWS


def resolve_source_column(
    rule: Rule, candidates: dict[str, list[tuple[int, int, str]]]
) -> tuple[int | None, str]:
    override = SOURCE_COLUMN_OVERRIDES.get(rule.target_key)
    if override:
        return override, "explicit legacy-layout exception"

    wanted = normalize(rule.legacy_label)
    if not wanted:
        return None, "mapping has no legacy label"
    permitted_rows = set(hinted_rows(rule.legacy_row_hint))
    exact = [item for item in candidates.get(wanted, []) if item[0] in permitted_rows]
    if len(exact) == 1:
        return exact[0][1], "exact header match"
    if len(exact) > 1 and len({item[1] for item in exact}) == 1:
        return exact[0][1], "same header repeated in the same column"

    # Conservative fallback: only accept a unique, very close match.
    scored: list[tuple[float, int, str]] = []
    for key, items in candidates.items():
        for row_number, column, label in items:
            if row_number in permitted_rows:
                scored.append((SequenceMatcher(None, wanted, key).ratio(), column, label))
    scored.sort(reverse=True)
    if scored and scored[0][0] >= 0.92:
        if len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.05:
            return scored[0][1], f"close header match: {scored[0][2]!r}"
    return None, "legacy header not resolved safely"


def load_allowed_values(template_path: Path) -> dict[str, set[str]]:
    rows = load_sheet(template_path, VALUES_SHEET)
    allowed: dict[str, set[str]] = defaultdict(set)
    for values in rows.values():
        parameter = values.get(4)  # D
        value = values.get(6)      # F
        if parameter not in (None, "") and value not in (None, ""):
            allowed[normalize(parameter)].add(normalize(value))
    return allowed


def load_status_codes(target_rows: dict[int, dict[int, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for row_number in range(3, 20):
        code = target_rows.get(row_number, {}).get(5)   # E
        label = target_rows.get(row_number, {}).get(7)  # G
        if label not in (None, "") and code not in (None, ""):
            result[normalize(label)] = code
    # The supplied template also contains status examples in A:B.
    for row_number in range(TARGET_DATA_START_ROW, TARGET_DATA_START_ROW + 40):
        code = target_rows.get(row_number, {}).get(1)
        label = target_rows.get(row_number, {}).get(2)
        if label not in (None, "") and code not in (None, ""):
            result[normalize(label)] = code
    return result


def resolve_status_code(label: Any, status_codes: dict[str, Any]) -> tuple[Any | None, str]:
    wanted = normalize(label)
    if wanted in status_codes:
        return status_codes[wanted], "exact status-code match"
    scored = sorted(
        (
            (SequenceMatcher(None, wanted, key).ratio(), code, key)
            for key, code in status_codes.items()
        ),
        key=lambda item: item[0],
        reverse=True,
    )
    if scored and scored[0][0] >= 0.72 and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.06):
        return scored[0][1], f"status-code label matched at {scored[0][0]:.0%}"
    return None, "status code not resolved safely"


def transform_value(rule: Rule, old_value: Any) -> tuple[Any, str, str]:
    if is_blank(old_value):
        return None, "SKIPPED", "source is blank"

    comparison_value = old_value
    if rule.target_key == "STATUS" and isinstance(old_value, str):
        comparison_value = old_value.split("/", 1)[0].strip()

    comparison_key = normalize(comparison_value)
    for source_value, target_value in sorted(
        rule.replacements, key=lambda pair: len(normalize(pair[0])), reverse=True
    ):
        source_key = normalize(source_value)
        matches = comparison_key == source_key
        if rule.target_key == "STATUS":
            matches = matches or comparison_key.startswith(source_key)
        if matches:
            return target_value, "CONVERTED", f"{display(source_value)} → {display(target_value)}"

    new_value = old_value
    note = "copied without conversion"
    if "semicolon" in rule.todo.casefold() or "semikolon" in rule.todo.casefold():
        if isinstance(new_value, str):
            cleaned = new_value.replace(";", ".").replace(",", ".")
            if cleaned != new_value:
                return cleaned, "CONVERTED", "replaced semicolons and commas with periods"
    return new_value, "COPIED", note
