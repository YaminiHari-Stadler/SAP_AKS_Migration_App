"""Shared constants, data structures, and value helpers.

Everything in this module is lifted unchanged from the V1 pilot script
`aks_migrate.py`. Behaviour must stay identical: the pilot numbers in
PILOT_FINDINGS.md are the regression baseline for this tool.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from typing import Any


SOURCE_SHEET = "Übersicht"
TARGET_SHEET = "Übersicht"
MAPPING_SHEET = "Tabelle1"
VALUES_SHEET = "MIG-Werte Merkmale"
SOURCE_HEADER_ROWS = (17, 18, 2)
SOURCE_DATA_START_ROW = 20
TARGET_KEY_ROW = 24
TARGET_PARAMETER_ROW = 25
TARGET_DATA_START_ROW = 28

PLACEHOLDERS = {"", "-", "---", "----"}

# Explicit exceptions are safer than guessing when legacy labels changed.
# Columns are one-based Excel column numbers.
SOURCE_COLUMN_OVERRIDES = {
    "STATUS": 1,              # A
    "CUST_SPEC_DESI": 4,      # D: mapping says row 18; supplied legacy header is on row 17
    "CONV_DIRE": 26,          # Z: supplied legacy header is "reversierend / reversible"
    "VA_DRUM": 14,           # N: legacy template calls this SS Trommel
    "VA_KO_SW_FORM": 14,     # N: same legacy source drives a second SAP field
    "LIMI_SWIT": 48,          # AV: legacy header says quantity rather than X/---
    "COMMENT": 62,            # BJ: BI is Schmierpatrone in the supplied legacy file
}

# Columns whose legacy value identifies an equipment row for a human reader.
IDENTITY_KEYS = ("PROC_POSI_NUMB", "CUST_SPEC_DESI", "TITLE")

REPORT_COLUMNS = [
    "source_row", "target_row", "source_label", "source_cell", "old_value",
    "target_parameter", "target_key", "target_cell", "new_value", "status",
    "rule", "note", "apply",
]


@dataclass(frozen=True)
class Rule:
    mapping_row: int
    target_parameter: str
    target_label: str
    target_key: str
    legacy_row_hint: str
    legacy_label: str
    todo: str
    replacements: tuple[tuple[Any, Any], ...]


@dataclass
class MigrationEntry:
    source_row: int
    target_row: int
    source_label: str
    source_cell: str
    old_value: Any
    target_parameter: str
    target_key: str
    target_cell: str
    new_value: Any
    status: str
    rule: str
    note: str = ""
    apply: bool = True
    # Presentation-only metadata. These are not part of REPORT_COLUMNS, so the
    # audit produced by this tool stays byte-comparable with the V1 pilot CSV.
    target_label: str = ""
    has_mapping_table: bool = False


class MigrationError(RuntimeError):
    """A migration was stopped deliberately; the message is user-facing."""


def normalize(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else format(value, ".15g")
    text = unicodedata.normalize("NFKC", str(value)).casefold().strip()
    return "".join(ch for ch in text if ch.isalnum())


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def display(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def excel_col(number: int) -> str:
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result
