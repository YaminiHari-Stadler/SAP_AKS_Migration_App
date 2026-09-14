"""Read-only access to .xlsb workbooks through pyxlsb.

Microsoft Excel is not involved here. The analyse/dry-run step therefore works
on any Windows machine, and Excel is only needed for the write phase.

A small cache keyed on (path, size, mtime) avoids re-parsing the 4 MB V2
template on every Streamlit rerun. The key includes the file stat, so replacing
the template on disk invalidates the entry automatically.
"""

from __future__ import annotations

from collections import OrderedDict
from pathlib import Path
from threading import Lock
from typing import Any

from pyxlsb import open_workbook

from .model import MigrationError


_CACHE: "OrderedDict[tuple[str, int, int, str], dict[int, dict[int, Any]]]" = OrderedDict()
_CACHE_LOCK = Lock()
_CACHE_MAX = 4


def sheet_names(path: Path) -> list[str]:
    try:
        with open_workbook(path) as workbook:
            return list(workbook.sheets)
    except MigrationError:
        raise
    except Exception as exc:  # pyxlsb raises bare exceptions for bad files
        raise MigrationError(
            f"{Path(path).name} could not be opened as an Excel binary workbook (.xlsb): {exc}"
        ) from exc


def _read_sheet(path: Path, sheet_name: str) -> dict[int, dict[int, Any]]:
    rows: dict[int, dict[int, Any]] = {}
    with open_workbook(path) as workbook:
        if sheet_name not in workbook.sheets:
            raise MigrationError(f"Sheet {sheet_name!r} not found in {Path(path).name}")
        with workbook.get_sheet(sheet_name) as sheet:
            for row in sheet.rows(sparse=True):
                if not row:
                    continue
                row_number = row[0].r + 1
                values = {
                    cell.c + 1: cell.v
                    for cell in row
                    if cell.v is not None and cell.v != ""
                }
                if values:
                    rows[row_number] = values
    return rows


def load_sheet(path: Path, sheet_name: str, use_cache: bool = True) -> dict[int, dict[int, Any]]:
    """Return {row_number: {column_number: value}} for one .xlsb worksheet."""
    path = Path(path)
    if not use_cache:
        return _read_sheet(path, sheet_name)

    try:
        stat = path.stat()
        key = (str(path.resolve()), stat.st_size, int(stat.st_mtime_ns), sheet_name)
    except OSError as exc:
        raise MigrationError(f"File not readable: {path}") from exc

    with _CACHE_LOCK:
        cached = _CACHE.get(key)
        if cached is not None:
            _CACHE.move_to_end(key)
            return cached

    rows = _read_sheet(path, sheet_name)

    with _CACHE_LOCK:
        _CACHE[key] = rows
        _CACHE.move_to_end(key)
        while len(_CACHE) > _CACHE_MAX:
            _CACHE.popitem(last=False)
    return rows


def clear_cache() -> None:
    with _CACHE_LOCK:
        _CACHE.clear()
