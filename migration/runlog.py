"""One log file per migration run.

Normal users must never see a Python traceback. The traceback goes here; the
user interface shows the friendly message and the path to this file.
"""

from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path

from .model import MigrationError


def new_run_id(source_name: str = "") -> str:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = Path(source_name).stem if source_name else ""
    slug = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("_")[:48]
    return f"{stamp}_{slug}" if slug else stamp


class RunLog:
    """A logger writing to logs/<run_id>.log, plus an in-memory copy for the UI."""

    def __init__(self, run_id: str, log_dir: Path):
        self.run_id = run_id
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.path = self.log_dir / f"{run_id}.log"
        self.lines: list[str] = []

        self._logger = logging.getLogger(f"aks_migration.{run_id}")
        self._logger.setLevel(logging.DEBUG)
        self._logger.propagate = False
        self.open()

    def open(self) -> None:
        """Attach a file handler, or re-attach one after `close`.

        A run log is written in two stages - the analysis, then the migration -
        and an analysis can be migrated more than once. Without this, the second
        stage would log into a closed handler and the file on disk would silently
        disagree with what the interface shows.
        """
        for handler in list(self._logger.handlers):
            self._logger.removeHandler(handler)
            handler.close()
        # Append mode: re-opening continues the same run's log.
        handler = logging.FileHandler(self.path, mode="a", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s"))
        self._logger.addHandler(handler)

    # -- writing -----------------------------------------------------------

    def _record(self, level: str, message: str) -> None:
        self.lines.append(f"{datetime.now().strftime('%H:%M:%S')}  {level:<8}  {message}")

    def info(self, message: str) -> None:
        self._logger.info(message)
        self._record("INFO", message)

    def warning(self, message: str) -> None:
        self._logger.warning(message)
        self._record("WARNING", message)

    def error(self, message: str) -> None:
        self._logger.error(message)
        self._record("ERROR", message)

    def exception(self, message: str, exc: BaseException) -> None:
        self._logger.error(message, exc_info=exc)
        self._record("ERROR", message)

    def section(self, title: str) -> None:
        self.info(f"--- {title} ---")

    def close(self) -> None:
        for handler in list(self._logger.handlers):
            self._logger.removeHandler(handler)
            handler.close()

    # -- reading -----------------------------------------------------------

    @property
    def text(self) -> str:
        return "\n".join(self.lines)


def friendly_error(exc: BaseException) -> str:
    """Translate an exception into something an engineer without Python can act on."""
    if isinstance(exc, MigrationError):
        return str(exc)

    message = str(exc).strip()
    name = type(exc).__name__

    if isinstance(exc, PermissionError) or "being used by another process" in message:
        return (
            "A file is locked. Close the legacy AKS file, the AKS V2 template, and any "
            "previously migrated workbook in Excel, then try again."
        )
    if isinstance(exc, FileNotFoundError):
        return f"A required file was not found: {message}"
    if "com_error" in name.casefold() or "-2147" in message:
        return (
            "Microsoft Excel refused the request. Close all Excel windows, end any leftover "
            "EXCEL.EXE process in Task Manager, and run the migration again. If it keeps "
            "happening, the workbook may be blocked by the Excel Trust Center — contact IT."
        )
    if isinstance(exc, MemoryError):
        return "The machine ran out of memory while processing this workbook."
    return (
        "The migration stopped because of an unexpected problem. No files were changed. "
        "Please send the log file to the tool maintainer."
    )
