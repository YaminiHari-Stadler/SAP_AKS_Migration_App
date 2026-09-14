"""Where the application keeps its files.

Everything lives under the application folder so the tool can be copied to a
network share or a USB stick without any configuration.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent.parent

CONFIG_DIR = APP_ROOT / "config"
TEMPLATES_DIR = APP_ROOT / "templates"
OUTPUTS_DIR = APP_ROOT / "outputs"
LOGS_DIR = APP_ROOT / "logs"
WORK_DIR = APP_ROOT / "outputs" / "_uploads"

DEFAULT_MAPPING_NAME = "MigrationsMapping_00.xlsx"
DEFAULT_TEMPLATE_NAME = "AKS_V2_06.xlsb"


def _first_existing(*candidates: Path) -> Path:
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def default_mapping() -> Path:
    """The shipped migration mapping workbook."""
    return _first_existing(CONFIG_DIR / DEFAULT_MAPPING_NAME, APP_ROOT / DEFAULT_MAPPING_NAME)


def default_template() -> Path:
    """The shipped AKS V2 template."""
    return _first_existing(TEMPLATES_DIR / DEFAULT_TEMPLATE_NAME, APP_ROOT / DEFAULT_TEMPLATE_NAME)


def ensure_directories() -> None:
    for directory in (CONFIG_DIR, TEMPLATES_DIR, OUTPUTS_DIR, LOGS_DIR, WORK_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def local_staging_root() -> Path:
    """A plain local folder that Excel can write into.

    The application folder is often on OneDrive. Excel rewrites paths inside a
    synced folder to the SharePoint URL behind them, opens the workbook
    read-only, and then discards the save without reporting an error. The write
    phase therefore happens here, outside any synced location, and the finished
    file is moved into `outputs\\` afterwards.
    """
    candidate = os.environ.get("LOCALAPPDATA", "")
    if candidate and "onedrive" not in candidate.casefold():
        root = Path(candidate) / "AKS Migration App" / "work"
    else:
        root = Path(tempfile.gettempdir()) / "AKS Migration App" / "work"
    root.mkdir(parents=True, exist_ok=True)
    return root


def local_staging_dir(name: str) -> Path:
    """A unique local folder for one Excel write."""
    prune_local_staging()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    directory = local_staging_root() / f"{stamp}_{name[:40]}"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def prune_local_staging(max_age_hours: int = 24) -> None:
    """Remove staging folders left behind by an interrupted or failed write."""
    cutoff = time.time() - max_age_hours * 3600
    try:
        entries = list(local_staging_root().iterdir())
    except OSError:
        return
    for entry in entries:
        try:
            if entry.is_dir() and entry.stat().st_mtime < cutoff:
                shutil.rmtree(entry, ignore_errors=True)
        except OSError:
            continue


def version_label(path: Path) -> str:
    """A short 'which file am I actually using' description for the interface."""
    path = Path(path)
    if not path.exists():
        return f"{path.name} — missing"
    stat = path.stat()
    modified = datetime.fromtimestamp(stat.st_mtime).strftime("%d %b %Y %H:%M")
    return f"{path.name} — {stat.st_size / 1_048_576:.1f} MB, modified {modified}"
