"""Recursive folder scanning."""

from __future__ import annotations

from pathlib import Path


def find_excel_files(root_folder: str | Path) -> list[Path]:
    """Return all .xlsx files recursively, ignoring temporary Office files."""
    root = Path(root_folder)
    files = []
    for path in root.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        files.append(path)
    return sorted(files)
