"""Data models used by the extractor."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ExtractedRow:
    """Single consolidated output row."""

    values: dict[str, Any] = field(default_factory=dict)


@dataclass
class ProcessingError:
    """Structured error that is later exported to Error_Report.xlsx."""

    file: str
    sheet: str = ""
    reason: str = ""
    exception: str = ""


@dataclass
class DebugRecord:
    """Worksheet-level diagnostics for Debug_Report.xlsx."""

    workbook: str
    worksheet: str
    detected_header_row: str = ""
    detected_bom_row: str = ""
    detected_columns: str = ""
    confidence_score: int = 0
    confidence_reasons: str = ""
    missing_columns: str = ""
    validation_status: str = ""


@dataclass
class UnmappedColumnRecord:
    """Value captured from a BOM column that did not map to a canonical field."""

    workbook: str
    worksheet: str
    source_row: int
    source_column: int
    source_header: str
    value: Any


@dataclass
class ProcessingResult:
    """Result of processing one workbook."""

    source_file: str
    rows: list[ExtractedRow] = field(default_factory=list)
    errors: list[ProcessingError] = field(default_factory=list)
    debug_records: list[DebugRecord] = field(default_factory=list)
    unmapped_columns: list[UnmappedColumnRecord] = field(default_factory=list)
    processed_successfully: bool = False
    processing_seconds: float = 0.0
