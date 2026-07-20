"""Excel export for consolidated TCR database, debug report and error report."""

from __future__ import annotations

from pathlib import Path
import pandas as pd

from .config import OUTPUT_COLUMNS
from .models import DebugRecord, ExtractedRow, ProcessingError, UnmappedColumnRecord


def _format_sheet(writer: pd.ExcelWriter, sheet_name: str) -> None:
    ws = writer.book[sheet_name]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for column_cells in ws.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column_letter].width = min(max_length + 2, 70)


def export_to_excel(
    rows: list[ExtractedRow],
    output_file: str | Path,
    unmapped_columns: list[UnmappedColumnRecord] | None = None,
) -> None:
    """Export extracted rows to a formatted Excel workbook.

    Backward compatible: callers can still pass only rows and output_file.
    """
    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame([row.values for row in rows], columns=OUTPUT_COLUMNS)

    unmapped_data = []
    for item in unmapped_columns or []:
        unmapped_data.append(
            {
                "Workbook": item.workbook,
                "Worksheet": item.worksheet,
                "Source Row": item.source_row,
                "Source Column": item.source_column,
                "Source Header": item.source_header,
                "Value": item.value,
            }
        )
    unmapped_df = pd.DataFrame(
        unmapped_data,
        columns=["Workbook", "Worksheet", "Source Row", "Source Column", "Source Header", "Value"],
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="TCR_Database")
        _format_sheet(writer, "TCR_Database")
        unmapped_df.to_excel(writer, index=False, sheet_name="Unmapped Columns")
        _format_sheet(writer, "Unmapped Columns")


def export_error_report(errors: list[ProcessingError], output_file: str | Path) -> Path:
    """Export workbook/sheet-level processing issues to Error_Report.xlsx."""
    output_path = Path(output_file)
    error_path = output_path.with_name("Error_Report.xlsx")
    data = [
        {"File": e.file, "Sheet": e.sheet, "Reason": e.reason, "Exception": e.exception}
        for e in errors
    ]
    df = pd.DataFrame(data, columns=["File", "Sheet", "Reason", "Exception"])
    with pd.ExcelWriter(error_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Errors")
        _format_sheet(writer, "Errors")
    return error_path


def export_debug_report(debug_records: list[DebugRecord], output_file: str | Path) -> Path:
    """Export worksheet-level diagnostics to Debug_Report.xlsx."""
    output_path = Path(output_file)
    debug_path = output_path.with_name("Debug_Report.xlsx")
    data = [
        {
            "Workbook": d.workbook,
            "Worksheet": d.worksheet,
            "Detected header row": d.detected_header_row,
            "Detected BOM row": d.detected_bom_row,
            "Detected columns": d.detected_columns,
            "Confidence score": d.confidence_score,
            "Confidence reasons": d.confidence_reasons,
            "Missing columns": d.missing_columns,
            "Validation status": d.validation_status,
        }
        for d in debug_records
    ]
    df = pd.DataFrame(
        data,
        columns=[
            "Workbook",
            "Worksheet",
            "Detected header row",
            "Detected BOM row",
            "Detected columns",
            "Confidence score",
            "Confidence reasons",
            "Missing columns",
            "Validation status",
        ],
    )
    with pd.ExcelWriter(debug_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Debug")
        _format_sheet(writer, "Debug")
    return debug_path
