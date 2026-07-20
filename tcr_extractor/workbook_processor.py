"""Workbook-level orchestration."""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook

from .bom_detector import find_bom_table, iter_bom_rows, iter_unmapped_values
from .config import MANDATORY_FIELDS, OUTPUT_COLUMNS
from .header_reader import read_header_info
from .models import DebugRecord, ExtractedRow, ProcessingError, ProcessingResult, UnmappedColumnRecord
from .utils import (
    clean_part_number,
    extract_part_number_from_text,
    extract_region_from_path,
    extract_year_from_path,
    first_non_empty,
    is_empty,
)


ProgressCallback = Callable[[str], None]


def validate_row(row: dict[str, Any]) -> str:
    """Return validation string without stopping processing."""
    missing = [field for field in MANDATORY_FIELDS if is_empty(row.get(field, ""))]
    if not missing:
        return "OK"
    return "; ".join(f"Missing {field}" for field in missing)


def _build_output_row(
    *,
    header: dict[str, Any],
    bom: dict[str, Any],
    source_file: Path,
    sheet_name: str,
) -> ExtractedRow:
    """Create one normalized output row from header + BOM data."""
    comments: list[str] = []

    reference_part_number = bom.get("Reference Part Number", "")
    direct_part_number = bom.get("Part Number", "")
    part_name = bom.get("Part Name", "")

    part_number = clean_part_number(first_non_empty(reference_part_number, direct_part_number))
    cleaned_part_name = part_name

    if is_empty(part_number):
        extracted_pn, cleaned_name = extract_part_number_from_text(part_name)
        part_number = extracted_pn
        cleaned_part_name = cleaned_name
        if extracted_pn:
            comments.append("Part Number extracted from Part Name")
    elif not is_empty(reference_part_number):
        comments.append("Reference Part Number used")

    row = {
        "Project": header.get("Project", ""),
        "SOP": header.get("SOP", ""),
        "FOTP": header.get("FOTP", ""),
        "MOB Presentation Date": header.get("MOB Presentation Date", ""),
        "Production Plant": header.get("Production Plant", ""),
        "EOP": header.get("EOP", ""),
        "Volume": header.get("Volume", ""),
        "Part Number": part_number,
        "Complexity Category": bom.get("Complexity Category", ""),
        "Final Decision Board": bom.get("Final Decision Board", ""),
        "Part Name": cleaned_part_name,
        "Part Size X": bom.get("Part Size X", ""),
        "Part Size Y": bom.get("Part Size Y", ""),
        "Part Size Z": bom.get("Part Size Z", ""),
        "Material": bom.get("Material", ""),
        "Material Key Number": bom.get("Material Key Number", ""),
        "Surface": bom.get("Surface", ""),
        "Metallized": bom.get("Metallized", ""),
        "Special Surface": bom.get("Special Surface", ""),
        "Weight": bom.get("Weight", ""),
        "Machine/Tonnage": bom.get("Machine/Tonnage", ""),
        "Process (1K/2K/3K)": bom.get("Process (1K/2K/3K)", ""),
        "Source File": str(source_file),
        "Sheet Name": sheet_name,
        "Year": extract_year_from_path(source_file),
        "Region": extract_region_from_path(source_file),
        "Validation Status": "",
        "Comments": "; ".join(comments),
    }
    row["Validation Status"] = validate_row(row)
    return ExtractedRow({column: row.get(column, "") for column in OUTPUT_COLUMNS})


def _worksheet_validation_status(rows: list[ExtractedRow], prior_status: str = "") -> str:
    """Create debug validation summary for one worksheet."""
    if prior_status:
        return prior_status
    if not rows:
        return "No rows exported"
    statuses = sorted({r.values.get("Validation Status", "") for r in rows if r.values.get("Validation Status", "")})
    return "; ".join(statuses) if statuses else "OK"


def _log_step(logger, workbook: Path, worksheet: str | None, message: str, callback: ProgressCallback | None = None) -> None:
    """Log a diagnostic step and optionally notify the GUI."""
    sheet_text = worksheet if worksheet else "-"
    full_message = f"{message} | workbook={workbook.name} | worksheet={sheet_text}"
    if logger:
        logger.info(full_message)
    if callback:
        callback(full_message)


def process_workbook(
    path: str | Path,
    logger=None,
    debug_mode: bool = False,
    progress_callback: ProgressCallback | None = None,
) -> ProcessingResult:
    """Process a single workbook and never raise errors to the caller.

    Args:
        debug_mode: when True, populate worksheet-level diagnostics.
        progress_callback: optional GUI callback for live diagnostics.
    """
    start = time.perf_counter()
    source_file = Path(path)
    result = ProcessingResult(source_file=str(source_file))

    wb = None
    try:
        _log_step(logger, source_file, None, "Opening workbook...", progress_callback)
        wb = load_workbook(source_file, data_only=True, read_only=True)
        _log_step(logger, source_file, None, "Workbook opened.", progress_callback)
    except Exception as exc:
        err = ProcessingError(str(source_file), "", "Cannot open workbook", repr(exc))
        result.errors.append(err)
        if logger:
            logger.exception("Cannot open workbook | workbook=%s | worksheet=-", source_file.name)
        result.processing_seconds = time.perf_counter() - start
        return result

    try:
        for ws in wb.worksheets:
            sheet_rows_before = len(result.rows)
            debug_record = DebugRecord(workbook=str(source_file), worksheet=ws.title)
            _log_step(logger, source_file, ws.title, f"Processing worksheet: {ws.title}", progress_callback)
            try:
                _log_step(logger, source_file, ws.title, "Reading header...", progress_callback)
                header, detected_header_rows = read_header_info(ws)
                _log_step(logger, source_file, ws.title, "Header finished.", progress_callback)

                _log_step(logger, source_file, ws.title, "Detecting BOM...", progress_callback)
                header_row, mapping, unmapped_headers, score, reasons, missing_columns = find_bom_table(ws)
                _log_step(logger, source_file, ws.title, "BOM detected.", progress_callback)

                debug_record.detected_header_row = detected_header_rows
                debug_record.detected_bom_row = str(header_row or "")
                debug_record.detected_columns = ", ".join(sorted(mapping.keys()))
                debug_record.confidence_score = score
                debug_record.confidence_reasons = "; ".join(reasons)
                debug_record.missing_columns = ", ".join(missing_columns)

                if header_row is None:
                    debug_record.validation_status = "No BOM table detected"
                    result.errors.append(
                        ProcessingError(
                            str(source_file),
                            ws.title,
                            "No BOM table detected",
                            f"best_score={score}; reasons={'; '.join(reasons)}",
                        )
                    )
                    if logger:
                        logger.warning(
                            "No BOM selected | workbook=%s | worksheet=%s | best_score=%s | reasons=%s",
                            source_file.name,
                            ws.title,
                            score,
                            "; ".join(reasons),
                        )
                    _log_step(logger, source_file, ws.title, "Worksheet finished.", progress_callback)
                    continue

                if logger:
                    logger.info(
                        "Selected BOM table | workbook=%s | worksheet=%s | row=%s | confidence=%s | reasons=%s",
                        source_file.name,
                        ws.title,
                        header_row,
                        score,
                        "; ".join(reasons),
                    )

                _log_step(logger, source_file, ws.title, "Reading BOM rows...", progress_callback)
                extracted_count = 0
                for source_row, bom_row in iter_bom_rows(ws, header_row, mapping):
                    result.rows.append(
                        _build_output_row(
                            header=header,
                            bom=bom_row,
                            source_file=source_file,
                            sheet_name=ws.title,
                        )
                    )
                    extracted_count += 1

                _log_step(logger, source_file, ws.title, f"Rows exported: {extracted_count}", progress_callback)

                for src_row, src_col, src_header, value in iter_unmapped_values(ws, header_row, unmapped_headers):
                    result.unmapped_columns.append(
                        UnmappedColumnRecord(
                            workbook=str(source_file),
                            worksheet=ws.title,
                            source_row=src_row,
                            source_column=src_col,
                            source_header=src_header,
                            value=value,
                        )
                    )

                sheet_rows = result.rows[sheet_rows_before:]
                debug_record.validation_status = _worksheet_validation_status(sheet_rows)

                if logger and extracted_count:
                    logger.info(
                        "Extracted rows summary | workbook=%s | worksheet=%s | rows=%s | bom_row=%s | confidence=%s",
                        source_file.name,
                        ws.title,
                        extracted_count,
                        header_row,
                        score,
                    )

            except Exception as exc:
                debug_record.validation_status = "Sheet processing failed"
                result.errors.append(ProcessingError(str(source_file), ws.title, "Sheet processing failed", repr(exc)))
                if logger:
                    logger.exception("Sheet failed | workbook=%s | worksheet=%s", source_file.name, ws.title)
            finally:
                _log_step(logger, source_file, ws.title, "Worksheet finished.", progress_callback)
                if debug_mode:
                    result.debug_records.append(debug_record)

        result.processed_successfully = len(result.rows) > 0
        result.processing_seconds = time.perf_counter() - start
        _log_step(logger, source_file, None, "Workbook finished.", progress_callback)
        return result

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
