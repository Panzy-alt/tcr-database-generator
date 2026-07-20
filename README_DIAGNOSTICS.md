# TCR diagnostics build

Changes made only for diagnostics:

- GUI now shows `Files processed: 0/<total_files>` immediately after scanning files and before the first workbook starts.
- Log file is created automatically as `extractor.log` next to the selected output Excel file.
- `workbook_processor.py` logs before and after major stages:
  - Opening workbook...
  - Workbook opened.
  - Processing worksheet: <sheet>
  - Reading header...
  - Header finished.
  - Detecting BOM...
  - BOM detected.
  - Reading BOM rows...
  - Rows exported: X
  - Worksheet finished.
  - Workbook finished.
- GUI status and log panel update after each diagnostic step via the existing background thread.
- Existing extraction logic is unchanged.
