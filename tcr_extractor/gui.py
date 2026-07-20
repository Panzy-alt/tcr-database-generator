"""Tkinter Windows desktop user interface."""

from __future__ import annotations

import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .exporter import export_debug_report, export_error_report, export_to_excel
from .logger import setup_logger
from .scanner import find_excel_files
from .workbook_processor import process_workbook


def _format_duration(seconds: float) -> str:
    seconds = max(0, int(seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h}h {m}m {s}s"
    if m:
        return f"{m}m {s}s"
    return f"{s}s"


def _safe_rate(count: int, seconds: float) -> float:
    return count / seconds if seconds > 0 else 0.0


class TcrExtractorApp(tk.Tk):
    """Desktop GUI for the TCR extraction workflow."""

    def __init__(self):
        super().__init__()
        self.title("TCR Desktop Extractor")
        self.geometry("920x640")
        self.minsize(840, 560)

        self.input_folder = tk.StringVar()
        self.output_file = tk.StringVar()
        self.debug_mode = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar(value="Select an input folder and output file.")
        self.current_file_text = tk.StringVar(value="Current file: -")
        self.files_processed_text = tk.StringVar(value="Files processed: 0/0")
        self.eta_text = tk.StringVar(value="Estimated remaining time: -")
        self.rows_exported_text = tk.StringVar(value="Rows exported: 0")

        self._build_ui()

    def _build_ui(self) -> None:
        padding = {"padx": 10, "pady": 7}

        title = ttk.Label(self, text="Tool Capacity Report Extractor", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w", padx=14, pady=(14, 4))

        frame = ttk.Frame(self)
        frame.pack(fill="x", **padding)

        ttk.Label(frame, text="Input folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.input_folder).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(frame, text="Browse", command=self._choose_folder).grid(row=0, column=2)

        ttk.Label(frame, text="Output Excel:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_file).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(frame, text="Save as", command=self._choose_output).grid(row=1, column=2)
        frame.columnconfigure(1, weight=1)

        options_frame = ttk.Frame(self)
        options_frame.pack(fill="x", padx=14, pady=(0, 4))
        ttk.Checkbutton(
            options_frame,
            text="Debug Mode - generate Debug_Report.xlsx",
            variable=self.debug_mode,
        ).pack(anchor="w")

        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", **padding)
        self.run_button = ttk.Button(action_frame, text="Run extraction", command=self._start_processing)
        self.run_button.pack(side="left")
        self.progress = ttk.Progressbar(action_frame, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=12)

        progress_frame = ttk.LabelFrame(self, text="Progress")
        progress_frame.pack(fill="x", padx=14, pady=(0, 8))
        for idx, var in enumerate([
            self.current_file_text,
            self.files_processed_text,
            self.eta_text,
            self.rows_exported_text,
        ]):
            ttk.Label(progress_frame, textvariable=var).grid(row=idx // 2, column=idx % 2, sticky="w", padx=10, pady=4)
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.columnconfigure(1, weight=1)

        ttk.Label(self, textvariable=self.status_text).pack(anchor="w", padx=14, pady=(0, 6))

        self.log_box = tk.Text(self, height=18, wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.log_box.configure(state="disabled")

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder containing TCR files")
        if folder:
            self.input_folder.set(folder)
            if not self.output_file.get():
                self.output_file.set(str(Path(folder) / "TCR_Consolidated_Database.xlsx"))

    def _choose_output(self) -> None:
        output = filedialog.asksaveasfilename(
            title="Save consolidated database as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if output:
            self.output_file.set(output)

    def _append_log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _start_processing(self) -> None:
        input_folder = self.input_folder.get().strip()
        output_file = self.output_file.get().strip()

        if not input_folder:
            messagebox.showwarning("Missing input", "Please select an input folder.")
            return
        if not output_file:
            messagebox.showwarning("Missing output", "Please select an output Excel file.")
            return

        self.run_button.configure(state="disabled")
        self.progress["value"] = 0
        self.status_text.set("Processing...")
        self._append_log("Starting extraction...")

        thread = threading.Thread(
            target=self._process,
            args=(input_folder, output_file, self.debug_mode.get()),
            daemon=True,
        )
        thread.start()

    def _update_progress(self, *, current_file: Path | None, index: int, total: int, start_time: float, rows: int) -> None:
        elapsed = time.perf_counter() - start_time
        avg_per_file = elapsed / index if index else 0
        remaining = avg_per_file * max(total - index, 0)

        self.progress.configure(value=index, maximum=max(total, 1))
        self.current_file_text.set(f"Current file: {current_file.name if current_file else '-'}")
        self.files_processed_text.set(f"Files processed: {index}/{total}")
        self.eta_text.set(f"Estimated remaining time: {_format_duration(remaining) if index else '-'}")
        self.rows_exported_text.set(f"Rows exported: {rows}")
        self.status_text.set(f"Processed {index}/{total} files")

    def _process(self, input_folder: str, output_file: str, debug_mode: bool) -> None:
        start_time = time.perf_counter()
        try:
            log_file = Path(output_file).with_suffix(".log")
            logger = setup_logger(log_file)

            excel_files = find_excel_files(input_folder)
            total_files = len(excel_files)
            all_rows = []
            all_errors = []
            all_debug_records = []
            all_unmapped_columns = []
            files_successful = 0
            files_skipped = 0
            workbook_seconds_total = 0.0

            self.after(0, lambda: self.progress.configure(maximum=max(total_files, 1)))
            self.after(0, lambda: self._append_log(f"Found {total_files} Excel files."))

            for index, file_path in enumerate(excel_files, start=1):
                self.after(0, lambda p=file_path: self._append_log(f"Processing: {p}"))
                result = process_workbook(file_path, logger=logger, debug_mode=debug_mode)
                all_rows.extend(result.rows)
                all_errors.extend(result.errors)
                all_debug_records.extend(result.debug_records)
                all_unmapped_columns.extend(result.unmapped_columns)
                workbook_seconds_total += result.processing_seconds

                if result.processed_successfully:
                    files_successful += 1
                else:
                    files_skipped += 1

                for error in result.errors:
                    logger.warning("%s | %s | %s | %s", error.file, error.sheet, error.reason, error.exception)

                self.after(
                    0,
                    lambda p=file_path, i=index, t=total_files, r=len(all_rows): self._update_progress(
                        current_file=p, index=i, total=t, start_time=start_time, rows=r
                    ),
                )

            export_to_excel(all_rows, output_file, unmapped_columns=all_unmapped_columns)
            error_report = export_error_report(all_errors, output_file)
            debug_report_text = "Debug report: not generated"
            if debug_mode:
                debug_report = export_debug_report(all_debug_records, output_file)
                debug_report_text = f"Debug report: {debug_report}"

            execution_time = time.perf_counter() - start_time
            files_per_sec = _safe_rate(total_files, execution_time)
            rows_per_sec = _safe_rate(len(all_rows), execution_time)
            avg_workbook_time = workbook_seconds_total / total_files if total_files else 0.0

            perf_line = (
                f"Performance | files/sec={files_per_sec:.2f} | rows/sec={rows_per_sec:.2f} | "
                f"avg_workbook_time={avg_workbook_time:.2f}s"
            )
            logger.info(perf_line)

            summary = (
                "Extraction finished.\n\n"
                f"Files scanned: {total_files}\n"
                f"Files processed successfully: {files_successful}\n"
                f"Files skipped: {files_skipped}\n"
                f"Rows exported: {len(all_rows)}\n"
                f"Execution time: {_format_duration(execution_time)}\n"
                f"Files/sec: {files_per_sec:.2f}\n"
                f"Rows/sec: {rows_per_sec:.2f}\n"
                f"Average workbook processing time: {avg_workbook_time:.2f}s\n"
                f"Error report: {error_report}\n"
                f"{debug_report_text}\n"
                f"Log file: {log_file}"
            )
            self.after(0, lambda: self._append_log(perf_line))
            self.after(0, lambda: self._append_log(summary))
            self.after(0, lambda: self.status_text.set("Extraction complete."))
            self.after(0, lambda: self.rows_exported_text.set(f"Rows exported: {len(all_rows)}"))
            self.after(0, lambda: messagebox.showinfo("Extraction complete", summary))

        except Exception as exc:
            self.after(0, lambda: self._append_log(f"FATAL ERROR: {exc}"))
            self.after(0, lambda: messagebox.showerror("Extraction failed", str(exc)))
        finally:
            self.after(0, lambda: self.run_button.configure(state="normal"))


def run_app() -> None:
    app = TcrExtractorApp()
    app.mainloop()
