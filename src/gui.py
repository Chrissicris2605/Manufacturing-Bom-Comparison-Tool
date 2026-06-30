"""Simple desktop interface for the public BOM comparison demo."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from bom_loader import load_bom_csv
from comparison_engine import compare_boms
from report_generator import generate_excel_report


class BomComparisonApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Engineering BOM Comparison Tool - Public Demo")
        self.geometry("760x300")
        self.resizable(False, False)

        self.old_file_var = tk.StringVar()
        self.new_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar(value=str(Path("output_examples/delta_report.xlsx")))
        self.changes_only_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=20)
        container.pack(fill="both", expand=True)

        title = ttk.Label(
            container,
            text="Engineering BOM Comparison Tool",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        subtitle = ttk.Label(
            container,
            text="Public demo: compare two generic CSV files by code and generate a formatted Excel report.",
        )
        subtitle.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 20))

        self._add_file_row(container, 2, "Baseline file", self.old_file_var, self._select_old_file)
        self._add_file_row(container, 3, "Revised file", self.new_file_var, self._select_new_file)
        self._add_file_row(container, 4, "Output report", self.output_file_var, self._select_output_file, save=True)

        changes_only = ttk.Checkbutton(
            container,
            text="Include only Added, Removed and Modified rows",
            variable=self.changes_only_var,
        )
        changes_only.grid(row=5, column=1, sticky="w", pady=(8, 0))

        run_button = ttk.Button(container, text="Generate Delta Report", command=self._run_comparison)
        run_button.grid(row=6, column=1, sticky="e", pady=(24, 0))

        container.columnconfigure(1, weight=1)

    def _add_file_row(self, parent, row: int, label: str, variable: tk.StringVar, command, save: bool = False) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        entry = ttk.Entry(parent, textvariable=variable, width=72)
        entry.grid(row=row, column=1, sticky="ew", padx=(12, 8), pady=6)
        button_text = "Save as" if save else "Browse"
        ttk.Button(parent, text=button_text, command=command).grid(row=row, column=2, pady=6)

    def _select_old_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.old_file_var.set(path)

    def _select_new_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.new_file_var.set(path)

    def _select_output_file(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="delta_report.xlsx",
        )
        if path:
            self.output_file_var.set(path)

    def _run_comparison(self) -> None:
        try:
            baseline_df = load_bom_csv(self.old_file_var.get())
            revised_df = load_bom_csv(self.new_file_var.get())
            delta_df = compare_boms(
                baseline_df,
                revised_df,
                include_unchanged=not self.changes_only_var.get(),
            )
            output_path = generate_excel_report(delta_df, self.output_file_var.get())
            messagebox.showinfo("Success", f"Report generated successfully:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))


if __name__ == "__main__":
    app = BomComparisonApp()
    app.mainloop()
