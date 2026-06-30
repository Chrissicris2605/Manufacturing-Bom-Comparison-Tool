"""Command-line entry point for the public BOM comparison demo."""

from __future__ import annotations

import argparse
from pathlib import Path

from bom_loader import load_bom_csv
from comparison_engine import compare_boms
from report_generator import generate_excel_report


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a formatted Excel delta report from two generic CSV files.")
    parser.add_argument("old_file", help="Path to the baseline CSV file")
    parser.add_argument("new_file", help="Path to the revised CSV file")
    parser.add_argument("output_file", help="Path to the output Excel report")
    parser.add_argument("--changes-only", action="store_true", help="Exclude unchanged rows from the detailed report")
    args = parser.parse_args()

    baseline_df = load_bom_csv(args.old_file)
    revised_df = load_bom_csv(args.new_file)
    delta_df = compare_boms(baseline_df, revised_df, include_unchanged=not args.changes_only)
    output_path = generate_excel_report(delta_df, Path(args.output_file))
    print(f"Delta report generated: {output_path}")


if __name__ == "__main__":
    main()
