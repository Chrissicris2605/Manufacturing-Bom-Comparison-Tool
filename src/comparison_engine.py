"""Comparison logic for the public BOM comparison demo."""

from __future__ import annotations

import pandas as pd

from models import CODE_COLUMN, COMPARE_COLUMNS


def _build_change_summary(old_row: pd.Series, new_row: pd.Series) -> str:
    changes: list[str] = []
    for column in COMPARE_COLUMNS:
        old_value = str(old_row[column])
        new_value = str(new_row[column])
        if old_value != new_value:
            changes.append(f"{column}: '{old_value}' -> '{new_value}'")
    return "; ".join(changes)


def compare_boms(old_df: pd.DataFrame, new_df: pd.DataFrame, include_unchanged: bool = True) -> pd.DataFrame:
    """Compare two generic BOM-like dataframes by code.

    Status rules:
    - Added: code exists only in the new file
    - Removed: code exists only in the old file
    - Modified: code exists in both files and at least one compared field changed
    - Unchanged: code exists in both files and all compared fields are equal
    """
    old_by_code = old_df.set_index(CODE_COLUMN)
    new_by_code = new_df.set_index(CODE_COLUMN)

    all_codes = sorted(set(old_by_code.index) | set(new_by_code.index))
    rows: list[dict[str, str]] = []

    for code in all_codes:
        in_old = code in old_by_code.index
        in_new = code in new_by_code.index

        if in_old and not in_new:
            old_row = old_by_code.loc[code]
            rows.append({
                "status": "Removed",
                "code": code,
                "old_description": old_row["description"],
                "new_description": "",
                "old_quantity": old_row["quantity"],
                "new_quantity": "",
                "old_unit": old_row["unit"],
                "new_unit": "",
                "old_revision": old_row["revision"],
                "new_revision": "",
                "changes": "Code exists only in old file",
            })
            continue

        if in_new and not in_old:
            new_row = new_by_code.loc[code]
            rows.append({
                "status": "Added",
                "code": code,
                "old_description": "",
                "new_description": new_row["description"],
                "old_quantity": "",
                "new_quantity": new_row["quantity"],
                "old_unit": "",
                "new_unit": new_row["unit"],
                "old_revision": "",
                "new_revision": new_row["revision"],
                "changes": "Code exists only in new file",
            })
            continue

        old_row = old_by_code.loc[code]
        new_row = new_by_code.loc[code]
        change_summary = _build_change_summary(old_row, new_row)
        status = "Modified" if change_summary else "Unchanged"

        if status == "Unchanged" and not include_unchanged:
            continue

        rows.append({
            "status": status,
            "code": code,
            "old_description": old_row["description"],
            "new_description": new_row["description"],
            "old_quantity": old_row["quantity"],
            "new_quantity": new_row["quantity"],
            "old_unit": old_row["unit"],
            "new_unit": new_row["unit"],
            "old_revision": old_row["revision"],
            "new_revision": new_row["revision"],
            "changes": change_summary or "No changes detected",
        })

    return pd.DataFrame(rows)
