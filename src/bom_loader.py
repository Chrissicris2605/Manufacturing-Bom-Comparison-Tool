"""Data loading utilities for the public BOM comparison demo."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from models import CODE_COLUMN, EXPECTED_COLUMNS


class BomValidationError(ValueError):
    """Raised when an input file does not match the expected demo schema."""


def load_bom_csv(path: str | Path) -> pd.DataFrame:
    """Load and validate a generic BOM-like CSV file.

    The public demo intentionally uses generic column names and fictional data.
    Items are identified by the `code` column.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    df = pd.read_csv(path, dtype=str).fillna("")
    df.columns = [column.strip().lower() for column in df.columns]

    missing_columns = [column for column in EXPECTED_COLUMNS if column not in df.columns]
    if missing_columns:
        raise BomValidationError(
            "Missing required column(s): " + ", ".join(missing_columns)
        )

    df = df[EXPECTED_COLUMNS].copy()
    df[CODE_COLUMN] = df[CODE_COLUMN].astype(str).str.strip()

    if df[CODE_COLUMN].eq("").any():
        raise BomValidationError("Every row must contain a non-empty code.")

    duplicated_codes = df[df[CODE_COLUMN].duplicated()][CODE_COLUMN].tolist()
    if duplicated_codes:
        raise BomValidationError(
            "Duplicated code(s) found: " + ", ".join(sorted(set(duplicated_codes)))
        )

    for column in EXPECTED_COLUMNS:
        df[column] = df[column].astype(str).str.strip()

    return df
