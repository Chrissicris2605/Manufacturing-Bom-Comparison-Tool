"""Shared constants for the public BOM comparison demo."""

from dataclasses import dataclass


CODE_COLUMN = "code"
COMPARE_COLUMNS = ["description", "quantity", "unit", "revision", "notes"]
EXPECTED_COLUMNS = [CODE_COLUMN, *COMPARE_COLUMNS]


@dataclass(frozen=True)
class StatusStyle:
    label: str
    fill_color: str


STATUS_STYLES = {
    "Added": StatusStyle(label="Added", fill_color="C6EFCE"),      # light green
    "Removed": StatusStyle(label="Removed", fill_color="FFC7CE"),  # light red
    "Modified": StatusStyle(label="Modified", fill_color="FFEB9C"),# light yellow
    "Unchanged": StatusStyle(label="Unchanged", fill_color="BDD7EE"),# light blue
}
