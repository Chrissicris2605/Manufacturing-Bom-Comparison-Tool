# Engineering BOM Comparison Tool

**Public demo – Engineering Automation / Python / Excel Reporting**

This repository contains a **safe public demo** of an engineering automation tool designed to compare structured BOM-like spreadsheets and generate a technical delta report.

The original professional work that inspired this project was developed in an industrial engineering context. This public repository is **not** the production tool. It is a simplified, rebuilt-from-scratch implementation using fictional data, generic field names, and non-proprietary business rules.

---

## Why this project exists

Engineering teams often need to compare structured technical lists across revisions. When this is done manually, the work can become slow, repetitive, and error-prone.

This demo shows how that workflow can be transformed into a small automation system that:

- loads two structured files;
- compares items by `code`;
- classifies each item as `Added`, `Removed`, `Modified`, or `Unchanged`;
- generates a formatted Excel report;
- uses colors to make the output easier to review.

---

## Status classification

| Status | Meaning | Report color |
|---|---|---|
| Added | Code exists only in the new file | Green |
| Removed | Code exists only in the old file | Red |
| Modified | Code exists in both files, but one or more fields changed | Yellow |
| Unchanged | Code exists in both files and all compared fields are equal | Blue |

---

## Demo scope

This public version compares generic CSV files with the following columns:

- `code`
- `description`
- `quantity`
- `unit`
- `revision`
- `notes`

The comparison is intentionally generic and can be adapted to many structured engineering workflows.

---

## Repository structure

```text
.
├── src/
│   ├── main.py                 # CLI entry point
│   ├── gui.py                  # Simple desktop interface
│   ├── bom_loader.py           # CSV loading and validation
│   ├── comparison_engine.py    # Comparison rules
│   ├── report_generator.py     # Excel report generation
│   └── models.py               # Shared constants and data structures
├── sample_data/
│   ├── old_bom.csv
│   └── new_bom.csv
├── output_examples/
│   └── README.md
├── docs/
│   ├── confidentiality.md
│   └── public-demo-scope.md
├── requirements.txt
└── README.md
```

---

## How to run

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Run with the command line

```bash
python src/main.py sample_data/old_bom.csv sample_data/new_bom.csv output_examples/delta_report.xlsx
```

To generate only changed rows:

```bash
python src/main.py sample_data/old_bom.csv sample_data/new_bom.csv output_examples/delta_report.xlsx --changes-only
```

### 3. Run with the desktop interface

```bash
python src/gui.py
```

The interface allows selecting the baseline file, the revised file, and the output path before generating the report.

---

## Example output

The generated Excel workbook contains:

- a `Summary` sheet with totals by status;
- a `Delta Report` sheet with row-level comparison results;
- color-coded status cells for quick visual review.

---

## Professional relevance

This project demonstrates skills in:

- engineering automation;
- structured data comparison;
- Python development;
- Excel report generation;
- process improvement;
- desktop tooling;
- workflow-oriented software design.

It reflects the kind of work I enjoy building: practical software that reduces repetitive engineering effort and improves consistency in technical workflows.

---

## Confidentiality notice

This repository does **not** include:

- production source code;
- proprietary algorithms;
- real datasets;
- internal company files;
- client information;
- confidential business rules;
- private naming conventions.

All sample data is fictional. All logic was rebuilt from scratch for public demonstration purposes.
