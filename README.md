# The-Big-Beautiful-Budgeter

Reads statement CSV files from `statements/` and generates Excel reports.

## Main category report

- Uses keyword matching from `json/categories.json`
- Writes state to `json/processed_statements.json`
- Exports `statement_report.xlsx`

Run:

- `python3 main.py`
- `python3 main.py --rebuild` after changing `json/categories.json`

## Personal budget report (separate subsystem)

- Uses fixed monthly budgets from `json/personal_budget_categories.json`
- Tracks only UI category decisions in `json/personal_budget_assignments.json`
- Includes only purchases from June 2026 and later
- Uses Date+Time id dedupe when reading CSV purchases
- Exports `personal_budget_report.xlsx`
- Pops a Tkinter assignment window automatically when uncategorized June-2026+ purchases exist
- The popup supports single-category assignments or a split of 2–10 exact-dollar allocations across categories

You can run personal budget directly with:

- `python3 personal_budget.py`
- `python3 personal_budget.py --rebuild` to clear prior personal-budget decisions and re-categorize from scratch

Running `python3 main.py` now builds both reports.

You can trigger personal-budget rebuild from main with:

- `python3 main.py --personal-budget-rebuild`

## Requirements

- `openpyxl`
- `tkinter` (built into standard Python on macOS system Python)

## Startup

- `source .venv/bin/activate`
