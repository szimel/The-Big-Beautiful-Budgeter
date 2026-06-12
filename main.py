import csv
import json
import calendar
import argparse
from collections import defaultdict
from pathlib import Path
from datetime import UTC, datetime

from excel_exporter import export_sheet_data_to_excel


BASE_DIR = Path(__file__).resolve().parent
CATEGORIES_FILE = str(BASE_DIR / "json" / "categories.json")
OUTPUT_EXCEL_FILE = str(BASE_DIR / "statement_report.xlsx")
PROCESSING_STATE_FILE = str(BASE_DIR / "json" / "processed_statements.json")
STATEMENTS_FOLDER = str(BASE_DIR / "statements")


def parse_csv_date(date_value: str) -> datetime:
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(date_value.strip(), fmt)
        except ValueError:
            continue
    raise ValueError(f"Unsupported date format: {date_value}")


def format_short_date(dt: datetime) -> str:
    return f"{dt.month}/{dt.day}/{str(dt.year)[2:]}"


def format_short_time(time_value: str) -> str:
    return (time_value or "").strip().upper()


def normalize_cardholder(name: str) -> str:
    text = (name or "").strip()
    lower = text.casefold()
    if lower.startswith("samuel"):
        return "Samuel"
    if lower.startswith("kamrie"):
        return "Kamrie"
    if not text:
        return ""
    return text.split()[0].title()


def parse_points(value: str) -> int:
    raw = (value or "").replace(",", "").strip()
    if not raw:
        return 0
    return int(float(raw))


def parse_amount(value: str) -> float:
    raw = (value or "").replace(",", "").replace("$", "").strip()
    if not raw:
        return 0.0
    return float(raw)


def load_processing_state(state_path: str) -> dict:
    if not Path(state_path).exists():
        return {"processed_files": [], "sheets": {}}

    with open(state_path, "r", encoding="utf-8") as file:
        data = json.load(file)

    if "processed_files" not in data:
        data["processed_files"] = []

    if "sheets" not in data or not isinstance(data["sheets"], dict):
        data["sheets"] = {}

    return data


def save_processing_state(state_path: str, state: dict) -> None:
    payload = {
        "processed_files": state.get("processed_files", []),
        "sheets": state.get("sheets", {}),
    }
    with open(state_path, "w", encoding="utf-8") as file:
        json.dump(payload, file, indent=2)


def get_statement_files(statements_folder: str) -> list[str]:
    folder = Path(statements_folder)
    if not folder.exists():
        return []

    return sorted([path.name for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".csv"])


def load_categories(categories_path: str) -> dict[str, list[str]]:
    with open(categories_path, "r", encoding="utf-8") as file:
        data = json.load(file)

    if not isinstance(data, dict):
        raise ValueError("categories.json must contain an object of category names to keyword lists")

    for category, keywords in data.items():
        if not isinstance(category, str) or not isinstance(keywords, list) or not all(isinstance(k, str) for k in keywords):
            raise ValueError("Each category must be a string mapped to a list of strings")

    return data


def save_categories(categories_path: str, categories: dict[str, list[str]]) -> None:
    with open(categories_path, "w", encoding="utf-8") as file:
        json.dump(categories, file, indent=2)


def add_category_keyword(categories: dict[str, list[str]], category: str, keyword: str) -> bool:
    if category not in categories:
        categories[category] = []

    if keyword in categories[category]:
        return False

    categories[category].append(keyword)
    return True


def remove_category_keyword(categories: dict[str, list[str]], category: str, keyword: str) -> bool:
    if category not in categories:
        return False
    if keyword not in categories[category]:
        return False

    categories[category].remove(keyword)
    return True


def update_category_keyword(
    categories_path: str, category: str, keyword: str, action: str
) -> bool:
    categories = load_categories(categories_path)

    if action == "add":
        changed = add_category_keyword(categories, category, keyword)
    elif action == "remove":
        changed = remove_category_keyword(categories, category, keyword)
    else:
        raise ValueError("action must be 'add' or 'remove'")

    if changed:
        save_categories(categories_path, categories)

    return changed


def find_categories_for_description(description: str, categories: dict[str, list[str]]) -> list[str]:
    description_lower = (description or "").lower()
    matches = []

    for category, keywords in categories.items():
        for keyword in keywords:
            if (keyword or "").lower() in description_lower:
                matches.append(category)
                break

    return matches


def transaction_unique_id(transaction: dict) -> str | None:
    date_value = (transaction.get("Date") or "").strip()
    time_value = (transaction.get("Time") or "").strip().upper()
    if not date_value or not time_value:
        return None
    return f"{date_value}|{time_value}"


def collect_seen_transaction_ids(sheets: dict[str, list[dict]]) -> set[str]:
    seen_ids: set[str] = set()
    for sheet_name, sections in sheets.items():
        if sheet_name == "Summary":
            continue
        for section in sections:
            for transaction in section.get("transactions", []):
                unique_id = transaction_unique_id(transaction)
                if unique_id:
                    seen_ids.add(unique_id)
    return seen_ids


def dedupe_report_entries_by_seen_ids(report_entries: list[dict], seen_ids: set[str]) -> list[dict]:
    deduped_reports = []

    for report in report_entries:
        unique_transactions = []
        for transaction in report.get("transactions", []):
            unique_id = transaction_unique_id(transaction)
            if unique_id and unique_id in seen_ids:
                continue
            if unique_id:
                seen_ids.add(unique_id)
            unique_transactions.append(transaction)

        if unique_transactions:
            updated_report = dict(report)
            updated_report["transactions"] = unique_transactions
            deduped_reports.append(updated_report)

    return deduped_reports


def parse_csv_statement(file_name: str, categories: dict[str, list[str]]) -> list[dict]:
    file_path = Path(STATEMENTS_FOLDER) / file_name
    grouped: dict[tuple[int, int], list[dict]] = defaultdict(list)

    with open(file_path, "r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            if (row.get("Type") or "").strip().casefold() != "purchase":
                continue

            date_raw = (row.get("Date") or "").strip()
            if not date_raw:
                continue

            dt = parse_csv_date(date_raw)
            merchant = (row.get("Merchant") or "").strip()
            if not merchant:
                merchant = (row.get("Description") or "").strip()
            if not merchant:
                continue

            transaction = {
                "description": merchant,
                "amount": parse_amount(row.get("Amount") or ""),
                "categories": find_categories_for_description(merchant, categories),
                "Purchased By": normalize_cardholder(row.get("Cardholder") or ""),
                "Points": parse_points(row.get("Points") or ""),
                "Date": format_short_date(dt),
                "Time": format_short_time(row.get("Time") or ""),
            }
            grouped[(dt.year, dt.month)].append(transaction)

    processed_at = datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")
    reports = []
    for (year, month), transactions in grouped.items():
        reports.append(
            {
                "file_name": file_name,
                "year": year,
                "month": month,
                "processed_at": processed_at,
                "transactions": transactions,
            }
        )

    return sort_reports_descending(reports)


def sort_reports_descending(reports: list[dict]) -> list[dict]:
    return sorted(reports, key=lambda r: (r["year"], r["month"], r["file_name"]), reverse=True)


def month_year_label(year: int, month: int) -> str:
    return f"{calendar.month_name[month].upper()} {year}"


def summarize_transactions_for_category(transactions: list[dict], category: str) -> list[dict]:
    return [transaction for transaction in transactions if category in transaction.get("categories", [])]


def summarize_uncategorized_transactions(transactions: list[dict]) -> list[dict]:
    return [transaction for transaction in transactions if not transaction.get("categories")]


def build_summary_rows(transactions: list[dict], categories: dict[str, list[str]]) -> list[dict]:
    rows = []
    for category in categories.keys():
        category_transactions = summarize_transactions_for_category(transactions, category)
        rows.append(
            {
                "category": category,
                "transactions": len(category_transactions),
                "total_amount": sum(tx.get("amount", 0.0) for tx in category_transactions),
            }
        )

    uncategorized_transactions = summarize_uncategorized_transactions(transactions)
    rows.append(
        {
            "category": "Uncategorized",
            "transactions": len(uncategorized_transactions),
            "total_amount": sum(tx.get("amount", 0.0) for tx in uncategorized_transactions),
        }
    )
    return rows


def section_totals(transactions: list[dict]) -> dict:
    count = len(transactions)
    total = sum(transaction.get("amount", 0.0) for transaction in transactions)
    points = sum(int(transaction.get("Points", 0) or 0) for transaction in transactions)
    average = total / count if count else 0.0
    return {
        "transactions": count,
        "total_amount": total,
        "points": points,
        "average_amount": average,
    }


def build_sheet_data(reports: list[dict], categories: dict[str, list[str]]) -> dict[str, list[dict]]:
    sorted_reports = sort_reports_descending(reports)
    category_order = list(categories.keys())

    summary_years = []
    years = sorted({report["year"] for report in sorted_reports}, reverse=True)
    for year in years:
        year_reports = [report for report in sorted_reports if report["year"] == year]
        months = []
        for report in year_reports:
            transactions = report.get("transactions", [])
            rows = build_summary_rows(transactions, categories)

            samuel_transactions = [tx for tx in transactions if tx.get("Purchased By") == "Samuel"]
            kamrie_transactions = [tx for tx in transactions if tx.get("Purchased By") == "Kamrie"]

            months.append(
                {
                    "year": report["year"],
                    "month": report["month"],
                    "label": month_year_label(report["year"], report["month"]),
                    "totals": section_totals(transactions),
                    "rows": rows,
                    "by_person": {
                        "Samuel": {
                            "rows": build_summary_rows(samuel_transactions, categories),
                            "totals": section_totals(samuel_transactions),
                        },
                        "Kamrie": {
                            "rows": build_summary_rows(kamrie_transactions, categories),
                            "totals": section_totals(kamrie_transactions),
                        },
                    },
                }
            )

        summary_years.append({"year": year, "months": months})

    sheets: dict[str, list[dict]] = {"Summary": summary_years, "Uncategorized": []}
    for category in category_order:
        sheets[category] = []

    for report in sorted_reports:
        label = month_year_label(report["year"], report["month"])
        transactions = report.get("transactions", [])

        uncategorized_transactions = summarize_uncategorized_transactions(transactions)
        sheets["Uncategorized"].append(
            {
                "year": report["year"],
                "month": report["month"],
                "label": label,
                "totals": section_totals(uncategorized_transactions),
                "transactions": uncategorized_transactions,
            }
        )

        for category in category_order:
            category_transactions = summarize_transactions_for_category(transactions, category)
            sheets[category].append(
                {
                    "year": report["year"],
                    "month": report["month"],
                    "label": label,
                    "totals": section_totals(category_transactions),
                    "transactions": category_transactions,
                }
            )

    return sheets


def initialize_sheet_data(categories: dict[str, list[str]]) -> dict[str, list[dict]]:
    sheet_data: dict[str, list[dict]] = {"Summary": [], "Uncategorized": []}
    for category in categories.keys():
        sheet_data[category] = []
    return sheet_data


def upsert_month_section(sections: list[dict], new_section: dict) -> None:
    for index, section in enumerate(sections):
        if section.get("year") == new_section.get("year") and section.get("month") == new_section.get("month"):
            sections[index] = new_section
            return
    sections.append(new_section)


def append_report_to_sheet_data(sheets: dict[str, list[dict]], report: dict, categories: dict[str, list[str]]) -> None:
    year = report["year"]
    month = report["month"]
    label = month_year_label(year, month)
    transactions = report.get("transactions", [])

    summary_rows = build_summary_rows(transactions, categories)
    samuel_transactions = [tx for tx in transactions if tx.get("Purchased By") == "Samuel"]
    kamrie_transactions = [tx for tx in transactions if tx.get("Purchased By") == "Kamrie"]
    uncategorized_transactions = summarize_uncategorized_transactions(transactions)

    summary_block = next((item for item in sheets["Summary"] if item.get("year") == year), None)
    if summary_block is None:
        summary_block = {"year": year, "months": []}
        sheets["Summary"].append(summary_block)

    upsert_month_section(
        summary_block["months"],
        {
            "year": year,
            "month": month,
            "label": label,
            "totals": section_totals(transactions),
            "rows": summary_rows,
            "by_person": {
                "Samuel": {
                    "rows": build_summary_rows(samuel_transactions, categories),
                    "totals": section_totals(samuel_transactions),
                },
                "Kamrie": {
                    "rows": build_summary_rows(kamrie_transactions, categories),
                    "totals": section_totals(kamrie_transactions),
                },
            },
        },
    )

    upsert_month_section(
        sheets["Uncategorized"],
        {
            "year": year,
            "month": month,
            "label": label,
            "totals": section_totals(uncategorized_transactions),
            "transactions": uncategorized_transactions,
        },
    )

    for category in categories.keys():
        category_transactions = summarize_transactions_for_category(transactions, category)
        upsert_month_section(
            sheets[category],
            {
                "year": year,
                "month": month,
                "label": label,
                "totals": section_totals(category_transactions),
                "transactions": category_transactions,
            },
        )


def sort_sheet_data(sheets: dict[str, list[dict]]) -> None:
    summary_years = sheets.get("Summary", [])
    summary_years.sort(key=lambda y: y.get("year", 0), reverse=True)
    for year_block in summary_years:
        year_block.setdefault("months", []).sort(key=lambda m: m.get("month", 0), reverse=True)

    for sheet_name, sections in sheets.items():
        if sheet_name == "Summary":
            continue
        sections.sort(key=lambda s: (s.get("year", 0), s.get("month", 0)), reverse=True)


def normalize_sheet_transaction_schema(sheets: dict[str, list[dict]]) -> None:
    for sheet_name, sections in sheets.items():
        if sheet_name == "Summary":
            continue
        for section in sections:
            for transaction in section.get("transactions", []):
                transaction.setdefault("description", "")
                transaction.setdefault("amount", 0.0)
                transaction.setdefault("categories", [])
                transaction.setdefault("Purchased By", "")
                transaction.setdefault("Points", 0)
                transaction.setdefault("Date", "")
                transaction.setdefault("Time", "")


def rebuild_sheets_from_csv_files(
    csv_file_names: list[str], categories: dict[str, list[str]]
) -> tuple[dict[str, list[dict]], list[str]]:
    sheets = initialize_sheet_data(categories)

    for file_name in sorted(csv_file_names):
        report_entries = parse_csv_statement(file_name, categories)
        for report_entry in report_entries:
            append_report_to_sheet_data(sheets, report_entry, categories)

    normalize_sheet_transaction_schema(sheets)
    normalize_summary_schema(sheets, categories)
    sort_sheet_data(sheets)
    return sheets, sorted(set(csv_file_names))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Budget statement processor")
    parser.add_argument(
        "--rebuild",
        action="store_true",
        help="Force reprocess all CSV files using current categories and rebuild all sheets/state from scratch.",
    )
    return parser.parse_args()


def normalize_summary_schema(sheets: dict[str, list[dict]], categories: dict[str, list[str]]) -> None:
    summary = sheets.get("Summary", [])
    default_categories = list(categories.keys()) + ["Uncategorized"]

    def default_rows() -> list[dict]:
        return [{"category": c, "transactions": 0, "total_amount": 0.0} for c in default_categories]

    for year_block in summary:
        for month_block in year_block.get("months", []):
            rows = month_block.get("rows")
            if not isinstance(rows, list) or not rows:
                month_block["rows"] = default_rows()

            month_block.setdefault("totals", {})
            month_block["totals"].setdefault("transactions", 0)
            month_block["totals"].setdefault("total_amount", 0.0)
            month_block["totals"].setdefault("points", 0)
            month_block["totals"].setdefault("average_amount", 0.0)

            by_person = month_block.setdefault("by_person", {})
            for person in ("Samuel", "Kamrie"):
                person_block = by_person.setdefault(person, {})
                person_rows = person_block.get("rows")
                if not isinstance(person_rows, list) or not person_rows:
                    person_block["rows"] = [{"category": r.get("category", ""), "transactions": 0, "total_amount": 0.0} for r in month_block["rows"]]

                person_totals = person_block.setdefault("totals", {})
                person_totals.setdefault("transactions", 0)
                person_totals.setdefault("total_amount", 0.0)
                person_totals.setdefault("points", 0)
                person_totals.setdefault("average_amount", 0.0)


if __name__ == "__main__":
    args = parse_args()
    categories = load_categories(CATEGORIES_FILE)
    state = load_processing_state(PROCESSING_STATE_FILE)
    statement_files = get_statement_files(STATEMENTS_FOLDER)

    if args.rebuild:
        sheets, processed_files = rebuild_sheets_from_csv_files(statement_files, categories)
        state["sheets"] = sheets
        state["processed_files"] = processed_files
        new_statement_files = statement_files
        run_mode = "full-rebuild"
    else:
        sheets = state.get("sheets") or initialize_sheet_data(categories)
        for category_name in categories.keys():
            sheets.setdefault(category_name, [])
        sheets.setdefault("Summary", [])
        sheets.setdefault("Uncategorized", [])

        legacy_reports = state.get("reports", [])
        if legacy_reports and not sheets.get("Summary"):
            sheets = build_sheet_data(legacy_reports, categories)

        processed_file_names = set(state.get("processed_files", []))
        new_statement_files = [name for name in statement_files if name not in processed_file_names]
        seen_transaction_ids = collect_seen_transaction_ids(sheets)

        for file_name in new_statement_files:
            report_entries = parse_csv_statement(file_name, categories)
            report_entries = dedupe_report_entries_by_seen_ids(report_entries, seen_transaction_ids)
            for report_entry in report_entries:
                append_report_to_sheet_data(sheets, report_entry, categories)
            state.setdefault("processed_files", []).append(file_name)

        normalize_sheet_transaction_schema(sheets)
        normalize_summary_schema(sheets, categories)
        sort_sheet_data(sheets)
        state["processed_files"] = sorted(set(state.get("processed_files", [])))
        state["sheets"] = sheets
        run_mode = "incremental"

    # Remove legacy bulk report storage from state payload to keep file size down.
    if "reports" in state:
        del state["reports"]

    save_processing_state(PROCESSING_STATE_FILE, state)

    export_sheet_data_to_excel(state["sheets"], list(categories.keys()), OUTPUT_EXCEL_FILE)
    print(
        "Export complete: "
        f"mode={run_mode} | "
        f"{OUTPUT_EXCEL_FILE} | New files processed: {len(new_statement_files)} | "
        f"Total statements tracked: {len(state['processed_files'])}"
    )