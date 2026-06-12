import re
import json
import calendar
from pathlib import Path
from datetime import UTC, datetime

import pymupdf
from excel_exporter import export_sheet_data_to_excel


SECTION_HEADER = "TRANSACTION DESCRIPTION"
AMOUNT_HEADER = "AMOUNT"
STOP_MARKER = "TOTAL FEES FOR THIS PERIOD"
CATEGORIES_FILE = "./json/categories.json"
OUTPUT_EXCEL_FILE = "./statement_report.xlsx"
PROCESSING_STATE_FILE = "./json/processed_statements.json"
STATEMENTS_FOLDER = "./statements"
DATE_RE = re.compile(r"^\d{2}/\d{2}$")
AMOUNT_RE = re.compile(
    r"^\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})(?:-)?$|^-?\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})$"
)
STATEMENT_PERIOD_RE = re.compile(
    r"(?P<month>January|February|March|April|May|June|July|August|September|October|November|December|"
    r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+(?P<year>\d{4})",
    re.IGNORECASE,
)
MONTH_NAME_TO_NUMBER = {
    "january": 1,
    "jan": 1,
    "february": 2,
    "feb": 2,
    "march": 3,
    "mar": 3,
    "april": 4,
    "apr": 4,
    "may": 5,
    "june": 6,
    "jun": 6,
    "july": 7,
    "jul": 7,
    "august": 8,
    "aug": 8,
    "september": 9,
    "sep": 9,
    "sept": 9,
    "october": 10,
    "oct": 10,
    "november": 11,
    "nov": 11,
    "december": 12,
    "dec": 12,
}


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

    return sorted([path.name for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"])


def parse_statement_period_from_filename(file_name: str) -> tuple[int, int]:
    match = STATEMENT_PERIOD_RE.search(file_name)
    if not match:
        raise ValueError(f"Could not infer statement month/year from file name: {file_name}")

    month_raw = match.group("month").lower()
    year = int(match.group("year"))
    month = MONTH_NAME_TO_NUMBER[month_raw]
    return year, month


def extract_pdf_text(pdf_path: str) -> str:
    doc = pymupdf.open(pdf_path)
    text_parts = []
    for page in doc:
        text_parts.append(page.get_text())
    return "\n".join(text_parts).upper()


def normalize_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def is_date(line: str) -> bool:
    return bool(DATE_RE.match(line))


def is_amount(line: str) -> bool:
    return bool(AMOUNT_RE.match(line))


def is_reference_id(line: str) -> bool:
    # Statement IDs are usually long alphanumeric tokens with no spaces.
    return " " not in line and len(line) >= 10 and any(ch.isalpha() for ch in line) and any(ch.isdigit() for ch in line)


def parse_amount(line: str) -> float:
    cleaned = line.replace("$", "").replace(",", "")
    sign = -1.0 if cleaned.endswith("-") else 1.0
    if cleaned.endswith("-"):
        cleaned = cleaned[:-1]
    return sign * float(cleaned)


def is_restart_marker(line: str) -> bool:
    upper = line.upper()
    return upper.startswith("PAYMENT") or " CREDIT " in f" {upper} "


def find_data_section_starts(lines: list[str]) -> list[int]:
    starts = []
    for i in range(len(lines) - 1):
        if lines[i] == SECTION_HEADER and lines[i + 1] == AMOUNT_HEADER:
            starts.append(i + 2)
    return starts


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
    description_folded = description.casefold()
    matches = []

    for category, keywords in categories.items():
        for keyword in keywords:
            if keyword.casefold() in description_folded:
                matches.append(category)
                break

    return matches


def add_categories_to_transactions(transactions: list[dict], categories: dict[str, list[str]]) -> list[dict]:
    for transaction in transactions:
        transaction["categories"] = find_categories_for_description(transaction["description"], categories)
    return transactions


def extract_transactions(text: str) -> list[dict]:
    lines = normalize_lines(text)
    if STOP_MARKER in lines:
        lines = lines[: lines.index(STOP_MARKER)]
    starts = find_data_section_starts(lines)
    transactions = []

    for section_index, start in enumerate(starts):
        end = starts[section_index + 1] - 2 if section_index + 1 < len(starts) else len(lines)
        i = start

        while i < end:
            if i + 1 < end and is_date(lines[i]) and is_date(lines[i + 1]):
                j = i + 2

                while j < end and is_reference_id(lines[j]):
                    j += 1

                description_parts = []
                while j < end and not is_amount(lines[j]):
                    if lines[j] == SECTION_HEADER:
                        break
                    if is_restart_marker(lines[j]):
                        break
                    description_parts.append(lines[j])
                    j += 1

                if j < end and description_parts and is_amount(lines[j]):
                    transactions.append(
                        {
                            "description": " ".join(description_parts).upper(),
                            "amount": parse_amount(lines[j]),
                        }
                    )
                    i = j + 1
                    continue

            i += 1

    return transactions


def process_statement_file(file_name: str, categories: dict[str, list[str]]) -> dict:
    file_path = str(Path(STATEMENTS_FOLDER) / file_name)
    year, month = parse_statement_period_from_filename(file_name)

    text = extract_pdf_text(file_path)
    parsed_transactions = extract_transactions(text)
    categorized_transactions = add_categories_to_transactions(parsed_transactions, categories)

    return {
        "file_name": file_name,
        "year": year,
        "month": month,
        "processed_at": datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z"),
        "transactions": categorized_transactions,
    }


def normalize_report_descriptions_to_uppercase(reports: list[dict]) -> None:
    for report in reports:
        for transaction in report.get("transactions", []):
            description = transaction.get("description")
            if isinstance(description, str):
                transaction["description"] = description.upper()


def normalize_sheet_descriptions_to_uppercase(sheets: dict[str, list[dict]]) -> None:
    for sheet_name, sections in sheets.items():
        if sheet_name == "Summary":
            continue

        for section in sections:
            for transaction in section.get("transactions", []):
                description = transaction.get("description")
                if isinstance(description, str):
                    transaction["description"] = description.upper()


def sort_reports_descending(reports: list[dict]) -> list[dict]:
    return sorted(reports, key=lambda r: (r["year"], r["month"], r["file_name"]), reverse=True)


def month_year_label(year: int, month: int) -> str:
    return f"{calendar.month_name[month].upper()} {year}"


def summarize_transactions_for_category(transactions: list[dict], category: str) -> list[dict]:
    return [transaction for transaction in transactions if category in transaction.get("categories", [])]


def summarize_uncategorized_transactions(transactions: list[dict]) -> list[dict]:
    return [transaction for transaction in transactions if not transaction.get("categories")]


def section_totals(transactions: list[dict]) -> dict:
    count = len(transactions)
    total = sum(transaction.get("amount", 0.0) for transaction in transactions)
    average = total / count if count else 0.0
    return {
        "transactions": count,
        "total_amount": total,
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
            rows = []
            for category in category_order:
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

            months.append(
                {
                    "year": report["year"],
                    "month": report["month"],
                    "label": month_year_label(report["year"], report["month"]),
                    "rows": rows,
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

    summary_rows = []
    for category in categories.keys():
        category_transactions = summarize_transactions_for_category(transactions, category)
        summary_rows.append(
            {
                "category": category,
                "transactions": len(category_transactions),
                "total_amount": sum(tx.get("amount", 0.0) for tx in category_transactions),
            }
        )

    uncategorized_transactions = summarize_uncategorized_transactions(transactions)
    summary_rows.append(
        {
            "category": "Uncategorized",
            "transactions": len(uncategorized_transactions),
            "total_amount": sum(tx.get("amount", 0.0) for tx in uncategorized_transactions),
        }
    )

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
            "rows": summary_rows,
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


if __name__ == "__main__":
    categories = load_categories(CATEGORIES_FILE)
    state = load_processing_state(PROCESSING_STATE_FILE)
    statement_files = get_statement_files(STATEMENTS_FOLDER)

    sheets = state.get("sheets") or initialize_sheet_data(categories)
    for category_name in categories.keys():
        sheets.setdefault(category_name, [])
    sheets.setdefault("Summary", [])
    sheets.setdefault("Uncategorized", [])

    legacy_reports = state.get("reports", [])
    if legacy_reports and not sheets.get("Summary"):
        normalize_report_descriptions_to_uppercase(legacy_reports)
        sheets = build_sheet_data(legacy_reports, categories)

    processed_file_names = set(state.get("processed_files", []))
    new_statement_files = [name for name in statement_files if name not in processed_file_names]

    for file_name in new_statement_files:
        report_entry = process_statement_file(file_name, categories)
        append_report_to_sheet_data(sheets, report_entry, categories)
        state.setdefault("processed_files", []).append(file_name)

    normalize_sheet_descriptions_to_uppercase(sheets)
    sort_sheet_data(sheets)
    state["processed_files"] = sorted(set(state.get("processed_files", [])))
    state["sheets"] = sheets

    # Remove legacy bulk report storage from state payload to keep file size down.
    if "reports" in state:
        del state["reports"]

    save_processing_state(PROCESSING_STATE_FILE, state)

    export_sheet_data_to_excel(state["sheets"], list(categories.keys()), OUTPUT_EXCEL_FILE)
    print(
        "Export complete: "
        f"{OUTPUT_EXCEL_FILE} | New files processed: {len(new_statement_files)} | "
        f"Total statements tracked: {len(state['processed_files'])}"
    )