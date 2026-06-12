from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from collections import Counter
from collections import defaultdict
import re


SHEET_FORBIDDEN_CHARS = set('[]:*?/\\')
SECTION_COLORS = ["1F4E78", "2E75B6", "4F81BD", "5B9BD5"]
YEAR_PALETTES = [
    {"banner": "1F4E78", "month": "DCE6F1", "card": "FDE9D9"},
    {"banner": "7F6000", "month": "FCE4D6", "card": "FFF2CC"},
    {"banner": "375623", "month": "E2F0D9", "card": "E2EFDA"},
    {"banner": "5B2C6F", "month": "EAD1F2", "card": "EDE2F7"},
    {"banner": "0B6E4F", "month": "D1F2EB", "card": "D6F5E3"},
    {"banner": "9C6500", "month": "FFE9CC", "card": "FFEFD5"},
]


def safe_sheet_title(name: str) -> str:
    cleaned = "".join("_" if char in SHEET_FORBIDDEN_CHARS else char for char in name).strip()
    return cleaned[:31] if cleaned else "Sheet"


def apply_header_style(cell) -> None:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center")


def auto_fit_columns(worksheet, min_width: int = 12, max_width: int = 70) -> None:
    for col_cells in worksheet.columns:
        values = ["" if cell.value is None else str(cell.value) for cell in col_cells]
        width = max(len(value) for value in values) + 2
        worksheet.column_dimensions[col_cells[0].column_letter].width = max(min_width, min(width, max_width))


def set_currency(cell) -> None:
    cell.number_format = "$#,##0.00"


def get_year_palette(year: int, ordered_years: list[int]) -> dict[str, str]:
    index = ordered_years.index(year) if year in ordered_years else 0
    return YEAR_PALETTES[index % len(YEAR_PALETTES)]


def write_summary_sheet(worksheet, summary_data: list[dict]) -> None:
    worksheet.title = "Summary"
    worksheet["A1"] = "MULTI-YEAR SUMMARY"
    worksheet["A1"].font = Font(size=18, bold=True)
    worksheet["A2"] = "YEARS STACK VERTICALLY. MONTHS SPREAD HORIZONTALLY (NEWEST TO OLDEST)."
    worksheet["A2"].font = Font(color="666666", italic=True)

    row = 4
    thin = Side(style="thin", color="D9D9D9")

    ordered_years = [item.get("year", 0) for item in summary_data]

    for year_index, year_block in enumerate(summary_data):
        year = year_block.get("year", 0)
        palette = get_year_palette(year, ordered_years)
        months = year_block.get("months", [])
        month_block_width = 4
        month_block_data_rows = max((len(month.get("rows", [])) for month in months), default=1)
        section_height = 3 + month_block_data_rows
        last_col = max(1, len(months) * month_block_width)

        year_total_transactions = 0
        year_total_amount = 0.0
        for month in months:
            for item in month.get("rows", []):
                if item.get("category") == "Uncategorized":
                    continue
                year_total_transactions += int(item.get("transactions", 0))
                year_total_amount += float(item.get("total_amount", 0.0))

        worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        year_cell = worksheet.cell(row=row, column=1)
        year_cell.value = f"YEAR {year}"
        year_cell.font = Font(size=14, bold=True, color="FFFFFF")
        year_cell.fill = PatternFill("solid", fgColor=palette["banner"])
        year_cell.alignment = Alignment(horizontal="left")

        totals_start_col = last_col + 2
        worksheet.merge_cells(start_row=row, start_column=totals_start_col, end_row=row, end_column=totals_start_col + 2)
        totals_banner = worksheet.cell(row=row, column=totals_start_col)
        totals_banner.value = "YEAR TOTALS"
        totals_banner.font = Font(size=12, bold=True, color="FFFFFF")
        totals_banner.fill = PatternFill("solid", fgColor=palette["banner"])
        totals_banner.alignment = Alignment(horizontal="center")

        totals_rows = [
            ("Months", len(months), False),
            ("Transactions", year_total_transactions, False),
            ("Total Amount", year_total_amount, True),
        ]
        for offset, (label, value, is_currency) in enumerate(totals_rows, start=1):
            label_cell = worksheet.cell(row=row + offset, column=totals_start_col, value=label)
            label_cell.font = Font(bold=True, color="1F4E78")
            label_cell.fill = PatternFill("solid", fgColor=palette["card"])
            value_cell = worksheet.cell(row=row + offset, column=totals_start_col + 1, value=value)
            value_cell.fill = PatternFill("solid", fgColor=palette["card"])
            if is_currency:
                set_currency(value_cell)

        for month_index, month_block in enumerate(months):
            start_col = 1 + month_index * month_block_width
            end_col = start_col + 2

            worksheet.merge_cells(start_row=row + 1, start_column=start_col, end_row=row + 1, end_column=end_col)
            month_cell = worksheet.cell(row=row + 1, column=start_col)
            month_cell.value = month_block.get("label", "UNKNOWN MONTH")
            month_cell.font = Font(bold=True, color="1F4E78")
            month_cell.fill = PatternFill("solid", fgColor=palette["month"])
            month_cell.alignment = Alignment(horizontal="center")

            headers = ["Category", "Transactions", "Total Amount"]
            for offset, header in enumerate(headers):
                header_cell = worksheet.cell(row=row + 2, column=start_col + offset)
                header_cell.value = header
                apply_header_style(header_cell)

            data_rows = month_block.get("rows", [])
            for data_index, item in enumerate(data_rows):
                data_row = row + 3 + data_index
                worksheet.cell(row=data_row, column=start_col, value=item.get("category", ""))
                worksheet.cell(row=data_row, column=start_col + 1, value=item.get("transactions", 0))
                amount_cell = worksheet.cell(row=data_row, column=start_col + 2, value=item.get("total_amount", 0.0))
                set_currency(amount_cell)

            for block_row in range(row + 2, row + 3 + max(1, month_block_data_rows)):
                for block_col in range(start_col, end_col + 1):
                    worksheet.cell(row=block_row, column=block_col).border = Border(
                        left=thin, right=thin, top=thin, bottom=thin
                    )

        row += section_height + 2

    auto_fit_columns(worksheet)


def simplify_company_name(description: str) -> str:
    name = description.upper().strip()
    name = re.sub(r"\s+[A-Z]{2}$", "", name)
    name = re.sub(r"#\d+", "", name)
    name = re.sub(r"\b\d{3,}\b", "", name)
    name = re.sub(r"\s+", " ", name).strip(" -*")
    return name or description.upper().strip()


def write_uncategorized_top_companies(worksheet, sections: list[dict]) -> None:
    counter = Counter()
    amount_by_name: dict[str, float] = defaultdict(float)

    for section in sections:
        for tx in section.get("transactions", []):
            name = simplify_company_name(tx.get("description", ""))
            if not name:
                continue
            counter[name] += 1
            amount_by_name[name] += float(tx.get("amount", 0.0))

    top_items = counter.most_common(10)
    start_col = 6

    worksheet.merge_cells(start_row=4, start_column=start_col, end_row=4, end_column=start_col + 2)
    title = worksheet.cell(row=4, column=start_col)
    title.value = "TOP 10 UNCATEGORIZED NAMES"
    title.font = Font(size=12, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="1F4E78")
    title.alignment = Alignment(horizontal="center")

    headers = ["Name", "Count", "Total Amount"]
    for offset, label in enumerate(headers):
        header_cell = worksheet.cell(row=5, column=start_col + offset, value=label)
        apply_header_style(header_cell)

    row = 6
    for name, count in top_items:
        worksheet.cell(row=row, column=start_col, value=name)
        worksheet.cell(row=row, column=start_col + 1, value=count)
        amount_cell = worksheet.cell(row=row, column=start_col + 2, value=float(amount_by_name[name]))
        set_currency(amount_cell)
        row += 1


def write_ledger_sheet(worksheet, sheet_name: str, sections: list[dict]) -> None:
    worksheet.title = safe_sheet_title(sheet_name)
    worksheet["A1"] = f"{sheet_name.upper()} LEDGER"
    worksheet["A1"].font = Font(size=18, bold=True)
    worksheet["A2"] = "MONTHLY SECTIONS STACK NEWEST TO OLDEST"
    worksheet["A2"].font = Font(color="666666", italic=True)

    row = 4

    ordered_years = sorted({section.get("year", 0) for section in sections}, reverse=True)

    for index, section in enumerate(sections):
        palette = get_year_palette(section.get("year", 0), ordered_years)
        header_color = palette["banner"]
        label = section.get("label", "UNKNOWN MONTH")
        totals = section.get("totals", {})
        section_transactions = section.get("transactions", [])

        worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        header_cell = worksheet.cell(row=row, column=1)
        header_cell.value = label
        header_cell.font = Font(size=13, bold=True, color="FFFFFF")
        header_cell.fill = PatternFill("solid", fgColor=header_color)
        header_cell.alignment = Alignment(horizontal="left")

        worksheet.cell(row=row + 1, column=1, value="Transactions")
        worksheet.cell(row=row + 1, column=2, value=totals.get("transactions", 0))
        worksheet.cell(row=row + 1, column=3, value="Total Amount")
        amount_cell = worksheet.cell(row=row + 1, column=4, value=totals.get("total_amount", 0.0))
        set_currency(amount_cell)

        for col in (1, 3):
            metric_cell = worksheet.cell(row=row + 1, column=col)
            metric_cell.font = Font(bold=True, color="1F4E78")

        headers = ["Description", "Amount", "All Matched Categories"]
        for col, text in enumerate(headers, start=1):
            header = worksheet.cell(row=row + 3, column=col, value=text)
            apply_header_style(header)

        data_row = row + 4
        if not section_transactions:
            worksheet.cell(row=data_row, column=1, value="(No transactions)")
            worksheet.cell(row=data_row, column=2, value=0.0)
            set_currency(worksheet.cell(row=data_row, column=2))
            data_row += 1
        else:
            for tx in section_transactions:
                worksheet.cell(row=data_row, column=1, value=tx.get("description", ""))
                amount = worksheet.cell(row=data_row, column=2, value=tx.get("amount", 0.0))
                set_currency(amount)
                worksheet.cell(row=data_row, column=3, value=", ".join(tx.get("categories", [])))
                data_row += 1

        row = data_row + 2

    auto_fit_columns(worksheet)

    if sheet_name.lower() == "uncategorized":
        write_uncategorized_top_companies(worksheet, sections)
        auto_fit_columns(worksheet)


def export_sheet_data_to_excel(
    sheet_data: dict[str, list[dict]], categories_order: list[str], output_file_path: str
) -> None:
    workbook = Workbook()
    first_sheet = workbook.active
    if first_sheet is not None:
        workbook.remove(first_sheet)

    summary_sheet = workbook.create_sheet(title="Summary")
    write_summary_sheet(summary_sheet, sheet_data.get("Summary", []))

    uncategorized_sheet = workbook.create_sheet(title="Uncategorized")
    write_ledger_sheet(uncategorized_sheet, "Uncategorized", sheet_data.get("Uncategorized", []))

    for category in categories_order:
        category_sheet = workbook.create_sheet(title=safe_sheet_title(category))
        write_ledger_sheet(category_sheet, category, sheet_data.get(category, []))

    workbook.save(output_file_path)
