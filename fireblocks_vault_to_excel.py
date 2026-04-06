from __future__ import annotations

# Built-in tools used to read CSV files, work with dates, and handle folders/files.
import csv
import re
from datetime import datetime
from pathlib import Path
from typing import Iterable

# openpyxl is the library used to create and format the Excel workbook.
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# -----------------------------------------------------------------------------
# SETTINGS
# These values tell the script where to find the input files, where to save the
# output file, and which currencies we want to keep.
# -----------------------------------------------------------------------------
BASE_DIR = Path(r"G:\Shared drives\Lumepay (Restricted)\Accounting\Miscellaneous\Accounting Devs\Fireblocks Vault Build-Up")
CSV_DIR = BASE_DIR / "Vault CSV"
OUTPUT_DIR = BASE_DIR / "Output"

# These are the only assets we want to keep after importing the CSV data.
ALLOWED_ASSETS = ["TRX", "ETH", "USDT_ERC20", "TRX_USDT_S2UZ"]

# These two asset codes should be combined into one USDT worksheet.
USDT_ASSETS = {"USDT_ERC20", "TRX_USDT_S2UZ"}

# These formats control how dates and balance numbers appear in Excel.
DATE_NUMBER_FORMAT = "yyyy/mm/dd hh:mm"
ROUNDED_DATE_NUMBER_FORMAT = "yyyy/mm/dd"
COMMA_STYLE_FORMAT = "#,##0.00"
TABLE_COLUMN_WIDTH = 16.5

# These fill colours are used on the USDT sheet for date and balance sections.
BLUE_HEADER_FILL = PatternFill(fill_type="solid", fgColor="538DD5")
BLUE_DATA_FILL = PatternFill(fill_type="solid", fgColor="C5D9F1")
YELLOW_HEADER_FILL = PatternFill(fill_type="solid", fgColor="FFE599")
YELLOW_DATA_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")


# -----------------------------------------------------------------------------
# FILE DISCOVERY
# This function looks inside the Vault CSV folder and makes sure there are
# exactly two CSV files to import.
# -----------------------------------------------------------------------------
def list_csv_files(folder: Path) -> list[Path]:
    csv_files = sorted(folder.glob("*.csv"))
    if len(csv_files) != 2:
        raise ValueError(
            f"Expected exactly 2 CSV files in '{folder}', found {len(csv_files)}."
        )
    return csv_files


# -----------------------------------------------------------------------------
# DESTINATION FILE DETAILS
# The Destination CSV gives us the vault name and the two destination wallet
# addresses we need for the workbook layout and the USDT balance formulas.
# -----------------------------------------------------------------------------
def find_destination_csv(csv_files: list[Path]) -> Path:
    for csv_file in csv_files:
        if "destination" in csv_file.name.lower():
            return csv_file
    raise ValueError("Could not find the CSV file with 'Destination' in the filename.")



def read_destination_rows(destination_csv: Path) -> list[list[str]]:
    with destination_csv.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        rows = list(reader)

    if len(rows) < 2:
        raise ValueError(f"Destination CSV does not contain a data row: {destination_csv}")

    return rows



def read_vault_details(destination_csv: Path) -> tuple[str, str, str]:
    rows = read_destination_rows(destination_csv)
    first_data_row = rows[1]

    # Column R = Destination, which is the vault name.
    vault_name = first_data_row[17].strip()
    if not vault_name:
        raise ValueError("Vault name in column R of the Destination CSV is blank.")

    trx_address = ""
    eth_address = ""

    for row in rows[1:]:
        if len(row) <= 18:
            continue

        asset = row[4].strip()
        destination_address = row[18].strip()

        if asset == "TRX" and not trx_address:
            trx_address = destination_address
        if asset == "ETH" and not eth_address:
            eth_address = destination_address

        if trx_address and eth_address:
            break

    if not trx_address:
        raise ValueError("Could not find a TRX destination address in the Destination CSV.")
    if not eth_address:
        raise ValueError("Could not find an ETH destination address in the Destination CSV.")

    return vault_name, trx_address, eth_address



def build_output_file(vault_name: str) -> Path:
    safe_name = re.sub(r'[<>:"/\\|?*]', "_", vault_name).strip()
    if not safe_name:
        safe_name = "Vault"
    return OUTPUT_DIR / f"{safe_name} Asset Build-Up.xlsx"


# -----------------------------------------------------------------------------
# CSV IMPORT
# This function reads one CSV file, keeps only columns A:AC, and returns:
# 1. the header row
# 2. the data rows
# Blank rows are ignored.
# -----------------------------------------------------------------------------
def read_trimmed_rows(csv_path: Path) -> tuple[list[str], list[list[str]]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        rows = list(reader)

    if not rows:
        raise ValueError(f"CSV file is empty: {csv_path}")

    # Keep only the first 29 columns, which is Excel columns A to AC.
    header = rows[0][:29]
    data_rows = [row[:29] for row in rows[1:] if any(cell.strip() for cell in row)]
    return header, data_rows


# -----------------------------------------------------------------------------
# DATE CONVERSION
# Fireblocks exports dates like "05 Apr 2026 08:32:53 GMT".
# This function converts that text into a real Excel-friendly date/time value.
# -----------------------------------------------------------------------------
def parse_fireblocks_date(raw_value: str) -> datetime:
    # The first 20 characters contain the part we need, without the timezone text.
    trimmed = raw_value[:20].strip()
    return datetime.strptime(trimmed, "%d %b %Y %H:%M:%S")


# -----------------------------------------------------------------------------
# FILTERING RULES
# These small helper functions keep only the rows we want.
# -----------------------------------------------------------------------------
def filter_completed(rows: Iterable[list[str]]) -> list[list[str]]:
    # Keep only rows where Status (column B) is COMPLETED.
    return [row for row in rows if len(row) > 1 and row[1].strip().upper() == "COMPLETED"]



def filter_assets(rows: Iterable[list]) -> list[list]:
    # Keep only the four asset types requested by the user.
    return [row for row in rows if len(row) > 4 and str(row[4]).strip() in ALLOWED_ASSETS]


# -----------------------------------------------------------------------------
# NUMBER CLEANUP
# This helper safely turns a text number into a Python number.
# If the value is blank, it returns 0.
# -----------------------------------------------------------------------------
def to_decimal_amount(value) -> float:
    if value in (None, ""):
        return 0.0
    return float(value)



def convert_numeric_columns(row: list[str]) -> list:
    converted_row = list(row)

    # Columns H to M should be stored as real numbers in Excel, not as text.
    for index in range(7, 13):
        converted_row[index] = to_decimal_amount(converted_row[index])

    return converted_row


# -----------------------------------------------------------------------------
# EXCEL FORMATTING HELPERS
# These functions make the workbook easier to read once it has been created.
# -----------------------------------------------------------------------------
def set_table_column_widths(worksheet, max_column: int) -> None:
    # Set all table columns to the fixed width requested by the user.
    for column_idx in range(1, max_column + 1):
        worksheet.column_dimensions[get_column_letter(column_idx)].width = TABLE_COLUMN_WIDTH



def apply_table_formats(worksheet, header_row: int, first_data_row: int, max_column: int) -> None:
    # Make the header row bold.
    for cell in worksheet[header_row]:
        cell.font = Font(bold=True)

    # Turn on filters only for the actual table range.
    worksheet.auto_filter.ref = f"A{header_row}:{get_column_letter(max_column)}{worksheet.max_row}"

    # Freeze the sheet just below the header row.
    worksheet.freeze_panes = f"A{first_data_row}"

    # Columns H to M contain numeric values and should display in comma style.
    for row_idx in range(first_data_row, worksheet.max_row + 1):
        for column_idx in range(8, 14):
            worksheet.cell(row=row_idx, column=column_idx).number_format = COMMA_STYLE_FORMAT

    # Columns AD and AE contain real Excel date values, so we format them here.
    for row_idx in range(first_data_row, worksheet.max_row + 1):
        worksheet.cell(row=row_idx, column=30).number_format = DATE_NUMBER_FORMAT
        worksheet.cell(row=row_idx, column=31).number_format = ROUNDED_DATE_NUMBER_FORMAT

    set_table_column_widths(worksheet, max_column)



def style_top_address_rows(worksheet) -> None:
    worksheet["A1"].font = Font(bold=True)
    worksheet["A2"].font = Font(bold=True)



def apply_usdt_colour_blocks(worksheet, header_row: int, first_data_row: int) -> None:
    # Colour the date columns AD and AE.
    for column_idx in range(30, 32):
        worksheet.cell(row=header_row, column=column_idx).fill = BLUE_HEADER_FILL
        for row_idx in range(first_data_row, worksheet.max_row + 1):
            worksheet.cell(row=row_idx, column=column_idx).fill = BLUE_DATA_FILL

    # Colour the balance columns AF to AH.
    for column_idx in range(32, 35):
        worksheet.cell(row=header_row, column=column_idx).fill = YELLOW_HEADER_FILL
        for row_idx in range(first_data_row, worksheet.max_row + 1):
            worksheet.cell(row=row_idx, column=column_idx).fill = YELLOW_DATA_FILL


# -----------------------------------------------------------------------------
# SHEET WRITING
# These functions create the workbook tabs and place the data in the requested
# layout.
# -----------------------------------------------------------------------------
def write_standard_sheet(workbook: Workbook, title: str, header: list[str], rows: list[list]) -> None:
    worksheet = workbook.create_sheet(title=title)
    worksheet.append(header)

    for row in rows:
        worksheet.append(row)

    apply_table_formats(worksheet, header_row=1, first_data_row=2, max_column=len(header))



def write_currency_sheet(
    workbook: Workbook,
    title: str,
    header: list[str],
    rows: list[list],
    vault_name: str,
    trx_address: str,
    eth_address: str,
) -> None:
    worksheet = workbook.create_sheet(title=title)

    # Rows 1 and 2 show the two reference wallet addresses for the vault.
    worksheet["A1"] = f"{vault_name} TRX Address"
    worksheet["B1"] = trx_address
    worksheet["A2"] = f"{vault_name} ETH Address"
    worksheet["B2"] = eth_address

    # Row 3 is intentionally left blank.
    # Row 4 contains the table headers.
    worksheet.append([])
    worksheet.append(header)

    # Data starts on row 5.
    for row in rows:
        worksheet.append(row)

    style_top_address_rows(worksheet)
    apply_table_formats(worksheet, header_row=4, first_data_row=5, max_column=len(header))


# -----------------------------------------------------------------------------
# USDT BALANCE CALCULATION
# This creates the USDT worksheet data by combining USDT_ERC20 and
# TRX_USDT_S2UZ into one sheet, then adding live Excel formulas for:
# AF = Opening balance
# AG = Inflow/(Outflow)
# AH = Closing balance
# -----------------------------------------------------------------------------
def build_usdt_rows(rows: list[list]) -> tuple[list[str], list[list]]:
    # Add the three extra balance columns after the existing AE column.
    header = rows[0][:31] + ["Opening balance", "Inflow/(Outflow)", "Closing balance"]

    # Keep only the rows that belong in the combined USDT sheet.
    usdt_source_rows = [row[:31] for row in rows[1:] if str(row[4]).strip() in USDT_ASSETS]

    output_rows: list[list] = []

    # Data starts on row 5 in the currency sheets, so the first formula row is 5.
    for excel_row_number, row in enumerate(usdt_source_rows, start=5):
        if excel_row_number == 5:
            opening_balance = 0
        else:
            opening_balance = f"=AH{excel_row_number - 1}"

        # If the destination address in column S matches either wallet address
        # shown at the top of the sheet, treat it as an inflow. Otherwise it is
        # an outflow and should be negative.
        inflow_outflow = f"=IF(OR(S{excel_row_number}=$B$1,S{excel_row_number}=$B$2),J{excel_row_number},-J{excel_row_number})"

        # Closing balance is opening balance plus movement.
        closing_balance = f"=AF{excel_row_number}+AG{excel_row_number}"

        output_rows.append(row + [opening_balance, inflow_outflow, closing_balance])

    return header, output_rows


# -----------------------------------------------------------------------------
# USDT NUMBER FORMATTING
# This applies number formatting to the three balance columns on the USDT sheet.
# -----------------------------------------------------------------------------
def apply_usdt_number_formats(worksheet, first_data_row: int) -> None:
    for row_idx in range(first_data_row, worksheet.max_row + 1):
        for column_idx in range(32, 35):
            worksheet.cell(row=row_idx, column=column_idx).number_format = COMMA_STYLE_FORMAT


# -----------------------------------------------------------------------------
# MAIN PROGRAM
# This is the part that runs the full process from start to finish.
# -----------------------------------------------------------------------------
def main() -> None:
    # Find the two CSV files in the input folder.
    csv_files = list_csv_files(CSV_DIR)

    # Read vault details and the TRX/ETH destination addresses from the file
    # whose name contains "Destination".
    destination_csv = find_destination_csv(csv_files)
    vault_name, trx_address, eth_address = read_vault_details(destination_csv)
    output_file = build_output_file(vault_name)

    # These variables will hold all imported rows and the shared header row.
    combined_rows: list[list[str]] = []
    trimmed_header: list[str] | None = None

    # Import each CSV and combine the rows into one list.
    for csv_file in csv_files:
        header, rows = read_trimmed_rows(csv_file)
        if trimmed_header is None:
            trimmed_header = header
        elif header != trimmed_header:
            raise ValueError(f"CSV headers do not match in file: {csv_file}")
        combined_rows.extend(rows)

    if trimmed_header is None:
        raise ValueError("No CSV header was loaded.")

    # Keep only rows where Status is COMPLETED.
    completed_rows = filter_completed(combined_rows)

    # Add the two new date columns that will become AD and AE in Excel.
    extended_header = trimmed_header + ["Date", "Date rounded"]
    enriched_rows: list[list] = []

    for row in completed_rows:
        # Turn H:M into real numbers before writing to Excel.
        converted_row = convert_numeric_columns(row)

        # Create the full date/time value from column D.
        parsed_date = parse_fireblocks_date(converted_row[3])

        # Create the rounded date value with time set to midnight.
        rounded_date = parsed_date.replace(hour=0, minute=0, second=0, microsecond=0)

        # Save the original row plus the two new date fields.
        enriched_rows.append(converted_row + [parsed_date, rounded_date])

    # Keep only the four requested assets.
    filtered_rows = filter_assets(enriched_rows)

    # Sort the remaining rows by the new Date column (AD), oldest to newest.
    filtered_rows.sort(key=lambda row: row[29])

    # Start a new Excel workbook and remove the default blank sheet.
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    # Create the main consolidated sheet containing all kept transactions.
    write_standard_sheet(workbook, "Consolidated", extended_header, filtered_rows)

    # Split out TRX and ETH into their own separate sheets.
    trx_rows = [row for row in filtered_rows if str(row[4]).strip() == "TRX"]
    eth_rows = [row for row in filtered_rows if str(row[4]).strip() == "ETH"]

    write_currency_sheet(workbook, "TRX", extended_header, trx_rows, vault_name, trx_address, eth_address)
    write_currency_sheet(workbook, "ETH", extended_header, eth_rows, vault_name, trx_address, eth_address)

    # Build the combined USDT sheet with running balance formulas.
    usdt_header, usdt_rows = build_usdt_rows([extended_header] + filtered_rows)
    write_currency_sheet(workbook, "USDT", usdt_header, usdt_rows, vault_name, trx_address, eth_address)
    apply_usdt_number_formats(workbook["USDT"], first_data_row=5)
    apply_usdt_colour_blocks(workbook["USDT"], header_row=4, first_data_row=5)

    # Make sure the output folder exists, then save the workbook.
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook.save(output_file)

    # Print a short summary in the terminal after the file is created.
    print(f"Vault name: {vault_name}")
    print(f"TRX address: {trx_address}")
    print(f"ETH address: {eth_address}")
    print(f"Workbook created: {output_file}")
    print(f"Imported CSV files: {', '.join(str(path.name) for path in csv_files)}")
    print(f"Completed rows kept: {len(completed_rows)}")
    print(f"Rows kept after asset filter: {len(filtered_rows)}")
    print(f"USDT rows kept: {len(usdt_rows)}")


# This makes sure the script only runs when you execute this file directly.
if __name__ == "__main__":
    main()
