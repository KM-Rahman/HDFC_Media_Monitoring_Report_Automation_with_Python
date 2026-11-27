"""
HDFC Bank Limited – Media Monitoring Report Automation
---------------------------------------------------
Portfolio demonstration script for processing News and Print Advertisement
Excel datasets. Uses anonymized/demo data and is not affiliated with HDFC Bank.

Run:
    python HDFC_Bank_Excel_Report_Automated_with_Python.py

Dependencies:
    pandas
    openpyxl
"""

from datetime import datetime
from pathlib import Path
import logging
import zipfile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Place the input Excel files in the same folder as this script.
PROJECT_DIR = Path(__file__).resolve().parent

NEWS_INPUT = PROJECT_DIR / "HDFC_News.xlsx"
NEWS_OUTPUT = PROJECT_DIR / "HDFC_news_processed.xlsx"

ADS_INPUT = PROJECT_DIR / "HDFC_Ads.xlsx"
ADS_TEMP = PROJECT_DIR / "ads_cleaned.xlsx"
ADS_OUTPUT = PROJECT_DIR / "HDFC_ads_processed.xlsx"

DATE_FORMAT = "%d-%b-%y"

NEWS_COLUMNS_TO_REMOVE = [
    "Region",
    "News Title & Link",
    "Genre",
    "Campaign",
    "Sub Category",
    "Source",
    "News Reporter",
    "Cir/Reach",
    "Keyword",
]

ADS_COLUMNS_TO_REMOVE = ["Ad Link", "Reporter"]

NEWS_COMPANIES = [
    "Bangladesh Bank (BB)",
    "BRAC Bank PLC",
    "Eastern Bank PLC",
    "Prime Bank PLC",
    "National Board of Revenue (NBR)",
    "HDFC Bank Limited",
    "Standard Chartered Bank (SCB)",
    "Banking Industry",
    "Citibank, N.A.",
    "Bangladesh Institute of Bank Management (BIBM)",
    "Bangladesh Association of Banks (BAB)",
    "Bangladesh Financial Intelligence Unit (BFIU)",
]

# News sorting priority is intentionally separate from the filtering list.
# The order below preserves the original priority sequence, using full names.
NEWS_SORT_PRIORITY = [
    ("HDFC Bank Limited", 1),
    ("Standard Chartered Bank (SCB)", 2),
    ("BRAC Bank PLC", 3),
    ("Eastern Bank PLC", 4),
    ("Prime Bank PLC", 5),
    ("Bangladesh Bank", 6),
    ("Bangladesh Financial Intelligence Unit (BFIU)", 7),
    ("National Board of Revenue (NBR)", 8),
    ("Banking Industry", 9),
]

ADVERTISING_COMPANIES = [
    "Digital Financial Service Industry",
    "Bangladesh Bank (BB)",
    "Mobile Financial Service Industry",
    "BRAC Bank PLC",
    "Eastern Bank PLC",
    "Prime Bank PLC",
    "National Board of Revenue (NBR)",
    "HDFC Bank Limited",
    "Standard Chartered Bank (SCB)",
    "Banking Industry",
    "Citibank, N.A.",
    "Bangladesh Institute of Bank Management (BIBM)",
    "Bangladesh Association of Banks (BAB)",
    "Bangladesh Financial Intelligence Unit (BFIU)",
]

# Configure readable console output.
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------

def strip_styles(input_xlsx: Path, cleaned_xlsx: Path) -> None:
    """Create a copy of an Excel workbook without its styles.xml component.

    This workaround is used when the source workbook's style definitions cause
    OpenPyXL loading errors. The original input workbook is left untouched.
    """
    with zipfile.ZipFile(input_xlsx, "r") as source_archive:
        with zipfile.ZipFile(cleaned_xlsx, "w") as cleaned_archive:
            for item in source_archive.infolist():
                if item.filename == "xl/styles.xml":
                    continue
                cleaned_archive.writestr(item, source_archive.read(item.filename))


def apply_formatting(excel_path: Path) -> None:
    """Apply borders, header styling, and basic column-width adjustment."""
    workbook = load_workbook(excel_path)
    worksheet = workbook.active

    thin_border_side = Side(border_style="thin", color="000000")
    cell_border = Border(
        top=thin_border_side,
        left=thin_border_side,
        right=thin_border_side,
        bottom=thin_border_side,
    )

    header_fill = PatternFill(fill_type="solid", fgColor="99CCFF")
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")

    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = cell_border

            if cell.row == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment

    # Adjust column widths based on the longest value in each column.
    for column_cells in worksheet.columns:
        column_letter = column_cells[0].column_letter
        longest_value = max(
            (len(str(cell.value)) for cell in column_cells if cell.value is not None),
            default=0,
        )
        worksheet.column_dimensions[column_letter].width = longest_value + 2

    workbook.save(excel_path)
    logger.info("Formatting applied: %s", excel_path.name)


def require_columns(dataframe: pd.DataFrame, required_columns: list[str], dataset_name: str) -> None:
    """Raise a clear error if expected columns are missing."""
    missing_columns = [column for column in required_columns if column not in dataframe.columns]
    if missing_columns:
        raise ValueError(
            f"{dataset_name} dataset is missing required column(s): "
            + ", ".join(missing_columns)
        )


# ---------------------------------------------------------------------------
# News report processing
# ---------------------------------------------------------------------------

def process_news_report() -> None:
    """Clean, prioritize, sort, export, and format the News dataset."""
    if not NEWS_INPUT.exists():
        logger.warning("News input file not found; skipping: %s", NEWS_INPUT.name)
        return

    logger.info("Processing News report...")
    dataframe = pd.read_excel(NEWS_INPUT, engine="openpyxl")
    require_columns(dataframe, ["Media Type", "News Date", "Company/Org"], "News")

    # Filter to the selected organizations before applying other transformations.
    dataframe = dataframe[dataframe["Company/Org"].isin(NEWS_COMPANIES)].copy()
    logger.info("News records retained for selected organizations: %d", len(dataframe))

    dataframe["Media Type"] = dataframe["Media Type"].replace("National", "Print")
    dataframe["News Date"] = datetime.now().strftime(DATE_FORMAT)
    dataframe = dataframe.drop(columns=NEWS_COLUMNS_TO_REMOVE, errors="ignore")

    # Sort by the original priority sequence, using full organization names.
    # Filtering and sorting are separate: other retained organizations receive
    # priority 10 and remain after the nine priority-ranked organizations.
    def get_news_sort_order(company_name):
        company_name = str(company_name)
        return next(
            (priority for full_name, priority in NEWS_SORT_PRIORITY
             if full_name.lower() in company_name.lower()),
            10,
        )

    dataframe["_sort_order"] = dataframe["Company/Org"].apply(get_news_sort_order)
    dataframe = dataframe.sort_values("_sort_order", kind="stable").drop(columns="_sort_order")

    dataframe.to_excel(NEWS_OUTPUT, index=False, engine="openpyxl")
    apply_formatting(NEWS_OUTPUT)
    logger.info("News report completed: %s", NEWS_OUTPUT.name)


# ---------------------------------------------------------------------------
# Advertisement report processing
# ---------------------------------------------------------------------------

def process_advertisement_report() -> None:
    """Clean, filter, export, and format the Print Advertisement dataset."""
    if not ADS_INPUT.exists():
        logger.warning("Advertisement input file not found; skipping: %s", ADS_INPUT.name)
        return

    logger.info("Processing Advertisement report...")

    try:
        # Remove problematic style definitions from a temporary workbook copy.
        strip_styles(ADS_INPUT, ADS_TEMP)

        dataframe = pd.read_excel(ADS_TEMP, engine="openpyxl")
        require_columns(
            dataframe,
            ["Media Type", "Ad Date", "Company"],
            "Advertisement",
        )

        # Filter to the selected organizations before applying other transformations.
        dataframe = dataframe[dataframe["Company"].isin(ADVERTISING_COMPANIES)].copy()
        logger.info("Advertisement records retained for selected organizations: %d", len(dataframe))

        dataframe["Media Type"] = dataframe["Media Type"].replace("National", "Print")
        dataframe["Ad Date"] = datetime.now().strftime(DATE_FORMAT)
        dataframe = dataframe.drop(columns=ADS_COLUMNS_TO_REMOVE, errors="ignore")

        # Retain Print advertisements only.
        dataframe = dataframe[dataframe["Media Type"] == "Print"]

        # Sort alphabetically according to the full company names in the configured list.
        ads_sort_priority = {company: priority for priority, company in enumerate(ADVERTISING_COMPANIES, start=1)}
        dataframe["_sort_order"] = dataframe["Company"].map(ads_sort_priority)
        dataframe = dataframe.sort_values("_sort_order", kind="stable").drop(columns="_sort_order")

        dataframe.to_excel(ADS_OUTPUT, index=False, engine="openpyxl")
        apply_formatting(ADS_OUTPUT)
        logger.info("Advertisement report completed: %s", ADS_OUTPUT.name)

    finally:
        # Remove the intermediate cleaned copy, whether processing succeeds or fails.
        if ADS_TEMP.exists():
            ADS_TEMP.unlink()


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main() -> None:
    """Run both report-processing workflows."""
    logger.info("Starting media monitoring report automation.")

    process_news_report()
    process_advertisement_report()

    logger.info("Automation run finished.")


if __name__ == "__main__":
    main()
