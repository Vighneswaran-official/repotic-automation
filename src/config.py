"""
src/config.py - Central configuration and constants for REPOTIC Automation.
"""

import os
import tempfile
from pathlib import Path
from openpyxl.styles import PatternFill

# Project root directory (repotic-automation/)
PROJECT_ROOT = Path(__file__).resolve().parent.parent

def get_writable_dir() -> Path:
    """
    Return a guaranteed writable directory.
    Falls back to system temp dir in serverless or read-only environments (e.g. AWS Lambda, Vercel).
    """
    # 1. Custom env var if specified
    if os.environ.get("OUTPUT_DIR"):
        p = Path(os.environ["OUTPUT_DIR"])
        try:
            p.mkdir(parents=True, exist_ok=True)
            return p
        except Exception:
            pass

    # 2. Serverless detection (AWS Lambda / Vercel /var/task)
    if (
        os.environ.get("VERCEL")
        or os.environ.get("AWS_LAMBDA_FUNCTION_NAME")
        or str(Path(__file__)).startswith("/var/task")
    ):
        p = Path(tempfile.gettempdir()) / "repotic_output"
        p.mkdir(parents=True, exist_ok=True)
        return p

    # 3. Check if local directory is writable
    local_dir = PROJECT_ROOT / "output"
    try:
        local_dir.mkdir(exist_ok=True)
        test_file = local_dir / ".write_test"
        test_file.touch()
        test_file.unlink()
        return local_dir
    except (OSError, PermissionError):
        p = Path(tempfile.gettempdir()) / "repotic_output"
        p.mkdir(parents=True, exist_ok=True)
        return p


OUTPUT_DIR = get_writable_dir()
FIXED_OUTPUT = OUTPUT_DIR / "Final_Output_Tamilnadu.xlsx"

# ── Invoice numbering and accounting period ────────────────────────────────────
MARKETPLACE_PREFIX_MAP = {
    "flipkart": "FL",
    "amazon": "AM",
    "snapdeal": "SN",
    "myntra": "MY",
    "meesho": "ME",
}

DEFAULT_INVOICE_PREFIX = "08/26-27/"
DEFAULT_INVOICE_DATE   = "31-08-2026"


def compute_month_defaults(year: int, month: int) -> tuple:
    """
    Returns (inv_prefix, inv_date_str).
    Financial year starts in April (month 4).
    e.g. year=2026, month=8 -> ("08/26-27/", "31-08-2026")
    """
    import calendar
    if month >= 4:
        fy = f"{(year % 100):02d}-{((year + 1) % 100):02d}"
    else:
        fy = f"{((year - 1) % 100):02d}-{(year % 100):02d}"
    prefix = f"{month:02d}/{fy}/"
    last_day = calendar.monthrange(year, month)[1]
    inv_date = f"{last_day:02d}-{month:02d}-{year}"
    return prefix, inv_date


# ── Processing constants ───────────────────────────────────────────────────────
TARGET_SECTION = "STATE WISE SALES"
REPOTIC_SHEETS = ["Flipkart", "Meesho", "Snapdeal", "Amazon", "Myntra"]

# REPOTIC column (uppercase) -> template column name
COL_MAP = {
    "STATE"         : "StateOfSupply",
    "HSN CODE"      : "HSNCode",
    "RATE"          : "TaxPer",
    "QTY"           : "Qty",
    "TAXABLE VALUE" : "TaxableAmt",
    "IGST"          : "IGSTAmt",
    "CGST"          : "CGSTAmt",
    "SGST"          : "SGSTAmt",
    "INVOICE AMOUNT": "Net_Amt",
}

# Exact 26 output columns (matches template header order)
OUTPUT_COLS = [
    "InvNo", "Inv_Dt", "Pty_Name", "Vch_Type", "GSTIN",
    "StateOfSupply", "Product_Name", "HSNCode", "Qty", "UOM",
    "TaxPer", "TaxableAmt", "IGSTAmt", "SGSTAmt", "CGSTAmt",
    "Cess", "OtherAmt", "Net_Amt", "Narration", "Discount",
    "Sales Ledger", "PO No", "PO Date", "DC No", "DC Date",
    "Bill of Supply",
]

NUMERIC_COLS = {"TaxPer", "TaxableAmt", "IGSTAmt", "SGSTAmt", "CGSTAmt", "Net_Amt", "Qty"}

# Columns to validate output sums against source sums
VALIDATION_MAP = {
    "TaxableAmt" : "TAXABLE VALUE",
    "IGSTAmt"    : "IGST",
    "CGSTAmt"    : "CGST",
    "SGSTAmt"    : "SGST",
    "Net_Amt"    : "INVOICE AMOUNT",
}

# Yellow fill for unmatched cells in Excel
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
