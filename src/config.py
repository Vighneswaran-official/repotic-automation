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

# ── Processing constants ───────────────────────────────────────────────────────
TARGET_SECTION = "STATE WISE SALES"
REPOTIC_SHEETS = ["Flipkart", "Meesho", "Snapdeal"]

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
