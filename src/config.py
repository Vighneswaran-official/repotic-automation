"""
src/config.py - Central configuration and constants for REPOTIC Automation.
"""

import os
import calendar
import tempfile
from datetime import datetime, date
from collections import Counter
from pathlib import Path
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

# ── Marketplace Codes (editable dict) ──────────────────────────────────────────
PREFIX = {
    "FLIPKART": "FL",
    "AMAZON": "AM",
    "SNAPDEAL": "SN",
    "MYNTRA": "MY",
    "MEESHO": "ME",
}

# Project root directory (repotic-automation/)
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def get_writable_dir() -> Path:
    """
    Return a guaranteed writable directory.
    Falls back to system temp dir in serverless or read-only environments (e.g. AWS Lambda, Vercel).
    """
    if os.environ.get("OUTPUT_DIR"):
        p = Path(os.environ["OUTPUT_DIR"])
        try:
            p.mkdir(parents=True, exist_ok=True)
            return p
        except Exception:
            pass

    if (
        os.environ.get("VERCEL")
        or os.environ.get("AWS_LAMBDA_FUNCTION_NAME")
        or str(Path(__file__)).startswith("/var/task")
    ):
        p = Path(tempfile.gettempdir()) / "repotic_output"
        p.mkdir(parents=True, exist_ok=True)
        return p

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


# ── Financial Year & Period Helpers ───────────────────────────────────────────
def fin_year(month: int, year: int) -> str:
    """
    Calculate financial year (1 April to 31 March).
    e.g. (3, 2027) -> "26-27", (4, 2027) -> "27-28"
    """
    start = year if month >= 4 else year - 1
    return f"{start % 100:02d}-{(start + 1) % 100:02d}"


def get_previous_calendar_month() -> tuple:
    """Return (year, month) of the previous calendar month relative to today."""
    today = date.today()
    if today.month == 1:
        return (today.year - 1, 12)
    return (today.year, today.month - 1)


def detect_month_from_repotic(repotic_path: Path):
    """
    Auto-detect (year, month) from the DATE column in the "B2B INVOICE" section
    of the REPOTIC sheet "Other B2B & CDNR".
    Returns (year, month) or None if not found.
    """
    try:
        wb = load_workbook(repotic_path, data_only=True)
        sheet_candidates = [
            s for s in wb.sheetnames
            if "other b2b" in s.strip().lower() or "cdnr" in s.strip().lower()
        ]
        if not sheet_candidates:
            return None

        ws = wb[sheet_candidates[0]]
        all_rows = list(ws.iter_rows(values_only=True))

        # Find "B2B INVOICE" section
        b2b_idx = None
        for i, row in enumerate(all_rows):
            for cell in row:
                if cell is not None and "B2B INVOICE" in str(cell).strip().upper():
                    b2b_idx = i
                    break
            if b2b_idx is not None:
                break

        start_row = (b2b_idx + 1) if b2b_idx is not None else 0

        # Find header row containing "DATE"
        date_col_idx = None
        header_row_idx = None
        for i in range(start_row, min(start_row + 15, len(all_rows))):
            row = all_rows[i]
            for col_idx, cell in enumerate(row):
                if cell is not None and str(cell).strip().upper() == "DATE":
                    date_col_idx = col_idx
                    header_row_idx = i
                    break
            if date_col_idx is not None:
                break

        # Fallback search if not found directly under section title
        if date_col_idx is None:
            for i, row in enumerate(all_rows):
                for col_idx, cell in enumerate(row):
                    if cell is not None and str(cell).strip().upper() == "DATE":
                        date_col_idx = col_idx
                        header_row_idx = i
                        break
                if date_col_idx is not None:
                    break

        if date_col_idx is None or header_row_idx is None:
            return None

        detected_months = []
        for i in range(header_row_idx + 1, len(all_rows)):
            row = all_rows[i]
            if date_col_idx >= len(row):
                continue
            val = row[date_col_idx]
            if val is None:
                continue

            if isinstance(val, (datetime, date)):
                detected_months.append((val.year, val.month))
                continue

            val_str = str(val).strip()
            if not val_str:
                continue

            for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y"):
                try:
                    dt = datetime.strptime(val_str, fmt)
                    detected_months.append((dt.year, dt.month))
                    break
                except ValueError:
                    pass

        if not detected_months:
            return None

        most_common = Counter(detected_months).most_common(1)[0][0]
        return most_common
    except Exception:
        return None


def resolve_report_period(repotic_path: Path = None, month_override: str = None) -> tuple:
    """
    Resolve (year, month, source_desc).
    Priority:
      1. UI / CLI month_override ("YYYY-MM")
      2. Auto-detect from B2B dates in "Other B2B & CDNR"
      3. Previous calendar month relative to today
    """
    if month_override and str(month_override).strip():
        parts = str(month_override).strip().split("-")
        if len(parts) == 2:
            try:
                y, m = int(parts[0]), int(parts[1])
                if 1 <= m <= 12:
                    return y, m, "UI / user selection"
            except ValueError:
                pass

    if repotic_path and Path(repotic_path).exists():
        detected = detect_month_from_repotic(Path(repotic_path))
        if detected:
            return detected[0], detected[1], "B2B dates"

    prev = get_previous_calendar_month()
    return prev[0], prev[1], "assumed"


# ── Processing constants ───────────────────────────────────────────────────────
TARGET_SECTION = "STATE WISE SALES"

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
