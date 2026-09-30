"""
scripts/repotic_to_final.py - Standalone CLI runner for the REPOTIC Automation.

Usage:
    python scripts/repotic_to_final.py [repotic_path] [ledger_path] [template_path] [output_path] [--month YYYY-MM]

Examples:
    python scripts/repotic_to_final.py
    python scripts/repotic_to_final.py --month 2027-04
    python scripts/repotic_to_final.py input/repotic.xlsx input/ledger.xlsx input/template.xlsx output/final.xlsx --month 2026-08
"""

import sys
import argparse
from pathlib import Path

# Add project root to sys.path
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from src.config import OUTPUT_DIR, FIXED_OUTPUT, PREFIX
from src.services.processor import run_automation


def main():
    parser = argparse.ArgumentParser(description="REPOTIC Automation CLI")
    parser.add_argument("repotic", nargs="?", default=None, help="Path to REPOTIC file")
    parser.add_argument("ledger", nargs="?", default=None, help="Path to ledger file")
    parser.add_argument("template", nargs="?", default=None, help="Path to template file")
    parser.add_argument("output", nargs="?", default=None, help="Path to output file")
    parser.add_argument("--month", default=None, help="Report month in YYYY-MM format (e.g. 2026-08 or 2027-04)")

    args = parser.parse_args()

    repotic_path  = Path(args.repotic) if args.repotic else OUTPUT_DIR / "upload_repotic.xlsx"
    ledger_path   = Path(args.ledger) if args.ledger else OUTPUT_DIR / "upload_ledger.xlsx"
    template_path = Path(args.template) if args.template else OUTPUT_DIR / "upload_template.xlsx"
    output_path   = Path(args.output) if args.output else FIXED_OUTPUT

    for label, p in [("REPOTIC", repotic_path), ("Ledger", ledger_path), ("Template", template_path)]:
        if not p.exists():
            print(f"[ERROR] Required input file not found: {label} ({p})")
            print("Please provide valid file paths or ensure files exist in output/.")
            sys.exit(1)

    print(f"[START] Running REPOTIC Automation via CLI...")
    print(f"  REPOTIC : {repotic_path}")
    print(f"  Ledger  : {ledger_path}")
    print(f"  Template: {template_path}")
    print(f"  Output  : {output_path}")
    if args.month:
        print(f"  Month override: {args.month}")

    res = run_automation(
        repotic_path, ledger_path, template_path, output_path,
        month_override=args.month
    )

    for line in res["logs"]:
        print(line)

    print(f"\n{'='*80}")
    print("SUMMARY")
    print(f"{'='*80}")
    for s in res["summary"]:
        print(f"  Sheet: {s['sheet']:<10} | Rows: {s['rows']:<4} | Series: {s.get('invoice_series', 'N/A'):<32} | Net: ₹{s['Net_Amt']:>12,.2f} | Status: {s['status']}")

    if res["unmatched"]:
        print(f"\nUnmatched states:")
        for u in res["unmatched"]:
            print(f"  [{u['sheet']}] {u['state']}")

    print(f"\n[DONE] Successfully processed {res['total_rows']} rows -> {output_path}")


if __name__ == "__main__":
    main()
