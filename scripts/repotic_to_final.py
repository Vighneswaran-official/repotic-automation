"""
scripts/repotic_to_final.py - Standalone CLI runner for the REPOTIC Automation.

Usage:
    python scripts/repotic_to_final.py [repotic_path] [ledger_path] [template_path] [output_path]

If paths are omitted, defaults to files inside the output/ directory:
    - output/upload_repotic.xlsx
    - output/upload_ledger.xlsx
    - output/upload_template.xlsx
    - output/Final_Output_Tamilnadu.xlsx
"""

import sys
from pathlib import Path

# Add project root to sys.path
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from src.config import OUTPUT_DIR, FIXED_OUTPUT
from src.services.processor import run_automation


def main():
    args = sys.argv[1:]
    repotic_path  = Path(args[0]) if len(args) > 0 else OUTPUT_DIR / "upload_repotic.xlsx"
    ledger_path   = Path(args[1]) if len(args) > 1 else OUTPUT_DIR / "upload_ledger.xlsx"
    template_path = Path(args[2]) if len(args) > 2 else OUTPUT_DIR / "upload_template.xlsx"
    output_path   = Path(args[3]) if len(args) > 3 else FIXED_OUTPUT

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

    res = run_automation(repotic_path, ledger_path, template_path, output_path)

    for line in res["logs"]:
        print(line)

    print(f"\n{'='*80}")
    print("SUMMARY")
    print(f"{'='*80}")
    for s in res["summary"]:
        print(f"  Sheet: {s['sheet']:<10} | Rows: {s['rows']:<4} | Net: ₹{s['Net_Amt']:>12,.2f} | Status: {s['status']}")

    if res["unmatched"]:
        print(f"\nUnmatched states:")
        for u in res["unmatched"]:
            print(f"  [{u['sheet']}] {u['state']}")

    print(f"\n[DONE] Successfully processed {res['total_rows']} rows -> {output_path}")


if __name__ == "__main__":
    main()
