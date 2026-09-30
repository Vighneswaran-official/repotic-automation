# REPOTIC Automation — Project Context for Claude

> Paste this file (or its contents) at the start of any Claude conversation
> so Claude has full context about this project.

---

## What This Project Does

Reads marketplace sales data from **REPOTIC.xlsx** (Flipkart, Meesho, Snapdeal),
performs a state-level ledger lookup from **AMAZON_LEDGER FILE.xlsx**, and
writes all mapped rows into a copy of **Final Sample Format Tamilnadu.xlsx**
(26-column Tally import template), saving the result as
**Final_Output_Tamilnadu.xlsx**.

There is also a **Flask web app** (`app.py`) that wraps the same logic
behind a drag-and-drop browser UI.

---

## File Structure

```
repotic_automation/
├── app.py                          ← Flask web server (main entry point)
├── repotic_to_final.py             ← Standalone CLI script (same logic)
├── templates/
│   └── index.html                  ← Dark-themed drag-and-drop UI
├── output/                         ← Generated at runtime (gitignore this)
│   ├── Final_Output_Tamilnadu.xlsx ← The generated output (fixed path)
│   ├── upload_repotic.xlsx         ← Last uploaded REPOTIC file
│   ├── upload_ledger.xlsx          ← Last uploaded ledger file
│   └── upload_template.xlsx        ← Last uploaded template file
└── AMAZON/                         ← Original source files (read-only)
    ├── REPOTIC.xlsx
    ├── AMAZON_LEDGER  FILE.xlsx
    └── Final Sample Format Tamilnadu.xlsx
```

---

## Input Files

### 1. REPOTIC.xlsx
Sheets: `Flipkart`, `Meesho`, `Snapdeal`, `Other B2B & CDNR`

Each sheet has **3 stacked sections** (detect by title text in column A):
- `INTER STATE SALES`  → **skip** (summary totals only)
- `INTRA STATE SALES`  → **skip** (summary totals only)
- `STATE WISE SALES`   → **read this section only**

The STATE WISE SALES section structure:
```
Row N:   "STATE WISE SALES"   ← title row (col A)
Row N+1: STATE | HSN CODE | RATE | QTY | TAXABLE VALUE | IGST | CGST | SGST | INVOICE AMOUNT
Row N+2: MAHARASHTRA | 61091000 | 5 | 89 | 22017.12 | 1100.96 | 0 | 0 | 23118.08
...      (one row per state, until end of sheet)
```

**Special case — Snapdeal:** HSN CODE is blank and QTY is 0 for all rows.
Handle gracefully (leave those cells empty, don't error out).

### 2. AMAZON_LEDGER FILE.xlsx  (sheet: `Amazon`)
Columns: `StateOfSupply` | `Sales Ledger` | `Pty_Name`

Lookup key matching rules:
- Case-insensitive
- Collapse extra whitespace
- REPOTIC uses ALL CAPS (`TAMIL NADU`); ledger uses Title Case (`Tamil Nadu`)
- Tamil Nadu → `AMAZON SALES LOCAL GST 5%` / `AMAZON SALES LOCAL`
- All other states → `AMAZON SALES IGST 5%` / `Unreg Ecom Amazon Sales Igst(<State>)`

Known states **not in the ledger** (handled by fallback):
- `TRIPURA` → auto-assigned IGST ledger
- `DADRA AND NAGAR HAVELI AND DAMAN AND DIU` → auto-assigned IGST ledger

Fallback logic (in `lookup_state()`):
```python
if CGSTAmt > 0:   # intra-state
    return "AMAZON SALES LOCAL", "AMAZON SALES LOCAL GST 5%"
else:             # inter-state
    pty  = f"Unreg Ecom Amazon Sales Igst({title_case_state})"
    ledg = "AMAZON SALES IGST 5%"
    return pty, ledg
```

### 3. Final Sample Format Tamilnadu.xlsx  (sheet: `Sheet1`)
Empty 26-column template. Headers in row 1, bold, yellow fill on some cols.
The output renames the sheet to `"Final Sample Format Tamilnadu"`.

---

## Column Mapping  (REPOTIC → Template)

| REPOTIC Column | Template Column |
|----------------|-----------------|
| STATE          | StateOfSupply   |
| RATE           | TaxPer          |
| QTY            | Qty             |
| TAXABLE VALUE  | TaxableAmt      |
| IGST           | IGSTAmt         |
| CGST           | CGSTAmt         |
| SGST           | SGSTAmt         |
| *(auto)*       | InvNo           | Format: `MM/YY-YY/XX-NN` (e.g. `08/26-27/FL-01`) |
| *(auto)*       | Inv_Dt          | Month-end date (e.g. `31-08-2026`) |

Other template columns (HSNCode, GSTIN, Product_Name, UOM, Cess, OtherAmt,
Narration, Discount, PO No, PO Date, DC No, DC Date) are left **blank**.
HSN code is not required and is intentionally left empty.
`Vch_Type` is set to `"Auto Sales"`, and `Bill of Supply` is set to `StateOfSupply`.
`InvNo` is stored explicitly as text (format `@`).

Numeric columns are stored as `float`, rounded to 2 decimal places,
with Excel number format `#,##0.00`.

---

## Output File

- Path: `output/Final_Output_Tamilnadu.xlsx`
- Sheet name: `Final Sample Format Tamilnadu`
- Row 1: 26 headers (preserved from template, original formatting kept)
- Rows 2+: Flipkart data, then Meesho data, then Snapdeal data (in order)
- Total rows (current data): 85 data rows across 3 sheets

---

## Flask Web App  (`app.py`)

### Routes
| Route | Method | Description |
|-------|--------|-------------|
| `/`   | GET    | Serves `templates/index.html` with no-cache headers |
| `/run`| POST   | Accepts 3 file uploads, runs automation, returns JSON |
| `/download` | GET | Serves `output/Final_Output_Tamilnadu.xlsx` as download |

### Key Design Decisions
1. **Fixed output path** (`OUTPUT_DIR / "Final_Output_Tamilnadu.xlsx"`) — not temp dir,
   survives Flask debug-mode auto-reloads.
2. **No-cache headers** on `/` — browser always gets fresh HTML/JS.
3. **Download uses `make_response(bytes)`** + explicit `Content-Disposition` header —
   avoids Flask `send_file()` filename issues.
4. **Frontend download uses `fetch()` + blob + `a.download`** — guarantees correct
   `.xlsx` filename in browser download bar.

### JSON Response from `/run`
```json
{
  "ok": true,
  "logs": ["Ledger loaded: 32 state entries", "─── Processing: Flipkart ───", ...],
  "summary": [
    {
      "sheet": "Flipkart",
      "rows": 31,
      "TaxableAmt": 182004.88,
      "IGSTAmt": 8264.63,
      "CGSTAmt": 418.26,
      "SGSTAmt": 418.26,
      "Net_Amt": 191106.03,
      "status": "OK",
      "validation": [{"col": "TaxableAmt", "source": 182004.88, "output": 182004.88, "ok": true}, ...]
    },
    ...
  ],
  "unmatched": ["DADRA AND NAGAR HAVELI AND DAMAN AND DIU", "TRIPURA"],
  "preview": [{"StateOfSupply": "MAHARASHTRA", "Qty": 89, ...}, ...],
  "total_rows": 85
}
```
> All numeric values are plain Python `float`/`int` (not numpy types) to avoid
> `TypeError: Object of type bool is not JSON serializable`.

---

## Known Bugs Fixed

| Bug | Root Cause | Fix |
|-----|-----------|-----|
| `Sales Ledger` column always blank | `itertuples()` converts `"Sales Ledger"` → `"Sales_Ledger"` (spaces → underscores) | Switched to `.to_dict(orient="records")` which preserves exact column names |
| `TypeError: bool not JSON serializable` | `abs(src - out) <= 0.02` returns numpy `bool_`, not Python `bool` | Wrapped with `bool()`, `float()`, `int()` |
| Download served with UUID filename | Old `<a href="/download">` caused page navigation; Flask temp dir wiped on reload | Switched to `fetch()` + blob + `a.download`; fixed output path |
| Sheet named `Sheet1` | Template sheet not renamed | Added `ws_out.title = "Final Sample Format Tamilnadu"` |
| Stale JS after code change | Browser cached old HTML | Added `Cache-Control: no-store` header to `/` route |
| TRIPURA / DADRA left blank | Not in ledger file | Added smart fallback: IGST if inter-state, LOCAL if intra-state |

---

## Running the Project

### CLI script
```powershell
cd c:\Users\knitk\.gemini\antigravity-ide\scratch\repotic_automation
python repotic_to_final.py
```

### Flask web app
```powershell
cd c:\Users\knitk\.gemini\antigravity-ide\scratch\repotic_automation
python app.py
# Open http://127.0.0.1:5000
```

### Dependencies
```
pandas
openpyxl
flask
```
Install: `pip install pandas openpyxl flask`

Python version: **3.14** (Windows)

---

## Validation Logic

After processing each sheet, sums are compared against the source:

| Output Column | Source Column  | Tolerance |
|---------------|----------------|-----------|
| TaxableAmt    | TAXABLE VALUE  | ±0.02     |
| IGSTAmt       | IGST           | ±0.02     |
| CGSTAmt       | CGST           | ±0.02     |
| SGSTAmt       | SGST           | ±0.02     |
| Net_Amt       | INVOICE AMOUNT | ±0.02     |

Current run result: **all 3 sheets validated OK, zero mismatches**.

---

## Current Grand Totals (for reference)

| Sheet    | Rows | TaxableAmt   | IGSTAmt   | CGSTAmt  | SGSTAmt  | Net_Amt      |
|----------|------|--------------|-----------|----------|----------|--------------|
| Flipkart | 31   | 1,82,004.88  | 8,264.63  | 418.26   | 418.26   | 1,91,106.03  |
| Meesho   | 26   | 80,077.27    | 3,503.20  | 250.38   | 250.38   | 84,081.23    |
| Snapdeal | 28   | 1,86,162.44  | 8,433.82  | 437.15   | 437.15   | 1,95,470.56  |
| **TOTAL**| **85**| **4,48,244.59** | **20,201.65** | **1,105.79** | **1,105.79** | **4,70,657.82** |
