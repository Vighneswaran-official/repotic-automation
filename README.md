# REPOTIC Automation

An automated data processing pipeline and Flask web application that converts multi-channel e-commerce sales reports (**Flipkart**, **Meesho**, **Snapdeal**) into a standardized **26-column Tally import format** (`Final_Output_Tamilnadu.xlsx`).

---

## 🚀 Features

- **Automated Multi-Channel Parsing**: Extracts state-wise sales data from Flipkart, Meesho, and Snapdeal sheets while ignoring summary-only sections.
- **Smart Ledger Mapping**: Lookups state-level GST ledgers and party names dynamically with intelligent fallbacks for unmapped states (e.g. Tripura, Dadra & Nagar Haveli).
- **Tally Template Compliance**: Generates formatted, audit-ready Excel workbooks matching Tally's 26-column template schema.
- **Validation Engine**: Computes and validates column totals (Taxable Value, IGST, CGST, SGST, Invoice Amount) against source values within a ±0.02 tolerance to ensure 100% data integrity.
- **Interactive Web App**: Modern drag-and-drop web UI powered by Flask for non-technical users to upload source files, review processing summaries, inspect validation status, and download the output.
- **Standalone CLI Mode**: Scriptable execution via `repotic_to_final.py` for headless or batch automation.

---

## 📁 Repository Structure

```
repotic-automation/
├── app.py                      # Production and serverless entry point
├── requirements.txt            # Python dependencies
├── README.md                   # Project overview & documentation
├── .gitignore                  # Git exclusions (runtime xlsx & cache)
│
├── src/                        # Core application code
│   ├── config.py               # Constants, column mappings, and writable temp dir
│   ├── services/
│   │   ├── __init__.py
│   │   └── processor.py        # Excel extraction, ledger lookup & validation pipeline
│   └── web/
│       ├── __init__.py
│       └── routes.py           # Flask routes (/, /run, /download)
│
├── static/                     # Frontend static assets
│   ├── css/
│   │   └── style.css           # Styling, themes, responsive layout
│   └── js/
│       └── main.js             # Drag-and-drop, execution AJAX, base64 blob download
│
├── templates/                  # Server-rendered HTML templates
│   └── index.html              # Clean semantic template
│
├── scripts/                    # Headless CLI & batch scripts
│   ├── __init__.py
│   └── repotic_to_final.py     # Standalone CLI processing runner
│
├── docs/                       # Developer documentation & references
│   └── CLAUDE.md
│
└── output/                     # Generated files staging directory (.gitkeep)
```

---

## 🛠️ Installation & Setup

### Prerequisites
- Python 3.10+
- Git

### 1. Clone Repository
```bash
git clone https://github.com/Vighneswaran-official/repotic-automation.git
cd repotic-automation
```

### 2. Set Up Virtual Environment (Recommended)
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux / macOS
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## 💻 Usage

### Option A: Web Application (Browser UI)
1. Run the Flask server:
   ```bash
   python app.py
   ```
2. Open your browser and navigate to:
   ```
   http://127.0.0.1:5000
   ```
3. Upload the 3 required Excel files:
   - **REPOTIC File** (`REPOTIC.xlsx` with Flipkart, Meesho, Snapdeal sheets)
   - **Ledger File** (`AMAZON_LEDGER FILE.xlsx`)
   - **Template File** (`Final Sample Format Tamilnadu.xlsx`)
4. Click **Run Automation**. Once validated, download the generated `Final_Output_Tamilnadu.xlsx`.

### Option B: Command Line (CLI)
Place your files in `output/` as:
- `output/upload_repotic.xlsx`
- `output/upload_ledger.xlsx`
- `output/upload_template.xlsx`

Run:
```bash
python scripts/repotic_to_final.py
# Or override month:
python scripts/repotic_to_final.py --month 2027-04
# Or supply custom paths:
python scripts/repotic_to_final.py path/to/repotic.xlsx path/to/ledger.xlsx path/to/template.xlsx path/to/output.xlsx --month 2026-08
```
The output file will be written to `output/Final_Output_Tamilnadu.xlsx`.

---

## 📊 Data Mapping & Invoice Numbering

| REPOTIC Source Column | Output Column (Tally Template) | Description / Logic |
|-----------------------|---------------------------------|---------------------|
| *(Auto-generated)*    | InvNo                           | `MM/YY-YY/XX-NN` (e.g. `08/26-27/FL-01`, text format) |
| *(Auto-generated)*    | Inv_Dt                          | Month-end date (e.g. `31-08-2026`) |
| STATE                 | StateOfSupply                   | Exact state name from sales sheet |
| HSN CODE              | HSNCode                         | Mapped HSN code |
| RATE                  | TaxPer                          | Applicable GST rate percentage |
| QTY                   | Qty                             | Quantity sold |
| TAXABLE VALUE         | TaxableAmt                      | Taxable sales value |
| IGST                  | IGSTAmt                         | Integrated GST amount |
| CGST                  | CGSTAmt                         | Central GST amount |
| SGST                  | SGSTAmt                         | State GST amount |
| INVOICE AMOUNT        | Net_Amt                         | Total invoice amount |
| *(Ledger Lookup)*     | Pty_Name                        | Looked up party name |
| *(Ledger Lookup)*     | Sales Ledger                    | Looked up sales ledger |

All other template columns (`GSTIN`, `Product_Name`, `UOM`, etc.) are left blank according to Tally import specifications.

---

## 📄 License
MIT License
