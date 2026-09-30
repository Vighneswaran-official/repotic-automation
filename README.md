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
├── app.py                  # Flask web server and API endpoints
├── repotic_to_final.py      # Standalone CLI processing script
├── templates/
│   └── index.html          # Drag-and-drop web UI
├── output/                 # Output and upload staging folder (.gitkeep)
├── requirements.txt        # Python package dependencies
├── .gitignore              # Ignores runtime Excel files and cache
└── README.md               # Project documentation
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
python repotic_to_final.py
```
The output file will be written to `output/Final_Output_Tamilnadu.xlsx`.

---

## 📊 Data Mapping Overview

| REPOTIC Source Column | Output Column (Tally Template) |
|-----------------------|---------------------------------|
| STATE                 | StateOfSupply                   |
| HSN CODE              | HSNCode                         |
| RATE                  | TaxPer                          |
| QTY                   | Qty                             |
| TAXABLE VALUE         | TaxableAmt                      |
| IGST                  | IGSTAmt                         |
| CGST                  | CGSTAmt                         |
| SGST                  | SGSTAmt                         |
| INVOICE AMOUNT        | Net_Amt                         |
| *(Ledger Lookup)*     | Pty_Name                        |
| *(Ledger Lookup)*     | Sales Ledger                    |

All other template columns (`InvNo`, `Inv_Dt`, `Vch_Type`, `GSTIN`, `Product_Name`, `UOM`, etc.) are left blank according to Tally import specifications.

---

## 📄 License
MIT License
