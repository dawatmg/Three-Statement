# Three-Statement Model Export

Automated three-statement financial export tool using Financial Modeling Prep (FMP) API. Fetches the latest income statement, balance sheet, and cash flow statement for any US stock ticker and exports to Excel with multiple sheets.

## Features

✅ **Fast & Reliable**: Fetches data in ~0.3 seconds with retry logic  
✅ **Comprehensive Sheets**: Latest data + 5-year/5-quarter historical data  
✅ **Auto-formatting**: Column widths auto-sized for readability  
✅ **Summary Sheet**: Key metrics consolidated from all three statements  
✅ **Both Annual & Quarterly**: Toggle between periods easily  

## Setup

1. Create virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
```

2. Install dependencies (already installed in `.venv`):

```bash
pip install pandas requests tenacity openpyxl
```

## Usage

### Run with ticker:

```bash
python three_statements_export.py AAPL
```

### Run interactively:

```bash
python three_statements_export.py
```

Then enter the ticker when prompted.

## Output

For each run, a single Excel file is generated: `<TICKER>_three_statements.xlsx`

The workbook contains 13 sheets:

**Latest Data (1 row each):**
- Annual Income Latest
- Annual Balance Latest  
- Annual CashFlow Latest
- Quarterly Income Latest
- Quarterly Balance Latest
- Quarterly CashFlow Latest

**Summary Sheet:**
- Key metrics from all six statements in one place

**Historical Data:**
- Annual Income 5Y (5 years of annual data)
- Annual Balance 5Y  
- Annual CashFlow 5Y
- Quarterly Income 5Q (5 quarters)
- Quarterly Balance 5Q
- Quarterly CashFlow 5Q

## Example

```bash
$ python three_statements_export.py TSLA
Saved workbook: /workspaces/Three-Statement/TSLA_three_statements.xlsx
Done in 0.35s
```

Open `TSLA_three_statements.xlsx` in Excel to view all statements and metrics.

## API Key

Your API key is embedded in the script. To change it, edit the `API_KEY` variable in `three_statements_export.py`.