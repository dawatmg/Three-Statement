# Three-Statement — FMP API Excel Exporter

Export **Income Statement**, **Balance Sheet**, and **Cash Flow Statement**
(annual *and* quarterly) for any publicly traded company to a professionally
formatted Excel workbook using the
[Financial Modeling Prep (FMP) API](https://financialmodelingprep.com/).

---

## Features

- Fetches all three financial statements via the FMP REST API
- Exports annual **and** quarterly data to a single `.xlsx` workbook (6 sheets)
- Professional Excel formatting:
  - Styled header rows (dark blue background, white text)
  - Section headers (light blue background)
  - Subtotal rows (bold font, top border)
  - Alternating row shading for readability
  - Number formatting with thousands separators
  - Auto-sized columns
- Retry logic with exponential backoff (via [tenacity](https://tenacity.readthedocs.io/))
- Graceful API error handling and logging
- Simple command-line interface

---

## Project Structure

```
Three-Statement/
├── three_statements_export.py   # Main script — CLI entry point
├── utils.py                     # Excel formatting helpers
├── config.py                    # Configuration & constants
├── example_usage.py             # Programmatic usage example
├── requirements.txt             # Python dependencies
├── .gitignore
└── README.md
```

---

## Requirements

- Python 3.10 or higher
- A free or paid [FMP API key](https://financialmodelingprep.com/developer/docs)

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/dawatmg/Three-Statement.git
cd Three-Statement

# 2. Create and activate a virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Configuration

Set your FMP API key as an environment variable:

```bash
export FMP_API_KEY=your_api_key_here
```

Or pass it directly via the `--api-key` flag (see below).

---

## Usage

### Command-line

```bash
# Basic usage — uses FMP_API_KEY environment variable
python three_statements_export.py --ticker AAPL

# Specify API key explicitly
python three_statements_export.py --ticker AAPL --api-key YOUR_KEY

# Customise output directory and number of periods
python three_statements_export.py \
    --ticker MSFT \
    --annual-limit 10 \
    --quarterly-limit 16 \
    --output ./output
```

The generated file will be saved as `<TICKER>_three_statements.xlsx` in the
output directory (default: current directory).

### Programmatic

```python
from three_statements_export import export_three_statements

output_path = export_three_statements(
    ticker="AAPL",
    api_key="YOUR_API_KEY",
    annual_limit=5,
    quarterly_limit=8,
    output_dir="./output",
)
print(f"Saved to: {output_path}")
```

See `example_usage.py` for a working example.

---

## Output

The workbook contains six sheets:

| Sheet | Description |
|---|---|
| Annual Income Statement | Last N fiscal years |
| Annual Balance Sheet | Last N fiscal years |
| Annual Cash Flow | Last N fiscal years |
| Quarterly Income Statement | Last N quarters |
| Quarterly Balance Sheet | Last N quarters |
| Quarterly Cash Flow | Last N quarters |

---

## CLI Options

| Option | Default | Description |
|---|---|---|
| `--ticker` | *(required)* | Stock ticker symbol (e.g. `AAPL`) |
| `--api-key` | `$FMP_API_KEY` | FMP API key |
| `--annual-limit` | `5` | Number of annual periods to fetch |
| `--quarterly-limit` | `12` | Number of quarterly periods to fetch |
| `--output` | `.` | Output directory for the Excel file |

---

## Dependencies

| Package | Purpose |
|---|---|
| `requests` | HTTP calls to the FMP API |
| `openpyxl` | Read/write Excel `.xlsx` files |
| `tenacity` | Retry logic with exponential backoff |
| `pandas` | Data manipulation and analysis (available for custom extensions) |

---

## License

MIT
