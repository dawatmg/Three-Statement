"""
Three Statements Export — main script.

Fetches Income Statement, Balance Sheet, and Cash Flow data from the
Financial Modeling Prep (FMP) API and exports them to a professionally
formatted Excel workbook.

Usage
-----
    python three_statements_export.py --ticker AAPL
    python three_statements_export.py --ticker AAPL --output ./output
    python three_statements_export.py --ticker AAPL --api-key <YOUR_KEY>
"""

import argparse
import logging
import os
import sys
from typing import Optional

import requests
from openpyxl import Workbook
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

from config import (
    DEFAULT_ANNUAL_LIMIT,
    DEFAULT_QUARTERLY_LIMIT,
    FMP_API_KEY,
    FMP_BASE_URL,
    OUTPUT_DIR,
    RETRY_ATTEMPTS,
    RETRY_WAIT_MAX,
    RETRY_WAIT_MIN,
)
from utils import (
    freeze_header,
    set_column_widths,
    write_data_row,
    write_header_row,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# API helpers
# ---------------------------------------------------------------------------

class FMPAPIError(Exception):
    """Raised when the FMP API returns an error response."""


@retry(
    retry=retry_if_exception_type(requests.exceptions.RequestException),
    stop=stop_after_attempt(RETRY_ATTEMPTS),
    wait=wait_exponential(multiplier=1, min=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
    reraise=True,
)
def fetch(endpoint: str, params: dict) -> list[dict]:
    """
    Fetch data from the FMP API with retry logic and error handling.

    Parameters
    ----------
    endpoint : str
        API endpoint path (e.g. ``"/income-statement/AAPL"``).
    params : dict
        Query parameters (API key, period, limit, etc.).

    Returns
    -------
    list[dict]
        Parsed JSON response as a list of period dictionaries.

    Raises
    ------
    FMPAPIError
        When the API returns a non-200 status code or an error payload.
    """
    url = f"{FMP_BASE_URL}{endpoint}"
    logger.debug("GET %s  params=%s", url, {k: v for k, v in params.items() if k != "apikey"})

    response = requests.get(url, params=params, timeout=30)

    if response.status_code == 429:
        raise requests.exceptions.RequestException("Rate limit exceeded (HTTP 429)")

    if not response.ok:
        raise FMPAPIError(
            f"HTTP {response.status_code} for {endpoint}: {response.text[:200]}"
        )

    data = response.json()

    if isinstance(data, dict) and "Error Message" in data:
        raise FMPAPIError(data["Error Message"])

    if not isinstance(data, list) or len(data) == 0:
        logger.warning("Empty response for %s", endpoint)
        return []

    return data


def _fetch_statement(
    ticker: str,
    statement: str,
    period: str,
    limit: int,
    api_key: str,
) -> list[dict]:
    """
    Wrapper around :func:`fetch` for a specific financial statement.

    Parameters
    ----------
    statement : str
        One of ``"income-statement"``, ``"balance-sheet-statement"``,
        ``"cash-flow-statement"``.
    period : str
        ``"annual"`` or ``"quarter"``.
    """
    params = {
        "period": period,
        "limit": limit,
        "apikey": api_key,
    }
    try:
        return fetch(f"/{statement}/{ticker.upper()}", params)
    except FMPAPIError as exc:
        logger.error("FMP API error fetching %s (%s): %s", statement, period, exc)
        return []
    except requests.exceptions.RequestException as exc:
        logger.error("Network error fetching %s (%s): %s", statement, period, exc)
        return []


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def _period_headers(records: list[dict]) -> list[str]:
    """Extract period labels (date strings) from a list of statement records."""
    return [r.get("date", f"Period {i + 1}") for i, r in enumerate(records)]


def build_income_statement_sheet(ws, records: list[dict], period_label: str) -> None:
    """Populate a worksheet with an income statement."""
    if not records:
        ws.cell(row=1, column=1, value="No data available.")
        return

    headers = ["Income Statement"] + _period_headers(records)
    write_header_row(ws, headers)
    freeze_header(ws)
    set_column_widths(ws, len(records))

    def _vals(key: str) -> list:
        return [r.get(key) for r in records]

    row = 2

    # Revenue section
    write_data_row(ws, row, "Revenue", [], style="section"); row += 1
    write_data_row(ws, row, "Revenue", _vals("revenue")); row += 1
    write_data_row(ws, row, "Cost of Revenue", _vals("costOfRevenue")); row += 1
    write_data_row(ws, row, "Gross Profit", _vals("grossProfit"), style="subtotal"); row += 1

    # Operating expenses
    write_data_row(ws, row, "Operating Expenses", [], style="section"); row += 1
    write_data_row(ws, row, "Research & Development", _vals("researchAndDevelopmentExpenses")); row += 1
    write_data_row(ws, row, "Selling, General & Admin", _vals("sellingGeneralAndAdministrativeExpenses")); row += 1
    write_data_row(ws, row, "Operating Expenses (Total)", _vals("operatingExpenses")); row += 1
    write_data_row(ws, row, "Operating Income (EBIT)", _vals("operatingIncome"), style="subtotal"); row += 1

    # Below-the-line items
    write_data_row(ws, row, "Other Income / Expenses", [], style="section"); row += 1
    write_data_row(ws, row, "Interest Expense", _vals("interestExpense")); row += 1
    write_data_row(ws, row, "Total Other Income", _vals("totalOtherIncomeExpensesNet")); row += 1
    write_data_row(ws, row, "EBITDA", _vals("ebitda"), style="subtotal"); row += 1

    # Net income
    write_data_row(ws, row, "Pre-tax Income", [], style="section"); row += 1
    write_data_row(ws, row, "Income Before Tax", _vals("incomeBeforeTax")); row += 1
    write_data_row(ws, row, "Income Tax Expense", _vals("incomeTaxExpense")); row += 1
    write_data_row(ws, row, "Net Income", _vals("netIncome"), style="subtotal"); row += 1

    # Per-share data
    write_data_row(ws, row, "Per Share Data", [], style="section"); row += 1
    write_data_row(ws, row, "EPS (Basic)", _vals("eps")); row += 1
    write_data_row(ws, row, "EPS (Diluted)", _vals("epsdiluted")); row += 1
    write_data_row(ws, row, "Shares Outstanding (Basic)", _vals("weightedAverageShsOut")); row += 1
    write_data_row(ws, row, "Shares Outstanding (Diluted)", _vals("weightedAverageShsOutDil")); row += 1

    ws.title = f"{period_label} Income Statement"


def build_balance_sheet_sheet(ws, records: list[dict], period_label: str) -> None:
    """Populate a worksheet with a balance sheet."""
    if not records:
        ws.cell(row=1, column=1, value="No data available.")
        return

    headers = ["Balance Sheet"] + _period_headers(records)
    write_header_row(ws, headers)
    freeze_header(ws)
    set_column_widths(ws, len(records))

    def _vals(key: str) -> list:
        return [r.get(key) for r in records]

    row = 2

    # Current assets
    write_data_row(ws, row, "Current Assets", [], style="section"); row += 1
    write_data_row(ws, row, "Cash & Equivalents", _vals("cashAndCashEquivalents")); row += 1
    write_data_row(ws, row, "Short-term Investments", _vals("shortTermInvestments")); row += 1
    write_data_row(ws, row, "Net Receivables", _vals("netReceivables")); row += 1
    write_data_row(ws, row, "Inventory", _vals("inventory")); row += 1
    write_data_row(ws, row, "Other Current Assets", _vals("otherCurrentAssets")); row += 1
    write_data_row(ws, row, "Total Current Assets", _vals("totalCurrentAssets"), style="subtotal"); row += 1

    # Non-current assets
    write_data_row(ws, row, "Non-Current Assets", [], style="section"); row += 1
    write_data_row(ws, row, "Property, Plant & Equipment (net)", _vals("propertyPlantEquipmentNet")); row += 1
    write_data_row(ws, row, "Goodwill", _vals("goodwill")); row += 1
    write_data_row(ws, row, "Intangible Assets", _vals("intangibleAssets")); row += 1
    write_data_row(ws, row, "Long-term Investments", _vals("longTermInvestments")); row += 1
    write_data_row(ws, row, "Other Non-Current Assets", _vals("otherNonCurrentAssets")); row += 1
    write_data_row(ws, row, "Total Non-Current Assets", _vals("totalNonCurrentAssets"), style="subtotal"); row += 1
    write_data_row(ws, row, "Total Assets", _vals("totalAssets"), style="subtotal"); row += 1

    # Current liabilities
    write_data_row(ws, row, "Current Liabilities", [], style="section"); row += 1
    write_data_row(ws, row, "Accounts Payable", _vals("accountPayables")); row += 1
    write_data_row(ws, row, "Short-term Debt", _vals("shortTermDebt")); row += 1
    write_data_row(ws, row, "Tax Payable", _vals("taxPayables")); row += 1
    write_data_row(ws, row, "Deferred Revenue", _vals("deferredRevenue")); row += 1
    write_data_row(ws, row, "Other Current Liabilities", _vals("otherCurrentLiabilities")); row += 1
    write_data_row(ws, row, "Total Current Liabilities", _vals("totalCurrentLiabilities"), style="subtotal"); row += 1

    # Non-current liabilities
    write_data_row(ws, row, "Non-Current Liabilities", [], style="section"); row += 1
    write_data_row(ws, row, "Long-term Debt", _vals("longTermDebt")); row += 1
    write_data_row(ws, row, "Deferred Tax Liabilities", _vals("deferredTaxLiabilitiesNonCurrent")); row += 1
    write_data_row(ws, row, "Other Non-Current Liabilities", _vals("otherNonCurrentLiabilities")); row += 1
    write_data_row(ws, row, "Total Non-Current Liabilities", _vals("totalNonCurrentLiabilities"), style="subtotal"); row += 1
    write_data_row(ws, row, "Total Liabilities", _vals("totalLiabilities"), style="subtotal"); row += 1

    # Equity
    write_data_row(ws, row, "Shareholders' Equity", [], style="section"); row += 1
    write_data_row(ws, row, "Common Stock", _vals("commonStock")); row += 1
    write_data_row(ws, row, "Retained Earnings", _vals("retainedEarnings")); row += 1
    write_data_row(ws, row, "Other Equity", _vals("othertotalStockholdersEquity")); row += 1
    write_data_row(ws, row, "Total Shareholders' Equity", _vals("totalStockholdersEquity"), style="subtotal"); row += 1
    write_data_row(ws, row, "Total Liabilities & Equity", _vals("totalLiabilitiesAndStockholdersEquity"), style="subtotal"); row += 1

    ws.title = f"{period_label} Balance Sheet"


def build_cash_flow_sheet(ws, records: list[dict], period_label: str) -> None:
    """Populate a worksheet with a cash flow statement."""
    if not records:
        ws.cell(row=1, column=1, value="No data available.")
        return

    headers = ["Cash Flow Statement"] + _period_headers(records)
    write_header_row(ws, headers)
    freeze_header(ws)
    set_column_widths(ws, len(records))

    def _vals(key: str) -> list:
        return [r.get(key) for r in records]

    row = 2

    # Operating activities
    write_data_row(ws, row, "Operating Activities", [], style="section"); row += 1
    write_data_row(ws, row, "Net Income", _vals("netIncome")); row += 1
    write_data_row(ws, row, "Depreciation & Amortisation", _vals("depreciationAndAmortization")); row += 1
    write_data_row(ws, row, "Stock-based Compensation", _vals("stockBasedCompensation")); row += 1
    write_data_row(ws, row, "Changes in Working Capital", _vals("changeInWorkingCapital")); row += 1
    write_data_row(ws, row, "Accounts Receivable Change", _vals("accountsReceivables")); row += 1
    write_data_row(ws, row, "Inventory Change", _vals("inventory")); row += 1
    write_data_row(ws, row, "Accounts Payable Change", _vals("accountsPayables")); row += 1
    write_data_row(ws, row, "Other Operating Activities", _vals("otherWorkingCapital")); row += 1
    write_data_row(ws, row, "Cash from Operations", _vals("netCashProvidedByOperatingActivities"), style="subtotal"); row += 1

    # Investing activities
    write_data_row(ws, row, "Investing Activities", [], style="section"); row += 1
    write_data_row(ws, row, "Capital Expenditures", _vals("capitalExpenditure")); row += 1
    write_data_row(ws, row, "Acquisitions", _vals("acquisitionsNet")); row += 1
    write_data_row(ws, row, "Purchases of Investments", _vals("purchasesOfInvestments")); row += 1
    write_data_row(ws, row, "Sales of Investments", _vals("salesMaturitiesOfInvestments")); row += 1
    write_data_row(ws, row, "Other Investing Activities", _vals("otherInvestingActivites")); row += 1
    write_data_row(ws, row, "Cash from Investing", _vals("netCashUsedForInvestingActivites"), style="subtotal"); row += 1

    # Financing activities
    write_data_row(ws, row, "Financing Activities", [], style="section"); row += 1
    write_data_row(ws, row, "Debt Repayment", _vals("debtRepayment")); row += 1
    write_data_row(ws, row, "Common Stock Issuance", _vals("commonStockIssued")); row += 1
    write_data_row(ws, row, "Common Stock Repurchased", _vals("commonStockRepurchased")); row += 1
    write_data_row(ws, row, "Dividends Paid", _vals("dividendsPaid")); row += 1
    write_data_row(ws, row, "Other Financing Activities", _vals("otherFinancingActivites")); row += 1
    write_data_row(ws, row, "Cash from Financing", _vals("netCashUsedProvidedByFinancingActivities"), style="subtotal"); row += 1

    # Summary
    write_data_row(ws, row, "Summary", [], style="section"); row += 1
    write_data_row(ws, row, "Net Change in Cash", _vals("netChangeInCash"), style="subtotal"); row += 1
    write_data_row(ws, row, "Cash at Beginning of Period", _vals("cashAtBeginningOfPeriod")); row += 1
    write_data_row(ws, row, "Cash at End of Period", _vals("cashAtEndOfPeriod"), style="subtotal"); row += 1
    write_data_row(ws, row, "Free Cash Flow", _vals("freeCashFlow"), style="subtotal"); row += 1

    ws.title = f"{period_label} Cash Flow"


# ---------------------------------------------------------------------------
# Main export orchestrator
# ---------------------------------------------------------------------------

def export_three_statements(
    ticker: str,
    api_key: str,
    annual_limit: int = DEFAULT_ANNUAL_LIMIT,
    quarterly_limit: int = DEFAULT_QUARTERLY_LIMIT,
    output_dir: str = OUTPUT_DIR,
) -> str:
    """
    Fetch all three financial statements (annual + quarterly) and write them
    to a single Excel workbook.

    Parameters
    ----------
    ticker : str
        Stock ticker symbol (e.g. ``"AAPL"``).
    api_key : str
        FMP API key.
    annual_limit : int
        Number of annual periods to retrieve.
    quarterly_limit : int
        Number of quarterly periods to retrieve.
    output_dir : str
        Directory where the output ``.xlsx`` file is saved.

    Returns
    -------
    str
        Absolute path to the generated Excel file.
    """
    ticker = ticker.upper()
    logger.info("Fetching financial statements for %s …", ticker)

    # --- Fetch data ---------------------------------------------------------
    statements = {
        "income-statement": ("income-statement", "Income Statement"),
        "balance-sheet-statement": ("balance-sheet-statement", "Balance Sheet"),
        "cash-flow-statement": ("cash-flow-statement", "Cash Flow"),
    }

    annual_data: dict[str, list] = {}
    quarterly_data: dict[str, list] = {}

    for key, (endpoint, label) in statements.items():
        logger.info("  → %s (annual)", label)
        annual_data[key] = _fetch_statement(ticker, endpoint, "annual", annual_limit, api_key)

        logger.info("  → %s (quarterly)", label)
        quarterly_data[key] = _fetch_statement(ticker, endpoint, "quarter", quarterly_limit, api_key)

    # --- Build workbook -----------------------------------------------------
    wb = Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    sheet_configs = [
        ("Annual",    "annual",    annual_data),
        ("Quarterly", "quarterly", quarterly_data),
    ]

    for period_label, _, data in sheet_configs:
        build_income_statement_sheet(wb.create_sheet(), data["income-statement"], period_label)
        build_balance_sheet_sheet(wb.create_sheet(), data["balance-sheet-statement"], period_label)
        build_cash_flow_sheet(wb.create_sheet(), data["cash-flow-statement"], period_label)

    # --- Save ---------------------------------------------------------------
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{ticker}_three_statements.xlsx")
    wb.save(output_path)
    logger.info("Saved: %s", os.path.abspath(output_path))
    return os.path.abspath(output_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export FMP three financial statements to Excel.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--ticker",
        required=True,
        help="Stock ticker symbol, e.g. AAPL",
    )
    parser.add_argument(
        "--api-key",
        default=FMP_API_KEY,
        help="FMP API key. Falls back to FMP_API_KEY environment variable.",
    )
    parser.add_argument(
        "--annual-limit",
        type=int,
        default=DEFAULT_ANNUAL_LIMIT,
        help="Number of annual periods to fetch.",
    )
    parser.add_argument(
        "--quarterly-limit",
        type=int,
        default=DEFAULT_QUARTERLY_LIMIT,
        help="Number of quarterly periods to fetch.",
    )
    parser.add_argument(
        "--output",
        default=OUTPUT_DIR,
        help="Output directory for the Excel file.",
    )
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> None:
    args = _parse_args(argv)

    if not args.api_key:
        logger.error(
            "No API key provided. Set the FMP_API_KEY environment variable "
            "or pass --api-key <key>."
        )
        sys.exit(1)

    try:
        output_path = export_three_statements(
            ticker=args.ticker,
            api_key=args.api_key,
            annual_limit=args.annual_limit,
            quarterly_limit=args.quarterly_limit,
            output_dir=args.output,
        )
        print(f"\n✅  Done! File saved to: {output_path}")
    except Exception as exc:
        logger.error("Unexpected error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
