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
    SheetWriter,
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


def build_income_statement_sheet(ws, records: list[dict], period_label: str) -> None:
    """Populate a worksheet with an income statement."""
    if not records:
        ws.cell(row=1, column=1, value="No data available.")
        return

    sw = SheetWriter(ws, "Income Statement", records)

    sw.section("Revenue")
    sw.normal("Revenue", "revenue")
    sw.normal("Cost of Revenue", "costOfRevenue")
    sw.subtotal("Gross Profit", "grossProfit")

    sw.section("Operating Expenses")
    sw.normal("Research & Development", "researchAndDevelopmentExpenses")
    sw.normal("Selling, General & Admin", "sellingGeneralAndAdministrativeExpenses")
    sw.normal("Operating Expenses (Total)", "operatingExpenses")
    sw.subtotal("Operating Income (EBIT)", "operatingIncome")

    sw.section("Other Income / Expenses")
    sw.normal("Interest Expense", "interestExpense")
    sw.normal("Total Other Income", "totalOtherIncomeExpensesNet")
    sw.subtotal("EBITDA", "ebitda")

    sw.section("Pre-tax Income")
    sw.normal("Income Before Tax", "incomeBeforeTax")
    sw.normal("Income Tax Expense", "incomeTaxExpense")
    sw.subtotal("Net Income", "netIncome")

    sw.section("Per Share Data")
    sw.normal("EPS (Basic)", "eps")
    sw.normal("EPS (Diluted)", "epsdiluted")
    sw.normal("Shares Outstanding (Basic)", "weightedAverageShsOut")
    sw.normal("Shares Outstanding (Diluted)", "weightedAverageShsOutDil")

    ws.title = f"{period_label} Income Statement"


def build_balance_sheet_sheet(ws, records: list[dict], period_label: str) -> None:
    """Populate a worksheet with a balance sheet."""
    if not records:
        ws.cell(row=1, column=1, value="No data available.")
        return

    sw = SheetWriter(ws, "Balance Sheet", records)

    sw.section("Current Assets")
    sw.normal("Cash & Equivalents", "cashAndCashEquivalents")
    sw.normal("Short-term Investments", "shortTermInvestments")
    sw.normal("Net Receivables", "netReceivables")
    sw.normal("Inventory", "inventory")
    sw.normal("Other Current Assets", "otherCurrentAssets")
    sw.subtotal("Total Current Assets", "totalCurrentAssets")

    sw.section("Non-Current Assets")
    sw.normal("Property, Plant & Equipment (net)", "propertyPlantEquipmentNet")
    sw.normal("Goodwill", "goodwill")
    sw.normal("Intangible Assets", "intangibleAssets")
    sw.normal("Long-term Investments", "longTermInvestments")
    sw.normal("Other Non-Current Assets", "otherNonCurrentAssets")
    sw.subtotal("Total Non-Current Assets", "totalNonCurrentAssets")
    sw.subtotal("Total Assets", "totalAssets")

    sw.section("Current Liabilities")
    sw.normal("Accounts Payable", "accountPayables")
    sw.normal("Short-term Debt", "shortTermDebt")
    sw.normal("Tax Payable", "taxPayables")
    sw.normal("Deferred Revenue", "deferredRevenue")
    sw.normal("Other Current Liabilities", "otherCurrentLiabilities")
    sw.subtotal("Total Current Liabilities", "totalCurrentLiabilities")

    sw.section("Non-Current Liabilities")
    sw.normal("Long-term Debt", "longTermDebt")
    sw.normal("Deferred Tax Liabilities", "deferredTaxLiabilitiesNonCurrent")
    sw.normal("Other Non-Current Liabilities", "otherNonCurrentLiabilities")
    sw.subtotal("Total Non-Current Liabilities", "totalNonCurrentLiabilities")
    sw.subtotal("Total Liabilities", "totalLiabilities")

    sw.section("Shareholders' Equity")
    sw.normal("Common Stock", "commonStock")
    sw.normal("Retained Earnings", "retainedEarnings")
    sw.normal("Other Equity", "othertotalStockholdersEquity")
    sw.subtotal("Total Shareholders' Equity", "totalStockholdersEquity")
    sw.subtotal("Total Liabilities & Equity", "totalLiabilitiesAndStockholdersEquity")

    ws.title = f"{period_label} Balance Sheet"


def build_cash_flow_sheet(ws, records: list[dict], period_label: str) -> None:
    """Populate a worksheet with a cash flow statement."""
    if not records:
        ws.cell(row=1, column=1, value="No data available.")
        return

    sw = SheetWriter(ws, "Cash Flow Statement", records)

    sw.section("Operating Activities")
    sw.normal("Net Income", "netIncome")
    sw.normal("Depreciation & Amortisation", "depreciationAndAmortization")
    sw.normal("Stock-based Compensation", "stockBasedCompensation")
    sw.normal("Changes in Working Capital", "changeInWorkingCapital")
    sw.normal("Accounts Receivable Change", "accountsReceivables")
    sw.normal("Inventory Change", "inventory")
    sw.normal("Accounts Payable Change", "accountsPayables")
    sw.normal("Other Operating Activities", "otherWorkingCapital")
    sw.subtotal("Cash from Operations", "netCashProvidedByOperatingActivities")

    sw.section("Investing Activities")
    sw.normal("Capital Expenditures", "capitalExpenditure")
    sw.normal("Acquisitions", "acquisitionsNet")
    sw.normal("Purchases of Investments", "purchasesOfInvestments")
    sw.normal("Sales of Investments", "salesMaturitiesOfInvestments")
    sw.normal("Other Investing Activities", "otherInvestingActivites")
    sw.subtotal("Cash from Investing", "netCashUsedForInvestingActivites")

    sw.section("Financing Activities")
    sw.normal("Debt Repayment", "debtRepayment")
    sw.normal("Common Stock Issuance", "commonStockIssued")
    sw.normal("Common Stock Repurchased", "commonStockRepurchased")
    sw.normal("Dividends Paid", "dividendsPaid")
    sw.normal("Other Financing Activities", "otherFinancingActivites")
    sw.subtotal("Cash from Financing", "netCashUsedProvidedByFinancingActivities")

    sw.section("Summary")
    sw.subtotal("Net Change in Cash", "netChangeInCash")
    sw.normal("Cash at Beginning of Period", "cashAtBeginningOfPeriod")
    sw.subtotal("Cash at End of Period", "cashAtEndOfPeriod")
    sw.subtotal("Free Cash Flow", "freeCashFlow")

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
