import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.worksheet import Worksheet
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type


BASE_URL = "https://financialmodelingprep.com/stable"
API_KEY = "QXYMQR8PKhe0CtSUlNERXsZd255vgL0g".strip()

if not API_KEY:
    raise RuntimeError("Missing FMP API key.")

SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": "three-statements-export/1.0",
        "Accept": "application/json",
    }
)


class FMPError(Exception):
    pass


@retry(
    stop=stop_after_attempt(4),
    wait=wait_exponential(multiplier=1, min=1, max=8),
    retry=retry_if_exception_type((requests.RequestException, FMPError)),
)
def fetch(endpoint: str, params: Dict[str, Any]) -> List[Dict[str, Any]]:
    query = dict(params)
    query["apikey"] = API_KEY

    url = f"{BASE_URL}/{endpoint}"
    response = SESSION.get(url, params=query, timeout=30)

    if response.status_code == 429:
        raise FMPError("Rate limited by FMP (429). Try again later.")

    try:
        response.raise_for_status()
    except requests.HTTPError as e:
        raise FMPError(f"HTTP {response.status_code}: {response.text}")
    
    data = response.json()

    if isinstance(data, dict):
        if data.get("Error Message"):
            raise FMPError(data["Error Message"])
        if data.get("error"):
            raise FMPError(str(data["error"]))

    if not isinstance(data, list):
        raise FMPError(f"Unexpected response type from {endpoint}: {type(data).__name__}")

    if len(data) == 0:
        raise FMPError(f"No data returned from {endpoint}")

    return data

def to_df(records: List[Dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(records)
    if "date" in df.columns:
        try:
            df = df.sort_values("date", ascending=False)
        except Exception:
            pass
    return df


def latest_row(df: pd.DataFrame) -> pd.DataFrame:
    return df.head(1).copy()


def keep_columns_first(df: pd.DataFrame, preferred: List[str]) -> pd.DataFrame:
    cols = list(df.columns)
    ordered = [c for c in preferred if c in cols] + [c for c in cols if c not in preferred]
    return df[ordered]


def write_sheet(writer: pd.ExcelWriter, name: str, df: pd.DataFrame) -> None:
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    for idx, col in enumerate(df.columns, start=1):
        max_len = max(
            len(str(col)),
            *(len(str(x)) for x in df[col].head(100).tolist())
        )
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 40)


def build_summary(
    ticker: str,
    inc_a: pd.DataFrame,
    bal_a: pd.DataFrame,
    cf_a: pd.DataFrame,
    inc_q: pd.DataFrame,
    bal_q: pd.DataFrame,
    cf_q: pd.DataFrame,
) -> pd.DataFrame:
    rows = []

    def add(section: str, df: pd.DataFrame, fields: List[str]) -> None:
        if df.empty:
            return
        row = df.iloc[0].to_dict()
        for f in fields:
            rows.append(
                {
                    "ticker": ticker,
                    "section": section,
                    "date": row.get("date"),
                    "period": row.get("period"),
                    "field": f,
                    "value": row.get(f),
                }
            )

    add("Annual Income", inc_a, ["revenue", "grossProfit", "operatingIncome", "netIncome", "eps", "weightedAverageShsOut"])
    add("Annual Balance", bal_a, ["cashAndCashEquivalents", "totalAssets", "totalDebt", "totalLiabilities", "totalStockholdersEquity"])
    add("Annual Cash Flow", cf_a, ["operatingCashFlow", "capitalExpenditure", "freeCashFlow", "netCashProvidedByOperatingActivities"])
    add("Quarterly Income", inc_q, ["revenue", "grossProfit", "operatingIncome", "netIncome", "eps", "weightedAverageShsOut"])
    add("Quarterly Balance", bal_q, ["cashAndCashEquivalents", "totalAssets", "totalDebt", "totalLiabilities", "totalStockholdersEquity"])
    add("Quarterly Cash Flow", cf_q, ["operatingCashFlow", "capitalExpenditure", "freeCashFlow", "netCashProvidedByOperatingActivities"])

    return pd.DataFrame(rows)


def fmt_num(val: Any, decimals: int = 0) -> str:
    """Format number for display (in millions)."""
    if val is None or pd.isna(val):
        return "-"
    try:
        num = float(val)
        if abs(num) >= 1_000_000:
            return f"${num/1_000_000:.{decimals}f}M"
        elif abs(num) >= 1_000:
            return f"${num/1_000:.{decimals}f}K"
        else:
            return f"${num:.{decimals}f}"
    except:
        return "-"


def style_header_row(ws: Worksheet, row: int, col_count: int) -> None:
    """Style a header row."""
    header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_section_row(ws: Worksheet, row: int, col_count: int) -> None:
    """Style a section header row."""
    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    section_font = Font(bold=True, size=10)
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = section_fill
        cell.font = section_font


def style_subtotal_row(ws: Worksheet, row: int, col_count: int) -> None:
    """Style a subtotal row."""
    subtotal_font = Font(bold=True)
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = subtotal_font
        cell.border = Border(top=Side(style='thin'))


def build_income_statement_sheet(ws: Worksheet, ticker: str, data_list: List[Dict[str, Any]]) -> None:
    """Create professionally formatted income statement."""
    ws.title = "Income Statement"
    
    # Header
    ws['A1'] = f"{ticker} - Consolidated Income Statement"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:E1')
    
    # Column headers
    dates = [row.get("date", "") for row in data_list]
    ws['A3'] = "Description"
    for idx, date in enumerate(dates):
        ws.cell(row=3, column=idx+2).value = date
    style_header_row(ws, 3, len(dates) + 1)
    
    row = 4
    
    # Revenue section
    ws[f'A{row}'] = "Revenue"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    for rec in data_list:
        ws.cell(row=row, column=1).value = f"  Revenue"
        for idx, data_rec in enumerate(data_list):
            val = data_rec.get("revenue")
            ws.cell(row=row, column=idx+2).value = val
            ws.cell(row=row, column=idx+2).number_format = '#,##0'
        break
    row += 1
    
    ws[f'A{row}'] = "  Cost of Revenue"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("costOfRevenue")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "Gross Profit"
    style_subtotal_row(ws, row, len(dates) + 1)
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("grossProfit")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    # Operating Expenses
    ws[f'A{row}'] = "Operating Expenses"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    expense_items = [
        ("researchAndDevelopmentExpenses", "Research & Development"),
        ("sellingAndMarketingExpenses", "Sales & Marketing"),
        ("generalAndAdministrativeExpenses", "General & Administrative"),
    ]
    
    for key, label in expense_items:
        ws[f'A{row}'] = f"  {label}"
        for idx, data_rec in enumerate(data_list):
            val = data_rec.get(key, 0)
            ws.cell(row=row, column=idx+2).value = val
            ws.cell(row=row, column=idx+2).number_format = '#,##0'
        row += 1
    
    ws[f'A{row}'] = "Total Operating Expenses"
    style_subtotal_row(ws, row, len(dates) + 1)
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("operatingExpenses")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    # Operating Income
    ws[f'A{row}'] = "Operating Income (EBIT)"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True, size=11)
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("operatingIncome")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    # Non-operating
    ws[f'A{row}'] = "Interest Income & Other"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Interest Income"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("interestIncome", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Other Income"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("nonOperatingIncomeExcludingInterest", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Income Before Tax"
    style_subtotal_row(ws, row, len(dates) + 1)
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("incomeBeforeTax")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Income Tax Expense"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("incomeTaxExpense")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Net Income"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True, size=12, color="FF0000")
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("netIncome")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
        ws.cell(row=row, column=idx+2).font = Font(bold=True, size=11)
    row += 2
    
    ws[f'A{row}'] = "Earnings Per Share (EPS)"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("eps")
        ws.cell(row=row, column=idx+2).number_format = '0.00'
    
    # Auto-size columns
    ws.column_dimensions['A'].width = 30
    for col in range(2, len(dates) + 2):
        ws.column_dimensions[chr(64+col)].width = 18


def build_balance_sheet_sheet(ws: Worksheet, ticker: str, data_list: List[Dict[str, Any]]) -> None:
    """Create professionally formatted balance sheet."""
    ws.title = "Balance Sheet"
    
    # Header
    ws['A1'] = f"{ticker} - Consolidated Balance Sheet"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:E1')
    
    # Column headers
    dates = [row.get("date", "") for row in data_list]
    ws['A3'] = "Description"
    for idx, date in enumerate(dates):
        ws.cell(row=3, column=idx+2).value = date
    style_header_row(ws, 3, len(dates) + 1)
    
    row = 4
    
    # ASSETS
    ws[f'A{row}'] = "ASSETS"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    row += 1
    
    ws[f'A{row}'] = "Current Assets"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Cash and Cash Equivalents"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("cashAndCashEquivalents", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Short-term Investments"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("shortTermInvestments", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Accounts Receivable"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("accountsReceivable", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Inventory"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("inventory", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Non-Current Assets"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Property, Plant & Equipment"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("propertyPlantEquipmentNet", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Intangibles & Goodwill"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("intangibleAssets", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "TOTAL ASSETS"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True, size=12, color="FF0000")
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("totalAssets")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
        ws.cell(row=row, column=idx+2).font = Font(bold=True)
    row += 3
    
    # LIABILITIES & EQUITY
    ws[f'A{row}'] = "LIABILITIES & EQUITY"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    row += 1
    
    ws[f'A{row}'] = "Current Liabilities"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Accounts Payable"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("accountsPayable", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Short-term Debt"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("shortTermDebt", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Long-term Liabilities"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Long-term Debt"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("longTermDebt", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "TOTAL LIABILITIES"
    style_subtotal_row(ws, row, len(dates) + 1)
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("totalLiabilities")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Stockholders' Equity"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Common Stock"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("commonStock", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Retained Earnings"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("retainedEarnings", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "TOTAL EQUITY"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True)
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("totalStockholdersEquity")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "TOTAL LIABILITIES & EQUITY"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True, size=12, color="FF0000")
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("totalLiabilities", 0) + data_rec.get("totalStockholdersEquity", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
        ws.cell(row=row, column=idx+2).font = Font(bold=True)
    
    # Auto-size columns
    ws.column_dimensions['A'].width = 30
    for col in range(2, len(dates) + 2):
        ws.column_dimensions[chr(64+col)].width = 18


def build_cash_flow_sheet(ws: Worksheet, ticker: str, data_list: List[Dict[str, Any]]) -> None:
    """Create professionally formatted cash flow statement."""
    ws.title = "Cash Flow"
    
    # Header
    ws['A1'] = f"{ticker} - Consolidated Cash Flow Statement"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:E1')
    
    # Column headers
    dates = [row.get("date", "") for row in data_list]
    ws['A3'] = "Description"
    for idx, date in enumerate(dates):
        ws.cell(row=3, column=idx+2).value = date
    style_header_row(ws, 3, len(dates) + 1)
    
    row = 4
    
    # Operating Activities
    ws[f'A{row}'] = "Cash Flow from Operating Activities"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Net Income"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("netIncome", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Depreciation & Amortization"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("depreciationAndAmortization", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Stock-Based Compensation"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("stockBasedCompensation", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Changes in Working Capital"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("changeInWorkingCapital", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Net Cash from Operations"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True)
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("netCashProvidedByOperatingActivities", data_rec.get("operatingCashFlow", 0))
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    # Investing Activities
    ws[f'A{row}'] = "Cash Flow from Investing Activities"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Capital Expenditures"
    for idx, data_rec in enumerate(data_list):
        capex = data_rec.get("capitalExpenditure", 0)
        ws.cell(row=row, column=idx+2).value = capex
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Acquisitions"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("acquisitions", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Net Cash from Investing"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True)
    for idx, data_rec in enumerate(data_list):
        investing = (data_rec.get("capitalExpenditure", 0) * -1) + data_rec.get("acquisitions", 0)
        ws.cell(row=row, column=idx+2).value = investing
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    # Financing Activities
    ws[f'A{row}'] = "Cash Flow from Financing Activities"
    style_section_row(ws, row, len(dates) + 1)
    row += 1
    
    ws[f'A{row}'] = "  Debt Issued"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("debtIssued", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Debt Repaid"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("debtRepaid", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "  Dividends Paid"
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("dividendsPaid", 0)
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Net Cash from Financing"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True)
    for idx, data_rec in enumerate(data_list):
        financing = data_rec.get("debtIssued", 0) - data_rec.get("debtRepaid", 0) - data_rec.get("dividendsPaid", 0)
        ws.cell(row=row, column=idx+2).value = financing
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
    row += 2
    
    ws[f'A{row}'] = "Free Cash Flow"
    style_subtotal_row(ws, row, len(dates) + 1)
    ws[f'A{row}'].font = Font(bold=True, size=12, color="FF0000")
    for idx, data_rec in enumerate(data_list):
        ws.cell(row=row, column=idx+2).value = data_rec.get("freeCashFlow")
        ws.cell(row=row, column=idx+2).number_format = '#,##0'
        ws.cell(row=row, column=idx+2).font = Font(bold=True)
    
    # Auto-size columns
    ws.column_dimensions['A'].width = 30
    for col in range(2, len(dates) + 2):
        ws.column_dimensions[chr(64+col)].width = 18


def export_three_statements(ticker: str) -> Path:
    ticker = ticker.upper().strip()

    annual_income = to_df(fetch("income-statement", {"symbol": ticker, "period": "annual", "limit": 5}))
    annual_balance = to_df(fetch("balance-sheet-statement", {"symbol": ticker, "period": "annual", "limit": 5}))
    annual_cash = to_df(fetch("cash-flow-statement", {"symbol": ticker, "period": "annual", "limit": 5}))

    quarterly_income = to_df(fetch("income-statement", {"symbol": ticker, "period": "quarter", "limit": 5}))
    quarterly_balance = to_df(fetch("balance-sheet-statement", {"symbol": ticker, "period": "quarter", "limit": 5}))
    quarterly_cash = to_df(fetch("cash-flow-statement", {"symbol": ticker, "period": "quarter", "limit": 5}))

    out = Path(f"{ticker}_three_statements.xlsx")

    # Create workbook with professional formatting
    wb = Workbook()
    wb.remove(wb.active)
    
    # Create sheets with professional formatting
    if not annual_income.empty:
        ws = wb.create_sheet("Annual Income")
        build_income_statement_sheet(ws, ticker, annual_income.to_dict('records'))
    
    if not annual_balance.empty:
        ws = wb.create_sheet("Annual Balance Sheet")
        build_balance_sheet_sheet(ws, ticker, annual_balance.to_dict('records'))
    
    if not annual_cash.empty:
        ws = wb.create_sheet("Annual Cash Flow")
        build_cash_flow_sheet(ws, ticker, annual_cash.to_dict('records'))
    
    if not quarterly_income.empty:
        ws = wb.create_sheet("Quarterly Income")
        build_income_statement_sheet(ws, ticker, quarterly_income.to_dict('records'))
    
    if not quarterly_balance.empty:
        ws = wb.create_sheet("Quarterly Balance Sheet")
        build_balance_sheet_sheet(ws, ticker, quarterly_balance.to_dict('records'))
    
    if not quarterly_cash.empty:
        ws = wb.create_sheet("Quarterly Cash Flow")
        build_cash_flow_sheet(ws, ticker, quarterly_cash.to_dict('records'))

    wb.save(out)
    return out


def main() -> None:
    if len(sys.argv) > 1:
        ticker = sys.argv[1]
    else:
        ticker = input("Enter ticker: ").strip()

    start = time.time()
    try:
        out = export_three_statements(ticker)
        elapsed = time.time() - start
        print(f"Saved workbook: {out.resolve()}")
        print(f"Done in {elapsed:.2f}s")
    except Exception as e:
        if hasattr(e, 'last_attempt'):
            print(f"ERROR: {e.last_attempt.exception()}")
        else:
            print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
