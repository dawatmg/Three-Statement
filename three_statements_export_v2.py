import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

API_KEY = "QXYMQR8PKhe0CtSUlNERXsZd255vgL0g"
BASE_URL = "https://financialmodelingprep.com/stable"


class FMPError(Exception):
    pass


SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": "three-statement-model/2.0",
        "Accept": "application/json",
    }
)

# ---------- Flexible field aliases ----------
IS_ALIASES = {
    "Revenue": ["revenue", "totalRevenue"],
    "COGS": ["costOfRevenue", "costAndExpenses"],
    "Reported Gross Profit": ["grossProfit"],
    "R&D": ["researchAndDevelopmentExpenses"],
    "SG&A": ["sellingGeneralAndAdministrativeExpenses"],
    "Reported Operating Expenses": ["operatingExpenses"],
    "Reported Operating Income": ["operatingIncome", "ebit"],
    "Interest Expense": ["interestExpense"],
    "Pre-Tax Income": ["incomeBeforeTax", "incomeBeforeTaxExpense"],
    "Taxes": ["incomeTaxExpense"],
    "Reported Net Income": ["netIncome"],
    "Shares Outstanding": ["weightedAverageShsOut", "weightedAverageShsOutDil"],
    "Reported EPS": ["eps", "epsdiluted"],
}

BS_ALIASES = {
    "Cash & Equivalents": ["cashAndCashEquivalents"],
    "Short-Term Investments": ["shortTermInvestments"],
    "Receivables": ["netReceivables", "accountsReceivables"],
    "Inventory": ["inventory"],
    "Other Current Assets": ["otherCurrentAssets"],
    "Reported Total Current Assets": ["totalCurrentAssets"],
    "PP&E, Net": ["propertyPlantEquipmentNet"],
    "Goodwill": ["goodwill"],
    "Intangibles": ["intangibleAssets"],
    "Other Non-Current Assets": ["otherNonCurrentAssets", "totalNonCurrentAssets"],
    "Reported Total Assets": ["totalAssets"],
    "Accounts Payable": ["accountPayables"],
    "Short-Term Debt": ["shortTermDebt"],
    "Other Current Liabilities": ["otherCurrentLiabilities"],
    "Reported Total Current Liabilities": ["totalCurrentLiabilities"],
    "Long-Term Debt": ["longTermDebt"],
    "Other Non-Current Liabilities": ["otherNonCurrentLiabilities", "totalNonCurrentLiabilities"],
    "Reported Total Liabilities": ["totalLiabilities"],
    "Common Stock": ["commonStock"],
    "Retained Earnings": ["retainedEarnings"],
    "Reported Total Equity": ["totalStockholdersEquity", "totalEquity"],
    "Reported Total Debt": ["totalDebt"],
}

CF_ALIASES = {
    "Reported Net Income": ["netIncome"],
    "D&A": ["depreciationAndAmortization"],
    "Stock-Based Compensation": ["stockBasedCompensation"],
    "Change in Working Capital": ["changeInWorkingCapital"],
    "Reported CFO": ["netCashProvidedByOperatingActivities", "operatingCashFlow"],
    "Capex": ["capitalExpenditure"],
    "Acquisitions": ["acquisitionsNet"],
    "Reported CFI": ["netCashUsedForInvestingActivites"],
    "Debt Repayment": ["debtRepayment"],
    "Stock Issued": ["commonStockIssued"],
    "Stock Repurchased": ["commonStockRepurchased"],
    "Dividends Paid": ["dividendsPaid"],
    "Reported CFF": ["netCashUsedProvidedByFinancingActivities"],
    "Reported Free Cash Flow": ["freeCashFlow"],
}

THIN = Side(style="thin", color="D9D9D9")


@retry(
    stop=stop_after_attempt(4),
    wait=wait_exponential(multiplier=1, min=1, max=8),
    retry=retry_if_exception_type((requests.RequestException, FMPError)),
)
def fetch(endpoint: str, ticker: str, period: str, limit: int) -> pd.DataFrame:
    url = f"{BASE_URL}/{endpoint}"
    params = {
        "symbol": ticker,
        "period": period,
        "limit": limit,
        "apikey": API_KEY,
    }
    r = SESSION.get(url, params=params, timeout=30)

    if r.status_code == 429:
        raise FMPError("Rate limited by FMP (429).")

    r.raise_for_status()
    data = r.json()

    if isinstance(data, dict):
        if data.get("Error Message"):
            raise FMPError(data["Error Message"])
        if data.get("error"):
            raise FMPError(str(data["error"]))

    if not isinstance(data, list) or len(data) == 0:
        raise FMPError(f"No data returned for {endpoint} / {ticker} / {period}")

    df = pd.DataFrame(data)
    if "date" not in df.columns:
        raise FMPError(f"Missing date column in {endpoint}")
    df = df.sort_values("date", ascending=False).reset_index(drop=True)
    return df


def get_statements(ticker: str):
    inc = fetch("income-statement", ticker, "annual", 5)
    bal = fetch("balance-sheet-statement", ticker, "annual", 5)
    cf = fetch("cash-flow-statement", ticker, "annual", 5)
    return inc, bal, cf


def style_sheet(ws, header_color="1F1F1F", freeze="B2"):
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=header_color)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, min(ws.max_row, 100) + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            max_len = max(max_len, len(str(val)) if val is not None else 0)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 28)

    ws.freeze_panes = freeze


def write_df(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, index: bool = False):
    df.to_excel(writer, sheet_name=sheet_name, index=index)
    style_sheet(writer.sheets[sheet_name], header_color="1F4E78", freeze="A2" if not index else "B2")


def choose_periods(inc: pd.DataFrame, bal: pd.DataFrame, cf: pd.DataFrame) -> List[str]:
    # safest: use income periods and only keep dates present in all 3
    common = set(inc["date"]).intersection(set(bal["date"])).intersection(set(cf["date"]))
    periods = [d for d in inc["date"].tolist() if d in common]
    return periods[:5]


def choose_series(df: pd.DataFrame, aliases: List[str], periods: List[str]) -> Tuple[List[Optional[float]], str]:
    for alias in aliases:
        if alias in df.columns:
            by_date = dict(zip(df["date"], df[alias]))
            return [by_date.get(p) for p in periods], alias
    return [None] * len(periods), ""


def build_mapped_df(
    df: pd.DataFrame,
    aliases_dict: Dict[str, List[str]],
    periods: List[str],
    statement_name: str,
) -> Tuple[pd.DataFrame, List[Dict[str, str]]]:
    rows = []
    log = []

    for canonical, aliases in aliases_dict.items():
        values, source = choose_series(df, aliases, periods)
        rows.append([canonical] + values)
        log.append(
            {
                "statement": statement_name,
                "canonical_line": canonical,
                "source_field_used": source if source else "MISSING",
                "status": "mapped" if source else "missing",
            }
        )

    mapped = pd.DataFrame(rows, columns=["Line Item"] + periods)
    return mapped, log


def highlight_model(ws):
    section_fill = PatternFill("solid", fgColor="D9E2F3")
    formula_fill = PatternFill("solid", fgColor="FFF2CC")

    important = {
        "Revenue",
        "Calculated Gross Profit",
        "Reported Gross Profit",
        "Gross Profit Check",
        "Calculated Operating Income",
        "Reported Operating Income",
        "Operating Income Check",
        "Calculated Net Income",
        "Reported Net Income",
        "Net Income Check",
        "Calculated Total Current Assets",
        "Reported Total Current Assets",
        "Current Assets Check",
        "Calculated Total Assets",
        "Reported Total Assets",
        "Total Assets Check",
        "Calculated Total Current Liabilities",
        "Reported Total Current Liabilities",
        "Current Liabilities Check",
        "Calculated Total Liabilities",
        "Reported Total Liabilities",
        "Total Liabilities Check",
        "Calculated Total Equity",
        "Reported Total Equity",
        "Equity Check",
        "L+E Check",
        "Calculated CFO",
        "Reported CFO",
        "CFO Check",
        "Calculated CFI",
        "Reported CFI",
        "CFI Check",
        "Calculated CFF",
        "Reported CFF",
        "CFF Check",
        "Calculated FCF",
        "Reported Free Cash Flow",
        "FCF Check",
    }

    for row in range(2, ws.max_row + 1):
        label = ws[f"A{row}"].value
        if label in important:
            for cell in ws[row]:
                cell.fill = section_fill
                cell.font = Font(bold=True)

    for row in range(2, ws.max_row + 1):
        for col in range(2, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.fill = formula_fill

    ws.column_dimensions["A"].width = 36
    for col in range(2, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16


def add_rows(ws, periods: List[str], rows: List[Tuple[str, List[Optional[float]]]]):
    ws.append(["Line Item"] + periods)
    row_map = {}
    r = 2
    for label, values in rows:
        ws.append([label] + values)
        row_map[label] = r
        r += 1
    return row_map


def col_letters(ws) -> List[str]:
    return [get_column_letter(c) for c in range(2, ws.max_column + 1)]


def make_is_model(writer: pd.ExcelWriter, mapped_df: pd.DataFrame, periods: List[str]):
    ws = writer.book.create_sheet("IS_Model")
    style_sheet(ws, header_color="1F1F1F", freeze="B2")

    value_lookup = {row["Line Item"]: row[periods].tolist() for _, row in mapped_df.iterrows()}
    zeros = [None] * len(periods)

    rows = [
        ("Revenue", value_lookup.get("Revenue", zeros)),
        ("COGS", value_lookup.get("COGS", zeros)),
        ("Calculated Gross Profit", zeros),
        ("Reported Gross Profit", value_lookup.get("Reported Gross Profit", zeros)),
        ("Gross Profit Check", zeros),
        ("R&D", value_lookup.get("R&D", zeros)),
        ("SG&A", value_lookup.get("SG&A", zeros)),
        ("Calculated Total OpEx", zeros),
        ("Reported Operating Expenses", value_lookup.get("Reported Operating Expenses", zeros)),
        ("Calculated Operating Income", zeros),
        ("Reported Operating Income", value_lookup.get("Reported Operating Income", zeros)),
        ("Operating Income Check", zeros),
        ("Interest Expense", value_lookup.get("Interest Expense", zeros)),
        ("Pre-Tax Income", value_lookup.get("Pre-Tax Income", zeros)),
        ("Taxes", value_lookup.get("Taxes", zeros)),
        ("Calculated Net Income", zeros),
        ("Reported Net Income", value_lookup.get("Reported Net Income", zeros)),
        ("Net Income Check", zeros),
        ("Shares Outstanding", value_lookup.get("Shares Outstanding", zeros)),
        ("Calculated EPS", zeros),
        ("Reported EPS", value_lookup.get("Reported EPS", zeros)),
        ("EPS Check", zeros),
    ]

    rm = add_rows(ws, periods, rows)

    for L in col_letters(ws):
        ws[f"{L}{rm['Calculated Gross Profit']}"] = f"={L}{rm['Revenue']}-{L}{rm['COGS']}"
        ws[f"{L}{rm['Gross Profit Check']}"] = f'=IF(OR({L}{rm["Reported Gross Profit"]}="",{L}{rm["Calculated Gross Profit"]}=""),"",{L}{rm["Reported Gross Profit"]}-{L}{rm["Calculated Gross Profit"]})'
        ws[f"{L}{rm['Calculated Total OpEx']}"] = f"={L}{rm['R&D']}+{L}{rm['SG&A']}"
        ws[f"{L}{rm['Calculated Operating Income']}"] = f"={L}{rm['Calculated Gross Profit']}-{L}{rm['Calculated Total OpEx']}"
        ws[f"{L}{rm['Operating Income Check']}"] = f'=IF(OR({L}{rm["Reported Operating Income"]}="",{L}{rm["Calculated Operating Income"]}=""),"",{L}{rm["Reported Operating Income"]}-{L}{rm["Calculated Operating Income"]})'
        ws[f"{L}{rm['Calculated Net Income']}"] = f'=IF(OR({L}{rm["Pre-Tax Income"]}="",{L}{rm["Taxes"]}=""),"",{L}{rm["Pre-Tax Income"]}-{L}{rm["Taxes"]})'
        ws[f"{L}{rm['Net Income Check']}"] = f'=IF(OR({L}{rm["Reported Net Income"]}="",{L}{rm["Calculated Net Income"]}=""),"",{L}{rm["Reported Net Income"]}-{L}{rm["Calculated Net Income"]})'
        ws[f"{L}{rm['Calculated EPS']}"] = f'=IFERROR({L}{rm["Calculated Net Income"]}/{L}{rm["Shares Outstanding"]},"")'
        ws[f"{L}{rm['EPS Check']}"] = f'=IF(OR({L}{rm["Reported EPS"]}="",{L}{rm["Calculated EPS"]}=""),"",{L}{rm["Reported EPS"]}-{L}{rm["Calculated EPS"]})'

    highlight_model(ws)


def make_bs_model(writer: pd.ExcelWriter, mapped_df: pd.DataFrame, periods: List[str]):
    ws = writer.book.create_sheet("BS_Model")
    style_sheet(ws, header_color="1F1F1F", freeze="B2")

    value_lookup = {row["Line Item"]: row[periods].tolist() for _, row in mapped_df.iterrows()}
    zeros = [None] * len(periods)

    rows = [
        ("Cash & Equivalents", value_lookup.get("Cash & Equivalents", zeros)),
        ("Short-Term Investments", value_lookup.get("Short-Term Investments", zeros)),
        ("Receivables", value_lookup.get("Receivables", zeros)),
        ("Inventory", value_lookup.get("Inventory", zeros)),
        ("Other Current Assets", value_lookup.get("Other Current Assets", zeros)),
        ("Calculated Total Current Assets", zeros),
        ("Reported Total Current Assets", value_lookup.get("Reported Total Current Assets", zeros)),
        ("Current Assets Check", zeros),
        ("PP&E, Net", value_lookup.get("PP&E, Net", zeros)),
        ("Goodwill", value_lookup.get("Goodwill", zeros)),
        ("Intangibles", value_lookup.get("Intangibles", zeros)),
        ("Other Non-Current Assets", value_lookup.get("Other Non-Current Assets", zeros)),
        ("Calculated Total Assets", zeros),
        ("Reported Total Assets", value_lookup.get("Reported Total Assets", zeros)),
        ("Total Assets Check", zeros),
        ("Accounts Payable", value_lookup.get("Accounts Payable", zeros)),
        ("Short-Term Debt", value_lookup.get("Short-Term Debt", zeros)),
        ("Other Current Liabilities", value_lookup.get("Other Current Liabilities", zeros)),
        ("Calculated Total Current Liabilities", zeros),
        ("Reported Total Current Liabilities", value_lookup.get("Reported Total Current Liabilities", zeros)),
        ("Current Liabilities Check", zeros),
        ("Long-Term Debt", value_lookup.get("Long-Term Debt", zeros)),
        ("Other Non-Current Liabilities", value_lookup.get("Other Non-Current Liabilities", zeros)),
        ("Calculated Total Liabilities", zeros),
        ("Reported Total Liabilities", value_lookup.get("Reported Total Liabilities", zeros)),
        ("Total Liabilities Check", zeros),
        ("Common Stock", value_lookup.get("Common Stock", zeros)),
        ("Retained Earnings", value_lookup.get("Retained Earnings", zeros)),
        ("Calculated Total Equity", zeros),
        ("Reported Total Equity", value_lookup.get("Reported Total Equity", zeros)),
        ("Equity Check", zeros),
        ("Calculated Total Liabilities + Equity", zeros),
        ("L+E Check", zeros),
        ("Reported Total Debt", value_lookup.get("Reported Total Debt", zeros)),
    ]

    rm = add_rows(ws, periods, rows)

    for L in col_letters(ws):
        ws[f"{L}{rm['Calculated Total Current Assets']}"] = (
            f"={L}{rm['Cash & Equivalents']}+{L}{rm['Short-Term Investments']}+"
            f"{L}{rm['Receivables']}+{L}{rm['Inventory']}+{L}{rm['Other Current Assets']}"
        )
        ws[f"{L}{rm['Current Assets Check']}"] = f'=IF(OR({L}{rm["Reported Total Current Assets"]}="",{L}{rm["Calculated Total Current Assets"]}=""),"",{L}{rm["Reported Total Current Assets"]}-{L}{rm["Calculated Total Current Assets"]})'
        ws[f"{L}{rm['Calculated Total Assets']}"] = (
            f"={L}{rm['Calculated Total Current Assets']}+{L}{rm['PP&E, Net']}+"
            f"{L}{rm['Goodwill']}+{L}{rm['Intangibles']}+{L}{rm['Other Non-Current Assets']}"
        )
        ws[f"{L}{rm['Total Assets Check']}"] = f'=IF(OR({L}{rm["Reported Total Assets"]}="",{L}{rm["Calculated Total Assets"]}=""),"",{L}{rm["Reported Total Assets"]}-{L}{rm["Calculated Total Assets"]})'
        ws[f"{L}{rm['Calculated Total Current Liabilities']}"] = (
            f"={L}{rm['Accounts Payable']}+{L}{rm['Short-Term Debt']}+{L}{rm['Other Current Liabilities']}"
        )
        ws[f"{L}{rm['Current Liabilities Check']}"] = f'=IF(OR({L}{rm["Reported Total Current Liabilities"]}="",{L}{rm["Calculated Total Current Liabilities"]}=""),"",{L}{rm["Reported Total Current Liabilities"]}-{L}{rm["Calculated Total Current Liabilities"]})'
        ws[f"{L}{rm['Calculated Total Liabilities']}"] = (
            f"={L}{rm['Calculated Total Current Liabilities']}+{L}{rm['Long-Term Debt']}+{L}{rm['Other Non-Current Liabilities']}"
        )
        ws[f"{L}{rm['Total Liabilities Check']}"] = f'=IF(OR({L}{rm["Reported Total Liabilities"]}="",{L}{rm["Calculated Total Liabilities"]}=""),"",{L}{rm["Reported Total Liabilities"]}-{L}{rm["Calculated Total Liabilities"]})'
        ws[f"{L}{rm['Calculated Total Equity']}"] = f"={L}{rm['Common Stock']}+{L}{rm['Retained Earnings']}"
        ws[f"{L}{rm['Equity Check']}"] = f'=IF(OR({L}{rm["Reported Total Equity"]}="",{L}{rm["Calculated Total Equity"]}=""),"",{L}{rm["Reported Total Equity"]}-{L}{rm["Calculated Total Equity"]})'
        ws[f"{L}{rm['Calculated Total Liabilities + Equity']}"] = f"={L}{rm['Calculated Total Liabilities']}+{L}{rm['Calculated Total Equity']}"
        ws[f"{L}{rm['L+E Check']}"] = f'=IF(OR({L}{rm["Reported Total Assets"]}="",{L}{rm["Calculated Total Liabilities + Equity"]}=""),"",{L}{rm["Reported Total Assets"]}-{L}{rm["Calculated Total Liabilities + Equity"]})'

    highlight_model(ws)


def make_cf_model(writer: pd.ExcelWriter, mapped_df: pd.DataFrame, periods: List[str]):
    ws = writer.book.create_sheet("CF_Model")
    style_sheet(ws, header_color="1F1F1F", freeze="B2")

    value_lookup = {row["Line Item"]: row[periods].tolist() for _, row in mapped_df.iterrows()}
    zeros = [None] * len(periods)

    rows = [
        ("Reported Net Income", value_lookup.get("Reported Net Income", zeros)),
        ("D&A", value_lookup.get("D&A", zeros)),
        ("Stock-Based Compensation", value_lookup.get("Stock-Based Compensation", zeros)),
        ("Change in Working Capital", value_lookup.get("Change in Working Capital", zeros)),
        ("Calculated CFO", zeros),
        ("Reported CFO", value_lookup.get("Reported CFO", zeros)),
        ("CFO Check", zeros),
        ("Capex", value_lookup.get("Capex", zeros)),
        ("Acquisitions", value_lookup.get("Acquisitions", zeros)),
        ("Calculated CFI", zeros),
        ("Reported CFI", value_lookup.get("Reported CFI", zeros)),
        ("CFI Check", zeros),
        ("Debt Repayment", value_lookup.get("Debt Repayment", zeros)),
        ("Stock Issued", value_lookup.get("Stock Issued", zeros)),
        ("Stock Repurchased", value_lookup.get("Stock Repurchased", zeros)),
        ("Dividends Paid", value_lookup.get("Dividends Paid", zeros)),
        ("Calculated CFF", zeros),
        ("Reported CFF", value_lookup.get("Reported CFF", zeros)),
        ("CFF Check", zeros),
        ("Calculated FCF", zeros),
        ("Reported Free Cash Flow", value_lookup.get("Reported Free Cash Flow", zeros)),
        ("FCF Check", zeros),
    ]

    rm = add_rows(ws, periods, rows)

    for L in col_letters(ws):
        ws[f"{L}{rm['Calculated CFO']}"] = (
            f"={L}{rm['Reported Net Income']}+{L}{rm['D&A']}+"
            f"{L}{rm['Stock-Based Compensation']}-{L}{rm['Change in Working Capital']}"
        )
        ws[f"{L}{rm['CFO Check']}"] = f'=IF(OR({L}{rm["Reported CFO"]}="",{L}{rm["Calculated CFO"]}=""),"",{L}{rm["Reported CFO"]}-{L}{rm["Calculated CFO"]})'
        ws[f"{L}{rm['Calculated CFI']}"] = f"={L}{rm['Capex']}+{L}{rm['Acquisitions']}"
        ws[f"{L}{rm['CFI Check']}"] = f'=IF(OR({L}{rm["Reported CFI"]}="",{L}{rm["Calculated CFI"]}=""),"",{L}{rm["Reported CFI"]}-{L}{rm["Calculated CFI"]})'
        ws[f"{L}{rm['Calculated CFF']}"] = (
            f"={L}{rm['Debt Repayment']}+{L}{rm['Stock Issued']}+"
            f"{L}{rm['Stock Repurchased']}+{L}{rm['Dividends Paid']}"
        )
        ws[f"{L}{rm['CFF Check']}"] = f'=IF(OR({L}{rm["Reported CFF"]}="",{L}{rm["Calculated CFF"]}=""),"",{L}{rm["Reported CFF"]}-{L}{rm["Calculated CFF"]})'
        ws[f"{L}{rm['Calculated FCF']}"] = f"={L}{rm['Calculated CFO']}+{L}{rm['Capex']}"
        ws[f"{L}{rm['FCF Check']}"] = f'=IF(OR({L}{rm["Reported Free Cash Flow"]}="",{L}{rm["Calculated FCF"]}=""),"",{L}{rm["Reported Free Cash Flow"]}-{L}{rm["Calculated FCF"]})'

    highlight_model(ws)


def make_checks(writer: pd.ExcelWriter):
    ws = writer.book.create_sheet("Checks")
    ws.append(["Check", "Meaning"])
    ws.append(["IS Checks", "Compare reported vs calculated Gross Profit / Operating Income / Net Income / EPS"])
    ws.append(["BS Checks", "Compare reported vs calculated Current Assets / Assets / Liabilities / Equity / L+E"])
    ws.append(["CF Checks", "Compare reported vs calculated CFO / CFI / CFF / FCF"])
    style_sheet(ws, header_color="9E480E", freeze="A2")
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 90


def export_model(ticker: str) -> Path:
    ticker = ticker.upper().strip()
    inc, bal, cf = get_statements(ticker)
    periods = choose_periods(inc, bal, cf)

    inc = inc[inc["date"].isin(periods)].copy()
    bal = bal[bal["date"].isin(periods)].copy()
    cf = cf[cf["date"].isin(periods)].copy()

    is_mapped, is_log = build_mapped_df(inc, IS_ALIASES, periods, "IS")
    bs_mapped, bs_log = build_mapped_df(bal, BS_ALIASES, periods, "BS")
    cf_mapped, cf_log = build_mapped_df(cf, CF_ALIASES, periods, "CF")

    mapping_log = pd.DataFrame(is_log + bs_log + cf_log)

    out = Path(f"{ticker}_three_statement_model.xlsx")

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        write_df(writer, "IS_Raw", inc)
        write_df(writer, "BS_Raw", bal)
        write_df(writer, "CF_Raw", cf)
        write_df(writer, "IS_Mapped", is_mapped)
        write_df(writer, "BS_Mapped", bs_mapped)
        write_df(writer, "CF_Mapped", cf_mapped)
        write_df(writer, "Mapping_Log", mapping_log)

        make_is_model(writer, is_mapped, periods)
        make_bs_model(writer, bs_mapped, periods)
        make_cf_model(writer, cf_mapped, periods)
        make_checks(writer)

        # number formatting
        int_fmt = '#,##0_);(#,##0)'
        for s in ["IS_Mapped", "BS_Mapped", "CF_Mapped", "IS_Model", "BS_Model", "CF_Model"]:
            ws = writer.book[s]
            for row in ws.iter_rows(min_row=2, min_col=2):
                for cell in row:
                    cell.number_format = int_fmt

    return out


def main():
    if not API_KEY or API_KEY == "PASTE_YOUR_NEW_FMP_KEY_HERE":
        raise RuntimeError("Add your new FMP API key to the API_KEY line first.")

    ticker = sys.argv[1].strip() if len(sys.argv) > 1 else input("Enter ticker: ").strip()
    output = export_model(ticker)
    print(f"Saved: {output.resolve()}")


if __name__ == "__main__":
    main()
