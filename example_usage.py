"""
Example usage of the Three Statements Export tool.

This script demonstrates how to import and call the export function
programmatically from your own Python code.
"""

import os

from three_statements_export import export_three_statements

# ---------------------------------------------------------------------------
# Set your FMP API key here or via the FMP_API_KEY environment variable.
# Get a free key at: https://financialmodelingprep.com/developer/docs
# ---------------------------------------------------------------------------
API_KEY = os.environ.get("FMP_API_KEY", "YOUR_API_KEY_HERE")


def run_example() -> None:
    """Export three financial statements for a handful of companies."""

    tickers = ["AAPL", "MSFT", "GOOGL"]

    for ticker in tickers:
        print(f"\n{'=' * 60}")
        print(f"  Exporting three statements for: {ticker}")
        print(f"{'=' * 60}")

        output_path = export_three_statements(
            ticker=ticker,
            api_key=API_KEY,
            annual_limit=5,     # last 5 fiscal years
            quarterly_limit=8,  # last 8 quarters
            output_dir="./output",
        )

        print(f"  ✅  Saved to: {output_path}\n")


if __name__ == "__main__":
    if API_KEY == "YOUR_API_KEY_HERE":
        print(
            "⚠️  Please set your FMP API key.\n"
            "   Option 1 – environment variable:  export FMP_API_KEY=<key>\n"
            "   Option 2 – edit example_usage.py and replace 'YOUR_API_KEY_HERE'.\n"
        )
    else:
        run_example()
