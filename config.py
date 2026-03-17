"""
Configuration management for the Three Statements Export tool.
"""

import os

# FMP API base URL
FMP_BASE_URL = "https://financialmodelingprep.com/api/v3"

# API key - set via environment variable FMP_API_KEY or pass directly
FMP_API_KEY = os.environ.get("FMP_API_KEY", "")

# Retry settings
RETRY_ATTEMPTS = 3
RETRY_WAIT_MIN = 1   # seconds
RETRY_WAIT_MAX = 10  # seconds

# Default periods to fetch
DEFAULT_ANNUAL_LIMIT = 5
DEFAULT_QUARTERLY_LIMIT = 12

# Output directory for generated Excel files
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", ".")

# Number format for numeric cells in Excel
NUMBER_FORMAT = "#,##0"

# Excel styling colours (openpyxl hex format without leading #)
HEADER_FILL_COLOR = "1F4E79"       # dark blue
HEADER_FONT_COLOR = "FFFFFF"       # white
SECTION_FILL_COLOR = "BDD7EE"      # light blue
SUBTOTAL_FILL_COLOR = "DDEBF7"     # very light blue
ALTERNATING_ROW_COLOR = "F2F2F2"   # light grey

# Column widths (approximate)
LABEL_COLUMN_WIDTH = 40
DATA_COLUMN_WIDTH = 16
